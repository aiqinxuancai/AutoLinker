// Claude/Gemini 缓存请求适配和用量解析，不依赖 IDE 或网络，便于离线验证。
#pragma once

#include <algorithm>
#include <cstdint>
#include <limits>
#include "..\\thirdparty\\json.hpp"

namespace AIProviderCache {
using Json = nlohmann::json;

struct Usage {
    bool reported = false;
    bool readReported = false;
    bool writeReported = false;
    int input = 0;
    int total = 0;
    int read = 0;
    int write = 0;
};

inline bool HasCount(const Json& value, const char* key)
{
    if (!value.is_object() || !value.contains(key)) return false;
    const auto& count = value[key];
    return count.is_number_unsigned() || (count.is_number_integer() && count.get<int64_t>() >= 0);
}

inline int Count(const Json& value, const char* key)
{
    if (!HasCount(value, key)) return 0;
    return static_cast<int>((std::min)(value[key].get<uint64_t>(),
        static_cast<uint64_t>((std::numeric_limits<int>::max)())));
}

inline int Sum(int a, int b, int c = 0)
{
    return static_cast<int>((std::min)(static_cast<int64_t>(a) + b + c,
        static_cast<int64_t>((std::numeric_limits<int>::max)())));
}

inline Usage ParseClaude(const Json& value)
{
    Usage usage;
    usage.readReported = HasCount(value, "cache_read_input_tokens");
    usage.writeReported = HasCount(value, "cache_creation_input_tokens");
    usage.reported = HasCount(value, "input_tokens") || usage.readReported || usage.writeReported;
    usage.read = Count(value, "cache_read_input_tokens");
    usage.write = Count(value, "cache_creation_input_tokens");
    // Claude 的 input_tokens 不包含缓存读量和写量，预算必须使用三者之和。
    usage.input = Sum(Count(value, "input_tokens"), usage.read, usage.write);
    usage.total = Sum(usage.input, Count(value, "output_tokens"));
    return usage;
}

inline Usage ParseGemini(const Json& value)
{
    Usage usage;
    usage.reported = HasCount(value, "promptTokenCount");
    usage.readReported = HasCount(value, "cachedContentTokenCount");
    usage.read = Count(value, "cachedContentTokenCount");
    // Gemini 的 promptTokenCount 已包含缓存读量，不可再次相加。
    usage.input = Count(value, "promptTokenCount");
    usage.total = HasCount(value, "totalTokenCount")
        ? (std::max)(usage.input, Count(value, "totalTokenCount"))
        : Sum(usage.input, Count(value, "candidatesTokenCount"), Count(value, "thoughtsTokenCount"));
    return usage;
}

inline size_t BreakpointCount(const Json& value)
{
    size_t count = 0;
    if (value.is_object()) {
        count += value.contains("cache_control") ? 1 : 0;
        for (const auto& item : value.items()) count += BreakpointCount(item.value());
    } else if (value.is_array()) {
        for (const auto& item : value) count += BreakpointCount(item);
    }
    return count;
}

inline bool MarkLastBlock(Json& content)
{
    if (content.is_string()) {
        if (content.get_ref<const std::string&>().empty()) return false;
        content = Json::array({{{"type", "text"}, {"text", content.get<std::string>()}}});
    }
    if (!content.is_array()) return false;
    for (auto it = content.rbegin(); it != content.rend(); ++it) {
        if (!it->is_object()) continue;
        const auto type = it->find("type");
        if (type == it->end() || !type->is_string()) continue;
        if (*type != "text" && *type != "tool_result" && *type != "image" && *type != "document") continue;
        if (it->contains("cache_control")) return false;
        if (*type == "text" && (!it->contains("text") || (*it)["text"] == "")) continue;
        (*it)["cache_control"] = {{"type", "ephemeral"}};
        return true;
    }
    return false;
}

inline void ApplyClaude(Json& request)
{
    if (!request.is_object() || request.contains("cache_control")) return;
    size_t count = BreakpointCount(request);
    if (count < 4 && request.contains("system") && MarkLastBlock(request["system"])) ++count;
    if (count >= 4 || !request.contains("messages") || !request["messages"].is_array()) return;
    // 只标记当前请求副本；历史消息不积累旧断点，思考块/签名保持原样。
    for (auto it = request["messages"].rbegin(); it != request["messages"].rend(); ++it) {
        if (it->is_object() && it->contains("content")) {
            MarkLastBlock((*it)["content"]);
            break;
        }
    }
}
}
