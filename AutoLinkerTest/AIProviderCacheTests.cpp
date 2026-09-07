// 独立缓存协议测试，不加载 IDE，也不发送付费请求。
#include "../src/AIProviderCache.h"
#include <iostream>
#include <stdexcept>

using namespace AIProviderCache;
int main()
{
    int checks = 0;
    const auto check = [&checks](bool ok, const char* name) {
        if (!ok) throw std::runtime_error(name);
        ++checks;
    };
    try {
        auto c = ParseClaude({{"input_tokens", 100}, {"output_tokens", 200},
            {"cache_read_input_tokens", 8000}, {"cache_creation_input_tokens", 2000}});
        check(c.reported && c.readReported && c.writeReported && c.input == 10100 &&
            c.total == 10300 && c.read == 8000 && c.write == 2000, "Claude totals");
        c = ParseClaude({{"input_tokens", 100}, {"output_tokens", 200}});
        check(c.input == 100 && c.total == 300 && !c.readReported && !c.writeReported, "Claude missing details");
        c = ParseClaude({{"input_tokens", 100}, {"cache_read_input_tokens", 0}, {"cache_creation_input_tokens", 0}});
        check(c.readReported && c.writeReported && c.read == 0, "Claude explicit zero");
        const auto huge = (std::numeric_limits<uint64_t>::max)();
        c = ParseClaude({{"input_tokens", huge}, {"cache_read_input_tokens", huge}, {"output_tokens", huge}});
        check(c.input == (std::numeric_limits<int>::max)() && c.total == c.input, "overflow saturation");
        c = ParseClaude({{"input_tokens", -3}, {"cache_read_input_tokens", "8000"}, {"cache_creation_input_tokens", nullptr}});
        check(!c.reported && !c.readReported && !c.writeReported && c.input == 0, "malformed counts");
        auto g = ParseGemini({{"promptTokenCount", 10000}, {"cachedContentTokenCount", 8000}, {"totalTokenCount", 10400}});
        check(g.input == 10000 && g.total == 10400 && g.read == 8000 && !g.writeReported, "Gemini no double counting");
        g = ParseGemini({{"promptTokenCount", 100}, {"candidatesTokenCount", 20}, {"thoughtsTokenCount", 30}});
        check(g.total == 150 && !g.readReported, "Gemini fallback totals");
        check(!ParseGemini(nullptr).reported && !ParseClaude(Json::array()).reported, "missing usage");
        g = ParseGemini({{"promptTokenCount", 100}, {"cachedContentTokenCount", 0}});
        check(g.readReported && g.read == 0, "Gemini explicit zero");

        Json original = {{"system", "stable system"}, {"messages", Json::array({
            {{"role", "user"}, {"content", "hello"}}
        })}};
        Json request = original;
        ApplyClaude(request);
        check(BreakpointCount(request) == 2 && request["system"][0]["text"] == "stable system" &&
            request["messages"][0]["content"][0]["text"] == "hello", "two breakpoints");
        check(BreakpointCount(original) == 0 && original["messages"][0]["content"].is_string(), "history unmodified");
        const auto once = request;
        ApplyClaude(request);
        check(once == request, "idempotence");
        Json thinking = {{"type", "thinking"}, {"thinking", "reason"}, {"signature", "sig"}};
        original["messages"].push_back({{"role", "assistant"}, {"content", Json::array({thinking})}});
        original["messages"].push_back({{"role", "user"}, {"content", Json::array({
            {{"type", "tool_result"}, {"tool_use_id", "call_1"}, {"content", "done"}}
        })}});
        request = original;
        ApplyClaude(request);
        check(request["messages"][1]["content"][0] == thinking &&
            request["messages"][2]["content"][0].contains("cache_control"), "tool results and thinking preserved");
        check(!request["messages"][0]["content"].is_array(), "old turn markers not accumulated");
        request["tools"] = Json::array();
        for (int i = 0; i < 2; ++i) request["tools"].push_back({{"cache_control", {{"type", "ephemeral"}}}});
        const auto four = request;
        ApplyClaude(request);
        check(request == four && BreakpointCount(request) == 4, "four marker limit");
        request = original;
        request["cache_control"] = {{"type", "ephemeral"}};
        const auto automatic = request;
        ApplyClaude(request);
        check(request == automatic, "preserve automatic caching");
        Json onlyThinking = Json::array({thinking});
        check(!MarkLastBlock(onlyThinking) && onlyThinking[0] == thinking, "never mark thinking");
        Json empty = "";
        check(!MarkLastBlock(empty), "skip empty text");
        std::cout << "PASS: " << checks << " provider cache checks\n";
        return 0;
    } catch (const std::exception& error) {
        std::cerr << "FAIL: " << error.what() << '\n';
        return 1;
    }
}
