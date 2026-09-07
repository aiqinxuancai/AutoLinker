// AI HTTP 诊断边界：请求前落盘，流数据逐块记录，不改变传输或重试行为。
#pragma once
#include <atomic>
#include <cctype>
#include <chrono>
#include "Global.h"
#include "Logger.h"
#include "WinINetUtil.h"
#include "..\\thirdparty\\json.hpp"

namespace AIHttpTrace {
inline std::atomic_uint64_t sequence{1};

inline void Write(uint64_t id, const char* event, nlohmann::json data) noexcept
{
    try {
        data["http_id"] = id;
        data["event"] = event;
        data["thread_id"] = GetCurrentThreadId();
        data["pid"] = GetCurrentProcessId();
        Logger::Instance().Write("AI-DEBUG", data.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace));
    } catch (...) {}
}

inline nlohmann::json SafeHeaders(const std::vector<HttpResponseHeaderEntry>& headers)
{
    nlohmann::json result = nlohmann::json::array();
    for (const auto& header : headers) {
        std::string name = header.name;
        for (char& ch : name) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
        const bool safe = name == "content-type" || name == "server" || name == "date" ||
            name == "x-request-id" || name == "request-id" || name == "traceparent" ||
            name == "cf-ray" || name == "retry-after" || name == "x-new-api-request-id";
        result.push_back({{"name", header.name}, {"value", safe ? header.value : "<redacted>"}});
    }
    return result;
}

inline uint64_t Begin(const std::string& url, const std::string& body, int timeout, int idle)
{
    if (!IsAIDebugLogEnabled()) return 0;
    const uint64_t id = sequence.fetch_add(1);
    // 查询串和认证头不落盘；模型参数、工具定义及上下文完整保留在请求正文中。
    Write(id, "request", {{"url", url.substr(0, url.find('?'))}, {"request_body", body},
        {"request_bytes", body.size()}, {"timeout_ms", timeout}, {"idle_timeout_ms", idle}});
    return id;
}

// 网络分块可能截断 UTF-8 字符；十六进制副本用于无损重建原始 SSE。
inline std::string Hex(const std::string& bytes)
{
    constexpr char digits[] = "0123456789abcdef";
    std::string result;
    result.reserve(bytes.size() * 2);
    for (unsigned char ch : bytes) {
        result.push_back(digits[ch >> 4]);
        result.push_back(digits[ch & 15]);
    }
    return result;
}

inline std::pair<std::string, int> Post(const std::string& url, const std::string& body,
    const std::string& headers, int timeout, bool cookies, bool redirect,
    HttpRequestCancellation* cancellation)
{
    const auto id = Begin(url, body, timeout, 0);
    auto result = PerformPostRequestDetailed(url, body, headers, timeout, cookies, redirect, cancellation);
    if (id) Write(id, "response", {{"status", result.statusCode}, {"body", result.body},
        {"headers", SafeHeaders(result.headers)}});
    return {std::move(result.body), result.statusCode};
}

inline std::pair<std::string, int> Stream(const std::string& url, const std::string& body,
    const std::function<bool(const std::string&)>& callback, const std::string& headers,
    int timeout, bool cookies, bool redirect, HttpRequestCancellation* cancellation, int idle = 300000)
{
    const auto id = Begin(url, body, timeout, idle);
    size_t chunkIndex = 0;
    std::vector<HttpResponseHeaderEntry> responseHeaders;
    auto result = PerformPostRequestStreaming(url, body, [&](const std::string& chunk) {
        if (id) Write(id, "stream_chunk", {{"index", chunkIndex++}, {"bytes", chunk.size()},
            {"data", chunk}, {"data_hex", Hex(chunk)}});
        return callback ? callback(chunk) : true;
    }, headers, timeout, cookies, redirect, cancellation, idle, id ? &responseHeaders : nullptr);
    if (id) Write(id, "response", {{"status", result.second}, {"body", result.first},
        {"chunks", chunkIndex}, {"headers", SafeHeaders(responseHeaders)}});
    return result;
}
}
