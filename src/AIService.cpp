#include "AIService.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <string_view>
#include <unordered_set>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"

#include "AIChatMcpClient.h"
#include "AIChatToolRegistry.h"
#include "AIChatToolPolicy.h"
#include "AIJsonConfig.h"
#include "ConfigManager.h"
#include "Global.h"
#include "IDEFacade.h"
#include "Logger.h"
#include "WinINetUtil.h"
#include <chrono>

namespace {
using PerfClock = std::chrono::steady_clock;

long long ElapsedMs(const PerfClock::time_point& start)
{
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(PerfClock::now() - start).count());
}

std::string ToLowerAsciiCopy(const std::string& text)
{
	std::string lowered = text;
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
}

bool EndsWithInsensitive(const std::string& text, const std::string& suffix)
{
	if (text.size() < suffix.size()) {
		return false;
	}
	const std::string tail = text.substr(text.size() - suffix.size());
	return ToLowerAsciiCopy(tail) == ToLowerAsciiCopy(suffix);
}

std::string DetectProjectTypeText()
{
	struct Candidate {
		INT fnCode;
		const char* text;
	};

	const Candidate candidates[] = {
		{FN_COMPILE_WINDOWS_DLL, "DLL"},
		{FN_COMPILE_WINDOWS_EXE, "窗口程序 EXE"},
		{FN_COMPILE_WINDOWS_CONOLE_EXE, "控制台程序 EXE"},
		{FN_COMPILE_WINDOWS_ECOM, "易模块"}
	};

	auto& ide = IDEFacade::Instance();
	std::string detected;
	for (const auto& candidate : candidates) {
		if (!ide.IsFunctionEnabled(candidate.fnCode)) {
			continue;
		}
		if (!detected.empty()) {
			detected += " / ";
		}
		detected += candidate.text;
	}

	return detected.empty() ? std::string("未知") : detected;
}

std::string TruncateForLog(const std::string& text, size_t maxLen = 240)
{
	if (text.size() <= maxLen) {
		return text;
	}
	return text.substr(0, maxLen) + "...";
}

constexpr int kAiRequestRetryCount = 5;
constexpr int kAiChatRequestRetryCount = 2;
constexpr int kAiChatRequestTimeoutMs = 60000;
constexpr int kAiRequestCancelledHttpStatus = 499;
constexpr int kMaxToolRounds = 32;

int GetChatRequestTimeoutMs(const AISettings& settings)
{
	return (std::clamp)(settings.timeoutMs, 1000, kAiChatRequestTimeoutMs);
}

bool IsCancelRequested(
	const std::function<bool()>& cancelCallback,
	const HttpRequestCancellation* cancelContext = nullptr)
{
	return (cancelCallback && cancelCallback()) ||
		(cancelContext != nullptr && cancelContext->IsCancelled());
}

bool SleepForRetryWithCancel(
	DWORD delayMs,
	const std::function<bool()>& cancelCallback,
	const HttpRequestCancellation* cancelContext = nullptr)
{
	DWORD sleptMs = 0;
	while (sleptMs < delayMs) {
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return false;
		}
		const DWORD sliceMs = (std::min)(delayMs - sleptMs, static_cast<DWORD>(50));
		::Sleep(sliceMs);
		sleptMs += sliceMs;
	}
	return !IsCancelRequested(cancelCallback, cancelContext);
}

bool IsSuccessfulHttpStatus(int statusCode)
{
	return statusCode >= 200 && statusCode < 300;
}

bool IsRetryableHttpStatus(int statusCode)
{
	return statusCode == 0 ||
		statusCode == 408 ||
		statusCode == 409 ||
		statusCode == 425 ||
		statusCode == 429 ||
		statusCode == 500 ||
		statusCode == 502 ||
		statusCode == 503 ||
		statusCode == 504;
}

bool ContainsRetryableTransportHint(const std::string& responseBody)
{
	const std::string lower = ToLowerAsciiCopy(responseBody);
	return lower.find("error in internetopen") != std::string::npos ||
		lower.find("error in internetcrackurl") != std::string::npos ||
		lower.find("error in internetconnect") != std::string::npos ||
		lower.find("error in httpopenrequest") != std::string::npos ||
		lower.find("error in httpsendrequest") != std::string::npos ||
		lower.find("timeout") != std::string::npos ||
		lower.find("timed out") != std::string::npos ||
		lower.find("cannot connect") != std::string::npos ||
		lower.find("connection") != std::string::npos;
}

bool ShouldRetryAiHttpRequest(int statusCode, const std::string& responseBody)
{
	if (IsSuccessfulHttpStatus(statusCode) || statusCode == kAiRequestCancelledHttpStatus) {
		return false;
	}
	if (statusCode == 0) {
		return responseBody.empty() || ContainsRetryableTransportHint(responseBody);
	}
	return IsRetryableHttpStatus(statusCode);
}

DWORD ComputeAiRetryDelayMs(int retryIndex)
{
	switch (retryIndex) {
	case 0:
		return 250;
	case 1:
		return 500;
	case 2:
		return 1000;
	case 3:
		return 1500;
	default:
		return 2000;
	}
}

void LogAiRetryAttempt(
	const std::string& tag,
	int nextAttemptIndex,
	int maxAttempts,
	int statusCode,
	const std::string& responseBody)
{
	// reason 来自接口/网关返回的响应体；文件日志保留完整内容，IDE 仅显示摘要。
	std::string fileReason = responseBody;
	if (fileReason.empty()) {
		fileReason = statusCode == 0
			? "<no response: network/transport failure>"
			: "<empty response body>";
	}
	std::string ideReason = AIService::Trim(responseBody);
	if (ideReason.empty()) {
		ideReason = fileReason;
	}
	const std::string prefix = std::format(
		"[AI Chat][Retry] {} attempt {}/{} http={} reason=",
		tag,
		nextAttemptIndex,
		maxAttempts,
		statusCode);
	Logger::Instance().WriteSplit(
		"AI",
		prefix + fileReason,
		std::format(
		"[AI Chat][Retry] {} attempt {}/{} http={} reason={}",
		tag,
		nextAttemptIndex,
		maxAttempts,
		statusCode,
		TruncateForLog(ideReason, 120)));
}

void LogAiHttpFailure(const std::string& tag, int statusCode, const std::string& responseBody)
{
	// HTTP 响应体可能包含换行和完整错误 JSON；文件日志保留原文，IDE 仅显示摘要。
	std::string fileResponse = responseBody;
	if (fileResponse.empty()) {
		fileResponse = statusCode == 0
			? "<no response: network/transport failure>"
			: "<empty response body>";
	}
	std::string ideResponse = AIService::Trim(responseBody);
	if (ideResponse.empty()) {
		ideResponse = fileResponse;
	}
	const std::string prefix = std::format(
		"[AI Chat][HTTP Failure] {} http={} response=",
		tag,
		statusCode);
	Logger::Instance().WriteSplit(
		"AI",
		prefix + fileResponse,
		std::format(
		"[AI Chat][HTTP Failure] {} http={} response={}",
		tag,
		statusCode,
		TruncateForLog(ideResponse, 120)));
}

std::string BuildHttpStatusErrorForUi(int statusCode, const std::string& responseBody)
{
	return std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
}

AIChatResult BuildCancelledChatResult(const std::string& partialContentLocal = std::string())
{
	AIChatResult result = {};
	result.cancelled = true;
	result.content = partialContentLocal;
	result.error = "chat request cancelled by user";
	result.httpStatus = kAiRequestCancelledHttpStatus;
	return result;
}

AIChatResult MarkChatResultCancelled(AIChatResult result, const std::string& partialContentLocal = std::string())
{
	result.ok = false;
	result.cancelled = true;
	result.content = partialContentLocal;
	result.error = "chat request cancelled by user";
	result.httpStatus = kAiRequestCancelledHttpStatus;
	return result;
}

std::pair<std::string, int> PerformPostRequestWithRetry(
	const std::string& url,
	const std::string& postData,
	const std::string& customHeaders,
	int timeout,
	bool autoCookies,
	bool neverRedirect,
	const char* retryTag,
	const std::function<bool()>& cancelCallback = {},
	HttpRequestCancellation* cancelContext = nullptr,
	int maxRetryCount = kAiRequestRetryCount,
	int* outAttemptCount = nullptr)
{
	std::pair<std::string, int> lastResult;
	const int boundedRetryCount = (std::clamp)(maxRetryCount, 0, kAiRequestRetryCount);
	if (outAttemptCount != nullptr) {
		*outAttemptCount = 0;
	}
	for (int attempt = 0; attempt <= boundedRetryCount; ++attempt) {
		if (outAttemptCount != nullptr) {
			*outAttemptCount = attempt + 1;
		}
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return std::make_pair(std::string("Request cancelled"), kAiRequestCancelledHttpStatus);
		}
		lastResult = PerformPostRequest(url, postData, customHeaders, timeout, autoCookies, neverRedirect, cancelContext);
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return std::make_pair(std::string("Request cancelled"), kAiRequestCancelledHttpStatus);
		}
		if (!ShouldRetryAiHttpRequest(lastResult.second, lastResult.first) || attempt >= boundedRetryCount) {
			return lastResult;
		}

		LogAiRetryAttempt(
			retryTag == nullptr ? "post" : retryTag,
			attempt + 2,
			boundedRetryCount + 1,
			lastResult.second,
			lastResult.first);
		if (!SleepForRetryWithCancel(ComputeAiRetryDelayMs(attempt), cancelCallback, cancelContext)) {
			return std::make_pair(std::string("Request cancelled"), kAiRequestCancelledHttpStatus);
		}
	}
	return lastResult;
}

std::pair<std::string, int> PerformPostRequestStreamingWithRetry(
	const std::string& url,
	const std::string& postData,
	const std::function<bool(const std::string& chunk)>& onChunk,
	const std::string& customHeaders,
	int timeout,
	bool autoCookies,
	bool neverRedirect,
	const char* retryTag,
	const std::function<bool()>& cancelCallback = {},
	HttpRequestCancellation* cancelContext = nullptr,
	int maxRetryCount = kAiRequestRetryCount,
	int* outAttemptCount = nullptr)
{
	std::pair<std::string, int> lastResult;
	const int boundedRetryCount = (std::clamp)(maxRetryCount, 0, kAiRequestRetryCount);
	if (outAttemptCount != nullptr) {
		*outAttemptCount = 0;
	}
	for (int attempt = 0; attempt <= boundedRetryCount; ++attempt) {
		if (outAttemptCount != nullptr) {
			*outAttemptCount = attempt + 1;
		}
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return std::make_pair(std::string("Request cancelled"), kAiRequestCancelledHttpStatus);
		}
		bool sawChunk = false;
		lastResult = PerformPostRequestStreaming(
			url,
			postData,
			[&onChunk, &sawChunk, &cancelCallback, cancelContext](const std::string& chunk) -> bool {
				if (IsCancelRequested(cancelCallback, cancelContext)) {
					return false;
				}
				if (!chunk.empty()) {
					sawChunk = true;
				}
				return onChunk ? onChunk(chunk) : true;
			},
			customHeaders,
			timeout,
			autoCookies,
			neverRedirect,
			cancelContext);
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return std::make_pair(std::string("Request cancelled"), kAiRequestCancelledHttpStatus);
		}
		const bool streamAccepted = sawChunk && IsSuccessfulHttpStatus(lastResult.second);
		if (streamAccepted || !ShouldRetryAiHttpRequest(lastResult.second, lastResult.first) || attempt >= boundedRetryCount) {
			return lastResult;
		}

		LogAiRetryAttempt(
			retryTag == nullptr ? "stream" : retryTag,
			attempt + 2,
			boundedRetryCount + 1,
			lastResult.second,
			lastResult.first);
		if (!SleepForRetryWithCancel(ComputeAiRetryDelayMs(attempt), cancelCallback, cancelContext)) {
			return std::make_pair(std::string("Request cancelled"), kAiRequestCancelledHttpStatus);
		}
	}
	return lastResult;
}

std::string LocalToUtf8(const std::string& text);

size_t ClampUtf8PrefixBoundary(const std::string& text, size_t maxBytes)
{
	size_t end = (std::min)(maxBytes, text.size());
	while (end > 0 && end < text.size() &&
		(static_cast<unsigned char>(text[end]) & 0xC0) == 0x80) {
		--end;
	}
	return end;
}

size_t ClampUtf8SuffixStartBoundary(const std::string& text, size_t tailBytes)
{
	if (tailBytes >= text.size()) {
		return 0;
	}

	size_t start = text.size() - tailBytes;
	while (start < text.size() &&
		(static_cast<unsigned char>(text[start]) & 0xC0) == 0x80) {
		++start;
	}
	return start;
}

std::string TruncateUtf8Text(const std::string& text, size_t maxBytes)
{
	if (text.size() <= maxBytes) {
		return text;
	}

	const size_t end = ClampUtf8PrefixBoundary(text, maxBytes);
	return text.substr(0, end) + std::format(
		"...[omitted UTF-8 byte range [{}, {}), {} bytes of {} total]",
		end,
		text.size(),
		text.size() - end,
		text.size());
}

std::string BuildUtf8Excerpt(const std::string& text, size_t headBytes, size_t tailBytes)
{
	if (text.size() <= headBytes + tailBytes + 64) {
		return text;
	}

	const size_t headEnd = ClampUtf8PrefixBoundary(text, headBytes);
	const size_t tailStart = ClampUtf8SuffixStartBoundary(text, tailBytes);
	if (tailStart <= headEnd) {
		return TruncateUtf8Text(text, headBytes + tailBytes);
	}

	return text.substr(0, headEnd) +
		std::format(
			"\n...[omitted UTF-8 byte range [{}, {}), {} bytes of {} total]...\n",
			headEnd,
			tailStart,
			tailStart - headEnd,
			text.size()) +
		text.substr(tailStart);
}

bool EndsWithAsciiInsensitive(std::string_view text, std::string_view suffix)
{
	if (text.size() < suffix.size()) {
		return false;
	}

	const size_t offset = text.size() - suffix.size();
	for (size_t i = 0; i < suffix.size(); ++i) {
		const unsigned char left = static_cast<unsigned char>(text[offset + i]);
		const unsigned char right = static_cast<unsigned char>(suffix[i]);
		if (std::tolower(left) != std::tolower(right)) {
			return false;
		}
	}
	return true;
}

std::string ToLowerAsciiCopy(std::string_view text)
{
	std::string lowered(text.begin(), text.end());
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
}

bool IsCodeLikeKey(std::string_view key)
{
	const std::string lowered = ToLowerAsciiCopy(key);
	return lowered == "code" ||
		lowered == "proposed_code" ||
		lowered == "plain_text" ||
		EndsWithAsciiInsensitive(lowered, "_code");
}

bool IsTraceLikeKey(std::string_view key)
{
	const std::string lowered = ToLowerAsciiCopy(key);
	return lowered == "trace" ||
		EndsWithAsciiInsensitive(lowered, "_trace");
}

size_t GetCompactArrayLimit(const std::string& toolName, std::string_view key)
{
	const std::string lowered = ToLowerAsciiCopy(key);
	if (toolName == "read_files" && (lowered == "files" || lowered == "omitted_requests")) {
		return 12;
	}
	if (toolName == "search_code") {
		if (lowered == "queries" || lowered == "continuations") {
			return 16;
		}
		if (lowered == "results" || lowered == "context") {
			return 20;
		}
	}
	if (lowered == "hunks") {
		return 4;
	}
	if (lowered == "matches") {
		return 5;
	}
	if (lowered == "symbols") {
		return 16;
	}
	if (lowered == "results") {
		return 8;
	}
	return 6;
}

void RebuildCompactResultPagination(
	nlohmann::json& result,
	std::string_view arrayKey,
	size_t visibleLimit)
{
	if (!result.is_object()) {
		return;
	}
	const std::string key(arrayKey);
	auto arrayIt = result.find(key);
	if (arrayIt == result.end() || !arrayIt->is_array() || arrayIt->size() <= visibleLimit) {
		return;
	}

	const size_t sourceReturned = arrayIt->size();
	arrayIt->erase(
		arrayIt->begin() + static_cast<nlohmann::json::difference_type>(visibleLimit),
		arrayIt->end());
	const int visibleReturned = static_cast<int>(arrayIt->size());
	const int requestedOffset = result.value("requested_offset", result.value("offset", 0));
	const int effectiveOffset = (std::max)(0, result.value("offset", requestedOffset));
	const long long minimumTotal = static_cast<long long>(effectiveOffset) + static_cast<long long>(sourceReturned);
	const long long reportedTotal = result.value("total_results", minimumTotal);
	const long long totalResults = (std::max)(minimumTotal, reportedTotal);
	const long long endOffset = (std::min)(
		totalResults,
		static_cast<long long>(effectiveOffset) + visibleReturned);
	const bool hasMore = endOffset < totalResults;

	result["requested_offset"] = requestedOffset;
	result["offset"] = effectiveOffset;
	result["returned"] = visibleReturned;
	result["total_results"] = totalResults;
	result["has_more"] = hasMore;
	result["next_offset"] = hasMore ? nlohmann::json(endOffset) : nlohmann::json(nullptr);
	result["visible_result_range"] = visibleReturned > 0
		? nlohmann::json({
			{"start_offset", effectiveOffset},
			{"end_offset_exclusive", endOffset}
		})
		: nlohmann::json(nullptr);

	nlohmann::json omittedRanges = nlohmann::json::array();
	if (effectiveOffset > 0) {
		omittedRanges.push_back({
			{"start_offset", 0},
			{"end_offset_exclusive", effectiveOffset}
		});
	}
	if (hasMore) {
		omittedRanges.push_back({
			{"start_offset", endOffset},
			{"end_offset_exclusive", totalResults}
		});
	}
	result["omitted_result_ranges"] = std::move(omittedRanges);
	result["truncated"] = true;
	result["context_page_compacted"] = true;
	result["context_source_page_returned"] = sourceReturned;
	result["context_omitted_items"] = sourceReturned - arrayIt->size();
	result["context_note"] =
		"The model-visible page was compacted. Continue from next_offset, which now follows the last visible item.";
	if (result.contains("page_limit") && result["page_limit"].is_number_integer()) {
		result["source_page_limit"] = result["page_limit"];
		result["page_limit"] = visibleLimit;
	}
}

void PreparePaginatedToolResultForCompactContext(
	const std::string& toolName,
	nlohmann::json& result)
{
	if (!result.is_object()) {
		return;
	}
	if (toolName == "list_files") {
		RebuildCompactResultPagination(result, "files", GetCompactArrayLimit(toolName, "files"));
		return;
	}
	if (toolName != "search_code") {
		return;
	}

	const size_t resultLimit = GetCompactArrayLimit(toolName, "results");
	if (result.contains("queries") && result["queries"].is_array()) {
		bool hasMore = false;
		bool truncated = result.value("truncated", false);
		nlohmann::json continuations = nlohmann::json::array();
		for (auto& query : result["queries"]) {
			RebuildCompactResultPagination(query, "results", resultLimit);
			const bool queryHasMore = query.is_object() && query.value("has_more", false);
			hasMore = hasMore || queryHasMore;
			truncated = truncated || (query.is_object() && query.value("truncated", false));
			if (queryHasMore) {
				continuations.push_back({
					{"pattern", query.value("pattern", std::string())},
					{"next_offset", query.contains("next_offset") ? query["next_offset"] : nlohmann::json(nullptr)},
					{"mirror_generation", result.value("mirror_generation", 0ull)}
				});
			}
		}
		result["has_more"] = hasMore;
		result["truncated"] = truncated;
		result["continuations"] = std::move(continuations);
		if (result.contains("page_limit") && result["page_limit"].is_number_integer() &&
			result["page_limit"].get<long long>() > static_cast<long long>(resultLimit)) {
			result["source_page_limit"] = result["page_limit"];
			result["page_limit"] = resultLimit;
		}
		return;
	}

	RebuildCompactResultPagination(result, "results", resultLimit);
}

nlohmann::json CompactToolContextJsonValue(
	const nlohmann::json& value,
	const std::string& toolName,
	std::string_view key,
	int depth)
{
	if (depth >= 6) {
		return "[omitted: max depth]";
	}

	if (value.is_null() || value.is_boolean() || value.is_number()) {
		return value;
	}

	if (value.is_string()) {
		const std::string text = value.get<std::string>();
		const std::string loweredKey = ToLowerAsciiCopy(key);
		if ((toolName == "read_file" || toolName == "read_code_item") && loweredKey == "content") {
			return BuildUtf8Excerpt(
				text,
				AIChatToolPolicy::kReadFileContextBytes - 4096,
				4096);
		}
		if (toolName == "read_files" && loweredKey == "content") {
			return BuildUtf8Excerpt(
				text,
				AIChatToolPolicy::kReadFilesPerFileContextBytes - 2048,
				2048);
		}
		if (toolName == "read_real_file") {
			if (loweredKey == "content") {
				return BuildUtf8Excerpt(
					text,
					AIChatToolPolicy::kReadFileContextBytes - 4096,
					4096);
			}
			if (loweredKey == "code") {
				return BuildUtf8Excerpt(
					text,
					AIChatToolPolicy::kReadRealFileCodeContextBytes - 8192,
					8192);
			}
		}
		if (toolName == "search_code" && loweredKey == "text") {
			return TruncateUtf8Text(text, 1600);
		}
		if (IsCodeLikeKey(key)) {
			return BuildUtf8Excerpt(text, 2400, 900);
		}
		if (IsTraceLikeKey(key)) {
			return TruncateUtf8Text(text, 1200);
		}
		return TruncateUtf8Text(text, 800);
	}

	if (value.is_array()) {
		const size_t limit = GetCompactArrayLimit(toolName, key);
		nlohmann::json out = nlohmann::json::array();
		for (size_t i = 0; i < value.size() && i < limit; ++i) {
			out.push_back(CompactToolContextJsonValue(value[i], toolName, key, depth + 1));
		}
		if (value.size() > limit) {
			out.push_back({
				{"_truncated", true},
				{"omitted_items", value.size() - limit},
				{"omitted_index_range", {
					{"start_index", limit},
					{"end_index_exclusive", value.size()}
				}}
			});
		}
		return out;
	}

	if (value.is_object()) {
		nlohmann::json out = nlohmann::json::object();
		for (auto it = value.begin(); it != value.end(); ++it) {
			out[it.key()] = CompactToolContextJsonValue(it.value(), toolName, it.key(), depth + 1);
		}
		return out;
	}

	return TruncateUtf8Text(value.dump(), 800);
}

struct CompactToolResultPayload {
	std::string textUtf8;
	nlohmann::json jsonValue = nlohmann::json::object();
};

struct HttpHeaderEntry {
	std::string name;
	std::string value;
};

CompactToolResultPayload BuildCompactToolResultPayload(const std::string& toolName, const std::string& toolResultLocal)
{
	CompactToolResultPayload payload;
	const std::string resultUtf8 = LocalToUtf8(toolResultLocal);

	try {
		nlohmann::json parsed = nlohmann::json::parse(resultUtf8);
		if (parsed.is_object() &&
			parsed.value("ok", false) &&
			(toolName == "edit_file" ||
			 toolName == "multi_edit_file" ||
			 toolName == "write_file" ||
			 toolName == "restore_file_snapshot")) {
			parsed.erase("code");
			parsed.erase("real_code");
			parsed["context_note"] = "Verified write output omitted from model context; use code_hash/new_hash and continue without re-reading.";
		}
		PreparePaginatedToolResultForCompactContext(toolName, parsed);
		nlohmann::json compact = CompactToolContextJsonValue(parsed, toolName, "", 0);
		if (!compact.is_object()) {
			payload.jsonValue = {
				{"tool_name", toolName},
				{"result", compact}
			};
		}
		else {
			payload.jsonValue = std::move(compact);
		}
	}
	catch (...) {
		payload.jsonValue = {
			{"tool_name", toolName},
			{"ok", false},
			{"text", BuildUtf8Excerpt(resultUtf8, 1800, 600)}
		};
	}

	payload.textUtf8 = payload.jsonValue.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
	return payload;
}

bool IsValidHttpHeaderName(std::string_view name)
{
	if (name.empty()) {
		return false;
	}

	for (const unsigned char ch : name) {
		if (ch <= 32 || ch >= 127 || ch == ':') {
			return false;
		}
	}
	return true;
}

bool ParseCustomHeadersTextInternal(
	const std::string& headerText,
	std::vector<HttpHeaderEntry>& outHeaders,
	std::string& outError)
{
	outHeaders.clear();
	outError.clear();

	size_t lineStart = 0;
	int lineNumber = 1;
	while (lineStart <= headerText.size()) {
		size_t lineEnd = headerText.find_first_of("\r\n", lineStart);
		std::string line = lineEnd == std::string::npos
			? headerText.substr(lineStart)
			: headerText.substr(lineStart, lineEnd - lineStart);
		if (lineEnd != std::string::npos && headerText[lineEnd] == '\r' &&
			lineEnd + 1 < headerText.size() && headerText[lineEnd + 1] == '\n') {
			lineStart = lineEnd + 2;
		}
		else if (lineEnd != std::string::npos) {
			lineStart = lineEnd + 1;
		}
		else {
			lineStart = headerText.size() + 1;
		}

		const std::string trimmedLine = AIService::Trim(line);
		if (trimmedLine.empty()) {
			++lineNumber;
			continue;
		}

		const size_t colonPos = trimmedLine.find(':');
		if (colonPos == std::string::npos) {
			outError = std::format("custom header line {} missing ':' separator", lineNumber);
			return false;
		}

		const std::string name = AIService::Trim(trimmedLine.substr(0, colonPos));
		if (!IsValidHttpHeaderName(name)) {
			outError = std::format("custom header line {} has invalid header name", lineNumber);
			return false;
		}

		outHeaders.push_back({
			name,
			AIService::Trim(trimmedLine.substr(colonPos + 1))
		});
		++lineNumber;
	}

	return true;
}

void UpsertHeaderEntry(
	std::vector<HttpHeaderEntry>& headers,
	const std::string& name,
	const std::string& value)
{
	const std::string loweredName = ToLowerAsciiCopy(name);
	for (auto& entry : headers) {
		if (ToLowerAsciiCopy(entry.name) == loweredName) {
			entry.name = name;
			entry.value = value;
			return;
		}
	}
	headers.push_back({ name, value });
}

std::string BuildMergedHeaders(
	const std::vector<HttpHeaderEntry>& baseHeaders,
	const AISettings& settings)
{
	std::vector<HttpHeaderEntry> merged = baseHeaders;
	std::vector<HttpHeaderEntry> customHeaders;
	std::string parseError;
	if (!ParseCustomHeadersTextInternal(settings.customHeadersText, customHeaders, parseError)) {
		return std::string();
	}

	for (const auto& entry : customHeaders) {
		UpsertHeaderEntry(merged, entry.name, entry.value);
	}

	std::string serialized;
	for (const auto& entry : merged) {
		serialized += entry.name;
		serialized += ": ";
		serialized += entry.value;
		serialized += "\r\n";
	}
	return serialized;
}

bool ValidateRequestSettings(const AISettings& settings, std::string& outError)
{
	std::string missingField;
	if (!AIService::HasRequiredSettings(settings, missingField)) {
		outError = "AI settings missing: " + missingField;
		return false;
	}

	if (!AIService::ValidateCustomHeadersText(settings.customHeadersText, outError)) {
		return false;
	}

	outError.clear();
	return true;
}

bool ContainsAsciiInsensitive(std::string_view text, std::string_view needle)
{
	if (needle.empty()) {
		return true;
	}
	return ToLowerAsciiCopy(text).find(ToLowerAsciiCopy(needle)) != std::string::npos;
}

bool IsDeepSeekCompatibleSettings(const AISettings& settings)
{
	return ContainsAsciiInsensitive(settings.baseUrl, "deepseek") ||
		ContainsAsciiInsensitive(settings.model, "deepseek");
}

bool IsGemini25Model(std::string_view model)
{
	return ContainsAsciiInsensitive(model, "gemini-2.5");
}

bool IsGemini25ProModel(std::string_view model)
{
	return IsGemini25Model(model) && ContainsAsciiInsensitive(model, "pro");
}

bool IsGemini3Model(std::string_view model)
{
	return ContainsAsciiInsensitive(model, "gemini-3");
}

bool IsGemini3FlashModel(std::string_view model)
{
	return IsGemini3Model(model) && ContainsAsciiInsensitive(model, "flash");
}

bool IsClaudeMythosPreviewModel(std::string_view model)
{
	return ContainsAsciiInsensitive(model, "claude-mythos-preview");
}

bool IsClaudeAdaptiveThinkingModel(std::string_view model)
{
	return IsClaudeMythosPreviewModel(model) ||
		ContainsAsciiInsensitive(model, "claude-opus-4-7") ||
		ContainsAsciiInsensitive(model, "claude-opus-4-6") ||
		ContainsAsciiInsensitive(model, "claude-sonnet-4-6");
}

bool IsOpenAIGpt5Model(std::string_view model)
{
	return ContainsAsciiInsensitive(model, "gpt-5");
}

void ApplyOpenAITemperatureIfSupported(nlohmann::json& requestBody, const AISettings& settings)
{
	if (IsOpenAIGpt5Model(settings.model)) {
		return;
	}
	requestBody["temperature"] = settings.temperature;
}

bool ShouldSkipOpenAIChatReasoningForToolUse(const AISettings& settings)
{
	return !IsDeepSeekCompatibleSettings(settings) &&
		IsOpenAIGpt5Model(settings.model);
}

std::string GetOpenAIReasoningEffort(AIThinkingLevel level)
{
	switch (level) {
	case AIThinkingLevel::Low:
		return "low";
	case AIThinkingLevel::Medium:
		return "medium";
	case AIThinkingLevel::High:
		return "high";
	case AIThinkingLevel::XHigh:
		return "xhigh";
	case AIThinkingLevel::Max:
		return "max";
	case AIThinkingLevel::Ultra:
		// 与 Codex CLI 一致：Ultra 在请求层使用 max，额外能力来自客户端多代理调度。
		return "max";
	case AIThinkingLevel::Off:
	default:
		return "none";
	}
}

std::string GetClaudeEffort(AIThinkingLevel level)
{
	switch (level) {
	case AIThinkingLevel::Low:
		return "low";
	case AIThinkingLevel::Medium:
		return "medium";
	case AIThinkingLevel::High:
	case AIThinkingLevel::XHigh:
	case AIThinkingLevel::Max:
	case AIThinkingLevel::Ultra:
		return "high";
	case AIThinkingLevel::Off:
	default:
		return std::string();
	}
}

int GetClaudeThinkingBudget(AIThinkingLevel level)
{
	switch (level) {
	case AIThinkingLevel::Low:
		return 1024;
	case AIThinkingLevel::Medium:
		return 4096;
	case AIThinkingLevel::High:
	case AIThinkingLevel::XHigh:
	case AIThinkingLevel::Max:
	case AIThinkingLevel::Ultra:
		return 8192;
	case AIThinkingLevel::Off:
	default:
		return 0;
	}
}

int GetGemini25ThinkingBudget(const AISettings& settings)
{
	switch (settings.thinkingLevel) {
	case AIThinkingLevel::Off:
		return IsGemini25ProModel(settings.model) ? 128 : 0;
	case AIThinkingLevel::Low:
		return 1024;
	case AIThinkingLevel::Medium:
		return 4096;
	case AIThinkingLevel::High:
	case AIThinkingLevel::XHigh:
	case AIThinkingLevel::Max:
	case AIThinkingLevel::Ultra:
	default:
		return -1;
	}
}

std::string GetGemini3ThinkingLevel(const AISettings& settings)
{
	if (IsGemini3FlashModel(settings.model)) {
		switch (settings.thinkingLevel) {
		case AIThinkingLevel::Off:
			return "minimal";
		case AIThinkingLevel::Low:
			return "low";
		case AIThinkingLevel::Medium:
			return "medium";
		case AIThinkingLevel::High:
		case AIThinkingLevel::XHigh:
		case AIThinkingLevel::Max:
		case AIThinkingLevel::Ultra:
		default:
			return "high";
		}
	}

	switch (settings.thinkingLevel) {
	case AIThinkingLevel::High:
	case AIThinkingLevel::XHigh:
	case AIThinkingLevel::Max:
	case AIThinkingLevel::Ultra:
		return "high";
	case AIThinkingLevel::Off:
	case AIThinkingLevel::Low:
	case AIThinkingLevel::Medium:
	default:
		return "low";
	}
}

void EnsureDeepSeekAssistantMessageCompat(nlohmann::json& message)
{
	if (!message.is_object()) {
		return;
	}

	if (!message.contains("content") || message["content"].is_null()) {
		message["content"] = "";
	}
	if (message.contains("reasoning_content") && message["reasoning_content"].is_null()) {
		message["reasoning_content"] = "";
	}
}

void ApplyThinkingConfigToOpenAIChatRequest(nlohmann::json& requestBody, const AISettings& settings)
{
	if (IsDeepSeekCompatibleSettings(settings)) {
		if (settings.thinkingLevel == AIThinkingLevel::Off) {
			requestBody["thinking"] = {
				{"type", "disabled"}
			};
			return;
		}

		requestBody["thinking"] = {
			{"type", "enabled"}
		};
		requestBody["reasoning_effort"] = settings.thinkingLevel >= AIThinkingLevel::High ? "max" : "high";
		return;
	}

	requestBody["reasoning_effort"] = GetOpenAIReasoningEffort(settings.thinkingLevel);
}

void ApplyThinkingConfigToOpenAIResponsesRequest(nlohmann::json& requestBody, const AISettings& settings)
{
	requestBody["reasoning"] = {
		{"effort", GetOpenAIReasoningEffort(settings.thinkingLevel)}
	};
	requestBody["include"] = nlohmann::json::array({ "reasoning.encrypted_content" });
}

void ApplyThinkingConfigToClaudeRequest(nlohmann::json& requestBody, const AISettings& settings)
{
	if (IsClaudeAdaptiveThinkingModel(settings.model)) {
		const std::string effort = GetClaudeEffort(settings.thinkingLevel);
		if (effort.empty()) {
			if (!IsClaudeMythosPreviewModel(settings.model)) {
				requestBody["thinking"] = {
					{"type", "disabled"}
				};
			}
			return;
		}

		requestBody["thinking"] = {
			{"type", "adaptive"}
		};
		if (!requestBody.contains("output_config") || !requestBody["output_config"].is_object()) {
			requestBody["output_config"] = nlohmann::json::object();
		}
		requestBody["output_config"]["effort"] = effort;
		return;
	}

	const int budget = GetClaudeThinkingBudget(settings.thinkingLevel);
	if (budget <= 0) {
		return;
	}

	requestBody["thinking"] = {
		{"type", "enabled"},
		{"budget_tokens", budget}
	};
}

void ApplyThinkingConfigToGeminiRequest(nlohmann::json& requestBody, const AISettings& settings)
{
	nlohmann::json thinkingConfig = nlohmann::json::object();
	if (IsGemini25Model(settings.model)) {
		thinkingConfig["thinkingBudget"] = GetGemini25ThinkingBudget(settings);
	}
	else if (IsGemini3Model(settings.model)) {
		thinkingConfig["thinkingLevel"] = GetGemini3ThinkingLevel(settings);
	}
	else {
		thinkingConfig["thinkingBudget"] = settings.thinkingLevel >= AIThinkingLevel::High ? -1 : 1024;
	}

	if (!requestBody.contains("generationConfig") || !requestBody["generationConfig"].is_object()) {
		requestBody["generationConfig"] = nlohmann::json::object();
	}
	requestBody["generationConfig"]["thinkingConfig"] = std::move(thinkingConfig);
}

std::string BuildToolRoundsExceededError(int maxToolRounds, const std::vector<AIChatToolEvent>& toolEvents)
{
	std::string message = std::format("tool call rounds exceeded limit ({})", maxToolRounds);
	if (!toolEvents.empty()) {
		message += " after ";
		message += std::to_string(toolEvents.size());
		message += " tool calls";
	}
	return message;
}

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		&wide[0],
		wideLen) <= 0) {
		return text;
	}

	const int outLen = WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		&out[0],
		outLen,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return out;
}

std::string LocalToUtf8(const std::string& text)
{
	// AutoLinker/IDE strings are typically local ANSI (GBK on zh-CN Windows).
	// Convert before feeding nlohmann::json, which requires UTF-8.
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

void NormalizeJsonStringsToUtf8InPlace(nlohmann::json& value)
{
	if (value.is_string()) {
		value = LocalToUtf8(value.get_ref<const std::string&>());
		return;
	}
	if (value.is_array()) {
		for (auto& item : value) {
			NormalizeJsonStringsToUtf8InPlace(item);
		}
		return;
	}
	if (value.is_object()) {
		for (auto& item : value.items()) {
			NormalizeJsonStringsToUtf8InPlace(item.value());
		}
	}
}

std::string Utf8ToLocal(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

using ChatToolCallback = std::function<std::string(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk)>;

struct ChatToolExecutionResult {
	std::string resultLocal;
	bool ok = false;
};

std::string AppendToolPolicyNotice(
	const std::string& resultLocal,
	const std::string& notice,
	const AIChatToolPolicy::Session& policy)
{
	if (notice.empty()) {
		return resultLocal;
	}
	try {
		nlohmann::json result = nlohmann::json::parse(LocalToUtf8(resultLocal));
		if (!result.is_object()) {
			return resultLocal;
		}
		result["_tool_policy"] = {
			{"message", notice},
			{"exploration_calls_used", policy.ExplorationCalls()},
			{"exploration_call_limit", AIChatToolPolicy::kHardExplorationCallLimit}
		};
		return Utf8ToLocal(result.dump());
	}
	catch (...) {
		return resultLocal;
	}
}

ChatToolExecutionResult ExecuteChatToolWithPolicy(
	AIChatToolPolicy::Session& policy,
	const ChatToolCallback& toolCallback,
	const std::string& toolName,
	const std::string& argumentsJsonUtf8)
{
	ChatToolExecutionResult execution;
	const AIChatToolPolicy::Decision decision = policy.BeforeToolCall(toolName, argumentsJsonUtf8);
	if (!decision.allowed) {
		execution.resultLocal = Utf8ToLocal(decision.resultJsonUtf8);
		return execution;
	}

	if (toolCallback) {
		execution.resultLocal = toolCallback(toolName, argumentsJsonUtf8, execution.ok);
	}
	else {
		execution.resultLocal = R"({"ok":false,"error":"tool callback not set"})";
	}
	const std::string notice = policy.AfterToolCall(
		toolName,
		argumentsJsonUtf8,
		LocalToUtf8(execution.resultLocal),
		execution.ok);
	execution.resultLocal = AppendToolPolicyNotice(execution.resultLocal, notice, policy);
	return execution;
}

AISettings BuildChatRoundSettings(const AISettings& settings, const AIChatToolPolicy::Session& policy)
{
	AISettings roundSettings = settings;
	if (policy.PreferLowThinkingForNextRound() &&
		roundSettings.thinkingLevel > AIThinkingLevel::Low) {
		roundSettings.thinkingLevel = AIThinkingLevel::Low;
	}
	return roundSettings;
}

void LogChatRoundMetrics(
	const char* tag,
	int round,
	size_t requestBytes,
	long long elapsedMs,
	int statusCode,
	int attempts,
	int explorationCalls)
{
	Logger::Instance().Write(
		"AI",
		std::format(
			"[AI Chat][Round] tag={} round={} request_bytes={} elapsed_ms={} http={} attempts={} exploration_calls={}",
			tag == nullptr ? "chat" : tag,
			round + 1,
			requestBytes,
			elapsedMs,
			statusCode,
			attempts,
			explorationCalls));
}

// 读取与当前源文件同目录、同名的 {stem}.AGENTS.md 项目规范文件。
// 文件不存在时返回空串；存在时返回去除 UTF-8 BOM 后的内容（已转为本地编码）。
std::string ReadProjectAgentsMd()
{
	if (AIService::Trim(g_nowOpenSourceFilePath).empty()) {
		return {};
	}
	try {
		const std::filesystem::path src(g_nowOpenSourceFilePath);
		const std::filesystem::path agentsMdPath =
			src.parent_path() / (src.stem().string() + ".AGENTS.md");
		if (!std::filesystem::exists(agentsMdPath)) {
			return {};
		}
		std::ifstream f(agentsMdPath, std::ios::binary);
		if (!f.is_open()) {
			return {};
		}
		std::string content((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
		// 去除 UTF-8 BOM（EF BB BF）
		if (content.size() >= 3 &&
			static_cast<unsigned char>(content[0]) == 0xEF &&
			static_cast<unsigned char>(content[1]) == 0xBB &&
			static_cast<unsigned char>(content[2]) == 0xBF) {
			content.erase(0, 3);
		}
		return Utf8ToLocal(content);
	}
	catch (...) {
		return {};
	}
}

std::string RemoveCodeFence(const std::string& text)
{
	const std::string content = AIService::Trim(text);
	size_t fenceBegin = content.find("```");
	if (fenceBegin == std::string::npos) {
		return content;
	}

	size_t firstLineEnd = content.find('\n', fenceBegin + 3);
	if (firstLineEnd == std::string::npos) {
		return content;
	}

	size_t fenceEnd = content.find("```", firstLineEnd + 1);
	if (fenceEnd == std::string::npos) {
		return content;
	}

	return AIService::Trim(content.substr(firstLineEnd + 1, fenceEnd - (firstLineEnd + 1)));
}

std::string MergeMessageContentUtf8(const nlohmann::json& message)
{
	std::string merged;
	if (!message.contains("content")) {
		return merged;
	}

	const nlohmann::json& content = message["content"];
	if (content.is_string()) {
		return content.get<std::string>();
	}
	if (!content.is_array()) {
		return merged;
	}

	for (const auto& item : content) {
		if (item.is_string()) {
			merged += item.get<std::string>();
			continue;
		}
		if (item.is_object() && item.contains("text") && item["text"].is_string()) {
			merged += item["text"].get<std::string>();
		}
	}
	return merged;
}

bool ExtractChatResponseMessage(const nlohmann::json& parsed, nlohmann::json& outMessage, std::string& outError)
{
	if (!parsed.contains("choices") || !parsed["choices"].is_array() || parsed["choices"].empty()) {
		outError = "AI response choices is empty";
		return false;
	}
	const nlohmann::json& choice = parsed["choices"][0];
	if (!choice.contains("message") || !choice["message"].is_object()) {
		outError = "AI response message missing";
		return false;
	}
	outMessage = choice["message"];
	return true;
}

struct StreamToolCallState {
	std::string id;
	std::string name;
	std::string arguments;
};

struct ChatStreamParseState {
	bool sawDataEvent = false;
	std::string pendingLine;
	std::string mergedUtf8;
	std::string reasoningContentUtf8;
	std::vector<StreamToolCallState> toolCalls;
	std::string parseError;
	bool hasUsage = false;
	int promptTokens = 0;
	int totalTokens = 0;
};

StreamToolCallState& EnsureToolCallSlot(std::vector<StreamToolCallState>& toolCalls, size_t index)
{
	if (toolCalls.size() <= index) {
		toolCalls.resize(index + 1);
	}
	return toolCalls[index];
}

bool ProcessStreamDataPayload(
	const std::string& payload,
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	if (payload.empty()) {
		return true;
	}
	if (payload == "[DONE]") {
		state.sawDataEvent = true;
		return true;
	}

	state.sawDataEvent = true;
	nlohmann::json packet;
	try {
		packet = nlohmann::json::parse(payload);
	}
	catch (const std::exception& ex) {
		state.parseError = std::string("Failed to parse streaming chunk JSON: ") + ex.what();
		return false;
	}

	if (packet.contains("error") && packet["error"].is_object()) {
		const auto& err = packet["error"];
		if (err.contains("message") && err["message"].is_string()) {
			state.parseError = Utf8ToLocal(err["message"].get<std::string>());
		}
		else {
			state.parseError = "AI streaming response contains error";
		}
		return false;
	}

	// usage 通常随最后一个 chunk 下发（需 stream_options.include_usage），其 choices 为空数组，
	// 故须在下面的 choices 早退之前捕获。
	if (packet.contains("usage") && packet["usage"].is_object()) {
		const auto& u = packet["usage"];
		if (u.contains("prompt_tokens") && u["prompt_tokens"].is_number_integer()) {
			state.promptTokens = u["prompt_tokens"].get<int>();
		}
		if (u.contains("total_tokens") && u["total_tokens"].is_number_integer()) {
			state.totalTokens = u["total_tokens"].get<int>();
		}
		state.hasUsage = true;
	}

	if (!packet.contains("choices") || !packet["choices"].is_array() || packet["choices"].empty()) {
		return true;
	}

	const auto& choice = packet["choices"][0];
	if (!choice.contains("delta") || !choice["delta"].is_object()) {
		return true;
	}
	const auto& delta = choice["delta"];

	const std::string deltaContentUtf8 = MergeMessageContentUtf8(delta);
	if (!deltaContentUtf8.empty()) {
		state.mergedUtf8 += deltaContentUtf8;
		if (streamCallback) {
			streamCallback(Utf8ToLocal(deltaContentUtf8));
		}
	}

	if (delta.contains("reasoning_content") && delta["reasoning_content"].is_string()) {
		state.reasoningContentUtf8 += delta["reasoning_content"].get<std::string>();
	}

	if (!delta.contains("tool_calls") || !delta["tool_calls"].is_array()) {
		return true;
	}

	for (const auto& toolCallDelta : delta["tool_calls"]) {
		if (!toolCallDelta.is_object()) {
			continue;
		}

		size_t index = state.toolCalls.size();
		if (toolCallDelta.contains("index") && toolCallDelta["index"].is_number_integer()) {
			const int idx = toolCallDelta["index"].get<int>();
			if (idx >= 0) {
				index = static_cast<size_t>(idx);
			}
		}

		auto& slot = EnsureToolCallSlot(state.toolCalls, index);
		if (toolCallDelta.contains("id") && toolCallDelta["id"].is_string()) {
			const std::string deltaId = toolCallDelta["id"].get<std::string>();
			if (!deltaId.empty()) {
				slot.id = deltaId;
			}
		}

		if (!toolCallDelta.contains("function") || !toolCallDelta["function"].is_object()) {
			continue;
		}

		const auto& fn = toolCallDelta["function"];
		if (fn.contains("name") && fn["name"].is_string()) {
			slot.name += fn["name"].get<std::string>();
		}
		if (fn.contains("arguments") && fn["arguments"].is_string()) {
			slot.arguments += fn["arguments"].get<std::string>();
		}
	}
	return true;
}

bool ProcessStreamLine(
	const std::string& rawLine,
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	std::string line = rawLine;
	if (!line.empty() && line.back() == '\r') {
		line.pop_back();
	}
	if (line.empty()) {
		return true;
	}
	if (line.rfind("data:", 0) != 0) {
		return true;
	}

	std::string payload = line.substr(5);
	if (!payload.empty() && payload[0] == ' ') {
		payload.erase(payload.begin());
	}
	return ProcessStreamDataPayload(payload, state, streamCallback);
}

bool ConsumeStreamChunk(
	const std::string& chunk,
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	state.pendingLine += chunk;
	size_t lineEnd = 0;
	while ((lineEnd = state.pendingLine.find('\n')) != std::string::npos) {
		const std::string line = state.pendingLine.substr(0, lineEnd);
		state.pendingLine.erase(0, lineEnd + 1);
		if (!ProcessStreamLine(line, state, streamCallback)) {
			return false;
		}
	}
	return true;
}

bool FlushStreamParseState(
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	if (state.pendingLine.empty()) {
		return true;
	}
	const std::string line = state.pendingLine;
	state.pendingLine.clear();
	return ProcessStreamLine(line, state, streamCallback);
}

nlohmann::json BuildAssistantMessageFromStreamState(const ChatStreamParseState& state)
{
	nlohmann::json message;
	message["role"] = "assistant";

	if (!state.toolCalls.empty()) {
		message["content"] = state.mergedUtf8;
		if (!state.reasoningContentUtf8.empty()) {
			message["reasoning_content"] = state.reasoningContentUtf8;
		}
		message["tool_calls"] = nlohmann::json::array();
		for (size_t i = 0; i < state.toolCalls.size(); ++i) {
			const auto& call = state.toolCalls[i];
			std::string callId = call.id;
			if (callId.empty()) {
				callId = std::format("call_auto_{}", i + 1);
			}
			message["tool_calls"].push_back({
				{"id", callId},
				{"type", "function"},
				{"function", {
					{"name", call.name},
					{"arguments", call.arguments}
				}}
			});
		}
		return message;
	}

	message["content"] = state.mergedUtf8;
	if (!state.reasoningContentUtf8.empty()) {
		message["reasoning_content"] = state.reasoningContentUtf8;
	}
	return message;
}

struct ResponsesStreamParseState {
	bool sawSseEvent = false;
	std::string pendingLine;
	std::string eventName;
	std::string eventData;
	std::string mergedTextUtf8;
	std::string parseError;
	nlohmann::json completedResponse = nlohmann::json::object();
	nlohmann::json outputItems = nlohmann::json::array();
};

std::string ExtractResponsesStreamError(const nlohmann::json& packet)
{
	const nlohmann::json* error = nullptr;
	if (packet.contains("error") && packet["error"].is_object()) {
		error = &packet["error"];
	}
	else if (packet.contains("response") && packet["response"].is_object() &&
		packet["response"].contains("error") && packet["response"]["error"].is_object()) {
		error = &packet["response"]["error"];
	}
	if (error != nullptr && error->contains("message") && (*error)["message"].is_string()) {
		return Utf8ToLocal((*error)["message"].get<std::string>());
	}
	return "Responses streaming request failed";
}

void AddResponsesOutputItemUnique(ResponsesStreamParseState& state, const nlohmann::json& item)
{
	if (!item.is_object()) {
		return;
	}
	const std::string id = item.value("id", std::string());
	if (!id.empty()) {
		for (auto& existing : state.outputItems) {
			if (existing.is_object() && existing.value("id", std::string()) == id) {
				existing = item;
				return;
			}
		}
	}
	state.outputItems.push_back(item);
}

bool ProcessResponsesStreamEvent(
	ResponsesStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	if (state.eventName.empty() && state.eventData.empty()) {
		return true;
	}
	const std::string eventName = state.eventName;
	const std::string eventData = state.eventData;
	state.eventName.clear();
	state.eventData.clear();
	state.sawSseEvent = true;
	if (eventData.empty() || eventData == "[DONE]") {
		return true;
	}

	nlohmann::json packet;
	try {
		packet = nlohmann::json::parse(eventData);
	}
	catch (const std::exception& ex) {
		state.parseError = std::string("Failed to parse Responses streaming event: ") + ex.what();
		return false;
	}

	const std::string type = packet.value("type", eventName);
	if (type == "response.output_text.delta") {
		const std::string delta = packet.value("delta", std::string());
		if (!delta.empty()) {
			state.mergedTextUtf8 += delta;
			if (streamCallback) {
				streamCallback(Utf8ToLocal(delta));
			}
		}
		return true;
	}
	if (type == "response.output_item.done") {
		if (packet.contains("item")) {
			AddResponsesOutputItemUnique(state, packet["item"]);
		}
		return true;
	}
	if (type == "response.completed") {
		if (packet.contains("response") && packet["response"].is_object()) {
			state.completedResponse = packet["response"];
		}
		return true;
	}
	if (type == "response.failed" || type == "response.incomplete" || type == "error") {
		state.parseError = ExtractResponsesStreamError(packet);
		return false;
	}
	return true;
}

bool ProcessResponsesStreamLine(
	const std::string& rawLine,
	ResponsesStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	std::string line = rawLine;
	if (!line.empty() && line.back() == '\r') {
		line.pop_back();
	}
	if (line.empty()) {
		return ProcessResponsesStreamEvent(state, streamCallback);
	}
	if (line.front() == ':') {
		return true;
	}
	if (line.rfind("event:", 0) == 0) {
		state.eventName = AIService::Trim(line.substr(6));
		return true;
	}
	if (line.rfind("data:", 0) == 0) {
		std::string data = line.substr(5);
		if (!data.empty() && data.front() == ' ') {
			data.erase(data.begin());
		}
		if (!state.eventData.empty()) {
			state.eventData.push_back('\n');
		}
		state.eventData += data;
	}
	return true;
}

bool ConsumeResponsesStreamChunk(
	const std::string& chunk,
	ResponsesStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	state.pendingLine += chunk;
	size_t lineEnd = 0;
	while ((lineEnd = state.pendingLine.find('\n')) != std::string::npos) {
		const std::string line = state.pendingLine.substr(0, lineEnd);
		state.pendingLine.erase(0, lineEnd + 1);
		if (!ProcessResponsesStreamLine(line, state, streamCallback)) {
			return false;
		}
	}
	return true;
}

bool FlushResponsesStreamState(
	ResponsesStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	if (!state.pendingLine.empty()) {
		const std::string line = state.pendingLine;
		state.pendingLine.clear();
		if (!ProcessResponsesStreamLine(line, state, streamCallback)) {
			return false;
		}
	}
	return ProcessResponsesStreamEvent(state, streamCallback);
}

nlohmann::json BuildResponsesParsedFromStream(const ResponsesStreamParseState& state)
{
	if (state.completedResponse.is_object() && !state.completedResponse.empty()) {
		return state.completedResponse;
	}
	nlohmann::json parsed;
	parsed["output"] = state.outputItems;
	return parsed;
}

bool TryParseResponsesResponseBody(
	const std::string& responseBody,
	nlohmann::json& outParsed,
	std::string& outStreamTextUtf8,
	std::string& outError)
{
	outParsed = nlohmann::json();
	outStreamTextUtf8.clear();
	outError.clear();

	std::string jsonParseError;
	try {
		outParsed = nlohmann::json::parse(responseBody);
		return true;
	}
	catch (const std::exception& ex) {
		jsonParseError = ex.what();
	}

	ResponsesStreamParseState streamState;
	const bool streamParsed =
		ConsumeResponsesStreamChunk(responseBody, streamState, {}) &&
		FlushResponsesStreamState(streamState, {});
	if (!streamParsed || !streamState.parseError.empty()) {
		outError = streamState.parseError.empty()
			? "Failed to parse Responses streaming response"
			: streamState.parseError;
		return false;
	}
	if (!streamState.sawSseEvent) {
		outError = "Failed to parse Responses API response: " + jsonParseError;
		return false;
	}

	outParsed = BuildResponsesParsedFromStream(streamState);
	outStreamTextUtf8 = std::move(streamState.mergedTextUtf8);
	return true;
}

bool TryParseRawChatMessageJson(const std::string& rawMessageJsonUtf8, nlohmann::json& outMessage)
{
	if (AIService::Trim(rawMessageJsonUtf8).empty()) {
		return false;
	}
	try {
		outMessage = nlohmann::json::parse(rawMessageJsonUtf8);
		return outMessage.is_object();
	}
	catch (...) {
		return false;
	}
}

struct OpenAIChatToolSequenceRepairStats {
	size_t removedMessages = 0;
	size_t removedIncompleteGroups = 0;
};

std::vector<std::string> ReadAssistantToolCallIds(const nlohmann::json& message)
{
	std::vector<std::string> ids;
	if (!message.is_object() ||
		!message.contains("tool_calls") ||
		!message["tool_calls"].is_array() ||
		message["tool_calls"].empty()) {
		return ids;
	}

	std::unordered_set<std::string> seen;
	for (const auto& toolCall : message["tool_calls"]) {
		if (!toolCall.is_object() || !toolCall.contains("id") || !toolCall["id"].is_string()) {
			continue;
		}
		const std::string id = AIService::Trim(toolCall["id"].get<std::string>());
		if (!id.empty() && seen.insert(id).second) {
			ids.push_back(id);
		}
	}
	return ids;
}

OpenAIChatToolSequenceRepairStats RepairOpenAIChatToolMessageSequence(nlohmann::json& messages)
{
	OpenAIChatToolSequenceRepairStats stats;
	if (!messages.is_array() || messages.empty()) {
		return stats;
	}

	nlohmann::json repaired = nlohmann::json::array();
	for (size_t i = 0; i < messages.size();) {
		const nlohmann::json& message = messages[i];
		const std::string role = message.is_object()
			? ToLowerAsciiCopy(AIService::Trim(message.value("role", std::string())))
			: std::string();
		if (role == "tool") {
			++stats.removedMessages;
			++i;
			continue;
		}

		const bool hasToolCalls = role == "assistant" &&
			message.contains("tool_calls") &&
			message["tool_calls"].is_array() &&
			!message["tool_calls"].empty();
		if (!hasToolCalls) {
			repaired.push_back(message);
			++i;
			continue;
		}

		const std::vector<std::string> expectedIds = ReadAssistantToolCallIds(message);
		const std::unordered_set<std::string> expected(expectedIds.begin(), expectedIds.end());
		std::unordered_set<std::string> answered;
		nlohmann::json matchingToolMessages = nlohmann::json::array();
		size_t next = i + 1;
		while (next < messages.size()) {
			const nlohmann::json& candidate = messages[next];
			const std::string candidateRole = candidate.is_object()
				? ToLowerAsciiCopy(AIService::Trim(candidate.value("role", std::string())))
				: std::string();
			if (candidateRole != "tool") {
				break;
			}

			const std::string callId = candidate.value("tool_call_id", std::string());
			if (expected.contains(callId) && answered.insert(callId).second) {
				matchingToolMessages.push_back(candidate);
			}
			++next;
		}

		const bool complete = !expectedIds.empty() && answered.size() == expectedIds.size();
		const size_t originalGroupSize = next - i;
		if (complete) {
			repaired.push_back(message);
			for (const auto& toolMessage : matchingToolMessages) {
				repaired.push_back(toolMessage);
			}
			stats.removedMessages += originalGroupSize - 1 - matchingToolMessages.size();
		}
		else {
			stats.removedMessages += originalGroupSize;
			++stats.removedIncompleteGroups;
		}
		i = next;
	}

	messages = std::move(repaired);
	return stats;
}

nlohmann::json BuildPublicToolCatalog()
{
	nlohmann::json tools = nlohmann::json::array();
	tools.push_back({
		{"name", "refresh_workspace_mirror"},
		{"description", "Refresh the current e-packager workspace mirror from live IDE memory. External MCP clients MUST call this successfully before their first source read/edit tool in each MCP session. AutoLinker's built-in AI chat refreshes automatically and does not expose this tool. mode=auto keeps the default strategy, main_only refreshes only main source files when possible, and full rebuilds the complete mirror."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"mode", {{"type", "string"}, {"enum", nlohmann::json::array({"auto", "main_only", "full"})}, {"description", "Defaults to auto."}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "update_plan"},
		{"description", "Update the visible task plan only for genuinely complex, multi-file, or explicitly planned work. Do not use it for a localized single-file change. It does not modify source code and does not replace the <proposed_plan> approval flow in plan mode."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"explanation", {{"type", "string"}, {"description", "Optional short note about why the plan changed."}}},
				{"plan", {{"type", "array"}, {"items", {
					{"type", "object"},
					{"properties", {
						{"step", {{"type", "string"}}},
						{"status", {{"type", "string"}, {"enum", nlohmann::json::array({"pending", "in_progress", "completed"})}}}
					}},
					{"required", nlohmann::json::array({"step", "status"})},
					{"additionalProperties", false}
				}}, {"description", "Task steps. At most one item may be in_progress."}}}
			}},
			{"required", nlohmann::json::array({"plan"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_files"},
		{"description", "List files in the current e-packager workspace mirror. Paths are relative to the mirror root. External MCP clients must call refresh_workspace_mirror first; built-in AI chat is refreshed automatically. By default focuses on src, ecom, elib and header text areas. Continue paginated results with next_offset."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"glob", {{"type", "string"}, {"description", "Optional glob such as src/**/*.txt or ecom/**/*.txt."}}},
				{"path", {{"type", "string"}, {"description", "Optional relative path prefix."}}},
				{"mirror_generation", {{"type", "integer"}, {"minimum", 0}, {"description", "Optional generation returned by the previous page; rejects stale pagination after a refresh or write."}}},
				{"offset", {{"type", "integer"}, {"minimum", 0}, {"description", "Zero-based result offset for pagination. Use next_offset from the prior result."}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 5000}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_code"},
		{"description", "Batch-stream search complete text files inside the current e-packager workspace mirror. External MCP clients must call refresh_workspace_mirror first; built-in AI chat is refreshed automatically. Literal substring search is the default; set regex=true only for simple bounded regular expressions. Use patterns to match up to 16 related names in one file pass."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"pattern", {{"type", "string"}, {"maxLength", 1024}, {"description", "Single literal substring by default; set regex=true for regular-expression search."}}},
				{"patterns", {{"type", "array"}, {"maxItems", 16}, {"items", {{"type", "string"}, {"maxLength", 1024}}}, {"description", "Optional batch of patterns using the same regex/case/glob/context options."}}},
				{"glob", {{"type", "string"}, {"description", "Optional file glob filter such as src/**/*.txt."}}},
				{"output_mode", {{"type", "string"}, {"enum", nlohmann::json::array({"files_with_matches", "content", "count"})}}},
				{"regex", {{"type", "boolean"}, {"description", "Defaults to false."}}},
				{"case_insensitive", {{"type", "boolean"}}},
				{"context", {{"type", "integer"}, {"minimum", 0}, {"maximum", 20}}},
				{"mirror_generation", {{"type", "integer"}, {"minimum", 0}, {"description", "Optional generation returned by the previous page; rejects stale pagination."}}},
				{"offset", {{"type", "integer"}, {"minimum", 0}, {"description", "Zero-based result offset. Continue with next_offset returned by the previous page."}}},
				{"head_limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 2000}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_file"},
		{"description", "Read one text file from the current e-packager workspace mirror. External MCP clients must call refresh_workspace_mirror first. Returns numbered text, line pagination, mirror_generation and a 1 MiB source byte window. Continue large files with next_source_byte_offset as byte_offset."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"mirror_generation", {{"type", "integer"}, {"minimum", 0}, {"description", "Optional generation returned by the previous page."}}},
				{"byte_offset", {{"type", "integer"}, {"minimum", 0}, {"description", "Source byte cursor for files larger than the 1 MiB read window. Continue with next_source_byte_offset and reset line offset to 0."}}},
				{"offset", {{"type", "integer"}, {"minimum", 0}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 20000}}}
			}},
			{"required", nlohmann::json::array({"file_path"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_files"},
		{"description", "Batch read one or more text files or byte windows from the current e-packager workspace mirror. External MCP clients must call refresh_workspace_mirror first. Prefer this over repeated read_file calls. Partial failures are explicit and large files continue with per-file next_source_byte_offset."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_paths", {{"type", "array"}, {"maxItems", 12}, {"items", {{"type", "string"}}}, {"description", "Simple list of up to 12 mirror-relative file paths."}}},
				{"mirror_generation", {{"type", "integer"}, {"minimum", 0}, {"description", "Optional generation returned by a prior mirror read."}}},
				{"files", {{"type", "array"}, {"maxItems", 12}, {"items", {
					{"type", "object"},
					{"properties", {
						{"file_path", {{"type", "string"}}},
						{"byte_offset", {{"type", "integer"}, {"minimum", 0}}},
						{"offset", {{"type", "integer"}, {"minimum", 0}}},
						{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 2000}}}
					}},
					{"required", nlohmann::json::array({"file_path"})},
					{"additionalProperties", false}
				}}, {"description", "Optional per-file offset/limit entries, up to 12 per call."}}},
				{"byte_offset", {{"type", "integer"}, {"minimum", 0}, {"description", "Default source byte cursor for all entries."}}},
				{"offset", {{"type", "integer"}, {"minimum", 0}, {"description", "Default offset for file_paths/files entries."}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 2000}, {"description", "Default per-file line limit."}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_code_item"},
		{"description", "Stream-locate and read one complete E-language top-level code item (especially a .子程序), including items beyond the first 1 MiB. External MCP clients must call refresh_workspace_mirror first. Use occurrence to disambiguate duplicate names. Optional include_references adds a bounded case-insensitive literal reference search."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"item_name", {{"type", "string"}}},
				{"occurrence", {{"type", "integer"}, {"minimum", 1}, {"description", "Select the Nth matching declaration when duplicate top-level names exist."}}},
				{"mirror_generation", {{"type", "integer"}, {"minimum", 0}, {"description", "Optional generation returned by a prior mirror read."}}},
				{"include_references", {{"type", "boolean"}, {"description", "Defaults to false."}}},
				{"reference_glob", {{"type", "string"}, {"description", "Optional reference search glob such as src/**/*.txt."}}},
				{"reference_limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 50}}}
			}},
			{"required", nlohmann::json::array({"file_path", "item_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_real_file"},
		{"description", "Read a bounded numbered view directly from the live IDE page mapped by mirror-relative file_path. External MCP clients must call refresh_workspace_mirror first. Returns code_hash plus paginated content; continue with next_offset when has_more=true. Call it immediately before editing and base old_text/full_code/expected_base_hash on this live real_source view."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"offset", {{"type", "integer"}, {"minimum", 0}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 20000}}}
			}},
			{"required", nlohmann::json::array({"file_path"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "edit_file"},
		{"description", "Edit one current-project source file by mirror-relative file_path. External MCP clients must call refresh_workspace_mirror first. The edit is always based on the live IDE page; expected_base_hash optionally rejects a stale caller base. Successful results are verified."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"old_text", {{"type", "string"}}},
				{"new_text", {{"type", "string"}}},
				{"expected_base_hash", {{"type", "string"}, {"description", "code_hash from read_real_file; required for external MCP calls and rejects stale live source."}}}
			}},
			{"required", nlohmann::json::array({"file_path", "old_text", "new_text"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "multi_edit_file"},
		{"description", "Apply multiple text edits to one current-project source file. External MCP clients must call refresh_workspace_mirror first. The edits are always based on the live IDE page; expected_base_hash optionally rejects stale live source."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"edits", {{"type", "array"}, {"items", {
					{"type", "object"},
					{"properties", {
						{"old_text", {{"type", "string"}}},
						{"new_text", {{"type", "string"}}},
						{"replace_all", {{"type", "boolean"}}}
					}},
					{"required", nlohmann::json::array({"old_text", "new_text"})},
					{"additionalProperties", false}
				}}}},
				{"fail_on_unmatched", {{"type", "boolean"}}},
				{"atomic", {{"type", "boolean"}}},
				{"expected_base_hash", {{"type", "string"}, {"description", "code_hash from read_real_file; required for external MCP calls and rejects stale live source."}}}
			}},
			{"required", nlohmann::json::array({"file_path", "edits"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "write_file"},
		{"description", "Overwrite one current-project source file with full_code. External MCP clients must call refresh_workspace_mirror first. expected_base_hash detects stale source. Successful results include verified=true and do not require confirmation reads."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"full_code", {{"type", "string"}}},
				{"expected_base_hash", {{"type", "string"}, {"description", "code_hash from read_real_file; required for external MCP calls."}}}
			}},
			{"required", nlohmann::json::array({"file_path", "full_code"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "diff_file"},
		{"description", "Preview a structured diff for one current-project source file without writing anything. The diff is based on the live IDE page. External MCP clients must call refresh_workspace_mirror first. Accepts new_code/full_code or text edit parameters."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"new_code", {{"type", "string"}}},
				{"full_code", {{"type", "string"}}},
				{"old_text", {{"type", "string"}}},
				{"new_text", {{"type", "string"}}},
				{"edits", {{"type", "array"}, {"items", {
					{"type", "object"},
					{"properties", {
						{"old_text", {{"type", "string"}}},
						{"new_text", {{"type", "string"}}},
						{"replace_all", {{"type", "boolean"}}}
					}},
					{"required", nlohmann::json::array({"old_text", "new_text"})},
					{"additionalProperties", false}
				}}}},
				{"fail_on_unmatched", {{"type", "boolean"}}},
				{"expected_base_hash", {{"type", "string"}, {"description", "code_hash from read_real_file; required for external MCP calls and rejects stale live source."}}}
			}},
			{"required", nlohmann::json::array({"file_path"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "restore_file_snapshot"},
		{"description", "Restore one current-project source file from the latest real-page snapshot or a specified snapshot_id. The current live page is read first; expected_current_hash optionally prevents overwriting newer changes."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"file_path", {{"type", "string"}}},
				{"snapshot_id", {{"type", "string"}}},
				{"restore_latest", {{"type", "boolean"}}},
				{"expected_current_hash", {{"type", "string"}, {"description", "Current live code hash; required for external MCP restore calls."}}}
			}},
			{"required", nlohmann::json::array({"file_path"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_current_page_info"},
		{"description", "Get current IDE page name, page type and the trace/source used to resolve that page name."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_current_eide_info"},
		{"description", "Get current E-language IDE instance information, including current source file path, current page info, MCP port/endpoint, process id, executable path and supported compile modes."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "refresh_dependency_catalog"},
		{"description", "Explicit-user-request only. Refresh AutoLinker dependency catalog cache for available .ec modules under ecom and .fne support libraries under lib."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"force", {{"type", "boolean"}, {"description", "Rebuild cache directories. Defaults to false."}}},
				{"wait", {{"type", "boolean"}, {"description", "Wait until refresh finishes. Defaults to false so the IDE main thread is not blocked."}}},
				{"timeout_ms", {{"type", "integer"}, {"minimum", 0}, {"maximum", 600000}, {"description", "Defaults to 30000 when wait=true; 0 also uses the bounded default."}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_available_modules"},
		{"description", "Explicit-user-request only. Search currently available .ec modules from the E-language ecom directory by module name, path, or cached main source."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"query", {{"type", "string"}, {"description", "Module name, file name, path or source keyword. Empty query lists top cached modules."}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 200}}},
				{"include_snippets", {{"type", "boolean"}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_available_support_libraries"},
		{"description", "Explicit-user-request only. Search currently available .fne support libraries from the E-language lib directory by library name, file name, path, or decoded GetNewInf information."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"query", {{"type", "string"}, {"description", "Support library name, file name, path, command name or type keyword. Empty query lists top cached libraries."}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 200}}},
				{"include_snippets", {{"type", "boolean"}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_imported_modules"},
		{"description", "Explicit-user-request only. List .ec modules currently imported by the active E-language project."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "add_module_to_project"},
		{"description", "Explicit-user-request only. Import one .ec module into the current project by module_path or module_name. Real import requires allow_blocking_import=true."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"module_path", {{"type", "string"}, {"description", "Absolute .ec file path."}}},
				{"module_name", {{"type", "string"}, {"description", "Module name or file name to resolve from the E-language module directory."}}},
				{"prefer_new_method", {{"type", "boolean"}, {"description", "Optional diagnostic switch. false uses the legacy IDE import path; true tries the newer AddECOM2 path first."}}},
				{"allow_blocking_import", {{"type", "boolean"}, {"description", "Required true to execute the real IDE module import call. Default false only resolves the module and avoids blocking MCP if the IDE API hangs."}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "remove_module_from_project"},
		{"description", "Explicit-user-request only. Remove one imported .ec module from the current project by module_index, module_path, or module_name."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"module_index", {{"type", "integer"}, {"minimum", 0}}},
				{"module_path", {{"type", "string"}}},
				{"module_name", {{"type", "string"}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "add_support_library_to_project"},
		{"description", "Explicit-user-request only. Add one .fne support library to the current project by library_path or library_name, then verify it is loaded."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"library_path", {{"type", "string"}, {"description", "Absolute .fne file path, or path relative to the E-language lib directory."}}},
				{"library_name", {{"type", "string"}, {"description", "Support library name or file name to resolve from the E-language lib directory."}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "compile_with_output_path"},
		{"description", "Compile the current project with a specified output path, suppressing the IDE save-file dialog. target=auto detects the current project type. A successful result requires the output artifact to exist and be updated."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"target", {{"type", "string"}, {"enum", nlohmann::json::array({"auto", "win_exe", "win_console_exe", "win_dll", "ecom"})}, {"description", "Defaults to auto."}}},
				{"output_path", {{"type", "string"}}},
				{"static_compile", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"output_path"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "run_powershell_command"},
		{"description", "Run one PowerShell command on the local machine after explicit user confirmation."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"command", {{"type", "string"}}},
				{"working_directory", {{"type", "string"}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 600}}}
			}},
			{"required", nlohmann::json::array({"command"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_web_tavily"},
		{"description", "Search the public web via Tavily and return normalized result snippets."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"query", {{"type", "string"}}},
				{"max_results", {{"type", "integer"}, {"minimum", 1}, {"maximum", 10}}},
				{"topic", {{"type", "string"}}}
			}},
			{"required", nlohmann::json::array({"query"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "fetch_url"},
		{"description", "Fetch one URL via HTTP GET and return normalized text response plus basic HTTP metadata."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"url", {{"type", "string"}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 300}}},
				{"max_bytes", {{"type", "integer"}, {"minimum", 4096}, {"maximum", 2097152}}}
			}},
			{"required", nlohmann::json::array({"url"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "extract_web_document"},
		{"description", "Fetch one web page or text document and extract readable plain-text content, title and a small set of links."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"url", {{"type", "string"}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 300}}},
				{"max_bytes", {{"type", "integer"}, {"minimum", 4096}, {"maximum", 2097152}}}
			}},
			{"required", nlohmann::json::array({"url"})},
			{"additionalProperties", false}
		}}
	});
	NormalizeJsonStringsToUtf8InPlace(tools);
	return tools;
}

bool IsRealPageReadToolVisible(AISourceEditMode mode)
{
	return mode == AISourceEditMode::RealPageFirst;
}

nlohmann::json FilterToolCatalogForSourceEditMode(const nlohmann::json& catalog, AISourceEditMode mode)
{
	if (IsRealPageReadToolVisible(mode)) {
		return catalog;
	}

	nlohmann::json filtered = nlohmann::json::array();
	for (const auto& item : catalog) {
		if (item.is_object() && item.value("name", std::string()) == "read_real_file") {
			continue;
		}
		filtered.push_back(item);
	}
	return filtered;
}

nlohmann::json BuildConfiguredToolCatalog(const AISettings& settings)
{
	return AIChatMcpClient::AppendMcpToolsToCatalog(
		FilterToolCatalogForSourceEditMode(BuildPublicToolCatalog(), settings.sourceEditMode));
}

nlohmann::json BuildInternalToolCatalog(const AISettings& settings);

nlohmann::json FilterInternalChatToolCatalog(const nlohmann::json& catalog)
{
	nlohmann::json filtered = nlohmann::json::array();
	for (const auto& item : catalog) {
		if (!item.is_object() || item.value("name", std::string()) == "refresh_workspace_mirror") {
			continue;
		}
		filtered.push_back(item);
	}
	return filtered;
}

nlohmann::json BuildChatToolDefinitions(const AISettings& settings)
{
	const nlohmann::json catalog = BuildInternalToolCatalog(settings);
	nlohmann::json tools = nlohmann::json::array();
	for (const auto& item : catalog) {
		tools.push_back({
			{"type", "function"},
			{"function", {
				{"name", item.value("name", "")},
				{"description", item.value("description", "")},
				{"parameters", item.contains("inputSchema") ? item["inputSchema"] : nlohmann::json::object()}
			}}
		});
	}
	return tools;
}

nlohmann::json BuildChatToolDefinitions(
	const AISettings& settings,
	const std::vector<AIChatMessage>&)
{
	return BuildChatToolDefinitions(settings);
}

std::string TruncateGeminiDescription(const std::string& text)
{
	if (text.size() <= 240) {
		return text;
	}
	return text.substr(0, 240);
}

nlohmann::json SanitizeGeminiSchema(const nlohmann::json& schema)
{
	if (!schema.is_object()) {
		return nlohmann::json::object();
	}

	nlohmann::json out = nlohmann::json::object();
	if (schema.contains("type") && (schema["type"].is_string() || schema["type"].is_array())) {
		out["type"] = schema["type"];
	}
	if (schema.contains("description") && schema["description"].is_string()) {
		out["description"] = TruncateGeminiDescription(schema["description"].get<std::string>());
	}
	if (schema.contains("enum") && schema["enum"].is_array()) {
		out["enum"] = schema["enum"];
	}
	if (schema.contains("required") && schema["required"].is_array()) {
		out["required"] = schema["required"];
	}
	if (schema.contains("format") && schema["format"].is_string()) {
		out["format"] = schema["format"];
	}
	if (schema.contains("pattern") && schema["pattern"].is_string()) {
		out["pattern"] = schema["pattern"];
	}
	if (schema.contains("nullable") && schema["nullable"].is_boolean()) {
		out["nullable"] = schema["nullable"];
	}
	const auto copyNumber = [&schema, &out](const char* key) {
		if (schema.contains(key) && schema[key].is_number()) {
			out[key] = schema[key];
		}
	};
	copyNumber("minimum");
	copyNumber("maximum");
	copyNumber("exclusiveMinimum");
	copyNumber("exclusiveMaximum");
	copyNumber("multipleOf");
	copyNumber("minLength");
	copyNumber("maxLength");
	copyNumber("minItems");
	copyNumber("maxItems");
	copyNumber("minProperties");
	copyNumber("maxProperties");
	if (schema.contains("additionalProperties")) {
		if (schema["additionalProperties"].is_boolean()) {
			out["additionalProperties"] = schema["additionalProperties"];
		}
		else if (schema["additionalProperties"].is_object()) {
			out["additionalProperties"] = SanitizeGeminiSchema(schema["additionalProperties"]);
		}
	}
	if (schema.contains("items") && schema["items"].is_object()) {
		out["items"] = SanitizeGeminiSchema(schema["items"]);
	}
	if (schema.contains("properties") && schema["properties"].is_object()) {
		nlohmann::json properties = nlohmann::json::object();
		for (auto it = schema["properties"].begin(); it != schema["properties"].end(); ++it) {
			properties[it.key()] = SanitizeGeminiSchema(it.value());
		}
		out["properties"] = std::move(properties);
	}
	const auto copyCombinator = [&schema, &out](const char* key) {
		if (!schema.contains(key) || !schema[key].is_array()) {
			return;
		}
		nlohmann::json values = nlohmann::json::array();
		for (const auto& item : schema[key]) {
			values.push_back(SanitizeGeminiSchema(item));
		}
		out[key] = std::move(values);
	};
	copyCombinator("anyOf");
	copyCombinator("oneOf");
	copyCombinator("allOf");
	return out;
}

nlohmann::json BuildInternalToolCatalog(const AISettings& settings)
{
	return FilterInternalChatToolCatalog(BuildConfiguredToolCatalog(settings));
}

nlohmann::json BuildGeminiTools(const std::vector<AIChatMessage>&, bool, const AISettings& settings)
{
	const nlohmann::json catalog = BuildInternalToolCatalog(settings);
	nlohmann::json declarations = nlohmann::json::array();
	for (const auto& item : catalog) {
		if (!item.is_object()) {
			continue;
		}
		declarations.push_back({
			{"name", item.value("name", "")},
			{"description", TruncateGeminiDescription(item.value("description", ""))},
			{"parameters", item.contains("inputSchema") ? SanitizeGeminiSchema(item["inputSchema"]) : nlohmann::json::object()}
		});
	}
	return declarations.empty()
		? nlohmann::json::array()
		: nlohmann::json::array({ {{"functionDeclarations", declarations}} });
}

nlohmann::json BuildResponsesToolDefinitions(
	const AISettings& settings,
	const std::vector<AIChatMessage>&)
{
	const nlohmann::json catalog = BuildInternalToolCatalog(settings);
	nlohmann::json tools = nlohmann::json::array();
	for (const auto& item : catalog) {
		tools.push_back({
			{"type", "function"},
			{"name", item.value("name", "")},
			{"description", item.value("description", "")},
			{"parameters", item.contains("inputSchema") ? item["inputSchema"] : nlohmann::json::object()}
		});
	}
	return tools;
}

std::string BuildChatSystemPrompt(const AISettings& settings)
{
	std::string projectName;
	if (!AIService::Trim(g_nowOpenSourceFilePath).empty()) {
		try {
			const std::filesystem::path sourcePath(g_nowOpenSourceFilePath);
			projectName = sourcePath.stem().string();
			if (projectName.empty()) {
				projectName = sourcePath.filename().string();
			}
		}
		catch (...) {
			projectName.clear();
		}
	}

	const std::string projectType = DetectProjectTypeText();
	const bool mirrorSourceBase = settings.sourceEditMode == AISourceEditMode::MirrorSourceBase;
	const std::string sourceReadRule = mirrorSourceBase
		? "4) 已知子程序/代码项名称时优先用 read_code_item；未知位置时只做一次批量 search_code，再用一次 read_files 批量读取必要文件。写入工具以 read_files/read_code_item 的镜像文本和哈希为基准。\n"
		: "4) 已知子程序/代码项名称时优先用 read_code_item；未知位置时只做一次批量 search_code，再用一次 read_files 批量读取必要镜像；编辑当前工程源码前，再用 read_real_file 读取同一 file_path 的 IDE 真实页文本。\n";
	{
		std::string prompt =
			"你是AutoLinker，一个内置于易语言IDE的插件形式的助手。\n"
			"优先使用最少量的批量工具获取准确上下文，不要臆测当前页面、源码、模块、支持库或搜索结果。\n\n"
			"当前项目名称：" + (projectName.empty() ? std::string("未知") : projectName) + "\n\n"
			"当前项目类型：" + projectType + "\n\n"
			"统一源码工具规则：\n"
			"1) list_files / search_code / read_files / read_code_item 基于 e-packager 解包出的当前工程镜像，路径一律是镜像内相对路径，并返回 mirror_source。\n"
			"2) 内置 AI 对话在本轮开始前已经自动刷新工程镜像；不要调用 refresh_workspace_mirror，该显式工具只保留给外部 MCP 客户端。\n"
			"3) 普通单文件修改的探索预算最多 6 次只读调用；达到 4 次时立即收敛。独立查询应在同一响应中一次发出，多个文件必须使用 read_files，不要串行重复 read_file。\n"
			+ sourceReadRule +
			"5) 修改当前工程源码时只能用 edit_file / multi_edit_file / write_file / diff_file / restore_file_snapshot，并以 file_path 作为目标。\n"
			+ std::string(mirrorSourceBase
				? "6) edit_file / multi_edit_file / write_file / diff_file 的匹配和 expected_base_hash 校验基于 read_files/read_code_item 的镜像文本；大块修改优先一次生成 full_code 后调用 write_file，避免反复 exact old_text 失败。\n"
				  "7) 写入前不会读取真实页源码；写入仍会按 file_path 映射到 IDE 程序项并整页写回。\n"
				: "6) 当前工具列表提供 read_real_file：编辑 src/*.txt 或固定表文件前，必须先对同一 file_path 调用 read_real_file；edit_file / multi_edit_file / write_file / diff_file 的 old_text、full_code 和 expected_base_hash 应基于 read_real_file 返回的 real_source/code_hash，不要用镜像源码作为编辑基准。\n"
				  "7) edit_file / write_file 会把 file_path 映射到 IDE 真实程序项，基于真实页文本匹配，再整页写回。\n")
			+ "8) 写工具返回 ok=true、verified=true 后，写入和结构校验已经完成；禁止为了确认而再次读取同一源码。需要验证时按用户要求编译/测试，只有失败后才重新读取定位。\n"
			"9) src/*.xml 是窗口界面 XML，只读；窗口程序集代码应编辑对应 src/*.txt。\n"
			"10) ecom/、elib/、header/ 是依赖/公开信息参考，可读可搜但不可写。\n"
			"11) 固定表文件 src/.数据类型.txt、src/.DLL声明.txt、src/.常量.txt、src/.全局变量.txt 可作为对应真实表页的编辑目标。\n"
			"12) 需要预览改动用 diff_file；需要回滚最近写入用 restore_file_snapshot。\n"
			"13) 通常我们只读取常量，不编辑和写入常量值，因为会覆盖一些长文本常量无法正确覆盖，所以我们通常用固定的程序集变量或局部变量来写固定的值，但需要给与一些注释，不要看起来像是魔法数字或文本。\n"
			"14) 只有用户要求编译验证时，才调用compile_with_output_path。编译前可用 get_current_eide_info 确认 project_type 和可用编译模式。\n"
			"15) 除非用户明确要求搜索、刷新、列出、添加或移除模块/支持库，否则不要调用 refresh_dependency_catalog、search_available_modules、search_available_support_libraries、list_imported_modules、add_module_to_project、remove_module_from_project、add_support_library_to_project。\n\n"
			"其他工具：\n"
			"- 仅复杂、多文件或用户明确要求计划时使用 update_plan；局部单文件修改不要创建计划卡片。\n"
			"- 需要确认当前页名/页类型时用 get_current_page_info，不要臆测当前页。\n"
			"- 涉及联网、查文档、搜最新资料时用 search_web_tavily 搜索、extract_web_document 取正文、fetch_url 取原始响应。\n"
			"- 需要本地命令时用 run_powershell_command（会经用户确认后执行）。\n\n"
			"计划模式：\n"
			"- 如果上下文系统消息说明当前处于计划模式，只能探索、阅读、搜索和制定方案，不要写入文件、回滚、编译或执行 PowerShell。\n"
			"- 计划准备好时，必须用单独的 <proposed_plan>...</proposed_plan> 块提交方案，等待用户批准后再实施。\n"
			"- 用户批准计划后再按批准方案执行；若用户要求修改计划，先重新提交新的 <proposed_plan>。\n\n"
			"易语言基础约定：\n"
			"- 以 # 开头的标识通常表示常量；图片/音频等二进制资源也按常量资源引用，例如 #启动画面。\n"
			"- 以 & 开头通常表示对子程序取址，用于回调或传递函数地址，例如 到整数 (&枚举窗口过程)。\n"
			"- 以 . 开头的是易语言系统指令/关键字，例如 .版本、.程序集、.程序集变量、.子程序、.参数、.局部变量、.全局变量、.常量、.DLL声明、.数据类型、.成员、.如果、.如果真、.否则、.返回；编辑时不要删掉前导的 .，也不要改成 C/C++/JS 风格。\n"
			"- 单引号 ' 开头表示整行注释，不要把注释内容当成代码，也不要改成 // 或 /* */。\n"
			"- 真 / 假 是布尔值。\n"
			"- 数组下标通常从 1 开始，第一个元素是 数组 [1]，不要按多数语言习惯推导成从 0 开始。\n"
			"- .计次循环首 (次数, i) 中 i 通常从 1 递增到 次数；遍历数组常写 .变量循环首 (1, 取数组成员数 (数组), 1, i)，但其它合法起止范围也可能存在。\n"
			"- 赋值常写作 `变量 ＝ 值`，不要误写成半角 `=`。\n"
			"- 自增/自减通常写作 `a ＝ a ＋ 1`、`a ＝ a － 1`，不要写 a++、--a。\n"
			"- 易语言字符串中不支持转义序列，不要使用 \\n、\\t 等。\n"
			"- 易语言字符串可使用加号连接，例如 `\"Hello\" ＋ \"World\"`。\n"
			"- 返回常见写法是 `返回 (...)`。\n"
			"- 条件分支有两套常用结构：需要 else/双分支时用 `.如果 (条件)` + `.否则` + `.如果结束`；只有单分支时可用 `.如果真 (条件)` + `.如果真结束`。\n"
			"- 禁止把 `.如果真` 和 `.否则` 混用；`.如果真` 块中不能出现 `.否则`，也不能用 `.如果结束` 去闭合。需要否则分支时必须改用 `.如果 (条件)`。\n"
			"- 多路分支优先使用 `.判断开始` / `.判断` / `.默认` / `.判断结束`，不要写 switch/case/default，也不要用 C/C++/JS 的花括号分支。\n"
			"- 固定次数循环用 `.计次循环首 (次数, i)` / `.计次循环尾 ()`；范围循环用 `.变量循环首 (起始, 结束, 步长, i)` / `.变量循环尾 ()`；条件循环用 `.循环判断首 ()` / `.循环判断尾 (条件)`。\n"
			"- 循环内提前结束用 `跳出循环 ()`，跳到本轮循环末尾用 `到循环尾 ()`；不要写 break、continue、for、while、do while。\n"
			"- 子程序返回用 `返回 (...)`，不要写 return；流程条件不要写成 Python 冒号缩进、JS 花括号或 C 风格小语句块。\n"
			"- 流程控制必须精确配对闭合，不要漏掉 .如果结束、.如果真结束、.判断结束、.计次循环尾 ()、.变量循环尾 ()、.循环判断尾 (条件) 等结尾，也不要交叉嵌套闭合。\n"
			"- 普通程序集/窗口程序集的子程序名按全局解析，没有命名空间隔离同名；新增或重命名子程序时必须保证全工程唯一，避免与其它程序集重名。\n"
			"- 控件事件子程序（如 _按钮_Clear_被单击）依赖窗口界面与窗口程序集的名称绑定，必须留在所属窗口程序集页内，不要挪到普通程序集、其它窗口或类。\n"
			"- 易语言的类不支持方法覆盖/重写/多态，调用父类方法直接写 `父类方法名 ()`，没有 this./self./super.；类中无法实现单例，需要全局唯一实例应在 src/.全局变量.txt 声明该类类型的全局变量来访问。\n"
			"- { ... } 字面量本质是字节型数组，只能直接赋值给字节集变量；在函数参数、数组成员等表达式位置需要字节集时应显式写 到字节集 ({ ... })。\n"
			"- 全角中文标点和全角运算符在代码里较常见，分析与编辑时不要误判，也不要擅自替换成其它语言写法。\n"
			"- 只修改某个子程序时不要重写整个页面，也不要重复输出 .版本 2，保持原有缩进、空行与注释风格。\n\n"
			"工具失败时先分析失败原因并换更合适的工具，不要机械重试同一个调用；写工具一旦返回 ok=true、verified=true，就停止源码探索，不要复读或二次写回。\n";
		const std::string extraPrompt = AIService::Trim(settings.extraSystemPrompt);
		if (!extraPrompt.empty()) {
			prompt += "\n附加系统提示：\n";
			prompt += extraPrompt;
		}
		const std::string agentsMd = AIService::Trim(ReadProjectAgentsMd());
		if (!agentsMd.empty()) {
			prompt += "\n\n项目规范（来自 .AGENTS.md）：\n";
			prompt += agentsMd;
		}
		return prompt;
	}
}
std::string BuildGeminiChatSystemPrompt(const AISettings& settings, bool minimal)
{
	std::string prompt =
		"你是 AutoLinker 内置的易语言项目助手。\n"
		"回答要直接、准确，优先使用已提供的工具获取工程上下文。\n"
		"不要臆测当前页面、模块、支持库或源码内容。\n"
		"工程镜像已在本轮开始前自动刷新；已知代码项优先 read_code_item，多个文件使用 read_files，不要重复读取相同范围。\n"
		"仅复杂、多文件或明确要求计划时使用 update_plan；写入 verified=true 后不要为了确认而复读源码。\n"
		"如果需要读取网页或文档，优先调用 extract_web_document；需要原始响应时调用 fetch_url。\n"
		"除非用户明确要求搜索、刷新、列出、添加或移除模块/支持库，否则不要调用依赖管理工具。\n"
		"如果工具不可用或调用失败，说明限制并基于已有信息继续。\n"
		"只输出对用户有用的结果，不输出内部推理过程。\n";

	if (!minimal) {
		prompt +=
			"\n易语言要点：\n"
			"- .版本、.子程序、.参数、.局部变量、.如果 等是易语言指令。\n"
			"- 单引号 ' 开头表示注释；真/假 是布尔值。\n"
			"- 赋值常写作 `变量 ＝ 值`，不要写成 C/C++ 风格。\n"
			"- 修改代码前先读取真实源码；编译前先确认项目类型。\n";
	}

	const std::string extraPrompt = AIService::Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n附加系统提示：\n";
		prompt += extraPrompt.size() > 1200 ? extraPrompt.substr(0, 1200) : extraPrompt;
	}
	return prompt;
}

std::string UrlEncode(const std::string& value)
{
	static constexpr char kHex[] = "0123456789ABCDEF";
	std::string encoded;
	encoded.reserve(value.size() + 16);
	for (unsigned char c : value) {
		if ((c >= 'a' && c <= 'z') ||
			(c >= 'A' && c <= 'Z') ||
			(c >= '0' && c <= '9') ||
			c == '-' || c == '_' || c == '.' || c == '~') {
			encoded.push_back(static_cast<char>(c));
			continue;
		}
		encoded.push_back('%');
		encoded.push_back(kHex[(c >> 4) & 0x0F]);
		encoded.push_back(kHex[c & 0x0F]);
	}
	return encoded;
}

std::string AppendQueryParam(std::string url, const std::string& key, const std::string& value)
{
	if (key.empty()) {
		return url;
	}
	const char sep = (url.find('?') == std::string::npos) ? '?' : '&';
	url.push_back(sep);
	url += UrlEncode(key);
	url += "=";
	url += UrlEncode(value);
	return url;
}

std::string ReplaceSuffixIfPresent(const std::string& text, const std::string& oldSuffix, const std::string& newSuffix)
{
	if (!EndsWithInsensitive(text, oldSuffix)) {
		return text;
	}
	return text.substr(0, text.size() - oldSuffix.size()) + newSuffix;
}

bool EndsWithOpenAIVersionSegment(const std::string& text)
{
	const size_t slash = text.find_last_of('/');
	if (slash == std::string::npos || slash + 2 >= text.size()) {
		return false;
	}
	if (text[slash + 1] != 'v' && text[slash + 1] != 'V') {
		return false;
	}
	for (size_t i = slash + 2; i < text.size(); ++i) {
		const unsigned char ch = static_cast<unsigned char>(text[i]);
		if (std::isdigit(ch) == 0 && ch != '.') {
			return false;
		}
	}
	return true;
}

std::string BuildClaudeEndpoint(const std::string& baseUrl)
{
	std::string url = AIService::Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}
	if (EndsWithInsensitive(url, "/v1/messages")) {
		return url;
	}
	if (EndsWithInsensitive(url, "/v1")) {
		return url + "/messages";
	}
	return url + "/v1/messages";
}

std::string BuildOpenAIResponsesEndpoint(const std::string& baseUrl)
{
	std::string url = AIService::Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}

	url = ReplaceSuffixIfPresent(url, "/chat/completions", "/responses");
	if (EndsWithInsensitive(url, "/responses")) {
		return url;
	}
	if (EndsWithOpenAIVersionSegment(url)) {
		return url + "/responses";
	}
	return url + "/v1/responses";
}

std::string BuildGeminiEndpoint(const std::string& baseUrl, const std::string& model, bool stream)
{
	std::string url = AIService::Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}

	const std::string suffix = stream ? ":streamGenerateContent" : ":generateContent";
	const std::string otherSuffix = stream ? ":generateContent" : ":streamGenerateContent";

	url = ReplaceSuffixIfPresent(url, otherSuffix, suffix);
	if (EndsWithInsensitive(url, suffix)) {
		return stream ? AppendQueryParam(url, "alt", "sse") : url;
	}

	if (url.find("/models/") != std::string::npos) {
		url += suffix;
		return stream ? AppendQueryParam(url, "alt", "sse") : url;
	}

	if (EndsWithInsensitive(url, "/v1beta") || EndsWithInsensitive(url, "/v1")) {
		url += "/models/" + UrlEncode(model) + suffix;
		return stream ? AppendQueryParam(url, "alt", "sse") : url;
	}

	url += "/v1beta/models/" + UrlEncode(model) + suffix;
	return stream ? AppendQueryParam(url, "alt", "sse") : url;
}

std::string BuildOpenAIHeaders(const AISettings& settings)
{
	return BuildMergedHeaders({
		{ "Content-Type", "application/json" },
		{ "Authorization", "Bearer " + settings.apiKey }
	}, settings);
}

std::string BuildClaudeHeaders(const AISettings& settings)
{
	return BuildMergedHeaders({
		{ "Content-Type", "application/json" },
		{ "x-api-key", settings.apiKey },
		{ "anthropic-version", "2023-06-01" }
	}, settings);
}

std::string BuildJsonHeadersOnly(const AISettings& settings)
{
	return BuildMergedHeaders({
		{ "Content-Type", "application/json" }
	}, settings);
}

nlohmann::json BuildClaudeTools(
	const AISettings& settings,
	const std::vector<AIChatMessage>& contextMessages)
{
	nlohmann::json out = nlohmann::json::array();
	const nlohmann::json openAiTools = BuildChatToolDefinitions(settings, contextMessages);
	for (const auto& tool : openAiTools) {
		if (!tool.contains("function") || !tool["function"].is_object()) {
			continue;
		}
		const nlohmann::json& fn = tool["function"];
		out.push_back({
			{"name", fn.value("name", "")},
			{"description", fn.value("description", "")},
			{"input_schema", fn.value("parameters", nlohmann::json::object())}
		});
	}
	return out;
}

nlohmann::json BuildGeminiTools(const AISettings& settings)
{
	nlohmann::json declarations = nlohmann::json::array();
	const nlohmann::json openAiTools = BuildChatToolDefinitions(settings);
	for (const auto& tool : openAiTools) {
		if (!tool.contains("function") || !tool["function"].is_object()) {
			continue;
		}
		const nlohmann::json& fn = tool["function"];
		declarations.push_back({
			{"name", fn.value("name", "")},
			{"description", TruncateGeminiDescription(fn.value("description", ""))},
			{"parameters", SanitizeGeminiSchema(fn.value("parameters", nlohmann::json::object()))}
		});
	}
	return nlohmann::json::array({ {{"functionDeclarations", declarations}} });
}

std::string ParseErrorMessageUtf8(const nlohmann::json& parsed)
{
	if (!parsed.contains("error")) {
		return std::string();
	}
	const auto& errorNode = parsed["error"];
	if (errorNode.is_object() && errorNode.contains("message") && errorNode["message"].is_string()) {
		return errorNode["message"].get<std::string>();
	}
	if (errorNode.is_string()) {
		return errorNode.get<std::string>();
	}
	return std::string();
}

std::string ExtractClaudeTextUtf8(const nlohmann::json& parsed)
{
	if (!parsed.contains("content") || !parsed["content"].is_array()) {
		return std::string();
	}
	std::string textUtf8;
	for (const auto& item : parsed["content"]) {
		if (!item.is_object()) {
			continue;
		}
		if (item.value("type", std::string()) == "text" && item.contains("text") && item["text"].is_string()) {
			textUtf8 += item["text"].get<std::string>();
		}
	}
	return textUtf8;
}

std::string ExtractGeminiTextUtf8(const nlohmann::json& parsed)
{
	if (!parsed.contains("candidates") || !parsed["candidates"].is_array() || parsed["candidates"].empty()) {
		return std::string();
	}
	const auto& candidate = parsed["candidates"][0];
	if (!candidate.contains("content") || !candidate["content"].is_object()) {
		return std::string();
	}
	const auto& content = candidate["content"];
	if (!content.contains("parts") || !content["parts"].is_array()) {
		return std::string();
	}
	std::string textUtf8;
	for (const auto& part : content["parts"]) {
		if (part.is_object() && part.value("thought", false)) {
			continue;
		}
		if (part.is_object() && part.contains("text") && part["text"].is_string()) {
			textUtf8 += part["text"].get<std::string>();
		}
	}
	return textUtf8;
}

std::string ExtractResponsesTextUtf8(const nlohmann::json& parsed)
{
	if (parsed.contains("output_text") && parsed["output_text"].is_string()) {
		return parsed["output_text"].get<std::string>();
	}
	if (!parsed.contains("output") || !parsed["output"].is_array()) {
		return std::string();
	}

	std::string textUtf8;
	for (const auto& item : parsed["output"]) {
		if (!item.is_object() || item.value("type", std::string()) != "message") {
			continue;
		}
		if (!item.contains("content") || !item["content"].is_array()) {
			continue;
		}
		for (const auto& contentItem : item["content"]) {
			if (!contentItem.is_object()) {
				continue;
			}
			const std::string contentType = contentItem.value("type", std::string());
			if ((contentType == "output_text" || contentType == "text") &&
				contentItem.contains("text") &&
				contentItem["text"].is_string()) {
				textUtf8 += contentItem["text"].get<std::string>();
			}
		}
	}
	return textUtf8;
}

bool IsGeminiResourceExhaustedResponse(int statusCode, const std::string& responseBody)
{
	if (statusCode != 429 && statusCode != 500 && statusCode != 503) {
		return false;
	}
	const std::string lower = ToLowerAsciiCopy(responseBody);
	return lower.find("resource has been exhausted") != std::string::npos ||
		lower.find("quota") != std::string::npos ||
		lower.find("exhausted") != std::string::npos;
}

struct ClaudeToolCall {
	std::string id;
	std::string name;
	std::string argumentsUtf8;
};

struct GeminiToolCall {
	std::string name;
	std::string argumentsUtf8;
};

struct ResponsesToolCall {
	std::string itemId;
	std::string callId;
	std::string name;
	std::string argumentsUtf8;
};

std::vector<ClaudeToolCall> ExtractClaudeToolCalls(const nlohmann::json& parsed)
{
	std::vector<ClaudeToolCall> calls;
	if (!parsed.contains("content") || !parsed["content"].is_array()) {
		return calls;
	}

	for (const auto& item : parsed["content"]) {
		if (!item.is_object() || item.value("type", std::string()) != "tool_use") {
			continue;
		}
		ClaudeToolCall call;
		call.id = item.value("id", "");
		call.name = item.value("name", "");
		if (item.contains("input")) {
			call.argumentsUtf8 = item["input"].dump();
		}
		else {
			call.argumentsUtf8 = "{}";
		}
		calls.push_back(std::move(call));
	}
	return calls;
}

std::vector<GeminiToolCall> ExtractGeminiToolCalls(const nlohmann::json& parsed)
{
	std::vector<GeminiToolCall> calls;
	if (!parsed.contains("candidates") || !parsed["candidates"].is_array() || parsed["candidates"].empty()) {
		return calls;
	}
	const auto& candidate = parsed["candidates"][0];
	if (!candidate.contains("content") || !candidate["content"].is_object()) {
		return calls;
	}
	const auto& content = candidate["content"];
	if (!content.contains("parts") || !content["parts"].is_array()) {
		return calls;
	}

	for (const auto& part : content["parts"]) {
		if (!part.is_object() || !part.contains("functionCall") || !part["functionCall"].is_object()) {
			continue;
		}
		const auto& fn = part["functionCall"];
		GeminiToolCall call;
		call.name = fn.value("name", "");
		if (fn.contains("args")) {
			call.argumentsUtf8 = fn["args"].dump();
		}
		else {
			call.argumentsUtf8 = "{}";
		}
		calls.push_back(std::move(call));
	}
	return calls;
}

std::vector<ResponsesToolCall> ExtractResponsesToolCalls(const nlohmann::json& parsed)
{
	std::vector<ResponsesToolCall> calls;
	if (!parsed.contains("output") || !parsed["output"].is_array()) {
		return calls;
	}

	for (const auto& item : parsed["output"]) {
		if (!item.is_object() || item.value("type", std::string()) != "function_call") {
			continue;
		}

		ResponsesToolCall call;
		if (item.contains("id") && item["id"].is_string()) {
			call.itemId = item["id"].get<std::string>();
		}
		if (item.contains("call_id") && item["call_id"].is_string()) {
			call.callId = item["call_id"].get<std::string>();
		}
		if (item.contains("name") && item["name"].is_string()) {
			call.name = item["name"].get<std::string>();
		}
		if (item.contains("arguments") && item["arguments"].is_string()) {
			call.argumentsUtf8 = item["arguments"].get<std::string>();
		}
		calls.push_back(std::move(call));
	}
	return calls;
}

bool IsResponsesInputItem(const nlohmann::json& item)
{
	if (!item.is_object()) {
		return false;
	}
	const std::string type = item.value("type", std::string());
	if (type == "message" || type == "reasoning" || type == "function_call" || type == "function_call_output") {
		return true;
	}
	const std::string role = item.value("role", std::string());
	return type.empty() &&
		(role == "user" || role == "assistant" || role == "developer" || role == "system") &&
		item.contains("content");
}

bool HasNonEmptyJsonString(const nlohmann::json& item, const char* key)
{
	return item.contains(key) && item[key].is_string() && !item[key].get<std::string>().empty();
}

bool IsResponsesReasoningItemReusableStateless(const nlohmann::json& item)
{
	return item.is_object() &&
		item.value("type", std::string()) == "reasoning" &&
		HasNonEmptyJsonString(item, "encrypted_content");
}

void ClearResponsesItemIdForStatelessRequest(nlohmann::json& item)
{
	if (!item.is_object()) {
		return;
	}
	item.erase("id");
}

bool PrepareResponsesInputItemForStatelessRequest(nlohmann::json& item)
{
	if (!IsResponsesInputItem(item)) {
		return false;
	}

	const std::string type = item.value("type", std::string());
	if (type == "reasoning" && !IsResponsesReasoningItemReusableStateless(item)) {
		return false;
	}

	ClearResponsesItemIdForStatelessRequest(item);
	return true;
}

std::string GetResponsesCallId(const nlohmann::json& item)
{
	if (!item.is_object() || !item.contains("call_id") || !item["call_id"].is_string()) {
		return std::string();
	}
	return item["call_id"].get<std::string>();
}

void RemoveOrphanResponsesFunctionCallItems(nlohmann::json& input)
{
	if (!input.is_array()) {
		return;
	}

	std::unordered_set<std::string> functionCallIds;
	std::unordered_set<std::string> outputCallIds;
	for (const auto& item : input) {
		if (!item.is_object()) {
			continue;
		}
		const std::string type = item.value("type", std::string());
		const std::string callId = GetResponsesCallId(item);
		if (callId.empty()) {
			continue;
		}
		if (type == "function_call") {
			functionCallIds.insert(callId);
		}
		else if (type == "function_call_output") {
			outputCallIds.insert(callId);
		}
	}

	nlohmann::json filtered = nlohmann::json::array();
	for (auto& item : input) {
		if (!item.is_object()) {
			filtered.push_back(std::move(item));
			continue;
		}
		const std::string type = item.value("type", std::string());
		if (type != "function_call" && type != "function_call_output") {
			filtered.push_back(std::move(item));
			continue;
		}

		const std::string callId = GetResponsesCallId(item);
		const bool paired =
			!callId.empty() &&
			functionCallIds.find(callId) != functionCallIds.end() &&
			outputCallIds.find(callId) != outputCallIds.end();
		if (paired) {
			filtered.push_back(std::move(item));
		}
	}
	input = std::move(filtered);
}

void PrepareResponsesInputItemsForStatelessRequest(nlohmann::json& input)
{
	if (!input.is_array()) {
		return;
	}

	nlohmann::json prepared = nlohmann::json::array();
	for (auto item : input) {
		if (PrepareResponsesInputItemForStatelessRequest(item)) {
			prepared.push_back(std::move(item));
		}
	}
	RemoveOrphanResponsesFunctionCallItems(prepared);
	input = std::move(prepared);
}

bool TryGetResponsesPreviousResponseId(const nlohmann::json& item, std::string& outResponseId)
{
	if (!item.is_object()) {
		return false;
	}
	if (item.value("type", std::string()) != "previous_response_ref") {
		return false;
	}
	if (!item.contains("response_id") || !item["response_id"].is_string()) {
		return false;
	}
	outResponseId = item["response_id"].get<std::string>();
	return !outResponseId.empty();
}

std::string BuildResponsesInstructions(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings)
{
	std::string instructionsUtf8 = LocalToUtf8(BuildChatSystemPrompt(settings));
	for (const AIChatMessage& msg : contextMessages) {
		if (ToLowerAsciiCopy(AIService::Trim(msg.role)) != "system") {
			continue;
		}
		const std::string contentUtf8 = LocalToUtf8(msg.content);
		if (contentUtf8.empty()) {
			continue;
		}
		if (!instructionsUtf8.empty()) {
			instructionsUtf8 += "\n\n";
		}
		instructionsUtf8 += contentUtf8;
	}
	return instructionsUtf8;
}

nlohmann::json BuildResponsesTextMessage(const std::string& role, const std::string& textUtf8)
{
	const std::string contentType = ToLowerAsciiCopy(role) == "assistant"
		? "output_text"
		: "input_text";
	return {
		{"role", role},
		{"content", nlohmann::json::array({
			{
				{"type", contentType},
				{"text", textUtf8}
			}
		})}
	};
}

void AppendResponsesOutputItemsToInput(const nlohmann::json& parsed, nlohmann::json& input)
{
	if (!input.is_array() || !parsed.contains("output") || !parsed["output"].is_array()) {
		return;
	}
	for (const auto& item : parsed["output"]) {
		nlohmann::json inputItem = item;
		if (PrepareResponsesInputItemForStatelessRequest(inputItem)) {
			input.push_back(std::move(inputItem));
		}
	}
}

AIResult ExecuteTaskClaude(
	const std::string& systemPrompt,
	const std::string& inputText,
	const AISettings& settings,
	int maxRetryCount = kAiRequestRetryCount)
{
	AIResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	const std::string endpoint = BuildClaudeEndpoint(settings.baseUrl);

	nlohmann::json requestBody;
	requestBody["model"] = LocalToUtf8(settings.model);
	requestBody["max_tokens"] = 4096;
	requestBody["temperature"] = settings.temperature;
	requestBody["system"] = LocalToUtf8(systemPrompt);
	requestBody["messages"] = nlohmann::json::array({
		{
			{"role", "user"},
			{"content", nlohmann::json::array({ {{"type", "text"}, {"text", LocalToUtf8(inputText)}} })}
		}
	});
	ApplyThinkingConfigToClaudeRequest(requestBody, settings);
	NormalizeJsonStringsToUtf8InPlace(requestBody);

	const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
		endpoint,
		requestBody.dump(),
		BuildClaudeHeaders(settings),
		settings.timeoutMs,
		false,
		false,
		"claude-task",
		{},
		nullptr,
		maxRetryCount);
	result.httpStatus = statusCode;
	if (statusCode < 200 || statusCode >= 300) {
		LogAiHttpFailure("claude-task", statusCode, responseBody);
		result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}
		const std::string textUtf8 = ExtractClaudeTextUtf8(parsed);
		if (textUtf8.empty()) {
			result.error = "Claude response content is empty";
			return result;
		}
		result.ok = true;
		result.content = Utf8ToLocal(textUtf8);
		return result;
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse Claude response: ") + ex.what();
		return result;
	}
}

AIResult ExecuteTaskGemini(
	const std::string& systemPrompt,
	const std::string& inputText,
	const AISettings& settings,
	int maxRetryCount = kAiRequestRetryCount)
{
	AIResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	std::string endpoint = BuildGeminiEndpoint(settings.baseUrl, LocalToUtf8(settings.model), false);
	endpoint = AppendQueryParam(endpoint, "key", settings.apiKey);

	nlohmann::json requestBody;
	requestBody["system_instruction"] = {
		{"parts", nlohmann::json::array({ {{"text", LocalToUtf8(systemPrompt)}} })}
	};
	requestBody["generationConfig"] = { {"temperature", settings.temperature} };
	requestBody["contents"] = nlohmann::json::array({
		{
			{"role", "user"},
			{"parts", nlohmann::json::array({ {{"text", LocalToUtf8(inputText)}} })}
		}
	});
	ApplyThinkingConfigToGeminiRequest(requestBody, settings);
	NormalizeJsonStringsToUtf8InPlace(requestBody);

	const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
		endpoint,
		requestBody.dump(),
		BuildJsonHeadersOnly(settings),
		settings.timeoutMs,
		false,
		false,
		"gemini-task",
		{},
		nullptr,
		maxRetryCount);
	result.httpStatus = statusCode;
	if (statusCode < 200 || statusCode >= 300) {
		LogAiHttpFailure("gemini-task", statusCode, responseBody);
		result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}
		const std::string textUtf8 = ExtractGeminiTextUtf8(parsed);
		if (textUtf8.empty()) {
			result.error = "Gemini response content is empty";
			return result;
		}
		result.ok = true;
		result.content = Utf8ToLocal(textUtf8);
		return result;
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse Gemini response: ") + ex.what();
		return result;
	}
}

AIResult ExecuteTaskOpenAIResponses(
	const std::string& systemPrompt,
	const std::string& inputText,
	const AISettings& settings,
	int maxRetryCount = kAiRequestRetryCount)
{
	AIResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	const std::string endpoint = BuildOpenAIResponsesEndpoint(settings.baseUrl);

	nlohmann::json requestBody;
	requestBody["model"] = LocalToUtf8(settings.model);
	ApplyOpenAITemperatureIfSupported(requestBody, settings);
	requestBody["instructions"] = LocalToUtf8(systemPrompt);
	requestBody["input"] = nlohmann::json::array({
		BuildResponsesTextMessage("user", LocalToUtf8(inputText))
	});
	requestBody["stream"] = false;
	requestBody["store"] = false;
	ApplyThinkingConfigToOpenAIResponsesRequest(requestBody, settings);
	NormalizeJsonStringsToUtf8InPlace(requestBody);

	const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
		endpoint,
		requestBody.dump(),
		BuildOpenAIHeaders(settings),
		settings.timeoutMs,
		false,
		false,
		"openai-responses-task",
		{},
		nullptr,
		maxRetryCount);
	result.httpStatus = statusCode;
	if (statusCode < 200 || statusCode >= 300) {
		LogAiHttpFailure("openai-responses-task", statusCode, responseBody);
		result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
		return result;
	}

	nlohmann::json parsed;
	std::string streamTextUtf8;
	if (!TryParseResponsesResponseBody(responseBody, parsed, streamTextUtf8, result.error)) {
		return result;
	}

	const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
	if (!errUtf8.empty()) {
		result.error = Utf8ToLocal(errUtf8);
		return result;
	}
	std::string textUtf8 = ExtractResponsesTextUtf8(parsed);
	if (textUtf8.empty()) {
		textUtf8 = std::move(streamTextUtf8);
	}
	if (textUtf8.empty()) {
		result.error = "Responses API response content is empty";
		return result;
	}
	result.ok = true;
	result.content = Utf8ToLocal(textUtf8);
	return result;
}

AIChatResult ExecuteChatWithToolsClaude(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback,
	const std::function<bool()>& cancelCallback,
	HttpRequestCancellation* cancelContext)
{
	AIChatResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	const std::string endpoint = BuildClaudeEndpoint(settings.baseUrl);
	const nlohmann::json tools = BuildClaudeTools(settings, contextMessages);

	std::string systemUtf8 = LocalToUtf8(BuildChatSystemPrompt(settings));
	nlohmann::json messages = nlohmann::json::array();
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(AIService::Trim(msg.role));
		if (role == "system") {
			systemUtf8 += "\n\n";
			systemUtf8 += LocalToUtf8(msg.content);
			continue;
		}

		nlohmann::json rawMessage;
		if (TryParseRawChatMessageJson(msg.rawMessageJsonUtf8, rawMessage)) {
			std::string rawRole;
			if (rawMessage.contains("role") && rawMessage["role"].is_string()) {
				rawRole = ToLowerAsciiCopy(AIService::Trim(rawMessage["role"].get<std::string>()));
			}
			if ((rawRole == "user" || rawRole == "assistant") && rawMessage.contains("content")) {
				rawMessage["role"] = rawRole;
				messages.push_back(std::move(rawMessage));
				continue;
			}
			if (rawRole == "tool" && rawMessage.contains("content")) {
				messages.push_back({
					{"role", "user"},
					{"content", rawMessage["content"]}
				});
				continue;
			}
		}

		if (role != "user" && role != "assistant") {
			continue;
		}
		messages.push_back({
			{"role", role},
			{"content", nlohmann::json::array({
				{{"type", "text"}, {"text", LocalToUtf8(msg.content)}}
			})}
		});
	}

	const int maxToolRounds = kMaxToolRounds;
	AIChatToolPolicy::Session toolPolicy;
	for (int round = 0; round < maxToolRounds; ++round) {
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return MarkChatResultCancelled(std::move(result));
		}

		const AISettings roundSettings = BuildChatRoundSettings(settings, toolPolicy);
		nlohmann::json requestBody;
		requestBody["model"] = LocalToUtf8(settings.model);
		requestBody["max_tokens"] = 4096;
		requestBody["temperature"] = roundSettings.temperature;
		requestBody["system"] = systemUtf8;
		requestBody["messages"] = messages;
		requestBody["tools"] = tools;
		requestBody["tool_choice"] = { {"type", "auto"} };
		requestBody["stream"] = false;
		ApplyThinkingConfigToClaudeRequest(requestBody, roundSettings);
		NormalizeJsonStringsToUtf8InPlace(requestBody);

		const std::string requestBodyText = requestBody.dump();
		int attemptCount = 0;
		const auto roundStart = PerfClock::now();
		const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
			endpoint,
			requestBodyText,
			BuildClaudeHeaders(settings),
			settings.timeoutMs,
			false,
			false,
			"claude-chat",
			cancelCallback,
			cancelContext,
			kAiChatRequestRetryCount,
			&attemptCount);
		LogChatRoundMetrics(
			"claude-chat",
			round,
			requestBodyText.size(),
			ElapsedMs(roundStart),
			statusCode,
			attemptCount,
			toolPolicy.ExplorationCalls());
		result.httpStatus = statusCode;
		if (IsCancelRequested(cancelCallback, cancelContext) || statusCode == kAiRequestCancelledHttpStatus) {
			return MarkChatResultCancelled(std::move(result));
		}
		if (statusCode < 200 || statusCode >= 300) {
			LogAiHttpFailure("claude-chat", statusCode, responseBody);
			result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
			return result;
		}

		nlohmann::json parsed;
		try {
			parsed = nlohmann::json::parse(responseBody);
		}
		catch (const std::exception& ex) {
			result.error = std::string("Failed to parse Claude response: ") + ex.what();
			return result;
		}

		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}

		const std::vector<ClaudeToolCall> toolCalls = ExtractClaudeToolCalls(parsed);
		const std::string textUtf8 = ExtractClaudeTextUtf8(parsed);
		if (toolCalls.empty()) {
			if (textUtf8.empty()) {
				result.error = "Claude response content is empty";
				return result;
			}
			result.ok = true;
			result.content = Utf8ToLocal(textUtf8);
			if (parsed.contains("usage") && parsed["usage"].is_object()) {
				const auto& u = parsed["usage"];
				if (u.contains("input_tokens") && u["input_tokens"].is_number_integer()) {
					result.promptTokens = u["input_tokens"].get<int>();
				}
				const int out = (u.contains("output_tokens") && u["output_tokens"].is_number_integer())
					? u["output_tokens"].get<int>() : 0;
				result.totalTokens = result.promptTokens + out;
				result.hasUsage = true;
			}
			if (streamCallback) {
				streamCallback(result.content);
			}
			return result;
		}
		if (streamCallback && !textUtf8.empty()) {
			streamCallback(Utf8ToLocal(textUtf8));
		}

		if (parsed.contains("content") && parsed["content"].is_array()) {
			nlohmann::json assistantMessage = {
				{"role", "assistant"},
				{"content", parsed["content"]}
			};
			try {
				result.contextPrefixRawMessagesUtf8.push_back(assistantMessage.dump());
			}
			catch (...) {
			}
			messages.push_back(std::move(assistantMessage));
		}

		for (size_t i = 0; i < toolCalls.size(); ++i) {
			const ClaudeToolCall& call = toolCalls[i];
			const std::string callId = call.id.empty()
				? std::format("toolu_auto_{}_{}", round + 1, i + 1)
				: call.id;

			const ChatToolExecutionResult toolExecution = ExecuteChatToolWithPolicy(
				toolPolicy,
				toolCallback,
				call.name,
				call.argumentsUtf8);
			const bool toolOk = toolExecution.ok;
			const std::string& toolResultLocal = toolExecution.resultLocal;
			const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(call.name, toolResultLocal);

			AIChatToolEvent evt = {};
			evt.name = call.name;
			evt.argumentsJson = Utf8ToLocal(call.argumentsUtf8);
			evt.resultJson = toolResultLocal;
			evt.ok = toolOk;
			result.toolEvents.push_back(std::move(evt));
			if (IsCancelRequested(cancelCallback, cancelContext)) {
				return MarkChatResultCancelled(std::move(result));
			}

			nlohmann::json toolResultContent = nlohmann::json::array({
				{
					{"type", "tool_result"},
					{"tool_use_id", callId},
					{"content", compactPayload.textUtf8}
				}
			});
			nlohmann::json rawToolMessage = {
				{"role", "tool"},
				{"content", toolResultContent}
			};
			try {
				result.contextPrefixRawMessagesUtf8.push_back(rawToolMessage.dump());
			}
			catch (...) {
			}
			messages.push_back({
				{"role", "user"},
				{"content", std::move(toolResultContent)}
			});
		}
	}

	result.toolRoundsExceeded = true;
	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}

AIChatResult ExecuteChatWithToolsGemini(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback,
	const std::function<bool()>& cancelCallback,
	HttpRequestCancellation* cancelContext)
{
	AIChatResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	std::string endpoint = BuildGeminiEndpoint(settings.baseUrl, LocalToUtf8(settings.model), false);
	endpoint = AppendQueryParam(endpoint, "key", settings.apiKey);
	nlohmann::json tools = BuildGeminiTools(contextMessages, false, settings);

	bool degradedRequestMode = false;
	std::string systemUtf8 = LocalToUtf8(BuildGeminiChatSystemPrompt(settings, degradedRequestMode));
	nlohmann::json contents = nlohmann::json::array();
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(AIService::Trim(msg.role));
		if (role == "system") {
			systemUtf8 += "\n\n";
			systemUtf8 += LocalToUtf8(msg.content);
			continue;
		}
		if (role != "user" && role != "assistant") {
			continue;
		}
		contents.push_back({
			{"role", role == "assistant" ? "model" : "user"},
			{"parts", nlohmann::json::array({
				{{"text", LocalToUtf8(msg.content)}}
			})}
		});
	}

	const int maxToolRounds = kMaxToolRounds;
	AIChatToolPolicy::Session toolPolicy;
	for (int round = 0; round < maxToolRounds; ++round) {
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return MarkChatResultCancelled(std::move(result));
		}

		const AISettings roundSettings = BuildChatRoundSettings(settings, toolPolicy);
		nlohmann::json requestBody;
		requestBody["system_instruction"] = {
			{"parts", nlohmann::json::array({ {{"text", systemUtf8}} })}
		};
		requestBody["generationConfig"] = { {"temperature", roundSettings.temperature} };
		requestBody["contents"] = contents;
		if (tools.is_array() && !tools.empty()) {
			requestBody["tools"] = tools;
		}
		ApplyThinkingConfigToGeminiRequest(requestBody, roundSettings);
		NormalizeJsonStringsToUtf8InPlace(requestBody);

		const std::string requestBodyText = requestBody.dump();
		int attemptCount = 0;
		const auto roundStart = PerfClock::now();
		const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
			endpoint,
			requestBodyText,
			BuildJsonHeadersOnly(settings),
			settings.timeoutMs,
			false,
			false,
			"gemini-chat",
			cancelCallback,
			cancelContext,
			kAiChatRequestRetryCount,
			&attemptCount);
		LogChatRoundMetrics(
			"gemini-chat",
			round,
			requestBodyText.size(),
			ElapsedMs(roundStart),
			statusCode,
			attemptCount,
			toolPolicy.ExplorationCalls());
		result.httpStatus = statusCode;
		if (IsCancelRequested(cancelCallback, cancelContext) || statusCode == kAiRequestCancelledHttpStatus) {
			return MarkChatResultCancelled(std::move(result));
		}
		if (statusCode < 200 || statusCode >= 300) {
			if (!degradedRequestMode && IsGeminiResourceExhaustedResponse(statusCode, responseBody)) {
				degradedRequestMode = true;
				systemUtf8 = LocalToUtf8(BuildGeminiChatSystemPrompt(settings, degradedRequestMode));
				tools = BuildGeminiTools(contextMessages, true, settings);
				--round;
				continue;
			}
			LogAiHttpFailure("gemini-chat", statusCode, responseBody);
			result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
			return result;
		}

		nlohmann::json parsed;
		try {
			parsed = nlohmann::json::parse(responseBody);
		}
		catch (const std::exception& ex) {
			result.error = std::string("Failed to parse Gemini response: ") + ex.what();
			return result;
		}

		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}

		if (!parsed.contains("candidates") || !parsed["candidates"].is_array() || parsed["candidates"].empty()) {
			result.error = "Gemini response candidates is empty";
			return result;
		}

		const auto& candidate = parsed["candidates"][0];
		if (!candidate.contains("content") || !candidate["content"].is_object()) {
			result.error = "Gemini response content missing";
			return result;
		}
		const auto& candidateContent = candidate["content"];

		const std::vector<GeminiToolCall> toolCalls = ExtractGeminiToolCalls(parsed);
		const std::string textUtf8 = ExtractGeminiTextUtf8(parsed);
		if (toolCalls.empty()) {
			if (textUtf8.empty()) {
				result.error = "Gemini response content is empty";
				return result;
			}
			result.ok = true;
			result.content = Utf8ToLocal(textUtf8);
			if (parsed.contains("usageMetadata") && parsed["usageMetadata"].is_object()) {
				const auto& u = parsed["usageMetadata"];
				if (u.contains("promptTokenCount") && u["promptTokenCount"].is_number_integer()) {
					result.promptTokens = u["promptTokenCount"].get<int>();
				}
				if (u.contains("totalTokenCount") && u["totalTokenCount"].is_number_integer()) {
					result.totalTokens = u["totalTokenCount"].get<int>();
				}
				result.hasUsage = true;
			}
			if (streamCallback) {
				streamCallback(result.content);
			}
			return result;
		}
		if (streamCallback && !textUtf8.empty()) {
			streamCallback(Utf8ToLocal(textUtf8));
		}

		contents.push_back(candidateContent);

		for (const GeminiToolCall& call : toolCalls) {
			const ChatToolExecutionResult toolExecution = ExecuteChatToolWithPolicy(
				toolPolicy,
				toolCallback,
				call.name,
				call.argumentsUtf8);
			const bool toolOk = toolExecution.ok;
			const std::string& toolResultLocal = toolExecution.resultLocal;
			const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(call.name, toolResultLocal);

			AIChatToolEvent evt = {};
			evt.name = call.name;
			evt.argumentsJson = Utf8ToLocal(call.argumentsUtf8);
			evt.resultJson = toolResultLocal;
			evt.ok = toolOk;
			result.toolEvents.push_back(std::move(evt));
			if (IsCancelRequested(cancelCallback, cancelContext)) {
				return MarkChatResultCancelled(std::move(result));
			}

			contents.push_back({
				{"role", "user"},
				{"parts", nlohmann::json::array({
					{
						{"functionResponse", {
							{"name", call.name},
							{"response", compactPayload.jsonValue}
						}}
					}
				})}
			});
		}
	}

	result.toolRoundsExceeded = true;
	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}

AIChatResult ExecuteChatWithToolsOpenAIResponses(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback,
	const std::function<bool()>& cancelCallback,
	HttpRequestCancellation* cancelContext)
{
	AIChatResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	const std::string endpoint = BuildOpenAIResponsesEndpoint(settings.baseUrl);
	const nlohmann::json tools = BuildResponsesToolDefinitions(settings, contextMessages);

	nlohmann::json input = nlohmann::json::array();
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(AIService::Trim(msg.role));
		nlohmann::json rawInputItem;
		if (TryParseRawChatMessageJson(msg.rawMessageJsonUtf8, rawInputItem)) {
			std::string ignoredResponseId;
			if (TryGetResponsesPreviousResponseId(rawInputItem, ignoredResponseId)) {
				continue;
			}
			if (IsResponsesInputItem(rawInputItem)) {
				input.push_back(std::move(rawInputItem));
				continue;
			}
		}
		if (role != "user" && role != "assistant") {
			continue;
		}
		input.push_back(BuildResponsesTextMessage(role, LocalToUtf8(msg.content)));
	}

	const std::string instructionsUtf8 = BuildResponsesInstructions(contextMessages, settings);
	const int maxToolRounds = kMaxToolRounds;
	AIChatToolPolicy::Session toolPolicy;
	for (int round = 0; round < maxToolRounds; ++round) {
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return MarkChatResultCancelled(std::move(result));
		}

		const AISettings roundSettings = BuildChatRoundSettings(settings, toolPolicy);
		nlohmann::json requestBody;
		requestBody["model"] = LocalToUtf8(settings.model);
		ApplyOpenAITemperatureIfSupported(requestBody, roundSettings);
		requestBody["instructions"] = instructionsUtf8;
		nlohmann::json requestInput = input;
		PrepareResponsesInputItemsForStatelessRequest(requestInput);
		requestBody["input"] = std::move(requestInput);
		requestBody["tools"] = tools;
		requestBody["parallel_tool_calls"] = true;
		requestBody["stream"] = true;
		requestBody["store"] = false;
		ApplyThinkingConfigToOpenAIResponsesRequest(requestBody, roundSettings);
		NormalizeJsonStringsToUtf8InPlace(requestBody);

		const std::string requestBodyText = requestBody.dump();
		ResponsesStreamParseState streamState;
		int attemptCount = 0;
		const auto roundStart = PerfClock::now();
		const auto [responseBody, statusCode] = PerformPostRequestStreamingWithRetry(
			endpoint,
			requestBodyText,
			[&streamState, &streamCallback](const std::string& chunk) -> bool {
				return ConsumeResponsesStreamChunk(chunk, streamState, streamCallback);
			},
			BuildOpenAIHeaders(settings),
			GetChatRequestTimeoutMs(settings),
			false,
			false,
			"openai-responses-chat",
			cancelCallback,
			cancelContext,
			kAiChatRequestRetryCount,
			&attemptCount);
		LogChatRoundMetrics(
			"openai-responses-chat",
			round,
			requestBodyText.size(),
			ElapsedMs(roundStart),
			statusCode,
			attemptCount,
			toolPolicy.ExplorationCalls());
		result.httpStatus = statusCode;
		if (IsCancelRequested(cancelCallback, cancelContext) || statusCode == kAiRequestCancelledHttpStatus) {
			return MarkChatResultCancelled(std::move(result), Utf8ToLocal(streamState.mergedTextUtf8));
		}
		if (statusCode < 200 || statusCode >= 300) {
			LogAiHttpFailure("openai-responses-chat", statusCode, responseBody);
			result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
			return result;
		}

		if (!FlushResponsesStreamState(streamState, streamCallback)) {
			result.error = streamState.parseError.empty()
				? "Failed to parse Responses streaming response"
				: streamState.parseError;
			return result;
		}

		nlohmann::json parsed;
		if (streamState.sawSseEvent) {
			if (!streamState.parseError.empty()) {
				result.error = streamState.parseError;
				return result;
			}
			parsed = BuildResponsesParsedFromStream(streamState);
		}
		else {
			try {
				parsed = nlohmann::json::parse(responseBody);
			}
			catch (const std::exception& ex) {
				result.error = std::string("Failed to parse Responses API response: ") + ex.what();
				return result;
			}
		}

		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}

		const std::vector<ResponsesToolCall> toolCalls = ExtractResponsesToolCalls(parsed);
		std::string textUtf8 = ExtractResponsesTextUtf8(parsed);
		if (textUtf8.empty()) {
			textUtf8 = streamState.mergedTextUtf8;
		}
		if (toolCalls.empty()) {
			if (textUtf8.empty()) {
				result.error = "Responses API response content is empty";
				return result;
			}
			result.ok = true;
			result.content = Utf8ToLocal(textUtf8);
			{
				const nlohmann::json* up = nullptr;
				if (parsed.contains("response") && parsed["response"].is_object() &&
					parsed["response"].contains("usage") && parsed["response"]["usage"].is_object()) {
					up = &parsed["response"]["usage"];
				}
				else if (parsed.contains("usage") && parsed["usage"].is_object()) {
					up = &parsed["usage"];
				}
				if (up != nullptr) {
					const auto& u = *up;
					if (u.contains("input_tokens") && u["input_tokens"].is_number_integer()) {
						result.promptTokens = u["input_tokens"].get<int>();
					}
					if (u.contains("total_tokens") && u["total_tokens"].is_number_integer()) {
						result.totalTokens = u["total_tokens"].get<int>();
					}
					result.hasUsage = true;
				}
			}
			if (!streamState.sawSseEvent && streamCallback) {
				streamCallback(result.content);
			}
			return result;
		}
		if (!streamState.sawSseEvent && streamCallback && !textUtf8.empty()) {
			streamCallback(Utf8ToLocal(textUtf8));
		}

		AppendResponsesOutputItemsToInput(parsed, input);
		if (parsed.contains("output") && parsed["output"].is_array()) {
			for (const auto& item : parsed["output"]) {
				nlohmann::json contextItem = item;
				if (!PrepareResponsesInputItemForStatelessRequest(contextItem)) {
					continue;
				}
				try {
					result.contextPrefixRawMessagesUtf8.push_back(contextItem.dump());
				}
				catch (...) {
				}
			}
		}

		for (size_t i = 0; i < toolCalls.size(); ++i) {
			const ResponsesToolCall& call = toolCalls[i];
			const std::string callId = call.callId.empty()
				? std::format("call_auto_round{}_{}", round + 1, i + 1)
				: call.callId;

			const ChatToolExecutionResult toolExecution = ExecuteChatToolWithPolicy(
				toolPolicy,
				toolCallback,
				call.name,
				call.argumentsUtf8);
			const bool toolOk = toolExecution.ok;
			const std::string& toolResultLocal = toolExecution.resultLocal;
			const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(call.name, toolResultLocal);

			AIChatToolEvent evt = {};
			evt.name = call.name;
			evt.argumentsJson = Utf8ToLocal(call.argumentsUtf8);
			evt.resultJson = toolResultLocal;
			evt.ok = toolOk;
			result.toolEvents.push_back(std::move(evt));
			if (IsCancelRequested(cancelCallback, cancelContext)) {
				return MarkChatResultCancelled(std::move(result), Utf8ToLocal(textUtf8));
			}

			nlohmann::json toolOutputItem = {
				{"type", "function_call_output"},
				{"call_id", callId},
				{"output", compactPayload.textUtf8}
			};
			input.push_back(toolOutputItem);
			try {
				result.contextPrefixRawMessagesUtf8.push_back(toolOutputItem.dump());
			}
			catch (...) {
			}
		}
	}

	result.toolRoundsExceeded = true;
	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}
} // namespace

bool AIService::LoadSettings(AIJsonConfig& jsonConfig, ConfigManager* iniConfig, AISettings& outSettings)
{
	outSettings = {};

	// 若 JSON 无数据，尝试从 INI 迁移 AI 相关配置
	if (!jsonConfig.hasAnyData() && iniConfig != nullptr) {
		const std::string iniApiKey  = iniConfig->getValue("ai.api_key");
		const std::string iniBaseUrl = iniConfig->getValue("ai.base_url");
		if (!iniApiKey.empty() || !iniBaseUrl.empty()) {
			// INI 键名到 JSON 键名的映射（去掉 "ai." 前缀）
			const std::pair<const char*, const char*> mapping[] = {
				{ "protocol_type",      "ai.protocol_type"        },
				{ "thinking_level",     "ai.thinking_level"       },
				{ "base_url",           "ai.base_url"             },
				{ "api_key",            "ai.api_key"              },
				{ "model",              "ai.model"                },
				{ "system_prompt_extra","ai.system_prompt_extra"  },
				{ "custom_headers",     "ai.custom_headers"       },
				{ "timeout_ms",         "ai.timeout_ms"           },
				{ "temperature",        "ai.temperature"          },
				{ "context_window",     "ai.context_window"       },
			};
			std::map<std::string, std::string> toMigrate;
			for (const auto& [jsonKey, iniKey] : mapping) {
				const std::string val = iniConfig->getValue(iniKey);
				if (!val.empty()) {
					toMigrate[jsonKey] = val;
				}
			}
			if (!toMigrate.empty()) {
				jsonConfig.setValues(toMigrate);
			}
			std::map<std::string, std::string> globalToMigrate;
			if (const std::string val = iniConfig->getValue("ai.source_edit_mode"); !val.empty()) {
				globalToMigrate["source_edit_mode"] = val;
			}
			if (const std::string val = iniConfig->getValue("ai.tavily_api_key"); !val.empty()) {
				globalToMigrate["tavily_api_key"] = val;
			}
			if (!globalToMigrate.empty()) {
				jsonConfig.setGlobalValues(globalToMigrate);
			}
		}
	}

	// 从 JSON 读取设置（getValueLocal 将 UTF-8 转换为本地编码供 AISettings 使用）
	outSettings.protocolType     = ParseProtocolType(jsonConfig.getValue("protocol_type"));
	outSettings.thinkingLevel    = ParseThinkingLevel(jsonConfig.getValue("thinking_level"));
	outSettings.sourceEditMode   = ParseSourceEditMode(jsonConfig.getGlobalValue("source_edit_mode"));
	outSettings.baseUrl          = jsonConfig.getValueLocal("base_url");
	outSettings.apiKey           = jsonConfig.getValueLocal("api_key");
	outSettings.model            = jsonConfig.getValueLocal("model");
	outSettings.extraSystemPrompt= jsonConfig.getValueLocal("system_prompt_extra");
	outSettings.customHeadersText= jsonConfig.getValueLocal("custom_headers");
	outSettings.tavilyApiKey     = jsonConfig.getGlobalValueLocal("tavily_api_key");

	const std::string timeoutValue = jsonConfig.getValue("timeout_ms");
	if (!timeoutValue.empty()) {
		try {
			outSettings.timeoutMs = (std::max)(1000, std::stoi(timeoutValue));
		}
		catch (...) {
			outSettings.timeoutMs = 120000;
		}
	}

	const std::string temperatureValue = jsonConfig.getValue("temperature");
	if (!temperatureValue.empty()) {
		try {
			outSettings.temperature = std::stod(temperatureValue);
		}
		catch (...) {
			outSettings.temperature = 0.2;
		}
	}

	const std::string ctxWindowValue = jsonConfig.getValue("context_window");
	if (!ctxWindowValue.empty()) {
		try {
			outSettings.contextWindowTokens = (std::max)(0, std::stoi(ctxWindowValue));
		}
		catch (...) {
			outSettings.contextWindowTokens = 0;
		}
	}

	return true;
}

void AIService::SaveSettings(AIJsonConfig& jsonConfig, const AISettings& settings)
{
	jsonConfig.setValues({
		{ "protocol_type",       ProtocolTypeToString(settings.protocolType) },
		{ "thinking_level",      ThinkingLevelToString(settings.thinkingLevel) },
		{ "base_url",            settings.baseUrl                            },
		{ "api_key",             settings.apiKey                             },
		{ "model",               settings.model                              },
		{ "system_prompt_extra", settings.extraSystemPrompt                  },
		{ "custom_headers",      settings.customHeadersText                  },
		{ "timeout_ms",          std::to_string(settings.timeoutMs)          },
		{ "temperature",         std::format("{:.2f}", settings.temperature) },
		{ "context_window",      std::to_string(settings.contextWindowTokens) },
	});
	jsonConfig.removeValues({
		"source_edit_mode",
		"tavily_api_key"
	});
	jsonConfig.setGlobalValues({
		{ "source_edit_mode", SourceEditModeToString(settings.sourceEditMode) },
		{ "tavily_api_key",   settings.tavilyApiKey }
	});
}

bool AIService::HasRequiredSettings(const AISettings& settings, std::string& outMissingField)
{
	if (Trim(settings.baseUrl).empty()) {
		outMissingField = "baseUrl";
		return false;
	}
	if (Trim(settings.apiKey).empty()) {
		outMissingField = "apiKey";
		return false;
	}
	if (Trim(settings.model).empty()) {
		outMissingField = "model";
		return false;
	}
	outMissingField.clear();
	return true;
}

namespace {

// 在 model 中定位 family 前缀，并解析其后紧邻的小版本号。
// 例：prefix="gpt-5" 对 "gpt-5.6-codex" → outMinor=6；对 "gpt-5"/"gpt-5-codex" → outMinor=0。
// 主版本号已包含在 prefix 内（如 "opus-4"），这里只取其后的小版本。
bool ParseFamilyMinorVersion(const std::string& model, const char* prefix, int& outMinor)
{
	const std::string prefixStr(prefix);
	const size_t pos = model.find(prefixStr);
	if (pos == std::string::npos) {
		return false;
	}
	size_t i = pos + prefixStr.size();
	if (i < model.size() && (model[i] == '.' || model[i] == '-')) {
		++i; // 跳过 major 与 minor 间的一个分隔符（OpenAI 用 '.'，Claude 用 '-'）
	}
	int minor = 0;
	bool hasDigit = false;
	while (i < model.size() && model[i] >= '0' && model[i] <= '9') {
		minor = minor * 10 + (model[i] - '0');
		hasDigit = true;
		++i;
	}
	outMinor = hasDigit ? minor : 0;
	return true;
}

// 递增族：在已知小版本里取 version<=modelMinor 的最近条目窗口（carry-forward），
// 使未登记的更高小版本默认继承上一代（如 opus-4-9 继承 opus-4-8 的窗口，而非回落）。
struct FamilyMinorWindow { int minor; int window; };

// versions 须按 minor 升序。modelMinor 比所有已知都小则返回 false（交回子串表/默认）。
bool ResolveVersionedFamilyWindow(
	const std::string& model, const char* prefix,
	const FamilyMinorWindow* versions, size_t count, int& outWindow)
{
	int minor = 0;
	if (!ParseFamilyMinorVersion(model, prefix, minor)) {
		return false;
	}
	int best = -1;
	for (size_t k = 0; k < count; ++k) {
		if (versions[k].minor <= minor) {
			best = versions[k].window; // 升序遍历，最后一个满足 <= 的即最近版本
		}
	}
	if (best < 0) {
		return false;
	}
	outWindow = best;
	return true;
}

bool ResolveGpt5Window(const std::string& model, int& outWindow)
{
	int minor = 0;
	if (!ParseFamilyMinorVersion(model, "gpt-5", minor)) {
		return false;
	}

	const bool compactVariant =
		model.find("-mini") != std::string::npos ||
		model.find("-nano") != std::string::npos ||
		model.find("-codex") != std::string::npos;
	if (compactVariant) {
		outWindow = 400000;
		return true;
	}

	const bool proVariant = model.find("-pro") != std::string::npos;
	if (minor >= 6) {
		outWindow = 1050000;
		return true;
	}
	if (minor >= 5) {
		outWindow = proVariant ? 1050000 : 1000000;
		return true;
	}
	outWindow = minor >= 4 ? 1050000 : 400000;
	return true;
}

} // namespace

int AIService::ResolveContextWindowTokens(const AISettings& settings)
{
	if (settings.contextWindowTokens > 0) {
		return settings.contextWindowTokens; // P1: 用户配置
	}

	const std::string m = ToLowerAsciiCopy(settings.model);

	// P2a: OpenAI GPT-5 系列 —— 5.4 与 5.6 为 1.05M，5.5 主模型为 1M，mini/nano/codex 仍按 400K。
	int gpt5Window = 0;
	if (ResolveGpt5Window(m, gpt5Window)) {
		return gpt5Window;
	}

	// P2b: 版本递增族 —— 未登记的更高小版本继承上一代窗口（窗口值核实于 2026-07）。
	// versions 按 minor 升序，主版本号写在前缀里。
	struct VersionedFamily { const char* prefix; const FamilyMinorWindow* versions; size_t count; };
	static const FamilyMinorWindow kOpus4[]   = { { 0, 200000 }, { 1, 200000 }, { 6, 1000000 }, { 7, 1000000 }, { 8, 1000000 } };
	static const FamilyMinorWindow kSonnet4[] = { { 5, 200000 }, { 6, 1000000 } };
	static const FamilyMinorWindow kHaiku4[]  = { { 5, 200000 } };
	static const VersionedFamily kFamilies[] = {
		{ "opus-4",   kOpus4,   sizeof(kOpus4) / sizeof(kOpus4[0]) },
		{ "sonnet-4", kSonnet4, sizeof(kSonnet4) / sizeof(kSonnet4[0]) },
		{ "haiku-4",  kHaiku4,  sizeof(kHaiku4) / sizeof(kHaiku4[0]) },
	};
	for (const auto& fam : kFamilies) {
		int window = 0;
		if (ResolveVersionedFamilyWindow(m, fam.prefix, fam.versions, fam.count, window)) {
			return window;
		}
	}

	// P2c: 子串表 —— 不规则命名或非递增族。更具体的在前。
	struct Entry { const char* key; int window; };
	static const Entry kTable[] = {
		// OpenAI（gpt-5 系由上面的版本族处理）
		{ "gpt-4.1",        1047576 },
		{ "gpt-4o",          128000 },
		{ "gpt-4-turbo",     128000 },
		{ "o4",              200000 },
		{ "o3",              200000 },
		{ "o1",              200000 },
		{ "gpt-4",             8192 },
		{ "gpt-3.5",          16385 },
		// Anthropic Claude（opus/sonnet/haiku-4 系由版本族处理）
		{ "fable-5",        1000000 },
		{ "mythos-5",       1000000 },
		{ "sonnet-5",       1000000 },
		{ "claude",          200000 }, // 兜底：未识别的 Claude 一律按 200K
		// Google Gemini —— 1.5/2.x/3 普遍 1M+
		{ "gemini-1.5-pro", 2097152 },
		{ "gemini-1.5",     1048576 },
		{ "gemini-2.5",     1048576 },
		{ "gemini-2.0",     1048576 },
		{ "gemini-3",       1048576 },
		{ "gemini",         1048576 },
		// DeepSeek —— V4/官方 chat/reasoner 别名为 1M，旧 R1/V3.2 保持 128K
		{ "deepseek-v4",    1000000 },
		{ "deepseek-chat",  1000000 },
		{ "deepseek-reasoner", 1000000 },
		{ "deepseek-v3.2",   128000 },
		{ "deepseek-r1",     128000 },
		{ "deepseek",        128000 },
		// 智谱 GLM
		{ "glm-5.2",        1000000 },
		{ "glm-5.1",         200000 },
		{ "glm-5-turbo",     200000 },
		{ "glm-5",           200000 },
		{ "glm-4.7",         200000 },
		{ "glm-4.6",         200000 },
		{ "glm-4.5",         128000 },
		// 通义千问 Qwen
		{ "qwen3.7",        1000000 },
		{ "qwen3.6",        1000000 },
		{ "qwen3-coder-plus", 1000000 },
		{ "qwen3-coder-flash", 1000000 },
		{ "qwen3-coder-next", 262144 },
		{ "qwen3-coder",     262144 },
		{ "qwen3.5",         262144 },
		{ "qwen-plus",      1000000 },
		{ "qwen-max",        262144 },
		// Kimi / Moonshot
		{ "kimi-k2",         262144 },
		{ "kimi",            262144 },
		{ "moonshot-v1-128k", 131072 },
		{ "moonshot-v1-32k",   32768 },
		{ "moonshot-v1-8k",     8192 },
		// 豆包
		{ "doubao-seed-2.0", 262144 },
		{ "doubao-seed-1.8", 262144 },
		// MiniMax
		{ "minimax-m3",     1000000 },
		{ "minimax-m2",      196608 },
		{ "minimax",         196608 },
	};
	for (const auto& e : kTable) {
		if (m.find(e.key) != std::string::npos) {
			return e.window; // P2c: 子串表
		}
	}
	return 200000; // P3: 默认（保守）
}

AIThinkingLevel AIService::ParseThinkingLevel(const std::string& text)
{
	const std::string v = ToLowerAsciiCopy(Trim(text));
	if (v == "low") {
		return AIThinkingLevel::Low;
	}
	if (v == "medium" || v == "med") {
		return AIThinkingLevel::Medium;
	}
	if (v == "high") {
		return AIThinkingLevel::High;
	}
	if (v == "xhigh" || v == "extra_high" || v == "extra high") {
		return AIThinkingLevel::XHigh;
	}
	if (v == "max") {
		return AIThinkingLevel::Max;
	}
	if (v == "ultra") {
		return AIThinkingLevel::Ultra;
	}
	return AIThinkingLevel::Off;
}

std::string AIService::ThinkingLevelToString(AIThinkingLevel thinkingLevel)
{
	switch (thinkingLevel) {
	case AIThinkingLevel::Low:
		return "low";
	case AIThinkingLevel::Medium:
		return "medium";
	case AIThinkingLevel::High:
		return "high";
	case AIThinkingLevel::XHigh:
		return "xhigh";
	case AIThinkingLevel::Max:
		return "max";
	case AIThinkingLevel::Ultra:
		return "ultra";
	case AIThinkingLevel::Off:
	default:
		return "off";
	}
}

std::string AIService::ThinkingLevelDisplayName(AIThinkingLevel thinkingLevel)
{
	switch (thinkingLevel) {
	case AIThinkingLevel::Low:
		return "低";
	case AIThinkingLevel::Medium:
		return "中";
	case AIThinkingLevel::High:
		return "高";
	case AIThinkingLevel::XHigh:
		return "超高";
	case AIThinkingLevel::Max:
		return "最大";
	case AIThinkingLevel::Ultra:
		return "Ultra";
	case AIThinkingLevel::Off:
	default:
		return "关闭";
	}
}

AISourceEditMode AIService::ParseSourceEditMode(const std::string& text)
{
	const std::string v = ToLowerAsciiCopy(Trim(text));
	if (v == "mirror_source_base" ||
		v == "mirror_direct_write" ||
		v == "mirror" ||
		v == "mirror_first" ||
		v == "mirror_direct" ||
		v == "e_packager" ||
		v == "epackager") {
		return AISourceEditMode::MirrorSourceBase;
	}
	return AISourceEditMode::RealPageFirst;
}

std::string AIService::SourceEditModeToString(AISourceEditMode mode)
{
	switch (mode) {
	case AISourceEditMode::MirrorSourceBase:
		return "mirror_source_base";
	case AISourceEditMode::RealPageFirst:
	default:
		return "real_page_first";
	}
}

std::string AIService::SourceEditModeDisplayName(AISourceEditMode mode)
{
	switch (mode) {
	case AISourceEditMode::MirrorSourceBase:
		return "解包镜像基准（测试）";
	case AISourceEditMode::RealPageFirst:
	default:
		return "真实页优先";
	}
}

bool AIService::ValidateCustomHeadersText(const std::string& headerText, std::string& outError)
{
	std::vector<HttpHeaderEntry> headers;
	return ParseCustomHeadersTextInternal(headerText, headers, outError);
}

AIProtocolType AIService::ParseProtocolType(const std::string& text)
{
	const std::string v = ToLowerAsciiCopy(Trim(text));
	if (v == "openai_responses" || v == "openai-responses" || v == "openai responses" || v == "responses" || v == "openairesponses") {
		return AIProtocolType::OpenAIResponses;
	}
	if (v == "gemini") {
		return AIProtocolType::Gemini;
	}
	if (v == "claude") {
		return AIProtocolType::Claude;
	}
	return AIProtocolType::OpenAI;
}

std::string AIService::ProtocolTypeToString(AIProtocolType protocolType)
{
	switch (protocolType) {
	case AIProtocolType::OpenAIResponses:
		return "openai_responses";
	case AIProtocolType::Gemini:
		return "gemini";
	case AIProtocolType::Claude:
		return "claude";
	case AIProtocolType::OpenAI:
	default:
		return "openai";
	}
}

std::string AIService::ProtocolTypeDisplayName(AIProtocolType protocolType)
{
	switch (protocolType) {
	case AIProtocolType::OpenAIResponses:
		return "OpenAI Responses";
	case AIProtocolType::Gemini:
		return "Gemini";
	case AIProtocolType::Claude:
		return "Claude";
	case AIProtocolType::OpenAI:
	default:
		return "OpenAI";
	}
}

std::string AIService::BuildTaskDisplayName(AITaskKind kind)
{
	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		return "AI优化函数";
	case AITaskKind::AddCommentsToFunction:
		return "AI为当前函数添加注释";
	case AITaskKind::TranslateFunctionAndVariables:
		return "AI翻译当前函数+变量名";
	case AITaskKind::TranslateText:
		return "AI翻译选中文本";
	case AITaskKind::AddByCurrentPageType:
		return "AI按当前页类型添加代码";
	default:
		return "AI任务";
	}
}

AIResult AIService::TestConnection(const AISettings& settings)
{
	AIResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}

	const std::string systemPrompt = "你是一个 API 连通性测试助手。请只返回 OK。";
	const std::string inputText = "请只返回 OK。";
	if (settings.protocolType == AIProtocolType::Claude) {
		return ExecuteTaskClaude(systemPrompt, inputText, settings, 0);
	}
	if (settings.protocolType == AIProtocolType::Gemini) {
		return ExecuteTaskGemini(systemPrompt, inputText, settings, 0);
	}
	if (settings.protocolType == AIProtocolType::OpenAIResponses) {
		return ExecuteTaskOpenAIResponses(systemPrompt, inputText, settings, 0);
	}

	const std::string modelUtf8 = LocalToUtf8(settings.model);
	const std::string systemPromptUtf8 = LocalToUtf8(systemPrompt);
	const std::string inputTextUtf8 = LocalToUtf8(inputText);

	nlohmann::json requestBody;
	requestBody["model"] = modelUtf8;
	if (!IsOpenAIGpt5Model(settings.model)) {
		requestBody["temperature"] = 0;
	}
	requestBody["stream"] = false;
	requestBody["messages"] = nlohmann::json::array({
		{
			{"role", "system"},
			{"content", systemPromptUtf8}
		},
		{
			{"role", "user"},
			{"content", inputTextUtf8}
		}
	});
	ApplyThinkingConfigToOpenAIChatRequest(requestBody, settings);
	NormalizeJsonStringsToUtf8InPlace(requestBody);

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers = BuildOpenAIHeaders(settings);

	std::string requestBodyText;
	try {
		requestBodyText = requestBody.dump();
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to build AI request JSON: ") + ex.what();
		return result;
	}

	const auto [responseBody, statusCode] =
		PerformPostRequestWithRetry(
			endpoint,
			requestBodyText,
			headers,
			settings.timeoutMs,
			false,
			false,
			"openai-test",
			{},
			nullptr,
			0);
	result.httpStatus = statusCode;

	if (statusCode < 200 || statusCode >= 300) {
		LogAiHttpFailure("openai-test", statusCode, responseBody);
		result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		if (parsed.contains("choices") && parsed["choices"].is_array() && !parsed["choices"].empty()) {
			const nlohmann::json& choice = parsed["choices"][0];
			if (choice.contains("message") && choice["message"].contains("content")) {
				if (choice["message"]["content"].is_string()) {
					result.ok = true;
					result.content = Utf8ToLocal(choice["message"]["content"].get<std::string>());
					return result;
				}
				if (choice["message"]["content"].is_array()) {
					std::string merged;
					for (const auto& item : choice["message"]["content"]) {
						if (item.is_string()) {
							merged += item.get<std::string>();
							continue;
						}
						if (item.is_object() && item.contains("text") && item["text"].is_string()) {
							merged += item["text"].get<std::string>();
						}
					}
					if (!merged.empty()) {
						result.ok = true;
						result.content = Utf8ToLocal(merged);
						return result;
					}
				}
			}
		}

		if (parsed.contains("error") && parsed["error"].contains("message") && parsed["error"]["message"].is_string()) {
			result.error = Utf8ToLocal(parsed["error"]["message"].get<std::string>());
			return result;
		}

		result.error = "AI response does not match expected chat/completions schema";
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse AI response: ") + ex.what();
	}

	return result;
}

AIResult AIService::ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings)
{
	AIResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}

	const std::string systemPrompt = BuildSystemPrompt(kind, settings);
	if (settings.protocolType == AIProtocolType::Claude) {
		return ExecuteTaskClaude(systemPrompt, inputText, settings);
	}
	if (settings.protocolType == AIProtocolType::Gemini) {
		return ExecuteTaskGemini(systemPrompt, inputText, settings);
	}
	if (settings.protocolType == AIProtocolType::OpenAIResponses) {
		return ExecuteTaskOpenAIResponses(systemPrompt, inputText, settings);
	}

	const std::string modelUtf8 = LocalToUtf8(settings.model);
	const std::string systemPromptUtf8 = LocalToUtf8(systemPrompt);
	const std::string inputTextUtf8 = LocalToUtf8(inputText);

	nlohmann::json requestBody;
	requestBody["model"] = modelUtf8;
	ApplyOpenAITemperatureIfSupported(requestBody, settings);
	requestBody["stream"] = false;
	requestBody["messages"] = nlohmann::json::array({
		{
			{"role", "system"},
			{"content", systemPromptUtf8}
		},
		{
			{"role", "user"},
			{"content", inputTextUtf8}
		}
	});
	ApplyThinkingConfigToOpenAIChatRequest(requestBody, settings);
	NormalizeJsonStringsToUtf8InPlace(requestBody);

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers = BuildOpenAIHeaders(settings);

	std::string requestBodyText;
	try {
		requestBodyText = requestBody.dump();
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to build AI request JSON: ") + ex.what();
		return result;
	}

	const auto [responseBody, statusCode] =
		PerformPostRequestWithRetry(endpoint, requestBodyText, headers, settings.timeoutMs, false, false, "openai-task");
	result.httpStatus = statusCode;

	if (statusCode < 200 || statusCode >= 300) {
		LogAiHttpFailure("openai-task", statusCode, responseBody);
		result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		if (parsed.contains("choices") && parsed["choices"].is_array() && !parsed["choices"].empty()) {
			const nlohmann::json& choice = parsed["choices"][0];
			if (choice.contains("message") && choice["message"].contains("content")) {
				if (choice["message"]["content"].is_string()) {
					result.ok = true;
					result.content = Utf8ToLocal(choice["message"]["content"].get<std::string>());
					return result;
				}
				if (choice["message"]["content"].is_array()) {
					std::string merged;
					for (const auto& item : choice["message"]["content"]) {
						if (item.is_string()) {
							merged += item.get<std::string>();
							continue;
						}
						if (item.is_object() && item.contains("text") && item["text"].is_string()) {
							merged += item["text"].get<std::string>();
						}
					}
					if (!merged.empty()) {
						result.ok = true;
						result.content = Utf8ToLocal(merged);
						return result;
					}
				}
			}
		}

		if (parsed.contains("error") && parsed["error"].contains("message") && parsed["error"]["message"].is_string()) {
			result.error = Utf8ToLocal(parsed["error"]["message"].get<std::string>());
			return result;
		}

		result.error = "AI response does not match expected chat/completions schema";
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse AI response: ") + ex.what();
	}

	return result;
}

AIChatResult AIService::ExecuteChatWithTools(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback,
	const std::function<bool()>& cancelCallback,
	HttpRequestCancellation* cancelContext)
{
	AIChatResult result = {};
	std::string validationError;
	if (!ValidateRequestSettings(settings, validationError)) {
		result.error = validationError;
		return result;
	}
	if (IsCancelRequested(cancelCallback, cancelContext)) {
		return MarkChatResultCancelled(std::move(result));
	}

	if (settings.protocolType == AIProtocolType::Claude) {
		return ExecuteChatWithToolsClaude(contextMessages, settings, toolCallback, streamCallback, cancelCallback, cancelContext);
	}
	if (settings.protocolType == AIProtocolType::Gemini) {
		return ExecuteChatWithToolsGemini(contextMessages, settings, toolCallback, streamCallback, cancelCallback, cancelContext);
	}
	if (settings.protocolType == AIProtocolType::OpenAIResponses) {
		return ExecuteChatWithToolsOpenAIResponses(contextMessages, settings, toolCallback, streamCallback, cancelCallback, cancelContext);
	}

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers = BuildOpenAIHeaders(settings);
	const uint64_t traceId = GetCurrentAIPerfTraceId();

	nlohmann::json requestMessages = nlohmann::json::array();
	requestMessages.push_back({
		{"role", "system"},
		{"content", LocalToUtf8(BuildChatSystemPrompt(settings))}
	});
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(Trim(msg.role));
		if (role != "system" && role != "user" && role != "assistant" && role != "tool") {
			continue;
		}
		nlohmann::json requestMessage;
		if ((role == "assistant" || role == "tool") && TryParseRawChatMessageJson(msg.rawMessageJsonUtf8, requestMessage)) {
			requestMessage["role"] = role;
			if (!requestMessage.contains("content") || requestMessage["content"].is_null()) {
				requestMessage["content"] = LocalToUtf8(msg.content);
			}
			if (role == "assistant" &&
				!msg.reasoningContent.empty() &&
				(!requestMessage.contains("reasoning_content") || !requestMessage["reasoning_content"].is_string())) {
				requestMessage["reasoning_content"] = msg.reasoningContent;
			}
		}
		else {
			requestMessage = {
				{"role", role},
				{"content", LocalToUtf8(msg.content)}
			};
			if (role == "assistant" && !msg.reasoningContent.empty()) {
				requestMessage["reasoning_content"] = msg.reasoningContent;
			}
		}
		if (role == "assistant" && IsDeepSeekCompatibleSettings(settings)) {
			EnsureDeepSeekAssistantMessageCompat(requestMessage);
		}
		requestMessages.push_back(std::move(requestMessage));
	}
	const OpenAIChatToolSequenceRepairStats repairStats = RepairOpenAIChatToolMessageSequence(requestMessages);
	if (repairStats.removedMessages > 0) {
		Logger::Instance().Write(
			"AIService",
			std::format(
				"repaired OpenAI chat tool sequence removed_messages={} incomplete_groups={}",
				repairStats.removedMessages,
				repairStats.removedIncompleteGroups));
	}

	const nlohmann::json tools = BuildChatToolDefinitions(settings, contextMessages);
	const int maxToolRounds = kMaxToolRounds;
	AIChatToolPolicy::Session toolPolicy;

	for (int round = 0; round < maxToolRounds; ++round) {
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return MarkChatResultCancelled(std::move(result));
		}
		const AISettings roundSettings = BuildChatRoundSettings(settings, toolPolicy);
		nlohmann::json requestBody;
		requestBody["model"] = LocalToUtf8(settings.model);
		ApplyOpenAITemperatureIfSupported(requestBody, roundSettings);
		requestBody["stream"] = true;
		requestBody["stream_options"] = { {"include_usage", true} };
		requestBody["messages"] = requestMessages;
		requestBody["tools"] = tools;
		if (!IsDeepSeekCompatibleSettings(settings)) {
			requestBody["tool_choice"] = "auto";
			requestBody["parallel_tool_calls"] = true;
		}
		if (!ShouldSkipOpenAIChatReasoningForToolUse(roundSettings)) {
			ApplyThinkingConfigToOpenAIChatRequest(requestBody, roundSettings);
		}
		NormalizeJsonStringsToUtf8InPlace(requestBody);

		std::string requestBodyText;
		try {
			requestBodyText = requestBody.dump();
		}
		catch (const std::exception& ex) {
			result.error = std::string("Failed to build AI chat request JSON: ") + ex.what();
			return result;
		}

		ChatStreamParseState streamState;
		int attemptCount = 0;
		const auto networkStart = PerfClock::now();
		const auto [responseBody, statusCode] =
			PerformPostRequestStreamingWithRetry(
				endpoint,
				requestBodyText,
				[&streamState, &streamCallback](const std::string& chunk) -> bool {
					return ConsumeStreamChunk(chunk, streamState, streamCallback);
				},
				headers,
				GetChatRequestTimeoutMs(settings),
				false,
				false,
				"openai-chat",
				cancelCallback,
				cancelContext,
				kAiChatRequestRetryCount,
				&attemptCount);
		LogAIPerfCost(
			traceId,
			"AIService.ExecuteChat.network_total",
			ElapsedMs(networkStart),
			"http=" + std::to_string(statusCode) + " endpoint=" + endpoint);
		LogChatRoundMetrics(
			"openai-chat",
			round,
			requestBodyText.size(),
			ElapsedMs(networkStart),
			statusCode,
			attemptCount,
			toolPolicy.ExplorationCalls());
		result.httpStatus = statusCode;
		if (IsCancelRequested(cancelCallback, cancelContext) || statusCode == kAiRequestCancelledHttpStatus) {
			return MarkChatResultCancelled(std::move(result), Utf8ToLocal(streamState.mergedUtf8));
		}
		if (statusCode < 200 || statusCode >= 300) {
			LogAiHttpFailure("openai-chat", statusCode, responseBody);
			result.error = BuildHttpStatusErrorForUi(statusCode, responseBody);
			return result;
		}

		if (!FlushStreamParseState(streamState, streamCallback)) {
			result.error = streamState.parseError.empty() ? "Failed to parse AI streaming response" : streamState.parseError;
			return result;
		}
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return MarkChatResultCancelled(std::move(result), Utf8ToLocal(streamState.mergedUtf8));
		}

		nlohmann::json message;
		if (streamState.sawDataEvent) {
			if (!streamState.parseError.empty()) {
				result.error = streamState.parseError;
				return result;
			}
			message = BuildAssistantMessageFromStreamState(streamState);
		}
		else {
			nlohmann::json parsed;
			try {
				parsed = nlohmann::json::parse(responseBody);
			}
			catch (const std::exception& ex) {
				result.error = std::string("Failed to parse AI response: ") + ex.what();
				return result;
			}

			std::string parseError;
			if (!ExtractChatResponseMessage(parsed, message, parseError)) {
				if (parsed.contains("error") && parsed["error"].contains("message") && parsed["error"]["message"].is_string()) {
					result.error = Utf8ToLocal(parsed["error"]["message"].get<std::string>());
				}
				else {
					result.error = parseError.empty() ? "AI response parse failed" : parseError;
				}
				return result;
			}
		}

		// Tool-call path.
		if (message.contains("tool_calls") && message["tool_calls"].is_array() && !message["tool_calls"].empty()) {
			const std::string toolIntroUtf8 = MergeMessageContentUtf8(message);
			if (!streamState.sawDataEvent && streamCallback && !toolIntroUtf8.empty()) {
				streamCallback(Utf8ToLocal(toolIntroUtf8));
			}
			if (IsDeepSeekCompatibleSettings(settings)) {
				EnsureDeepSeekAssistantMessageCompat(message);
			}
			try {
				result.contextPrefixRawMessagesUtf8.push_back(message.dump());
			}
			catch (...) {
			}
			requestMessages.push_back(message);

			for (const auto& toolCall : message["tool_calls"]) {
				std::string callId;
				std::string toolName;
				std::string argsUtf8;
				if (toolCall.contains("id") && toolCall["id"].is_string()) {
					callId = toolCall["id"].get<std::string>();
				}
				if (toolCall.contains("function") && toolCall["function"].is_object()) {
					const auto& fn = toolCall["function"];
					if (fn.contains("name") && fn["name"].is_string()) {
						toolName = fn["name"].get<std::string>();
					}
					if (fn.contains("arguments") && fn["arguments"].is_string()) {
						argsUtf8 = fn["arguments"].get<std::string>();
					}
				}

				if (callId.empty()) {
					callId = std::format("call_auto_round{}_{}", round + 1, result.toolEvents.size() + 1);
				}

				const ChatToolExecutionResult toolExecution = ExecuteChatToolWithPolicy(
					toolPolicy,
					toolCallback,
					toolName,
					argsUtf8);
				const bool toolOk = toolExecution.ok;
				const std::string& toolResultLocal = toolExecution.resultLocal;
				const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(toolName, toolResultLocal);

				AIChatToolEvent evt = {};
				evt.name = toolName;
				evt.argumentsJson = Utf8ToLocal(argsUtf8);
				evt.resultJson = toolResultLocal;
				evt.ok = toolOk;
				result.toolEvents.push_back(std::move(evt));
				if (IsCancelRequested(cancelCallback, cancelContext)) {
					return MarkChatResultCancelled(std::move(result), Utf8ToLocal(streamState.mergedUtf8));
				}

				nlohmann::json toolMessage = {
					{"role", "tool"},
					{"tool_call_id", callId},
					{"name", toolName},
					{"content", compactPayload.textUtf8}
				};
				try {
					result.contextPrefixRawMessagesUtf8.push_back(toolMessage.dump());
				}
				catch (...) {
				}
				requestMessages.push_back(std::move(toolMessage));
			}
			continue;
		}

		// Final assistant content path.
		std::string mergedUtf8 = MergeMessageContentUtf8(message);
		if (!streamState.sawDataEvent && streamCallback && !mergedUtf8.empty()) {
			streamCallback(Utf8ToLocal(mergedUtf8));
		}
		if (mergedUtf8.empty()) {
			result.error = "AI response content is empty";
			return result;
		}
		if (IsCancelRequested(cancelCallback, cancelContext)) {
			return MarkChatResultCancelled(std::move(result), Utf8ToLocal(mergedUtf8));
		}

		result.ok = true;
		result.content = Utf8ToLocal(mergedUtf8);
		if (message.contains("reasoning_content") && message["reasoning_content"].is_string()) {
			result.reasoningContent = message["reasoning_content"].get<std::string>();
		}
		if (streamState.hasUsage) {
			result.hasUsage = true;
			result.promptTokens = streamState.promptTokens;
			result.totalTokens = streamState.totalTokens;
		}
		return result;
	}

	result.toolRoundsExceeded = true;
	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}

std::string AIService::BuildPublicToolCatalogJson()
{
	const nlohmann::json nativeCatalog = BuildPublicToolCatalog();
	return AIChatToolRegistry::FilterExternalPublicCatalog(nativeCatalog).dump();
}

std::string AIService::BuildAgentOptimizationSelfTestJson()
{
	nlohmann::json checks = nlohmann::json::array();
	bool allOk = true;

	{
		nlohmann::json messages = nlohmann::json::array({
			{{"role", "system"}, {"content", "system"}},
			{{"role", "tool"}, {"tool_call_id", "orphan"}, {"content", "bad"}},
			{{"role", "assistant"}, {"content", ""}, {"tool_calls", nlohmann::json::array({
				{{"id", "call_ok"}, {"type", "function"}, {"function", {{"name", "read_file"}, {"arguments", "{}"}}}}
			})}},
			{{"role", "tool"}, {"tool_call_id", "call_ok"}, {"content", "ok"}},
			{{"role", "assistant"}, {"content", ""}, {"tool_calls", nlohmann::json::array({
				{{"id", "call_missing_a"}, {"type", "function"}, {"function", {{"name", "read_file"}, {"arguments", "{}"}}}},
				{{"id", "call_missing_b"}, {"type", "function"}, {"function", {{"name", "read_file"}, {"arguments", "{}"}}}}
			})}},
			{{"role", "tool"}, {"tool_call_id", "call_missing_a"}, {"content", "partial"}},
			{{"role", "user"}, {"content", "continue"}}
		});
		const OpenAIChatToolSequenceRepairStats stats = RepairOpenAIChatToolMessageSequence(messages);
		const bool ok = messages.size() == 4 &&
			messages[0].value("role", std::string()) == "system" &&
			messages[1].value("role", std::string()) == "assistant" &&
			messages[2].value("role", std::string()) == "tool" &&
			messages[3].value("role", std::string()) == "user" &&
			messages[2].value("tool_call_id", std::string()) == "call_ok" &&
			stats.removedMessages == 3 &&
			stats.removedIncompleteGroups == 1;
		checks.push_back({
			{"name", "openai_chat_tool_sequence_repair"},
			{"ok", ok},
			{"remaining_messages", messages.size()},
			{"removed_messages", stats.removedMessages},
			{"removed_incomplete_groups", stats.removedIncompleteGroups}
		});
		allOk = allOk && ok;
	}

	{
		const std::array<std::pair<AIThinkingLevel, const char*>, 7> levels = {{
			{ AIThinkingLevel::Off, "off" },
			{ AIThinkingLevel::Low, "low" },
			{ AIThinkingLevel::Medium, "medium" },
			{ AIThinkingLevel::High, "high" },
			{ AIThinkingLevel::XHigh, "xhigh" },
			{ AIThinkingLevel::Max, "max" },
			{ AIThinkingLevel::Ultra, "ultra" },
		}};
		bool roundTripOk = true;
		for (const auto& [level, text] : levels) {
			roundTripOk = roundTripOk &&
				ThinkingLevelToString(level) == text &&
				ParseThinkingLevel(text) == level;
		}

		AISettings xhighSettings = {};
		xhighSettings.thinkingLevel = AIThinkingLevel::XHigh;
		nlohmann::json xhighRequest;
		ApplyThinkingConfigToOpenAIResponsesRequest(xhighRequest, xhighSettings);
		AISettings ultraSettings = {};
		ultraSettings.thinkingLevel = AIThinkingLevel::Ultra;
		nlohmann::json ultraRequest;
		ApplyThinkingConfigToOpenAIResponsesRequest(ultraRequest, ultraSettings);
		const bool requestMappingOk =
			xhighRequest["reasoning"].value("effort", std::string()) == "xhigh" &&
			ultraRequest["reasoning"].value("effort", std::string()) == "max";
		const bool ok = roundTripOk && requestMappingOk;
		checks.push_back({
			{"name", "gpt_5_6_reasoning_effort_presets"},
			{"ok", ok},
			{"xhigh_request_effort", xhighRequest["reasoning"].value("effort", std::string())},
			{"ultra_request_effort", ultraRequest["reasoning"].value("effort", std::string())}
		});
		allOk = allOk && ok;
	}

	{
		const std::array<const char*, 4> models = {
			"gpt-5.6", "gpt-5.6-sol", "gpt-5.6-terra", "gpt-5.6-luna"
		};
		bool ok = true;
		for (const char* model : models) {
			AISettings settings = {};
			settings.model = model;
			ok = ok && ResolveContextWindowTokens(settings) == 1050000;
		}
		checks.push_back({
			{"name", "gpt_5_6_context_window_presets"},
			{"ok", ok},
			{"context_window", 1050000}
		});
		allOk = allOk && ok;
	}

	{
		const std::string sourceContent(5000, 'x');
		nlohmann::json raw = {
			{"ok", true},
			{"file_path", "src/Test.txt"},
			{"content", sourceContent}
		};
		const CompactToolResultPayload compact = BuildCompactToolResultPayload(
			"read_file",
			Utf8ToLocal(raw.dump()));
		const size_t compactBytes = compact.jsonValue.value("content", std::string()).size();
		const bool ok = compactBytes == sourceContent.size();
		checks.push_back({
			{"name", "read_file_context_not_truncated_to_800_bytes"},
			{"ok", ok},
			{"content_bytes", compactBytes}
		});
		allOk = allOk && ok;
	}

	{
		const std::string sourceContent(40000, 'x');
		nlohmann::json raw = {
			{"ok", true},
			{"file_path", "src/Large.txt"},
			{"content", sourceContent}
		};
		const CompactToolResultPayload compact = BuildCompactToolResultPayload(
			"read_file",
			Utf8ToLocal(raw.dump()));
		const std::string excerpt = compact.jsonValue.value("content", std::string());
		const bool ok = excerpt.find("omitted UTF-8 byte range [") != std::string::npos &&
			excerpt.find("40000 total") != std::string::npos;
		checks.push_back({
			{"name", "context_excerpt_exact_omitted_byte_range"},
			{"ok", ok},
			{"excerpt_bytes", excerpt.size()}
		});
		allOk = allOk && ok;
	}

	{
		nlohmann::json files = nlohmann::json::array();
		for (int i = 0; i < 15; ++i) {
			files.push_back(std::format("src/File{:02}.txt", i));
		}
		nlohmann::json raw = {
			{"ok", true},
			{"files", std::move(files)},
			{"requested_offset", 10},
			{"offset", 10},
			{"returned", 15},
			{"total_results", 40},
			{"has_more", true},
			{"next_offset", 25},
			{"visible_result_range", {{"start_offset", 10}, {"end_offset_exclusive", 25}}}
		};
		const CompactToolResultPayload compact = BuildCompactToolResultPayload(
			"list_files",
			Utf8ToLocal(raw.dump()));
		const auto& visibleFiles = compact.jsonValue["files"];
		const bool ok = visibleFiles.is_array() &&
			visibleFiles.size() == 6 &&
			std::all_of(visibleFiles.begin(), visibleFiles.end(), [](const nlohmann::json& item) {
				return item.is_string();
			}) &&
			compact.jsonValue.value("returned", -1) == 6 &&
			compact.jsonValue.value("next_offset", -1) == 16 &&
			compact.jsonValue.value("has_more", false) &&
			compact.jsonValue["visible_result_range"].value("start_offset", -1) == 10 &&
			compact.jsonValue["visible_result_range"].value("end_offset_exclusive", -1) == 16 &&
			compact.jsonValue.value("context_omitted_items", 0u) == 9u;
		checks.push_back({
			{"name", "compacted_list_pagination_cursor"},
			{"ok", ok},
			{"visible_files", visibleFiles.size()},
			{"next_offset", compact.jsonValue.value("next_offset", -1)}
		});
		allOk = allOk && ok;
	}

	{
		nlohmann::json results = nlohmann::json::array();
		for (int i = 0; i < 25; ++i) {
			results.push_back({
				{"file_path", "src/Test.txt"},
				{"line_number", i + 1},
				{"text", std::format("match {}", i)}
			});
		}
		nlohmann::json raw = {
			{"ok", true},
			{"pattern", "match"},
			{"results", std::move(results)},
			{"requested_offset", 5},
			{"offset", 5},
			{"returned", 25},
			{"total_results", 60},
			{"has_more", true},
			{"next_offset", 30},
			{"page_limit", 200},
			{"truncated", true}
		};
		const CompactToolResultPayload compact = BuildCompactToolResultPayload(
			"search_code",
			Utf8ToLocal(raw.dump()));
		const auto& visibleResults = compact.jsonValue["results"];
		const bool ok = visibleResults.is_array() &&
			visibleResults.size() == 20 &&
			compact.jsonValue.value("returned", -1) == 20 &&
			compact.jsonValue.value("next_offset", -1) == 25 &&
			compact.jsonValue.value("page_limit", 0u) == 20u &&
			compact.jsonValue.value("source_page_limit", 0) == 200;
		checks.push_back({
			{"name", "compacted_search_pagination_cursor"},
			{"ok", ok},
			{"visible_results", visibleResults.size()},
			{"next_offset", compact.jsonValue.value("next_offset", -1)}
		});
		allOk = allOk && ok;
	}

	{
		const nlohmann::json catalog = BuildPublicToolCatalog();
		const auto readFilesIt = std::find_if(catalog.begin(), catalog.end(), [](const nlohmann::json& item) {
			return item.is_object() && item.value("name", std::string()) == "read_files";
		});
		bool ok = readFilesIt != catalog.end();
		if (ok) {
			const nlohmann::json properties = (*readFilesIt)["inputSchema"].value(
				"properties",
				nlohmann::json::object());
			ok = properties.contains("file_paths") &&
				properties["file_paths"].value("maxItems", 0) == 12 &&
				properties.contains("files") &&
				properties["files"].value("maxItems", 0) == 12;
		}
		checks.push_back({
			{"name", "read_files_context_batch_limit"},
			{"ok", ok},
			{"max_files", 12}
		});
		allOk = allOk && ok;
	}

	{
		ResponsesStreamParseState state;
		const std::string sse =
			"event: response.output_text.delta\n"
			"data: {\"type\":\"response.output_text.delta\",\"delta\":\"ok\"}\n\n"
			"event: response.output_item.done\n"
			"data: {\"type\":\"response.output_item.done\",\"item\":{\"id\":\"fc_1\",\"type\":\"function_call\",\"call_id\":\"call_1\",\"name\":\"read_files\",\"arguments\":\"{}\"}}\n\n"
			"event: response.completed\n"
			"data: {\"type\":\"response.completed\",\"response\":{\"output\":[{\"id\":\"fc_1\",\"type\":\"function_call\",\"call_id\":\"call_1\",\"name\":\"read_files\",\"arguments\":\"{}\"}],\"usage\":{\"input_tokens\":10,\"total_tokens\":12}}}\n\n";
		const size_t split = sse.size() / 2;
		const bool consumed = ConsumeResponsesStreamChunk(sse.substr(0, split), state, {}) &&
			ConsumeResponsesStreamChunk(sse.substr(split), state, {}) &&
			FlushResponsesStreamState(state, {});
		const nlohmann::json parsed = BuildResponsesParsedFromStream(state);
		const std::vector<ResponsesToolCall> calls = ExtractResponsesToolCalls(parsed);
		const bool ok = consumed &&
			state.parseError.empty() &&
			state.mergedTextUtf8 == "ok" &&
			calls.size() == 1 &&
			calls.front().name == "read_files";
		checks.push_back({
			{"name", "responses_stream_parser"},
			{"ok", ok},
			{"tool_calls", calls.size()},
			{"text", state.mergedTextUtf8},
			{"error", state.parseError}
		});
		allOk = allOk && ok;
	}

	{
		const std::string jsonBody =
			R"({"output":[{"type":"message","content":[{"type":"output_text","text":"json-ok"}]}]})";
		const std::string sseBody =
			"event: response.output_text.delta\n"
			"data: {\"type\":\"response.output_text.delta\",\"delta\":\"sse-ok\"}\n\n"
			"event: response.completed\n"
			"data: {\"type\":\"response.completed\",\"response\":{\"output\":[]}}\n\n";
		nlohmann::json jsonParsed;
		nlohmann::json sseParsed;
		std::string jsonStreamText;
		std::string sseStreamText;
		std::string jsonError;
		std::string sseError;
		const bool jsonOk = TryParseResponsesResponseBody(
			jsonBody,
			jsonParsed,
			jsonStreamText,
			jsonError);
		const bool sseOk = TryParseResponsesResponseBody(
			sseBody,
			sseParsed,
			sseStreamText,
			sseError);
		const bool ok = jsonOk &&
			ExtractResponsesTextUtf8(jsonParsed) == "json-ok" &&
			jsonStreamText.empty() &&
			sseOk &&
			sseParsed.is_object() &&
			sseStreamText == "sse-ok";
		checks.push_back({
			{"name", "responses_response_body_parser"},
			{"ok", ok},
			{"json_error", jsonError},
			{"sse_error", sseError},
			{"sse_text", sseStreamText}
		});
		allOk = allOk && ok;
	}

	{
		const nlohmann::json nativeCatalog = BuildPublicToolCatalog();
		const nlohmann::json realPageCatalog = FilterInternalChatToolCatalog(
			FilterToolCatalogForSourceEditMode(nativeCatalog, AISourceEditMode::RealPageFirst));
		const nlohmann::json mirrorCatalog = FilterInternalChatToolCatalog(
			FilterToolCatalogForSourceEditMode(nativeCatalog, AISourceEditMode::MirrorSourceBase));
		const auto contains = [](const nlohmann::json& catalog, const char* name) {
			return std::find_if(catalog.begin(), catalog.end(), [name](const nlohmann::json& item) {
				return item.is_object() && item.value("name", std::string()) == name;
			}) != catalog.end();
		};
		const std::array<const char*, 12> alwaysVisible = {{
			"update_plan",
			"read_file",
			"edit_file",
			"write_file",
			"refresh_dependency_catalog",
			"add_module_to_project",
			"compile_with_output_path",
			"run_powershell_command",
			"search_web_tavily",
			"fetch_url",
			"extract_web_document",
			"get_current_eide_info"
		}};
		bool ok = true;
		for (const char* name : alwaysVisible) {
			ok = ok && contains(realPageCatalog, name) && contains(mirrorCatalog, name);
		}
		ok = ok &&
			realPageCatalog.size() + 1 == nativeCatalog.size() &&
			mirrorCatalog.size() + 2 == nativeCatalog.size() &&
			contains(realPageCatalog, "read_real_file") &&
			!contains(mirrorCatalog, "read_real_file") &&
			!contains(realPageCatalog, "refresh_workspace_mirror") &&
			!contains(mirrorCatalog, "refresh_workspace_mirror");
		checks.push_back({
			{"name", "all_configured_tools_always_visible"},
			{"ok", ok},
			{"native_tool_count", nativeCatalog.size()},
			{"real_page_tool_count", realPageCatalog.size()},
			{"mirror_tool_count", mirrorCatalog.size()},
			{"real_page_read_visible", contains(realPageCatalog, "read_real_file")},
			{"mirror_real_page_read_hidden", !contains(mirrorCatalog, "read_real_file")}
		});
		allOk = allOk && ok;
	}

	{
		nlohmann::json catalog = nlohmann::json::array();
		for (int i = 0; i < 6; ++i) {
			const std::string originalName = i == 5
				? "decompile_function"
				: std::format("generic_tool_{}", i);
			catalog.push_back({
				{"name", std::format("mcp_ida_tool_{}", i)},
				{"description", "test"},
				{"inputSchema", nlohmann::json::object()},
				{"x_autolinker_mcp", {
					{"server_name", "IDA"},
					{"tool_name", originalName}
				}}
			});
		}
		const nlohmann::json visible = FilterInternalChatToolCatalog(catalog);
		const bool ok = visible.size() == catalog.size();
		checks.push_back({
			{"name", "all_external_mcp_tools_always_visible"},
			{"ok", ok},
			{"configured", catalog.size()},
			{"visible", visible.size()}
		});
		allOk = allOk && ok;
	}

	return nlohmann::json({
		{"name", "agent-optimization-self-test"},
		{"ok", allOk},
		{"checks", std::move(checks)}
	}).dump();
}

std::string AIService::NormalizeModelOutputToCode(const std::string& modelText)
{
	return RemoveCodeFence(modelText);
}

std::string AIService::Trim(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string AIService::BuildEndpoint(const std::string& baseUrl)
{
	std::string url = Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}
	url = ReplaceSuffixIfPresent(url, "/responses", "/chat/completions");
	if (EndsWithInsensitive(url, "/chat/completions")) {
		return url;
	}
	if (EndsWithOpenAIVersionSegment(url)) {
		return url + "/chat/completions";
	}
	return url + "/v1/chat/completions";
}

std::string AIService::BuildSystemPrompt(AITaskKind kind, const AISettings& settings)
{
	const std::string projectType = DetectProjectTypeText();
	std::string prompt =
		"你是一个易语言代码助手。\n"
		"当前项目类型：" + projectType + "\n"
		"规则：\n"
		"1) 代码注释使用单引号（'）。\n"
		"2) 必须遵循易语言语法与排版（如 .版本 2、.子程序 等）。\n"
		"3) 除非用户明确要求，否则只输出最终代码或文本，不要附加解释。\n"
		"4) 不要输出 Markdown 标题或其他包装。\n\n"
		"易语言格式示例：\n"
		".版本 2\n"
		".子程序 demo, 整数型\n"
		"返回 (0)\n\n";

	prompt += R"AL_REF(参考函数（完整示例，仅用于格式与风格参考，不要照抄无关变量）：
.版本 2

.子程序 FindWindowWithContainReturnMainList, 整数型, 公开, 函数注释：寻找所有符合条件的顶级窗口句柄，返回找到的数量
.参数 mainTitle, 文本型, , 参数；主窗口标题包含的文本
.参数 mainClass, 文本型, , 参数；主窗口类名包含的文本
.参数 childTitle, 文本型, , 参数；子窗口标题包含的文本
.参数 childClass, 文本型, , 参数；子窗口类名包含的文本
.参数 childHasChild, 逻辑型, 可空, 参数；可空；找到的子窗口需要包含子窗口(真)还是不包含(假)
.参数 mainHwndArray, 整数型, 参考 数组, 参数；参考（传址）；数组；用于存放找到的所有主窗口句柄的数组变量
.局部变量 i, 整数型, , , 这些都是变量
.局部变量 class, 文本型
.局部变量 text, 文本型
.局部变量 mid, 整数型, , "0", 这是个数组
.局部变量 x, 整数型
.局部变量 mainOK, 逻辑型
.局部变量 mainTitleOK, 逻辑型
.局部变量 mainClassOK, 逻辑型
.局部变量 childTitleOK, 逻辑型
.局部变量 childClassOK, 逻辑型
.局部变量 childOK, 逻辑型
.局部变量 hasChild, 逻辑型

' 1. 初始化结果数组
清除数组 (mainHwndArray)

' 2. 获取当前所有顶级窗口
清除数组 (m_hwnd_list)
EnumWindows (到整数 (&枚举窗口过程), 0)

' 3. 遍历顶级窗口
.计次循环首 (取数组成员数 (m_hwnd_list), i)
    text ＝ 窗口_取标题 (m_hwnd_list [i])
    class ＝ 窗口_取类名 (m_hwnd_list [i])

    ' 匹配主窗口条件 (空字符串视为匹配成功)
    mainTitleOK ＝ 选择 (mainTitle ＝ "", 真, IsContains (text, mainTitle))
    mainClassOK ＝ 选择 (mainClass ＝ "", 真, IsContains (class, mainClass))

    mainOK ＝ mainTitleOK 且 mainClassOK

    .如果真 (mainOK)
        ' 4. 如果主窗口匹配，枚举其所有子窗口进行深度检查
        清除数组 (mid)
        窗口_枚举所有子窗口 (m_hwnd_list [i], mid, )

        .计次循环首 (取数组成员数 (mid), x)
            text ＝ 窗口_取标题 (mid [x])
            class ＝ 窗口_取类名 (mid [x])

            ' 匹配子窗口条件
            childTitleOK ＝ 选择 (childTitle ＝ "", 真, IsContains (text, childTitle))
            childClassOK ＝ 选择 (childClass ＝ "", 真, IsContains (class, childClass))
            childOK ＝ childTitleOK 且 childClassOK

            ' 检查子窗口是否含有孙窗口
            hasChild ＝ hasChildWindow (mid [x])

            ' 综合判断子窗口是否符合要求
            .如果 (childHasChild)
                .如果真 (childOK 且 hasChild)
                    加入成员 (mainHwndArray, m_hwnd_list [i])
                    跳出循环 ()  ' 只要找到一个符合条件的子窗口，该主窗口就合格，跳出子窗口循环
                .如果真结束

            .否则
                .如果真 (childOK 且 hasChild ＝ 假)
                    加入成员 (mainHwndArray, m_hwnd_list [i])
                    跳出循环 ()
                .如果真结束

            .如果结束

        .计次循环尾 ()
    .如果真结束

.计次循环尾 ()

' 5. 返回找到的总数
返回 (取数组成员数 (mainHwndArray))

)AL_REF";

	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		prompt += "任务：优化给定函数代码，在保持行为等价的前提下提升可读性与健壮性。只返回完整可替换的函数代码。";
		break;
	case AITaskKind::AddCommentsToFunction:
		prompt += "任务：为给定函数添加合适注释（函数说明与关键行注释），不得改变原逻辑。只返回完整可替换的函数代码。";
		break;
	case AITaskKind::TranslateFunctionAndVariables:
		prompt +=
			"任务：将函数名、参数名、局部变量名翻译或重命名为英文 lowerCamelCase（首字母小写），并保持逻辑不变。\n"
			"禁止翻译或修改任何以 '.' 开头的易语言系统指令/关键字（例如：.版本/.子程序/.参数/.局部变量/.如果/.否则/.返回 等）。\n"
			"只允许修改标识符（函数名/参数名/局部变量名），系统指令与语句结构必须保持不变。\n"
			"只返回完整可替换的函数代码。";
		break;
	case AITaskKind::TranslateText:
		prompt += "任务：翻译用户提供的文本。只返回翻译后的纯文本，不要附加解释。";
		break;
	case AITaskKind::AddByCurrentPageType:
		prompt +=
			"任务：根据‘当前页类型 + 用户需求 + 当前页伪代码’生成一段可直接追加的代码片段。\n\n"
			"额外要求：\n"
			"1) 只输出要追加的代码，禁止重复整个页面。\n"
			"2) 不要输出解释或 Markdown 包装。\n"
			"3) 不要重复输出 .版本 行。\n"
			"4) 必须遵守当前页面已有结构和书写风格。\n"
			"5) 生成结果必须能直接粘贴到当前页面末尾。\n\n"
			"输出时保持原有换行与缩进，不要把多行代码压成一行。";
		break;
	default:
		break;
	}

	const std::string extraPrompt = Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n\n用户额外系统提示：\n";
		prompt += extraPrompt;
	}

	const std::string agentsMd = Trim(ReadProjectAgentsMd());
	if (!agentsMd.empty()) {
		prompt += "\n\n项目规范（来自 .AGENTS.md）：\n";
		prompt += agentsMd;
	}

	return prompt;
}

