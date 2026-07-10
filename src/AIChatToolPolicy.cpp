#include "AIChatToolPolicy.h"

#include <algorithm>
#include <cctype>
#include <limits>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "..\\thirdparty\\json.hpp"

namespace AIChatToolPolicy {
namespace {

using json = nlohmann::json;

struct ReadCoverage {
	std::string filePath;
	int begin = 0;
	int end = 0;
};

struct ReadRequestRange {
	std::string filePath;
	std::string normalizedPath;
	int begin = 0;
	int end = 0;
};

constexpr int kDefaultReadFileLimit = 2000;
constexpr int kDefaultReadFilesLimit = 1200;
constexpr int kMaxReadFilesRequests = 24;

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string NormalizePath(std::string path)
{
	std::replace(path.begin(), path.end(), '\\', '/');
	return ToLowerAscii(std::move(path));
}

std::string CanonicalArguments(const std::string& argumentsJsonUtf8)
{
	try {
		const json parsed = argumentsJsonUtf8.empty() ? json::object() : json::parse(argumentsJsonUtf8);
		return parsed.dump();
	}
	catch (...) {
		return argumentsJsonUtf8;
	}
}

bool IsSourceExplorationTool(const std::string& toolName)
{
	return toolName == "list_files" ||
		toolName == "search_code" ||
		toolName == "read_file" ||
		toolName == "read_files" ||
		toolName == "read_code_item";
}

bool IsSourceReadTool(const std::string& toolName)
{
	return IsSourceExplorationTool(toolName) || toolName == "read_real_file";
}

bool IsWriteTool(const std::string& toolName)
{
	return toolName == "edit_file" ||
		toolName == "multi_edit_file" ||
		toolName == "write_file" ||
		toolName == "restore_file_snapshot";
}

bool IsCompileTool(const std::string& toolName)
{
	return toolName == "compile_with_output_path";
}

std::string BuildSignature(const std::string& toolName, const std::string& argumentsJsonUtf8)
{
	return toolName + "\n" + CanonicalArguments(argumentsJsonUtf8);
}

int SafeRangeEnd(int begin, int limit)
{
	const long long value = static_cast<long long>((std::max)(0, begin)) + (std::max)(1, limit);
	return value > (std::numeric_limits<int>::max)()
		? (std::numeric_limits<int>::max)()
		: static_cast<int>(value);
}

int GetInteger(const json& value, const char* key, int defaultValue)
{
	if (!value.is_object() || !value.contains(key) || !value[key].is_number_integer()) {
		return defaultValue;
	}
	return value[key].get<int>();
}

void AddReadRequestRange(
	std::vector<ReadRequestRange>& requests,
	const std::string& filePath,
	int offset,
	int limit,
	bool deduplicateByFile)
{
	const std::string normalized = NormalizePath(filePath);
	if (normalized.empty()) {
		return;
	}
	if (deduplicateByFile) {
		const int normalizedBegin = (std::max)(0, offset);
		const int normalizedEnd = SafeRangeEnd(normalizedBegin, limit);
		const auto duplicate = std::find_if(
			requests.begin(),
			requests.end(),
			[&normalized, normalizedBegin, normalizedEnd](const ReadRequestRange& item) {
				return item.normalizedPath == normalized &&
					item.begin == normalizedBegin &&
					item.end == normalizedEnd;
			});
		if (duplicate != requests.end()) {
			return;
		}
	}
	if (static_cast<int>(requests.size()) >= kMaxReadFilesRequests) {
		return;
	}
	ReadRequestRange request;
	request.filePath = filePath;
	request.normalizedPath = normalized;
	request.begin = (std::max)(0, offset);
	request.end = SafeRangeEnd(request.begin, limit);
	requests.push_back(std::move(request));
}

std::vector<ReadRequestRange> ParseReadRequests(
	const std::string& toolName,
	const std::string& argumentsJsonUtf8)
{
	std::vector<ReadRequestRange> requests;
	try {
		const json args = argumentsJsonUtf8.empty() ? json::object() : json::parse(argumentsJsonUtf8);
		if (toolName == "read_file") {
			if (args.contains("file_path") && args["file_path"].is_string()) {
				AddReadRequestRange(
					requests,
					args["file_path"].get<std::string>(),
					GetInteger(args, "offset", 0),
					GetInteger(args, "limit", kDefaultReadFileLimit),
					false);
			}
			return requests;
		}
		if (toolName != "read_files") {
			return requests;
		}

		const int defaultOffset = (std::max)(0, GetInteger(args, "offset", 0));
		const int defaultLimit = (std::max)(1, GetInteger(args, "limit", kDefaultReadFilesLimit));
		if (args.contains("file_paths") && args["file_paths"].is_array()) {
			for (const auto& item : args["file_paths"]) {
				if (item.is_string()) {
					AddReadRequestRange(
						requests,
						item.get<std::string>(),
						defaultOffset,
						defaultLimit,
						true);
				}
			}
		}
		if (args.contains("files") && args["files"].is_array()) {
			for (const auto& item : args["files"]) {
				if (item.is_string()) {
					AddReadRequestRange(
						requests,
						item.get<std::string>(),
						defaultOffset,
						defaultLimit,
						true);
				}
				else if (item.is_object() && item.contains("file_path") && item["file_path"].is_string()) {
					AddReadRequestRange(
						requests,
						item["file_path"].get<std::string>(),
						GetInteger(item, "offset", defaultOffset),
						GetInteger(item, "limit", defaultLimit),
						true);
				}
			}
		}
	}
	catch (...) {
		requests.clear();
	}
	return requests;
}

bool RangesOverlap(int leftBegin, int leftEnd, int rightBegin, int rightEnd)
{
	return leftBegin < rightEnd && rightBegin < leftEnd;
}

std::vector<std::pair<int, int>> BuildMissingRanges(
	const ReadRequestRange& request,
	const std::vector<ReadCoverage>& coverage)
{
	std::vector<std::pair<int, int>> covered;
	for (const ReadCoverage& item : coverage) {
		if (item.filePath != request.normalizedPath ||
			!RangesOverlap(request.begin, request.end, item.begin, item.end)) {
			continue;
		}
		covered.push_back({
			(std::max)(request.begin, item.begin),
			(std::min)(request.end, item.end)
		});
	}
	std::sort(covered.begin(), covered.end());

	std::vector<std::pair<int, int>> missing;
	int cursor = request.begin;
	for (const auto& item : covered) {
		if (item.first > cursor) {
			missing.push_back({cursor, item.first});
		}
		cursor = (std::max)(cursor, item.second);
		if (cursor >= request.end) {
			break;
		}
	}
	if (cursor < request.end) {
		missing.push_back({cursor, request.end});
	}
	return missing;
}

bool HasCoveredOverlap(
	const ReadRequestRange& request,
	const std::vector<ReadCoverage>& coverage)
{
	return std::any_of(
		coverage.begin(),
		coverage.end(),
		[&request](const ReadCoverage& item) {
			return item.filePath == request.normalizedPath &&
				RangesOverlap(request.begin, request.end, item.begin, item.end);
		});
}

void CoalesceReadCoverage(std::vector<ReadCoverage>& coverage)
{
	std::sort(coverage.begin(), coverage.end(), [](const ReadCoverage& left, const ReadCoverage& right) {
		if (left.filePath != right.filePath) {
			return left.filePath < right.filePath;
		}
		if (left.begin != right.begin) {
			return left.begin < right.begin;
		}
		return left.end < right.end;
	});
	std::vector<ReadCoverage> merged;
	for (const ReadCoverage& item : coverage) {
		if (!merged.empty() &&
			merged.back().filePath == item.filePath &&
			item.begin <= merged.back().end) {
			merged.back().end = (std::max)(merged.back().end, item.end);
			continue;
		}
		merged.push_back(item);
	}
	coverage = std::move(merged);
}

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

int CountCompleteNumberedLines(const std::string& content, size_t begin, size_t end)
{
	int count = 0;
	for (size_t i = begin; i < end && i < content.size(); ++i) {
		if (content[i] == '\n') {
			++count;
		}
	}
	return count;
}

std::vector<std::pair<int, int>> BuildVisibleCoverageRanges(
	const std::string& content,
	int beginLineOffset,
	int returnedLines,
	size_t contextByteLimit,
	size_t tailBytes)
{
	std::vector<std::pair<int, int>> ranges;
	if (returnedLines <= 0) {
		return ranges;
	}
	const int endLineOffset = SafeRangeEnd(beginLineOffset, returnedLines);
	if (content.size() <= contextByteLimit + 64) {
		ranges.push_back({beginLineOffset, endLineOffset});
		return ranges;
	}

	const size_t headBytes = contextByteLimit > tailBytes
		? contextByteLimit - tailBytes
		: contextByteLimit;
	const size_t headEnd = ClampUtf8PrefixBoundary(content, headBytes);
	const size_t tailStart = ClampUtf8SuffixStartBoundary(content, tailBytes);
	const int headLines = CountCompleteNumberedLines(content, 0, headEnd);
	if (headLines > 0) {
		ranges.push_back({
			beginLineOffset,
			(std::min)(endLineOffset, SafeRangeEnd(beginLineOffset, headLines))
		});
	}
	if (tailStart <= headEnd) {
		return ranges;
	}

	size_t completeTailStart = tailStart;
	if (completeTailStart > 0 && content[completeTailStart - 1] != '\n') {
		const size_t firstLineEnd = content.find('\n', completeTailStart);
		if (firstLineEnd == std::string::npos) {
			return ranges;
		}
		completeTailStart = firstLineEnd + 1;
	}
	const int tailLines = CountCompleteNumberedLines(content, completeTailStart, content.size());
	if (tailLines > 0) {
		ranges.push_back({
			(std::max)(beginLineOffset, endLineOffset - tailLines),
			endLineOffset
		});
	}
	return ranges;
}

std::string BuildBlockedResult(int used, const std::string& reason, const std::string& hint)
{
	json result;
	result["ok"] = false;
	result["error"] = "tool_policy_blocked";
	result["reason"] = reason;
	result["hint"] = hint;
	result["exploration_calls_used"] = used;
	result["exploration_call_limit"] = kHardExplorationCallLimit;
	return result.dump();
}

std::string BuildReadOverlapBlockedResult(
	int used,
	const std::string& reason,
	const std::vector<ReadRequestRange>& requests,
	const std::vector<ReadCoverage>& coverage)
{
	json missingRows = json::array();
	json requestedRows = json::array();
	for (const ReadRequestRange& request : requests) {
		requestedRows.push_back({
			{"file_path", request.filePath},
			{"offset", request.begin},
			{"limit", request.end - request.begin}
		});
		for (const auto& missing : BuildMissingRanges(request, coverage)) {
			missingRows.push_back({
				{"file_path", request.filePath},
				{"offset", missing.first},
				{"limit", missing.second - missing.first}
			});
		}
	}

	json result;
	result["ok"] = false;
	result["error"] = "tool_policy_blocked";
	result["reason"] = reason;
	result["hint"] = missingRows.empty()
		? "Every requested line is already available in prior tool results. Reuse the existing content."
		: "The request overlaps prior reads. Call read_files once with only suggested_missing_ranges, or reuse the existing content.";
	result["requested_ranges"] = std::move(requestedRows);
	result["suggested_missing_ranges"] = missingRows;
	if (!missingRows.empty()) {
		result["suggested_tool"] = "read_files";
		result["suggested_arguments"] = {{"files", missingRows}};
	}
	result["exploration_calls_used"] = used;
	result["exploration_call_limit"] = kHardExplorationCallLimit;
	return result.dump();
}

} // namespace

struct Session::Impl {
	int explorationCalls = 0;
	bool postWriteReadBlocked = false;
	bool preferLowThinking = false;
	std::unordered_set<std::string> seenSignatures;
	std::vector<ReadCoverage> readCoverage;
	std::unordered_map<std::string, int> knownTotalLines;
};

Session::Session()
	: m_impl(std::make_unique<Impl>())
{
}

Session::~Session() = default;
Session::Session(Session&&) noexcept = default;
Session& Session::operator=(Session&&) noexcept = default;

Decision Session::BeforeToolCall(const std::string& toolName, const std::string& argumentsJsonUtf8)
{
	Impl* state = m_impl.get();
	Decision decision;
	state->preferLowThinking = false;

	const std::string signature = BuildSignature(toolName, argumentsJsonUtf8);
	if (state->seenSignatures.find(signature) != state->seenSignatures.end()) {
		decision.allowed = false;
		decision.reason = "duplicate_tool_call";
		decision.resultJsonUtf8 = BuildBlockedResult(
			state->explorationCalls,
			decision.reason,
			"The same tool call already ran in this request. Reuse the existing result instead of retrying it.");
		return decision;
	}

	if (state->postWriteReadBlocked && IsSourceReadTool(toolName)) {
		decision.allowed = false;
		decision.reason = "verified_write_already_completed";
		decision.resultJsonUtf8 = BuildBlockedResult(
			state->explorationCalls,
			decision.reason,
			"The source write already succeeded and was verified. Continue with compile/test or finish the response; do not re-read only for confirmation.");
		return decision;
	}

	if (toolName == "read_file" || toolName == "read_files") {
		std::vector<ReadRequestRange> requests = ParseReadRequests(toolName, argumentsJsonUtf8);
		bool hasOverlap = false;
		bool hasMissing = false;
		for (ReadRequestRange& request : requests) {
			const auto totalIt = state->knownTotalLines.find(request.normalizedPath);
			if (totalIt != state->knownTotalLines.end()) {
				request.begin = (std::min)(request.begin, totalIt->second);
				request.end = (std::min)(request.end, totalIt->second);
			}
			hasOverlap = hasOverlap || HasCoveredOverlap(request, state->readCoverage);
			hasMissing = hasMissing || !BuildMissingRanges(request, state->readCoverage).empty();
		}
		if (hasOverlap) {
			decision.allowed = false;
			decision.reason = hasMissing ? "overlapping_read_range" : "duplicate_read_range";
			decision.resultJsonUtf8 = BuildReadOverlapBlockedResult(
				state->explorationCalls,
				decision.reason,
				requests,
				state->readCoverage);
			return decision;
		}
	}

	if (IsSourceExplorationTool(toolName)) {
		if (state->explorationCalls >= kHardExplorationCallLimit) {
			decision.allowed = false;
			decision.reason = "exploration_budget_exhausted";
			decision.resultJsonUtf8 = BuildBlockedResult(
				state->explorationCalls,
				decision.reason,
				"Use the source context already collected and make the requested change. Only ask the user if a concrete missing fact prevents progress.");
			return decision;
		}
		++state->explorationCalls;
	}

	state->seenSignatures.insert(signature);
	return decision;
}

std::string Session::AfterToolCall(
	const std::string& toolName,
	const std::string& argumentsJsonUtf8,
	const std::string& resultJsonUtf8,
	bool ok)
{
	Impl* state = m_impl.get();
	const std::string signature = BuildSignature(toolName, argumentsJsonUtf8);
	if (!ok) {
		state->seenSignatures.erase(signature);
		state->preferLowThinking = false;
		if (state->postWriteReadBlocked && !IsSourceReadTool(toolName)) {
			state->postWriteReadBlocked = false;
		}
	}

	if (ok && (toolName == "read_file" || toolName == "read_files" || toolName == "read_code_item")) {
		try {
			const json result = json::parse(resultJsonUtf8);
			const auto recordRow = [state](
				const json& row,
				size_t contextByteLimit,
				size_t tailBytes,
				bool codeItem) {
				if (!row.is_object() || !row.value("ok", false)) {
					return;
				}
				const std::string content = row.value("content", std::string());
				const std::string normalizedPath = NormalizePath(row.value("file_path", std::string()));
				const int begin = codeItem
					? (std::max)(0, row.value("start_line", 1) - 1)
					: (std::max)(0, row.value("offset", 0));
				const int returnedLines = row.value("returned_lines", 0);
				if (returnedLines <= 0) {
					return;
				}
				for (const auto& visibleRange : BuildVisibleCoverageRanges(
					content,
					begin,
					returnedLines,
					contextByteLimit,
					tailBytes)) {
					if (!normalizedPath.empty() && visibleRange.second > visibleRange.first) {
						state->readCoverage.push_back({
							normalizedPath,
							visibleRange.first,
							visibleRange.second
						});
					}
				}
				if (!normalizedPath.empty() &&
					row.contains("total_lines") &&
					row["total_lines"].is_number_integer() &&
					row.value("total_lines_complete", true)) {
					state->knownTotalLines[normalizedPath] = (std::max)(0, row["total_lines"].get<int>());
				}
			};

			if (toolName == "read_files" && result.contains("files") && result["files"].is_array()) {
				for (const auto& row : result["files"]) {
					recordRow(row, kReadFilesPerFileContextBytes, 2048, false);
				}
			}
			else {
				recordRow(result, kReadFileContextBytes, 4096, toolName == "read_code_item");
			}
			CoalesceReadCoverage(state->readCoverage);
		}
		catch (...) {
		}
	}

	if (ok && IsWriteTool(toolName)) {
		state->postWriteReadBlocked = true;
		state->preferLowThinking = true;
		return "The write succeeded and was verified. Do not call source read tools again merely to confirm it; compile/test if requested, otherwise finish.";
	}
	if (ok && IsCompileTool(toolName) && state->postWriteReadBlocked) {
		state->preferLowThinking = true;
	}

	if (IsSourceExplorationTool(toolName) && state->explorationCalls >= kSoftExplorationCallLimit) {
		return "Exploration budget is nearly exhausted. Prefer read_files/read_code_item batching and proceed with the existing context instead of continuing broad searches.";
	}
	return std::string();
}

int Session::ExplorationCalls() const
{
	const auto* state = m_impl.get();
	return state == nullptr ? 0 : state->explorationCalls;
}

bool Session::PreferLowThinkingForNextRound() const
{
	const auto* state = m_impl.get();
	return state != nullptr && state->preferLowThinking;
}

} // namespace AIChatToolPolicy
