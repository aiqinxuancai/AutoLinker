#include "WorkspaceFileTools.h"

#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <regex>
#include <sstream>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "RealPageCodeToolSupport.h"
#include "WorkspaceMirror.h"

namespace WorkspaceFileTools {
namespace {

using json = nlohmann::json;

constexpr size_t kMaxReadBytes = 1024 * 1024;
constexpr int kDefaultListLimit = 500;
constexpr int kDefaultSearchLimit = 200;
constexpr int kDefaultReadLimit = 2000;
constexpr int kDefaultBatchReadLimit = 1200;
constexpr int kMaxBatchReadFiles = 24;
constexpr int kMaxBatchReadTotalLines = 12000;

std::wstring WideFromCodePage(const std::string& text, UINT codePage, DWORD flags = 0)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int wideLen = MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLen <= 0) {
		return std::wstring();
	}
	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), wide.data(), wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string StringFromWideCodePage(const std::wstring& text, UINT codePage)
{
	if (text.empty()) {
		return std::string();
	}
	const int outLen = WideCharToMultiByte(codePage, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (outLen <= 0) {
		return std::string();
	}
	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(codePage, 0, text.data(), static_cast<int>(text.size()), out.data(), outLen, nullptr, nullptr) <= 0) {
		return std::string();
	}
	return out;
}

bool IsValidUtf8Text(const std::string& text)
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
	const std::wstring wide = WideFromCodePage(text, fromCodePage, fromFlags);
	if (wide.empty()) {
		return text;
	}
	return StringFromWideCodePage(wide, toCodePage);
}

std::string LocalToUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8);
}

std::string Utf8ToLocalText(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string NormalizeSlash(std::string text)
{
	std::replace(text.begin(), text.end(), '\\', '/');
	return text;
}

std::string ToLowerAscii(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string TrimAsciiCopy(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin]))) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1]))) {
		--end;
	}
	return text.substr(begin, end - begin);
}

int ClampInt(int value, int minValue, int maxValue)
{
	return (std::max)(minValue, (std::min)(value, maxValue));
}

std::string GetJsonStringUtf8(const json& args, const char* key)
{
	if (!args.contains(key) || !args[key].is_string()) {
		return std::string();
	}
	return args[key].get<std::string>();
}

int GetJsonInt(const json& args, const char* key, int defaultValue)
{
	if (!args.contains(key) || !args[key].is_number_integer()) {
		return defaultValue;
	}
	return args[key].get<int>();
}

bool GetJsonBool(const json& args, const char* key, bool defaultValue)
{
	if (!args.contains(key) || !args[key].is_boolean()) {
		return defaultValue;
	}
	return args[key].get<bool>();
}

bool IsDefaultVisiblePath(const std::string& relativePath)
{
	const std::string path = ToLowerAscii(NormalizeSlash(relativePath));
	return path.rfind("src/", 0) == 0 ||
		path.rfind("ecom/", 0) == 0 ||
		path.rfind("elib/", 0) == 0 ||
		path.rfind("header/", 0) == 0 ||
		path == "agents.md" ||
		path == "info.json";
}

std::string EscapeRegexChar(char ch)
{
	static const std::string specials = R"(\.^$|()[]{}+?)";
	if (specials.find(ch) != std::string::npos) {
		return std::string("\\") + ch;
	}
	return std::string(1, ch);
}

std::regex GlobToRegex(const std::string& glob, bool caseInsensitive)
{
	std::string pattern = "^";
	for (size_t i = 0; i < glob.size(); ++i) {
		const char ch = glob[i];
		if (ch == '*') {
			if (i + 1 < glob.size() && glob[i + 1] == '*') {
				if (i + 2 < glob.size() && glob[i + 2] == '/') {
					pattern += "(.*/)?";
					i += 2;
				}
				else {
					pattern += ".*";
					++i;
				}
			}
			else {
				pattern += "[^/]*";
			}
		}
		else if (ch == '?') {
			pattern += "[^/]";
		}
		else {
			pattern += EscapeRegexChar(ch);
		}
	}
	pattern += "$";
	const auto flags = std::regex::ECMAScript | (caseInsensitive ? std::regex::icase : std::regex::flag_type{});
	return std::regex(pattern, flags);
}

bool GlobMatches(const std::string& relativePath, const std::string& glob)
{
	if (glob.empty() || glob == "**" || glob == "**/*") {
		return true;
	}
	try {
		return std::regex_match(NormalizeSlash(relativePath), GlobToRegex(NormalizeSlash(glob), true));
	}
	catch (...) {
		return false;
	}
}

bool ReadFileBytesLimited(
	const std::filesystem::path& path,
	std::string& outBytes,
	bool& outTruncated,
	size_t& outTotalBytes,
	std::string& outError)
{
	outBytes.clear();
	outTruncated = false;
	outTotalBytes = 0;
	outError.clear();

	std::ifstream file(path, std::ios::binary);
	if (!file) {
		outError = "open file failed";
		return false;
	}
	outBytes.assign(std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>());
	outTotalBytes = outBytes.size();
	if (outBytes.size() > kMaxReadBytes) {
		outBytes.resize(kMaxReadBytes);
		outTruncated = true;
	}
	return true;
}

bool LooksBinary(const std::string& bytes)
{
	const size_t limit = (std::min)(bytes.size(), static_cast<size_t>(4096));
	for (size_t i = 0; i < limit; ++i) {
		if (bytes[i] == '\0') {
			return true;
		}
	}
	return false;
}

std::string DecodeTextToUtf8(std::string bytes)
{
	if (bytes.size() >= 3 &&
		static_cast<unsigned char>(bytes[0]) == 0xEF &&
		static_cast<unsigned char>(bytes[1]) == 0xBB &&
		static_cast<unsigned char>(bytes[2]) == 0xBF) {
		bytes.erase(0, 3);
	}
	if (IsValidUtf8Text(bytes)) {
		return bytes;
	}
	return LocalToUtf8Text(bytes);
}

std::vector<std::string> SplitLinesUtf8(const std::string& text)
{
	std::vector<std::string> lines;
	std::string current;
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			lines.push_back(current);
			current.clear();
		}
		else if (ch == '\n') {
			lines.push_back(current);
			current.clear();
		}
		else {
			current.push_back(ch);
		}
	}
	if (!current.empty() || text.empty() || text.back() == '\n' || text.back() == '\r') {
		lines.push_back(current);
	}
	return lines;
}

std::string BuildNumberedView(const std::vector<std::string>& lines, int offset, int limit, int& outReturned, bool& outTruncated)
{
	outReturned = 0;
	outTruncated = false;
	const int total = static_cast<int>(lines.size());
	const int start = ClampInt(offset, 0, total);
	const int maxLines = limit <= 0 ? kDefaultReadLimit : ClampInt(limit, 1, 20000);
	const int end = (std::min)(total, start + maxLines);
	outTruncated = end < total;

	std::ostringstream stream;
	for (int i = start; i < end; ++i) {
		stream << (i + 1) << "\t" << lines[static_cast<size_t>(i)] << "\n";
		++outReturned;
	}
	return stream.str();
}

json BuildOffsetRange(int begin, int end)
{
	return {
		{"start_offset", begin},
		{"end_offset_exclusive", end}
	};
}

void AppendResultPaginationMetadata(
	json& result,
	int requestedOffset,
	int returned,
	int totalResults)
{
	const int effectiveOffset = ClampInt(requestedOffset, 0, totalResults);
	const int endOffset = (std::min)(totalResults, effectiveOffset + (std::max)(0, returned));
	const bool hasMore = endOffset < totalResults;
	result["requested_offset"] = requestedOffset;
	result["offset"] = effectiveOffset;
	result["returned"] = returned;
	result["total_results"] = totalResults;
	result["has_more"] = hasMore;
	if (hasMore) {
		result["next_offset"] = endOffset;
	}
	else {
		result["next_offset"] = nullptr;
	}
	result["visible_result_range"] = returned > 0
		? BuildOffsetRange(effectiveOffset, endOffset)
		: json(nullptr);
	json omittedRanges = json::array();
	if (effectiveOffset > 0) {
		omittedRanges.push_back(BuildOffsetRange(0, effectiveOffset));
	}
	if (hasMore) {
		omittedRanges.push_back(BuildOffsetRange(endOffset, totalResults));
	}
	result["omitted_result_ranges"] = std::move(omittedRanges);
}

void AppendLinePaginationMetadata(
	json& result,
	int requestedOffset,
	int effectiveOffset,
	int returned,
	int totalLines)
{
	const int endOffset = (std::min)(totalLines, effectiveOffset + (std::max)(0, returned));
	const bool hasMore = endOffset < totalLines;
	result["requested_offset"] = requestedOffset;
	result["offset"] = effectiveOffset;
	result["has_more"] = hasMore;
	if (hasMore) {
		result["next_offset"] = endOffset;
	}
	else {
		result["next_offset"] = nullptr;
	}
	result["visible_line_range"] = returned > 0
		? json({
			{"start_line", effectiveOffset + 1},
			{"end_line", endOffset}
		})
		: json(nullptr);
	json omittedRanges = json::array();
	if (effectiveOffset > 0) {
		omittedRanges.push_back({
			{"start_line", 1},
			{"end_line", effectiveOffset}
		});
	}
	if (hasMore) {
		omittedRanges.push_back({
			{"start_line", endOffset + 1},
			{"end_line", totalLines}
		});
	}
	result["omitted_line_ranges"] = std::move(omittedRanges);
}

void AppendSourceByteMetadata(
	json& result,
	size_t returnedBytes,
	size_t totalBytes,
	bool truncated)
{
	result["source_bytes_returned"] = returnedBytes;
	result["source_total_bytes"] = totalBytes;
	result["source_bytes_truncated"] = truncated;
	result["visible_source_byte_range"] = {
		{"start_offset", 0},
		{"end_offset_exclusive", returnedBytes}
	};
	json omittedRanges = json::array();
	if (truncated && totalBytes > returnedBytes) {
		omittedRanges.push_back({
			{"start_offset", returnedBytes},
			{"end_offset_exclusive", totalBytes}
		});
	}
	result["omitted_source_byte_ranges"] = std::move(omittedRanges);
}

std::string ToolResultToLocalJson(const json& result)
{
	// 页面内容/搜索命中可能含非法 UTF-8 字节（如易语言字节集里的 GBK 字节 0xB5）。
	// 裸 dump() 遇到会抛 type_error.316，异常一路逃出主线程 WndProc 导致 IDE 崩溃。
	// 用 error_handler=replace 把非法字节替换为 U+FFFD，保证序列化绝不抛异常。
	return Utf8ToLocalText(result.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace));
}

std::string BuildError(const std::string& error)
{
	json r;
	r["ok"] = false;
	r["error"] = error;
	return ToolResultToLocalJson(r);
}

bool BuildReadFileRow(
	const std::string& filePath,
	int offset,
	int limit,
	json& outRow,
	std::string& outError)
{
	outRow = json::object();
	outError.clear();

	std::filesystem::path fullPath;
	std::string relativePath;
	std::string error;
	if (!WorkspaceMirror::ResolveFilePath(filePath, fullPath, relativePath, error)) {
		outError = error;
		return false;
	}

	std::string bytes;
	bool fileTruncated = false;
	size_t totalBytes = 0;
	if (!ReadFileBytesLimited(fullPath, bytes, fileTruncated, totalBytes, error)) {
		outError = error;
		return false;
	}
	if (LooksBinary(bytes)) {
		outError = "read_file only supports text files: " + relativePath;
		return false;
	}

	const size_t returnedBytes = bytes.size();
	const std::string text = DecodeTextToUtf8(std::move(bytes));
	const std::string hashText = NormalizeRealCodeLineBreaksToCrLf(Utf8ToLocalText(text));
	const auto lines = SplitLinesUtf8(text);
	const int totalLines = static_cast<int>(lines.size());
	const int effectiveOffset = ClampInt(offset, 0, totalLines);
	int returned = 0;
	bool lineTruncated = false;
	const std::string view = BuildNumberedView(lines, effectiveOffset, limit, returned, lineTruncated);

	outRow["ok"] = true;
	outRow["file_path"] = relativePath;
	outRow["code_kind"] = "mirror_source";
	outRow["code_hash"] = BuildStableTextHashForRealCode(hashText);
	outRow["code_hash_complete"] = !fileTruncated;
	outRow["total_lines"] = totalLines;
	outRow["total_lines_complete"] = !fileTruncated;
	outRow["returned_lines"] = returned;
	outRow["truncated"] = fileTruncated || effectiveOffset > 0 || lineTruncated;
	outRow["content"] = view;
	AppendLinePaginationMetadata(outRow, offset, effectiveOffset, returned, totalLines);
	AppendSourceByteMetadata(outRow, returnedBytes, totalBytes, fileTruncated);
	return true;
}

std::string ExecuteReadFile(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}

	const std::string filePath = GetJsonStringUtf8(args, "file_path");
	const int offset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int limit = GetJsonInt(args, "limit", 0);

	json r;
	std::string error;
	if (!BuildReadFileRow(filePath, offset, limit, r, error)) {
		return BuildError(error);
	}
	outOk = true;
	return ToolResultToLocalJson(r);
}

struct BatchReadRequest {
	std::string filePath;
	int offset = 0;
	int limit = kDefaultBatchReadLimit;
};

void AddBatchReadRequest(
	std::vector<BatchReadRequest>& requests,
	const BatchReadRequest& request)
{
	if (TrimAsciiCopy(request.filePath).empty()) {
		return;
	}
	const auto exists = std::find_if(
		requests.begin(),
		requests.end(),
		[&request](const BatchReadRequest& item) {
			return ToLowerAscii(NormalizeSlash(item.filePath)) ==
				ToLowerAscii(NormalizeSlash(request.filePath));
		});
	if (exists != requests.end()) {
		return;
	}
	requests.push_back(request);
}

json BuildBatchReadRequestJson(const BatchReadRequest& request, size_t requestIndex)
{
	return {
		{"request_index", requestIndex},
		{"file_path", request.filePath},
		{"offset", request.offset},
		{"limit", request.limit}
	};
}

std::string ExecuteReadFiles(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}

	const int defaultOffset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int defaultLimit = ClampInt(GetJsonInt(args, "limit", kDefaultBatchReadLimit), 1, kDefaultReadLimit);
	std::vector<BatchReadRequest> requests;

	if (args.contains("file_paths") && args["file_paths"].is_array()) {
		for (const auto& item : args["file_paths"]) {
			if (!item.is_string()) {
				continue;
			}
			BatchReadRequest request;
			request.filePath = item.get<std::string>();
			request.offset = defaultOffset;
			request.limit = defaultLimit;
			AddBatchReadRequest(requests, request);
		}
	}
	if (args.contains("files") && args["files"].is_array()) {
		for (const auto& item : args["files"]) {
			BatchReadRequest request;
			request.offset = defaultOffset;
			request.limit = defaultLimit;
			if (item.is_string()) {
				request.filePath = item.get<std::string>();
			}
			else if (item.is_object()) {
				request.filePath = GetJsonStringUtf8(item, "file_path");
				request.offset = (std::max)(0, GetJsonInt(item, "offset", defaultOffset));
				request.limit = ClampInt(GetJsonInt(item, "limit", defaultLimit), 1, kDefaultReadLimit);
			}
			AddBatchReadRequest(requests, request);
		}
	}

	if (requests.empty()) {
		return BuildError("read_files requires file_paths or files");
	}

	json rows = json::array();
	int okCount = 0;
	int errorCount = 0;
	int returnedLines = 0;
	const size_t scheduledCount = (std::min)(requests.size(), static_cast<size_t>(kMaxBatchReadFiles));
	bool outputTruncated = scheduledCount < requests.size();
	json omittedRequests = json::array();
	for (size_t i = scheduledCount; i < requests.size(); ++i) {
		omittedRequests.push_back(BuildBatchReadRequestJson(requests[i], i));
	}

	for (size_t requestIndex = 0; requestIndex < scheduledCount; ++requestIndex) {
		const BatchReadRequest& request = requests[requestIndex];
		if (returnedLines >= kMaxBatchReadTotalLines) {
			outputTruncated = true;
			for (size_t i = requestIndex; i < scheduledCount; ++i) {
				omittedRequests.push_back(BuildBatchReadRequestJson(requests[i], i));
			}
			break;
		}
		json row;
		const int remainingLines = kMaxBatchReadTotalLines - returnedLines;
		const int effectiveLimit = (std::min)(request.limit, remainingLines);
		std::string error;
		if (BuildReadFileRow(request.filePath, request.offset, effectiveLimit, row, error)) {
			++okCount;
			returnedLines += row.value("returned_lines", 0);
			if (row.value("truncated", false)) {
				outputTruncated = true;
			}
		}
		else {
			++errorCount;
			row["ok"] = false;
			row["file_path"] = request.filePath;
			row["error"] = error;
		}
		row["request_index"] = requestIndex;
		rows.push_back(std::move(row));
	}

	json r;
	r["ok"] = okCount > 0;
	r["all_ok"] = errorCount == 0 && !outputTruncated;
	r["code_kind"] = "mirror_source";
	r["files"] = std::move(rows);
	r["requested"] = requests.size();
	r["scheduled"] = scheduledCount;
	r["returned"] = r["files"].size();
	r["ok_count"] = okCount;
	r["error_count"] = errorCount;
	r["returned_lines"] = returnedLines;
	r["truncated"] = outputTruncated;
	r["omitted_requests"] = std::move(omittedRequests);
	r["omitted_request_count"] = r["omitted_requests"].size();
	if (!r["omitted_requests"].empty()) {
		r["next_request_index"] = r["omitted_requests"][0].value("request_index", 0);
	}
	else {
		r["next_request_index"] = nullptr;
	}
	r["max_files"] = kMaxBatchReadFiles;
	r["max_total_lines"] = kMaxBatchReadTotalLines;
	outOk = okCount > 0;
	return ToolResultToLocalJson(r);
}

std::string ExecuteListFiles(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}

	const std::string glob = GetJsonStringUtf8(args, "glob");
	const std::string pathPrefix = NormalizeSlash(GetJsonStringUtf8(args, "path"));
	const int offset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int limit = ClampInt(GetJsonInt(args, "limit", kDefaultListLimit), 1, 5000);

	std::vector<std::string> files;
	std::string error;
	if (!WorkspaceMirror::ListMirrorFiles(files, error)) {
		return BuildError(error);
	}

	json rows = json::array();
	int matched = 0;
	for (const std::string& file : files) {
		if (!pathPrefix.empty() && NormalizeSlash(file).rfind(pathPrefix, 0) != 0) {
			continue;
		}
		if (!glob.empty() && !GlobMatches(file, glob)) {
			continue;
		}
		if (glob.empty() && pathPrefix.empty() && !IsDefaultVisiblePath(file)) {
			continue;
		}
		const int resultOffset = matched;
		++matched;
		if (resultOffset >= offset && static_cast<int>(rows.size()) < limit) {
			rows.push_back(file);
		}
	}

	json r;
	r["ok"] = true;
	r["files"] = std::move(rows);
	r["count"] = matched;
	const int returned = static_cast<int>(r["files"].size());
	AppendResultPaginationMetadata(r, offset, returned, matched);
	r["truncated"] = r["offset"].get<int>() > 0 || r["has_more"].get<bool>();
	outOk = true;
	return ToolResultToLocalJson(r);
}

bool LineMatches(
	const std::string& line,
	const std::string& pattern,
	bool regexMode,
	bool caseInsensitive,
	const std::regex* compiledRegex)
{
	if (regexMode) {
		return compiledRegex != nullptr && std::regex_search(line, *compiledRegex);
	}
	if (caseInsensitive) {
		return ToLowerAscii(line).find(ToLowerAscii(pattern)) != std::string::npos;
	}
	return line.find(pattern) != std::string::npos;
}

json BuildContextLines(const std::vector<std::string>& lines, int center, int context)
{
	json rows = json::array();
	const int start = (std::max)(0, center - context);
	const int end = (std::min)(static_cast<int>(lines.size()), center + context + 1);
	for (int i = start; i < end; ++i) {
		rows.push_back({
			{"line_number", i + 1},
			{"text", lines[static_cast<size_t>(i)]}
		});
	}
	return rows;
}

json BuildSearchErrorJson(const std::string& error)
{
	json r;
	r["ok"] = false;
	r["error"] = error;
	return r;
}

json ExecuteSearchCodeQuery(
	const std::string& pattern,
	const std::string& glob,
	const std::string& outputMode,
	bool caseInsensitive,
	bool regexMode,
	int context,
	int offset,
	int pageLimit)
{
	if (pattern.empty()) {
		return BuildSearchErrorJson("pattern is required");
	}
	if (outputMode != "content" && outputMode != "files_with_matches" && outputMode != "count") {
		return BuildSearchErrorJson("output_mode must be content, files_with_matches, or count");
	}

	std::regex compiledRegex;
	const std::regex* regexPtr = nullptr;
	if (regexMode) {
		try {
			compiledRegex = std::regex(
				pattern,
				std::regex::ECMAScript | (caseInsensitive ? std::regex::icase : std::regex::flag_type{}));
			regexPtr = &compiledRegex;
		}
		catch (const std::exception& ex) {
			return BuildSearchErrorJson(std::string("invalid regex pattern: ") + ex.what());
		}
	}

	std::vector<std::string> files;
	std::string error;
	if (!WorkspaceMirror::ListMirrorFiles(files, error)) {
		return BuildSearchErrorJson(error);
	}

	json results = json::array();
	int filesWithMatches = 0;
	int totalMatches = 0;
	int totalResults = 0;
	int skippedBinary = 0;
	int sourceFilesTruncated = 0;
	size_t sourceBytesOmitted = 0;

	for (const std::string& file : files) {
		if (!glob.empty() && !GlobMatches(file, glob)) {
			continue;
		}
		if (glob.empty() && !IsDefaultVisiblePath(file)) {
			continue;
		}

		std::filesystem::path fullPath;
		std::string resolvedRelative;
		if (!WorkspaceMirror::ResolveFilePath(file, fullPath, resolvedRelative, error)) {
			continue;
		}

		std::string bytes;
		bool fileTruncated = false;
		size_t totalBytes = 0;
		if (!ReadFileBytesLimited(fullPath, bytes, fileTruncated, totalBytes, error)) {
			continue;
		}
		if (fileTruncated) {
			++sourceFilesTruncated;
			if (totalBytes > bytes.size()) {
				sourceBytesOmitted += totalBytes - bytes.size();
			}
		}
		if (LooksBinary(bytes)) {
			++skippedBinary;
			continue;
		}
		const auto lines = SplitLinesUtf8(DecodeTextToUtf8(std::move(bytes)));

		int fileMatches = 0;
		for (int i = 0; i < static_cast<int>(lines.size()); ++i) {
			const std::string& line = lines[static_cast<size_t>(i)];
			if (!LineMatches(line, pattern, regexMode, caseInsensitive, regexPtr)) {
				continue;
			}
			++fileMatches;
			++totalMatches;
			if (outputMode == "content") {
				const int resultOffset = totalResults;
				++totalResults;
				if (resultOffset < offset || static_cast<int>(results.size()) >= pageLimit) {
					continue;
				}
				json row;
				row["file_path"] = file;
				row["line_number"] = i + 1;
				row["text"] = line;
				if (context > 0) {
					row["context"] = BuildContextLines(lines, i, context);
				}
				results.push_back(std::move(row));
			}
		}

		if (fileMatches <= 0) {
			continue;
		}
		++filesWithMatches;
		if (outputMode != "content") {
			const int resultOffset = totalResults;
			++totalResults;
			if (resultOffset >= offset && static_cast<int>(results.size()) < pageLimit) {
				if (outputMode == "files_with_matches") {
					results.push_back(file);
				}
				else {
					results.push_back({
						{"file_path", file},
						{"count", fileMatches}
					});
				}
			}
		}
	}

	json r;
	r["ok"] = true;
	r["pattern"] = pattern;
	r["regex"] = regexMode;
	r["case_insensitive"] = caseInsensitive;
	r["output_mode"] = outputMode;
	r["results"] = std::move(results);
	r["files_with_matches"] = filesWithMatches;
	r["match_count"] = totalMatches;
	r["skipped_binary"] = skippedBinary;
	r["source_files_truncated"] = sourceFilesTruncated;
	r["source_bytes_omitted"] = sourceBytesOmitted;
	r["results_complete"] = sourceFilesTruncated == 0;
	const int returned = static_cast<int>(r["results"].size());
	AppendResultPaginationMetadata(r, offset, returned, totalResults);
	r["page_limit"] = pageLimit;
	r["truncated"] = r["offset"].get<int>() > 0 ||
		r["has_more"].get<bool>() ||
		sourceFilesTruncated > 0;
	return r;
}

bool IsTopLevelDeclarationDirective(const std::string& directiveLower)
{
	return directiveLower == ".子程序" ||
		directiveLower == ".程序集" ||
		directiveLower == ".窗口程序集" ||
		directiveLower == ".类模块" ||
		directiveLower == ".程序集变量" ||
		directiveLower == ".全局变量" ||
		directiveLower == ".dll命令" ||
		directiveLower == ".dll声明" ||
		directiveLower == ".数据类型" ||
		directiveLower == ".常量" ||
		directiveLower == ".图片资源" ||
		directiveLower == ".声音资源";
}

bool TryParseTopLevelDeclaration(
	const std::string& line,
	std::string& outDirective,
	std::string& outName)
{
	outDirective.clear();
	outName.clear();
	const std::string trimmed = TrimAsciiCopy(line);
	if (trimmed.empty() || trimmed.front() != '.') {
		return false;
	}
	const size_t directiveEnd = trimmed.find_first_of(" \t");
	if (directiveEnd == std::string::npos) {
		return false;
	}
	outDirective = ToLowerAscii(trimmed.substr(0, directiveEnd));
	if (!IsTopLevelDeclarationDirective(outDirective)) {
		return false;
	}
	const size_t nameBegin = trimmed.find_first_not_of(" \t", directiveEnd);
	if (nameBegin == std::string::npos) {
		return false;
	}
	const size_t nameEnd = trimmed.find(',', nameBegin);
	outName = TrimAsciiCopy(trimmed.substr(
		nameBegin,
		nameEnd == std::string::npos ? std::string::npos : nameEnd - nameBegin));
	return !outName.empty();
}

std::string ExecuteReadCodeItem(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}

	const std::string filePath = GetJsonStringUtf8(args, "file_path");
	const std::string itemName = TrimAsciiCopy(GetJsonStringUtf8(args, "item_name"));
	if (filePath.empty() || itemName.empty()) {
		return BuildError("read_code_item requires file_path and item_name");
	}

	std::filesystem::path fullPath;
	std::string relativePath;
	std::string error;
	if (!WorkspaceMirror::ResolveFilePath(filePath, fullPath, relativePath, error)) {
		return BuildError(error);
	}

	std::string bytes;
	bool fileTruncated = false;
	size_t totalBytes = 0;
	if (!ReadFileBytesLimited(fullPath, bytes, fileTruncated, totalBytes, error)) {
		return BuildError(error);
	}
	if (LooksBinary(bytes)) {
		return BuildError("read_code_item only supports text files: " + relativePath);
	}

	const size_t returnedBytes = bytes.size();
	const std::string text = DecodeTextToUtf8(std::move(bytes));
	const std::vector<std::string> lines = SplitLinesUtf8(text);
	std::vector<int> matches;
	std::vector<std::string> matchDirectives;
	for (int i = 0; i < static_cast<int>(lines.size()); ++i) {
		std::string directive;
		std::string name;
		if (!TryParseTopLevelDeclaration(lines[static_cast<size_t>(i)], directive, name)) {
			continue;
		}
		if (ToLowerAscii(name) == ToLowerAscii(itemName)) {
			matches.push_back(i);
			matchDirectives.push_back(std::move(directive));
		}
	}

	if (matches.empty()) {
		json r;
		r["ok"] = false;
		r["error"] = fileTruncated
			? "code item not found in visible file prefix; source file was truncated"
			: "code item not found";
		r["file_path"] = relativePath;
		r["item_name"] = itemName;
		r["search_complete"] = !fileTruncated;
		AppendSourceByteMetadata(r, returnedBytes, totalBytes, fileTruncated);
		return ToolResultToLocalJson(r);
	}
	if (matches.size() > 1) {
		json candidates = json::array();
		for (size_t i = 0; i < matches.size(); ++i) {
			candidates.push_back({
				{"line_number", matches[i] + 1},
				{"directive", matchDirectives[i]}
			});
		}
		json r;
		r["ok"] = false;
		r["error"] = "code item name matched multiple declarations";
		r["file_path"] = relativePath;
		r["item_name"] = itemName;
		r["candidates"] = std::move(candidates);
		r["search_complete"] = !fileTruncated;
		AppendSourceByteMetadata(r, returnedBytes, totalBytes, fileTruncated);
		return ToolResultToLocalJson(r);
	}

	const int start = matches.front();
	int end = static_cast<int>(lines.size());
	bool foundNextDeclaration = false;
	for (int i = start + 1; i < static_cast<int>(lines.size()); ++i) {
		std::string directive;
		std::string name;
		if (TryParseTopLevelDeclaration(lines[static_cast<size_t>(i)], directive, name)) {
			end = i;
			foundNextDeclaration = true;
			break;
		}
	}

	std::ostringstream content;
	for (int i = start; i < end; ++i) {
		content << (i + 1) << "\t" << lines[static_cast<size_t>(i)] << "\n";
	}

	const std::string hashText = NormalizeRealCodeLineBreaksToCrLf(Utf8ToLocalText(text));
	json r;
	r["ok"] = true;
	r["file_path"] = relativePath;
	r["code_kind"] = "mirror_source";
	r["code_hash"] = BuildStableTextHashForRealCode(hashText);
	r["code_hash_complete"] = !fileTruncated;
	r["item_name"] = itemName;
	r["directive"] = matchDirectives.front();
	r["start_line"] = start + 1;
	r["end_line"] = end;
	r["returned_lines"] = end - start;
	r["total_lines"] = lines.size();
	r["total_lines_complete"] = !fileTruncated;
	r["item_complete"] = foundNextDeclaration || !fileTruncated;
	r["truncated"] = !r["item_complete"].get<bool>();
	r["content"] = content.str();
	AppendSourceByteMetadata(r, returnedBytes, totalBytes, fileTruncated);

	if (GetJsonBool(args, "include_references", false)) {
		const std::string referenceGlob = GetJsonStringUtf8(args, "reference_glob");
		const int referenceLimit = ClampInt(GetJsonInt(args, "reference_limit", 20), 1, 50);
		r["references"] = ExecuteSearchCodeQuery(
			itemName,
			referenceGlob,
			"content",
			false,
			false,
			1,
			0,
			referenceLimit);
	}

	outOk = true;
	return ToolResultToLocalJson(r);
}

std::vector<std::string> GetSearchPatterns(const json& args, bool& outPatternListTruncated)
{
	std::vector<std::string> patterns;
	const auto addPattern = [&patterns, &outPatternListTruncated](const std::string& pattern) {
		if (TrimAsciiCopy(pattern).empty()) {
			return;
		}
		const auto exists = std::find(patterns.begin(), patterns.end(), pattern);
		if (exists != patterns.end()) {
			return;
		}
		if (patterns.size() >= 16) {
			outPatternListTruncated = true;
			return;
		}
		patterns.push_back(pattern);
	};

	addPattern(GetJsonStringUtf8(args, "pattern"));
	if (args.contains("patterns") && args["patterns"].is_array()) {
		for (const auto& item : args["patterns"]) {
			if (item.is_string()) {
				addPattern(item.get<std::string>());
			}
		}
	}
	return patterns;
}

std::string ExecuteSearchCode(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}

	const std::string glob = GetJsonStringUtf8(args, "glob");
	const std::string outputMode = GetJsonStringUtf8(args, "output_mode").empty()
		? "content"
		: GetJsonStringUtf8(args, "output_mode");
	const bool caseInsensitive = GetJsonBool(args, "case_insensitive", false);
	const bool regexMode = GetJsonBool(args, "regex", true);
	const int context = ClampInt(GetJsonInt(args, "context", 0), 0, 20);
	const int offset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int headLimit = ClampInt(GetJsonInt(args, "head_limit", kDefaultSearchLimit), 1, 2000);
	bool patternListTruncated = false;
	const std::vector<std::string> patterns = GetSearchPatterns(args, patternListTruncated);
	if (patterns.empty()) {
		return BuildError("pattern or patterns is required");
	}

	if (patterns.size() == 1 && !patternListTruncated) {
		const json r = ExecuteSearchCodeQuery(
			patterns.front(),
			glob,
			outputMode,
			caseInsensitive,
			regexMode,
			context,
			offset,
			headLimit);
		outOk = r.value("ok", false);
		return ToolResultToLocalJson(r);
	}

	json queries = json::array();
	bool anyOk = false;
	bool anyTruncated = patternListTruncated;
	bool anyHasMore = false;
	for (const std::string& pattern : patterns) {
		json row = ExecuteSearchCodeQuery(
			pattern,
			glob,
			outputMode,
			caseInsensitive,
			regexMode,
			context,
			offset,
			headLimit);
		anyOk = anyOk || row.value("ok", false);
		anyTruncated = anyTruncated || row.value("truncated", false);
		anyHasMore = anyHasMore || row.value("has_more", false);
		queries.push_back(std::move(row));
	}

	json r;
	r["ok"] = anyOk;
	r["batch"] = true;
	r["regex"] = regexMode;
	r["case_insensitive"] = caseInsensitive;
	r["output_mode"] = outputMode;
	r["glob"] = glob;
	r["offset"] = offset;
	r["page_limit"] = headLimit;
	r["queries"] = std::move(queries);
	r["requested"] = patterns.size();
	r["returned"] = r["queries"].size();
	r["has_more"] = anyHasMore;
	r["truncated"] = anyTruncated;
	outOk = anyOk;
	return ToolResultToLocalJson(r);
}

} // namespace

bool CanHandleTool(const std::string& toolName)
{
	return toolName == "read_file" ||
		toolName == "read_files" ||
		toolName == "read_code_item" ||
		toolName == "list_files" ||
		toolName == "search_code";
}

std::string ExecuteTool(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	if (toolName == "read_file") {
		return ExecuteReadFile(argumentsJson, outOk);
	}
	if (toolName == "read_files") {
		return ExecuteReadFiles(argumentsJson, outOk);
	}
	if (toolName == "read_code_item") {
		return ExecuteReadCodeItem(argumentsJson, outOk);
	}
	if (toolName == "list_files") {
		return ExecuteListFiles(argumentsJson, outOk);
	}
	if (toolName == "search_code") {
		return ExecuteSearchCode(argumentsJson, outOk);
	}
	outOk = false;
	return BuildError("unknown workspace file tool: " + toolName);
}

std::string BuildPaginationSelfTestJson()
{
	json resultPage;
	AppendResultPaginationMetadata(resultPage, 10, 5, 30);
	json linePage;
	AppendLinePaginationMetadata(linePage, 20, 20, 10, 100);
	json sourcePage;
	AppendSourceByteMetadata(sourcePage, 1024, 4096, true);

	const bool resultPageOk =
		resultPage.value("offset", -1) == 10 &&
		resultPage.value("next_offset", -1) == 15 &&
		resultPage.value("total_results", -1) == 30 &&
		resultPage["omitted_result_ranges"].is_array() &&
		resultPage["omitted_result_ranges"].size() == 2 &&
		resultPage["omitted_result_ranges"][0].value("start_offset", -1) == 0 &&
		resultPage["omitted_result_ranges"][0].value("end_offset_exclusive", -1) == 10 &&
		resultPage["omitted_result_ranges"][1].value("start_offset", -1) == 15 &&
		resultPage["omitted_result_ranges"][1].value("end_offset_exclusive", -1) == 30;
	const bool linePageOk =
		linePage.value("next_offset", -1) == 30 &&
		linePage["visible_line_range"].value("start_line", -1) == 21 &&
		linePage["visible_line_range"].value("end_line", -1) == 30 &&
		linePage["omitted_line_ranges"].size() == 2;
	const bool sourcePageOk =
		sourcePage.value("source_bytes_truncated", false) &&
		sourcePage["omitted_source_byte_ranges"].size() == 1 &&
		sourcePage["omitted_source_byte_ranges"][0].value("start_offset", 0u) == 1024u &&
		sourcePage["omitted_source_byte_ranges"][0].value("end_offset_exclusive", 0u) == 4096u;

	return json({
		{"name", "workspace-file-pagination"},
		{"ok", resultPageOk && linePageOk && sourcePageOk},
		{"result_page", std::move(resultPage)},
		{"line_page", std::move(linePage)},
		{"source_page", std::move(sourcePage)}
	}).dump();
}

} // namespace WorkspaceFileTools
