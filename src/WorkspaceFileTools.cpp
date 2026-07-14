#include "WorkspaceFileTools.h"

#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <optional>
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
constexpr int kMaxBatchReadFiles = 12;
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

size_t GetJsonSize(const json& args, const char* key, size_t defaultValue)
{
	if (!args.contains(key)) {
		return defaultValue;
	}
	if (args[key].is_number_unsigned()) {
		return args[key].get<size_t>();
	}
	if (args[key].is_number_integer()) {
		const long long value = args[key].get<long long>();
		return value >= 0 ? static_cast<size_t>(value) : defaultValue;
	}
	return defaultValue;
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

bool TryCompileGlob(const std::string& glob, std::optional<std::regex>& outRegex, std::string& outError)
{
	outRegex.reset();
	outError.clear();
	if (glob.empty() || glob == "**" || glob == "**/*") {
		return true;
	}
	try {
		outRegex.emplace(GlobToRegex(NormalizeSlash(glob), true));
		return true;
	}
	catch (const std::exception& ex) {
		outError = std::string("invalid glob: ") + ex.what();
		return false;
	}
}

bool CompiledGlobMatches(const std::string& relativePath, const std::optional<std::regex>& regex)
{
	return !regex.has_value() || std::regex_match(NormalizeSlash(relativePath), *regex);
}

bool ReadFileBytesWindow(
	const std::filesystem::path& path,
	size_t requestedByteOffset,
	std::string& outBytes,
	size_t& outEffectiveByteOffset,
	bool& outTruncated,
	size_t& outTotalBytes,
	std::string& outError)
{
	outBytes.clear();
	outEffectiveByteOffset = 0;
	outTruncated = false;
	outTotalBytes = 0;
	outError.clear();

	std::ifstream file(path, std::ios::binary);
	if (!file) {
		outError = "open file failed";
		return false;
	}
	std::error_code ec;
	outTotalBytes = static_cast<size_t>(std::filesystem::file_size(path, ec));
	if (ec) {
		outError = "query file size failed";
		return false;
	}
	outEffectiveByteOffset = (std::min)(requestedByteOffset, outTotalBytes);
	if (outEffectiveByteOffset > 0) {
		file.seekg(static_cast<std::streamoff>(outEffectiveByteOffset), std::ios::beg);
		if (!file) {
			outError = "seek file failed";
			return false;
		}
	}
	const size_t readBytes = (std::min)(outTotalBytes - outEffectiveByteOffset, kMaxReadBytes);
	outBytes.assign(readBytes, '\0');
	if (readBytes > 0) {
		file.read(outBytes.data(), static_cast<std::streamsize>(readBytes));
		const std::streamsize got = file.gcount();
		if (got < 0) {
			outError = "read file failed";
			return false;
		}
		outBytes.resize(static_cast<size_t>(got));
	}
	while (outEffectiveByteOffset > 0 && !outBytes.empty() &&
		(static_cast<unsigned char>(outBytes.front()) & 0xC0) == 0x80) {
		outBytes.erase(outBytes.begin());
		++outEffectiveByteOffset;
	}
	outTruncated = outEffectiveByteOffset + outBytes.size() < outTotalBytes;
	if (outTruncated && !outBytes.empty()) {
		const size_t lastNewline = outBytes.find_last_of("\r\n");
		if (lastNewline != std::string::npos && lastNewline + 1 < outBytes.size()) {
			outBytes.resize(lastNewline + 1);
		}
		else if (!IsValidUtf8Text(outBytes)) {
			const size_t originalSize = outBytes.size();
			bool repaired = false;
			for (int removed = 0; removed < 4 && !outBytes.empty(); ++removed) {
				outBytes.pop_back();
				if (IsValidUtf8Text(outBytes)) {
					repaired = true;
					break;
				}
			}
			if (!repaired) {
				outBytes.resize(originalSize);
			}
		}
	}
	return true;
}

bool ReadFileBytesLimited(
	const std::filesystem::path& path,
	std::string& outBytes,
	bool& outTruncated,
	size_t& outTotalBytes,
	std::string& outError)
{
	size_t effectiveByteOffset = 0;
	return ReadFileBytesWindow(
		path,
		0,
		outBytes,
		effectiveByteOffset,
		outTruncated,
		outTotalBytes,
		outError);
}

size_t CountLineBreaksBefore(const std::filesystem::path& path, size_t byteOffset)
{
	if (byteOffset == 0) {
		return 0;
	}
	std::ifstream file(path, std::ios::binary);
	if (!file) {
		return 0;
	}
	size_t remaining = byteOffset;
	size_t lineBreaks = 0;
	char buffer[64 * 1024];
	while (remaining > 0 && file) {
		const size_t wanted = (std::min)(remaining, sizeof(buffer));
		file.read(buffer, static_cast<std::streamsize>(wanted));
		const std::streamsize got = file.gcount();
		if (got <= 0) {
			break;
		}
		for (std::streamsize i = 0; i < got; ++i) {
			if (buffer[i] == '\n') {
				++lineBreaks;
			}
		}
		remaining -= static_cast<size_t>(got);
	}
	return lineBreaks;
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

std::string BuildNumberedView(
	const std::vector<std::string>& lines,
	int offset,
	int limit,
	int& outReturned,
	bool& outTruncated,
	size_t lineNumberBase = 0)
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
		stream << (lineNumberBase + static_cast<size_t>(i) + 1) << "\t" << lines[static_cast<size_t>(i)] << "\n";
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
	int totalLines,
	size_t lineNumberBase = 0)
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
			{"start_line", lineNumberBase + static_cast<size_t>(effectiveOffset) + 1},
			{"end_line", lineNumberBase + static_cast<size_t>(endOffset)}
		})
		: json(nullptr);
	json omittedRanges = json::array();
	if (lineNumberBase + static_cast<size_t>(effectiveOffset) > 0) {
		omittedRanges.push_back({
			{"start_line", 1},
			{"end_line", lineNumberBase + static_cast<size_t>(effectiveOffset)}
		});
	}
	if (hasMore) {
		omittedRanges.push_back({
			{"start_line", lineNumberBase + static_cast<size_t>(endOffset) + 1},
			{"end_line", lineNumberBase + static_cast<size_t>(totalLines)}
		});
	}
	result["omitted_line_ranges"] = std::move(omittedRanges);
}

void AppendSourceByteMetadata(
	json& result,
	size_t returnedBytes,
	size_t totalBytes,
	bool truncated,
	size_t startOffset = 0)
{
	result["source_bytes_returned"] = returnedBytes;
	result["source_total_bytes"] = totalBytes;
	result["source_bytes_truncated"] = truncated;
	result["visible_source_byte_range"] = {
		{"start_offset", startOffset},
		{"end_offset_exclusive", startOffset + returnedBytes}
	};
	result["has_more_source_bytes"] = startOffset + returnedBytes < totalBytes;
	result["next_source_byte_offset"] = startOffset + returnedBytes < totalBytes
		? json(startOffset + returnedBytes)
		: json(nullptr);
	json omittedRanges = json::array();
	if (startOffset > 0) {
		omittedRanges.push_back({
			{"start_offset", 0},
			{"end_offset_exclusive", startOffset}
		});
	}
	if (truncated && totalBytes > startOffset + returnedBytes) {
		omittedRanges.push_back({
			{"start_offset", startOffset + returnedBytes},
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

bool ValidateMirrorGeneration(
	const json& args,
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	std::string& outError)
{
	outError.clear();
	if (!args.contains("mirror_generation")) {
		return true;
	}
	if (!args["mirror_generation"].is_number_unsigned() && !args["mirror_generation"].is_number_integer()) {
		outError = "mirror_generation must be a non-negative integer";
		return false;
	}
	const long long requested = args["mirror_generation"].get<long long>();
	if (requested < 0 || static_cast<std::uint64_t>(requested) != snapshot.generation) {
		outError = "stale mirror_generation; refresh or restart pagination with current generation " +
			std::to_string(snapshot.generation);
		return false;
	}
	return true;
}

bool BuildReadFileRow(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& filePath,
	size_t byteOffset,
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
	if (!WorkspaceMirror::ResolvePreparedFilePath(snapshot, filePath, fullPath, relativePath, error)) {
		outError = error;
		return false;
	}

	std::string bytes;
	size_t effectiveByteOffset = 0;
	bool fileTruncated = false;
	size_t totalBytes = 0;
	if (!ReadFileBytesWindow(
			fullPath,
			byteOffset,
			bytes,
			effectiveByteOffset,
			fileTruncated,
			totalBytes,
			error)) {
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
	const size_t lineNumberBase = CountLineBreaksBefore(fullPath, effectiveByteOffset);
	const int totalLines = static_cast<int>(lines.size());
	const int effectiveOffset = ClampInt(offset, 0, totalLines);
	int returned = 0;
	bool lineTruncated = false;
	const std::string view = BuildNumberedView(
		lines,
		effectiveOffset,
		limit,
		returned,
		lineTruncated,
		lineNumberBase);

	outRow["ok"] = true;
	outRow["file_path"] = relativePath;
	outRow["mirror_generation"] = snapshot.generation;
	outRow["code_kind"] = "mirror_source";
	outRow["code_hash"] = BuildStableTextHashForRealCode(hashText);
	outRow["code_hash_complete"] = effectiveByteOffset == 0 && !fileTruncated;
	outRow["requested_source_byte_offset"] = byteOffset;
	outRow["source_byte_offset"] = effectiveByteOffset;
	outRow["line_number_base"] = lineNumberBase;
	outRow["total_lines"] = totalLines;
	outRow["total_lines_complete"] = effectiveByteOffset == 0 && !fileTruncated;
	outRow["returned_lines"] = returned;
	outRow["truncated"] = effectiveByteOffset > 0 || fileTruncated || effectiveOffset > 0 || lineTruncated;
	outRow["content"] = view;
	AppendLinePaginationMetadata(outRow, offset, effectiveOffset, returned, totalLines, lineNumberBase);
	AppendSourceByteMetadata(outRow, returnedBytes, totalBytes, fileTruncated, effectiveByteOffset);
	return true;
}

std::string ExecuteReadFile(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& argumentsJson,
	bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}
	std::string generationError;
	if (!ValidateMirrorGeneration(args, snapshot, generationError)) {
		return BuildError(generationError);
	}

	const std::string filePath = GetJsonStringUtf8(args, "file_path");
	const size_t byteOffset = GetJsonSize(args, "byte_offset", 0);
	const int offset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int limit = GetJsonInt(args, "limit", 0);

	json r;
	std::string error;
	if (!BuildReadFileRow(snapshot, filePath, byteOffset, offset, limit, r, error)) {
		return BuildError(error);
	}
	outOk = true;
	return ToolResultToLocalJson(r);
}

struct BatchReadRequest {
	std::string filePath;
	size_t byteOffset = 0;
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
				ToLowerAscii(NormalizeSlash(request.filePath)) &&
				item.byteOffset == request.byteOffset &&
				item.offset == request.offset &&
				item.limit == request.limit;
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
		{"byte_offset", request.byteOffset},
		{"offset", request.offset},
		{"limit", request.limit}
	};
}

std::string ExecuteReadFiles(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& argumentsJson,
	bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}
	std::string generationError;
	if (!ValidateMirrorGeneration(args, snapshot, generationError)) {
		return BuildError(generationError);
	}

	const int defaultOffset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const size_t defaultByteOffset = GetJsonSize(args, "byte_offset", 0);
	const int defaultLimit = ClampInt(GetJsonInt(args, "limit", kDefaultBatchReadLimit), 1, kDefaultReadLimit);
	std::vector<BatchReadRequest> requests;

	if (args.contains("file_paths") && args["file_paths"].is_array()) {
		for (const auto& item : args["file_paths"]) {
			if (!item.is_string()) {
				continue;
			}
			BatchReadRequest request;
			request.filePath = item.get<std::string>();
			request.byteOffset = defaultByteOffset;
			request.offset = defaultOffset;
			request.limit = defaultLimit;
			AddBatchReadRequest(requests, request);
		}
	}
	if (args.contains("files") && args["files"].is_array()) {
		for (const auto& item : args["files"]) {
			BatchReadRequest request;
			request.byteOffset = defaultByteOffset;
			request.offset = defaultOffset;
			request.limit = defaultLimit;
			if (item.is_string()) {
				request.filePath = item.get<std::string>();
			}
			else if (item.is_object()) {
				request.filePath = GetJsonStringUtf8(item, "file_path");
				request.byteOffset = GetJsonSize(item, "byte_offset", defaultByteOffset);
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
		if (BuildReadFileRow(snapshot, request.filePath, request.byteOffset, request.offset, effectiveLimit, row, error)) {
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
	r["status"] = okCount == 0
		? "error"
		: (errorCount == 0 && !outputTruncated ? "success" : "partial");
	r["code_kind"] = "mirror_source";
	r["mirror_generation"] = snapshot.generation;
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

std::string ExecuteListFiles(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& argumentsJson,
	bool& outOk)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}
	std::string generationError;
	if (!ValidateMirrorGeneration(args, snapshot, generationError)) {
		return BuildError(generationError);
	}

	const std::string glob = GetJsonStringUtf8(args, "glob");
	const std::string pathPrefix = NormalizeSlash(GetJsonStringUtf8(args, "path"));
	const int offset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int limit = ClampInt(GetJsonInt(args, "limit", kDefaultListLimit), 1, 5000);
	std::optional<std::regex> globRegex;
	std::string globError;
	if (!TryCompileGlob(glob, globRegex, globError)) {
		return BuildError(globError);
	}

	const std::vector<std::string>& files = snapshot.relativePathsUtf8;

	json rows = json::array();
	int matched = 0;
	for (const std::string& file : files) {
		if (!pathPrefix.empty() && NormalizeSlash(file).rfind(pathPrefix, 0) != 0) {
			continue;
		}
		if (!CompiledGlobMatches(file, globRegex)) {
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
	r["mirror_generation"] = snapshot.generation;
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
		return ToLowerAscii(line).find(pattern) != std::string::npos;
	}
	return line.find(pattern) != std::string::npos;
}

bool IsPotentiallyCatastrophicRegex(const std::string& pattern)
{
	if (pattern.size() > 256) {
		return true;
	}
	std::vector<bool> groupContainsQuantifier;
	bool escaped = false;
	bool inClass = false;
	for (size_t i = 0; i < pattern.size(); ++i) {
		const char ch = pattern[i];
		if (escaped) {
			if (std::isdigit(static_cast<unsigned char>(ch)) != 0) {
				return true;
			}
			escaped = false;
			continue;
		}
		if (ch == '\\') {
			escaped = true;
			continue;
		}
		if (ch == '[') {
			inClass = true;
			continue;
		}
		if (ch == ']' && inClass) {
			inClass = false;
			continue;
		}
		if (inClass) {
			continue;
		}
		if (ch == '(') {
			groupContainsQuantifier.push_back(false);
			continue;
		}
		if (ch == ')' && !groupContainsQuantifier.empty()) {
			const bool nestedQuantifier = groupContainsQuantifier.back();
			groupContainsQuantifier.pop_back();
			const char next = i + 1 < pattern.size() ? pattern[i + 1] : '\0';
			if (nestedQuantifier && (next == '*' || next == '+' || next == '?' || next == '{')) {
				return true;
			}
			continue;
		}
		if (ch == '*' || ch == '+' || ch == '?' || ch == '{') {
			for (auto&& groupFlag : groupContainsQuantifier) {
				groupFlag = true;
			}
		}
	}
	return false;
}

json BuildSearchErrorJson(const std::string& error)
{
	json r;
	r["ok"] = false;
	r["status"] = "error";
	r["error"] = error;
	return r;
}

struct SearchQueryRuntime {
	std::string pattern;
	std::string matchPattern;
	std::optional<std::regex> compiledRegex;
	json results = json::array();
	int filesWithMatches = 0;
	int totalMatches = 0;
	int totalResults = 0;
	bool active = false;
	std::string error;
};

struct SearchSourceStats {
	int skippedBinary = 0;
	int sourceFilesTruncated = 0;
	size_t sourceBytesOmitted = 0;
	int failedFileCount = 0;
	std::vector<std::string> failedFileSamples;
	bool cancelled = false;
};

struct PendingSearchResult {
	size_t queryIndex = 0;
	json row;
	int remainingContextLines = 0;
};

json FinalizeSearchQuery(
	SearchQueryRuntime& query,
	const SearchSourceStats& sourceStats,
	const std::string& outputMode,
	bool caseInsensitive,
	bool regexMode,
	int offset,
	int pageLimit)
{
	if (!query.active) {
		return BuildSearchErrorJson(query.error);
	}

	json r;
	r["ok"] = !sourceStats.cancelled;
	r["pattern"] = query.pattern;
	r["regex"] = regexMode;
	r["case_insensitive"] = caseInsensitive;
	r["output_mode"] = outputMode;
	r["results"] = std::move(query.results);
	r["files_with_matches"] = query.filesWithMatches;
	r["match_count"] = query.totalMatches;
	r["skipped_binary"] = sourceStats.skippedBinary;
	r["source_files_truncated"] = sourceStats.sourceFilesTruncated;
	r["source_bytes_omitted"] = sourceStats.sourceBytesOmitted;
	r["failed_file_count"] = sourceStats.failedFileCount;
	r["failed_file_samples"] = sourceStats.failedFileSamples;
	r["cancelled"] = sourceStats.cancelled;
	r["scan_complete"] = sourceStats.failedFileCount == 0 && !sourceStats.cancelled;
	r["results_complete"] = sourceStats.sourceFilesTruncated == 0 &&
		sourceStats.failedFileCount == 0 && !sourceStats.cancelled;
	r["status"] = sourceStats.cancelled
		? "partial"
		: (r["results_complete"].get<bool>() ? "success" : "partial");
	if (sourceStats.cancelled) {
		r["error"] = "search cancelled";
	}
	const int returned = static_cast<int>(r["results"].size());
	AppendResultPaginationMetadata(r, offset, returned, query.totalResults);
	r["page_limit"] = pageLimit;
	r["truncated"] = r["offset"].get<int>() > 0 ||
		r["has_more"].get<bool>() ||
		sourceStats.sourceFilesTruncated > 0;
	return r;
}

std::vector<json> ExecuteSearchCodeQueries(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::vector<std::string>& patterns,
	const std::string& glob,
	const std::string& outputMode,
	bool caseInsensitive,
	bool regexMode,
	int context,
	int offset,
	int pageLimit,
	const std::function<bool()>& cancelCallback = {})
{
	std::vector<SearchQueryRuntime> queries;
	queries.reserve(patterns.size());
	bool anyActive = false;
	for (const std::string& pattern : patterns) {
		SearchQueryRuntime query;
		query.pattern = pattern;
		if (pattern.empty()) {
			query.error = "pattern is required";
		}
		else if (outputMode != "content" && outputMode != "files_with_matches" && outputMode != "count") {
			query.error = "output_mode must be content, files_with_matches, or count";
		}
		else if (regexMode) {
			if (IsPotentiallyCatastrophicRegex(pattern)) {
				query.error = "regex pattern is too complex; use literal search or a simpler regex up to 256 bytes";
				queries.push_back(std::move(query));
				continue;
			}
			try {
				query.compiledRegex.emplace(
					pattern,
					std::regex::ECMAScript | (caseInsensitive ? std::regex::icase : std::regex::flag_type{}));
				query.active = true;
			}
			catch (const std::exception& ex) {
				query.error = std::string("invalid regex pattern: ") + ex.what();
			}
		}
		else {
			query.matchPattern = caseInsensitive ? ToLowerAscii(pattern) : pattern;
			query.active = true;
		}
		anyActive = anyActive || query.active;
		queries.push_back(std::move(query));
	}

	std::optional<std::regex> globRegex;
	std::string globError;
	if (anyActive && !TryCompileGlob(glob, globRegex, globError)) {
		for (SearchQueryRuntime& query : queries) {
			if (query.active) {
				query.active = false;
				query.error = globError;
			}
		}
		anyActive = false;
	}

	SearchSourceStats sourceStats;
	if (anyActive) {
		for (const std::string& file : snapshot.relativePathsUtf8) {
			if (cancelCallback && cancelCallback()) {
				sourceStats.cancelled = true;
				break;
			}
			if (!CompiledGlobMatches(file, globRegex)) {
				continue;
			}
			if (glob.empty() && !IsDefaultVisiblePath(file)) {
				continue;
			}

			std::filesystem::path fullPath;
			std::string resolvedRelative;
			std::string error;
			if (!WorkspaceMirror::ResolvePreparedFilePath(snapshot, file, fullPath, resolvedRelative, error)) {
				++sourceStats.failedFileCount;
				if (sourceStats.failedFileSamples.size() < 10) {
					sourceStats.failedFileSamples.push_back(file + ": " + error);
				}
				continue;
			}
			std::ifstream sourceFile(fullPath, std::ios::binary);
			if (!sourceFile) {
				++sourceStats.failedFileCount;
				if (sourceStats.failedFileSamples.size() < 10) {
					sourceStats.failedFileSamples.push_back(file + ": open file failed");
				}
				continue;
			}
			char sampleBuffer[4096] = {};
			sourceFile.read(sampleBuffer, static_cast<std::streamsize>(sizeof(sampleBuffer)));
			const std::streamsize sampleBytes = sourceFile.gcount();
			const std::string sample(
				sampleBuffer,
				sampleBytes > 0 ? static_cast<size_t>(sampleBytes) : 0);
			if (LooksBinary(sample)) {
				++sourceStats.skippedBinary;
				continue;
			}
			sourceFile.clear();
			sourceFile.seekg(0, std::ios::beg);
			if (!sourceFile) {
				++sourceStats.failedFileCount;
				if (sourceStats.failedFileSamples.size() < 10) {
					sourceStats.failedFileSamples.push_back(file + ": seek file failed");
				}
				continue;
			}

			std::vector<int> fileMatches(queries.size(), 0);
			std::vector<int> pendingCounts(queries.size(), 0);
			std::vector<std::pair<int, std::string>> previousContext;
			std::vector<PendingSearchResult> pendingResults;
			std::string rawLine;
			int lineIndex = 0;
			for (;;) {
				if ((lineIndex & 0xFF) == 0 && cancelCallback && cancelCallback()) {
					sourceStats.cancelled = true;
					break;
				}
				const std::streampos linePosition = sourceFile.tellg();
				if (!std::getline(sourceFile, rawLine)) {
					break;
				}
				if (!rawLine.empty() && rawLine.back() == '\r') {
					rawLine.pop_back();
				}
				std::string line = DecodeTextToUtf8(std::move(rawLine));
				rawLine.clear();
				++lineIndex;

				for (auto pendingIt = pendingResults.begin(); pendingIt != pendingResults.end();) {
					pendingIt->row["context"].push_back({
						{"line_number", lineIndex},
						{"text", line}
					});
					--pendingIt->remainingContextLines;
					if (pendingIt->remainingContextLines <= 0) {
						queries[pendingIt->queryIndex].results.push_back(std::move(pendingIt->row));
						--pendingCounts[pendingIt->queryIndex];
						pendingIt = pendingResults.erase(pendingIt);
					}
					else {
						++pendingIt;
					}
				}

				const std::string loweredLine = !regexMode && caseInsensitive
					? ToLowerAscii(line)
					: std::string();
				for (size_t queryIndex = 0; queryIndex < queries.size(); ++queryIndex) {
					SearchQueryRuntime& query = queries[queryIndex];
					const bool matched = query.active && (regexMode
						? LineMatches(line, query.matchPattern, true, caseInsensitive,
							query.compiledRegex ? &*query.compiledRegex : nullptr)
						: (caseInsensitive
							? loweredLine.find(query.matchPattern) != std::string::npos
							: line.find(query.matchPattern) != std::string::npos));
					if (!matched) {
						continue;
					}
					++fileMatches[queryIndex];
					++query.totalMatches;
					if (outputMode != "content") {
						continue;
					}
					const int resultOffset = query.totalResults++;
					if (resultOffset < offset ||
						static_cast<int>(query.results.size()) + pendingCounts[queryIndex] >= pageLimit) {
						continue;
					}
					json row;
					row["file_path"] = file;
					row["line_number"] = lineIndex;
					row["source_byte_offset"] = linePosition >= 0
						? json(static_cast<unsigned long long>(linePosition))
						: json(nullptr);
					row["text"] = line;
					if (context > 0) {
						json contextRows = json::array();
						for (const auto& [contextLineNumber, contextText] : previousContext) {
							contextRows.push_back({
								{"line_number", contextLineNumber},
								{"text", contextText}
							});
						}
						contextRows.push_back({
							{"line_number", lineIndex},
							{"text", line}
						});
						row["context"] = std::move(contextRows);
						pendingResults.push_back(PendingSearchResult{
							queryIndex,
							std::move(row),
							context
						});
						++pendingCounts[queryIndex];
					}
					else {
						query.results.push_back(std::move(row));
					}
				}

				if (context > 0) {
					previousContext.push_back({ lineIndex, line });
					if (previousContext.size() > static_cast<size_t>(context)) {
						previousContext.erase(previousContext.begin());
					}
				}
			}
			for (PendingSearchResult& pending : pendingResults) {
				queries[pending.queryIndex].results.push_back(std::move(pending.row));
			}
			if (sourceFile.bad()) {
				++sourceStats.failedFileCount;
				if (sourceStats.failedFileSamples.size() < 10) {
					sourceStats.failedFileSamples.push_back(file + ": read file failed");
				}
			}
			if (sourceStats.cancelled) {
				break;
			}

			for (size_t queryIndex = 0; queryIndex < queries.size(); ++queryIndex) {
				SearchQueryRuntime& query = queries[queryIndex];
				const int matchCount = fileMatches[queryIndex];
				if (!query.active || matchCount <= 0) {
					continue;
				}
				++query.filesWithMatches;
				if (outputMode == "content") {
					continue;
				}
				const int resultOffset = query.totalResults++;
				if (resultOffset < offset || static_cast<int>(query.results.size()) >= pageLimit) {
					continue;
				}
				if (outputMode == "files_with_matches") {
					query.results.push_back(file);
				}
				else {
					query.results.push_back({
						{"file_path", file},
						{"count", matchCount}
					});
				}
			}
		}
	}

	std::vector<json> rows;
	rows.reserve(queries.size());
	for (SearchQueryRuntime& query : queries) {
		rows.push_back(FinalizeSearchQuery(
			query,
			sourceStats,
			outputMode,
			caseInsensitive,
			regexMode,
			offset,
			pageLimit));
	}
	return rows;
}

json ExecuteSearchCodeQuery(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& pattern,
	const std::string& glob,
	const std::string& outputMode,
	bool caseInsensitive,
	bool regexMode,
	int context,
	int offset,
	int pageLimit,
	const std::function<bool()>& cancelCallback = {})
{
	std::vector<json> rows = ExecuteSearchCodeQueries(
		snapshot,
		{ pattern },
		glob,
		outputMode,
		caseInsensitive,
		regexMode,
		context,
		offset,
		pageLimit,
		cancelCallback);
	return rows.empty() ? BuildSearchErrorJson("pattern is required") : std::move(rows.front());
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

json BuildStreamingCodeItemResult(
	const std::filesystem::path& fullPath,
	const std::string& relativePath,
	const std::string& itemName,
	int occurrence,
	std::uint64_t mirrorGeneration,
	const std::function<bool()>& cancelCallback = {})
{
	struct CapturedItem {
		std::string directive;
		int startLine = 0;
		int endLine = 0;
		std::string content;
		std::string rawContent;
	};
	std::ifstream file(fullPath, std::ios::binary);
	if (!file) {
		return BuildSearchErrorJson("open file failed");
	}
	std::vector<CapturedItem> captured;
	std::optional<CapturedItem> active;
	std::string rawLine;
	int lineNumber = 0;
	while (std::getline(file, rawLine)) {
		if ((lineNumber & 0xFF) == 0 && cancelCallback && cancelCallback()) {
			json r = BuildSearchErrorJson("code item read cancelled");
			r["cancelled"] = true;
			return r;
		}
		if (!rawLine.empty() && rawLine.back() == '\r') {
			rawLine.pop_back();
		}
		const std::string line = DecodeTextToUtf8(std::move(rawLine));
		rawLine.clear();
		++lineNumber;
		std::string directive;
		std::string name;
		const bool declaration = TryParseTopLevelDeclaration(line, directive, name);
		if (declaration && active.has_value()) {
			active->endLine = lineNumber - 1;
			captured.push_back(std::move(*active));
			active.reset();
		}
		if (declaration && ToLowerAscii(name) == ToLowerAscii(itemName)) {
			active.emplace();
			active->directive = std::move(directive);
			active->startLine = lineNumber;
		}
		if (active.has_value()) {
			active->rawContent += line;
			active->rawContent += "\r\n";
			active->content += std::to_string(lineNumber);
			active->content += "\t";
			active->content += line;
			active->content += "\n";
		}
	}
	if (file.bad()) {
		return BuildSearchErrorJson("read file failed");
	}
	if (active.has_value()) {
		active->endLine = lineNumber;
		captured.push_back(std::move(*active));
	}
	if (captured.empty()) {
		json r = BuildSearchErrorJson("code item not found");
		r["file_path"] = relativePath;
		r["item_name"] = itemName;
		r["search_complete"] = true;
		return r;
	}
	if (captured.size() > 1 && occurrence == 0) {
		json candidates = json::array();
		for (const CapturedItem& item : captured) {
			candidates.push_back({
				{"line_number", item.startLine},
				{"directive", item.directive}
			});
		}
		json r = BuildSearchErrorJson("code item name matched multiple declarations");
		r["file_path"] = relativePath;
		r["item_name"] = itemName;
		r["candidates"] = std::move(candidates);
		r["search_complete"] = true;
		return r;
	}
	if (occurrence > static_cast<int>(captured.size())) {
		json r = BuildSearchErrorJson("occurrence exceeds the number of matching declarations");
		r["file_path"] = relativePath;
		r["item_name"] = itemName;
		r["match_count"] = captured.size();
		return r;
	}
	const size_t selectedIndex = occurrence > 0 ? static_cast<size_t>(occurrence - 1) : 0;
	const CapturedItem& item = captured[selectedIndex];
	json r;
	r["ok"] = true;
	r["status"] = "success";
	r["file_path"] = relativePath;
	r["mirror_generation"] = mirrorGeneration;
	r["code_kind"] = "mirror_source";
	r["code_hash_complete"] = false;
	r["item_hash"] = BuildStableTextHashForRealCode(Utf8ToLocalText(item.rawContent));
	r["item_name"] = itemName;
	r["directive"] = item.directive;
	r["occurrence"] = selectedIndex + 1;
	r["match_count"] = captured.size();
	r["start_line"] = item.startLine;
	r["end_line"] = item.endLine;
	r["returned_lines"] = item.endLine - item.startLine + 1;
	r["total_lines"] = lineNumber;
	r["total_lines_complete"] = true;
	r["item_complete"] = true;
	r["truncated"] = false;
	r["content"] = item.content;
	return r;
}

std::string ExecuteReadCodeItem(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& argumentsJson,
	bool& outOk,
	const std::function<bool()>& cancelCallback)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}
	std::string generationError;
	if (!ValidateMirrorGeneration(args, snapshot, generationError)) {
		return BuildError(generationError);
	}

	const std::string filePath = GetJsonStringUtf8(args, "file_path");
	const std::string itemName = TrimAsciiCopy(GetJsonStringUtf8(args, "item_name"));
	const int occurrence = (std::max)(0, GetJsonInt(args, "occurrence", 0));
	if (filePath.empty() || itemName.empty()) {
		return BuildError("read_code_item requires file_path and item_name");
	}

	std::filesystem::path fullPath;
	std::string relativePath;
	std::string error;
	if (!WorkspaceMirror::ResolvePreparedFilePath(snapshot, filePath, fullPath, relativePath, error)) {
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
	if (fileTruncated) {
		json streamed = BuildStreamingCodeItemResult(
			fullPath,
			relativePath,
			itemName,
			occurrence,
			snapshot.generation,
			cancelCallback);
		if (streamed.value("ok", false) && GetJsonBool(args, "include_references", false)) {
			const std::string referenceGlob = GetJsonStringUtf8(args, "reference_glob");
			const int referenceLimit = ClampInt(GetJsonInt(args, "reference_limit", 20), 1, 50);
			streamed["references"] = ExecuteSearchCodeQuery(
				snapshot,
				itemName,
				referenceGlob,
				"content",
				true,
				false,
				1,
				0,
				referenceLimit,
				cancelCallback);
		}
		outOk = streamed.value("ok", false);
		return ToolResultToLocalJson(streamed);
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
	if (matches.size() > 1 && occurrence == 0) {
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
	if (occurrence > static_cast<int>(matches.size())) {
		json r;
		r["ok"] = false;
		r["error"] = "occurrence exceeds the number of matching declarations";
		r["file_path"] = relativePath;
		r["item_name"] = itemName;
		r["match_count"] = matches.size();
		return ToolResultToLocalJson(r);
	}

	const size_t selectedIndex = occurrence > 0 ? static_cast<size_t>(occurrence - 1) : 0;
	const int start = matches[selectedIndex];
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
	r["mirror_generation"] = snapshot.generation;
	r["code_kind"] = "mirror_source";
	r["code_hash"] = BuildStableTextHashForRealCode(hashText);
	r["code_hash_complete"] = !fileTruncated;
	r["item_name"] = itemName;
	r["directive"] = matchDirectives[selectedIndex];
	r["occurrence"] = selectedIndex + 1;
	r["match_count"] = matches.size();
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
			snapshot,
			itemName,
			referenceGlob,
			"content",
			true,
			false,
			1,
			0,
			referenceLimit,
			cancelCallback);
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

std::string ExecuteSearchCode(
	const WorkspaceMirror::FileAccessSnapshot& snapshot,
	const std::string& argumentsJson,
	bool& outOk,
	const std::function<bool()>& cancelCallback)
{
	outOk = false;
	json args;
	try {
		args = argumentsJson.empty() ? json::object() : json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		return BuildError(std::string("invalid arguments json: ") + ex.what());
	}
	std::string generationError;
	if (!ValidateMirrorGeneration(args, snapshot, generationError)) {
		return BuildError(generationError);
	}

	const std::string glob = GetJsonStringUtf8(args, "glob");
	const std::string outputMode = GetJsonStringUtf8(args, "output_mode").empty()
		? "content"
		: GetJsonStringUtf8(args, "output_mode");
	const bool caseInsensitive = GetJsonBool(args, "case_insensitive", false);
	const bool regexMode = GetJsonBool(args, "regex", false);
	const int context = ClampInt(GetJsonInt(args, "context", 0), 0, 20);
	const int offset = (std::max)(0, GetJsonInt(args, "offset", 0));
	const int headLimit = ClampInt(GetJsonInt(args, "head_limit", kDefaultSearchLimit), 1, 2000);
	bool patternListTruncated = false;
	const std::vector<std::string> patterns = GetSearchPatterns(args, patternListTruncated);
	if (patterns.empty()) {
		return BuildError("pattern or patterns is required");
	}
	std::vector<json> queryResults = ExecuteSearchCodeQueries(
		snapshot,
		patterns,
		glob,
		outputMode,
		caseInsensitive,
		regexMode,
		context,
		offset,
		headLimit,
		cancelCallback);

	if (queryResults.size() == 1 && !patternListTruncated) {
		json r = std::move(queryResults.front());
		r["mirror_generation"] = snapshot.generation;
		outOk = r.value("ok", false);
		return ToolResultToLocalJson(r);
	}

	json queries = json::array();
	bool anyOk = false;
	int okQueryCount = 0;
	int errorQueryCount = 0;
	bool anyTruncated = patternListTruncated;
	bool anyHasMore = false;
	json continuations = json::array();
	for (json& row : queryResults) {
		const bool rowOk = row.value("ok", false);
		anyOk = anyOk || rowOk;
		okQueryCount += rowOk ? 1 : 0;
		errorQueryCount += rowOk ? 0 : 1;
		anyTruncated = anyTruncated || row.value("truncated", false);
		anyHasMore = anyHasMore || row.value("has_more", false);
		if (row.value("has_more", false)) {
			continuations.push_back({
				{"pattern", row.value("pattern", std::string())},
				{"next_offset", row.contains("next_offset") ? row["next_offset"] : json(nullptr)},
				{"mirror_generation", snapshot.generation}
			});
		}
		queries.push_back(std::move(row));
	}

	json r;
	r["ok"] = anyOk;
	r["batch"] = true;
	r["mirror_generation"] = snapshot.generation;
	r["regex"] = regexMode;
	r["case_insensitive"] = caseInsensitive;
	r["output_mode"] = outputMode;
	r["glob"] = glob;
	r["offset"] = offset;
	r["page_limit"] = headLimit;
	r["queries"] = std::move(queries);
	r["requested"] = patterns.size();
	r["returned"] = r["queries"].size();
	r["ok_count"] = okQueryCount;
	r["error_count"] = errorQueryCount;
	r["all_ok"] = errorQueryCount == 0 && !patternListTruncated;
	r["status"] = !anyOk
		? "error"
		: (errorQueryCount == 0 && !patternListTruncated ? "success" : "partial");
	r["pattern_list_truncated"] = patternListTruncated;
	r["has_more"] = anyHasMore;
	r["continuations"] = std::move(continuations);
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

std::string ExecuteTool(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk,
	const std::function<bool()>& cancelCallback)
{
	WorkspaceMirror::FileAccessSnapshot snapshot;
	std::string snapshotError;
	if (!WorkspaceMirror::GetPreparedFileAccessSnapshot(snapshot, snapshotError)) {
		outOk = false;
		return BuildError(snapshotError);
	}
	if (toolName == "read_file") {
		return ExecuteReadFile(snapshot, argumentsJson, outOk);
	}
	if (toolName == "read_files") {
		return ExecuteReadFiles(snapshot, argumentsJson, outOk);
	}
	if (toolName == "read_code_item") {
		return ExecuteReadCodeItem(snapshot, argumentsJson, outOk, cancelCallback);
	}
	if (toolName == "list_files") {
		return ExecuteListFiles(snapshot, argumentsJson, outOk);
	}
	if (toolName == "search_code") {
		return ExecuteSearchCode(snapshot, argumentsJson, outOk, cancelCallback);
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

	bool byteCursorOk = false;
	bool fullSearchOk = false;
	bool streamingCodeItemOk = false;
	bool staleGenerationRejected = false;
	std::error_code tempError;
	const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(tempError) /
		std::format("autolinker_workspace_file_test_{}_{}", GetCurrentProcessId(), GetTickCount64());
	if (!tempError) {
		const std::filesystem::path sourcePath = tempRoot / "src" / "Large.txt";
		std::filesystem::create_directories(sourcePath.parent_path(), tempError);
		if (!tempError) {
			std::ofstream file(sourcePath, std::ios::binary | std::ios::trunc);
			for (int i = 0; i < 15000; ++i) {
				file << std::string(80, 'x') << "\n";
			}
			file << ".子程序 TailProc\n";
			file << "TAIL_MARKER_UNIQUE\n";
			file << ".子程序 NextProc\n";
			file.close();

			WorkspaceMirror::FileAccessSnapshot snapshot;
			snapshot.mirrorRoot = tempRoot;
			snapshot.relativePathsUtf8 = { "src/Large.txt" };
			snapshot.generation = 42;
			json firstRow;
			std::string rowError;
			if (BuildReadFileRow(snapshot, "src/Large.txt", 0, 0, 20000, firstRow, rowError) &&
				firstRow.value("has_more_source_bytes", false)) {
				const size_t nextByteOffset = firstRow.value("next_source_byte_offset", static_cast<size_t>(0));
				json secondRow;
				byteCursorOk = nextByteOffset > 0 &&
					BuildReadFileRow(snapshot, "src/Large.txt", nextByteOffset, 0, 20000, secondRow, rowError) &&
					secondRow.value("content", std::string()).find("TAIL_MARKER_UNIQUE") != std::string::npos;
			}

			std::vector<json> searchRows = ExecuteSearchCodeQueries(
				snapshot,
				{ "TAIL_MARKER_UNIQUE" },
				"",
				"content",
				false,
				false,
				0,
				0,
				10);
			fullSearchOk = searchRows.size() == 1 &&
				searchRows.front().value("ok", false) &&
				searchRows.front().value("match_count", 0) == 1 &&
				searchRows.front()["results"].is_array() &&
				!searchRows.front()["results"].empty() &&
				searchRows.front()["results"][0].value("source_byte_offset", 0ull) > 1024ull * 1024ull;

			const json codeItem = BuildStreamingCodeItemResult(
				sourcePath,
				"src/Large.txt",
				"TailProc",
				0,
				snapshot.generation);
			streamingCodeItemOk = codeItem.value("ok", false) &&
				codeItem.value("content", std::string()).find("TAIL_MARKER_UNIQUE") != std::string::npos;

			std::string generationError;
			staleGenerationRejected = !ValidateMirrorGeneration(
				json({{"mirror_generation", 41}}),
				snapshot,
				generationError);
		}
		std::filesystem::remove_all(tempRoot, tempError);
	}

	return json({
		{"name", "workspace-file-pagination"},
		{"ok", resultPageOk && linePageOk && sourcePageOk && byteCursorOk &&
			fullSearchOk && streamingCodeItemOk && staleGenerationRejected},
		{"result_page", std::move(resultPage)},
		{"line_page", std::move(linePage)},
		{"source_page", std::move(sourcePage)},
		{"byte_cursor_ok", byteCursorOk},
		{"full_search_ok", fullSearchOk},
		{"streaming_code_item_ok", streamingCodeItemOk},
		{"stale_generation_rejected", staleGenerationRejected}
	}).dump();
}

} // namespace WorkspaceFileTools
