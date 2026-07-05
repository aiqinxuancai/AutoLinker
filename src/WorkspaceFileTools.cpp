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

bool ReadFileBytesLimited(const std::filesystem::path& path, std::string& outBytes, bool& outTruncated, std::string& outError)
{
	outBytes.clear();
	outTruncated = false;
	outError.clear();

	std::ifstream file(path, std::ios::binary);
	if (!file) {
		outError = "open file failed";
		return false;
	}
	outBytes.assign(std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>());
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
	if (!ReadFileBytesLimited(fullPath, bytes, fileTruncated, error)) {
		outError = error;
		return false;
	}
	if (LooksBinary(bytes)) {
		outError = "read_file only supports text files: " + relativePath;
		return false;
	}

	const std::string text = DecodeTextToUtf8(std::move(bytes));
	const std::string hashText = NormalizeRealCodeLineBreaksToCrLf(Utf8ToLocalText(text));
	const auto lines = SplitLinesUtf8(text);
	int returned = 0;
	bool lineTruncated = false;
	const std::string view = BuildNumberedView(lines, offset, limit, returned, lineTruncated);

	outRow["ok"] = true;
	outRow["file_path"] = relativePath;
	outRow["code_kind"] = "mirror_source";
	outRow["code_hash"] = BuildStableTextHashForRealCode(hashText);
	outRow["total_lines"] = lines.size();
	outRow["offset"] = offset;
	outRow["returned_lines"] = returned;
	outRow["truncated"] = fileTruncated || lineTruncated;
	outRow["content"] = view;
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
	const BatchReadRequest& request,
	bool& outRequestTruncated)
{
	if (TrimAsciiCopy(request.filePath).empty()) {
		return;
	}
	const auto exists = std::find_if(
		requests.begin(),
		requests.end(),
		[&request](const BatchReadRequest& item) {
			return NormalizeSlash(item.filePath) == NormalizeSlash(request.filePath);
		});
	if (exists != requests.end()) {
		return;
	}
	if (static_cast<int>(requests.size()) >= kMaxBatchReadFiles) {
		outRequestTruncated = true;
		return;
	}
	requests.push_back(request);
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
	bool requestTruncated = false;

	if (args.contains("file_paths") && args["file_paths"].is_array()) {
		for (const auto& item : args["file_paths"]) {
			if (!item.is_string()) {
				continue;
			}
			BatchReadRequest request;
			request.filePath = item.get<std::string>();
			request.offset = defaultOffset;
			request.limit = defaultLimit;
			AddBatchReadRequest(requests, request, requestTruncated);
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
			AddBatchReadRequest(requests, request, requestTruncated);
		}
	}

	if (requests.empty()) {
		return BuildError("read_files requires file_paths or files");
	}

	json rows = json::array();
	int okCount = 0;
	int errorCount = 0;
	int returnedLines = 0;
	bool outputTruncated = requestTruncated;

	for (const BatchReadRequest& request : requests) {
		if (returnedLines >= kMaxBatchReadTotalLines) {
			outputTruncated = true;
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
		rows.push_back(std::move(row));
	}

	json r;
	r["ok"] = okCount > 0;
	r["all_ok"] = errorCount == 0 && !outputTruncated;
	r["code_kind"] = "mirror_source";
	r["files"] = std::move(rows);
	r["requested"] = requests.size();
	r["returned"] = r["files"].size();
	r["ok_count"] = okCount;
	r["error_count"] = errorCount;
	r["returned_lines"] = returnedLines;
	r["truncated"] = outputTruncated;
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
		++matched;
		if (static_cast<int>(rows.size()) < limit) {
			rows.push_back(file);
		}
	}

	json r;
	r["ok"] = true;
	r["files"] = std::move(rows);
	r["count"] = matched;
	r["returned"] = r["files"].size();
	r["truncated"] = matched > limit;
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
	int headLimit)
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
	int skippedBinary = 0;
	bool truncated = false;

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
		if (!ReadFileBytesLimited(fullPath, bytes, fileTruncated, error)) {
			continue;
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
			if (outputMode == "content" && static_cast<int>(results.size()) < headLimit) {
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
		if (outputMode == "files_with_matches" && static_cast<int>(results.size()) < headLimit) {
			results.push_back(file);
		}
		else if (outputMode == "count" && static_cast<int>(results.size()) < headLimit) {
			results.push_back({
				{"file_path", file},
				{"count", fileMatches}
			});
		}
		if (static_cast<int>(results.size()) >= headLimit) {
			truncated = true;
			if (outputMode != "content") {
				break;
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
	r["truncated"] = truncated;
	return r;
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
	const int headLimit = ClampInt(GetJsonInt(args, "head_limit", kDefaultSearchLimit), 1, 2000);
	bool patternListTruncated = false;
	const std::vector<std::string> patterns = GetSearchPatterns(args, patternListTruncated);
	if (patterns.empty()) {
		return BuildError("pattern or patterns is required");
	}

	if (patterns.size() == 1 && !patternListTruncated) {
		const json r = ExecuteSearchCodeQuery(patterns.front(), glob, outputMode, caseInsensitive, regexMode, context, headLimit);
		outOk = r.value("ok", false);
		return ToolResultToLocalJson(r);
	}

	json queries = json::array();
	bool anyOk = false;
	bool anyTruncated = patternListTruncated;
	for (const std::string& pattern : patterns) {
		json row = ExecuteSearchCodeQuery(pattern, glob, outputMode, caseInsensitive, regexMode, context, headLimit);
		anyOk = anyOk || row.value("ok", false);
		anyTruncated = anyTruncated || row.value("truncated", false);
		queries.push_back(std::move(row));
	}

	json r;
	r["ok"] = anyOk;
	r["batch"] = true;
	r["regex"] = regexMode;
	r["case_insensitive"] = caseInsensitive;
	r["output_mode"] = outputMode;
	r["glob"] = glob;
	r["queries"] = std::move(queries);
	r["requested"] = patterns.size();
	r["returned"] = r["queries"].size();
	r["truncated"] = anyTruncated;
	outOk = anyOk;
	return ToolResultToLocalJson(r);
}

} // namespace

bool CanHandleTool(const std::string& toolName)
{
	return toolName == "read_file" || toolName == "read_files" || toolName == "list_files" || toolName == "search_code";
}

std::string ExecuteTool(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	if (toolName == "read_file") {
		return ExecuteReadFile(argumentsJson, outOk);
	}
	if (toolName == "read_files") {
		return ExecuteReadFiles(argumentsJson, outOk);
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

} // namespace WorkspaceFileTools
