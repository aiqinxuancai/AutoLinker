#include "ProjectSourceCacheManager.h"

#include <cctype>
#include <chrono>
#include <filesystem>
#include <format>
#include <mutex>
#include <string>
#include <system_error>

#include "..\\thirdparty\\json.hpp"

#include "EideProjectBinarySerializer.h"
#include "Global.h"
#include "IDEFacade.h"
#include "WindowHelper.h"
#include "e2txt.h"

namespace project_source_cache {
namespace {

struct CacheState {
	std::mutex mutex;
	Snapshot snapshot;
	uint64_t nextRevision = 1;
	uint64_t lastRefreshTraceId = 0;
	std::chrono::steady_clock::time_point lastRefreshAt = {};
};

CacheState g_cacheState;

constexpr auto kRefreshReuseWindow = std::chrono::milliseconds(1500);
constexpr char kProjectHitTokenPrefix[] = "e2txt_project_v1:";

std::string JoinTraceParts(const std::vector<std::string>& traces)
{
	std::string text;
	for (size_t i = 0; i < traces.size(); ++i) {
		if (i != 0) {
			text += "|";
		}
		text += traces[i];
	}
	return text;
}

std::string TrimAsciiCopyForProjectCache(const std::string& text)
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

bool IsLikelyClassModulePageNameForProjectCache(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopyForProjectCache(text);
	return trimmed.rfind("Class_", 0) == 0 || trimmed.rfind("类", 0) == 0;
}

bool IsValidUtf8TextForProjectCache(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0) > 0;
}

std::string ConvertCodePageForProjectCache(
	const std::string& text,
	UINT fromCodePage,
	UINT toCodePage,
	DWORD fromFlags = 0)
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
			wide.data(),
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
			out.data(),
			outLen,
			nullptr,
			nullptr) <= 0) {
		return text;
	}
	return out;
}

std::string LocalToUtf8TextForProjectCache(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8TextForProjectCache(text)) {
		return text;
	}
	return ConvertCodePageForProjectCache(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocalTextForProjectCache(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8TextForProjectCache(text)) {
		return text;
	}
	return ConvertCodePageForProjectCache(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string NormalizePathForProjectCache(const std::string& pathText)
{
	if (pathText.empty()) {
		return std::string();
	}

	try {
		std::filesystem::path path(pathText);
		path = path.lexically_normal();
		if (path.is_relative()) {
			path = std::filesystem::absolute(path);
		}
		path = path.lexically_normal();
		return path.string();
	}
	catch (...) {
		return pathText;
	}
}

std::string ResolveCurrentSourcePathForProjectCache()
{
	std::string sourcePath = TrimAsciiCopyForProjectCache(g_nowOpenSourceFilePath);
	if (sourcePath.empty()) {
		sourcePath = TrimAsciiCopyForProjectCache(GetSourceFilePath());
	}

	const std::string normalized = NormalizePathForProjectCache(sourcePath);
	if (!normalized.empty()) {
		g_nowOpenSourceFilePath = normalized;
	}
	return normalized;
}

bool IsProjectSourcePath(const std::string& sourcePath)
{
	if (sourcePath.empty()) {
		return false;
	}

	try {
		const std::filesystem::path path(sourcePath);
		std::string ext = path.extension().string();
		for (char& ch : ext) {
			ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
		}
		return ext == ".e";
	}
	catch (...) {
		return false;
	}
}

std::string BuildPageTypeKey(const e2txt::Page& page)
{
	const std::string typeName = TrimAsciiCopyForProjectCache(page.typeName);
	const std::string pageName = TrimAsciiCopyForProjectCache(page.name);
	if (typeName == "程序集") {
		return IsLikelyClassModulePageNameForProjectCache(pageName) ? "class_module" : "assembly";
	}
	if (typeName == "全局变量") {
		return "global_var";
	}
	if (typeName == "自定义数据类型") {
		return "user_data_type";
	}
	if (typeName == "DLL命令") {
		return "dll_command";
	}
	if (typeName == "常量资源") {
		return "const_resource";
	}
	if (typeName == "图片资源") {
		return "picture_resource";
	}
	if (typeName == "声音资源") {
		return "sound_resource";
	}
	if (typeName == "窗口/表单") {
		return "form";
	}
	return "unknown";
}

Snapshot BuildSnapshotFromDocument(
	const e2txt::Document& document,
	const std::string& sourcePath,
	const std::string& snapshotPath,
	const std::string& parsedInputPath,
	const std::string& parsedInputKind,
	uint64_t revision)
{
	Snapshot snapshot;
	snapshot.ok = true;
	snapshot.revision = revision;
	snapshot.sourcePath = sourcePath;
	snapshot.snapshotPath = snapshotPath;
	snapshot.parsedInputPath = parsedInputPath;
	snapshot.parsedInputKind = parsedInputKind;
	snapshot.projectName = document.projectName;
	snapshot.versionText = document.versionText;
	snapshot.pages.reserve(document.pages.size());

	for (size_t index = 0; index < document.pages.size(); ++index) {
		const auto& page = document.pages[index];
		Page item;
		item.pageIndex = static_cast<int>(index);
		item.typeName = page.typeName;
		item.typeKey = BuildPageTypeKey(page);
		item.name = page.name;
		item.lines = page.lines;
		snapshot.pages.push_back(std::move(item));
	}

	return snapshot;
}

bool TryGenerateDocumentSnapshot(
	const std::string& sourcePath,
	bool preferSaveSnapshot,
	Snapshot& outSnapshot,
	bool& outUsedSaveSnapshot,
	std::string& outTrace,
	std::string& outError)
{
	outSnapshot = {};
	outUsedSaveSnapshot = false;
	outTrace.clear();
	outError.clear();

	const std::string normalizedSourcePath = NormalizePathForProjectCache(sourcePath);
	if (!IsProjectSourcePath(normalizedSourcePath)) {
		outError = "current source is not an .e file";
		outTrace = "source_not_e_file";
		return false;
	}

	std::vector<std::string> traces;
	e2txt::Generator generator;
	e2txt::Document document;
	std::string parseInputPath;
	std::string parseInputKind;
	std::string lastParseError;

	const auto tryParseBytes = [&](
		const std::vector<std::uint8_t>& inputBytes,
		const std::string& inputKind,
		const std::string& tag)
		-> bool {
		document = {};
		lastParseError.clear();

		if (!generator.GenerateDocumentFromBytes(inputBytes, normalizedSourcePath, document, &lastParseError)) {
			if (lastParseError.empty()) {
				lastParseError = "e2txt generate document failed";
			}
			traces.push_back(std::format(
				"{}_parse_failed kind={} source={} bytes={} error={}",
				tag,
				inputKind,
				normalizedSourcePath,
				inputBytes.size(),
				lastParseError));
			return false;
		}

		parseInputPath = normalizedSourcePath;
		parseInputKind = inputKind;
		traces.push_back(std::format(
			"{}_parse_ok kind={} source={} bytes={} pages={}",
			tag,
			inputKind,
			normalizedSourcePath,
			inputBytes.size(),
			document.pages.size()));
		return true;
	};

	std::vector<std::uint8_t> memoryBytes;
	std::string memorySerializeError;
	std::string memorySerializeTrace;
	if (e571::ProjectBinarySerializer::Instance().SerializeCurrentProject(
			memoryBytes,
			&memorySerializeError,
			&memorySerializeTrace)) {
		traces.push_back(std::format(
			"memory_serialize_ok bytes={} trace={}",
			memoryBytes.size(),
			memorySerializeTrace));
		if (tryParseBytes(memoryBytes, "memory_serialize_bytes", "memory_serialize")) {
			outUsedSaveSnapshot = false;
		}
	}
	else {
		traces.push_back(std::format(
			"memory_serialize_failed error={} trace={}",
			memorySerializeError,
			memorySerializeTrace));
	}

	if (parseInputPath.empty()) {
		(void)preferSaveSnapshot;
		(void)lastParseError;
		outError = memorySerializeError.empty()
			? "memory serialize current project failed"
			: memorySerializeError;
		traces.push_back("fallback_disabled memory_only_required");
		outTrace = JoinTraceParts(traces);
		return false;
	}

	uint64_t revision = 0;
	{
		std::lock_guard<std::mutex> lock(g_cacheState.mutex);
		revision = g_cacheState.nextRevision++;
	}

	outSnapshot = BuildSnapshotFromDocument(
		document,
		normalizedSourcePath,
		std::string(),
		parseInputPath,
		parseInputKind,
		revision);

	traces.push_back(std::format(
		"parse_input={} parse_kind={} pages={} revision={}",
		parseInputPath,
		parseInputKind,
		outSnapshot.pages.size(),
		outSnapshot.revision));
	outTrace = JoinTraceParts(traces);
	return true;
}

bool ShouldReuseCurrentSnapshotLocked(
	const std::string& sourcePath,
	bool forceRefresh,
	uint64_t traceId)
{
	if (!g_cacheState.snapshot.ok) {
		return false;
	}
	if (NormalizePathForProjectCache(g_cacheState.snapshot.sourcePath) != NormalizePathForProjectCache(sourcePath)) {
		return false;
	}
	if (!forceRefresh) {
		return true;
	}
	if (traceId != 0 && traceId == g_cacheState.lastRefreshTraceId) {
		return true;
	}
	if (g_cacheState.lastRefreshAt.time_since_epoch().count() == 0) {
		return false;
	}
	return (std::chrono::steady_clock::now() - g_cacheState.lastRefreshAt) <= kRefreshReuseWindow;
}

bool TryCopyCurrentSnapshotLocked(Snapshot& outSnapshot)
{
	if (!g_cacheState.snapshot.ok) {
		return false;
	}
	outSnapshot = g_cacheState.snapshot;
	return true;
}

} // namespace

ProjectSourceCacheManager& ProjectSourceCacheManager::Instance()
{
	static ProjectSourceCacheManager instance;
	return instance;
}

bool ProjectSourceCacheManager::WarmupCurrentSource(std::string* outError, std::string* outTrace)
{
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	const std::string sourcePath = ResolveCurrentSourcePathForProjectCache();
	if (!IsProjectSourcePath(sourcePath)) {
		if (outError != nullptr) {
			*outError = "current source is not an .e file";
		}
		if (outTrace != nullptr) {
			*outTrace = "warmup_skip_not_e_file";
		}
		return false;
	}

	{
		std::lock_guard<std::mutex> lock(g_cacheState.mutex);
		if (ShouldReuseCurrentSnapshotLocked(sourcePath, false, 0)) {
			if (outTrace != nullptr) {
				*outTrace = std::format(
					"warmup_reuse_existing revision={} pages={}",
					g_cacheState.snapshot.revision,
					g_cacheState.snapshot.pages.size());
			}
			return true;
		}
	}

	Snapshot snapshot;
	bool usedSaveSnapshot = false;
	std::string trace;
	std::string error;
	if (!TryGenerateDocumentSnapshot(sourcePath, false, snapshot, usedSaveSnapshot, trace, error)) {
		if (outError != nullptr) {
			*outError = error;
		}
		if (outTrace != nullptr) {
			*outTrace = trace;
		}
		return false;
	}

	{
		std::lock_guard<std::mutex> lock(g_cacheState.mutex);
		g_cacheState.snapshot = snapshot;
		g_cacheState.lastRefreshTraceId = 0;
		g_cacheState.lastRefreshAt = std::chrono::steady_clock::now();
	}

	OutputStringToELog(std::format(
		"[ProjectSourceCacheWarmup] warmup_ok source={} pages={} parse_kind={} revision={}",
		snapshot.sourcePath,
		snapshot.pages.size(),
		snapshot.parsedInputKind,
		snapshot.revision));
	if (outTrace != nullptr) {
		*outTrace = trace;
	}
	return true;
}

bool ProjectSourceCacheManager::EnsureCurrentSourceLatest(
	Snapshot& outSnapshot,
	bool forceRefresh,
	bool* outRefreshed,
	std::string* outError,
	std::string* outTrace)
{
	outSnapshot = {};
	if (outRefreshed != nullptr) {
		*outRefreshed = false;
	}
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	const std::string sourcePath = ResolveCurrentSourcePathForProjectCache();
	if (!IsProjectSourcePath(sourcePath)) {
		if (outError != nullptr) {
			*outError = "current source is not an .e file";
		}
		if (outTrace != nullptr) {
			*outTrace = "current_source_not_e_file";
		}
		return false;
	}

	const uint64_t traceId = GetCurrentAIPerfTraceId();
	{
		std::lock_guard<std::mutex> lock(g_cacheState.mutex);
		if (ShouldReuseCurrentSnapshotLocked(sourcePath, forceRefresh, traceId) &&
			TryCopyCurrentSnapshotLocked(outSnapshot)) {
			if (outTrace != nullptr) {
				*outTrace = std::format(
					"reuse_snapshot revision={} pages={} forceRefresh={} traceId={}",
					outSnapshot.revision,
					outSnapshot.pages.size(),
					forceRefresh ? 1 : 0,
					traceId);
			}
			return true;
		}
	}

	Snapshot snapshot;
	bool usedSaveSnapshot = false;
	std::string trace;
	std::string error;
	if (!TryGenerateDocumentSnapshot(sourcePath, true, snapshot, usedSaveSnapshot, trace, error)) {
		if (!forceRefresh) {
			Snapshot existing;
			{
				std::lock_guard<std::mutex> lock(g_cacheState.mutex);
				if (TryCopyCurrentSnapshotLocked(existing)) {
					outSnapshot = existing;
					if (outTrace != nullptr) {
						*outTrace = "generate_latest_failed_fallback_existing|" + trace;
					}
					if (outError != nullptr) {
						*outError = error;
					}
					return true;
				}
			}
		}
		if (outError != nullptr) {
			*outError = error;
		}
		if (outTrace != nullptr) {
			*outTrace = trace;
		}
		return false;
	}

	{
		std::lock_guard<std::mutex> lock(g_cacheState.mutex);
		g_cacheState.snapshot = snapshot;
		g_cacheState.lastRefreshTraceId = traceId;
		g_cacheState.lastRefreshAt = std::chrono::steady_clock::now();
	}

	outSnapshot = snapshot;
	if (outRefreshed != nullptr) {
		*outRefreshed = true;
	}
	if (outTrace != nullptr) {
		*outTrace = trace;
	}
	return true;
}

bool ProjectSourceCacheManager::ResolveSnapshotForHit(
	const HitToken& token,
	bool refreshIfStale,
	Snapshot& outSnapshot,
	bool* outRefreshed,
	std::string* outError,
	std::string* outTrace)
{
	outSnapshot = {};
	if (outRefreshed != nullptr) {
		*outRefreshed = false;
	}
	if (outError != nullptr) {
		outError->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	const std::string normalizedTokenSource = NormalizePathForProjectCache(token.sourcePath);
	if (!IsProjectSourcePath(normalizedTokenSource)) {
		if (outError != nullptr) {
			*outError = "project hit token source_path is invalid";
		}
		if (outTrace != nullptr) {
			*outTrace = "token_source_invalid";
		}
		return false;
	}

	{
		std::lock_guard<std::mutex> lock(g_cacheState.mutex);
		if (g_cacheState.snapshot.ok &&
			NormalizePathForProjectCache(g_cacheState.snapshot.sourcePath) == normalizedTokenSource &&
			g_cacheState.snapshot.revision == token.revision) {
			outSnapshot = g_cacheState.snapshot;
			if (outTrace != nullptr) {
				*outTrace = std::format(
					"reuse_token_revision revision={} pages={}",
					outSnapshot.revision,
					outSnapshot.pages.size());
			}
			return true;
		}
	}

	const std::string currentSourcePath = ResolveCurrentSourcePathForProjectCache();
	if (NormalizePathForProjectCache(currentSourcePath) != normalizedTokenSource) {
		if (outError != nullptr) {
			*outError = "current project source changed, please rerun search_project_source_cache";
		}
		if (outTrace != nullptr) {
			*outTrace = std::format(
				"current_source_changed current={} token={}",
				currentSourcePath,
				normalizedTokenSource);
		}
		return false;
	}

	return EnsureCurrentSourceLatest(
		outSnapshot,
		refreshIfStale,
		outRefreshed,
		outError,
		outTrace);
}

bool ProjectSourceCacheManager::GetSnapshotCopy(Snapshot& outSnapshot) const
{
	std::lock_guard<std::mutex> lock(g_cacheState.mutex);
	return TryCopyCurrentSnapshotLocked(outSnapshot);
}

std::string ProjectSourceCacheManager::BuildHitToken(const HitToken& token) const
{
	nlohmann::json root;
	root["source_path"] = LocalToUtf8TextForProjectCache(token.sourcePath);
	root["revision"] = token.revision;
	root["page_index"] = token.pageIndex;
	root["line_number"] = token.lineNumber;
	root["page_name"] = LocalToUtf8TextForProjectCache(token.pageName);
	root["page_type_key"] = token.pageTypeKey;
	root["page_type_name"] = LocalToUtf8TextForProjectCache(token.pageTypeName);
	return std::string(kProjectHitTokenPrefix) + root.dump();
}

bool ProjectSourceCacheManager::ParseHitToken(const std::string& text, HitToken& outToken, std::string* outError) const
{
	outToken = {};
	if (outError != nullptr) {
		outError->clear();
	}

	const std::string tokenText = TrimAsciiCopyForProjectCache(text);
	if (tokenText.rfind(kProjectHitTokenPrefix, 0) != 0) {
		if (outError != nullptr) {
			*outError = "not a project cache token";
		}
		return false;
	}

	try {
		const nlohmann::json root = nlohmann::json::parse(
			tokenText.substr(sizeof(kProjectHitTokenPrefix) - 1));
		if (root.contains("source_path") && root["source_path"].is_string()) {
			outToken.sourcePath = Utf8ToLocalTextForProjectCache(root["source_path"].get<std::string>());
		}
		if (root.contains("revision") && root["revision"].is_number_unsigned()) {
			outToken.revision = root["revision"].get<uint64_t>();
		}
		else if (root.contains("revision") && root["revision"].is_number_integer()) {
			outToken.revision = static_cast<uint64_t>((std::max)(0, root["revision"].get<int>()));
		}
		if (root.contains("page_index") && root["page_index"].is_number_integer()) {
			outToken.pageIndex = root["page_index"].get<int>();
		}
		if (root.contains("line_number") && root["line_number"].is_number_integer()) {
			outToken.lineNumber = root["line_number"].get<int>();
		}
		if (root.contains("page_name") && root["page_name"].is_string()) {
			outToken.pageName = Utf8ToLocalTextForProjectCache(root["page_name"].get<std::string>());
		}
		if (root.contains("page_type_key") && root["page_type_key"].is_string()) {
			outToken.pageTypeKey = root["page_type_key"].get<std::string>();
		}
		if (root.contains("page_type_name") && root["page_type_name"].is_string()) {
			outToken.pageTypeName = Utf8ToLocalTextForProjectCache(root["page_type_name"].get<std::string>());
		}
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) {
			*outError = std::string("parse project cache token failed: ") + ex.what();
		}
		return false;
	}

	if (outToken.sourcePath.empty() || outToken.pageIndex < 0 || outToken.lineNumber <= 0) {
		if (outError != nullptr) {
			*outError = "project cache token missing required fields";
		}
		return false;
	}

	outToken.sourcePath = NormalizePathForProjectCache(outToken.sourcePath);
	return true;
}

} // namespace project_source_cache
