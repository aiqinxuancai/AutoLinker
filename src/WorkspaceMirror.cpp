#include "WorkspaceMirror.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <format>
#include <mutex>
#include <random>
#include <string>
#include <unordered_map>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "EPackagerIntegration.h"
#include "Global.h"

namespace WorkspaceMirror {
namespace {

using json = nlohmann::json;

struct MirrorState {
	std::filesystem::path sourcePath;
	std::filesystem::path mirrorRoot;
	bool valid = false;
	std::unordered_map<std::string, ProgramItemRef> itemByRelativePath;
};

std::mutex g_mutex;
MirrorState g_state;

std::wstring WideFromCodePage(const std::string& text, UINT codePage, DWORD flags = 0)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int wideLen = MultiByteToWideChar(
		codePage,
		flags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}
	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
			codePage,
			flags,
			text.data(),
			static_cast<int>(text.size()),
			wide.data(),
			wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string StringFromWideCodePage(const std::wstring& text, UINT codePage)
{
	if (text.empty()) {
		return std::string();
	}
	const int outLen = WideCharToMultiByte(
		codePage,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (outLen <= 0) {
		return std::string();
	}
	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(
			codePage,
			0,
			text.data(),
			static_cast<int>(text.size()),
			out.data(),
			outLen,
			nullptr,
			nullptr) <= 0) {
		return std::string();
	}
	return out;
}

std::wstring WideFromUtf8(const std::string& text)
{
	return WideFromCodePage(text, CP_UTF8, MB_ERR_INVALID_CHARS);
}

std::string Utf8FromWide(const std::wstring& text)
{
	return StringFromWideCodePage(text, CP_UTF8);
}

std::string LocalFromUtf8(const std::string& text)
{
	constexpr UINT kEideCodePage = 936;
	const std::wstring wide = WideFromUtf8(text);
	return wide.empty() && !text.empty() ? text : StringFromWideCodePage(wide, kEideCodePage);
}

std::string Utf8FromPath(const std::filesystem::path& path)
{
	return Utf8FromWide(path.wstring());
}

std::string LocalFromPath(const std::filesystem::path& path)
{
	return LocalFromUtf8(Utf8FromPath(path));
}

std::filesystem::path PathFromUtf8(const std::string& text)
{
	const std::wstring wide = WideFromUtf8(text);
	return std::filesystem::path(wide.empty() && !text.empty() ? std::wstring(text.begin(), text.end()) : wide);
}

std::string ToLowerAscii(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string NormalizeRelativePathUtf8(std::string text)
{
	std::replace(text.begin(), text.end(), '\\', '/');
	while (!text.empty() && text.front() == '/') {
		text.erase(text.begin());
	}
	return text;
}

bool StartsWithAsciiInsensitive(const std::string& text, const std::string& prefix)
{
	if (text.size() < prefix.size()) {
		return false;
	}
	return ToLowerAscii(text.substr(0, prefix.size())) == ToLowerAscii(prefix);
}

bool IsSamePath(const std::filesystem::path& lhs, const std::filesystem::path& rhs)
{
	std::error_code ec;
	const auto left = std::filesystem::weakly_canonical(lhs, ec);
	if (ec) {
		return _wcsicmp(lhs.wstring().c_str(), rhs.wstring().c_str()) == 0;
	}
	const auto right = std::filesystem::weakly_canonical(rhs, ec);
	if (ec) {
		return _wcsicmp(lhs.wstring().c_str(), rhs.wstring().c_str()) == 0;
	}
	return _wcsicmp(left.wstring().c_str(), right.wstring().c_str()) == 0;
}

bool IsPathInside(const std::filesystem::path& root, const std::filesystem::path& child)
{
	std::error_code ec;
	const auto rootCanonical = std::filesystem::weakly_canonical(root, ec);
	if (ec || rootCanonical.empty()) {
		return false;
	}
	const auto childCanonical = std::filesystem::weakly_canonical(child, ec);
	if (ec || childCanonical.empty()) {
		return false;
	}
	std::wstring rootText = rootCanonical.wstring();
	const std::wstring childText = childCanonical.wstring();
	if (!rootText.empty() && rootText.back() != L'\\' && rootText.back() != L'/') {
		rootText.push_back(L'\\');
	}
	return _wcsnicmp(childText.c_str(), rootText.c_str(), rootText.size()) == 0 ||
		_wcsicmp(childText.c_str(), rootCanonical.wstring().c_str()) == 0;
}

bool ReadUtf8File(const std::filesystem::path& path, std::string& outText)
{
	outText.clear();
	std::ifstream file(path, std::ios::binary);
	if (!file) {
		return false;
	}
	outText.assign(std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>());
	if (outText.size() >= 3 &&
		static_cast<unsigned char>(outText[0]) == 0xEF &&
		static_cast<unsigned char>(outText[1]) == 0xBB &&
		static_cast<unsigned char>(outText[2]) == 0xBF) {
		outText.erase(0, 3);
	}
	return true;
}

bool ShouldRemoveMirrorRoot(const std::filesystem::path& mirrorRoot)
{
	if (mirrorRoot.empty()) {
		return false;
	}
	const std::wstring name = mirrorRoot.filename().wstring();
	if (name.rfind(L"al_", 0) != 0) {
		return false;
	}
	std::error_code ec;
	const auto canonical = std::filesystem::weakly_canonical(mirrorRoot, ec);
	if (ec || canonical.empty()) {
		return false;
	}
	const std::wstring pathText = canonical.wstring();
	return pathText.find(L"\\.temp\\") != std::wstring::npos ||
		pathText.find(L"\\AutoLinker\\workspace-mirror\\") != std::wstring::npos;
}

void RemoveMirrorRootIfSafe(const std::filesystem::path& mirrorRoot)
{
	if (!ShouldRemoveMirrorRoot(mirrorRoot)) {
		return;
	}
	std::error_code ec;
	std::filesystem::remove_all(mirrorRoot, ec);
}

bool RemoveMainProjectArtifacts(const std::filesystem::path& mirrorRoot, std::string& outError)
{
	outError.clear();
	if (!ShouldRemoveMirrorRoot(mirrorRoot)) {
		outError = "拒绝清理非 AutoLinker 工程镜像目录：" + LocalFromPath(mirrorRoot);
		return false;
	}

	for (const auto& child : {
			L"src",
			L"project",
			L"image",
			L"audio",
			L"header",
			L"info.json",
			L"AGENTS.md" }) {
		std::error_code ec;
		std::filesystem::remove_all(mirrorRoot / child, ec);
		if (ec) {
			outError = "清理工程镜像主工程文件失败：" + ec.message();
			return false;
		}
	}
	return true;
}

void CleanupSiblingMirrors(const std::filesystem::path& baseDir)
{
	std::error_code ec;
	if (!std::filesystem::exists(baseDir, ec) || !std::filesystem::is_directory(baseDir, ec)) {
		return;
	}
	for (const auto& entry : std::filesystem::directory_iterator(baseDir, ec)) {
		if (ec) {
			break;
		}
		if (!entry.is_directory(ec)) {
			continue;
		}
		const std::wstring name = entry.path().filename().wstring();
		if (name.rfind(L"al_", 0) == 0) {
			RemoveMirrorRootIfSafe(entry.path());
		}
	}
}

// 判断路径是否位于云同步盘（OneDrive / Dropbox / Google Drive / 坚果云 等）下。
// 在同步盘里解包成百上千个文件会被同步客户端实时扫描/加锁/上传，
// 导致 e-packager unpack 从 ~1s 暴涨到几十秒，因此需将镜像改放到本地临时目录。
// 注意：本工程以 MultiByte 字符集且未启用 /utf-8 编译，源文件中的中文宽字面量
// 会被按 GBK 解释而错乱，因此这里只用纯 ASCII 标记，并自行做 ASCII 小写转换。
bool IsUnderCloudSyncPath(const std::filesystem::path& path)
{
	std::wstring text = path.wstring();
	for (wchar_t& c : text) {
		if (c >= L'A' && c <= L'Z') {
			c = static_cast<wchar_t>(c - L'A' + L'a');
		}
		else if (c == L'/') {
			c = L'\\';
		}
	}

	// 路径分段里出现这些目录名即视为同步盘（用分隔符界定，避免误伤子串）。
	static const wchar_t* const kMarkers[] = {
		L"onedrive",
		L"dropbox",
		L"google drive",
		L"googledrive",
		L"nutstore",   // 坚果云英文目录名
	};
	for (const wchar_t* marker : kMarkers) {
		const std::wstring needle = std::wstring(L"\\") + marker;
		size_t pos = text.find(needle);
		while (pos != std::wstring::npos) {
			const size_t after = pos + needle.size();
			// 命中 "\onedrive\"、以 "\onedrive" 结尾、或紧跟空格/连字符（如 "OneDrive - 公司"）。
			if (after >= text.size() || text[after] == L'\\' || text[after] == L' ' || text[after] == L'-') {
				return true;
			}
			pos = text.find(needle, pos + 1);
		}
	}

	// 兜底：OneDrive 环境变量指向的根目录前缀。
	for (const wchar_t* var : { L"OneDrive", L"OneDriveConsumer", L"OneDriveCommercial" }) {
		wchar_t buf[MAX_PATH] = {};
		const DWORD n = ::GetEnvironmentVariableW(var, buf, MAX_PATH);
		if (n == 0 || n >= MAX_PATH) {
			continue;
		}
		std::wstring root(buf, n);
		for (wchar_t& c : root) {
			if (c >= L'A' && c <= L'Z') {
				c = static_cast<wchar_t>(c - L'A' + L'a');
			}
			else if (c == L'/') {
				c = L'\\';
			}
		}
		if (!root.empty() && text.rfind(root, 0) == 0) {
			return true;
		}
	}
	return false;
}

std::filesystem::path BuildUniqueMirrorRoot(const std::filesystem::path& sourcePath)
{
	const DWORD pid = GetCurrentProcessId();
	const ULONGLONG tick = GetTickCount64();
	std::random_device rd;
	const unsigned int salt = rd();
	const std::wstring uniqueName = std::format(L"al_{}_{}_{:08x}", pid, tick, salt);

	std::error_code ec;
	const std::filesystem::path localBase =
		std::filesystem::temp_directory_path(ec) / L"AutoLinker" / L"workspace-mirror";

	// 工程在云同步盘下时，优先用本地系统临时目录，避免同步客户端拖慢解包。
	const bool preferLocalTemp = IsUnderCloudSyncPath(sourcePath.parent_path());
	if (preferLocalTemp) {
		std::error_code localEc;
		std::filesystem::create_directories(localBase, localEc);
		if (!localEc) {
			CleanupSiblingMirrors(localBase);
			OutputStringToELog(
				"[WorkspaceMirror] source under cloud-sync path, using local temp mirror: " +
				LocalFromPath(localBase));
			return localBase / uniqueName;
		}
		OutputStringToELog(
			"[WorkspaceMirror] cloud-sync detected but local temp unavailable, falling back to project .temp");
	}

	std::filesystem::path preferredBase = sourcePath.parent_path() / L".temp";
	std::filesystem::create_directories(preferredBase, ec);
	if (!ec) {
		CleanupSiblingMirrors(preferredBase);
		return preferredBase / uniqueName;
	}

	std::filesystem::create_directories(localBase, ec);
	CleanupSiblingMirrors(localBase);
	return localBase / uniqueName;
}

bool BuildSnapshot(const std::filesystem::path& sourcePath, std::filesystem::path& outSnapshotPath, std::string& outTrace, std::string& outError)
{
	outSnapshotPath = EPackagerIntegration::BuildCurrentProjectSnapshotPathForSource(sourcePath);
	size_t bytesWritten = 0;
	std::string snapshotTrace;
	if (!EPackagerIntegration::WriteCurrentProjectSnapshot(outSnapshotPath, bytesWritten, snapshotTrace, outError)) {
		if (!snapshotTrace.empty()) {
			outTrace = outTrace.empty() ? snapshotTrace : (outTrace + "|" + snapshotTrace);
		}
		std::error_code ec;
		if (std::filesystem::exists(sourcePath, ec) && std::filesystem::is_regular_file(sourcePath, ec)) {
			outSnapshotPath = sourcePath;
			const std::string memoryError = outError;
			outError.clear();
			outTrace =
				(outTrace.empty() ? std::string() : (outTrace + "|")) +
				"snapshot_kind=saved_file_fallback|memory_snapshot_error=" +
				memoryError;
			OutputStringToELog("[WorkspaceMirror] memory snapshot unavailable, unpack saved file: " + LocalFromPath(sourcePath));
			return true;
		}
		EPackagerIntegration::CleanupSnapshotRoot(outSnapshotPath.parent_path());
		if (!outTrace.empty()) {
			outError += " trace=" + outTrace;
		}
		return false;
	}
	if (!snapshotTrace.empty()) {
		outTrace = outTrace.empty() ? snapshotTrace : (outTrace + "|" + snapshotTrace);
	}
	return true;
}

void AddIndexEntry(MirrorState& state, ProgramItemRef item)
{
	item.relativePathUtf8 = NormalizeRelativePathUtf8(item.relativePathUtf8);
	if (item.relativePathUtf8.empty()) {
		return;
	}
	state.itemByRelativePath.insert_or_assign(ToLowerAscii(item.relativePathUtf8), std::move(item));
}

void AddFixedTableEntries(MirrorState& state)
{
	AddIndexEntry(state, ProgramItemRef{ "src/.数据类型.txt", LocalFromUtf8("自定义数据类型"), "user_data_type", true, true, false });
	AddIndexEntry(state, ProgramItemRef{ "src/.DLL声明.txt", LocalFromUtf8("Dll命令"), "dll_command", true, true, false });
	AddIndexEntry(state, ProgramItemRef{ "src/.常量.txt", LocalFromUtf8("常量表..."), "const_resource", true, true, false });
	AddIndexEntry(state, ProgramItemRef{ "src/.全局变量.txt", LocalFromUtf8("全局变量"), "global_var", true, true, false });
}

void ParseFileMetaArray(MirrorState& state, const json& root, const char* key, bool formXml)
{
	const auto it = root.find(key);
	if (it == root.end() || !it->is_array()) {
		return;
	}
	for (const auto& row : *it) {
		if (!row.is_object()) {
			continue;
		}
		const std::string relativePath = row.value("relativePath", std::string());
		const std::string logicalName = row.value("logicalName", std::string());
		if (relativePath.empty() || logicalName.empty()) {
			continue;
		}
		ProgramItemRef item;
		item.relativePathUtf8 = relativePath;
		item.pageNameLocal = LocalFromUtf8(logicalName);
		item.kind = formXml ? "form" : std::string();
		item.editable = !formXml && StartsWithAsciiInsensitive(NormalizeRelativePathUtf8(relativePath), "src/");
		item.fixedTable = false;
		item.formXml = formXml;
		AddIndexEntry(state, std::move(item));
	}
}

bool ParseMetadata(MirrorState& state, std::string& outError)
{
	std::filesystem::path metaPath = state.mirrorRoot / L"project" / L"_meta.json";
	std::error_code ec;
	if (!std::filesystem::exists(metaPath, ec)) {
		metaPath = state.mirrorRoot / L"src" / L"_meta.json";
	}
	std::string metaText;
	if (!ReadUtf8File(metaPath, metaText)) {
		outError = "读取 e-packager 元数据失败：" + Utf8FromPath(metaPath);
		return false;
	}

	json root;
	try {
		root = json::parse(metaText);
	}
	catch (const std::exception& ex) {
		outError = std::string("解析 e-packager 元数据失败：") + ex.what();
		return false;
	}

	state.itemByRelativePath.clear();
	ParseFileMetaArray(state, root, "sourceFiles", false);
	ParseFileMetaArray(state, root, "formFiles", true);
	AddFixedTableEntries(state);
	return true;
}

bool RebuildMirrorLocked(const std::filesystem::path& sourcePath, std::string& outError)
{
	RemoveMirrorRootIfSafe(g_state.mirrorRoot);
	g_state = {};
	g_state.sourcePath = sourcePath;
	g_state.mirrorRoot = BuildUniqueMirrorRoot(sourcePath);

	std::filesystem::path snapshotPath;
	std::string snapshotTrace;
	if (!BuildSnapshot(sourcePath, snapshotPath, snapshotTrace, outError)) {
		g_state = {};
		return false;
	}
	const bool snapshotIsTemporary = !IsSamePath(snapshotPath, sourcePath);

	std::filesystem::path toolPath;
	if (!EPackagerIntegration::EnsureToolReady(toolPath, outError)) {
		if (snapshotIsTemporary) {
			EPackagerIntegration::CleanupSnapshotRoot(snapshotPath.parent_path());
		}
		g_state = {};
		return false;
	}

	std::error_code ec;
	std::filesystem::create_directories(g_state.mirrorRoot, ec);
	if (ec) {
		outError = "创建工程镜像目录失败：" + ec.message();
		if (snapshotIsTemporary) {
			EPackagerIntegration::CleanupSnapshotRoot(snapshotPath.parent_path());
		}
		g_state = {};
		return false;
	}

	OutputStringToELog("[WorkspaceMirror] preparing workspace mirror: " + LocalFromPath(g_state.mirrorRoot));
	const EPackagerIntegration::ProcessRunResult result = EPackagerIntegration::RunProcessAndCapture(
		toolPath,
		{ L"unpack", snapshotPath.wstring(), g_state.mirrorRoot.wstring() },
		toolPath.parent_path());
	if (snapshotIsTemporary) {
		EPackagerIntegration::CleanupSnapshotRoot(snapshotPath.parent_path());
	}

	if (!result.ok) {
		outError = std::format(
			"e-packager unpack failed, exitCode={} {}",
			result.exitCode,
			result.error);
		if (!result.stdErrBytes.empty()) {
			outError += " stderr=" + result.stdErrBytes;
		}
		RemoveMirrorRootIfSafe(g_state.mirrorRoot);
		g_state = {};
		return false;
	}

	if (!ParseMetadata(g_state, outError)) {
		RemoveMirrorRootIfSafe(g_state.mirrorRoot);
		g_state = {};
		return false;
	}

	g_state.valid = true;
	OutputStringToELog("[WorkspaceMirror] workspace mirror ready: " + LocalFromPath(g_state.mirrorRoot));
	return true;
}

bool RefreshMirrorMainOnlyLocked(const std::filesystem::path& sourcePath, std::string& outError)
{
	if (g_state.mirrorRoot.empty() || !IsSamePath(g_state.sourcePath, sourcePath)) {
		return false;
	}

	std::error_code ec;
	if (!std::filesystem::exists(g_state.mirrorRoot, ec) ||
		!std::filesystem::is_directory(g_state.mirrorRoot, ec)) {
		return false;
	}

	std::filesystem::path snapshotPath;
	std::string snapshotTrace;
	if (!BuildSnapshot(sourcePath, snapshotPath, snapshotTrace, outError)) {
		return false;
	}
	const bool snapshotIsTemporary = !IsSamePath(snapshotPath, sourcePath);

	std::filesystem::path toolPath;
	if (!EPackagerIntegration::EnsureToolReady(toolPath, outError)) {
		if (snapshotIsTemporary) {
			EPackagerIntegration::CleanupSnapshotRoot(snapshotPath.parent_path());
		}
		return false;
	}

	if (!RemoveMainProjectArtifacts(g_state.mirrorRoot, outError)) {
		if (snapshotIsTemporary) {
			EPackagerIntegration::CleanupSnapshotRoot(snapshotPath.parent_path());
		}
		return false;
	}

	auto runMainOnlyUnpack = [&]() {
		return EPackagerIntegration::RunProcessAndCapture(
			toolPath,
			{ L"unpack", snapshotPath.wstring(), g_state.mirrorRoot.wstring(), L"--main-only" },
			toolPath.parent_path());
	};

	OutputStringToELog("[WorkspaceMirror] refreshing main workspace mirror: " + LocalFromPath(g_state.mirrorRoot));
	EPackagerIntegration::ProcessRunResult result = runMainOnlyUnpack();
	if (!result.ok) {
		OutputStringToELog("[WorkspaceMirror] main-only refresh failed, checking latest e-packager");
		std::string updateError;
		if (EPackagerIntegration::EnsureLatestToolReady(toolPath, updateError)) {
			result = runMainOnlyUnpack();
		}
		else {
			OutputStringToELog("[WorkspaceMirror] latest e-packager check failed: " + updateError);
		}
	}
	if (snapshotIsTemporary) {
		EPackagerIntegration::CleanupSnapshotRoot(snapshotPath.parent_path());
	}

	if (!result.ok) {
		outError = std::format(
			"e-packager main-only unpack failed, exitCode={} {}",
			result.exitCode,
			result.error);
		if (!result.stdErrBytes.empty()) {
			outError += " stderr=" + result.stdErrBytes;
		}
		return false;
	}

	if (!ParseMetadata(g_state, outError)) {
		g_state.valid = false;
		return false;
	}

	g_state.valid = true;
	OutputStringToELog("[WorkspaceMirror] main workspace mirror refreshed: " + LocalFromPath(g_state.mirrorRoot));
	return true;
}

bool EnsureMirrorFreshLocked(std::string& outError)
{
	outError.clear();
	std::filesystem::path sourcePath;
	if (!EPackagerIntegration::GetCurrentSourcePath(sourcePath, outError)) {
		return false;
	}
	if (g_state.valid && !g_state.mirrorRoot.empty() && IsSamePath(g_state.sourcePath, sourcePath)) {
		std::error_code ec;
		if (std::filesystem::exists(g_state.mirrorRoot, ec) &&
			std::filesystem::is_directory(g_state.mirrorRoot, ec) &&
			!g_state.itemByRelativePath.empty()) {
			return true;
		}
		g_state.valid = false;
	}
	return RebuildMirrorLocked(sourcePath, outError);
}

bool BuildSafeRelativePath(const std::string& filePathUtf8, std::filesystem::path& outRelative, std::string& outRelativeUtf8, std::string& outError)
{
	outRelative.clear();
	outRelativeUtf8.clear();
	outError.clear();

	const std::string normalized = NormalizeRelativePathUtf8(filePathUtf8);
	if (normalized.empty()) {
		outError = "file_path is required";
		return false;
	}
	if (normalized.find('\0') != std::string::npos || normalized.find("..") != std::string::npos) {
		outError = "file_path must be a relative path inside the workspace mirror";
		return false;
	}

	const std::filesystem::path candidate = PathFromUtf8(normalized);
	if (candidate.is_absolute()) {
		outError = "absolute file_path is not allowed";
		return false;
	}
	for (const auto& part : candidate) {
		const std::wstring value = part.wstring();
		if (value == L"." || value == L"..") {
			outError = "file_path must not contain . or .. segments";
			return false;
		}
	}
	outRelative = candidate;
	outRelativeUtf8 = normalized;
	return true;
}

} // namespace

bool EnsureMirrorFresh(std::string& outError)
{
	std::lock_guard<std::mutex> guard(g_mutex);
	return EnsureMirrorFreshLocked(outError);
}

bool RefreshMirror(std::string& outError, std::string* outMode)
{
	std::lock_guard<std::mutex> guard(g_mutex);
	outError.clear();
	if (outMode != nullptr) {
		outMode->clear();
	}

	std::filesystem::path sourcePath;
	if (!EPackagerIntegration::GetCurrentSourcePath(sourcePath, outError)) {
		return false;
	}

	if (RefreshMirrorMainOnlyLocked(sourcePath, outError)) {
		if (outMode != nullptr) {
			*outMode = "main_only";
		}
		return true;
	}

	if (!outError.empty()) {
		OutputStringToELog("[WorkspaceMirror] main-only refresh failed, rebuilding mirror: " + outError);
		outError.clear();
	}
	if (!RebuildMirrorLocked(sourcePath, outError)) {
		return false;
	}
	if (outMode != nullptr) {
		*outMode = "full";
	}
	return true;
}

void InvalidateMirror()
{
	std::lock_guard<std::mutex> guard(g_mutex);
	g_state.valid = false;
}

void ResetAndCleanup()
{
	std::lock_guard<std::mutex> guard(g_mutex);
	RemoveMirrorRootIfSafe(g_state.mirrorRoot);
	g_state = {};
}

bool GetMirrorRoot(std::filesystem::path& outRoot, std::string& outError)
{
	std::lock_guard<std::mutex> guard(g_mutex);
	if (!EnsureMirrorFreshLocked(outError)) {
		outRoot.clear();
		return false;
	}
	outRoot = g_state.mirrorRoot;
	return true;
}

bool ResolveFilePath(
	const std::string& filePathUtf8,
	std::filesystem::path& outFullPath,
	std::string& outRelativePathUtf8,
	std::string& outError)
{
	std::lock_guard<std::mutex> guard(g_mutex);
	if (!EnsureMirrorFreshLocked(outError)) {
		return false;
	}

	std::filesystem::path relativePath;
	if (!BuildSafeRelativePath(filePathUtf8, relativePath, outRelativePathUtf8, outError)) {
		return false;
	}

	const std::filesystem::path fullPath = g_state.mirrorRoot / relativePath;
	if (!IsPathInside(g_state.mirrorRoot, fullPath)) {
		outError = "file_path resolves outside workspace mirror";
		return false;
	}
	std::error_code ec;
	if (!std::filesystem::exists(fullPath, ec) || !std::filesystem::is_regular_file(fullPath, ec)) {
		outError = "file not found in workspace mirror: " + outRelativePathUtf8;
		return false;
	}
	outFullPath = fullPath;
	return true;
}

bool ResolveFileToProgramItem(
	const std::string& filePathUtf8,
	ProgramItemRef& outItem,
	std::string& outError)
{
	outItem = {};
	std::lock_guard<std::mutex> guard(g_mutex);
	if (!EnsureMirrorFreshLocked(outError)) {
		return false;
	}

	std::filesystem::path relativePath;
	std::string relativePathUtf8;
	if (!BuildSafeRelativePath(filePathUtf8, relativePath, relativePathUtf8, outError)) {
		return false;
	}

	const auto it = g_state.itemByRelativePath.find(ToLowerAscii(relativePathUtf8));
	if (it == g_state.itemByRelativePath.end()) {
		outError = "file_path is not mapped to an editable current-project source page: " + relativePathUtf8;
		return false;
	}
	if (!it->second.editable) {
		outError = it->second.formXml
			? "form .xml files are read-only; edit the matching src/*.txt code page instead"
			: "this file is read-only for source editing";
		return false;
	}
	outItem = it->second;
	return true;
}

bool ListMirrorFiles(std::vector<std::string>& outRelativePathsUtf8, std::string& outError)
{
	outRelativePathsUtf8.clear();
	std::lock_guard<std::mutex> guard(g_mutex);
	if (!EnsureMirrorFreshLocked(outError)) {
		return false;
	}

	std::error_code ec;
	for (const auto& entry : std::filesystem::recursive_directory_iterator(g_state.mirrorRoot, ec)) {
		if (ec) {
			break;
		}
		if (!entry.is_regular_file(ec)) {
			continue;
		}
		const std::filesystem::path relative = std::filesystem::relative(entry.path(), g_state.mirrorRoot, ec);
		if (ec || relative.empty()) {
			continue;
		}
		outRelativePathsUtf8.push_back(NormalizeRelativePathUtf8(Utf8FromPath(relative)));
	}
	std::sort(outRelativePathsUtf8.begin(), outRelativePathsUtf8.end());
	return true;
}

} // namespace WorkspaceMirror
