#include "WorkspaceMirror.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <cstdlib>
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
	std::vector<std::string> relativePathsUtf8;
};

std::mutex g_mutex;
MirrorState g_state;
std::uint64_t g_generation = 1;

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

std::string Utf8FromLocal(const std::string& text)
{
	constexpr UINT kEideCodePage = 936;
	const std::wstring wide = WideFromCodePage(text, kEideCodePage);
	return wide.empty() && !text.empty() ? text : StringFromWideCodePage(wide, CP_UTF8);
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

bool WriteUtf8BomFile(const std::filesystem::path& path, const std::string& textUtf8, std::string& outError)
{
	std::filesystem::path tempPath = path;
	tempPath += std::format(
		L".autolinker-{}-{}-{}.tmp",
		GetCurrentProcessId(),
		GetCurrentThreadId(),
		GetTickCount64());
	std::ofstream file(tempPath, std::ios::binary | std::ios::trunc);
	if (!file) {
		outError = "open mirror file for write failed: " + Utf8FromPath(path);
		return false;
	}
	const unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
	file.write(reinterpret_cast<const char*>(bom), sizeof(bom));
	file.write(textUtf8.data(), static_cast<std::streamsize>(textUtf8.size()));
	if (!file.good()) {
		file.close();
		std::error_code cleanupError;
		std::filesystem::remove(tempPath, cleanupError);
		outError = "write mirror file failed: " + Utf8FromPath(path);
		return false;
	}
	file.close();
	if (MoveFileExW(
			tempPath.c_str(),
			path.c_str(),
			MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == FALSE) {
		const DWORD moveError = GetLastError();
		std::error_code cleanupError;
		std::filesystem::remove(tempPath, cleanupError);
		outError = "replace mirror file failed: " + std::to_string(moveError);
		return false;
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

bool TryParseMirrorOwnerProcessId(const std::wstring& name, DWORD& outProcessId)
{
	outProcessId = 0;
	if (name.rfind(L"al_", 0) != 0) {
		return false;
	}
	const size_t end = name.find(L'_', 3);
	if (end == std::wstring::npos || end == 3) {
		return false;
	}
	const std::wstring processIdText = name.substr(3, end - 3);
	wchar_t* parseEnd = nullptr;
	const unsigned long value = std::wcstoul(processIdText.c_str(), &parseEnd, 10);
	if (parseEnd == nullptr || *parseEnd != L'\0' || value == 0) {
		return false;
	}
	outProcessId = static_cast<DWORD>(value);
	return true;
}

bool IsProcessDefinitelyStopped(DWORD processId)
{
	if (processId == 0) {
		return false;
	}
	HANDLE process = OpenProcess(SYNCHRONIZE, FALSE, processId);
	if (process == nullptr) {
		// 无权限查询时采取保守策略，避免误删另一个活动 IDE 的镜像。
		return false;
	}
	const DWORD waitResult = WaitForSingleObject(process, 0);
	CloseHandle(process);
	return waitResult == WAIT_OBJECT_0;
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
		DWORD ownerProcessId = 0;
		if (!TryParseMirrorOwnerProcessId(name, ownerProcessId)) {
			continue;
		}
		if (ownerProcessId != GetCurrentProcessId() && !IsProcessDefinitelyStopped(ownerProcessId)) {
			continue;
		}
		RemoveMirrorRootIfSafe(entry.path());
	}
}

bool TryGetSystemMirrorBase(std::filesystem::path& outBase, std::string& outError)
{
	outBase.clear();
	outError.clear();
	std::error_code ec;
	const std::filesystem::path systemTemp = std::filesystem::temp_directory_path(ec);
	if (ec || systemTemp.empty()) {
		outError = "获取系统临时目录失败";
		if (ec) {
			outError += "：" + ec.message();
		}
		return false;
	}
	outBase = systemTemp / L"AutoLinker" / L"workspace-mirror";
	return true;
}

bool BuildUniqueMirrorRoot(std::filesystem::path& outRoot, std::string& outError)
{
	outRoot.clear();
	std::filesystem::path baseDir;
	if (!TryGetSystemMirrorBase(baseDir, outError)) {
		return false;
	}

	std::error_code ec;
	std::filesystem::create_directories(baseDir, ec);
	if (ec) {
		outError = "创建系统临时镜像目录失败：" + ec.message();
		return false;
	}
	CleanupSiblingMirrors(baseDir);

	const DWORD pid = GetCurrentProcessId();
	const ULONGLONG tick = GetTickCount64();
	std::random_device rd;
	const unsigned int salt = rd();
	const std::wstring uniqueName = std::format(L"al_{}_{}_{:08x}", pid, tick, salt);
	outRoot = baseDir / uniqueName;
	return true;
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

bool RebuildFileIndexLocked(MirrorState& state, std::string& outError)
{
	state.relativePathsUtf8.clear();
	std::error_code ec;
	const auto options = std::filesystem::directory_options::skip_permission_denied;
	for (std::filesystem::recursive_directory_iterator it(state.mirrorRoot, options, ec), end;
		it != end;
		it.increment(ec)) {
		if (ec) {
			outError = "enumerate workspace mirror failed: " + ec.message();
			state.relativePathsUtf8.clear();
			return false;
		}
		if (!it->is_regular_file(ec)) {
			if (ec) {
				ec.clear();
			}
			continue;
		}
		const std::filesystem::path relative = std::filesystem::relative(it->path(), state.mirrorRoot, ec);
		if (ec || relative.empty()) {
			outError = "resolve workspace mirror relative path failed: " + ec.message();
			state.relativePathsUtf8.clear();
			return false;
		}
		state.relativePathsUtf8.push_back(NormalizeRelativePathUtf8(Utf8FromPath(relative)));
	}
	std::sort(state.relativePathsUtf8.begin(), state.relativePathsUtf8.end());
	return true;
}

bool RebuildMirrorLocked(const std::filesystem::path& sourcePath, std::string& outError)
{
	RemoveMirrorRootIfSafe(g_state.mirrorRoot);
	g_state = {};
	if (!BuildUniqueMirrorRoot(g_state.mirrorRoot, outError)) {
		return false;
	}
	g_state.sourcePath = sourcePath;

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
	if (!RebuildFileIndexLocked(g_state, outError)) {
		RemoveMirrorRootIfSafe(g_state.mirrorRoot);
		g_state = {};
		return false;
	}

	g_state.valid = true;
	++g_generation;
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
	if (!RebuildFileIndexLocked(g_state, outError)) {
		g_state.valid = false;
		return false;
	}

	g_state.valid = true;
	++g_generation;
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
			!g_state.itemByRelativePath.empty() &&
			!g_state.relativePathsUtf8.empty()) {
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
	if (normalized.find('\0') != std::string::npos) {
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

bool PrepareFileAccess(std::string& outError)
{
	return EnsureMirrorFresh(outError);
}

bool GetPreparedFileAccessSnapshot(FileAccessSnapshot& outSnapshot, std::string& outError)
{
	outSnapshot = {};
	outError.clear();
	std::lock_guard<std::mutex> guard(g_mutex);
	if (!g_state.valid || g_state.mirrorRoot.empty() || g_state.relativePathsUtf8.empty()) {
		outError = "workspace mirror is not prepared";
		return false;
	}
	std::error_code ec;
	if (!std::filesystem::exists(g_state.mirrorRoot, ec) ||
		!std::filesystem::is_directory(g_state.mirrorRoot, ec)) {
		outError = "prepared workspace mirror directory is unavailable";
		return false;
	}
	outSnapshot.mirrorRoot = g_state.mirrorRoot;
	outSnapshot.relativePathsUtf8 = g_state.relativePathsUtf8;
	outSnapshot.generation = g_generation;
	return true;
}

bool ResolvePreparedFilePath(
	const FileAccessSnapshot& snapshot,
	const std::string& filePathUtf8,
	std::filesystem::path& outFullPath,
	std::string& outRelativePathUtf8,
	std::string& outError)
{
	outFullPath.clear();
	outRelativePathUtf8.clear();
	if (snapshot.mirrorRoot.empty()) {
		outError = "workspace mirror snapshot is empty";
		return false;
	}
	std::filesystem::path relativePath;
	if (!BuildSafeRelativePath(filePathUtf8, relativePath, outRelativePathUtf8, outError)) {
		return false;
	}
	const std::filesystem::path fullPath = snapshot.mirrorRoot / relativePath;
	if (!IsPathInside(snapshot.mirrorRoot, fullPath)) {
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

bool RefreshMirror(std::string& outError, std::string* outMode, RefreshMode mode)
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

	if (mode == RefreshMode::Full) {
		if (!RebuildMirrorLocked(sourcePath, outError)) {
			return false;
		}
		if (outMode != nullptr) {
			*outMode = "full";
		}
		return true;
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
	else if (mode == RefreshMode::MainOnly) {
		OutputStringToELog("[WorkspaceMirror] main-only refresh unavailable, rebuilding mirror");
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
	++g_generation;
}

bool UpdateMirrorTextFile(
	const std::string& filePathUtf8,
	const std::string& textLocal,
	std::string& outError)
{
	std::lock_guard<std::mutex> guard(g_mutex);
	outError.clear();
	if (!EnsureMirrorFreshLocked(outError)) {
		return false;
	}

	std::filesystem::path relativePath;
	std::string relativePathUtf8;
	if (!BuildSafeRelativePath(filePathUtf8, relativePath, relativePathUtf8, outError)) {
		return false;
	}

	const std::filesystem::path fullPath = g_state.mirrorRoot / relativePath;
	if (!IsPathInside(g_state.mirrorRoot, fullPath)) {
		outError = "file_path resolves outside workspace mirror";
		return false;
	}

	std::error_code ec;
	if (!std::filesystem::exists(fullPath, ec) || !std::filesystem::is_regular_file(fullPath, ec)) {
		outError = "file not found in workspace mirror: " + relativePathUtf8;
		return false;
	}

	if (!WriteUtf8BomFile(fullPath, Utf8FromLocal(textLocal), outError)) {
		return false;
	}

	g_state.valid = true;
	++g_generation;
	OutputStringToELog("[WorkspaceMirror] mirror text file updated: " + relativePathUtf8);
	return true;
}

void ResetAndCleanup()
{
	std::lock_guard<std::mutex> guard(g_mutex);
	RemoveMirrorRootIfSafe(g_state.mirrorRoot);
	g_state = {};
	++g_generation;
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

std::uint64_t GetGeneration()
{
	std::lock_guard<std::mutex> guard(g_mutex);
	return g_generation;
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

	outRelativePathsUtf8 = g_state.relativePathsUtf8;
	return true;
}

std::string BuildSelfTestReportJson()
{
	DWORD processId = 0;
	const bool ownerParsed = TryParseMirrorOwnerProcessId(L"al_1234_5678_deadbeef", processId) && processId == 1234;
	DWORD invalidProcessId = 0;
	const bool invalidOwnerRejected =
		!TryParseMirrorOwnerProcessId(L"al_bad_5678_deadbeef", invalidProcessId) &&
		!TryParseMirrorOwnerProcessId(L"other_1234_5678", invalidProcessId);

	std::filesystem::path relativePath;
	std::string normalizedPath;
	std::string error;
	const bool dottedNameAccepted = BuildSafeRelativePath(
		"src/name..with-dots.txt",
		relativePath,
		normalizedPath,
		error);
	const bool traversalRejected = !BuildSafeRelativePath(
		"src/../secret.txt",
		relativePath,
		normalizedPath,
		error);
	std::filesystem::path systemMirrorBase;
	std::string systemMirrorError;
	const bool systemMirrorBaseResolved = TryGetSystemMirrorBase(systemMirrorBase, systemMirrorError);
	std::error_code tempEc;
	const std::filesystem::path systemTemp = std::filesystem::temp_directory_path(tempEc);
	const bool systemTempOnly =
		systemMirrorBaseResolved &&
		!tempEc &&
		IsPathInside(systemTemp, systemMirrorBase) &&
		systemMirrorBase.filename() == L"workspace-mirror" &&
		systemMirrorBase.parent_path().filename() == L"AutoLinker";

	return json({
		{"name", "workspace-mirror-safety"},
		{"ok", ownerParsed && invalidOwnerRejected && dottedNameAccepted && traversalRejected && systemTempOnly},
		{"owner_pid_parsed", ownerParsed},
		{"invalid_owner_rejected", invalidOwnerRejected},
		{"dotted_name_accepted", dottedNameAccepted},
		{"traversal_rejected", traversalRejected},
		{"system_temp_only", systemTempOnly},
		{"mirror_base", Utf8FromPath(systemMirrorBase)}
	}).dump();
}

} // namespace WorkspaceMirror
