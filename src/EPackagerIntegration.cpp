#include "EPackagerIntegration.h"

#include <Windows.h>
#include <ShlObj.h>
#include <process.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <memory>
#include <sstream>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AutoLinkerInternal.h"
#include "EideProjectBinarySerializer.h"
#include "Global.h"
#include "PathHelper.h"
#include "PowerShellToolRunner.h"
#include "WinINetUtil.h"

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")

namespace EPackagerIntegration {
namespace {

using json = nlohmann::json;

constexpr const char* kLatestReleaseApi = "https://api.github.com/repos/aiqinxuancai/e-packager/releases/latest";
constexpr const char* kGitHubHeaders =
	"User-Agent: AutoLinker\r\n"
	"Accept: application/vnd.github+json\r\n";
constexpr long long kUpdateCheckIntervalSeconds = 7LL * 24LL * 60LL * 60LL;

std::atomic_bool g_unpackTaskRunning = false;

struct ScopedComInit {
	HRESULT hr = E_FAIL;
	bool initialized = false;

	ScopedComInit()
	{
		hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
		initialized = SUCCEEDED(hr);
		if (hr == RPC_E_CHANGED_MODE) {
			initialized = false;
		}
	}

	~ScopedComInit()
	{
		if (initialized) {
			CoUninitialize();
		}
	}
};

struct LatestReleaseInfo {
	std::string tag;
	std::string assetName;
	std::string downloadUrl;
};

struct UnpackRequest {
	std::filesystem::path originalSourcePath;
	std::filesystem::path snapshotPath;
	std::filesystem::path snapshotRoot;
	std::filesystem::path unpackDir;
};

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

std::wstring WideFromLocal(const std::string& text)
{
	return WideFromCodePage(text, CP_ACP);
}

std::wstring WideFromUtf8(const std::string& text)
{
	return WideFromCodePage(text, CP_UTF8, MB_ERR_INVALID_CHARS);
}

std::string LocalFromWide(const std::wstring& text)
{
	return StringFromWideCodePage(text, CP_ACP);
}

std::string Utf8FromWide(const std::wstring& text)
{
	return StringFromWideCodePage(text, CP_UTF8);
}

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string TrimAsciiCopy(std::string text)
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

bool EndsWithInsensitive(const std::string& text, const std::string& suffix)
{
	if (text.size() < suffix.size()) {
		return false;
	}
	return ToLowerAscii(text.substr(text.size() - suffix.size())) == ToLowerAscii(suffix);
}

std::string StripLeadingVersionPrefix(std::string version)
{
	if (!version.empty() && (version.front() == 'v' || version.front() == 'V')) {
		version.erase(version.begin());
	}
	return version;
}

std::filesystem::path PathFromLocal(const std::string& path)
{
	return std::filesystem::path(WideFromLocal(path));
}

std::string LocalPathString(const std::filesystem::path& path)
{
	return LocalFromWide(path.wstring());
}

std::string Utf8PathString(const std::filesystem::path& path)
{
	return Utf8FromWide(path.wstring());
}

std::filesystem::path GetTempDirectory()
{
	wchar_t buffer[MAX_PATH] = {};
	const DWORD size = GetTempPathW(static_cast<DWORD>(_countof(buffer)), buffer);
	if (size > 0 && size < _countof(buffer)) {
		return std::filesystem::path(buffer);
	}
	return std::filesystem::temp_directory_path();
}

std::filesystem::path BuildCurrentProjectSnapshotRoot()
{
	const ULONGLONG tick = GetTickCount64();
	const DWORD pid = GetCurrentProcessId();
	return GetTempDirectory() / L"AutoLinker" / L"unpack-snapshots" / std::format(L"{}.{}", pid, tick);
}

std::filesystem::path BuildCurrentProjectSnapshotPath(const std::filesystem::path& sourcePath)
{
	std::filesystem::path fileName = sourcePath.filename();
	if (fileName.empty()) {
		fileName = L"current_project.e";
	}
	return BuildCurrentProjectSnapshotRoot() / fileName;
}

bool WriteCurrentProjectSnapshotImpl(
	const std::filesystem::path& snapshotPath,
	size_t& outBytesWritten,
	std::string& outTrace,
	std::string& outError)
{
	outBytesWritten = 0;
	outTrace.clear();
	outError.clear();

	const std::string localPath = LocalPathString(snapshotPath);
	if (TrimAsciiCopy(localPath).empty()) {
		outError = "snapshot path is empty";
		return false;
	}

	return e571::ProjectBinarySerializer::Instance().WriteCurrentProjectToFile(
		localPath,
		&outBytesWritten,
		&outError,
		&outTrace);
}

bool ShouldRemoveSnapshotRoot(const std::filesystem::path& snapshotRoot)
{
	if (snapshotRoot.empty()) {
		return false;
	}

	std::error_code ec;
	const std::filesystem::path root = std::filesystem::weakly_canonical(snapshotRoot, ec);
	if (ec || root.empty()) {
		return false;
	}

	const std::filesystem::path allowedRoot = std::filesystem::weakly_canonical(
		GetTempDirectory() / L"AutoLinker" / L"unpack-snapshots",
		ec);
	if (ec || allowedRoot.empty()) {
		return false;
	}

	const std::wstring rootText = root.wstring();
	std::wstring allowedText = allowedRoot.wstring();
	if (!allowedText.empty() && allowedText.back() != L'\\' && allowedText.back() != L'/') {
		allowedText.push_back(L'\\');
	}
	return _wcsnicmp(rootText.c_str(), allowedText.c_str(), allowedText.size()) == 0;
}

void CleanupSnapshotRootImpl(const std::filesystem::path& snapshotRoot)
{
	if (!ShouldRemoveSnapshotRoot(snapshotRoot)) {
		return;
	}
	std::error_code ec;
	std::filesystem::remove_all(snapshotRoot, ec);
}

std::filesystem::path GetToolsDirectory()
{
	return std::filesystem::path(WideFromLocal(GetBasePath())) / L"tools";
}

std::filesystem::path GetEPackagerExePath()
{
	return GetToolsDirectory() / L"e-packager.exe";
}

std::filesystem::path GetEPackagerMetaPath()
{
	return GetToolsDirectory() / L"e-packager.autolinker.json";
}

long long NowUnixSeconds()
{
	return std::chrono::duration_cast<std::chrono::seconds>(
		std::chrono::system_clock::now().time_since_epoch()).count();
}

std::wstring QuoteCommandLineArg(const std::wstring& arg)
{
	if (arg.empty()) {
		return L"\"\"";
	}

	bool needsQuote = false;
	for (wchar_t ch : arg) {
		if (ch == L' ' || ch == L'\t' || ch == L'\n' || ch == L'\v' || ch == L'"') {
			needsQuote = true;
			break;
		}
	}
	if (!needsQuote) {
		return arg;
	}

	std::wstring quoted = L"\"";
	size_t backslashes = 0;
	for (wchar_t ch : arg) {
		if (ch == L'\\') {
			++backslashes;
			continue;
		}
		if (ch == L'"') {
			quoted.append(backslashes * 2 + 1, L'\\');
			quoted.push_back(ch);
			backslashes = 0;
			continue;
		}
		quoted.append(backslashes, L'\\');
		backslashes = 0;
		quoted.push_back(ch);
	}
	quoted.append(backslashes * 2, L'\\');
	quoted.push_back(L'"');
	return quoted;
}

void ReadPipeToBytes(HANDLE pipe, std::string* output)
{
	if (pipe == nullptr || output == nullptr) {
		return;
	}

	char buffer[4096] = {};
	DWORD bytesRead = 0;
	while (ReadFile(pipe, buffer, static_cast<DWORD>(sizeof(buffer)), &bytesRead, nullptr) != FALSE && bytesRead > 0) {
		output->append(buffer, bytesRead);
	}
}

ProcessRunResult RunProcessAndCaptureImpl(
	const std::filesystem::path& exePath,
	const std::vector<std::wstring>& args,
	const std::filesystem::path& workingDirectory)
{
	ProcessRunResult result = {};

	std::wstring commandLine = QuoteCommandLineArg(exePath.wstring());
	for (const auto& arg : args) {
		commandLine.push_back(L' ');
		commandLine += QuoteCommandLineArg(arg);
	}

	SECURITY_ATTRIBUTES sa = {};
	sa.nLength = sizeof(sa);
	sa.bInheritHandle = TRUE;

	HANDLE stdOutRead = nullptr;
	HANDLE stdOutWrite = nullptr;
	HANDLE stdErrRead = nullptr;
	HANDLE stdErrWrite = nullptr;
	if (CreatePipe(&stdOutRead, &stdOutWrite, &sa, 0) == FALSE) {
		result.error = "CreatePipe stdout failed";
		return result;
	}
	if (CreatePipe(&stdErrRead, &stdErrWrite, &sa, 0) == FALSE) {
		CloseHandle(stdOutRead);
		CloseHandle(stdOutWrite);
		result.error = "CreatePipe stderr failed";
		return result;
	}

	SetHandleInformation(stdOutRead, HANDLE_FLAG_INHERIT, 0);
	SetHandleInformation(stdErrRead, HANDLE_FLAG_INHERIT, 0);

	STARTUPINFOW si = {};
	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE;
	si.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
	si.hStdOutput = stdOutWrite;
	si.hStdError = stdErrWrite;

	PROCESS_INFORMATION pi = {};
	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');
	const std::wstring cwd = workingDirectory.wstring();
	const BOOL created = CreateProcessW(
		nullptr,
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		TRUE,
		CREATE_NO_WINDOW,
		nullptr,
		cwd.empty() ? nullptr : cwd.c_str(),
		&si,
		&pi);

	CloseHandle(stdOutWrite);
	CloseHandle(stdErrWrite);

	if (created == FALSE) {
		CloseHandle(stdOutRead);
		CloseHandle(stdErrRead);
		result.error = std::format("CreateProcessW failed, error={}", GetLastError());
		return result;
	}

	std::thread stdoutReader(ReadPipeToBytes, stdOutRead, &result.stdOutBytes);
	std::thread stderrReader(ReadPipeToBytes, stdErrRead, &result.stdErrBytes);

	WaitForSingleObject(pi.hProcess, INFINITE);
	DWORD exitCode = 0;
	if (GetExitCodeProcess(pi.hProcess, &exitCode) != FALSE) {
		result.exitCode = exitCode;
	}

	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);

	if (stdoutReader.joinable()) {
		stdoutReader.join();
	}
	if (stderrReader.joinable()) {
		stderrReader.join();
	}
	CloseHandle(stdOutRead);
	CloseHandle(stdErrRead);

	result.ok = result.exitCode == 0 && result.error.empty();
	return result;
}

std::string BytesToLocalText(const std::string& bytes)
{
	if (bytes.empty()) {
		return std::string();
	}

	std::wstring wide = WideFromUtf8(bytes);
	if (!wide.empty()) {
		return LocalFromWide(wide);
	}
	wide = WideFromCodePage(bytes, CP_ACP);
	if (!wide.empty()) {
		return LocalFromWide(wide);
	}
	return bytes;
}

void OutputTextBlock(const std::string& title, const std::string& text)
{
	const std::string localText = BytesToLocalText(text);
	if (localText.empty()) {
		return;
	}

	OutputStringToELog(title);
	std::istringstream stream(localText);
	std::string line;
	while (std::getline(stream, line)) {
		if (!line.empty() && line.back() == '\r') {
			line.pop_back();
		}
		if (!line.empty()) {
			OutputStringToELog("  " + line);
		}
	}
}

std::string ReadFileText(const std::filesystem::path& path)
{
	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		return std::string();
	}
	std::ostringstream ss;
	ss << in.rdbuf();
	return ss.str();
}

void WriteFileText(const std::filesystem::path& path, const std::string& text)
{
	std::ofstream out(path, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		return;
	}
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
}

json LoadMeta()
{
	const auto text = ReadFileText(GetEPackagerMetaPath());
	if (text.empty()) {
		return json::object();
	}
	json value = json::parse(text, nullptr, false);
	return value.is_object() ? value : json::object();
}

void SaveMeta(const LatestReleaseInfo& info)
{
	json value = json::object();
	value["last_check_unix"] = NowUnixSeconds();
	value["tag"] = info.tag;
	value["asset_name"] = info.assetName;
	WriteFileText(GetEPackagerMetaPath(), value.dump(2));
}

bool IsUpdateCheckDue(bool toolExists)
{
	if (!toolExists) {
		return true;
	}
	const json meta = LoadMeta();
	const long long lastCheck = meta.value("last_check_unix", 0LL);
	return lastCheck <= 0 || NowUnixSeconds() - lastCheck >= kUpdateCheckIntervalSeconds;
}

bool FetchLatestRelease(LatestReleaseInfo& outInfo, std::string& outError)
{
	OutputStringToELog("[e-packager] 正在检查最新版本...");
	auto response = PerformGetRequest(kLatestReleaseApi, kGitHubHeaders, 60000, false, false);
	if (response.second != 200) {
		outError = std::format("GitHub API HTTP {}", response.second);
		if (!response.first.empty()) {
			outError += ": " + response.first.substr(0, (std::min<size_t>)(response.first.size(), 300));
		}
		return false;
	}

	json release = json::parse(response.first, nullptr, false);
	if (!release.is_object()) {
		outError = "GitHub API 返回的 release JSON 无效";
		return false;
	}

	LatestReleaseInfo info = {};
	info.tag = release.value("tag_name", std::string());
	const json assets = release.value("assets", json::array());
	for (const auto& asset : assets) {
		if (!asset.is_object()) {
			continue;
		}
		const std::string name = asset.value("name", std::string());
		const std::string lowered = ToLowerAscii(name);
		if (lowered.find("windows-win32") == std::string::npos || !EndsWithInsensitive(name, ".zip")) {
			continue;
		}
		info.assetName = name;
		info.downloadUrl = asset.value("browser_download_url", std::string());
		break;
	}

	if (info.tag.empty() || info.assetName.empty() || info.downloadUrl.empty()) {
		outError = "未找到 e-packager Windows Win32 zip 资源";
		return false;
	}

	outInfo = std::move(info);
	OutputStringToELog(std::format("[e-packager] 最新版本：{} ({})", outInfo.tag, outInfo.assetName));
	return true;
}

std::string DetectInstalledVersion(const std::filesystem::path& exePath)
{
	if (!std::filesystem::exists(exePath)) {
		return std::string();
	}

	ProcessRunResult result = RunProcessAndCaptureImpl(exePath, { L"version" }, exePath.parent_path());
	std::string text = BytesToLocalText(result.stdOutBytes);
	if (!result.stdErrBytes.empty()) {
		text += "\n";
		text += BytesToLocalText(result.stdErrBytes);
	}
	if (!result.ok && text.empty()) {
		return std::string();
	}
	return text;
}

bool WriteBinaryFile(const std::filesystem::path& path, const std::string& bytes, std::string& outError)
{
	std::ofstream out(path, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		outError = "无法写入文件：" + LocalPathString(path);
		return false;
	}
	out.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
	if (!out.good()) {
		outError = "写入文件失败：" + LocalPathString(path);
		return false;
	}
	return true;
}

bool DownloadZip(const LatestReleaseInfo& info, const std::filesystem::path& zipPath, std::string& outError)
{
	OutputStringToELog(std::format("[e-packager] 开始下载：{}", info.downloadUrl));
	auto response = PerformGetRequest(info.downloadUrl, kGitHubHeaders, 300000, false, false);
	if (response.second != 200) {
		outError = std::format("下载失败，HTTP {}", response.second);
		if (!response.first.empty()) {
			outError += ": " + response.first.substr(0, (std::min<size_t>)(response.first.size(), 300));
		}
		return false;
	}

	if (!WriteBinaryFile(zipPath, response.first, outError)) {
		return false;
	}

	OutputStringToELog(std::format("[e-packager] 下载完成，大小 {} 字节", response.first.size()));
	return true;
}

std::string EscapePowerShellSingleQuoted(const std::string& text)
{
	std::string escaped;
	escaped.reserve(text.size() + 8);
	for (char ch : text) {
		escaped.push_back(ch);
		if (ch == '\'') {
			escaped.push_back('\'');
		}
	}
	return escaped;
}

bool ExtractZip(const std::filesystem::path& zipPath, const std::filesystem::path& destination, std::string& outError)
{
	OutputStringToELog("[e-packager] 正在解压工具包...");
	const std::string command =
		"Expand-Archive -LiteralPath '" + EscapePowerShellSingleQuoted(Utf8PathString(zipPath)) +
		"' -DestinationPath '" + EscapePowerShellSingleQuoted(Utf8PathString(destination)) + "' -Force";
	PowerShellRunResult result = PowerShellToolRunner::Run(command, Utf8PathString(GetToolsDirectory()), 120);
	if (!result.stdOut.empty()) {
		OutputTextBlock("[e-packager] 解压输出：", result.stdOut);
	}
	if (!result.stdErr.empty()) {
		OutputTextBlock("[e-packager] 解压错误输出：", result.stdErr);
	}
	if (!result.ok || result.exitCode != 0) {
		outError = result.error.empty()
			? std::format("Expand-Archive exitCode={}", result.exitCode)
			: result.error;
		return false;
	}
	return true;
}

std::filesystem::path FindExtractedExe(const std::filesystem::path& root)
{
	std::error_code ec;
	for (const auto& entry : std::filesystem::recursive_directory_iterator(root, ec)) {
		if (ec) {
			break;
		}
		if (!entry.is_regular_file(ec)) {
			continue;
		}
		if (_wcsicmp(entry.path().filename().c_str(), L"e-packager.exe") == 0) {
			return entry.path();
		}
	}
	return {};
}

bool CopyExtractedPackage(const std::filesystem::path& extractedExe, std::string& outError)
{
	if (extractedExe.empty()) {
		outError = "解压后未找到 e-packager.exe";
		return false;
	}

	std::error_code ec;
	const std::filesystem::path packageRoot = extractedExe.parent_path();
	const std::filesystem::path toolsDir = GetToolsDirectory();
	for (const auto& entry : std::filesystem::directory_iterator(packageRoot, ec)) {
		if (ec) {
			outError = "枚举解压目录失败：" + ec.message();
			return false;
		}
		const auto target = toolsDir / entry.path().filename();
		std::filesystem::copy(
			entry.path(),
			target,
			std::filesystem::copy_options::recursive | std::filesystem::copy_options::overwrite_existing,
			ec);
		if (ec) {
			outError = "复制工具文件失败：" + ec.message();
			return false;
		}
	}
	return true;
}

bool DownloadAndInstallTool(const LatestReleaseInfo& info, std::string& outError)
{
	const std::filesystem::path toolsDir = GetToolsDirectory();
	std::error_code ec;
	std::filesystem::create_directories(toolsDir, ec);
	if (ec) {
		outError = "创建 tools 目录失败：" + ec.message();
		return false;
	}

	const ULONGLONG tick = GetTickCount64();
	const DWORD pid = GetCurrentProcessId();
	const std::filesystem::path tempRoot = toolsDir / std::format(L"e-packager.download.{}.{}", pid, tick);
	const std::filesystem::path zipPath = tempRoot / L"e-packager.zip";
	const std::filesystem::path extractDir = tempRoot / L"extract";
	std::filesystem::create_directories(extractDir, ec);
	if (ec) {
		outError = "创建临时目录失败：" + ec.message();
		return false;
	}

	bool ok = DownloadZip(info, zipPath, outError) &&
		ExtractZip(zipPath, extractDir, outError) &&
		CopyExtractedPackage(FindExtractedExe(extractDir), outError);

	std::filesystem::remove_all(tempRoot, ec);
	if (!ok) {
		return false;
	}

	SaveMeta(info);
	OutputStringToELog("[e-packager] 工具已安装到：" + LocalPathString(GetEPackagerExePath()));
	return true;
}

bool EnsureToolReadyImpl(std::filesystem::path& outToolPath, std::string& outError)
{
	const std::filesystem::path toolPath = GetEPackagerExePath();
	const bool toolExists = std::filesystem::exists(toolPath);
	if (!IsUpdateCheckDue(toolExists)) {
		outToolPath = toolPath;
		return true;
	}

	LatestReleaseInfo latest;
	std::string fetchError;
	if (!FetchLatestRelease(latest, fetchError)) {
		if (toolExists) {
			OutputStringToELog("[e-packager] 检查更新失败，将继续使用现有工具：" + fetchError);
			outToolPath = toolPath;
			return true;
		}
		outError = "无法下载 e-packager：" + fetchError;
		return false;
	}

	bool needsDownload = !toolExists;
	if (toolExists) {
		const std::string installedVersion = ToLowerAscii(DetectInstalledVersion(toolPath));
		const std::string latestTag = ToLowerAscii(latest.tag);
		const std::string latestTagNoPrefix = StripLeadingVersionPrefix(latestTag);
		if (installedVersion.find(latestTag) == std::string::npos &&
			installedVersion.find(latestTagNoPrefix) == std::string::npos) {
			needsDownload = true;
		}
	}

	if (!needsDownload) {
		OutputStringToELog("[e-packager] 本地工具已是最新版本：" + latest.tag);
		SaveMeta(latest);
		outToolPath = toolPath;
		return true;
	}

	if (!DownloadAndInstallTool(latest, outError)) {
		if (toolExists) {
			OutputStringToELog("[e-packager] 更新失败，将继续使用现有工具：" + outError);
			outError.clear();
			outToolPath = toolPath;
			return true;
		}
		return false;
	}

	outToolPath = toolPath;
	return true;
}

bool IsCurrentSourceEFile()
{
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		return false;
	}
	std::filesystem::path sourcePath = PathFromLocal(g_nowOpenSourceFilePath);
	const std::wstring ext = sourcePath.extension().wstring();
	return _wcsicmp(ext.c_str(), L".e") == 0;
}

std::wstring CurrentSourceFileNameW()
{
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		return std::wstring();
	}
	return PathFromLocal(g_nowOpenSourceFilePath).filename().wstring();
}

bool PickOutputParentDirectoryLegacy(std::filesystem::path& outDirectory)
{
	BROWSEINFOW browseInfo = {};
	browseInfo.hwndOwner = g_hwnd;
	browseInfo.lpszTitle = L"选择反编译输出目录";
	browseInfo.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_USENEWUI;

	PIDLIST_ABSOLUTE itemList = SHBrowseForFolderW(&browseInfo);
	if (itemList == nullptr) {
		return false;
	}

	wchar_t pathBuffer[MAX_PATH] = {};
	const BOOL ok = SHGetPathFromIDListW(itemList, pathBuffer);
	CoTaskMemFree(itemList);
	if (ok == FALSE || pathBuffer[0] == L'\0') {
		return false;
	}

	outDirectory = std::filesystem::path(pathBuffer);
	return true;
}

bool PickOutputParentDirectory(std::filesystem::path& outDirectory)
{
	ScopedComInit com;

	IFileOpenDialog* dialog = nullptr;
	HRESULT hr = CoCreateInstance(
		CLSID_FileOpenDialog,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&dialog));
	if (FAILED(hr) || dialog == nullptr) {
		OutputStringToELog(std::format("[e-packager] 创建目录选择窗口失败：0x{:08X}", static_cast<unsigned int>(hr)));
		return PickOutputParentDirectoryLegacy(outDirectory);
	}

	DWORD options = 0;
	if (SUCCEEDED(dialog->GetOptions(&options))) {
		dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
	}
	dialog->SetTitle(L"选择反编译输出目录");

	hr = dialog->Show(g_hwnd);
	if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) {
		dialog->Release();
		return false;
	}
	if (FAILED(hr)) {
		OutputStringToELog(std::format("[e-packager] 目录选择失败：0x{:08X}", static_cast<unsigned int>(hr)));
		dialog->Release();
		return PickOutputParentDirectoryLegacy(outDirectory);
	}

	IShellItem* item = nullptr;
	hr = dialog->GetResult(&item);
	dialog->Release();
	if (FAILED(hr) || item == nullptr) {
		OutputStringToELog("[e-packager] 目录选择失败：无法读取选择结果");
		return false;
	}

	PWSTR rawPath = nullptr;
	hr = item->GetDisplayName(SIGDN_FILESYSPATH, &rawPath);
	item->Release();
	if (FAILED(hr) || rawPath == nullptr) {
		OutputStringToELog("[e-packager] 目录选择失败：无法读取目录路径");
		return false;
	}

	outDirectory = std::filesystem::path(rawPath);
	CoTaskMemFree(rawPath);
	return true;
}

void OpenDirectoryInExplorer(const std::filesystem::path& directory)
{
	ShellExecuteW(g_hwnd, L"open", L"explorer.exe", QuoteCommandLineArg(directory.wstring()).c_str(), nullptr, SW_SHOWNORMAL);
}

void UnpackWorker(void* param)
{
	std::unique_ptr<UnpackRequest> request(static_cast<UnpackRequest*>(param));
	if (!request) {
		g_unpackTaskRunning.store(false);
		return;
	}

	const auto originalSourcePath = request->originalSourcePath;
	const auto snapshotPath = request->snapshotPath;
	const auto snapshotRoot = request->snapshotRoot;
	const auto unpackDir = request->unpackDir;

	try {
		OutputStringToELog("[e-packager] 开始反编译当前内存快照：" + LocalPathString(originalSourcePath));
		OutputStringToELog("[e-packager] 快照文件：" + LocalPathString(snapshotPath));
		std::error_code ec;
		std::filesystem::create_directories(unpackDir, ec);
		if (ec) {
			OutputStringToELog("[e-packager] 创建输出目录失败：" + ec.message());
			CleanupSnapshotRootImpl(snapshotRoot);
			g_unpackTaskRunning.store(false);
			return;
		}

		std::filesystem::path toolPath;
		std::string error;
		if (!EnsureToolReadyImpl(toolPath, error)) {
			OutputStringToELog("[e-packager] " + error);
			CleanupSnapshotRootImpl(snapshotRoot);
			g_unpackTaskRunning.store(false);
			return;
		}

		OutputStringToELog("[e-packager] 输出目录：" + LocalPathString(unpackDir));
		ProcessRunResult result = RunProcessAndCaptureImpl(
			toolPath,
			{ L"unpack", snapshotPath.wstring(), unpackDir.wstring() },
			toolPath.parent_path());
		OutputTextBlock("[e-packager] 标准输出：", result.stdOutBytes);
		OutputTextBlock("[e-packager] 错误输出：", result.stdErrBytes);

		if (!result.ok) {
			OutputStringToELog(std::format(
				"[e-packager] 反编译失败，exitCode={} {}",
				result.exitCode,
				result.error));
			CleanupSnapshotRootImpl(snapshotRoot);
			g_unpackTaskRunning.store(false);
			return;
		}

		OutputStringToELog("[e-packager] 反编译完成");
		CleanupSnapshotRootImpl(snapshotRoot);
		OpenDirectoryInExplorer(unpackDir);
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[e-packager] 反编译异常：") + ex.what());
		CleanupSnapshotRootImpl(snapshotRoot);
	}
	catch (...) {
		OutputStringToELog("[e-packager] 反编译发生未知异常");
		CleanupSnapshotRootImpl(snapshotRoot);
	}

	g_unpackTaskRunning.store(false);
}

} // namespace

bool GetCurrentSourcePath(std::filesystem::path& outSourcePath, std::string& outError)
{
	outSourcePath.clear();
	outError.clear();
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		outError = "当前没有打开易语言源码文件";
		return false;
	}

	outSourcePath = PathFromLocal(g_nowOpenSourceFilePath);
	const std::wstring ext = outSourcePath.extension().wstring();
	if (_wcsicmp(ext.c_str(), L".e") != 0) {
		outError = "当前打开的不是 .e 源码文件";
		return false;
	}
	return true;
}

std::filesystem::path BuildCurrentProjectSnapshotPathForSource(const std::filesystem::path& sourcePath)
{
	return BuildCurrentProjectSnapshotPath(sourcePath);
}

bool WriteCurrentProjectSnapshot(
	const std::filesystem::path& snapshotPath,
	size_t& outBytesWritten,
	std::string& outTrace,
	std::string& outError)
{
	return WriteCurrentProjectSnapshotImpl(snapshotPath, outBytesWritten, outTrace, outError);
}

void CleanupSnapshotRoot(const std::filesystem::path& snapshotRoot)
{
	CleanupSnapshotRootImpl(snapshotRoot);
}

bool EnsureToolReady(std::filesystem::path& outToolPath, std::string& outError)
{
	return EnsureToolReadyImpl(outToolPath, outError);
}

ProcessRunResult RunProcessAndCapture(
	const std::filesystem::path& exePath,
	const std::vector<std::wstring>& args,
	const std::filesystem::path& workingDirectory)
{
	return RunProcessAndCaptureImpl(exePath, args, workingDirectory);
}

std::wstring BuildUnpackMenuTitle()
{
	const std::wstring filename = CurrentSourceFileNameW();
	if (filename.empty()) {
		return L"将当前.e反编译到目录";
	}
	return L"将[" + filename + L"]反编译到目录";
}

bool CanUnpackCurrentSource()
{
	return IsCurrentSourceEFile();
}

void RunCurrentSourceUnpackToDirectory()
{
	UpdateCurrentOpenSourceFile();
	if (!IsCurrentSourceEFile()) {
		OutputStringToELog("[e-packager] 当前没有打开 .e 源文件，无法反编译");
		return;
	}

	if (g_unpackTaskRunning.exchange(true)) {
		OutputStringToELog("[e-packager] 已有反编译任务正在执行，请稍候");
		return;
	}

	std::filesystem::path parentDirectory;
	if (!PickOutputParentDirectory(parentDirectory)) {
		g_unpackTaskRunning.store(false);
		OutputStringToELog("[e-packager] 已取消选择输出目录");
		return;
	}

	const std::filesystem::path sourcePath = PathFromLocal(g_nowOpenSourceFilePath);
	const std::filesystem::path unpackDir = parentDirectory / (sourcePath.filename().wstring() + L".unpack");
	const std::filesystem::path snapshotPath = BuildCurrentProjectSnapshotPath(sourcePath);

	size_t snapshotBytes = 0;
	std::string snapshotTrace;
	std::string snapshotError;
	if (!WriteCurrentProjectSnapshotImpl(snapshotPath, snapshotBytes, snapshotTrace, snapshotError)) {
		g_unpackTaskRunning.store(false);
		OutputStringToELog("[e-packager] 内存快照导出失败：" + snapshotError);
		if (!snapshotTrace.empty()) {
			OutputStringToELog("[e-packager] 快照导出诊断：" + snapshotTrace);
		}
		OutputStringToELog("[e-packager] 不会替用户保存源文件，请先触发一次 IDE 自动备份或打开工程后再试");
		CleanupSnapshotRootImpl(snapshotPath.parent_path());
		return;
	}

	OutputStringToELog(std::format(
		"[e-packager] 已导出当前内存快照，bytes={} path={}",
		snapshotBytes,
		LocalPathString(snapshotPath)));
	if (!snapshotTrace.empty()) {
		OutputStringToELog("[e-packager] 快照导出诊断：" + snapshotTrace);
	}

	auto* request = new UnpackRequest{
		sourcePath,
		snapshotPath,
		snapshotPath.parent_path(),
		unpackDir
	};
	if (_beginthread(UnpackWorker, 0, request) == static_cast<uintptr_t>(-1)) {
		delete request;
		g_unpackTaskRunning.store(false);
		CleanupSnapshotRootImpl(snapshotPath.parent_path());
		OutputStringToELog("[e-packager] 启动后台反编译任务失败");
		return;
	}
}

} // namespace EPackagerIntegration
