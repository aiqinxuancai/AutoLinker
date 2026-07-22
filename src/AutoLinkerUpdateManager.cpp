#include "AutoLinkerUpdateManager.h"

#include <Windows.h>
#include <process.h>

#include <atomic>
#include <cctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <mutex>
#include <string>
#include <vector>

#include "AutoLinkerReleaseClient.h"
#include "AutoLinkerUpdateInstaller.h"
#include "AutoLinkerVersion.h"
#include "AIChatFeature.h"
#include "Global.h"
#include "PathHelper.h"
#include "PowerShellToolRunner.h"
#include "Version.h"

namespace AutoLinkerUpdateManager {
namespace {

std::atomic_bool g_updateRunning = false;
std::atomic_bool g_checkRunning = false;
std::mutex g_taskStartMutex;
std::mutex g_statusMutex;
ComponentUpdateStatus g_status = {
	ComponentUpdateState::Idle,
	AUTOLINKER_VERSION,
	{},
	"尚未检查更新。",
	-1
};
HWND g_statusNotificationWindow = nullptr;

void PublishStatus(
	ComponentUpdateState state,
	std::string message,
	std::string latestVersion = {},
	int progressPercent = -1)
{
	HWND notificationWindow = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_statusMutex);
		g_status.state = state;
		g_status.currentVersion = AUTOLINKER_VERSION;
		if (!latestVersion.empty()) {
			g_status.latestVersion = std::move(latestVersion);
		}
		g_status.message = std::move(message);
		g_status.progressPercent = progressPercent;
		notificationWindow = g_statusNotificationWindow;
	}
	if (notificationWindow != nullptr && IsWindow(notificationWindow)) {
		PostMessageW(notificationWindow, WM_AUTOLINKER_COMPONENT_UPDATE_STATUS, 0, 0);
	}
}

struct UpdateRunningGuard {
	~UpdateRunningGuard()
	{
		g_updateRunning.store(false);
	}
};

std::wstring WideFromUtf8(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (length <= 0) {
		return {};
	}
	std::wstring result(static_cast<size_t>(length), L'\0');
	if (MultiByteToWideChar(
			CP_UTF8,
			MB_ERR_INVALID_CHARS,
			text.data(),
			static_cast<int>(text.size()),
			result.data(),
			length) <= 0) {
		return {};
	}
	return result;
}

std::string Utf8FromWide(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (length <= 0) {
		return {};
	}
	std::string result(static_cast<size_t>(length), '\0');
	if (WideCharToMultiByte(
			CP_UTF8,
			0,
			text.data(),
			static_cast<int>(text.size()),
			result.data(),
			length,
			nullptr,
			nullptr) <= 0) {
		return {};
	}
	return result;
}

std::string LocalFromUtf8(const std::string& text)
{
	const std::wstring wide = WideFromUtf8(text);
	if (wide.empty()) {
		return text;
	}
	const int length = WideCharToMultiByte(
		CP_ACP, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, nullptr, nullptr);
	if (length <= 0) {
		return text;
	}
	std::string result(static_cast<size_t>(length), '\0');
	WideCharToMultiByte(
		CP_ACP, 0, wide.data(), static_cast<int>(wide.size()), result.data(), length, nullptr, nullptr);
	return result;
}

void OutputUpdateLog(const std::string& text)
{
	OutputStringToELog(LocalFromUtf8(text));
}

std::string NormalizeVersionText(std::string version)
{
	while (!version.empty() && std::isspace(static_cast<unsigned char>(version.front())) != 0) {
		version.erase(version.begin());
	}
	while (!version.empty() && std::isspace(static_cast<unsigned char>(version.back())) != 0) {
		version.pop_back();
	}
	if (!version.empty() && (version.front() == 'v' || version.front() == 'V')) {
		version.erase(version.begin());
	}
	return version;
}

bool IsNewerVersion(const std::string& latest, const std::string& current)
{
	const std::string normalizedLatest = NormalizeVersionText(latest);
	const std::string normalizedCurrent = NormalizeVersionText(current);
	try {
		return Version(normalizedLatest) > Version(normalizedCurrent);
	}
	catch (...) {
		return !normalizedLatest.empty() && normalizedLatest != normalizedCurrent;
	}
}

bool WriteBinaryFile(const std::filesystem::path& path, const std::string& bytes, std::string& outError)
{
	std::ofstream output(path, std::ios::binary | std::ios::trunc);
	if (!output.is_open()) {
		outError = "无法写入更新文件：" + Utf8FromWide(path.wstring());
		return false;
	}
	output.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
	if (!output.good()) {
		outError = "写入更新文件失败：" + Utf8FromWide(path.wstring());
		return false;
	}
	return true;
}

std::string EscapePowerShellSingleQuoted(const std::string& text)
{
	std::string escaped;
	escaped.reserve(text.size() + 8);
	for (const char ch : text) {
		escaped.push_back(ch);
		if (ch == '\'') {
			escaped.push_back('\'');
		}
	}
	return escaped;
}

std::string PowerShellLiteral(const std::filesystem::path& path)
{
	return "'" + EscapePowerShellSingleQuoted(Utf8FromWide(path.wstring())) + "'";
}

bool ExtractArchive(
	const std::filesystem::path& archivePath,
	const std::filesystem::path& destination,
	std::string& outError)
{
	const std::string command =
		"Expand-Archive -LiteralPath " + PowerShellLiteral(archivePath) +
		" -DestinationPath " + PowerShellLiteral(destination) + " -Force";
	const PowerShellRunResult result = PowerShellToolRunner::Run(
		command,
		Utf8FromWide(destination.parent_path().wstring()),
		120);
	if (!result.ok || result.exitCode != 0) {
		outError = result.error.empty() ? result.stdErr : result.error;
		if (outError.empty()) {
			outError = "Expand-Archive exitCode=" + std::to_string(result.exitCode);
		}
		return false;
	}
	return true;
}

std::filesystem::path FindExtractedFne(const std::filesystem::path& root)
{
	std::error_code ec;
	for (std::filesystem::recursive_directory_iterator it(root, ec), end; it != end && !ec; it.increment(ec)) {
		if (!it->is_regular_file(ec) || ec) {
			continue;
		}
		if (_wcsicmp(it->path().filename().c_str(), L"AutoLinker.fne") == 0) {
			return it->path();
		}
	}
	return {};
}

bool ValidateWin32Fne(const std::filesystem::path& path, std::string& outError)
{
	std::ifstream input(path, std::ios::binary);
	if (!input.is_open()) {
		outError = "无法读取解压后的 AutoLinker.fne";
		return false;
	}

	IMAGE_DOS_HEADER dosHeader = {};
	input.read(reinterpret_cast<char*>(&dosHeader), sizeof(dosHeader));
	if (!input.good() || dosHeader.e_magic != IMAGE_DOS_SIGNATURE || dosHeader.e_lfanew <= 0) {
		outError = "解压后的 AutoLinker.fne 不是有效的 PE 文件";
		return false;
	}
	input.seekg(dosHeader.e_lfanew, std::ios::beg);
	DWORD signature = 0;
	IMAGE_FILE_HEADER fileHeader = {};
	input.read(reinterpret_cast<char*>(&signature), sizeof(signature));
	input.read(reinterpret_cast<char*>(&fileHeader), sizeof(fileHeader));
	if (!input.good() || signature != IMAGE_NT_SIGNATURE ||
		fileHeader.Machine != IMAGE_FILE_MACHINE_I386 ||
		(fileHeader.Characteristics & IMAGE_FILE_DLL) == 0) {
		outError = "更新包中的 AutoLinker.fne 不是 Win32 DLL";
		return false;
	}
	return true;
}

std::filesystem::path GetIdeExecutablePath()
{
	std::vector<wchar_t> buffer(MAX_PATH);
	for (;;) {
		const DWORD length = GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
		if (length == 0) {
			return {};
		}
		if (length < buffer.size()) {
			return std::filesystem::path(std::wstring(buffer.data(), length));
		}
		buffer.resize(buffer.size() * 2);
	}
}

void ShowError(const std::string& message)
{
	MessageBoxW(
		g_hwnd,
		WideFromUtf8(message).c_str(),
		L"AutoLinker 更新",
		MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
}

void CleanupStaging(std::filesystem::path& stagingRoot)
{
	if (stagingRoot.empty()) {
		return;
	}
	std::error_code ec;
	std::filesystem::remove_all(stagingRoot, ec);
	stagingRoot.clear();
}

void UpdateWorker(void*)
{
	UpdateRunningGuard guard;
	std::filesystem::path stagingRoot;
	try {
		PublishStatus(ComponentUpdateState::Checking, "正在检查最新 Release...");
		OutputUpdateLog("[AutoLinker更新] 正在检查最新 Release...");
		AutoLinkerReleaseInfo release;
		std::string error;
		if (!AutoLinkerReleaseClient::FetchLatest(release, error)) {
			PublishStatus(ComponentUpdateState::Error, error);
			OutputUpdateLog("[AutoLinker更新] " + error);
			ShowError(error);
			return;
		}

		const std::string currentVersion = AUTOLINKER_VERSION;
		const std::wstring currentVersionWide = WideFromUtf8(currentVersion);
		const std::wstring releaseTagWide = WideFromUtf8(release.tag);
		if (!IsNewerVersion(release.tag, currentVersion)) {
			PublishStatus(ComponentUpdateState::UpToDate, "当前已是最新版本。", release.tag);
			const std::wstring message =
				L"当前版本：" + currentVersionWide +
				L"\r\n最新 Release：" + releaseTagWide +
				L"\r\n\r\n当前已是最新版本。";
			MessageBoxW(g_hwnd, message.c_str(), L"AutoLinker 更新", MB_OK | MB_ICONINFORMATION);
			return;
		}
		PublishStatus(ComponentUpdateState::UpdateAvailable, "发现可用的新版本。", release.tag);

		const std::filesystem::path ideExecutable = GetIdeExecutablePath();
		if (ideExecutable.empty()) {
			PublishStatus(ComponentUpdateState::Error, "无法获取当前易语言 IDE 路径。", release.tag);
			ShowError("无法获取当前易语言 IDE 路径，不能执行更新。");
			return;
		}
		const std::filesystem::path targetFne = ideExecutable.parent_path() / L"lib" / L"AutoLinker.fne";
		std::error_code ec;
		if (!std::filesystem::is_regular_file(targetFne, ec)) {
			PublishStatus(ComponentUpdateState::Error, "未找到当前 AutoLinker.fne。", release.tag);
			ShowError("未找到当前支持库：" + Utf8FromWide(targetFne.wstring()));
			return;
		}

		const std::wstring prompt =
			L"更新 AutoLinker.fne 需要结束当前易语言 IDE 进程后才能替换支持库。\r\n\r\n"
			L"请确认当前打开的源文件已经保存。\r\n\r\n"
			L"当前版本：" + currentVersionWide +
			L"\r\n将要更新到：" + releaseTagWide +
			L"\r\n\r\n下载并准备完成后，AutoLinker 将请求关闭当前 IDE。是否继续更新？";
		if (MessageBoxW(
				g_hwnd,
				prompt.c_str(),
				L"更新 AutoLinker 支持库",
				MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2 | MB_SETFOREGROUND) != IDYES) {
			OutputUpdateLog("[AutoLinker更新] 用户取消更新");
			PublishStatus(ComponentUpdateState::UpdateAvailable, "已取消本次更新。", release.tag);
			return;
		}

		stagingRoot = GetAutoLinkerCacheDirectoryPath() /
			std::format(L"AutoLinkerUpdate.{}.{}", GetCurrentProcessId(), GetTickCount64());
		const std::filesystem::path archivePath = stagingRoot / L"AutoLinker.zip";
		const std::filesystem::path extractDirectory = stagingRoot / L"extract";
		std::filesystem::create_directories(extractDirectory, ec);
		if (ec) {
			PublishStatus(ComponentUpdateState::Error, "创建更新临时目录失败：" + ec.message(), release.tag);
			ShowError("创建更新临时目录失败：" + ec.message());
			CleanupStaging(stagingRoot);
			return;
		}

		PublishStatus(ComponentUpdateState::Downloading, "正在下载并校验更新包...", release.tag);
		std::string archiveBytes;
		if (!AutoLinkerReleaseClient::DownloadArchive(release.asset, archiveBytes, error) ||
			!WriteBinaryFile(archivePath, archiveBytes, error)) {
			PublishStatus(ComponentUpdateState::Error, "下载更新包失败：" + error, release.tag);
			OutputUpdateLog("[AutoLinker更新] 下载失败：" + error);
			ShowError("下载 AutoLinker 更新包失败：\r\n" + error);
			CleanupStaging(stagingRoot);
			return;
		}
		OutputUpdateLog("[AutoLinker更新] 下载完成，" + std::to_string(archiveBytes.size()) + " 字节");
		archiveBytes.clear();
		archiveBytes.shrink_to_fit();

		if (!ExtractArchive(archivePath, extractDirectory, error)) {
			PublishStatus(ComponentUpdateState::Error, "解压更新包失败：" + error, release.tag);
			ShowError("解压 AutoLinker 更新包失败：\r\n" + error);
			CleanupStaging(stagingRoot);
			return;
		}
		const std::filesystem::path extractedFne = FindExtractedFne(extractDirectory);
		if (extractedFne.empty() || !ValidateWin32Fne(extractedFne, error)) {
			PublishStatus(
				ComponentUpdateState::Error,
				error.empty() ? "更新包中未找到 AutoLinker.fne。" : error,
				release.tag);
			ShowError(error.empty() ? "更新包中未找到 AutoLinker.fne" : error);
			CleanupStaging(stagingRoot);
			return;
		}

		const std::filesystem::path stagedFne = stagingRoot / L"AutoLinker.fne.new";
		std::filesystem::copy_file(extractedFne, stagedFne, std::filesystem::copy_options::overwrite_existing, ec);
		if (ec) {
			PublishStatus(ComponentUpdateState::Error, "暂存 AutoLinker.fne 失败：" + ec.message(), release.tag);
			ShowError("暂存 AutoLinker.fne 失败：" + ec.message());
			CleanupStaging(stagingRoot);
			return;
		}

		AutoLinkerUpdateInstallRequest installRequest;
		PublishStatus(ComponentUpdateState::Installing, "正在准备退出 IDE 后的替换任务...", release.tag);
		installRequest.processId = GetCurrentProcessId();
		installRequest.stagedFne = stagedFne;
		installRequest.targetFne = targetFne;
		installRequest.stagingRoot = stagingRoot;
		installRequest.logPath = GetAutoLinkerLogDirectoryPath() / L"autolinker_update.log";
		installRequest.targetVersion = release.tag;
		PROCESS_INFORMATION updaterProcess = {};
		if (!AutoLinkerUpdateInstaller::Launch(installRequest, updaterProcess, error)) {
			PublishStatus(ComponentUpdateState::Error, error, release.tag);
			ShowError(error);
			CleanupStaging(stagingRoot);
			return;
		}

		OutputUpdateLog(
			"[AutoLinker更新] 更新包已准备完成，退出 IDE 后将替换：" +
			Utf8FromWide(targetFne.wstring()));
		PublishStatus(ComponentUpdateState::ReadyToRestart, "更新已准备完成，IDE 退出后将自动替换。", release.tag);
		if (PostMessageW(g_hwnd, WM_CLOSE, 0, 0) == FALSE) {
			TerminateProcess(updaterProcess.hProcess, 1);
			WaitForSingleObject(updaterProcess.hProcess, 5000);
			CloseHandle(updaterProcess.hThread);
			CloseHandle(updaterProcess.hProcess);
			PublishStatus(ComponentUpdateState::Error, "无法请求关闭当前易语言 IDE。", release.tag);
			ShowError("无法请求关闭当前易语言 IDE，更新已取消。");
			CleanupStaging(stagingRoot);
			return;
		}

		CloseHandle(updaterProcess.hThread);
		CloseHandle(updaterProcess.hProcess);
		stagingRoot.clear();
	}
	catch (const std::exception& ex) {
		const std::string error = std::string("AutoLinker 更新异常：") + ex.what();
		OutputUpdateLog("[AutoLinker更新] " + error);
		PublishStatus(ComponentUpdateState::Error, error);
		ShowError(error);
	}
	catch (...) {
		OutputUpdateLog("[AutoLinker更新] 发生未知异常");
		PublishStatus(ComponentUpdateState::Error, "AutoLinker 更新发生未知异常。");
		ShowError("AutoLinker 更新发生未知异常。");
	}
	CleanupStaging(stagingRoot);
}

void CheckWorker(void*)
{
	try {
		PublishStatus(ComponentUpdateState::Checking, "正在检查最新 Release...");
		AutoLinkerReleaseInfo release;
		std::string error;
		if (!AutoLinkerReleaseClient::FetchLatest(release, error)) {
			PublishStatus(ComponentUpdateState::Error, error);
			g_checkRunning.store(false);
			return;
		}
		if (IsNewerVersion(release.tag, AUTOLINKER_VERSION)) {
			PublishStatus(ComponentUpdateState::UpdateAvailable, "发现可用的新版本。", release.tag);
			AIChatFeature::SetUpdateAvailable(release.tag);
		}
		else {
			PublishStatus(ComponentUpdateState::UpToDate, "当前已是最新版本。", release.tag);
		}
	}
	catch (const std::exception& ex) {
		PublishStatus(ComponentUpdateState::Error, std::string("检查更新失败：") + ex.what());
	}
	catch (...) {
		PublishStatus(ComponentUpdateState::Error, "检查更新发生未知异常。");
	}
	g_checkRunning.store(false);
}

} // namespace

void CheckForUpdatesInBackground()
{
	{
		std::lock_guard<std::mutex> lock(g_taskStartMutex);
		if (g_updateRunning.load(std::memory_order_acquire) ||
			g_checkRunning.load(std::memory_order_acquire)) {
			return;
		}
		g_checkRunning.store(true, std::memory_order_release);
	}
	if (_beginthread(CheckWorker, 0, nullptr) == static_cast<uintptr_t>(-1)) {
		g_checkRunning.store(false);
		PublishStatus(ComponentUpdateState::Error, "启动后台检查任务失败。");
	}
}

void RunUpdateInBackground()
{
	bool checkRunning = false;
	bool updateRunning = false;
	{
		std::lock_guard<std::mutex> lock(g_taskStartMutex);
		checkRunning = g_checkRunning.load(std::memory_order_acquire);
		updateRunning = g_updateRunning.load(std::memory_order_acquire);
		if (!checkRunning && !updateRunning) {
			g_updateRunning.store(true, std::memory_order_release);
		}
	}
	if (checkRunning) {
		MessageBoxW(
			g_hwnd,
			L"正在检查 AutoLinker 最新版本，请稍候。",
			L"AutoLinker 更新",
			MB_OK | MB_ICONINFORMATION);
		return;
	}
	if (updateRunning) {
		MessageBoxW(
			g_hwnd,
			L"已有 AutoLinker 更新任务正在执行，请稍候。",
			L"AutoLinker 更新",
			MB_OK | MB_ICONINFORMATION);
		return;
	}

	if (_beginthread(UpdateWorker, 0, nullptr) == static_cast<uintptr_t>(-1)) {
		g_updateRunning.store(false);
		PublishStatus(ComponentUpdateState::Error, "启动 AutoLinker 后台更新任务失败。");
		ShowError("启动 AutoLinker 后台更新任务失败。");
	}
}

void SetStatusNotificationWindow(HWND window)
{
	std::lock_guard<std::mutex> lock(g_statusMutex);
	g_statusNotificationWindow = window != nullptr && IsWindow(window) ? window : nullptr;
}

ComponentUpdateStatus GetStatus()
{
	std::lock_guard<std::mutex> lock(g_statusMutex);
	return g_status;
}

} // namespace AutoLinkerUpdateManager
