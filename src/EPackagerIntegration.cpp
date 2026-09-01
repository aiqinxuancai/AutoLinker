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
#include <mutex>
#include <sstream>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AutoLinkerInternal.h"
#include "ArchiveExtractor.h"
#include "DownloadProgressReporter.h"
#include "EideProjectBinarySerializer.h"
#include "Global.h"
#include "PathHelper.h"
#include "WinINetUtil.h"

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")

namespace EPackagerIntegration {
namespace {

using json = nlohmann::json;

constexpr const char* kLatestReleaseApi = "https://api.github.com/repos/aiqinxuancai/e-packager/releases/latest";
constexpr const char* kGitHubBaseUrl = "https://github.com/";
constexpr const char* kGitHubAcceleratorBaseUrl = "https://github-fast.apptest.dev/";
constexpr const char* kGitHubHeaders =
	"User-Agent: AutoLinker\r\n"
	"Accept: application/vnd.github+json\r\n";
constexpr long long kUpdateCheckIntervalSeconds = 7LL * 24LL * 60LL * 60LL;

std::atomic_bool g_unpackTaskRunning = false;
std::atomic_bool g_toolCheckTaskRunning = false;
std::atomic_bool g_toolUpdateTaskRunning = false;
std::mutex g_toolTaskStartMutex;
std::mutex g_updateStatusMutex;
std::string Utf8FromStatusText(const std::string& text);
ComponentUpdateStatus g_updateStatus = {
	ComponentUpdateState::Idle,
	{},
	{},
	Utf8FromStatusText("尚未检查组件更新。"),
	-1
};
HWND g_updateStatusNotificationWindow = nullptr;

void PublishUpdateStatus(
	ComponentUpdateState state,
	std::string message,
	std::string currentVersion = {},
	std::string latestVersion = {})
{
	message = Utf8FromStatusText(message);
	currentVersion = Utf8FromStatusText(currentVersion);
	latestVersion = Utf8FromStatusText(latestVersion);
	HWND notificationWindow = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_updateStatusMutex);
		g_updateStatus.state = state;
		if (!currentVersion.empty()) {
			g_updateStatus.currentVersion = std::move(currentVersion);
		}
		if (!latestVersion.empty()) {
			g_updateStatus.latestVersion = std::move(latestVersion);
		}
		g_updateStatus.message = std::move(message);
		g_updateStatus.progressPercent = -1;
		notificationWindow = g_updateStatusNotificationWindow;
	}
	if (notificationWindow != nullptr && IsWindow(notificationWindow)) {
		PostMessageW(notificationWindow, WM_AUTOLINKER_COMPONENT_UPDATE_STATUS, 1, 0);
	}
}

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
	std::uint64_t assetSize = 0;
};

LatestReleaseInfo g_checkedRelease;
bool g_checkedReleaseAvailable = false;

void ToolUpdateWorker(void* parameter);
void ToolMenuUpdateWorker(void*);

void ClearCheckedRelease()
{
	std::lock_guard<std::mutex> lock(g_updateStatusMutex);
	g_checkedRelease = {};
	g_checkedReleaseAvailable = false;
}

void StoreCheckedRelease(const LatestReleaseInfo& release)
{
	std::lock_guard<std::mutex> lock(g_updateStatusMutex);
	g_checkedRelease = release;
	g_checkedReleaseAvailable = true;
}

bool TryGetCheckedRelease(LatestReleaseInfo& release)
{
	std::lock_guard<std::mutex> lock(g_updateStatusMutex);
	if (!g_checkedReleaseAvailable || g_updateStatus.state != ComponentUpdateState::UpdateAvailable) {
		return false;
	}
	release = g_checkedRelease;
	return true;
}

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

std::string Utf8FromStatusText(const std::string& text)
{
	if (text.empty() || !WideFromUtf8(text).empty()) {
		return text;
	}
	const std::wstring wide = WideFromLocal(text);
	return wide.empty() ? std::string() : Utf8FromWide(wide);
}

std::string LocalFromUtf8(const std::string& text)
{
	const std::wstring wide = WideFromUtf8(text);
	return wide.empty() ? text : LocalFromWide(wide);
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

std::string BuildManualEPackagerInstallGuidance()
{
	return "如自动下载或解压失败，请从 https://github.com/aiqinxuancai/e-packager 或相关交流群群共享获取 e-packager，解压后将 e-packager.exe 放到 " +
		LocalPathString(GetEPackagerExePath()) +
		"，不要多套一层压缩包目录，然后重试。";
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

void ReadPipeToBytes(HANDLE pipe, std::string* output, std::atomic_bool* done)
{
	if (pipe == nullptr || output == nullptr) {
		if (done != nullptr) {
			done->store(true, std::memory_order_release);
		}
		return;
	}

	char buffer[4096] = {};
	DWORD bytesRead = 0;
	while (ReadFile(pipe, buffer, static_cast<DWORD>(sizeof(buffer)), &bytesRead, nullptr) != FALSE && bytesRead > 0) {
		output->append(buffer, bytesRead);
	}
	if (done != nullptr) {
		done->store(true, std::memory_order_release);
	}
}

constexpr DWORD kEPackagerProcessTimeoutMs = 180000;

bool IsIdeMainThread()
{
	if (g_hwnd == nullptr || !IsWindow(g_hwnd)) {
		return false;
	}
	const DWORD threadId = GetWindowThreadProcessId(g_hwnd, nullptr);
	return threadId != 0 && threadId == GetCurrentThreadId();
}

bool IsSafeMessageToDispatchWhileWaiting(const UINT message)
{
	return message == WM_PAINT ||
		message == WM_NCPAINT ||
		message == WM_ERASEBKGND ||
		message == WM_TIMER ||
		message == WM_SETCURSOR;
}

ProcessRunResult RunProcessAndCaptureImpl(
	const std::filesystem::path& exePath,
	const std::vector<std::wstring>& args,
	const std::filesystem::path& workingDirectory,
	const DWORD timeoutMs = kEPackagerProcessTimeoutMs)
{
	ProcessRunResult result = {};
	const DWORD effectiveTimeoutMs = timeoutMs == 0
		? kEPackagerProcessTimeoutMs
		: (std::min)(timeoutMs, kEPackagerProcessTimeoutMs);

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

	STARTUPINFOEXW si = {};
	si.StartupInfo.cb = sizeof(si);
	si.StartupInfo.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
	si.StartupInfo.wShowWindow = SW_HIDE;
	// e-packager is non-interactive; leaving stdin unset also avoids inheriting
	// an arbitrary IDE console handle through the restricted handle list.
	si.StartupInfo.hStdInput = nullptr;
	si.StartupInfo.hStdOutput = stdOutWrite;
	si.StartupInfo.hStdError = stdErrWrite;

	// Restrict inherited handles so a child process spawned by e-packager cannot
	// keep our anonymous pipes open after the main process exits.
	SIZE_T attributeListBytes = 0;
	InitializeProcThreadAttributeList(nullptr, 1, 0, &attributeListBytes);
	std::vector<BYTE> attributeStorage(attributeListBytes);
	LPPROC_THREAD_ATTRIBUTE_LIST attributeList =
		reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(attributeStorage.data());
	if (attributeListBytes == 0 ||
		InitializeProcThreadAttributeList(attributeList, 1, 0, &attributeListBytes) == FALSE) {
		CloseHandle(stdOutRead);
		CloseHandle(stdOutWrite);
		CloseHandle(stdErrRead);
		CloseHandle(stdErrWrite);
		result.error = std::format("InitializeProcThreadAttributeList failed, error={}", GetLastError());
		return result;
	}
	const HANDLE inheritedHandles[] = { stdOutWrite, stdErrWrite };
	if (UpdateProcThreadAttribute(
			attributeList,
			0,
			PROC_THREAD_ATTRIBUTE_HANDLE_LIST,
			const_cast<HANDLE*>(inheritedHandles),
			sizeof(inheritedHandles),
			nullptr,
			nullptr) == FALSE) {
		DeleteProcThreadAttributeList(attributeList);
		CloseHandle(stdOutRead);
		CloseHandle(stdOutWrite);
		CloseHandle(stdErrRead);
		CloseHandle(stdErrWrite);
		result.error = std::format("UpdateProcThreadAttribute handle list failed, error={}", GetLastError());
		return result;
	}
	si.lpAttributeList = attributeList;

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
		CREATE_NO_WINDOW | EXTENDED_STARTUPINFO_PRESENT,
		nullptr,
		cwd.empty() ? nullptr : cwd.c_str(),
		reinterpret_cast<LPSTARTUPINFOW>(&si),
		&pi);
	const DWORD createError = created == FALSE ? GetLastError() : ERROR_SUCCESS;
	DeleteProcThreadAttributeList(attributeList);

	CloseHandle(stdOutWrite);
	CloseHandle(stdErrWrite);

	if (created == FALSE) {
		CloseHandle(stdOutRead);
		CloseHandle(stdErrRead);
		result.error = std::format("CreateProcessW failed, error={}", createError);
		return result;
	}
	const ULONGLONG processStartTick = GetTickCount64();

	HANDLE job = CreateJobObjectW(nullptr, nullptr);
	if (job == nullptr) {
		const DWORD jobError = GetLastError();
		OutputStringToELog(std::format(
			"[e-packager] CreateJobObjectW unavailable, continuing without child cleanup job, error={}",
			jobError));
	}
	if (job != nullptr) {
		JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobInfo = {};
		jobInfo.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
		if (SetInformationJobObject(
				job,
				JobObjectExtendedLimitInformation,
				&jobInfo,
				static_cast<DWORD>(sizeof(jobInfo))) == FALSE ||
			AssignProcessToJobObject(job, pi.hProcess) == FALSE) {
			const DWORD jobAssignError = GetLastError();
			OutputStringToELog(std::format(
				"[e-packager] child cleanup job unavailable, continuing without it, error={}",
				jobAssignError));
			CloseHandle(job);
			job = nullptr;
		}
	}

	std::atomic_bool stdoutDone = false;
	std::atomic_bool stderrDone = false;
	std::thread stdoutReader(ReadPipeToBytes, stdOutRead, &result.stdOutBytes, &stdoutDone);
	std::thread stderrReader(ReadPipeToBytes, stdErrRead, &result.stdErrBytes, &stderrDone);

	DWORD waitResult = WAIT_FAILED;
	std::vector<MSG> deferredMessages;
	if (!IsIdeMainThread()) {
		waitResult = WaitForSingleObject(pi.hProcess, effectiveTimeoutMs);
	}
	else {
		// WorkspaceMirror is invoked by the IDE window procedure. Waiting on the
		// child directly would stop painting and make the IDE appear hung while a
		// large project is unpacked. Pump only redraw/timer messages here; tool,
		// command, and input messages are deferred to avoid re-entering the mirror
		// lock and are posted back after the child finishes.
		const ULONGLONG deadline = processStartTick + effectiveTimeoutMs;
		for (;;) {
			const ULONGLONG now = GetTickCount64();
			if (now >= deadline) {
				waitResult = WAIT_TIMEOUT;
				break;
			}
			const DWORD remaining = static_cast<DWORD>((std::min)(
				deadline - now,
				static_cast<ULONGLONG>(MAXDWORD)));
			const DWORD messageWait = MsgWaitForMultipleObjectsEx(
				1,
				&pi.hProcess,
				remaining,
				QS_ALLINPUT,
				MWMO_INPUTAVAILABLE);
			if (messageWait == WAIT_OBJECT_0) {
				waitResult = WAIT_OBJECT_0;
				break;
			}
			if (messageWait == WAIT_TIMEOUT) {
				waitResult = WAIT_TIMEOUT;
				break;
			}
			if (messageWait != WAIT_OBJECT_0 + 1) {
				waitResult = WAIT_FAILED;
				break;
			}

			MSG message = {};
			while (PeekMessageW(&message, nullptr, 0, 0, PM_REMOVE) != FALSE) {
				if (message.message == WM_QUIT) {
					deferredMessages.push_back(message);
					continue;
				}
				if (IsSafeMessageToDispatchWhileWaiting(message.message)) {
					TranslateMessage(&message);
					DispatchMessageW(&message);
				}
				else if (deferredMessages.size() < 1024) {
					deferredMessages.push_back(message);
				}
			}
		}
	}
	for (const MSG& message : deferredMessages) {
		if (message.message == WM_QUIT) {
			PostQuitMessage(static_cast<int>(message.wParam));
		}
		else if (message.hwnd != nullptr) {
			PostMessageW(message.hwnd, message.message, message.wParam, message.lParam);
		}
		else {
			PostThreadMessageW(GetCurrentThreadId(), message.message, message.wParam, message.lParam);
		}
	}
	if (waitResult == WAIT_TIMEOUT) {
		result.error = std::format("e-packager timed out after {} ms", effectiveTimeoutMs);
		OutputStringToELog("[e-packager] " + result.error);
		TerminateProcess(pi.hProcess, 124);
		WaitForSingleObject(pi.hProcess, 5000);
	}
	else if (waitResult == WAIT_FAILED) {
		result.error = std::format("WaitForSingleObject failed, error={}", GetLastError());
		TerminateProcess(pi.hProcess, 126);
		WaitForSingleObject(pi.hProcess, 5000);
	}
	result.elapsedMs = GetTickCount64() - processStartTick;
	DWORD exitCode = 0;
	if (GetExitCodeProcess(pi.hProcess, &exitCode) != FALSE) {
		result.exitCode = exitCode;
	}

	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);

	const auto finishPipeReader = [](std::thread& reader, std::atomic_bool& done) {
		if (!reader.joinable()) {
			return;
		}
		// Give a normally exiting process a short chance to drain its final bytes.
		for (int attempt = 0; attempt < 20 && !done.load(std::memory_order_acquire); ++attempt) {
			Sleep(5);
		}
		if (!done.load(std::memory_order_acquire)) {
			// A leaked descendant handle can leave ReadFile blocked forever.
			CancelSynchronousIo(static_cast<HANDLE>(reader.native_handle()));
		}
		reader.join();
	};
	finishPipeReader(stdoutReader, stdoutDone);
	finishPipeReader(stderrReader, stderrDone);
	CloseHandle(stdOutRead);
	CloseHandle(stdErrRead);
	CloseHandle(job);

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

std::string BytesToUtf8Text(const std::string& bytes)
{
	if (bytes.empty()) {
		return std::string();
	}

	std::wstring wide = WideFromUtf8(bytes);
	if (wide.empty()) {
		wide = WideFromCodePage(bytes, CP_ACP);
	}
	return wide.empty() ? std::string() : Utf8FromWide(wide);
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
		outError = response.second == 0
			? std::string("GitHub API 请求失败，未收到 HTTP 响应")
			: std::format("GitHub API HTTP {}", response.second);
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
		info.assetSize = asset.value("size", 0ULL);
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
	std::string text = BytesToUtf8Text(result.stdOutBytes);
	if (!result.stdErrBytes.empty()) {
		text += "\n";
		text += BytesToUtf8Text(result.stdErrBytes);
	}
	if (!result.ok && text.empty()) {
		return std::string();
	}
	return TrimAsciiCopy(std::move(text));
}

bool IsToolUpdateRequired(
	bool toolExists,
	const std::string& installedVersion,
	const LatestReleaseInfo& latest)
{
	if (!toolExists) {
		return true;
	}
	const std::string normalizedInstalledVersion = ToLowerAscii(installedVersion);
	const std::string latestTag = ToLowerAscii(latest.tag);
	const std::string latestTagNoPrefix = StripLeadingVersionPrefix(latestTag);
	return normalizedInstalledVersion.find(latestTag) == std::string::npos &&
		normalizedInstalledVersion.find(latestTagNoPrefix) == std::string::npos;
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

std::string BuildAcceleratedGitHubUrl(const std::string& url)
{
	const std::string githubBaseUrl = kGitHubBaseUrl;
	if (url.rfind(githubBaseUrl, 0) != 0) {
		return url;
	}
	return std::string(kGitHubAcceleratorBaseUrl) + url.substr(githubBaseUrl.size());
}

bool IsZipResponse(const std::string& bytes)
{
	if (bytes.size() < 4 || bytes[0] != 'P' || bytes[1] != 'K') {
		return false;
	}
	return (bytes[2] == '\x03' && bytes[3] == '\x04') ||
		(bytes[2] == '\x05' && bytes[3] == '\x06') ||
		(bytes[2] == '\x07' && bytes[3] == '\x08');
}

std::string DescribeDownloadFailure(const std::pair<std::string, int>& response)
{
	if (response.second == 200) {
		return std::format("响应内容不是有效的 ZIP 文件（{} 字节）", response.first.size());
	}

	std::string error = response.second == 0
		? std::string("未收到 HTTP 响应")
		: std::format("HTTP {}", response.second);
	if (!response.first.empty()) {
		error += ": " + response.first.substr(0, (std::min<size_t>)(response.first.size(), 300));
	}
	return error;
}

bool DownloadZip(const LatestReleaseInfo& info, const std::filesystem::path& zipPath, std::string& outError)
{
	const auto download = [&info](const std::string& url) {
		DownloadProgressReporter progress("[e-packager]", info.assetSize);
		return PerformGetRequest(
			url,
			kGitHubHeaders,
			300000,
			false,
			false,
			[&progress](std::uint64_t downloadedBytes, std::uint64_t totalBytes) {
				if (const auto message = progress.Update(downloadedBytes, totalBytes)) {
					OutputStringToELog(LocalFromUtf8(*message));
				}
			});
	};

	const std::string acceleratedUrl = BuildAcceleratedGitHubUrl(info.downloadUrl);
	OutputStringToELog(std::format("[e-packager] 通过 GitHub 加速地址下载：{}", acceleratedUrl));
	auto response = download(acceleratedUrl);
	if (response.second != 200 || !IsZipResponse(response.first)) {
		const std::string acceleratedError = DescribeDownloadFailure(response);
		OutputStringToELog("[e-packager] GitHub 加速地址下载失败，将尝试原始 GitHub 地址：" + acceleratedError);
		OutputStringToELog(std::format("[e-packager] 通过原始 GitHub 地址下载：{}", info.downloadUrl));
		response = download(info.downloadUrl);
		if (response.second != 200 || !IsZipResponse(response.first)) {
			outError = "GitHub 加速地址下载失败（" + acceleratedError + "）；原始 GitHub 地址下载失败（" +
				DescribeDownloadFailure(response) + "）";
			return false;
		}
	}

	if (!WriteBinaryFile(zipPath, response.first, outError)) {
		return false;
	}

	OutputStringToELog(std::format("[e-packager] 下载完成，大小 {} 字节", response.first.size()));
	return true;
}

bool ExtractZip(const std::filesystem::path& zipPath, const std::filesystem::path& destination, std::string& outError)
{
	OutputStringToELog("[e-packager] 正在解压工具包...");
	ArchiveExtractor::ExtractionResult result;
	if (!ArchiveExtractor::ExtractZip(
			zipPath,
			destination,
			GetToolsDirectory(),
			result,
			120)) {
		outError = result.error;
		return false;
	}
	OutputStringToELog(std::format(
		"[e-packager] 解压完成，extract_method={}",
		ArchiveExtractor::MethodName(result.method)));
	if (result.method == ArchiveExtractor::ExtractionMethod::Tar && !result.primaryError.empty()) {
		OutputStringToELog("[e-packager] PowerShell 解压失败，已使用 tar 后备：" + result.primaryError);
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

bool EnsureToolReadyImpl(
	std::filesystem::path& outToolPath,
	std::string& outError)
{
	const std::filesystem::path toolPath = GetEPackagerExePath();
	const bool toolExists = std::filesystem::exists(toolPath);
	const std::string installedVersion = toolExists ? DetectInstalledVersion(toolPath) : std::string();
	if (!IsUpdateCheckDue(toolExists)) {
		outToolPath = toolPath;
		PublishUpdateStatus(
			ComponentUpdateState::Completed,
			"e-packager 在最近 7 天内已检查，继续使用现有组件。",
			installedVersion.empty() ? "已安装" : installedVersion);
		return true;
	}

	PublishUpdateStatus(
		ComponentUpdateState::Checking,
		"正在检查 e-packager 最新版本...",
		installedVersion.empty() ? (toolExists ? "已安装" : "未安装") : installedVersion);

	LatestReleaseInfo latest;
	std::string fetchError;
	if (!FetchLatestRelease(latest, fetchError)) {
		if (toolExists) {
			OutputStringToELog("[e-packager] 7 天周期检查失败，将继续使用现有工具：" + fetchError);
			outToolPath = toolPath;
			PublishUpdateStatus(
				ComponentUpdateState::Error,
				"检查失败，继续使用已安装组件：" + fetchError,
				installedVersion.empty() ? "已安装" : installedVersion);
			return true;
		}
		outError = "无法下载 e-packager：" + fetchError + "\r\n" + BuildManualEPackagerInstallGuidance();
		PublishUpdateStatus(ComponentUpdateState::Error, outError);
		return false;
	}

	const bool needsDownload = IsToolUpdateRequired(toolExists, installedVersion, latest);

	if (!needsDownload) {
		OutputStringToELog("[e-packager] 本地工具已是最新版本：" + latest.tag);
		SaveMeta(latest);
		outToolPath = toolPath;
		PublishUpdateStatus(
			ComponentUpdateState::UpToDate,
			"本地组件已是最新版本。",
			installedVersion,
			latest.tag);
		return true;
	}

	PublishUpdateStatus(
		ComponentUpdateState::Downloading,
		"正在下载并安装 e-packager...",
		installedVersion.empty() ? (toolExists ? "已安装" : "未安装") : installedVersion,
		latest.tag);
	if (!DownloadAndInstallTool(latest, outError)) {
		outError += "\r\n" + BuildManualEPackagerInstallGuidance();
		if (toolExists) {
			OutputStringToELog("[e-packager] 自动更新失败，将继续使用现有工具：" + outError);
			PublishUpdateStatus(
				ComponentUpdateState::Error,
				"自动更新失败，继续使用已安装组件：" + outError,
				installedVersion,
				latest.tag);
			outError.clear();
			outToolPath = toolPath;
			return true;
		}
		PublishUpdateStatus(ComponentUpdateState::Error, outError, installedVersion, latest.tag);
		return false;
	}

	outToolPath = toolPath;
	PublishUpdateStatus(ComponentUpdateState::Completed, "e-packager 已更新完成。", latest.tag, latest.tag);
	return true;
}

struct ToolTaskRunningGuard {
	explicit ToolTaskRunningGuard(std::atomic_bool& runningFlag)
		: flag(runningFlag)
	{
	}

	~ToolTaskRunningGuard()
	{
		flag.store(false, std::memory_order_release);
	}

	std::atomic_bool& flag;
};

void ToolCheckWorker(void*)
{
	ToolTaskRunningGuard runningGuard(g_toolCheckTaskRunning);
	try {
		const std::filesystem::path toolPath = GetEPackagerExePath();
		const bool toolExists = std::filesystem::exists(toolPath);
		const std::string installedVersion = toolExists ? DetectInstalledVersion(toolPath) : std::string();
		const std::string displayedVersion = installedVersion.empty()
			? (toolExists ? "已安装" : "未安装")
			: installedVersion;
		ClearCheckedRelease();
		PublishUpdateStatus(
			ComponentUpdateState::Checking,
			"正在检查 e-packager 最新版本...",
			displayedVersion);

		LatestReleaseInfo latest;
		std::string error;
		if (!FetchLatestRelease(latest, error)) {
			const std::string guidance = toolExists ? std::string() : "\r\n" + BuildManualEPackagerInstallGuidance();
			PublishUpdateStatus(
				ComponentUpdateState::Error,
				"检查 e-packager 更新失败：" + error + guidance,
				displayedVersion);
			OutputStringToELog("[e-packager] 检查更新失败：" + error + guidance);
			return;
		}

		if (!IsToolUpdateRequired(toolExists, installedVersion, latest)) {
			SaveMeta(latest);
			PublishUpdateStatus(
				ComponentUpdateState::UpToDate,
				"本地组件已是最新版本。",
				displayedVersion,
				latest.tag);
			return;
		}

		StoreCheckedRelease(latest);
		OutputStringToELog(std::format(
			"[e-packager] 检测到新版本 {}，请通过关于页面或工具菜单手动更新。",
			latest.tag));
		PublishUpdateStatus(
			ComponentUpdateState::UpdateAvailable,
			toolExists ? "发现可用的新版本。" : "尚未安装，发现可用版本。",
			displayedVersion,
			latest.tag);
	}
	catch (const std::exception& ex) {
		ClearCheckedRelease();
		OutputStringToELog(std::string("[e-packager] 检查更新异常：") + ex.what());
		PublishUpdateStatus(ComponentUpdateState::Error, std::string("检查 e-packager 更新异常：") + ex.what());
	}
	catch (...) {
		ClearCheckedRelease();
		OutputStringToELog("[e-packager] 检查更新发生未知异常");
		PublishUpdateStatus(ComponentUpdateState::Error, "检查 e-packager 更新发生未知异常。");
	}
}

void ToolUpdateWorker(void* parameter)
{
	ToolTaskRunningGuard runningGuard(g_toolUpdateTaskRunning);
	std::unique_ptr<LatestReleaseInfo> latest(static_cast<LatestReleaseInfo*>(parameter));
	if (!latest) {
		PublishUpdateStatus(ComponentUpdateState::Error, "缺少已检查的 e-packager 版本信息。");
		return;
	}

	try {
		OutputStringToELog("[e-packager] 开始安装已检查的组件版本：" + latest->tag);
		const std::filesystem::path toolPath = GetEPackagerExePath();
		const bool toolExists = std::filesystem::exists(toolPath);
		const std::string installedVersion = toolExists ? DetectInstalledVersion(toolPath) : std::string();
		const std::string displayedVersion = installedVersion.empty()
			? (toolExists ? "已安装" : "未安装")
			: installedVersion;

		if (!IsToolUpdateRequired(toolExists, installedVersion, *latest)) {
			SaveMeta(*latest);
			ClearCheckedRelease();
			PublishUpdateStatus(
				ComponentUpdateState::UpToDate,
				"本地组件已是最新版本。",
				displayedVersion,
				latest->tag);
			return;
		}

		PublishUpdateStatus(
			ComponentUpdateState::Downloading,
			"正在下载并安装 e-packager...",
			displayedVersion,
			latest->tag);
		std::string error;
		if (!DownloadAndInstallTool(*latest, error)) {
			error += "\r\n" + BuildManualEPackagerInstallGuidance();
			PublishUpdateStatus(
				ComponentUpdateState::Error,
				"e-packager 更新失败：" + error,
				displayedVersion,
				latest->tag);
			OutputStringToELog("[e-packager] 手动更新失败：" + error);
			return;
		}

		ClearCheckedRelease();
		PublishUpdateStatus(
			ComponentUpdateState::Completed,
			"e-packager 已更新完成。",
			latest->tag,
			latest->tag);
		OutputStringToELog("[e-packager] 手动更新完成：" + LocalPathString(toolPath));
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[e-packager] 手动更新异常：") + ex.what());
		PublishUpdateStatus(ComponentUpdateState::Error, std::string("e-packager 更新异常：") + ex.what());
	}
	catch (...) {
		OutputStringToELog("[e-packager] 手动更新发生未知异常");
		PublishUpdateStatus(ComponentUpdateState::Error, "e-packager 更新发生未知异常。");
	}
}

// 工具菜单使用一键更新流程：在同一个后台任务中检查、下载并安装最新版本。
void ToolMenuUpdateWorker(void*)
{
	ToolTaskRunningGuard runningGuard(g_toolUpdateTaskRunning);
	try {
		const std::filesystem::path toolPath = GetEPackagerExePath();
		const bool toolExists = std::filesystem::exists(toolPath);
		const std::string installedVersion = toolExists ? DetectInstalledVersion(toolPath) : std::string();
		const std::string displayedVersion = installedVersion.empty()
			? (toolExists ? "已安装" : "未安装")
			: installedVersion;
		ClearCheckedRelease();
		PublishUpdateStatus(
			ComponentUpdateState::Checking,
			"正在检查 e-packager 最新版本...",
			displayedVersion);

		LatestReleaseInfo latest;
		std::string error;
		if (!FetchLatestRelease(latest, error)) {
			const std::string guidance = toolExists ? std::string() : "\r\n" + BuildManualEPackagerInstallGuidance();
			PublishUpdateStatus(
				ComponentUpdateState::Error,
				"检查 e-packager 更新失败：" + error + guidance,
				displayedVersion);
			OutputStringToELog("[e-packager] 一键更新检查失败：" + error + guidance);
			return;
		}

		if (!IsToolUpdateRequired(toolExists, installedVersion, latest)) {
			SaveMeta(latest);
			PublishUpdateStatus(
				ComponentUpdateState::UpToDate,
				"本地组件已是最新版本。",
				displayedVersion,
				latest.tag);
			OutputStringToELog("[e-packager] 本地工具已是最新版本：" + latest.tag);
			return;
		}

		PublishUpdateStatus(
			ComponentUpdateState::Downloading,
			"正在下载并安装 e-packager...",
			displayedVersion,
			latest.tag);
		if (!DownloadAndInstallTool(latest, error)) {
			error += "\r\n" + BuildManualEPackagerInstallGuidance();
			PublishUpdateStatus(
				ComponentUpdateState::Error,
				"e-packager 更新失败：" + error,
				displayedVersion,
				latest.tag);
			OutputStringToELog("[e-packager] 一键更新失败：" + error);
			return;
		}

		PublishUpdateStatus(
			ComponentUpdateState::Completed,
			"e-packager 已更新完成。",
			latest.tag,
			latest.tag);
		OutputStringToELog("[e-packager] 一键更新完成：" + LocalPathString(toolPath));
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[e-packager] 一键更新异常：") + ex.what());
		PublishUpdateStatus(ComponentUpdateState::Error, std::string("e-packager 更新异常：") + ex.what());
	}
	catch (...) {
		OutputStringToELog("[e-packager] 一键更新发生未知异常");
		PublishUpdateStatus(ComponentUpdateState::Error, "e-packager 更新发生未知异常。");
	}
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

void RunToolUpdateInBackground()
{
	LatestReleaseInfo latest;
	{
		std::lock_guard<std::mutex> taskLock(g_toolTaskStartMutex);
		if (g_toolCheckTaskRunning.load(std::memory_order_acquire) ||
			g_toolUpdateTaskRunning.load(std::memory_order_acquire)) {
			OutputStringToELog("[e-packager] 已有组件更新任务正在执行，请稍候");
			return;
		}
		if (!TryGetCheckedRelease(latest)) {
			PublishUpdateStatus(ComponentUpdateState::Error, "请先检查 e-packager 更新。");
			return;
		}
		g_toolUpdateTaskRunning.store(true, std::memory_order_release);
	}

	auto* request = new LatestReleaseInfo(std::move(latest));
	if (_beginthread(ToolUpdateWorker, 0, request) == static_cast<uintptr_t>(-1)) {
		delete request;
		g_toolUpdateTaskRunning.store(false, std::memory_order_release);
		PublishUpdateStatus(ComponentUpdateState::Error, "启动 e-packager 后台更新任务失败。");
		OutputStringToELog("[e-packager] 启动后台更新任务失败");
	}
}

void CheckForToolUpdatesInBackground()
{
	{
		std::lock_guard<std::mutex> taskLock(g_toolTaskStartMutex);
		if (g_toolCheckTaskRunning.load(std::memory_order_acquire) ||
			g_toolUpdateTaskRunning.load(std::memory_order_acquire)) {
			OutputStringToELog("[e-packager] 已有组件更新任务正在执行，请稍候");
			return;
		}
		g_toolCheckTaskRunning.store(true, std::memory_order_release);
	}

	if (_beginthread(ToolCheckWorker, 0, nullptr) == static_cast<uintptr_t>(-1)) {
		g_toolCheckTaskRunning.store(false, std::memory_order_release);
		PublishUpdateStatus(ComponentUpdateState::Error, "启动 e-packager 后台检查任务失败。");
		OutputStringToELog("[e-packager] 启动后台检查任务失败");
	}
}

void RunToolUpdateFromMenuInBackground()
{
	{
		std::lock_guard<std::mutex> taskLock(g_toolTaskStartMutex);
		if (g_toolCheckTaskRunning.load(std::memory_order_acquire) ||
			g_toolUpdateTaskRunning.load(std::memory_order_acquire)) {
			OutputStringToELog("[e-packager] 已有组件更新任务正在执行，请稍候");
			return;
		}
		g_toolUpdateTaskRunning.store(true, std::memory_order_release);
	}
	if (_beginthread(ToolMenuUpdateWorker, 0, nullptr) == static_cast<uintptr_t>(-1)) {
		g_toolUpdateTaskRunning.store(false, std::memory_order_release);
		PublishUpdateStatus(ComponentUpdateState::Error, "启动 e-packager 后台更新任务失败。");
		OutputStringToELog("[e-packager] 启动一键更新任务失败");
	}
}

void SetUpdateStatusNotificationWindow(HWND window)
{
	std::lock_guard<std::mutex> lock(g_updateStatusMutex);
	g_updateStatusNotificationWindow = window != nullptr && IsWindow(window) ? window : nullptr;
}

ComponentUpdateStatus GetUpdateStatus()
{
	std::lock_guard<std::mutex> lock(g_updateStatusMutex);
	return g_updateStatus;
}

ProcessRunResult RunProcessAndCapture(
	const std::filesystem::path& exePath,
	const std::vector<std::wstring>& args,
	const std::filesystem::path& workingDirectory,
	const unsigned long timeoutMs)
{
	return RunProcessAndCaptureImpl(exePath, args, workingDirectory, timeoutMs);
}

std::wstring BuildUnpackMenuTitle()
{
	const std::wstring filename = CurrentSourceFileNameW();
	if (filename.empty()) {
		return L"将当前.e反编译到目录";
	}
	return L"[" + filename + L"]反编译到目录";
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
