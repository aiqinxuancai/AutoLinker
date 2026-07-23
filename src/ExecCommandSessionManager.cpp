#include "ExecCommandSessionManager.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstdlib>
#include <cstdio>
#include <cwctype>
#include <filesystem>
#include <format>
#include <memory>
#include <mutex>
#include <random>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <utility>
#include <vector>

#include "..\thirdparty\json.hpp"

namespace {
constexpr size_t kOutputBufferBytes = 1024 * 1024;
constexpr size_t kOutputHalfBytes = kOutputBufferBytes / 2;
constexpr size_t kMaxOutputTokens = kOutputBufferBytes / 4;
constexpr size_t kMaxSessions = 64;
constexpr size_t kProtectedRecentSessions = 8;
constexpr unsigned int kMinYieldTimeMs = 250;
constexpr unsigned int kWindowsInitialYieldFloorMs = 2000;
constexpr unsigned int kMaxYieldTimeMs = 30000;
constexpr unsigned int kMinEmptyPollTimeMs = 5000;
constexpr unsigned int kMaxEmptyPollTimeMs = 300000;

std::wstring Utf8ToWide(const std::string& text)
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
	std::wstring wide(static_cast<size_t>(length), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		length) <= 0) {
		return {};
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
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
	std::string utf8(static_cast<size_t>(length), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		utf8.data(),
		length,
		nullptr,
		nullptr) <= 0) {
		return {};
	}
	return utf8;
}

std::string NormalizeCapturedText(const std::string& bytes)
{
	if (bytes.empty()) {
		return {};
	}
	const int wideLength = MultiByteToWideChar(
		CP_UTF8,
		0,
		bytes.data(),
		static_cast<int>(bytes.size()),
		nullptr,
		0);
	if (wideLength <= 0) {
		return bytes;
	}
	std::wstring wide(static_cast<size_t>(wideLength), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		0,
		bytes.data(),
		static_cast<int>(bytes.size()),
		wide.data(),
		wideLength) <= 0) {
		return bytes;
	}
	return WideToUtf8(wide);
}

std::wstring ToLowerWide(std::wstring value)
{
	std::transform(value.begin(), value.end(), value.begin(), [](wchar_t ch) {
		return static_cast<wchar_t>(towlower(ch));
	});
	return value;
}

bool IsAbsolutePath(const std::wstring& path)
{
	return (path.size() >= 2 && path[1] == L':') ||
		(path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\');
}

std::wstring ResolveExecutable(const std::wstring& shell)
{
	if (shell.empty()) {
		return {};
	}
	if (IsAbsolutePath(shell)) {
		const DWORD attrs = GetFileAttributesW(shell.c_str());
		return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0
			? shell
			: std::wstring();
	}
	const DWORD required = SearchPathW(nullptr, shell.c_str(), nullptr, 0, nullptr, nullptr);
	if (required == 0) {
		return {};
	}
	const size_t bufferLength = static_cast<size_t>(required) + 1;
	std::wstring resolved(bufferLength, L'\0');
	const DWORD copied = SearchPathW(
		nullptr,
		shell.c_str(),
		nullptr,
		static_cast<DWORD>(bufferLength),
		resolved.data(),
		nullptr);
	if (copied == 0 || copied >= bufferLength) {
		return {};
	}
	resolved.resize(static_cast<size_t>(copied));
	return resolved;
}

std::wstring QuoteCommandLineArgument(const std::wstring& argument)
{
	if (argument.empty()) {
		return L"\"\"";
	}
	if (argument.find_first_of(L" \t\n\v\"") == std::wstring::npos) {
		return argument;
	}

	std::wstring quoted = L"\"";
	size_t backslashes = 0;
	for (const wchar_t ch : argument) {
		if (ch == L'\\') {
			++backslashes;
			continue;
		}
		if (ch == L'\"') {
			quoted.append(backslashes * 2 + 1, L'\\');
			quoted.push_back(L'\"');
			backslashes = 0;
			continue;
		}
		quoted.append(backslashes, L'\\');
		backslashes = 0;
		quoted.push_back(ch);
	}
	quoted.append(backslashes * 2, L'\\');
	quoted.push_back(L'\"');
	return quoted;
}

std::wstring BuildCommandLine(const std::vector<std::wstring>& arguments)
{
	std::wstring commandLine;
	for (const std::wstring& argument : arguments) {
		if (!commandLine.empty()) {
			commandLine.push_back(L' ');
		}
		commandLine += QuoteCommandLineArgument(argument);
	}
	return commandLine;
}

bool ResolveShellArguments(
	const ExecCommandRequest& request,
	std::wstring& outExecutable,
	std::vector<std::wstring>& outArguments,
	std::string& outError)
{
	outError.clear();
	const std::wstring requestedShell = request.shellUtf8.empty()
		? L"powershell.exe"
		: Utf8ToWide(request.shellUtf8);
	if (requestedShell.empty()) {
		outError = "shell is not valid UTF-8";
		return false;
	}
	outExecutable = ResolveExecutable(requestedShell);
	if (outExecutable.empty()) {
		outError = "shell executable was not found";
		return false;
	}

	const std::wstring filename = ToLowerWide(std::filesystem::path(outExecutable).filename().wstring());
	const std::wstring command = Utf8ToWide(request.commandUtf8);
	if (command.empty()) {
		outError = "cmd is empty or not valid UTF-8";
		return false;
	}

	outArguments.clear();
	outArguments.push_back(outExecutable);
	if (filename == L"powershell.exe" || filename == L"powershell" ||
		filename == L"pwsh.exe" || filename == L"pwsh") {
		if (!request.loginShell) {
			outArguments.push_back(L"-NoProfile");
		}
		outArguments.push_back(L"-Command");
		outArguments.push_back(
			L"[Console]::InputEncoding=[System.Text.UTF8Encoding]::new($false);"
			L"[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);"
			L"$OutputEncoding=[System.Text.UTF8Encoding]::new($false);" + command);
		return true;
	}
	if (filename == L"cmd.exe" || filename == L"cmd") {
		outArguments.push_back(L"/c");
		outArguments.push_back(command);
		return true;
	}
	if (filename == L"bash.exe" || filename == L"bash" ||
		filename == L"sh.exe" || filename == L"sh" ||
		filename == L"zsh.exe" || filename == L"zsh") {
		outArguments.push_back(request.loginShell ? L"-lc" : L"-c");
		outArguments.push_back(command);
		return true;
	}
	outError = "unsupported shell; expected PowerShell, cmd, bash, sh, or zsh";
	return false;
}

std::vector<wchar_t> BuildEnvironmentBlock()
{
	std::vector<std::wstring> entries;
	LPWCH environment = GetEnvironmentStringsW();
	if (environment != nullptr) {
		for (const wchar_t* current = environment; *current != L'\0'; current += wcslen(current) + 1) {
			entries.emplace_back(current);
		}
		FreeEnvironmentStringsW(environment);
	}

	const std::array<std::pair<std::wstring, std::wstring>, 10> overrides = {{
		{L"NO_COLOR", L"1"},
		{L"TERM", L"dumb"},
		{L"LANG", L"C.UTF-8"},
		{L"LC_CTYPE", L"C.UTF-8"},
		{L"LC_ALL", L"C.UTF-8"},
		{L"COLORTERM", L""},
		{L"PAGER", L"cat"},
		{L"GIT_PAGER", L"cat"},
		{L"GH_PAGER", L"cat"},
		{L"CODEX_CI", L"1"}
	}};
	for (const auto& [name, value] : overrides) {
		const auto it = std::find_if(entries.begin(), entries.end(), [&name](const std::wstring& entry) {
			const size_t equals = entry.find(L'=');
			return equals != std::wstring::npos && _wcsicmp(entry.substr(0, equals).c_str(), name.c_str()) == 0;
		});
		const std::wstring replacement = name + L"=" + value;
		if (it == entries.end()) {
			entries.push_back(replacement);
		}
		else {
			*it = replacement;
		}
	}
	std::sort(entries.begin(), entries.end(), [](const std::wstring& left, const std::wstring& right) {
		return _wcsicmp(left.c_str(), right.c_str()) < 0;
	});

	size_t total = 1;
	for (const std::wstring& entry : entries) {
		total += entry.size() + 1;
	}
	std::vector<wchar_t> block;
	block.reserve(total);
	for (const std::wstring& entry : entries) {
		block.insert(block.end(), entry.begin(), entry.end());
		block.push_back(L'\0');
	}
	block.push_back(L'\0');
	return block;
}

unsigned int ClampInitialYield(unsigned int value)
{
	return (std::clamp)((std::max)(value, kWindowsInitialYieldFloorMs), kMinYieldTimeMs, kMaxYieldTimeMs);
}

unsigned int ClampWriteYield(unsigned int value, bool empty)
{
	const unsigned int atLeastMinimum = (std::max)(value, kMinYieldTimeMs);
	return empty
		? (std::clamp)(atLeastMinimum, kMinEmptyPollTimeMs, kMaxEmptyPollTimeMs)
		: (std::min)(atLeastMinimum, kMaxYieldTimeMs);
}

size_t ClampMaxTokens(size_t value)
{
	return (std::min)(value == 0 ? size_t{0} : value, kMaxOutputTokens);
}

std::string TruncateForTokens(const std::string& text, size_t maxTokens)
{
	const size_t maxBytes = (std::min)(maxTokens, kMaxOutputTokens) * 4;
	if (text.size() <= maxBytes) {
		return text;
	}
	if (maxBytes == 0) {
		return {};
	}
	if (maxBytes < 96) {
		static constexpr std::string_view kShortTruncationMarker = "... output truncated ...";
		return std::string(kShortTruncationMarker.substr(0, maxBytes));
	}
	const size_t markerReserve = 96;
	const size_t contentBytes = maxBytes - markerReserve;
	const size_t headBytes = contentBytes / 2;
	const size_t tailBytes = contentBytes - headBytes;
	const size_t omitted = text.size() - headBytes - tailBytes;
	std::string result = text.substr(0, headBytes);
	result += std::format("\n... {} bytes truncated ...\n", omitted);
	if (tailBytes > 0) {
		result += text.substr(text.size() - tailBytes);
	}
	return NormalizeCapturedText(result);
}

std::string GenerateChunkId()
{
	static std::mutex randomMutex;
	static std::mt19937 generator(std::random_device{}());
	static constexpr char kHex[] = "0123456789abcdef";
	std::lock_guard<std::mutex> guard(randomMutex);
	std::uniform_int_distribution<int> distribution(0, 15);
	std::string id(6, '0');
	for (char& ch : id) {
		ch = kHex[distribution(generator)];
	}
	return id;
}

struct DrainedOutput {
	std::string bytes;
	size_t originalBytes = 0;
};

struct Session : std::enable_shared_from_this<Session> {
	~Session()
	{
		Terminate(125);
		if (watcher.joinable()) {
			watcher.join();
		}
		if (reader.joinable()) {
			reader.join();
		}
		if (readPipe != nullptr) {
			CloseHandle(readPipe);
		}
		if (process != nullptr) {
			CloseHandle(process);
		}
		if (job != nullptr) {
			CloseHandle(job);
		}
	}

	void AppendOutput(const char* data, size_t size)
	{
		std::lock_guard<std::mutex> guard(mutex);
		unreadOriginalBytes += size;
		if (!unreadTruncated && unreadHead.size() + size <= kOutputBufferBytes) {
			unreadHead.append(data, size);
		}
		else if (!unreadTruncated) {
			std::string combined = std::move(unreadHead);
			combined.append(data, size);
			unreadHead = combined.substr(0, kOutputHalfBytes);
			unreadTail = combined.substr(combined.size() - kOutputHalfBytes);
			unreadTruncated = true;
		}
		else {
			unreadTail.append(data, size);
			if (unreadTail.size() > kOutputHalfBytes) {
				unreadTail.erase(0, unreadTail.size() - kOutputHalfBytes);
			}
		}
		cv.notify_all();
	}

	DrainedOutput DrainOutput()
	{
		std::lock_guard<std::mutex> guard(mutex);
		DrainedOutput drained;
		drained.originalBytes = unreadOriginalBytes;
		drained.bytes = std::move(unreadHead);
		if (unreadTruncated) {
			const size_t retained = drained.bytes.size() + unreadTail.size();
			const size_t omitted = unreadOriginalBytes > retained ? unreadOriginalBytes - retained : 0;
			drained.bytes += std::format("\n... {} bytes truncated ...\n", omitted);
			drained.bytes += unreadTail;
		}
		unreadHead.clear();
		unreadTail.clear();
		unreadOriginalBytes = 0;
		unreadTruncated = false;
		return drained;
	}

	void Terminate(DWORD exitCode)
	{
		bool expected = false;
		if (!terminationRequested.compare_exchange_strong(expected, true)) {
			return;
		}
		if (job != nullptr) {
			TerminateJobObject(job, exitCode);
		}
	}

	std::string ownerSessionId;
	int sessionId = 0;
	HANDLE job = nullptr;
	HANDLE process = nullptr;
	HANDLE readPipe = nullptr;
	std::thread reader;
	std::thread watcher;
	std::mutex mutex;
	std::condition_variable cv;
	std::string unreadHead;
	std::string unreadTail;
	size_t unreadOriginalBytes = 0;
	bool unreadTruncated = false;
	bool outputClosed = false;
	bool processExited = false;
	DWORD exitCode = STILL_ACTIVE;
	std::atomic_bool terminationRequested = false;
	std::chrono::steady_clock::time_point lastUsed = std::chrono::steady_clock::now();
};

void ReadSessionOutput(const std::shared_ptr<Session>& session)
{
	std::array<char, 8192> buffer = {};
	DWORD bytesRead = 0;
	while (ReadFile(
		session->readPipe,
		buffer.data(),
		static_cast<DWORD>(buffer.size()),
		&bytesRead,
		nullptr) != FALSE && bytesRead > 0) {
		session->AppendOutput(buffer.data(), bytesRead);
	}
	{
		std::lock_guard<std::mutex> guard(session->mutex);
		session->outputClosed = true;
	}
	session->cv.notify_all();
}

void WatchSessionProcess(const std::shared_ptr<Session>& session)
{
	WaitForSingleObject(session->process, INFINITE);
	DWORD exitCode = 0;
	GetExitCodeProcess(session->process, &exitCode);
	if (session->job != nullptr) {
		TerminateJobObject(session->job, exitCode);
	}
	{
		std::lock_guard<std::mutex> guard(session->mutex);
		session->processExited = true;
		session->exitCode = exitCode;
	}
	session->cv.notify_all();
}

bool StartSessionProcess(
	const ExecCommandRequest& request,
	const std::shared_ptr<Session>& session,
	std::string& outError)
{
	std::wstring executable;
	std::vector<std::wstring> arguments;
	if (!ResolveShellArguments(request, executable, arguments, outError)) {
		return false;
	}

	std::wstring workingDirectory = Utf8ToWide(request.workingDirectoryUtf8);
	if (workingDirectory.empty()) {
		outError = "workdir is empty or not valid UTF-8";
		return false;
	}
	if (!IsAbsolutePath(workingDirectory)) {
		outError = "workdir must resolve to an absolute path";
		return false;
	}
	const DWORD directoryAttrs = GetFileAttributesW(workingDirectory.c_str());
	if (directoryAttrs == INVALID_FILE_ATTRIBUTES || (directoryAttrs & FILE_ATTRIBUTE_DIRECTORY) == 0) {
		outError = "workdir does not exist";
		return false;
	}

	SECURITY_ATTRIBUTES security = {};
	security.nLength = sizeof(security);
	security.bInheritHandle = TRUE;
	HANDLE readPipe = nullptr;
	HANDLE writePipe = nullptr;
	if (CreatePipe(&readPipe, &writePipe, &security, 0) == FALSE) {
		outError = "CreatePipe output failed";
		return false;
	}
	if (SetHandleInformation(readPipe, HANDLE_FLAG_INHERIT, 0) == FALSE) {
		CloseHandle(readPipe);
		CloseHandle(writePipe);
		outError = "make output read pipe non-inheritable failed";
		return false;
	}
	HANDLE nullInput = CreateFileW(
		L"NUL",
		GENERIC_READ,
		FILE_SHARE_READ | FILE_SHARE_WRITE,
		&security,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,
		nullptr);
	if (nullInput == INVALID_HANDLE_VALUE) {
		CloseHandle(readPipe);
		CloseHandle(writePipe);
		outError = "open NUL stdin failed";
		return false;
	}

	STARTUPINFOEXW startup = {};
	startup.StartupInfo.cb = sizeof(startup);
	startup.StartupInfo.dwFlags = STARTF_USESTDHANDLES;
	startup.StartupInfo.hStdInput = nullInput;
	startup.StartupInfo.hStdOutput = writePipe;
	startup.StartupInfo.hStdError = writePipe;
	SIZE_T attributeBytes = 0;
	InitializeProcThreadAttributeList(nullptr, 1, 0, &attributeBytes);
	if (attributeBytes == 0) {
		CloseHandle(nullInput);
		CloseHandle(readPipe);
		CloseHandle(writePipe);
		outError = "query process attribute list size failed";
		return false;
	}
	std::vector<unsigned char> attributeStorage(attributeBytes);
	startup.lpAttributeList = reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(attributeStorage.data());
	if (InitializeProcThreadAttributeList(startup.lpAttributeList, 1, 0, &attributeBytes) == FALSE) {
		CloseHandle(nullInput);
		CloseHandle(readPipe);
		CloseHandle(writePipe);
		outError = "InitializeProcThreadAttributeList failed";
		return false;
	}
	std::array<HANDLE, 2> inheritedHandles = {writePipe, nullInput};
	if (UpdateProcThreadAttribute(
		startup.lpAttributeList,
		0,
		PROC_THREAD_ATTRIBUTE_HANDLE_LIST,
		inheritedHandles.data(),
		sizeof(inheritedHandles),
		nullptr,
		nullptr) == FALSE) {
		DeleteProcThreadAttributeList(startup.lpAttributeList);
		CloseHandle(nullInput);
		CloseHandle(readPipe);
		CloseHandle(writePipe);
		outError = "UpdateProcThreadAttribute handle list failed";
		return false;
	}

	std::wstring commandLine = BuildCommandLine(arguments);
	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');
	std::vector<wchar_t> environment = BuildEnvironmentBlock();
	PROCESS_INFORMATION processInfo = {};
	const BOOL created = CreateProcessW(
		executable.c_str(),
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		TRUE,
		CREATE_NO_WINDOW | CREATE_SUSPENDED | EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT,
		environment.data(),
		workingDirectory.c_str(),
		&startup.StartupInfo,
		&processInfo);
	const DWORD createError = created == FALSE ? GetLastError() : ERROR_SUCCESS;
	DeleteProcThreadAttributeList(startup.lpAttributeList);
	CloseHandle(nullInput);
	CloseHandle(writePipe);
	if (created == FALSE) {
		CloseHandle(readPipe);
		outError = std::format("CreateProcessW failed, error={}", createError);
		return false;
	}

	HANDLE job = CreateJobObjectW(nullptr, nullptr);
	if (job == nullptr) {
		TerminateProcess(processInfo.hProcess, 126);
		CloseHandle(processInfo.hThread);
		CloseHandle(processInfo.hProcess);
		CloseHandle(readPipe);
		outError = "CreateJobObjectW failed";
		return false;
	}
	JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobInfo = {};
	jobInfo.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
	if (SetInformationJobObject(
		job,
		JobObjectExtendedLimitInformation,
		&jobInfo,
		static_cast<DWORD>(sizeof(jobInfo))) == FALSE ||
		AssignProcessToJobObject(job, processInfo.hProcess) == FALSE) {
		TerminateProcess(processInfo.hProcess, 126);
		CloseHandle(job);
		CloseHandle(processInfo.hThread);
		CloseHandle(processInfo.hProcess);
		CloseHandle(readPipe);
		outError = "assign command process to job failed";
		return false;
	}

	session->job = job;
	session->process = processInfo.hProcess;
	session->readPipe = readPipe;
	if (ResumeThread(processInfo.hThread) == static_cast<DWORD>(-1)) {
		CloseHandle(processInfo.hThread);
		session->Terminate(126);
		outError = "ResumeThread command process failed";
		return false;
	}
	CloseHandle(processInfo.hThread);
	try {
		session->reader = std::thread(ReadSessionOutput, session);
		session->watcher = std::thread(WatchSessionProcess, session);
	}
	catch (const std::system_error& ex) {
		session->Terminate(126);
		if (session->watcher.joinable()) {
			session->watcher.join();
		}
		if (session->reader.joinable()) {
			session->reader.join();
		}
		outError = std::string("start command monitor thread failed: ") + ex.what();
		return false;
	}
	return true;
}

struct WaitSnapshot {
	DrainedOutput output;
	bool exited = false;
	DWORD exitCode = STILL_ACTIVE;
};

WaitSnapshot WaitAndCollect(
	const std::shared_ptr<Session>& session,
	unsigned int yieldTimeMs,
	const std::function<bool()>& cancelCallback)
{
	const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(yieldTimeMs);
	std::unique_lock<std::mutex> lock(session->mutex);
	while (!(session->outputClosed && session->processExited) &&
		std::chrono::steady_clock::now() < deadline) {
		if (cancelCallback && cancelCallback()) {
			break;
		}
		const auto remaining = deadline - std::chrono::steady_clock::now();
		const auto slice = (std::min)(
			remaining,
			std::chrono::duration_cast<std::chrono::steady_clock::duration>(
				std::chrono::milliseconds(50)));
		session->cv.wait_for(lock, slice);
	}
	WaitSnapshot snapshot;
	snapshot.exited = session->processExited;
	snapshot.exitCode = session->exitCode;
	lock.unlock();
	snapshot.output = session->DrainOutput();
	return snapshot;
}

std::string FormatResponse(
	const WaitSnapshot& snapshot,
	int sessionId,
	size_t maxOutputTokens,
	double wallTimeSeconds)
{
	const std::string normalized = NormalizeCapturedText(snapshot.output.bytes);
	const size_t originalTokenCount = (snapshot.output.originalBytes + 3) / 4;
	std::vector<std::string> sections;
	sections.push_back("Chunk ID: " + GenerateChunkId());
	sections.push_back(std::format("Wall time: {:.4f} seconds", wallTimeSeconds));
	if (snapshot.exited) {
		sections.push_back(std::format("Process exited with code {}", snapshot.exitCode));
	}
	else {
		sections.push_back(std::format("Process running with session ID {}", sessionId));
	}
	sections.push_back(std::format("Original token count: {}", originalTokenCount));
	sections.push_back("Output:");
	sections.push_back(TruncateForTokens(normalized, ClampMaxTokens(maxOutputTokens)));

	std::string response;
	for (const std::string& section : sections) {
		if (!response.empty()) {
			response.push_back('\n');
		}
		response += section;
	}
	return response;
}
} // namespace

class ExecCommandSessionManager::Impl {
public:
	int AllocateSessionIdLocked()
	{
		std::uniform_int_distribution<int> distribution(1000, 99999);
		for (;;) {
			const int candidate = distribution(generator);
			if (!sessions.contains(candidate)) {
				return candidate;
			}
		}
	}

	void Store(const std::shared_ptr<Session>& session)
	{
		std::shared_ptr<Session> evicted;
		{
			std::lock_guard<std::mutex> guard(mutex);
			session->sessionId = AllocateSessionIdLocked();
			if (sessions.size() >= kMaxSessions) {
				std::vector<std::shared_ptr<Session>> ordered;
				ordered.reserve(sessions.size());
				for (const auto& [id, item] : sessions) {
					(void)id;
					ordered.push_back(item);
				}
				std::sort(ordered.begin(), ordered.end(), [](const auto& left, const auto& right) {
					return left->lastUsed > right->lastUsed;
				});
				const size_t protectedCount = (std::min)(kProtectedRecentSessions, ordered.size());
				auto choose = ordered.end();
				for (auto it = ordered.begin() + static_cast<std::ptrdiff_t>(protectedCount); it != ordered.end(); ++it) {
					std::lock_guard<std::mutex> sessionGuard((*it)->mutex);
					if ((*it)->processExited) {
						choose = it;
						break;
					}
				}
				if (choose == ordered.end() && protectedCount < ordered.size()) {
					choose = ordered.end() - 1;
				}
				if (choose != ordered.end()) {
					evicted = *choose;
					sessions.erase(evicted->sessionId);
				}
			}
			sessions[session->sessionId] = session;
		}
		if (evicted) {
			evicted->Terminate(125);
		}
	}

	std::shared_ptr<Session> Find(int sessionId, const std::string& ownerSessionId)
	{
		std::lock_guard<std::mutex> guard(mutex);
		const auto it = sessions.find(sessionId);
		if (it == sessions.end() || it->second->ownerSessionId != ownerSessionId) {
			return {};
		}
		it->second->lastUsed = std::chrono::steady_clock::now();
		return it->second;
	}

	void RemoveIfSame(int sessionId, const std::shared_ptr<Session>& session)
	{
		std::lock_guard<std::mutex> guard(mutex);
		const auto it = sessions.find(sessionId);
		if (it != sessions.end() && it->second == session) {
			sessions.erase(it);
		}
	}

	std::mutex mutex;
	std::unordered_map<int, std::shared_ptr<Session>> sessions;
	std::mt19937 generator{std::random_device{}()};
};

ExecCommandSessionManager& ExecCommandSessionManager::Instance()
{
	static ExecCommandSessionManager instance;
	return instance;
}

ExecCommandSessionManager::ExecCommandSessionManager()
	: m_impl(new Impl())
{
}

ExecCommandSessionManager::~ExecCommandSessionManager()
{
	TerminateAll();
	delete m_impl;
	m_impl = nullptr;
}

ExecCommandResult ExecCommandSessionManager::Execute(
	const ExecCommandRequest& request,
	const std::function<bool()>& cancelCallback)
{
	ExecCommandResult result;
	if (request.ownerSessionId.empty()) {
		result.error = "internal chat session is unavailable";
		return result;
	}
	if (request.tty) {
		result.error = "tty=true is not supported by AutoLinker exec_command";
		return result;
	}
	if (request.commandUtf8.empty()) {
		result.error = "cmd is required";
		return result;
	}

	auto session = std::make_shared<Session>();
	session->ownerSessionId = request.ownerSessionId;
	if (!StartSessionProcess(request, session, result.error)) {
		return result;
	}
	m_impl->Store(session);

	const auto start = std::chrono::steady_clock::now();
	const WaitSnapshot snapshot = WaitAndCollect(session, ClampInitialYield(request.yieldTimeMs), cancelCallback);
	const double wallTime = std::chrono::duration<double>(std::chrono::steady_clock::now() - start).count();
	result.responseUtf8 = FormatResponse(snapshot, session->sessionId, request.maxOutputTokens, wallTime);
	result.ok = true;
	if (snapshot.exited) {
		m_impl->RemoveIfSame(session->sessionId, session);
	}
	return result;
}

ExecCommandResult ExecCommandSessionManager::WriteStdin(
	const ExecCommandWriteRequest& request,
	const std::function<bool()>& cancelCallback)
{
	ExecCommandResult result;
	const std::shared_ptr<Session> session = m_impl->Find(request.sessionId, request.ownerSessionId);
	if (!session) {
		result.error = std::format("write_stdin failed: unknown session ID {}", request.sessionId);
		return result;
	}
	if (!request.charsUtf8.empty() && request.charsUtf8 != std::string(1, '\x03')) {
		result.error = "write_stdin failed: stdin is closed for a non-TTY session";
		return result;
	}
	if (request.charsUtf8 == std::string(1, '\x03')) {
		session->Terminate(130);
	}

	const auto start = std::chrono::steady_clock::now();
	const WaitSnapshot snapshot = WaitAndCollect(
		session,
		ClampWriteYield(request.yieldTimeMs, request.charsUtf8.empty()),
		cancelCallback);
	const double wallTime = std::chrono::duration<double>(std::chrono::steady_clock::now() - start).count();
	result.responseUtf8 = FormatResponse(snapshot, session->sessionId, request.maxOutputTokens, wallTime);
	result.ok = true;
	if (snapshot.exited) {
		m_impl->RemoveIfSame(session->sessionId, session);
	}
	return result;
}

void ExecCommandSessionManager::TerminateOwnerSession(const std::string& ownerSessionId)
{
	if (ownerSessionId.empty()) {
		return;
	}
	std::vector<std::shared_ptr<Session>> removed;
	{
		std::lock_guard<std::mutex> guard(m_impl->mutex);
		for (auto it = m_impl->sessions.begin(); it != m_impl->sessions.end();) {
			if (it->second->ownerSessionId == ownerSessionId) {
				removed.push_back(it->second);
				it = m_impl->sessions.erase(it);
			}
			else {
				++it;
			}
		}
	}
	for (const auto& session : removed) {
		session->Terminate(125);
	}
}

void ExecCommandSessionManager::TerminateAll()
{
	if (m_impl == nullptr) {
		return;
	}
	std::vector<std::shared_ptr<Session>> removed;
	{
		std::lock_guard<std::mutex> guard(m_impl->mutex);
		for (auto& [id, session] : m_impl->sessions) {
			(void)id;
			removed.push_back(std::move(session));
		}
		m_impl->sessions.clear();
	}
	for (const auto& session : removed) {
		session->Terminate(125);
	}
}

std::string ExecCommandSessionManager::BuildSelfTestJson()
{
	nlohmann::json report = {
		{"name", "exec-command-session"},
		{"ok", false}
	};
	const std::string owner = "exec-command-self-test";

	ExecCommandRequest shortRequest;
	shortRequest.ownerSessionId = owner;
	shortRequest.commandUtf8 = "Write-Output 'short-ok'; exit 7";
	shortRequest.workingDirectoryUtf8 = WideToUtf8(std::filesystem::current_path().wstring());
	shortRequest.loginShell = false;
	shortRequest.yieldTimeMs = 2000;
	const ExecCommandResult shortResult = Execute(shortRequest);
	const bool shortOk = shortResult.ok &&
		shortResult.responseUtf8.find("Process exited with code 7") != std::string::npos &&
		shortResult.responseUtf8.find("short-ok") != std::string::npos;

	ExecCommandRequest longRequest = shortRequest;
	longRequest.commandUtf8 = "Write-Output 'begin'; Start-Sleep -Milliseconds 2600; Write-Output 'end'";
	const ExecCommandResult first = Execute(longRequest);
	int sessionId = 0;
	if (first.ok) {
		const std::string marker = "Process running with session ID ";
		const size_t markerAt = first.responseUtf8.find(marker);
		if (markerAt != std::string::npos) {
			sessionId = std::atoi(first.responseUtf8.c_str() + markerAt + marker.size());
		}
	}
	ExecCommandWriteRequest poll;
	poll.ownerSessionId = owner;
	poll.sessionId = sessionId;
	poll.yieldTimeMs = 5000;
	const ExecCommandResult second = sessionId > 0 ? WriteStdin(poll) : ExecCommandResult{};
	const bool sessionOk = first.ok && sessionId > 0 && second.ok &&
		second.responseUtf8.find("Process exited with code 0") != std::string::npos &&
		second.responseUtf8.find("end") != std::string::npos;

	ExecCommandRequest cancelRequest = shortRequest;
	cancelRequest.commandUtf8 = "Start-Sleep -Seconds 30";
	const ExecCommandResult running = Execute(cancelRequest, []() { return true; });
	int cancelSessionId = 0;
	if (running.ok) {
		const std::string marker = "Process running with session ID ";
		const size_t markerAt = running.responseUtf8.find(marker);
		if (markerAt != std::string::npos) {
			cancelSessionId = std::atoi(running.responseUtf8.c_str() + markerAt + marker.size());
		}
	}
	ExecCommandWriteRequest interrupt;
	interrupt.ownerSessionId = owner;
	interrupt.sessionId = cancelSessionId;
	interrupt.charsUtf8 = std::string(1, '\x03');
	interrupt.yieldTimeMs = 250;
	const ExecCommandResult interrupted = cancelSessionId > 0 ? WriteStdin(interrupt) : ExecCommandResult{};
	const bool interruptOk = interrupted.ok &&
		interrupted.responseUtf8.find("Process exited with code 130") != std::string::npos;

	TerminateOwnerSession(owner);
	report["short_command"] = {{"ok", shortOk}};
	report["background_poll"] = {{"ok", sessionOk}, {"session_id", sessionId}};
	report["interrupt"] = {{"ok", interruptOk}, {"session_id", cancelSessionId}};
	report["ok"] = shortOk && sessionOk && interruptOk;
	return report.dump();
}
