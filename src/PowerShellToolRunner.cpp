#include "PowerShellToolRunner.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <string>
#include <thread>
#include <vector>

namespace {
std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	const int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int utf8Len = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (utf8Len <= 0) {
		return std::string();
	}

	std::string utf8(static_cast<size_t>(utf8Len), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		utf8.data(),
		utf8Len,
		nullptr,
		nullptr) <= 0) {
		return std::string();
	}
	return utf8;
}

std::string Base64Encode(const unsigned char* data, size_t size)
{
	static constexpr char kTable[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	std::string out;
	out.reserve(((size + 2) / 3) * 4);
	for (size_t i = 0; i < size; i += 3) {
		const unsigned int a = data[i];
		const unsigned int b = (i + 1 < size) ? data[i + 1] : 0;
		const unsigned int c = (i + 2 < size) ? data[i + 2] : 0;
		const unsigned int n = (a << 16) | (b << 8) | c;

		out.push_back(kTable[(n >> 18) & 0x3F]);
		out.push_back(kTable[(n >> 12) & 0x3F]);
		out.push_back(i + 1 < size ? kTable[(n >> 6) & 0x3F] : '=');
		out.push_back(i + 2 < size ? kTable[n & 0x3F] : '=');
	}
	return out;
}

std::string BuildEncodedCommand(const std::wstring& userCommand)
{
	std::wstring script =
		L"[Console]::InputEncoding=[System.Text.Encoding]::UTF8\r\n"
		L"[Console]::OutputEncoding=[System.Text.Encoding]::UTF8\r\n"
		L"$OutputEncoding=[System.Text.Encoding]::UTF8\r\n";
	script += userCommand;

	const auto* bytes = reinterpret_cast<const unsigned char*>(script.data());
	return Base64Encode(bytes, script.size() * sizeof(wchar_t));
}

void ReadPipeToString(HANDLE hPipe, std::string* output, std::atomic_bool* doneFlag)
{
	if (output == nullptr || doneFlag == nullptr) {
		return;
	}

	char buffer[4096];
	DWORD bytesRead = 0;
	while (ReadFile(hPipe, buffer, static_cast<DWORD>(sizeof(buffer)), &bytesRead, nullptr) != FALSE && bytesRead > 0) {
		output->append(buffer, bytesRead);
	}
	*doneFlag = true;
}

bool IsAbsoluteOrEmpty(const std::wstring& path)
{
	if (path.empty()) {
		return true;
	}
	if (path.size() >= 2 && path[1] == L':') {
		return true;
	}
	if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\') {
		return true;
	}
	return false;
}

std::wstring GetCurrentDirectoryWide()
{
	const DWORD len = GetCurrentDirectoryW(0, nullptr);
	if (len == 0) {
		return std::wstring();
	}
	std::wstring dir(static_cast<size_t>(len), L'\0');
	if (GetCurrentDirectoryW(len, dir.data()) == 0) {
		return std::wstring();
	}
	if (!dir.empty() && dir.back() == L'\0') {
		dir.pop_back();
	}
	return dir;
}
} // namespace

PowerShellRunResult PowerShellToolRunner::Run(
	const std::string& commandUtf8,
	const std::string& workingDirectoryUtf8,
	int timeoutSeconds)
{
	PowerShellRunResult result = {};

	const std::wstring command = Utf8ToWide(commandUtf8);
	if (command.empty()) {
		result.error = "command is empty or not valid UTF-8";
		return result;
	}

	std::wstring workingDirectory = Utf8ToWide(workingDirectoryUtf8);
	if (workingDirectory.empty()) {
		workingDirectory = GetCurrentDirectoryWide();
	}
	else if (!IsAbsoluteOrEmpty(workingDirectory)) {
		result.error = "working_directory must be an absolute path";
		return result;
	}

	DWORD attrs = GetFileAttributesW(workingDirectory.c_str());
	if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0) {
		result.error = "working_directory does not exist";
		return result;
	}

	result.effectiveWorkingDirectory = WideToUtf8(workingDirectory);
	const int boundedTimeoutSeconds = (std::clamp)(timeoutSeconds, 1, 600);
	const std::string encodedCommand = BuildEncodedCommand(command);
	std::wstring commandLine = L"powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ";
	commandLine += Utf8ToWide(encodedCommand);

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
	si.dwFlags = STARTF_USESTDHANDLES;
	si.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
	si.hStdOutput = stdOutWrite;
	si.hStdError = stdErrWrite;

	PROCESS_INFORMATION pi = {};
	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');
	const BOOL created = CreateProcessW(
		nullptr,
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		TRUE,
		CREATE_NO_WINDOW,
		nullptr,
		workingDirectory.c_str(),
		&si,
		&pi);

	CloseHandle(stdOutWrite);
	CloseHandle(stdErrWrite);

	if (created == FALSE) {
		CloseHandle(stdOutRead);
		CloseHandle(stdErrRead);
		result.error = "CreateProcessW powershell.exe failed";
		return result;
	}

	std::atomic_bool stdoutDone = false;
	std::atomic_bool stderrDone = false;
	std::thread stdoutReader(ReadPipeToString, stdOutRead, &result.stdOut, &stdoutDone);
	std::thread stderrReader(ReadPipeToString, stdErrRead, &result.stdErr, &stderrDone);

	const DWORD waitResult = WaitForSingleObject(pi.hProcess, static_cast<DWORD>(boundedTimeoutSeconds * 1000));
	if (waitResult == WAIT_TIMEOUT) {
		result.timedOut = true;
		TerminateProcess(pi.hProcess, 124);
		WaitForSingleObject(pi.hProcess, 5000);
	}

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

	result.ok = !result.timedOut && result.error.empty();
	return result;
}
