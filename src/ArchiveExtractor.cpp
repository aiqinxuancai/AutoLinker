#include "ArchiveExtractor.h"

#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <format>
#include <sstream>
#include <string>
#include <thread>
#include <vector>

#include "PowerShellToolRunner.h"

namespace ArchiveExtractor {
namespace {

struct ProcessResult {
	bool ok = false;
	bool timedOut = false;
	DWORD exitCode = 0;
	std::string stdOut;
	std::string stdErr;
	std::string error;
};

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

std::wstring QuoteCommandLineArgument(const std::wstring& argument)
{
	if (argument.empty()) {
		return L"\"\"";
	}
	bool needsQuote = false;
	for (const wchar_t ch : argument) {
		if (ch == L' ' || ch == L'\t' || ch == L'\n' || ch == L'\v' || ch == L'\"') {
			needsQuote = true;
			break;
		}
	}
	if (!needsQuote) {
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
			quoted.push_back(ch);
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

std::filesystem::path GetWindowsDirectoryPath()
{
	std::vector<wchar_t> buffer(MAX_PATH);
	for (;;) {
		const UINT length = GetWindowsDirectoryW(buffer.data(), static_cast<UINT>(buffer.size()));
		if (length == 0) {
			return {};
		}
		if (length < buffer.size()) {
			return std::filesystem::path(std::wstring(buffer.data(), length));
		}
		buffer.resize(static_cast<size_t>(length) + 1);
	}
}

std::filesystem::path GetSystemDirectoryPath()
{
	std::vector<wchar_t> buffer(MAX_PATH);
	for (;;) {
		const UINT length = GetSystemDirectoryW(buffer.data(), static_cast<UINT>(buffer.size()));
		if (length == 0) {
			return {};
		}
		if (length < buffer.size()) {
			return std::filesystem::path(std::wstring(buffer.data(), length));
		}
		buffer.resize(static_cast<size_t>(length) + 1);
	}
}

void AddCandidate(std::vector<std::filesystem::path>& candidates, std::filesystem::path candidate)
{
	if (candidate.empty()) {
		return;
	}
	const auto duplicate = std::find_if(candidates.begin(), candidates.end(), [&candidate](const auto& existing) {
		return _wcsicmp(existing.c_str(), candidate.c_str()) == 0;
	});
	if (duplicate == candidates.end()) {
		candidates.push_back(std::move(candidate));
	}
}

std::filesystem::path ResolveSystemExecutable(const wchar_t* executableName)
{
	std::vector<std::filesystem::path> candidates;
	const std::filesystem::path systemDirectory = GetSystemDirectoryPath();
	const std::filesystem::path windowsDirectory = GetWindowsDirectoryPath();
	AddCandidate(candidates, systemDirectory / executableName);
	AddCandidate(candidates, windowsDirectory / L"System32" / executableName);

	BOOL wow64 = FALSE;
	if (IsWow64Process(GetCurrentProcess(), &wow64) != FALSE && wow64 != FALSE) {
		AddCandidate(candidates, windowsDirectory / L"Sysnative" / executableName);
		AddCandidate(candidates, windowsDirectory / L"SysWOW64" / executableName);
	}

	for (const auto& candidate : candidates) {
		std::error_code ec;
		if (std::filesystem::is_regular_file(candidate, ec) && !ec) {
			return candidate;
		}
	}
	return {};
}

void ReadPipeToString(HANDLE pipe, std::string* output)
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

ProcessResult RunProcess(
	const std::filesystem::path& executable,
	const std::vector<std::wstring>& arguments,
	const std::filesystem::path& workingDirectory,
	int timeoutSeconds)
{
	ProcessResult result;
	if (executable.empty()) {
		result.error = "executable was not found";
		return result;
	}

	std::wstring commandLine = QuoteCommandLineArgument(executable.wstring());
	for (const auto& argument : arguments) {
		commandLine.push_back(L' ');
		commandLine += QuoteCommandLineArgument(argument);
	}

	SECURITY_ATTRIBUTES security = {};
	security.nLength = sizeof(security);
	security.bInheritHandle = TRUE;
	HANDLE stdOutRead = nullptr;
	HANDLE stdOutWrite = nullptr;
	HANDLE stdErrRead = nullptr;
	HANDLE stdErrWrite = nullptr;
	if (CreatePipe(&stdOutRead, &stdOutWrite, &security, 0) == FALSE) {
		result.error = std::format("CreatePipe stdout failed, Win32 error={}", GetLastError());
		return result;
	}
	if (CreatePipe(&stdErrRead, &stdErrWrite, &security, 0) == FALSE) {
		const DWORD error = GetLastError();
		CloseHandle(stdOutRead);
		CloseHandle(stdOutWrite);
		result.error = std::format("CreatePipe stderr failed, Win32 error={}", error);
		return result;
	}
	SetHandleInformation(stdOutRead, HANDLE_FLAG_INHERIT, 0);
	SetHandleInformation(stdErrRead, HANDLE_FLAG_INHERIT, 0);

	STARTUPINFOW startup = {};
	startup.cb = sizeof(startup);
	startup.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
	startup.wShowWindow = SW_HIDE;
	startup.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
	startup.hStdOutput = stdOutWrite;
	startup.hStdError = stdErrWrite;

	PROCESS_INFORMATION process = {};
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
		workingDirectory.empty() ? nullptr : workingDirectory.c_str(),
		&startup,
		&process);
	const DWORD createError = created == FALSE ? GetLastError() : ERROR_SUCCESS;
	CloseHandle(stdOutWrite);
	CloseHandle(stdErrWrite);
	if (created == FALSE) {
		CloseHandle(stdOutRead);
		CloseHandle(stdErrRead);
		result.error = std::format("CreateProcessW failed, Win32 error={}", createError);
		return result;
	}

	std::thread stdoutReader(ReadPipeToString, stdOutRead, &result.stdOut);
	std::thread stderrReader(ReadPipeToString, stdErrRead, &result.stdErr);
	const DWORD waitMilliseconds = static_cast<DWORD>((std::clamp)(timeoutSeconds, 1, 600)) * 1000U;
	const DWORD waitResult = WaitForSingleObject(process.hProcess, waitMilliseconds);
	if (waitResult == WAIT_TIMEOUT) {
		result.timedOut = true;
		TerminateProcess(process.hProcess, 124);
		WaitForSingleObject(process.hProcess, 5000);
	}
	else if (waitResult != WAIT_OBJECT_0) {
		result.error = std::format("WaitForSingleObject failed, Win32 error={}", GetLastError());
		TerminateProcess(process.hProcess, 126);
		WaitForSingleObject(process.hProcess, 5000);
	}

	DWORD exitCode = 0;
	if (GetExitCodeProcess(process.hProcess, &exitCode) != FALSE) {
		result.exitCode = exitCode;
	}
	CloseHandle(process.hThread);
	CloseHandle(process.hProcess);
	if (stdoutReader.joinable()) {
		stdoutReader.join();
	}
	if (stderrReader.joinable()) {
		stderrReader.join();
	}
	CloseHandle(stdOutRead);
	CloseHandle(stdErrRead);

	if (result.timedOut) {
		result.error = "process timed out";
	}
	else if (result.error.empty() && result.exitCode != 0) {
		result.error = std::format("process exited with code {}", result.exitCode);
	}
	result.ok = result.error.empty() && result.exitCode == 0;
	return result;
}

std::string DescribeProcessFailure(const ProcessResult& result)
{
	std::string description = result.error;
	if (description.empty()) {
		description = std::format("exitCode={}", result.exitCode);
	}
	if (!result.stdErr.empty()) {
		description += " stderr=" + result.stdErr;
	}
	return description;
}

bool IsUnsafeArchiveEntry(std::string entry)
{
	while (!entry.empty() && (entry.back() == '\r' || entry.back() == '\n')) {
		entry.pop_back();
	}
	std::replace(entry.begin(), entry.end(), '\\', '/');
	while (entry.rfind("./", 0) == 0) {
		entry.erase(0, 2);
	}
	if (entry.empty()) {
		return false;
	}
	if (entry.front() == '/' ||
		(entry.size() >= 2 && std::isalpha(static_cast<unsigned char>(entry[0])) != 0 && entry[1] == ':')) {
		return true;
	}

	size_t begin = 0;
	while (begin <= entry.size()) {
		const size_t end = entry.find('/', begin);
		const std::string component = entry.substr(begin, end == std::string::npos ? std::string::npos : end - begin);
		if (component == "..") {
			return true;
		}
		if (end == std::string::npos) {
			break;
		}
		begin = end + 1;
	}
	return false;
}

bool ValidateTarEntries(const std::string& listing, std::string& outError)
{
	std::istringstream stream(listing);
	std::string line;
	size_t entryCount = 0;
	while (std::getline(stream, line)) {
		if (line.empty()) {
			continue;
		}
		++entryCount;
		if (entryCount > 100000) {
			outError = "archive contains too many entries";
			return false;
		}
		if (IsUnsafeArchiveEntry(line)) {
			outError = "archive contains unsafe entry: " + line;
			return false;
		}
	}
	if (entryCount == 0) {
		outError = "archive contains no entries";
		return false;
	}
	return true;
}

bool TryPowerShellExtract(
	const std::filesystem::path& archivePath,
	const std::filesystem::path& destination,
	const std::filesystem::path& workingDirectory,
	int timeoutSeconds,
	std::string& outError)
{
	const std::string command =
		"Expand-Archive -LiteralPath " + PowerShellLiteral(archivePath) +
		" -DestinationPath " + PowerShellLiteral(destination) + " -Force";
	const PowerShellRunResult result = PowerShellToolRunner::Run(
		command,
		Utf8FromWide(workingDirectory.wstring()),
		timeoutSeconds);
	if (result.ok && result.exitCode == 0) {
		return true;
	}
	outError = result.error.empty()
		? std::format("Expand-Archive exitCode={}", result.exitCode)
		: result.error;
	if (!result.stdErr.empty()) {
		outError += " stderr=" + result.stdErr;
	}
	return false;
}

bool TryTarExtract(
	const std::filesystem::path& archivePath,
	const std::filesystem::path& destination,
	const std::filesystem::path& workingDirectory,
	int timeoutSeconds,
	std::string& outError)
{
	const std::filesystem::path tarPath = ResolveSystemExecutable(L"tar.exe");
	if (tarPath.empty()) {
		outError = "system tar.exe was not found";
		return false;
	}

	const ProcessResult listResult = RunProcess(
		tarPath,
		{ L"-tf", archivePath.wstring() },
		workingDirectory,
		timeoutSeconds);
	if (!listResult.ok) {
		outError = "tar archive listing failed: " + DescribeProcessFailure(listResult);
		return false;
	}
	if (!ValidateTarEntries(listResult.stdOut, outError)) {
		return false;
	}

	const ProcessResult extractResult = RunProcess(
		tarPath,
		{ L"-xf", archivePath.wstring(), L"-C", destination.wstring() },
		workingDirectory,
		timeoutSeconds);
	if (!extractResult.ok) {
		outError = "tar extraction failed: " + DescribeProcessFailure(extractResult);
		return false;
	}
	return true;
}

} // namespace

bool ExtractZip(
	const std::filesystem::path& archivePath,
	const std::filesystem::path& destination,
	const std::filesystem::path& workingDirectory,
	ExtractionResult& outResult,
	int timeoutSeconds)
{
	outResult = {};
	std::error_code ec;
	if (!std::filesystem::is_regular_file(archivePath, ec) || ec) {
		outResult.error = "archive file is unavailable: " + Utf8FromWide(archivePath.wstring());
		return false;
	}
	std::filesystem::create_directories(destination, ec);
	if (ec) {
		outResult.error = "create extraction directory failed: " + ec.message();
		return false;
	}
	const std::filesystem::path effectiveWorkingDirectory = workingDirectory.empty()
		? destination.parent_path()
		: workingDirectory;

	if (TryPowerShellExtract(
			archivePath,
			destination,
			effectiveWorkingDirectory,
			timeoutSeconds,
			outResult.primaryError)) {
		outResult.ok = true;
		outResult.method = ExtractionMethod::PowerShell;
		return true;
	}

	if (TryTarExtract(
			archivePath,
			destination,
			effectiveWorkingDirectory,
			timeoutSeconds,
			outResult.fallbackError)) {
		outResult.ok = true;
		outResult.method = ExtractionMethod::Tar;
		return true;
	}

	outResult.error =
		"PowerShell extraction failed: " + outResult.primaryError +
		"; tar fallback failed: " + outResult.fallbackError;
	return false;
}

const char* MethodName(ExtractionMethod method)
{
	switch (method) {
	case ExtractionMethod::PowerShell:
		return "powershell";
	case ExtractionMethod::Tar:
		return "tar";
	default:
		return "none";
	}
}

} // namespace ArchiveExtractor
