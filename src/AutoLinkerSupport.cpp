#include "AutoLinkerInternal.h"

#include <Windows.h>

#include <algorithm>
#include <filesystem>
#include <format>
#include <fstream>
#include <mutex>
#include <string>

#include "PathHelper.h"

namespace {

std::mutex g_aiRoundtripLogMutex;

std::filesystem::path GetAIRoundtripLogPath()
{
	return GetAutoLinkerLogFilePath("ai_roundtrip_last.log");
}

bool IsValidUtf8ForLog(const std::string& text)
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

std::string BuildHexHead(const std::string& text, size_t maxBytes = 64)
{
	const size_t n = (std::min)(text.size(), maxBytes);
	std::string out;
	out.reserve(n * 3 + 16);
	constexpr char kHex[] = "0123456789ABCDEF";
	for (size_t i = 0; i < n; ++i) {
		const unsigned char ch = static_cast<unsigned char>(text[i]);
		out.push_back(kHex[(ch >> 4) & 0x0F]);
		out.push_back(kHex[ch & 0x0F]);
		if (i + 1 < n) {
			out.push_back(' ');
		}
	}
	if (text.size() > n) {
		out += " ...";
	}
	return out;
}

void AppendAIRoundtripLogLineUnlocked(std::ofstream& out, const std::string& line)
{
	out << line << "\r\n";
}

bool TryParseProcessIdDirectoryName(const std::wstring& directoryName, DWORD& processId)
{
	if (directoryName.empty()) {
		return false;
	}

	ULONGLONG parsedValue = 0;
	for (const wchar_t ch : directoryName) {
		if (ch < L'0' || ch > L'9') {
			return false;
		}
		parsedValue = parsedValue * 10 + static_cast<ULONGLONG>(ch - L'0');
		if (parsedValue > MAXDWORD) {
			return false;
		}
	}

	processId = static_cast<DWORD>(parsedValue);
	return true;
}

bool IsProcessLikelyAlive(DWORD processId)
{
	if (processId == 0) {
		return false;
	}

	HANDLE process = OpenProcess(SYNCHRONIZE, FALSE, processId);
	if (process == nullptr) {
		// 只有明确的无效进程号才视为进程已退出；权限不足等情况保守保留目录。
		return GetLastError() != ERROR_INVALID_PARAMETER;
	}

	const DWORD waitResult = WaitForSingleObject(process, 0);
	CloseHandle(process);
	if (waitResult == WAIT_OBJECT_0) {
		return false;
	}
	return true;
}

void RemoveExitedWebView2ProcessDirectories(
	const std::filesystem::path& rootPath,
	DWORD currentProcessId)
{
	std::error_code iterateError;
	std::filesystem::directory_iterator iterator(rootPath, iterateError);
	const std::filesystem::directory_iterator end;
	while (!iterateError && iterator != end) {
		const std::filesystem::directory_entry entry = *iterator;
		iterator.increment(iterateError);

		std::error_code typeError;
		if (!entry.is_directory(typeError) || typeError) {
			continue;
		}

		DWORD processId = 0;
		if (!TryParseProcessIdDirectoryName(entry.path().filename().wstring(), processId) ||
			processId == currentProcessId ||
			IsProcessLikelyAlive(processId)) {
			continue;
		}

		std::error_code removeError;
		std::filesystem::remove_all(entry.path(), removeError);
	}
}

std::wstring InitializeWebView2UserDataFolderPath()
{
	wchar_t tempPathBuffer[MAX_PATH] = {};
	const DWORD tempPathLength = GetTempPathW(
		static_cast<DWORD>(sizeof(tempPathBuffer) / sizeof(tempPathBuffer[0])),
		tempPathBuffer);

	std::filesystem::path tempPath;
	if (tempPathLength > 0 &&
		tempPathLength < static_cast<DWORD>(sizeof(tempPathBuffer) / sizeof(tempPathBuffer[0]))) {
		tempPath = tempPathBuffer;
	}
	else {
		std::error_code fallbackError;
		tempPath = std::filesystem::temp_directory_path(fallbackError);
		if (fallbackError) {
			return L"";
		}
	}

	const std::filesystem::path rootPath = tempPath / "AutoLinker" / "WebView2";
	std::error_code createRootError;
	std::filesystem::create_directories(rootPath, createRootError);

	const DWORD currentProcessId = GetCurrentProcessId();
	RemoveExitedWebView2ProcessDirectories(rootPath, currentProcessId);

	const std::filesystem::path processPath = rootPath / std::to_wstring(currentProcessId);
	std::error_code createProcessError;
	std::filesystem::create_directories(processPath, createProcessError);
	return processPath.wstring();
}

} // namespace

std::wstring GetWebView2UserDataFolderPath()
{
	static const std::wstring processPath = InitializeWebView2UserDataFolderPath();
	return processPath;
}

std::string EscapeOneLineForLog(std::string text)
{
	for (size_t i = 0; i < text.size(); ++i) {
		if (text[i] == '\r') {
			text.replace(i, 1, "\\r");
			++i;
			continue;
		}
		if (text[i] == '\n') {
			text.replace(i, 1, "\\n");
			++i;
			continue;
		}
	}
	return text;
}

void BeginAIRoundtripLogSession(uint64_t traceId, const std::string& scene, const std::string& taskName)
{
	const auto path = GetAIRoundtripLogPath();
	SYSTEMTIME st = {};
	GetLocalTime(&st);
	std::lock_guard<std::mutex> guard(g_aiRoundtripLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		return;
	}

	AppendAIRoundtripLogLineUnlocked(
		out,
		"[AI-ROUNDTRIP] session-start trace=" + std::to_string(traceId) +
		" scene=" + scene +
		" task=\"" + EscapeOneLineForLog(taskName) + "\"" +
		" time=" + std::to_string(st.wYear) + "-" + std::to_string(st.wMonth) + "-" + std::to_string(st.wDay) +
		" " + std::to_string(st.wHour) + ":" + std::to_string(st.wMinute) + ":" + std::to_string(st.wSecond) +
		"." + std::to_string(st.wMilliseconds));
}

void AppendAIRoundtripLogLine(const std::string& line)
{
	const auto path = GetAIRoundtripLogPath();
	std::lock_guard<std::mutex> guard(g_aiRoundtripLogMutex);
	std::ofstream out(path, std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}
	AppendAIRoundtripLogLineUnlocked(out, line);
}

void AppendAIRoundtripLogBlock(const std::string& title, const std::string& text)
{
	const auto path = GetAIRoundtripLogPath();
	std::lock_guard<std::mutex> guard(g_aiRoundtripLogMutex);
	std::ofstream out(path, std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}

	AppendAIRoundtripLogLineUnlocked(
		out,
		"[BLOCK-BEGIN] title=" + title +
		" bytes=" + std::to_string(text.size()) +
		" utf8Valid=" + std::to_string(IsValidUtf8ForLog(text) ? 1 : 0) +
		" hexHead=" + BuildHexHead(text));
	out << text << "\r\n";
	AppendAIRoundtripLogLineUnlocked(out, "[BLOCK-END] title=" + title);
}
