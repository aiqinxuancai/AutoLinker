#include "IDEFacade.h"
#include "Global.h"
#include "PathHelper.h"

#include <fnshare.h>
#include <lib2.h>
#include <algorithm>
#include <cctype>
#include <chrono>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <limits>
#include <mutex>
#include <utility>
#include <vector>

namespace {
using PerfClock = std::chrono::steady_clock;

long long ElapsedMs(const PerfClock::time_point& start)
{
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(PerfClock::now() - start).count());
}

std::string TruncateForPerfLog(const std::string& text, size_t maxLen = 120)
{
	if (text.size() <= maxLen) {
		return text;
	}
	return text.substr(0, maxLen) + "...";
}

std::string TrimAsciiCopyLocal(const std::string& text)
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

std::string GetWindowTextCopyLocalA(HWND hWnd)
{
	char buffer[512] = {};
	if (hWnd != nullptr && IsWindow(hWnd)) {
		GetWindowTextA(hWnd, buffer, static_cast<int>(sizeof(buffer)));
	}
	return buffer;
}

BOOL CALLBACK EnumChildProcCollectClass(HWND hWnd, LPARAM lParam)
{
	auto* pair = reinterpret_cast<std::pair<const char*, std::vector<HWND>*>*>(lParam);
	if (pair == nullptr || pair->first == nullptr || pair->second == nullptr) {
		return TRUE;
	}

	char className[128] = {};
	if (GetClassNameA(hWnd, className, static_cast<int>(sizeof(className))) > 0 &&
		_stricmp(className, pair->first) == 0) {
		pair->second->push_back(hWnd);
	}
	return TRUE;
}

std::vector<HWND> CollectChildWindowsByClassLocal(HWND root, const char* className)
{
	std::vector<HWND> windows;
	if (root == nullptr || !IsWindow(root) || className == nullptr || className[0] == '\0') {
		return windows;
	}

	std::pair<const char*, std::vector<HWND>*> ctx{ className, &windows };
	EnumChildWindows(root, EnumChildProcCollectClass, reinterpret_cast<LPARAM>(&ctx));
	return windows;
}

std::string DescribeActiveWindowTypeShort(IDEFacade::ActiveWindowType type)
{
	switch (type)
	{
	case IDEFacade::ActiveWindowType::Module:
		return "程序集/类";
	case IDEFacade::ActiveWindowType::UserDataType:
		return "自定义数据类型";
	case IDEFacade::ActiveWindowType::GlobalVar:
		return "全局变量";
	case IDEFacade::ActiveWindowType::DllCommand:
		return "DLL命令";
	case IDEFacade::ActiveWindowType::FormDesigner:
		return "窗口/表单";
	case IDEFacade::ActiveWindowType::ConstResource:
		return "常量资源";
	case IDEFacade::ActiveWindowType::PictureResource:
		return "图片资源";
	case IDEFacade::ActiveWindowType::SoundResource:
		return "声音资源";
	default:
		return std::string();
	}
}

bool TrySplitPageIdentityText(
	const std::string& rawText,
	std::string& outTypeText,
	std::string& outName)
{
	outTypeText.clear();
	outName.clear();

	const std::string text = TrimAsciiCopyLocal(rawText);
	if (text.empty()) {
		return false;
	}

	size_t colonPos = text.find(":");
	size_t colonWidth = 1;
	if (colonPos == std::string::npos) {
		colonPos = text.find("：");
		colonWidth = 3;
	}

	if (colonPos != std::string::npos) {
		std::string left = TrimAsciiCopyLocal(text.substr(0, colonPos));
		std::string right = TrimAsciiCopyLocal(text.substr(colonPos + colonWidth));
		if (!left.empty() && !right.empty()) {
			outTypeText = std::move(left);
			outName = std::move(right);
			return true;
		}
	}

	outName = text;
	return true;
}

bool TryExtractBracketPageIdentity(
	const std::string& mainTitle,
	std::string& outTypeText,
	std::string& outName)
{
	outTypeText.clear();
	outName.clear();

	const size_t closePos = mainTitle.rfind(']');
	if (closePos == std::string::npos) {
		return false;
	}
	const size_t openPos = mainTitle.rfind('[', closePos);
	if (openPos == std::string::npos || closePos <= openPos + 1) {
		return false;
	}
	return TrySplitPageIdentityText(mainTitle.substr(openPos + 1, closePos - openPos - 1), outTypeText, outName);
}

std::mutex g_codeFetchLogMutex;

std::filesystem::path GetCodeFetchLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "ai_code_fetch_last.log";
}

std::string EscapeOneLine(std::string text)
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

void AppendCodeFetchLogLine(const std::string& line)
{
	if (!IsAICodeFetchDebugEnabled()) {
		return;
	}
	const auto path = GetCodeFetchLogPath();
	std::lock_guard<std::mutex> guard(g_codeFetchLogMutex);
	std::ofstream out(path, std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}
	out << line << "\r\n";
}

void BeginCodeFetchLogSession(const char* scene)
{
	if (!IsAICodeFetchDebugEnabled()) {
		return;
	}
	const auto path = GetCodeFetchLogPath();
	const uint64_t traceId = GetCurrentAIPerfTraceId();
	SYSTEMTIME st = {};
	GetLocalTime(&st);
	std::lock_guard<std::mutex> guard(g_codeFetchLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		return;
	}
	out << "[AI-CODE-FETCH] session-start"
		<< " scene=" << (scene == nullptr ? "" : scene)
		<< " trace=" << traceId
		<< " time="
		<< st.wYear << "-"
		<< st.wMonth << "-"
		<< st.wDay << " "
		<< st.wHour << ":"
		<< st.wMinute << ":"
		<< st.wSecond
		<< "." << st.wMilliseconds
		<< "\r\n";
}
DWORD PtrToDWORD(const void* ptr)
{
	return static_cast<DWORD>(reinterpret_cast<ULONG_PTR>(ptr));
}

bool IsSubRelatedType(int type)
{
	switch (type)
	{
	case VT_SUB_NAME:
	case VT_SUB_RET_TYPE:
	case VT_SUB_EPK_NAME:
	case VT_SUB_EXPLAIN:
	case VT_SUB_EXPORT:
	case VT_SUB_ARG_NAME:
	case VT_SUB_ARG_TYPE:
	case VT_SUB_ARG_POINTER_TYPE:
	case VT_SUB_ARG_NULL_TYPE:
	case VT_SUB_ARG_ARY_TYPE:
	case VT_SUB_ARG_EXPLAIN:
	case VT_SUB_VAR_NAME:
	case VT_SUB_VAR_TYPE:
	case VT_SUB_VAR_STATIC_TYPE:
	case VT_SUB_VAR_ARY_TYPE:
	case VT_SUB_VAR_EXPLAIN:
	case VT_SUB_PRG_ITEM:
		return true;
	default:
		return false;
	}
}

void TrimTrailingLineBreaks(std::string& text)
{
	while (!text.empty() && (text.back() == '\r' || text.back() == '\n')) {
		text.pop_back();
	}
}

void TrimTrailingLineBreaks(std::wstring& text)
{
	while (!text.empty() && (text.back() == L'\r' || text.back() == L'\n')) {
		text.pop_back();
	}
}

void RemoveLastLine(std::string& text)
{
	TrimTrailingLineBreaks(text);
	size_t pos = text.find_last_of('\n');
	if (pos == std::string::npos) {
		text.clear();
		return;
	}
	size_t keep = pos;
	if (keep > 0 && text[keep - 1] == '\r') {
		--keep;
	}
	text.resize(keep);
}

void RemoveLastLine(std::wstring& text)
{
	TrimTrailingLineBreaks(text);
	size_t pos = text.find_last_of(L'\n');
	if (pos == std::wstring::npos) {
		text.clear();
		return;
	}
	size_t keep = pos;
	if (keep > 0 && text[keep - 1] == L'\r') {
		--keep;
	}
	text.resize(keep);
}

bool SetClipboardUnicodeText(const std::wstring& text)
{
	const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
	if (hMem == nullptr) {
		return false;
	}

	void* memPtr = GlobalLock(hMem);
	if (memPtr == nullptr) {
		GlobalFree(hMem);
		return false;
	}
	std::memcpy(memPtr, text.c_str(), bytes);
	GlobalUnlock(hMem);
	if (SetClipboardData(CF_UNICODETEXT, hMem) == nullptr) {
		GlobalFree(hMem);
		return false;
	}
	return true;
}

bool SetClipboardAnsiText(const std::string& text)
{
	const size_t bytes = text.size() + 1;
	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
	if (hMem == nullptr) {
		return false;
	}

	void* memPtr = GlobalLock(hMem);
	if (memPtr == nullptr) {
		GlobalFree(hMem);
		return false;
	}
	std::memcpy(memPtr, text.c_str(), bytes);
	GlobalUnlock(hMem);
	if (SetClipboardData(CF_TEXT, hMem) == nullptr) {
		GlobalFree(hMem);
		return false;
	}
	return true;
}

std::string WideToUtf8(const std::wstring& text);

std::wstring MultiByteSmartToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, nullptr, 0);
	UINT codePage = CP_UTF8;
	DWORD flags = MB_ERR_INVALID_CHARS;
	if (size <= 0) {
		size = MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, nullptr, 0);
		codePage = CP_ACP;
		flags = 0;
		if (size <= 0) {
			return std::wstring();
		}
	}

	std::wstring out(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(codePage, flags, text.c_str(), -1, out.data(), size) <= 0) {
		return std::wstring();
	}
	if (!out.empty() && out.back() == L'\0') {
		out.pop_back();
	}
	return out;
}

std::string ConvertUtf8ToCodePage(const std::string& text, UINT toCodePage)
{
	if (text.empty()) {
		return std::string();
	}

	// Only convert when input is strict UTF-8. Otherwise keep original bytes unchanged.
	int wideLen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, nullptr, 0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, wide.data(), wideLen) <= 0) {
		return text;
	}

	int outLen = WideCharToMultiByte(toCodePage, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(toCodePage, 0, wide.c_str(), -1, out.data(), outLen, nullptr, nullptr) <= 0) {
		return text;
	}
	if (!out.empty() && out.back() == '\0') {
		out.pop_back();
	}
	return out;
}

std::string EnsureGbkText(const std::string& text)
{
	// Keep non-UTF8 input untouched to avoid corrupting existing GBK/ANSI source bytes.
	return ConvertUtf8ToCodePage(text, 936);
}

bool SetClipboardTextForPaste(const std::string& text)
{
	if (!OpenClipboard(nullptr)) {
		return false;
	}

	bool ok = false;
	if (EmptyClipboard()) {
		const std::wstring wide = MultiByteSmartToWide(text);
		const std::string ansiText = EnsureGbkText(text);
		bool hasAnyFormat = false;
		if (!wide.empty()) {
			hasAnyFormat = SetClipboardUnicodeText(wide) || hasAnyFormat;
		}
		// Always provide ANSI(CF_TEXT) as GBK payload for hosts preferring ANSI paste.
		if (!ansiText.empty()) {
			hasAnyFormat = SetClipboardAnsiText(ansiText) || hasAnyFormat;
		}
		else {
			hasAnyFormat = SetClipboardAnsiText(text) || hasAnyFormat;
		}
		ok = hasAnyFormat;
	}

	CloseClipboard();
	return ok;
}

struct ClipboardTextSnapshot {
	bool captured = false;
	bool hasText = false;
	std::string textUtf8;
};

bool CaptureClipboardTextSnapshot(ClipboardTextSnapshot& outSnapshot)
{
	outSnapshot = {};
	if (!OpenClipboard(nullptr)) {
		return false;
	}

	outSnapshot.captured = true;
	if (IsClipboardFormatAvailable(CF_UNICODETEXT)) {
		HANDLE hData = GetClipboardData(CF_UNICODETEXT);
		if (hData != nullptr) {
			const wchar_t* textPtr = static_cast<const wchar_t*>(GlobalLock(hData));
			if (textPtr != nullptr) {
				outSnapshot.textUtf8 = WideToUtf8(std::wstring(textPtr));
				outSnapshot.hasText = true;
				GlobalUnlock(hData);
			}
		}
	}
	else if (IsClipboardFormatAvailable(CF_TEXT)) {
		HANDLE hData = GetClipboardData(CF_TEXT);
		if (hData != nullptr) {
			const char* textPtr = static_cast<const char*>(GlobalLock(hData));
			if (textPtr != nullptr) {
				outSnapshot.textUtf8.assign(textPtr);
				outSnapshot.hasText = true;
				GlobalUnlock(hData);
			}
		}
	}

	CloseClipboard();
	return true;
}

void RestoreClipboardTextSnapshot(const ClipboardTextSnapshot& snapshot)
{
	if (!snapshot.captured || !snapshot.hasText) {
		return;
	}
	SetClipboardTextForPaste(snapshot.textUtf8);
}

class ClipboardTextRestoreGuard {
public:
	ClipboardTextRestoreGuard()
	{
		m_enabled = CaptureClipboardTextSnapshot(m_snapshot);
	}

	~ClipboardTextRestoreGuard()
	{
		if (m_enabled) {
			RestoreClipboardTextSnapshot(m_snapshot);
		}
	}

	ClipboardTextRestoreGuard(const ClipboardTextRestoreGuard&) = delete;
	ClipboardTextRestoreGuard& operator=(const ClipboardTextRestoreGuard&) = delete;

private:
	ClipboardTextSnapshot m_snapshot = {};
	bool m_enabled = false;
};

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	int size = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return std::string();
	}

	std::string out(static_cast<size_t>(size), '\0');
	WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, out.data(), size, nullptr, nullptr);
	if (!out.empty() && out.back() == '\0') {
		out.pop_back();
	}
	return out;
}

bool ReadClipboardText(std::string& outText)
{
	outText.clear();
	if (!OpenClipboard(nullptr)) {
		return false;
	}

	bool ok = false;
	if (IsClipboardFormatAvailable(CF_UNICODETEXT)) {
		HANDLE hData = GetClipboardData(CF_UNICODETEXT);
		if (hData != nullptr) {
			const wchar_t* textPtr = static_cast<const wchar_t*>(GlobalLock(hData));
			if (textPtr != nullptr) {
				outText = WideToUtf8(std::wstring(textPtr));
				GlobalUnlock(hData);
				ok = true;
			}
		}
	}
	else if (IsClipboardFormatAvailable(CF_TEXT)) {
		HANDLE hData = GetClipboardData(CF_TEXT);
		if (hData != nullptr) {
			const char* textPtr = static_cast<const char*>(GlobalLock(hData));
			if (textPtr != nullptr) {
				outText.assign(textPtr);
				GlobalUnlock(hData);
				ok = true;
			}
		}
	}

	CloseClipboard();
	return ok;
}

bool TrimClipboardLastLine()
{
	if (!OpenClipboard(nullptr)) {
		return false;
	}

	bool ok = false;
	if (IsClipboardFormatAvailable(CF_UNICODETEXT)) {
		HANDLE hData = GetClipboardData(CF_UNICODETEXT);
		if (hData != nullptr) {
			const wchar_t* textPtr = static_cast<const wchar_t*>(GlobalLock(hData));
			if (textPtr != nullptr) {
				std::wstring text(textPtr);
				GlobalUnlock(hData);
				RemoveLastLine(text);
				EmptyClipboard();
				ok = SetClipboardUnicodeText(text);
			}
		}
	}
	else if (IsClipboardFormatAvailable(CF_TEXT)) {
		HANDLE hData = GetClipboardData(CF_TEXT);
		if (hData != nullptr) {
			const char* textPtr = static_cast<const char*>(GlobalLock(hData));
			if (textPtr != nullptr) {
				std::string text(textPtr);
				GlobalUnlock(hData);
				RemoveLastLine(text);
				EmptyClipboard();
				ok = SetClipboardAnsiText(text);
			}
		}
	}

	CloseClipboard();
	return ok;
}

#ifdef UNICODE
std::wstring ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	int size = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
	if (size <= 0) {
		size = MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, nullptr, 0);
		if (size <= 0) {
			return std::wstring();
		}
		std::wstring out(static_cast<size_t>(size), L'\0');
		MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, out.data(), size);
		if (!out.empty() && out.back() == L'\0') {
			out.pop_back();
		}
		return out;
	}

	std::wstring out(static_cast<size_t>(size), L'\0');
	MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, out.data(), size);
	if (!out.empty() && out.back() == L'\0') {
		out.pop_back();
	}
	return out;
}
#endif

constexpr int kMaxRowScan = 20000;
constexpr int kMaxFunctionScan = 4096;
constexpr int kMaxConsecutiveGetTextFail = 256;
constexpr int kMenuNameScanUpLimit = 1024;
constexpr int kMenuNameConsecutiveFailLimit = 256;
constexpr int kParentWalkLocateMaxHop = 128;

std::string TrimAsciiSpaceCopy(const std::string& value)
{
	size_t begin = 0;
	size_t end = value.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(value[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(value[end - 1])) != 0) {
		--end;
	}
	return value.substr(begin, end - begin);
}

std::string ParseSubNameFromHeader(const std::string& headerLine)
{
	std::string line = TrimAsciiSpaceCopy(headerLine);
	if (line.empty()) {
		return std::string();
	}

	// 兼容“.子程序 名称, 返回类型”以及直接传入“名称 返回类型”的情况。
	std::string remain = line;
	if (remain.front() == '.') {
		size_t pos = 1;
		while (pos < remain.size() && std::isspace(static_cast<unsigned char>(remain[pos])) == 0) {
			++pos;
		}
		while (pos < remain.size() && std::isspace(static_cast<unsigned char>(remain[pos])) != 0) {
			++pos;
		}
		remain = TrimAsciiSpaceCopy(remain.substr(pos));
	}

	size_t comma = remain.find(',');
	size_t commaCN_GBK = remain.find("\xA3\xAC");
	size_t commaCN_UTF8 = remain.find("\xEF\xBC\x8C");
	size_t cutPos = (std::min)(
		comma == std::string::npos ? std::numeric_limits<size_t>::max() : comma,
		(std::min)(
			commaCN_GBK == std::string::npos ? std::numeric_limits<size_t>::max() : commaCN_GBK,
			commaCN_UTF8 == std::string::npos ? std::numeric_limits<size_t>::max() : commaCN_UTF8));
	if (cutPos != std::numeric_limits<size_t>::max()) {
		remain = remain.substr(0, cutPos);
	}
	return TrimAsciiSpaceCopy(remain);
}

bool IsSubHeaderTextLine(const std::string& rawText)
{
	const std::string line = TrimAsciiSpaceCopy(rawText);
	if (line.empty() || line[0] != '.') {
		return false;
	}

	// UTF-8 / GBK bytes for "子程序"
	if (line.find("\xE5\xAD\x90\xE7\xA8\x8B\xE5\xBA\x8F") != std::string::npos ||
		line.find("\xD7\xD3\xB3\xCC\xD0\xF2") != std::string::npos) {
		return true;
	}

	std::string lower = line;
	std::transform(lower.begin(), lower.end(), lower.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return lower.rfind(".subroutine", 0) == 0 || lower.rfind(".sub", 0) == 0;
}

bool IsDataTypeHeaderTextLine(const std::string& rawText)
{
	const std::string line = TrimAsciiSpaceCopy(rawText);
	if (line.empty() || line[0] != '.') {
		return false;
	}

	std::string lower = line;
	std::transform(lower.begin(), lower.end(), lower.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	if (lower.rfind(".datatype", 0) == 0) {
		return true;
	}

	const std::wstring wide = MultiByteSmartToWide(line);
	if (wide.empty()) {
		return false;
	}
	return wide.rfind(L".数据类型", 0) == 0;
}

bool IsDllCommandHeaderTextLine(const std::string& rawText)
{
	const std::string line = TrimAsciiSpaceCopy(rawText);
	if (line.empty() || line[0] != '.') {
		return false;
	}

	std::string lower = line;
	std::transform(lower.begin(), lower.end(), lower.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	if (lower.rfind(".dll", 0) == 0 || lower.rfind(".dllcommand", 0) == 0) {
		return true;
	}

	const std::wstring wide = MultiByteSmartToWide(line);
	if (wide.empty()) {
		return false;
	}
	return wide.rfind(L".DLL命令", 0) == 0 || wide.rfind(L".dll命令", 0) == 0;
}

bool LocateLastBlockRowRangeByHeader(
	const std::vector<std::string>& lines,
	const std::function<bool(const std::string&)>& isHeader,
	int& outStartRow,
	int& outEndRow)
{
	outStartRow = -1;
	outEndRow = -1;
	if (lines.empty()) {
		return false;
	}

	int effectiveLast = static_cast<int>(lines.size()) - 1;
	while (effectiveLast >= 0 && TrimAsciiSpaceCopy(lines[static_cast<size_t>(effectiveLast)]).empty()) {
		--effectiveLast;
	}
	if (effectiveLast < 0) {
		return false;
	}

	for (int row = effectiveLast; row >= 0; --row) {
		if (isHeader(lines[static_cast<size_t>(row)])) {
			outStartRow = row;
			outEndRow = effectiveLast;
			return true;
		}
	}
	return false;
}

void AppendLineWithCrLf(std::string& target, const std::string& line)
{
	target += line;
	target += "\r\n";
}

std::vector<std::string> SplitLinesPreserveEmpty(const std::string& text)
{
	std::vector<std::string> lines;
	if (text.empty()) {
		return lines;
	}

	size_t start = 0;
	while (start <= text.size()) {
		size_t end = text.find_first_of("\r\n", start);
		if (end == std::string::npos) {
			lines.push_back(text.substr(start));
			break;
		}

		lines.push_back(text.substr(start, end - start));
		if (text[end] == '\r' && (end + 1) < text.size() && text[end + 1] == '\n') {
			start = end + 2;
		}
		else {
			start = end + 1;
		}
		if (start == text.size()) {
			lines.emplace_back();
			break;
		}
	}
	return lines;
}

bool LocateFunctionRangeFromPageTextByCaretRow(
	const std::string& pageCode,
	int caretRow,
	int& outStartRow,
	int& outEndRow,
	std::string* outFunctionName,
	std::string* outReason)
{
	outStartRow = -1;
	outEndRow = -1;
	if (outFunctionName != nullptr) {
		outFunctionName->clear();
	}
	if (outReason != nullptr) {
		outReason->clear();
	}

	const std::vector<std::string> lines = SplitLinesPreserveEmpty(pageCode);
	if (lines.empty()) {
		if (outReason != nullptr) {
			*outReason = "page_lines_empty";
		}
		return false;
	}

	const int maxRow = static_cast<int>(lines.size()) - 1;
	int row = caretRow;
	if (row < 0) {
		row = 0;
	}
	if (row > maxRow) {
		row = maxRow;
	}

	int startRow = -1;
	for (int i = row; i >= 0; --i) {
		if (IsSubHeaderTextLine(lines[static_cast<size_t>(i)])) {
			startRow = i;
			break;
		}
	}
	if (startRow < 0) {
		if (outReason != nullptr) {
			*outReason = "sub_header_not_found_above_row_" + std::to_string(row);
		}
		return false;
	}

	int endRow = maxRow;
	for (int i = startRow + 1; i <= maxRow; ++i) {
		if (IsSubHeaderTextLine(lines[static_cast<size_t>(i)])) {
			endRow = i - 1;
			break;
		}
	}
	while (endRow > startRow && TrimAsciiSpaceCopy(lines[static_cast<size_t>(endRow)]).empty()) {
		--endRow;
	}

	outStartRow = startRow;
	outEndRow = (std::max)(startRow, endRow);
	if (outFunctionName != nullptr) {
		std::string parsed = ParseSubNameFromHeader(lines[static_cast<size_t>(startRow)]);
		if (parsed.empty()) {
			parsed = TrimAsciiSpaceCopy(lines[static_cast<size_t>(startRow)]);
		}
		*outFunctionName = parsed;
	}
	if (outReason != nullptr) {
		*outReason =
			"ok row=" + std::to_string(outStartRow) + "-" + std::to_string(outEndRow) +
			" caretRow=" + std::to_string(caretRow) +
			" mappedRow=" + std::to_string(row);
	}
	return true;
}

BOOL CALLBACK EnumChildProcFindOutputWindow(HWND hwnd, LPARAM lParam)
{
	if (GetDlgCtrlID(hwnd) == 1011) {
		HWND* out = reinterpret_cast<HWND*>(lParam);
		*out = hwnd;
		return FALSE;
	}
	return TRUE;
}
}

IDEFacade& IDEFacade::Instance()
{
	static IDEFacade api;
	return api;
}

INT IDEFacade::RunFunctionRaw(INT code, DWORD p1, DWORD p2) const
{
	DWORD params[2] = { p1, p2 };
	return NotifySys(NES_RUN_FUNC, code, PtrToDWORD(params));
}

bool IDEFacade::RunFunction(INT code, DWORD p1, DWORD p2) const
{
	return RunFunctionRaw(code, p1, p2) != FALSE;
}

bool IDEFacade::Invoke(INT fnCode, DWORD p1, DWORD p2) const
{
	return RunFunction(fnCode, p1, p2);
}

bool IDEFacade::IsFnEnabled(INT fnCode) const
{
	bool enabled = false;
	return RunIsFuncEnabled(fnCode, enabled) && enabled;
}

bool IDEFacade::TryGetInt(INT fnCode, int& outValue) const
{
	outValue = 0;
	return Invoke(fnCode, PtrToDWORD(&outValue), 0);
}

bool IDEFacade::TryGetBool(INT fnCode, bool& outValue) const
{
	BOOL value = FALSE;
	if (!Invoke(fnCode, PtrToDWORD(&value), 0)) {
		outValue = false;
		return false;
	}
	outValue = (value == TRUE);
	return true;
}

#define IDEFACADE_DEFINE_NOARG_METHOD(methodName, fnCode) \
	bool IDEFacade::methodName() const { return Invoke(fnCode); }
IDEFACADE_NOARG_FN_LIST(IDEFACADE_DEFINE_NOARG_METHOD)
#undef IDEFACADE_DEFINE_NOARG_METHOD

bool IDEFacade::RunMoveOpenSpecRowArg(int rowIndex) const
{
	return Invoke(FN_MOVE_OPEN_SPEC_ROW_ARG, static_cast<DWORD>(rowIndex), 0);
}

bool IDEFacade::RunMoveCloseSpecRowArg(int rowIndex) const
{
	return Invoke(FN_MOVE_CLOSE_SPEC_ROW_ARG, static_cast<DWORD>(rowIndex), 0);
}

bool IDEFacade::RunMoveCaret(int rowIndex, int colIndex) const
{
	return Invoke(FN_MOVE_CARET, static_cast<DWORD>(rowIndex), static_cast<DWORD>(colIndex));
}

bool IDEFacade::RunScrollSpecHorzPos(int pos) const
{
	return Invoke(FN_SCROLL_SPEC_HORZ_POS, static_cast<DWORD>(pos), 0);
}

bool IDEFacade::RunScrollSpecVertPos(int pos) const
{
	return Invoke(FN_SCROLL_SPEC_VERT_POS, static_cast<DWORD>(pos), 0);
}

bool IDEFacade::RunBlkAddDef(int topRowIndex, int bottomRowIndex) const
{
	return Invoke(FN_BLK_ADD_DEF, static_cast<DWORD>(topRowIndex), static_cast<DWORD>(bottomRowIndex));
}

bool IDEFacade::RunBlkRemoveDef(int topRowIndex, int bottomRowIndex) const
{
	return Invoke(FN_BLK_REMOVE_DEF, static_cast<DWORD>(topRowIndex), static_cast<DWORD>(bottomRowIndex));
}

bool IDEFacade::RunInsertText(const std::string& text, bool asKeyboardInput) const
{
	return Invoke(FN_INSERT_TEXT, PtrToDWORD(text.c_str()), asKeyboardInput ? 1 : 0);
}

bool IDEFacade::RunPreCompile(bool& success) const
{
	BOOL compileOk = FALSE;
	if (!Invoke(FN_PRE_COMPILE, PtrToDWORD(&compileOk), 0)) {
		success = false;
		return false;
	}
	success = (compileOk == TRUE);
	return true;
}

bool IDEFacade::RunSetAndCompilePrgItemText(const std::string& text, bool preCompile) const
{
	return Invoke(FN_SET_AND_COMPILE_PRG_ITEM_TEXT, PtrToDWORD(text.c_str()), preCompile ? 1 : 0);
}

bool IDEFacade::RunReplaceAll2(const std::string& findText, const std::string& replaceText, bool caseSensitive) const
{
	REPLACE_ALL2_PARAM replaceParam = {};
	replaceParam.m_szFind = findText.c_str();
	replaceParam.m_szReplace = replaceText.c_str();
	replaceParam.m_blCase = caseSensitive ? TRUE : FALSE;
	return Invoke(FN_REPLACE_ALL2, PtrToDWORD(&replaceParam), 0);
}

bool IDEFacade::RunInputPrg2(const std::string& filePath) const
{
	return Invoke(FN_INPUT_PRG2, PtrToDWORD(filePath.c_str()), 0);
}

bool IDEFacade::RunAddNewEcom2(const std::string& filePath, bool& success) const
{
	BOOL addOk = FALSE;
	if (!Invoke(FN_ADD_NEW_ECOM2, PtrToDWORD(filePath.c_str()), PtrToDWORD(&addOk))) {
		success = false;
		return false;
	}
	success = (addOk == TRUE);
	return true;
}

bool IDEFacade::RunRemoveSpecEcom(int index) const
{
	return Invoke(FN_REMOVE_SPEC_ECOM, static_cast<DWORD>(index), 0);
}

bool IDEFacade::RunOpenFile2(const std::string& filePath) const
{
	return Invoke(FN_OPEN_FILE2, PtrToDWORD(filePath.c_str()), 0);
}

bool IDEFacade::RunAddTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon) const
{
	ADD_TAB_INF tabInf = {};
	tabInf.m_hWnd = hWnd;
	tabInf.m_hIcon = hIcon;
#ifdef UNICODE
	thread_local std::wstring captionW;
	thread_local std::wstring toolTipW;
	captionW = ToWide(caption);
	toolTipW = ToWide(toolTip);
	tabInf.m_szCaption = const_cast<LPWSTR>(captionW.c_str());
	tabInf.m_szToolTip = const_cast<LPWSTR>(toolTipW.c_str());
#else
	tabInf.m_szCaption = const_cast<LPSTR>(caption.c_str());
	tabInf.m_szToolTip = const_cast<LPSTR>(toolTip.c_str());
#endif
	return Invoke(FN_ADD_TAB, PtrToDWORD(&tabInf), 0);
}

bool IDEFacade::RunGetActiveWndType(int& outType) const
{
	outType = 0;
	return Invoke(FN_GET_ACTIVE_WND_TYPE, PtrToDWORD(&outType), 0);
}

bool IDEFacade::RunInputEcom(const std::string& filePath, bool& success) const
{
	BOOL importOk = FALSE;
	if (!Invoke(FN_INPUT_ECOM, PtrToDWORD(filePath.c_str()), PtrToDWORD(&importOk))) {
		success = false;
		return false;
	}
	success = (importOk == TRUE);
	return true;
}

bool IDEFacade::RunIsFuncEnabled(INT fnCode, bool& enabled) const
{
	BOOL ideEnabled = FALSE;
	if (!Invoke(FN_IS_FUNC_ENABLED, static_cast<DWORD>(fnCode), PtrToDWORD(&ideEnabled))) {
		enabled = false;
		return false;
	}
	enabled = (ideEnabled == TRUE);
	return true;
}

bool IDEFacade::RunClipGetEprgDataSize(int& size) const
{
	size = 0;
	return Invoke(FN_CLIP_GET_EPRG_DATA_SIZE, PtrToDWORD(&size), 0);
}

bool IDEFacade::RunClipGetEprgData(std::vector<uint8_t>& data) const
{
	data.clear();
	int size = 0;
	if (!RunClipGetEprgDataSize(size) || size <= 0) {
		return false;
	}

	data.resize(static_cast<size_t>(size));
	BOOL success = FALSE;
	if (!Invoke(FN_CLIP_GET_EPRG_DATA, PtrToDWORD(data.data()), PtrToDWORD(&success))) {
		data.clear();
		return false;
	}
	if (success != TRUE) {
		data.clear();
		return false;
	}
	return true;
}

bool IDEFacade::RunClipSetEprgData(const std::vector<uint8_t>& data) const
{
	if (data.empty()) {
		return false;
	}
	return Invoke(FN_CLIP_SET_EPRG_DATA, PtrToDWORD(data.data()), static_cast<DWORD>(data.size()));
}

bool IDEFacade::RunGetCaretRowIndex(int& rowIndex) const
{
	rowIndex = -1;
	return Invoke(FN_GET_CARET_ROW_INDEX, PtrToDWORD(&rowIndex), 0);
}

bool IDEFacade::RunGetCaretColIndex(int& colIndex) const
{
	colIndex = -1;
	return Invoke(FN_GET_CARET_COL_INDEX, PtrToDWORD(&colIndex), 0);
}

bool IDEFacade::ReadProgramLikeText(INT functionCode, int rowIndex, int colIndex, ProgramText& outText) const
{
	GET_PRG_TEXT_PARAM query = {};
	query.m_nRowIndex = rowIndex;
	query.m_nColIndex = colIndex;
	query.m_pBuf = nullptr;
	query.m_nBufSize = 0;

	if (!Invoke(functionCode, PtrToDWORD(&query), 0)) {
		return false;
	}

	outText.type = query.m_nType;
	outText.isTitle = (query.m_blIsTitle == TRUE);
	outText.text.clear();
	if (query.m_nBufSize <= 0) {
		return true;
	}

	std::vector<char> buffer(static_cast<size_t>(query.m_nBufSize) + 1, '\0');
	query.m_pBuf = buffer.data();
	query.m_nBufSize = static_cast<int>(buffer.size());

	if (!Invoke(functionCode, PtrToDWORD(&query), 0)) {
		return false;
	}

	outText.type = query.m_nType;
	outText.isTitle = (query.m_blIsTitle == TRUE);
	outText.text.assign(buffer.data());
	return true;
}

std::string IDEFacade::TrimAsciiSpace(const std::string& s)
{
	return TrimAsciiSpaceCopy(s);
}

std::string IDEFacade::NormalizeFunctionName(const std::string& name)
{
	std::string normalized = TrimAsciiSpace(name);
	if (!normalized.empty() && normalized.front() == '.') {
		const std::string parsed = ParseSubNameFromHeader(normalized);
		if (!parsed.empty()) {
			return parsed;
		}
	}

	size_t comma = normalized.find(',');
	size_t commaCN_GBK = normalized.find("\xA3\xAC");
	size_t commaCN_UTF8 = normalized.find("\xEF\xBC\x8C");
	size_t cutPos = (std::min)(
		comma == std::string::npos ? std::numeric_limits<size_t>::max() : comma,
		(std::min)(
			commaCN_GBK == std::string::npos ? std::numeric_limits<size_t>::max() : commaCN_GBK,
			commaCN_UTF8 == std::string::npos ? std::numeric_limits<size_t>::max() : commaCN_UTF8));
	if (cutPos != std::numeric_limits<size_t>::max()) {
		normalized = normalized.substr(0, cutPos);
	}
	return TrimAsciiSpace(normalized);
}

std::string IDEFacade::EnsureTrailingLineBreak(const std::string& text)
{
	if (text.empty()) {
		return "\r\n";
	}

	// Normalize all line breaks to CRLF for EIDE before replacement.
	std::string normalized;
	normalized.reserve(text.size() + 16);
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			normalized += "\r\n";
			continue;
		}
		if (ch == '\n') {
			normalized += "\r\n";
			continue;
		}
		normalized.push_back(ch);
	}

	if (normalized.empty()) {
		return "\r\n";
	}
	if (normalized.size() >= 2 && normalized[normalized.size() - 2] == '\r' && normalized.back() == '\n') {
		return normalized;
	}
	if (normalized.back() == '\r') {
		normalized.push_back('\n');
		return normalized;
	}
	normalized += "\r\n";
	return normalized;
}

bool IDEFacade::SelectRowRange(int startRow, int endRow) const
{
	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto totalStart = PerfClock::now();
	if (startRow < 0 || endRow < startRow) {
		AppendCodeFetchLogLine(
			"[STEP] SelectRowRange invalid range startRow=" + std::to_string(startRow) +
			" endRow=" + std::to_string(endRow));
		LogAIPerfCost(
			traceId,
			"SelectRowRange.total",
			ElapsedMs(totalStart),
			"ok=0 reason=invalid_range startRow=" + std::to_string(startRow) +
			" endRow=" + std::to_string(endRow));
		return false;
	}
	AppendCodeFetchLogLine(
		"[STEP] SelectRowRange startRow=" + std::to_string(startRow) +
		" endRow=" + std::to_string(endRow));
	RunBlkClearAllDef();
	const bool ok = RunBlkAddDef(startRow, endRow);
	AppendCodeFetchLogLine(
		"[STEP] SelectRowRange result ok=" + std::to_string(ok ? 1 : 0) +
		" startRow=" + std::to_string(startRow) +
		" endRow=" + std::to_string(endRow));
	LogAIPerfCost(
		traceId,
		"SelectRowRange.total",
		ElapsedMs(totalStart),
		"ok=" + std::to_string(ok ? 1 : 0) +
		" startRow=" + std::to_string(startRow) +
		" endRow=" + std::to_string(endRow));
	return ok;
}

bool IDEFacade::TranslateProgramRowRangeToBlockRange(int startProgramRow, int endProgramRow, int& outStartBlockRow, int& outEndBlockRow) const
{
	outStartBlockRow = -1;
	outEndBlockRow = -1;
	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto totalStart = PerfClock::now();
	AppendCodeFetchLogLine(
		"[STEP] enter TranslateProgramRowRangeToBlockRange startProgramRow=" + std::to_string(startProgramRow) +
		" endProgramRow=" + std::to_string(endProgramRow));
	if (startProgramRow < 0 || endProgramRow < startProgramRow) {
		AppendCodeFetchLogLine("[STEP] TranslateProgramRowRangeToBlockRange invalid range");
		LogAIPerfCost(
			traceId,
			"TranslateProgramRowRangeToBlockRange.total",
			ElapsedMs(totalStart),
			"ok=0 reason=invalid_range startProgramRow=" + std::to_string(startProgramRow) +
			" endProgramRow=" + std::to_string(endProgramRow));
		return false;
	}

	int nonTitleRowIndex = -1;
	for (int row = 0; row <= endProgramRow; ++row) {
		ProgramText rowText = {};
		if (!RunGetPrgText(row, -1, rowText)) {
			AppendCodeFetchLogLine("[STEP] TranslateProgramRowRangeToBlockRange read row failed row=" + std::to_string(row));
			LogAIPerfCost(
				traceId,
				"TranslateProgramRowRangeToBlockRange.total",
				ElapsedMs(totalStart),
				"ok=0 reason=read_row_failed row=" + std::to_string(row) +
				" endProgramRow=" + std::to_string(endProgramRow));
			return false;
		}

		if (!rowText.isTitle) {
			++nonTitleRowIndex;
		}

		if (row == startProgramRow) {
			if (rowText.isTitle) {
				// Map title row to the next editable row index.
				outStartBlockRow = nonTitleRowIndex + 1;
			}
			else {
				outStartBlockRow = nonTitleRowIndex;
			}
		}

		if (row == endProgramRow) {
			if (rowText.isTitle) {
				outEndBlockRow = nonTitleRowIndex;
			}
			else {
				outEndBlockRow = nonTitleRowIndex;
			}
		}
	}

	if (outStartBlockRow < 0 || outEndBlockRow < outStartBlockRow) {
		AppendCodeFetchLogLine(
			"[STEP] TranslateProgramRowRangeToBlockRange mapped invalid startBlockRow=" + std::to_string(outStartBlockRow) +
			" endBlockRow=" + std::to_string(outEndBlockRow));
		LogAIPerfCost(
			traceId,
			"TranslateProgramRowRangeToBlockRange.total",
			ElapsedMs(totalStart),
			"ok=0 reason=mapped_invalid startBlockRow=" + std::to_string(outStartBlockRow) +
			" endBlockRow=" + std::to_string(outEndBlockRow));
		return false;
	}
	AppendCodeFetchLogLine(
		"[STEP] TranslateProgramRowRangeToBlockRange mapped startBlockRow=" + std::to_string(outStartBlockRow) +
		" endBlockRow=" + std::to_string(outEndBlockRow));
	LogAIPerfCost(
		traceId,
		"TranslateProgramRowRangeToBlockRange.total",
		ElapsedMs(totalStart),
		"ok=1 startProgramRow=" + std::to_string(startProgramRow) +
		" endProgramRow=" + std::to_string(endProgramRow) +
		" startBlockRow=" + std::to_string(outStartBlockRow) +
		" endBlockRow=" + std::to_string(outEndBlockRow));
	return true;
}

bool IDEFacade::ReplaceSelectedRowsText(const std::string& text, bool preCompile) const
{
	ClipboardTextRestoreGuard clipboardGuard;
	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto totalStart = PerfClock::now();
	AppendCodeFetchLogLine(
		"[STEP] enter ReplaceSelectedRowsText inputLen=" + std::to_string(text.size()) +
		" preCompile=" + std::to_string(preCompile ? 1 : 0));

	const std::string finalText = EnsureTrailingLineBreak(text);
	bool usedPastePath = false;
	bool preCompileInvoked = false;
	bool preCompileOk = false;

	// Force paste-path to match manual copy/paste behavior in IDE.
	const bool setClipboardOk = SetClipboardTextForPaste(finalText);
	const bool pasteOk = setClipboardOk && RunEditPaste();
	if (!pasteOk) {
		AppendCodeFetchLogLine(
			"[STEP] ReplaceSelectedRowsText paste path failed setClipboardOk=" + std::to_string(setClipboardOk ? 1 : 0) +
			" pasteOk=" + std::to_string(pasteOk ? 1 : 0));
		LogAIPerfCost(
			traceId,
			"ReplaceSelectedRowsText.total",
			ElapsedMs(totalStart),
			"ok=0 reason=paste_path_failed setClipboardOk=" + std::to_string(setClipboardOk ? 1 : 0) +
			" pasteOk=" + std::to_string(pasteOk ? 1 : 0) +
			" preCompile=" + std::to_string(preCompile ? 1 : 0));
		return false;
	}
	usedPastePath = true;
	AppendCodeFetchLogLine("[STEP] ReplaceSelectedRowsText paste path success");

	if (!preCompile) {
		AppendCodeFetchLogLine(
			"[STEP] ReplaceSelectedRowsText done replaced=1 preCompile=0 usedPastePath=" + std::to_string(usedPastePath ? 1 : 0) +
			" usedFallbackPath=0");
		LogAIPerfCost(
			traceId,
			"ReplaceSelectedRowsText.total",
			ElapsedMs(totalStart),
			"ok=1 preCompile=0 usedPastePath=" + std::to_string(usedPastePath ? 1 : 0) +
			" usedFallbackPath=0" +
			" outputLen=" + std::to_string(finalText.size()));
		return true;
	}

	bool compileOk = false;
	preCompileInvoked = true;
	if (!RunPreCompile(compileOk)) {
		AppendCodeFetchLogLine("[STEP] ReplaceSelectedRowsText RunPreCompile invoke failed");
		LogAIPerfCost(
			traceId,
			"ReplaceSelectedRowsText.total",
			ElapsedMs(totalStart),
			"ok=0 reason=run_precompile_failed usedPastePath=" + std::to_string(usedPastePath ? 1 : 0) +
			" usedFallbackPath=0");
		return false;
	}
	preCompileOk = compileOk;
	AppendCodeFetchLogLine(
		"[STEP] ReplaceSelectedRowsText done replaced=1 preCompile=1 compileOk=" + std::to_string(preCompileOk ? 1 : 0) +
		" usedPastePath=" + std::to_string(usedPastePath ? 1 : 0) +
		" usedFallbackPath=0");
	LogAIPerfCost(
		traceId,
		"ReplaceSelectedRowsText.total",
		ElapsedMs(totalStart),
		"ok=" + std::to_string(preCompileOk ? 1 : 0) +
		" preCompileInvoked=" + std::to_string(preCompileInvoked ? 1 : 0) +
		" compileOk=" + std::to_string(preCompileOk ? 1 : 0) +
		" usedPastePath=" + std::to_string(usedPastePath ? 1 : 0) +
		" usedFallbackPath=0" +
		" outputLen=" + std::to_string(finalText.size()));
	return preCompileOk;
}

bool IDEFacade::BuildCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const
{
	outSnapshot = {};
	BeginCodeFetchLogSession("BuildCurrentPageSnapshot");
	AppendCodeFetchLogLine("[STEP] enter BuildCurrentPageSnapshot");

	int caretRow = -1;
	int caretCol = -1;
	if (!GetCaretPosition(caretRow, caretCol)) {
		AppendCodeFetchLogLine("[STEP] GetCaretPosition failed");
		return false;
	}
	AppendCodeFetchLogLine("[STEP] caret row=" + std::to_string(caretRow) + " col=" + std::to_string(caretCol));

	ProgramText currentRowText = {};
	if (!RunGetPrgText(caretRow, -1, currentRowText)) {
		AppendCodeFetchLogLine("[STEP] RunGetPrgText(caretRow) failed");
		return false;
	}

	int firstRow = caretRow;
	for (int i = 0; i < kMaxRowScan && firstRow > 0; ++i) {
		ProgramText probe = {};
		if (!RunGetPrgText(firstRow - 1, -1, probe)) {
			break;
		}
		--firstRow;
	}

	int lastRow = caretRow;
	for (int i = 0; i < kMaxRowScan; ++i) {
		ProgramText probe = {};
		if (!RunGetPrgText(lastRow + 1, -1, probe)) {
			break;
		}
		++lastRow;
	}

	if (firstRow < 0 || lastRow < firstRow) {
		AppendCodeFetchLogLine("[STEP] invalid page range first=" + std::to_string(firstRow) + " last=" + std::to_string(lastRow));
		return false;
	}
	AppendCodeFetchLogLine("[STEP] page range first=" + std::to_string(firstRow) + " last=" + std::to_string(lastRow));

	outSnapshot.firstRow = firstRow;
	outSnapshot.lastRow = lastRow;

	std::vector<ProgramText> rows;
	rows.reserve(static_cast<size_t>(lastRow - firstRow + 1));
	for (int row = firstRow; row <= lastRow; ++row) {
		ProgramText rowText = {};
		if (!RunGetPrgText(row, -1, rowText)) {
			AppendCodeFetchLogLine("[STEP] row read failed row=" + std::to_string(row));
			return false;
		}
		rows.push_back(rowText);
		AppendLineWithCrLf(outSnapshot.code, rowText.text);
	}

	std::vector<int> subStartRows;
	subStartRows.reserve(32);
	for (int row = firstRow; row <= lastRow; ++row) {
		const ProgramText& rowText = rows[static_cast<size_t>(row - firstRow)];
		if (!rowText.isTitle && rowText.type == VT_SUB_NAME) {
			subStartRows.push_back(row);
		}
	}

	for (size_t i = 0; i < subStartRows.size(); ++i) {
		const int start = subStartRows[i];
		const int end = (i + 1 < subStartRows.size()) ? (subStartRows[i + 1] - 1) : lastRow;
		if (end < start) {
			continue;
		}

		FunctionBlock block = {};
		block.startRow = start;
		block.endRow = end;
		block.headerCol = -1;
		block.name = NormalizeFunctionName(rows[static_cast<size_t>(start - firstRow)].text);

		for (int row = start; row <= end; ++row) {
			AppendLineWithCrLf(block.code, rows[static_cast<size_t>(row - firstRow)].text);
		}
		outSnapshot.functions.push_back(std::move(block));
	}
	AppendCodeFetchLogLine("[STEP] BuildCurrentPageSnapshot done functions=" + std::to_string(outSnapshot.functions.size()));

	return true;
}

bool IDEFacade::FindFunctionBlockByName(const PageCodeSnapshot& snapshot, const std::string& name, FunctionBlock& outBlock) const
{
	const std::string target = NormalizeFunctionName(name);
	if (target.empty()) {
		return false;
	}

	for (const auto& block : snapshot.functions) {
		if (NormalizeFunctionName(block.name) == target) {
			outBlock = block;
			return true;
		}
	}
	return false;
}

bool IDEFacade::FindCurrentFunctionBlock(const PageCodeSnapshot& snapshot, FunctionBlock& outBlock) const
{
	int caretRow = -1;
	int caretCol = -1;
	if (!GetCaretPosition(caretRow, caretCol)) {
		return false;
	}

	for (const auto& block : snapshot.functions) {
		if (caretRow >= block.startRow && caretRow <= block.endRow) {
			outBlock = block;
			return true;
		}
	}
	return false;
}

bool IDEFacade::LocateCurrentFunctionRowRange(int& outStartRow, int& outEndRow, std::string* outDiagnostics) const
{
	outStartRow = -1;
	outEndRow = -1;
	if (IsAICodeFetchDebugEnabled()) {
		AppendCodeFetchLogLine("[STEP] enter LocateCurrentFunctionRowRange");
	}
	if (outDiagnostics != nullptr) {
		outDiagnostics->clear();
	}
	auto setDiag = [&](const std::string& message) {
		if (outDiagnostics != nullptr) {
			*outDiagnostics = message;
		}
	};

	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const int scanLimit = (std::min)(kMaxRowScan, 20000);

	int caretRow = -1;
	int caretCol = -1;
	if (!GetCaretPosition(caretRow, caretCol)) {
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine("[STEP] GetCaretPosition failed in LocateCurrentFunctionRowRange");
		}
		setDiag("GetCaretPosition failed");
		return false;
	}
	if (caretRow < 0) {
		setDiag("invalid caret row");
		return false;
	}

	int startRow = -1;
	ProgramText startRowText = {};
	int pageCopyEndRow = -1;
	bool hasPageCopyRange = false;
	int scanCountUp = 0;
	int failCountUp = 0;
	int consecutiveFailUp = 0;
	const auto scanUpStart = PerfClock::now();
	for (int probeRow = caretRow; probeRow >= 0 && scanCountUp < scanLimit; --probeRow, ++scanCountUp) {
		ProgramText rowText = {};
		if (!RunGetPrgText(probeRow, -1, rowText)) {
			++failCountUp;
			++consecutiveFailUp;
			if (consecutiveFailUp >= kMaxConsecutiveGetTextFail) {
				if (IsAICodeFetchDebugEnabled()) {
					AppendCodeFetchLogLine(
						"[STEP] locate scan_up break on consecutive failures probeRow=" + std::to_string(probeRow) +
						" consecutiveFailUp=" + std::to_string(consecutiveFailUp) +
						" failLimit=" + std::to_string(kMaxConsecutiveGetTextFail));
				}
				break;
			}
			continue;
		}
		consecutiveFailUp = 0;
		if (!rowText.isTitle && rowText.type == VT_SUB_NAME) {
			startRow = probeRow;
			startRowText = rowText;
			break;
		}
		ProgramText helpText = {};
		if (RunGetPrgHelp(probeRow, -1, helpText) &&
			!helpText.isTitle &&
			helpText.type == VT_SUB_NAME) {
			startRow = probeRow;
			startRowText = helpText;
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine(
					"[STEP] locate scan_up found by FN_GET_PRG_HELP row=" + std::to_string(probeRow) +
					" text=\"" + EscapeOneLine(TruncateForPerfLog(helpText.text, 80)) + "\"");
			}
			break;
		}
	}
	LogAIPerfCost(
		traceId,
		"LocateCurrentFunctionRowRange.scan_up",
		ElapsedMs(scanUpStart),
		"caretRow=" + std::to_string(caretRow) +
		" scans=" + std::to_string(scanCountUp) +
		" fails=" + std::to_string(failCountUp) +
		" startRow=" + std::to_string(startRow));

	if (startRow < 0) {
		const int originalRow = caretRow;
		const int originalCol = caretCol;
		int walkRow = caretRow;
		int walkCol = caretCol;
		const auto walkStart = PerfClock::now();
		std::string walkStopReason = "max_hop";
		int walkFailRead = 0;
		int walkHops = 0;
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] locate fallback parent-walk begin row=" + std::to_string(walkRow) +
				" col=" + std::to_string(walkCol) +
				" maxHop=" + std::to_string(kParentWalkLocateMaxHop));
		}
		for (; walkHops < kParentWalkLocateMaxHop; ++walkHops) {
			ProgramText currentRowText = {};
			if (RunGetPrgText(walkRow, -1, currentRowText)) {
				if (!currentRowText.isTitle && currentRowText.type == VT_SUB_NAME) {
					startRow = walkRow;
					startRowText = currentRowText;
					walkStopReason = "found_sub_name";
					if (IsAICodeFetchDebugEnabled()) {
						AppendCodeFetchLogLine(
							"[STEP] locate fallback parent-walk found row=" + std::to_string(startRow) +
							" name=\"" + EscapeOneLine(NormalizeFunctionName(startRowText.text)) + "\"");
					}
					break;
				}
				if (IsAICodeFetchDebugEnabled()) {
					AppendCodeFetchLogLine(
						"[STEP] locate fallback parent-walk hop=" + std::to_string(walkHops) +
						" row=" + std::to_string(walkRow) +
						" type=" + std::to_string(currentRowText.type) +
						" isTitle=" + std::to_string(currentRowText.isTitle ? 1 : 0) +
						" text=\"" + EscapeOneLine(TruncateForPerfLog(currentRowText.text, 80)) + "\"");
				}
			}
			else {
				++walkFailRead;
				if (IsAICodeFetchDebugEnabled()) {
					AppendCodeFetchLogLine(
						"[STEP] locate fallback parent-walk read failed row=" + std::to_string(walkRow) +
						" failRead=" + std::to_string(walkFailRead));
				}
			}

			if (!MoveToParentCommand()) {
				walkStopReason = "move_parent_failed";
				break;
			}
			int nextRow = -1;
			int nextCol = -1;
			if (!GetCaretPosition(nextRow, nextCol)) {
				walkStopReason = "get_caret_failed";
				break;
			}
			if (nextRow == walkRow && nextCol == walkCol) {
				walkStopReason = "caret_not_moved";
				break;
			}
			walkRow = nextRow;
			walkCol = nextCol;
		}
		if (originalRow >= 0 && originalCol >= 0) {
			MoveCaret(originalRow, originalCol);
		}
		LogAIPerfCost(
			traceId,
			"LocateCurrentFunctionRowRange.parent_walk",
			ElapsedMs(walkStart),
			"hops=" + std::to_string(walkHops) +
			" failRead=" + std::to_string(walkFailRead) +
			" startRow=" + std::to_string(startRow) +
			" reason=" + walkStopReason);
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] locate fallback parent-walk done hops=" + std::to_string(walkHops) +
				" failRead=" + std::to_string(walkFailRead) +
				" startRow=" + std::to_string(startRow) +
				" reason=" + walkStopReason);
		}
	}

	if (startRow < 0) {
		const auto moveBackStart = PerfClock::now();
		const int originalRow = caretRow;
		const int originalCol = caretCol;
		int backRow = -1;
		int backCol = -1;
		std::string moveBackReason = "not_attempted";
		bool movedBack = false;
		bool gotBackCaret = false;
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine("[STEP] locate fallback move_back_sub begin");
		}
		movedBack = MoveBackSub();
		if (!movedBack) {
			moveBackReason = "move_back_sub_failed";
		}
		else if (!GetCaretPosition(backRow, backCol)) {
			moveBackReason = "get_caret_after_move_back_failed";
		}
		else {
			gotBackCaret = true;
			moveBackReason = "caret_ready";
			ProgramText backText = {};
			if (RunGetPrgText(backRow, -1, backText)) {
				if (!backText.isTitle && backText.type == VT_SUB_NAME) {
					startRow = backRow;
					startRowText = backText;
					moveBackReason = "found_sub_name_at_back_row";
				}
				else if (IsSubHeaderTextLine(backText.text)) {
					startRow = backRow;
					startRowText = backText;
					moveBackReason = "found_sub_header_at_back_row";
				}
			}

			if (startRow < 0) {
				for (int delta = -4; delta <= 4; ++delta) {
					const int probeRow = backRow + delta;
					if (probeRow < 0) {
						continue;
					}
					ProgramText probeText = {};
					if (RunGetPrgText(probeRow, -1, probeText) &&
						(!probeText.isTitle && probeText.type == VT_SUB_NAME)) {
						startRow = probeRow;
						startRowText = probeText;
						moveBackReason = "found_sub_name_near_back_row";
						break;
					}
					ProgramText probeHelp = {};
					if (RunGetPrgHelp(probeRow, -1, probeHelp) &&
						(!probeHelp.isTitle && probeHelp.type == VT_SUB_NAME)) {
						startRow = probeRow;
						startRowText = probeHelp;
						moveBackReason = "found_sub_name_by_help_near_back_row";
						break;
					}
				}
			}
		}

		if (originalRow >= 0 && originalCol >= 0) {
			MoveCaret(originalRow, originalCol);
		}

		LogAIPerfCost(
			traceId,
			"LocateCurrentFunctionRowRange.move_back_sub",
			ElapsedMs(moveBackStart),
			"movedBack=" + std::to_string(movedBack ? 1 : 0) +
			" gotBackCaret=" + std::to_string(gotBackCaret ? 1 : 0) +
			" backRow=" + std::to_string(backRow) +
			" backCol=" + std::to_string(backCol) +
			" startRow=" + std::to_string(startRow) +
			" reason=" + moveBackReason);
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] locate fallback move_back_sub done movedBack=" + std::to_string(movedBack ? 1 : 0) +
				" gotBackCaret=" + std::to_string(gotBackCaret ? 1 : 0) +
				" backRow=" + std::to_string(backRow) +
				" backCol=" + std::to_string(backCol) +
				" startRow=" + std::to_string(startRow) +
				" reason=" + moveBackReason);
		}
	}

	if (startRow < 0) {
		const auto pageCopyStart = PerfClock::now();
		std::string pageCode;
		std::string pageReason = "not_run";
		const bool pageCopied = GetCurrentPageCode(pageCode);
		if (!pageCopied) {
			pageReason = "GetCurrentPageCode_failed";
		}
		else {
			std::string mappedFunctionName;
			int mappedStartRow = -1;
			int mappedEndRow = -1;
			if (!LocateFunctionRangeFromPageTextByCaretRow(
				pageCode,
				caretRow,
				mappedStartRow,
				mappedEndRow,
				&mappedFunctionName,
				&pageReason)) {
				pageReason = "LocateByPageCopy_failed:" + pageReason;
			}
			else {
				startRow = mappedStartRow;
				pageCopyEndRow = mappedEndRow;
				hasPageCopyRange = true;
				ProgramText mappedStartText = {};
				if (RunGetPrgText(startRow, -1, mappedStartText)) {
					startRowText = mappedStartText;
				}
				else {
					startRowText.text = ".子程序 " + mappedFunctionName;
					startRowText.type = VT_SUB_NAME;
					startRowText.isTitle = false;
				}
			}
		}
		LogAIPerfCost(
			traceId,
			"LocateCurrentFunctionRowRange.page_copy_fallback",
			ElapsedMs(pageCopyStart),
			"ok=" + std::to_string(startRow >= 0 ? 1 : 0) +
			" startRow=" + std::to_string(startRow) +
			" reason=" + TruncateForPerfLog(pageReason));
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] locate fallback page-copy done ok=" + std::to_string(startRow >= 0 ? 1 : 0) +
				" startRow=" + std::to_string(startRow) +
				" reason=\"" + EscapeOneLine(pageReason) + "\"");
		}
	}

	if (startRow < 0) {
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] locate startRow failed caretRow=" + std::to_string(caretRow) +
				" scanUp=" + std::to_string(scanCountUp) +
				" failUp=" + std::to_string(failCountUp));
		}
		setDiag("cannot locate VT_SUB_NAME above caretRow=" + std::to_string(caretRow));
		return false;
	}

	// Compatibility fix: some hosts expose the visual ".子程序" header one row above VT_SUB_NAME.
	if (!hasPageCopyRange && !IsSubHeaderTextLine(startRowText.text) && startRow > 0) {
		ProgramText prevRow = {};
		if (RunGetPrgText(startRow - 1, -1, prevRow) && IsSubHeaderTextLine(prevRow.text)) {
			--startRow;
		}
	}

	if (hasPageCopyRange) {
		outStartRow = startRow;
		outEndRow = (std::max)(startRow, pageCopyEndRow);
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] locate done (page-copy range) startRow=" + std::to_string(outStartRow) +
				" endRow=" + std::to_string(outEndRow));
		}
		setDiag("ok(page-copy): caret-range row=" + std::to_string(outStartRow) + "-" + std::to_string(outEndRow));
		return true;
	}

	int endRow = startRow;
	int scanCountDown = 0;
	int failCountDown = 0;
	int repeatedRowCount = 0;
	std::string scanDownStopReason = "max_scan_limit";
	const int functionScanLimit = (std::min)(scanLimit, kMaxFunctionScan);
	ProgramText prevRow = startRowText;
	const auto scanDownStart = PerfClock::now();
	for (int scan = 0; scan < functionScanLimit; ++scan) {
		++scanCountDown;
		ProgramText nextRow = {};
		if (!RunGetPrgText(endRow + 1, -1, nextRow)) {
			++failCountDown;
			scanDownStopReason = "next_row_read_failed";
			break;
		}
		// Some hosts expose the next function boundary as a title row first (empty text + VT_SUB_NAME).
		// Stop before including it, otherwise range may eat into the next function.
		if (nextRow.isTitle && nextRow.type == VT_SUB_NAME) {
			scanDownStopReason = "next_sub_title_row";
			break;
		}
		if (!nextRow.isTitle && nextRow.type == VT_SUB_NAME) {
			scanDownStopReason = "next_sub_header";
			break;
		}
		if (!nextRow.isTitle && !IsSubRelatedType(nextRow.type)) {
			scanDownStopReason = "non_sub_related_type_" + std::to_string(nextRow.type);
			break;
		}

		if (nextRow.type == prevRow.type &&
			nextRow.isTitle == prevRow.isTitle &&
			nextRow.text == prevRow.text) {
			++repeatedRowCount;
			if (repeatedRowCount >= 64) {
				scanDownStopReason = "repeated_same_row";
				break;
			}
		}
		else {
			repeatedRowCount = 0;
		}

		if (!nextRow.isTitle &&
			TrimAsciiSpaceCopy(nextRow.text).empty() &&
			nextRow.type == 0) {
			scanDownStopReason = "empty_non_title_row";
			break;
		}

		++endRow;
		prevRow = nextRow;
	}

	int trimmedTailSubTitleRows = 0;
	while (endRow > startRow) {
		ProgramText tailRow = {};
		if (!RunGetPrgText(endRow, -1, tailRow)) {
			break;
		}
		if (!(tailRow.isTitle && tailRow.type == VT_SUB_NAME)) {
			break;
		}
		--endRow;
		++trimmedTailSubTitleRows;
	}

	LogAIPerfCost(
		traceId,
		"LocateCurrentFunctionRowRange.scan_down",
		ElapsedMs(scanDownStart),
		"startRow=" + std::to_string(startRow) +
		" scans=" + std::to_string(scanCountDown) +
		" fails=" + std::to_string(failCountDown) +
		" endRow=" + std::to_string(endRow) +
		" reason=" + scanDownStopReason +
		" trimmedTailSubTitleRows=" + std::to_string(trimmedTailSubTitleRows));

	outStartRow = startRow;
	outEndRow = (std::max)(startRow, endRow);
	if (IsAICodeFetchDebugEnabled()) {
		AppendCodeFetchLogLine(
			"[STEP] locate done startRow=" + std::to_string(outStartRow) +
			" endRow=" + std::to_string(outEndRow) +
			" scanUp=" + std::to_string(scanCountUp) +
			" failUp=" + std::to_string(failCountUp) +
			" scanDown=" + std::to_string(scanCountDown) +
			" failDown=" + std::to_string(failCountDown));
	}
	setDiag("ok: caret-range row=" + std::to_string(outStartRow) + "-" + std::to_string(outEndRow));
	return true;
}

bool IDEFacade::GetCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const
{
	return BuildCurrentPageSnapshot(outSnapshot);
}

bool IDEFacade::GetCurrentPageCode(std::string& outCode) const
{
	outCode.clear();
	ClipboardTextRestoreGuard clipboardGuard;
	BeginCodeFetchLogSession("GetCurrentPageCode");
	AppendCodeFetchLogLine("[STEP] enter GetCurrentPageCode");

	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto copyStart = PerfClock::now();
	bool movedSelectAll = false;
	bool copied = false;
	bool clipboardChanged = false;
	auto logCopy = [&](bool ok, const std::string& reason) {
		LogAIPerfCost(
			traceId,
			"GetCurrentPageCode.select_all_copy",
			ElapsedMs(copyStart),
			"ok=" + std::to_string(ok ? 1 : 0) +
			" reason=" + reason +
			" movedAll=" + std::to_string(movedSelectAll ? 1 : 0) +
			" copied=" + std::to_string(copied ? 1 : 0) +
			" clipboardChanged=" + std::to_string(clipboardChanged ? 1 : 0));
	};

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);
	AppendCodeFetchLogLine("[STEP] caret row=" + std::to_string(caretRow) + " col=" + std::to_string(caretCol));

	const DWORD beforeSeq = GetClipboardSequenceNumber();
	RunBlkClearAllDef();
	if (!RunMoveBlkSelAll()) {
		RunBlkClearAllDef();
		AppendCodeFetchLogLine("[STEP] RunMoveBlkSelAll failed");
		logCopy(false, "move_select_all_failed");
		return false;
	}
	movedSelectAll = true;

	copied = CopySelection();
	RunBlkClearAllDef();

	if (caretRow >= 0 && caretCol >= 0) {
		MoveCaret(caretRow, caretCol);
	}

	if (!copied) {
		AppendCodeFetchLogLine("[STEP] CopySelection failed");
		logCopy(false, "copy_selection_failed");
		return false;
	}
	clipboardChanged = (GetClipboardSequenceNumber() != beforeSeq);
	if (!clipboardChanged) {
		AppendCodeFetchLogLine("[STEP] clipboard sequence unchanged");
		logCopy(false, "clipboard_not_changed");
		return false;
	}

	const bool readOk = ReadClipboardText(outCode) && !outCode.empty();
	AppendCodeFetchLogLine(
		"[STEP] ReadClipboardText ok=" + std::to_string(readOk ? 1 : 0) +
		" len=" + std::to_string(outCode.size()) +
		" head=\"" + EscapeOneLine(outCode) + "\"");
	logCopy(readOk, readOk ? "ok" : "read_clipboard_failed");
	return readOk;
}

bool IDEFacade::GetCurrentPageName(std::string& outName, std::string* outTypeText, std::string* outDiagnostics) const
{
	outName.clear();
	if (outTypeText != nullptr) {
		outTypeText->clear();
	}
	if (outDiagnostics != nullptr) {
		outDiagnostics->clear();
	}

	const HWND mainHwnd = GetMainWindow();
	if (mainHwnd == nullptr || !IsWindow(mainHwnd)) {
		if (outDiagnostics != nullptr) {
			*outDiagnostics = "main_window_invalid";
		}
		return false;
	}

	const ActiveWindowType activeType = GetActiveWindowType();
	const std::string activeTypeText = DescribeActiveWindowTypeShort(activeType);

	const auto mdiClients = CollectChildWindowsByClassLocal(mainHwnd, "MDIClient");
	for (HWND mdiHwnd : mdiClients) {
		HWND activeChild = reinterpret_cast<HWND>(SendMessageA(mdiHwnd, WM_MDIGETACTIVE, 0, 0));
		if (activeChild == nullptr || !IsWindow(activeChild)) {
			continue;
		}

		const std::string childTitle = GetWindowTextCopyLocalA(activeChild);
		std::string typeText;
		std::string name;
		if (TrySplitPageIdentityText(childTitle, typeText, name) && !name.empty()) {
			outName = name;
			if (outTypeText != nullptr) {
				*outTypeText = !typeText.empty() ? typeText : activeTypeText;
			}
			if (outDiagnostics != nullptr) {
				*outDiagnostics = "mdi_active_child_title";
			}
			return true;
		}
	}

	const std::string mainTitle = GetWindowTextCopyLocalA(mainHwnd);
	std::string typeText;
	std::string name;
	if (TryExtractBracketPageIdentity(mainTitle, typeText, name) && !name.empty()) {
		outName = name;
		if (outTypeText != nullptr) {
			*outTypeText = !typeText.empty() ? typeText : activeTypeText;
		}
		if (outDiagnostics != nullptr) {
			*outDiagnostics = "main_window_bracket";
		}
		return true;
	}

	if (outDiagnostics != nullptr) {
		*outDiagnostics = "page_name_not_found";
	}
	return false;
}

bool IDEFacade::ReplaceCurrentPageCode(const std::string& newPageCode, bool preCompile) const
{
	if (newPageCode.empty()) {
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);

	RunBlkClearAllDef();
	if (!RunMoveBlkSelAll()) {
		RunBlkClearAllDef();
		return false;
	}

	const bool ok = ReplaceSelectedRowsText(newPageCode, preCompile);
	RunBlkClearAllDef();

	if (caretRow >= 0 && caretCol >= 0) {
		MoveCaret(caretRow, caretCol);
	}
	return ok;
}

bool IDEFacade::ReplaceRowRangeText(int startRow, int endRow, const std::string& newText, bool preCompile) const
{
	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto totalStart = PerfClock::now();
	AppendCodeFetchLogLine(
		"[STEP] enter ReplaceRowRangeText startRow=" + std::to_string(startRow) +
		" endRow=" + std::to_string(endRow) +
		" inputLen=" + std::to_string(newText.size()) +
		" preCompile=" + std::to_string(preCompile ? 1 : 0));
	auto logTotal = [&](bool ok, const std::string& reason) {
		LogAIPerfCost(
			traceId,
			"ReplaceRowRangeText.total",
			ElapsedMs(totalStart),
			"ok=" + std::to_string(ok ? 1 : 0) +
			" reason=" + TruncateForPerfLog(reason) +
			" startRow=" + std::to_string(startRow) +
			" endRow=" + std::to_string(endRow));
	};

	if (startRow < 0 || endRow < startRow) {
		AppendCodeFetchLogLine("[STEP] ReplaceRowRangeText invalid range");
		logTotal(false, "invalid_range");
		return false;
	}
	if (TrimAsciiSpace(newText).empty()) {
		AppendCodeFetchLogLine("[STEP] ReplaceRowRangeText input empty after trim");
		logTotal(false, "empty_input");
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	const bool hasCaret = GetCaretPosition(caretRow, caretCol);
	auto restoreCaret = [&]() {
		if (hasCaret && caretRow >= 0 && caretCol >= 0) {
			MoveCaret(caretRow, caretCol);
		}
	};
	auto failAndRestore = [&](const std::string& reason) {
		RunBlkClearAllDef();
		restoreCaret();
		AppendCodeFetchLogLine("[STEP] ReplaceRowRangeText failed reason=" + EscapeOneLine(reason));
		logTotal(false, reason);
		return false;
	};

	if (!MoveCaret(startRow, 0)) {
		return failAndRestore("move_caret_to_start_failed");
	}
	int selectedStartRow = startRow;
	int selectedEndRow = endRow;
	std::string selectStrategy = "program_rows";
	if (!SelectRowRange(selectedStartRow, selectedEndRow)) {
		int blockStartRow = -1;
		int blockEndRow = -1;
		if (!TranslateProgramRowRangeToBlockRange(startRow, endRow, blockStartRow, blockEndRow)) {
			return failAndRestore("select_row_range_failed_direct_and_translate");
		}
		selectedStartRow = blockStartRow;
		selectedEndRow = blockEndRow;
		selectStrategy = "translated_block_rows";
		AppendCodeFetchLogLine(
			"[STEP] ReplaceRowRangeText fallback translated blockStartRow=" + std::to_string(blockStartRow) +
			" blockEndRow=" + std::to_string(blockEndRow));
		if (!SelectRowRange(selectedStartRow, selectedEndRow)) {
			return failAndRestore("select_row_range_failed_after_translate");
		}
	}
	AppendCodeFetchLogLine(
		"[STEP] ReplaceRowRangeText select strategy=" + selectStrategy +
		" selectedStartRow=" + std::to_string(selectedStartRow) +
		" selectedEndRow=" + std::to_string(selectedEndRow));

	const bool ok = ReplaceSelectedRowsText(newText, preCompile);
	RunBlkClearAllDef();
	restoreCaret();
	if (!ok) {
		AppendCodeFetchLogLine("[STEP] ReplaceRowRangeText failed in ReplaceSelectedRowsText");
		logTotal(false, "replace_selected_rows_failed");
		return false;
	}
	AppendCodeFetchLogLine("[STEP] ReplaceRowRangeText done");
	logTotal(true, "ok");
	return true;
}

bool IDEFacade::GetFunctionCodeByName(const std::string& functionName, std::string& outCode) const
{
	PageCodeSnapshot snapshot = {};
	if (!BuildCurrentPageSnapshot(snapshot)) {
		return false;
	}

	FunctionBlock block = {};
	if (!FindFunctionBlockByName(snapshot, functionName, block)) {
		return false;
	}
	outCode = block.code;
	return true;
}

bool IDEFacade::GetCurrentFunctionCode(std::string& outCode, std::string* outDiagnostics) const
{
	outCode.clear();
	BeginCodeFetchLogSession("GetCurrentFunctionCode");
	AppendCodeFetchLogLine("[STEP] enter GetCurrentFunctionCode");
	if (outDiagnostics != nullptr) {
		outDiagnostics->clear();
	}
	auto setDiag = [&](const std::string& message) {
		if (outDiagnostics != nullptr) {
			*outDiagnostics = message;
		}
	};

	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](bool ok, int rowsRead, const std::string& reason) {
		LogAIPerfCost(
			traceId,
			"GetCurrentFunctionCode.total",
			ElapsedMs(totalStart),
			"ok=" + std::to_string(ok ? 1 : 0) +
			" rowsRead=" + std::to_string(rowsRead) +
			" reason=" + TruncateForPerfLog(reason));
	};

	int startRow = -1;
	int endRow = -1;
	std::string locateDiag;
	if (!LocateCurrentFunctionRowRange(startRow, endRow, &locateDiag)) {
		AppendCodeFetchLogLine("[STEP] LocateCurrentFunctionRowRange failed diag=\"" + EscapeOneLine(locateDiag) + "\"");
		setDiag(locateDiag.empty() ? "LocateCurrentFunctionRowRange failed" : locateDiag);
		logTotal(false, 0, locateDiag.empty() ? "locate_failed" : locateDiag);
		return false;
	}
	AppendCodeFetchLogLine("[STEP] function range start=" + std::to_string(startRow) + " end=" + std::to_string(endRow));

	int caretRow = -1;
	int caretCol = -1;
	const bool hasCaret = GetCaretPosition(caretRow, caretCol);
	const auto restoreCaret = [this, hasCaret, caretRow, caretCol]() {
		if (hasCaret && caretRow >= 0 && caretCol >= 0) {
			MoveCaret(caretRow, caretCol);
		}
	};

	const auto copyStart = PerfClock::now();
	int selectedStartRow = startRow;
	int selectedEndRow = endRow;
	std::string selectStrategy = "program_rows";

	const DWORD beforeSeq = GetClipboardSequenceNumber();
	RunBlkClearAllDef();
	if (!SelectRowRange(selectedStartRow, selectedEndRow)) {
		int blockStartRow = -1;
		int blockEndRow = -1;
		if (!TranslateProgramRowRangeToBlockRange(startRow, endRow, blockStartRow, blockEndRow)) {
			RunBlkClearAllDef();
			restoreCaret();
			AppendCodeFetchLogLine("[STEP] TranslateProgramRowRangeToBlockRange failed");
			setDiag("TranslateProgramRowRangeToBlockRange failed");
			LogAIPerfCost(
				traceId,
				"GetCurrentFunctionCode.copy_selection",
				ElapsedMs(copyStart),
				"ok=0 reason=select_row_range_failed_direct_and_translate startRow=" + std::to_string(startRow) +
				" endRow=" + std::to_string(endRow));
			logTotal(false, 0, "select_row_range_failed_direct_and_translate");
			return false;
		}
		selectedStartRow = blockStartRow;
		selectedEndRow = blockEndRow;
		selectStrategy = "translated_block_rows";
		AppendCodeFetchLogLine(
			"[STEP] GetCurrentFunctionCode fallback translated blockStartRow=" + std::to_string(blockStartRow) +
			" blockEndRow=" + std::to_string(blockEndRow));
		if (!SelectRowRange(selectedStartRow, selectedEndRow)) {
			RunBlkClearAllDef();
			restoreCaret();
			AppendCodeFetchLogLine("[STEP] SelectRowRange failed in GetCurrentFunctionCode after translate");
			setDiag("SelectRowRange failed after translate");
			LogAIPerfCost(
				traceId,
				"GetCurrentFunctionCode.copy_selection",
				ElapsedMs(copyStart),
				"ok=0 reason=select_row_range_failed_after_translate selectedStartRow=" + std::to_string(selectedStartRow) +
				" selectedEndRow=" + std::to_string(selectedEndRow));
			logTotal(false, 0, "select_row_range_failed_after_translate");
			return false;
		}
	}
	AppendCodeFetchLogLine(
		"[STEP] GetCurrentFunctionCode select strategy=" + selectStrategy +
		" selectedStartRow=" + std::to_string(selectedStartRow) +
		" selectedEndRow=" + std::to_string(selectedEndRow));

	const bool copied = CopySelection();
	RunBlkClearAllDef();
	restoreCaret();
	if (!copied) {
		AppendCodeFetchLogLine("[STEP] CopySelection failed in GetCurrentFunctionCode");
		setDiag("CopySelection failed");
		LogAIPerfCost(
			traceId,
			"GetCurrentFunctionCode.copy_selection",
			ElapsedMs(copyStart),
			"ok=0 reason=copy_selection_failed");
		logTotal(false, 0, "copy_selection_failed");
		return false;
	}

	const bool clipboardChanged = (GetClipboardSequenceNumber() != beforeSeq);
	std::string code;
	const bool readOk = ReadClipboardText(code);
	TrimTrailingLineBreaks(code);
	if (!readOk || TrimAsciiSpace(code).empty()) {
		AppendCodeFetchLogLine(
			"[STEP] ReadClipboardText failed/empty in GetCurrentFunctionCode readOk=" + std::to_string(readOk ? 1 : 0) +
			" clipboardChanged=" + std::to_string(clipboardChanged ? 1 : 0));
		setDiag("ReadClipboardText failed or empty");
		LogAIPerfCost(
			traceId,
			"GetCurrentFunctionCode.copy_selection",
			ElapsedMs(copyStart),
			"ok=0 reason=read_clipboard_failed_or_empty clipboardChanged=" + std::to_string(clipboardChanged ? 1 : 0));
		logTotal(false, 0, "read_clipboard_failed_or_empty");
		return false;
	}

	const int rowsRead = (std::max)(0, endRow - startRow + 1);
	LogAIPerfCost(
		traceId,
		"GetCurrentFunctionCode.copy_selection",
		ElapsedMs(copyStart),
		"ok=1 rowsRead=" + std::to_string(rowsRead) +
		" startRow=" + std::to_string(startRow) +
		" endRow=" + std::to_string(endRow) +
		" selectedStartRow=" + std::to_string(selectedStartRow) +
		" selectedEndRow=" + std::to_string(selectedEndRow) +
		" strategy=" + selectStrategy +
		" clipboardChanged=" + std::to_string(clipboardChanged ? 1 : 0));

	outCode = std::move(code);
	AppendCodeFetchLogLine(
		"[STEP] GetCurrentFunctionCode done len=" + std::to_string(outCode.size()) +
		" head=\"" + EscapeOneLine(outCode) + "\"");
	setDiag("ok: row=" + std::to_string(startRow) + "-" + std::to_string(endRow));
	logTotal(true, rowsRead, "ok");
	return true;
}
bool IDEFacade::ReplaceFunctionCodeByName(const std::string& functionName, const std::string& newFunctionCode, bool preCompile) const
{
	PageCodeSnapshot snapshot = {};
	if (!BuildCurrentPageSnapshot(snapshot)) {
		return false;
	}

	FunctionBlock block = {};
	if (!FindFunctionBlockByName(snapshot, functionName, block)) {
		return false;
	}
	return ReplaceRowRangeText(block.startRow, block.endRow, newFunctionCode, preCompile);
}

bool IDEFacade::ReplaceCurrentFunctionCode(const std::string& newFunctionCode, bool preCompile) const
{
	const uint64_t traceId = GetCurrentAIPerfTraceId();
	const auto totalStart = PerfClock::now();
	AppendCodeFetchLogLine(
		"[STEP] enter ReplaceCurrentFunctionCode inputLen=" + std::to_string(newFunctionCode.size()) +
		" preCompile=" + std::to_string(preCompile ? 1 : 0));
	auto logTotal = [&](bool ok, const std::string& reason, int startRow, int endRow) {
		LogAIPerfCost(
			traceId,
			"ReplaceCurrentFunctionCode.total",
			ElapsedMs(totalStart),
			"ok=" + std::to_string(ok ? 1 : 0) +
			" reason=" + TruncateForPerfLog(reason) +
			" startRow=" + std::to_string(startRow) +
			" endRow=" + std::to_string(endRow));
	};

	if (TrimAsciiSpace(newFunctionCode).empty()) {
		AppendCodeFetchLogLine("[STEP] ReplaceCurrentFunctionCode empty input after trim");
		logTotal(false, "empty_input", -1, -1);
		return false;
	}

	int startRow = -1;
	int endRow = -1;
	std::string locateDiag;
	if (!LocateCurrentFunctionRowRange(startRow, endRow, &locateDiag)) {
		AppendCodeFetchLogLine(
			"[STEP] ReplaceCurrentFunctionCode locate failed diag=\"" + EscapeOneLine(locateDiag) + "\"");
		logTotal(false, locateDiag.empty() ? "locate_failed" : locateDiag, startRow, endRow);
		return false;
	}
	AppendCodeFetchLogLine(
		"[STEP] ReplaceCurrentFunctionCode range startRow=" + std::to_string(startRow) +
		" endRow=" + std::to_string(endRow));

	const bool ok = ReplaceRowRangeText(startRow, endRow, newFunctionCode, preCompile);
	logTotal(ok, ok ? "ok" : "replace_row_range_failed", startRow, endRow);
	return ok;
}

bool IDEFacade::InsertCodeBelowFunction(const std::string& functionName, const std::string& codeToInsert, bool appendIfNotFound, bool preCompile) const
{
	if (TrimAsciiSpace(codeToInsert).empty()) {
		return false;
	}

	PageCodeSnapshot snapshot = {};
	if (!BuildCurrentPageSnapshot(snapshot)) {
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);
	const auto restoreCaret = [this, caretRow, caretCol]() {
		if (caretRow >= 0 && caretCol >= 0) {
			MoveCaret(caretRow, caretCol);
		}
	};

	std::string finalInsert = EnsureTrailingLineBreak(codeToInsert);
	bool inserted = false;

	FunctionBlock target = {};
	if (FindFunctionBlockByName(snapshot, functionName, target)) {
		if (MoveCaret(target.endRow + 1, 0)) {
			inserted = RunInsertText(finalInsert, false);
		}
		else if (MoveCaret(target.endRow, 0)) {
			RunMoveEditCaretToEnd();
			inserted = RunInsertText(std::string("\r\n") + finalInsert, false);
		}
	}
	else {
		if (!appendIfNotFound) {
			return false;
		}

		if (snapshot.lastRow >= 0 && MoveCaret(snapshot.lastRow + 1, 0)) {
			inserted = RunInsertText(finalInsert, false);
		}
		else if (snapshot.lastRow >= 0 && MoveCaret(snapshot.lastRow, 0)) {
			RunMoveEditCaretToEnd();
			inserted = RunInsertText(std::string("\r\n") + finalInsert, false);
		}
		else if (MoveCaret(0, 0)) {
			inserted = RunInsertText(finalInsert, false);
		}
	}

	if (!inserted) {
		restoreCaret();
		return false;
	}

	if (preCompile) {
		bool compileOk = false;
		if (!RunPreCompile(compileOk) || !compileOk) {
			restoreCaret();
			return false;
		}
	}

	restoreCaret();
	return true;
}

bool IDEFacade::InsertDllDeclaration(const std::string& dllDeclarationCode, bool preCompile) const
{
	if (TrimAsciiSpace(dllDeclarationCode).empty()) {
		return false;
	}

	PageCodeSnapshot snapshot = {};
	if (!BuildCurrentPageSnapshot(snapshot)) {
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);
	const auto restoreCaret = [this, caretRow, caretCol]() {
		if (caretRow >= 0 && caretCol >= 0) {
			MoveCaret(caretRow, caretCol);
		}
	};

	const std::string finalInsert = EnsureTrailingLineBreak(dllDeclarationCode);
	bool inserted = false;
	if (snapshot.lastRow >= 0 && MoveCaret(snapshot.lastRow + 1, 0)) {
		inserted = RunInsertText(finalInsert, false);
	}
	else if (snapshot.lastRow >= 0 && MoveCaret(snapshot.lastRow, 0)) {
		RunMoveEditCaretToEnd();
		inserted = RunInsertText(std::string("\r\n") + finalInsert, false);
	}
	else if (MoveCaret(0, 0)) {
		inserted = RunInsertText(finalInsert, false);
	}

	if (!inserted) {
		restoreCaret();
		return false;
	}

	if (preCompile) {
		bool compileOk = false;
		if (!RunPreCompile(compileOk) || !compileOk) {
			restoreCaret();
			return false;
		}
	}

	restoreCaret();
	return true;
}

bool IDEFacade::InsertCodeAtPageTop(const std::string& codeToInsert, bool preCompile) const
{
	if (TrimAsciiSpace(codeToInsert).empty()) {
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);
	const auto restoreCaret = [this, caretRow, caretCol]() {
		if (caretRow >= 0 && caretCol >= 0) {
			MoveCaret(caretRow, caretCol);
		}
	};

	int insertRow = 0;
	ProgramText firstRowText = {};
	if (RunGetPrgText(0, -1, firstRowText)) {
		const std::string firstLine = TrimAsciiSpace(firstRowText.text);
		if (!firstLine.empty() && firstLine.rfind(".", 0) == 0) {
			insertRow = 1;
		}
	}

	bool inserted = false;
	const std::string finalInsert = EnsureTrailingLineBreak(codeToInsert);
	if (MoveCaret(insertRow, 0)) {
		inserted = RunInsertText(finalInsert, false);
	}
	else if (MoveCaret(0, 0)) {
		inserted = RunInsertText(finalInsert, false);
	}

	if (!inserted) {
		restoreCaret();
		return false;
	}

	if (preCompile) {
		bool compileOk = false;
		if (!RunPreCompile(compileOk) || !compileOk) {
			restoreCaret();
			return false;
		}
	}

	restoreCaret();
	return true;
}

bool IDEFacade::InsertCodeAtPageBottom(const std::string& codeToInsert, bool preCompile) const
{
	if (TrimAsciiSpace(codeToInsert).empty()) {
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);
	const auto restoreCaret = [this, caretRow, caretCol]() {
		if (caretRow >= 0 && caretCol >= 0) {
			MoveCaret(caretRow, caretCol);
		}
	};

	const std::string finalInsert = EnsureTrailingLineBreak(codeToInsert);
	int lastRow = -1;
	int lastCol = -1;
	if (RunMoveBottom() && GetCaretPosition(lastRow, lastCol) && lastRow < 0) {
		lastRow = -1;
	}
	if (lastRow < 0) {
		ProgramText probe = {};
		const int seedRow = (caretRow >= 0) ? caretRow : 0;
		if (RunGetPrgText(seedRow, -1, probe)) {
			lastRow = seedRow;
			while (RunGetPrgText(lastRow + 1, -1, probe)) {
				++lastRow;
			}
		}
	}

	if (lastRow >= 0) {
		ProgramText lastRowText = {};
		if (RunGetPrgText(lastRow, -1, lastRowText)) {
			const std::string replacePayload = EnsureTrailingLineBreak(lastRowText.text) + finalInsert;
			const bool ok = ReplaceRowRangeText(lastRow, lastRow, replacePayload, preCompile);
			restoreCaret();
			return ok;
		}
	}

	bool inserted = false;
	if (MoveCaret(0, 0)) {
		inserted = RunInsertText(finalInsert, false);
	}
	if (!inserted) {
		restoreCaret();
		return false;
	}

	if (preCompile) {
		bool compileOk = false;
		if (!RunPreCompile(compileOk) || !compileOk) {
			restoreCaret();
			return false;
		}
	}

	restoreCaret();
	return true;
}

bool IDEFacade::InsertDllDeclarationByTemplate(const std::string& dllName, const std::string& commandName, const std::string& returnType, const std::string& argList, bool preCompile) const
{
	const std::string cmdName = TrimAsciiSpace(commandName);
	if (cmdName.empty()) {
		return false;
	}

	std::string decl = ".DLL鍛戒护 " + cmdName;
	if (!TrimAsciiSpace(returnType).empty()) {
		decl += ", " + TrimAsciiSpace(returnType);
	}
	if (!TrimAsciiSpace(dllName).empty()) {
		decl += ", \"" + TrimAsciiSpace(dllName) + "\"";
	}
	if (!TrimAsciiSpace(argList).empty()) {
		decl += "\r\n";
		decl += argList;
	}
	return InsertDllDeclaration(decl, preCompile);
}

bool IDEFacade::JumpToFunctionHeaderByName(const std::string& functionName) const
{
	PageCodeSnapshot snapshot = {};
	if (!BuildCurrentPageSnapshot(snapshot)) {
		return false;
	}

	FunctionBlock block = {};
	if (!FindFunctionBlockByName(snapshot, functionName, block)) {
		return false;
	}

	return MoveCaret(block.startRow, 0);
}

bool IDEFacade::RunGetPrgText(int rowIndex, int colIndex, ProgramText& outText) const
{
	const bool ok = ReadProgramLikeText(FN_GET_PRG_TEXT, rowIndex, colIndex, outText);
	if (IsAICodeFetchDebugEnabled()) {
		AppendCodeFetchLogLine(
			"[CALL] fn=FN_GET_PRG_TEXT row=" + std::to_string(rowIndex) +
			" col=" + std::to_string(colIndex) +
			" ok=" + std::to_string(ok ? 1 : 0) +
			" type=" + std::to_string(outText.type) +
			" isTitle=" + std::to_string(outText.isTitle ? 1 : 0) +
			" text=\"" + EscapeOneLine(outText.text) + "\"");
	}
	return ok;
}

bool IDEFacade::RunGetPrgHelp(int rowIndex, int colIndex, ProgramText& outText) const
{
	const bool ok = ReadProgramLikeText(FN_GET_PRG_HELP, rowIndex, colIndex, outText);
	if (IsAICodeFetchDebugEnabled()) {
		AppendCodeFetchLogLine(
			"[CALL] fn=FN_GET_PRG_HELP row=" + std::to_string(rowIndex) +
			" col=" + std::to_string(colIndex) +
			" ok=" + std::to_string(ok ? 1 : 0) +
			" type=" + std::to_string(outText.type) +
			" isTitle=" + std::to_string(outText.isTitle ? 1 : 0) +
			" text=\"" + EscapeOneLine(outText.text) + "\"");
	}
	return ok;
}

bool IDEFacade::RunGetNumEcom(int& count) const
{
	count = 0;
	return Invoke(FN_GET_NUM_ECOM, PtrToDWORD(&count), 0);
}

bool IDEFacade::RunGetEcomFileName(int index, std::string& path) const
{
	char buffer[MAX_PATH] = {};
	if (!Invoke(FN_GET_ECOM_FILE_NAME, static_cast<DWORD>(index), PtrToDWORD(buffer))) {
		return false;
	}
	path.assign(buffer);
	return true;
}

bool IDEFacade::RunGetNumLib(int& count) const
{
	count = 0;
	return Invoke(FN_GET_NUM_LIB, PtrToDWORD(&count), 0);
}

bool IDEFacade::RunGetLibInfoText(int index, std::string& text) const
{
	char* libText = nullptr;
	if (!Invoke(FN_GET_LIB_INFO_TEXT, static_cast<DWORD>(index), PtrToDWORD(&libText))) {
		return false;
	}
	text.assign(libText == nullptr ? "" : libText);
	return true;
}

HWND IDEFacade::GetMainWindow() const
{
	return reinterpret_cast<HWND>(NotifySys(NES_GET_MAIN_HWND, 0, 0));
}

bool IDEFacade::IsFunctionEnabled(INT code) const
{
	return IsFnEnabled(code);
}

IDEFacade::ActiveWindowType IDEFacade::GetActiveWindowType() const
{
	int activeType = 0;
	RunGetActiveWndType(activeType);
	return static_cast<ActiveWindowType>(activeType);
}

bool IDEFacade::GetCaretPosition(int& rowIndex, int& colIndex) const
{
	BOOL rowOk = RunGetCaretRowIndex(rowIndex);
	BOOL colOk = RunGetCaretColIndex(colIndex);
	return rowOk && colOk;
}

bool IDEFacade::GetProgramText(int rowIndex, int colIndex, ProgramText& outText) const
{
	return RunGetPrgText(rowIndex, colIndex, outText);
}

bool IDEFacade::GetProgramHelp(int rowIndex, int colIndex, ProgramText& outText) const
{
	return RunGetPrgHelp(rowIndex, colIndex, outText);
}

bool IDEFacade::InsertText(const std::string& text, bool asKeyboardInput) const
{
	return RunInsertText(text, asKeyboardInput);
}

bool IDEFacade::SetAndCompileCurrentItemText(const std::string& text, bool preCompile) const
{
	return RunSetAndCompilePrgItemText(text, preCompile);
}

bool IDEFacade::ReplaceAll(const std::string& findText, const std::string& replaceText, bool caseSensitive) const
{
	return RunReplaceAll2(findText, replaceText, caseSensitive);
}

bool IDEFacade::SelectAll() const
{
	return RunMoveBlkSelAll();
}

bool IDEFacade::CopySelection() const
{
	return RunEditCopy();
}

bool IDEFacade::GetSelectedText(std::string& outText) const
{
	outText.clear();
	ClipboardTextRestoreGuard clipboardGuard;
	const DWORD beforeSeq = GetClipboardSequenceNumber();
	if (!CopySelection()) {
		return false;
	}
	const DWORD afterSeq = GetClipboardSequenceNumber();
	if (afterSeq == beforeSeq) {
		return false;
	}
	if (!ReadClipboardText(outText)) {
		return false;
	}
	return !TrimAsciiSpace(outText).empty();
}

bool IDEFacade::SetClipboardText(const std::string& text) const
{
	if (text.empty()) {
		return false;
	}
	return SetClipboardTextForPaste(text);
}

bool IDEFacade::CopyCurrentFunctionCodeToClipboard() const
{
	std::string functionCode;
	if (!GetCurrentFunctionCode(functionCode, nullptr)) {
		return false;
	}
	if (TrimAsciiSpace(functionCode).empty()) {
		return false;
	}
	return SetClipboardTextForPaste(functionCode);
}

bool IDEFacade::MovePrevUnit() const
{
	return RunMovePrevUnit();
}

bool IDEFacade::MoveNextUnit() const
{
	return RunMoveNextUnit();
}

bool IDEFacade::MoveToParentCommand() const
{
	return RunMoveToParentCmd();
}

bool IDEFacade::MoveCaret(int rowIndex, int colIndex) const
{
	return RunMoveCaret(rowIndex, colIndex);
}

bool IDEFacade::MoveToReferencedSub() const
{
	return RunMoveSpecSub();
}

bool IDEFacade::OpenCurrentSub() const
{
	return RunMoveOpenSpecSub();
}

bool IDEFacade::CloseCurrentSub() const
{
	return RunMoveCloseSpecSub();
}

bool IDEFacade::MoveBackSub() const
{
	return RunMoveBackSub();
}

bool IDEFacade::OpenViewTab(ViewTab tab) const
{
	switch (tab)
	{
	case ViewTab::DataType:
		return RunViewDataTypeTab();
	case ViewTab::GlobalVar:
		return RunViewGlobalVarTab();
	case ViewTab::DllCommand:
		return RunViewDllcmdTab();
	case ViewTab::ConstResource:
		return RunViewConstTab();
	case ViewTab::PictureResource:
		return RunViewPicTab();
	case ViewTab::SoundResource:
		return RunViewSoundTab();
	default:
		return false;
	}
}

bool IDEFacade::SaveFile() const
{
	return RunSaveFile();
}

bool IDEFacade::OpenFile(const std::string& filePath) const
{
	return RunOpenFile2(filePath);
}

bool IDEFacade::Compile() const
{
	return RunCompile();
}

bool IDEFacade::CompileAndRun() const
{
	return RunCompileAndRun();
}

bool IDEFacade::AddOutputTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon) const
{
	return RunAddTab(hWnd, caption, toolTip, hIcon);
}

HWND IDEFacade::FindOutputWindowHandle() const
{
	HWND mainHwnd = GetMainWindow();
	if (mainHwnd == nullptr) {
		return nullptr;
	}

	HWND outputHwnd = nullptr;
	EnumChildWindows(mainHwnd, EnumChildProcFindOutputWindow, reinterpret_cast<LPARAM>(&outputHwnd));
	return outputHwnd;
}

bool IDEFacade::AppendOutputWindowText(const std::string& text) const
{
	HWND outputHwnd = FindOutputWindowHandle();
	if (outputHwnd == nullptr) {
		return false;
	}

	SendMessageA(outputHwnd, EM_SETSEL, static_cast<WPARAM>(-1), static_cast<LPARAM>(-1));
	SendMessageA(outputHwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(text.c_str()));
	return true;
}

bool IDEFacade::AppendOutputWindowLine(const std::string& text) const
{
	if (!AppendOutputWindowText(text)) {
		return false;
	}
	return AppendOutputWindowText("\r\n");
}

bool IDEFacade::GetImportedECOMCount(int& count) const
{
	return RunGetNumEcom(count);
}

bool IDEFacade::GetImportedECOMPath(int index, std::string& path) const
{
	return RunGetEcomFileName(index, path);
}

bool IDEFacade::InputECOM(const std::string& filePath, bool useNewAddMethod) const
{
	bool success = false;
	if (useNewAddMethod) {
		if (!RunAddNewEcom2(filePath, success)) {
			return false;
		}
		return success;
	}

	if (!RunInputEcom(filePath, success)) {
		return false;
	}
	return success;
}

bool IDEFacade::RemoveECOM(int index) const
{
	return RunRemoveSpecEcom(index);
}

bool IDEFacade::AddECOM(const std::string& filePath) const
{
	return InputECOM(filePath, false);
}

bool IDEFacade::AddECOM2(const std::string& filePath) const
{
	return InputECOM(filePath, true);
}

int IDEFacade::FindECOMIndex(const std::string& filePath) const
{
	int ecomCount = 0;
	if (!GetImportedECOMCount(ecomCount) || ecomCount <= 0) {
		return -1;
	}

	for (int i = 0; i < ecomCount; ++i) {
		std::string currentPath;
		if (!GetImportedECOMPath(i, currentPath)) {
			continue;
		}
		if (currentPath == filePath) {
			return i;
		}
	}
	return -1;
}

int IDEFacade::FindECOMNameIndex(const std::string& ecomName) const
{
	int ecomCount = 0;
	if (!GetImportedECOMCount(ecomCount) || ecomCount <= 0) {
		return -1;
	}

	for (int i = 0; i < ecomCount; ++i) {
		std::string currentPath;
		if (!GetImportedECOMPath(i, currentPath)) {
			continue;
		}

		std::filesystem::path pathObj(currentPath);
		if (pathObj.stem().string() == ecomName) {
			return i;
		}
	}
	return -1;
}

bool IDEFacade::RemoveECOM(const std::string& filePath) const
{
	int index = FindECOMIndex(filePath);
	if (index < 0) {
		return false;
	}
	return RemoveECOM(index);
}

void IDEFacade::RegisterContextMenuItem(UINT commandId, const std::string& text, MenuHandler handler)
{
	auto it = std::find_if(m_contextMenuItems.begin(), m_contextMenuItems.end(),
		[commandId](const ContextMenuItem& item) { return item.commandId == commandId; });

	if (it != m_contextMenuItems.end()) {
		it->text = text;
		it->handler = std::move(handler);
		return;
	}

	m_contextMenuItems.push_back(ContextMenuItem{ commandId, text, std::move(handler) });
}

void IDEFacade::ClearContextMenuItems()
{
	m_contextMenuItems.clear();
}

void IDEFacade::RefreshContextMenuEnabledState(HMENU popupMenu)
{
	if (popupMenu == nullptr || m_contextMenuItems.empty()) {
		return;
	}

	constexpr const char* kAiTranslateTextLabel = "AI翻译选中文本";
	bool hasSelectedTextResolved = false;
	bool hasSelectedText = false;

	for (const auto& item : m_contextMenuItems) {
		// Only touch items that are truly our own menu entries (same command id + same caption).
		if (GetMenuState(popupMenu, item.commandId, MF_BYCOMMAND) == 0xFFFFFFFF) {
			continue;
		}

		char title[256] = {};
		const int titleLen = GetMenuStringA(
			popupMenu,
			item.commandId,
			title,
			static_cast<int>(sizeof(title)),
			MF_BYCOMMAND);
		if (titleLen <= 0) {
			continue;
		}
		const std::string currentTitle(title, static_cast<size_t>(titleLen));
		if (currentTitle != item.text) {
			continue;
		}

		if (item.text == kAiTranslateTextLabel && !hasSelectedTextResolved) {
			hasSelectedText = IsFnEnabled(FN_EDIT_CUT);
			hasSelectedTextResolved = true;
		}
		const bool disableTranslateText = (item.text == kAiTranslateTextLabel) && !hasSelectedText;
		EnableMenuItem(
			popupMenu,
			item.commandId,
			MF_BYCOMMAND | (disableTranslateText ? MF_GRAYED : MF_ENABLED));
	}
}

bool IDEFacade::InjectContextMenuToPopup(HMENU popupMenu)
{
	if (popupMenu == nullptr || m_contextMenuItems.empty()) {
		return false;
	}

	bool containsTriggerItem = false;
	const int popupCount = GetMenuItemCount(popupMenu);
	for (int i = 0; i < popupCount; ++i) {
		wchar_t title[256] = {};
		const int len = GetMenuStringW(
			popupMenu,
			static_cast<UINT>(i),
			title,
			static_cast<int>(sizeof(title) / sizeof(title[0])),
			MF_BYPOSITION);
		if (len <= 0) {
			continue;
		}
		const std::wstring itemTitle(title, static_cast<size_t>(len));
		if (itemTitle.find(L"新方法") != std::wstring::npos ||
			itemTitle.find(L"新子程序") != std::wstring::npos) {
			containsTriggerItem = true;
			break;
		}
	}
	if (!containsTriggerItem) {
		return false;
	}

	for (int i = GetMenuItemCount(popupMenu) - 1; i >= 0; --i) {
		MENUITEMINFOW mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_SUBMENU;
		bool removeThis = false;
		if (GetMenuItemInfoW(popupMenu, static_cast<UINT>(i), TRUE, &mii) && mii.hSubMenu != nullptr) {
			wchar_t title[256] = {};
			const int len = GetMenuStringW(
				popupMenu,
				static_cast<UINT>(i),
				title,
				static_cast<int>(sizeof(title) / sizeof(title[0])),
				MF_BYPOSITION);
			if (len > 0) {
				const std::wstring itemTitle(title, static_cast<size_t>(len));
				if (itemTitle == L"AutoLinker") {
					removeThis = true;
				}
			}
		}
		if (removeThis) {
			DeleteMenu(popupMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (i > 0) {
				const UINT prevState = GetMenuState(popupMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				if (prevState != 0xFFFFFFFF && (prevState & MF_SEPARATOR) == MF_SEPARATOR) {
					DeleteMenu(popupMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				}
			}
		}
	}

	HMENU autoLinkerMenu = CreatePopupMenu();
	if (autoLinkerMenu == nullptr) {
		return false;
	}

	constexpr const char* kAiTranslateTextLabel = "AI翻译选中文本";
	constexpr const char* kCopyCurrentFunctionLabel = "复制当前函数代码";
	constexpr const char* kAiOptimizeFunctionLabel = "AI优化函数";
	constexpr const char* kAiCommentFunctionLabel = "AI为当前函数添加注释";
	constexpr const char* kAiTranslateFunctionLabel = "AI翻译当前函数+变量名";
	const bool hasSelectedText = IsFnEnabled(FN_EDIT_CUT);
	const auto isFunctionActionItem = [&](const ContextMenuItem& item) {
		return item.text == kCopyCurrentFunctionLabel ||
			item.text == kAiOptimizeFunctionLabel ||
			item.text == kAiCommentFunctionLabel ||
			item.text == kAiTranslateFunctionLabel;
	};

	std::vector<const ContextMenuItem*> functionItems;
	std::vector<const ContextMenuItem*> normalItems;
	functionItems.reserve(m_contextMenuItems.size());
	normalItems.reserve(m_contextMenuItems.size());
	for (const auto& item : m_contextMenuItems) {
		if (isFunctionActionItem(item)) {
			functionItems.push_back(&item);
		}
		else {
			normalItems.push_back(&item);
		}
	}

	const auto appendActionItem = [&](const ContextMenuItem& item) {
		const bool disableTranslateText = (item.text == kAiTranslateTextLabel) && !hasSelectedText;
		const UINT flags = MF_STRING | (disableTranslateText ? MF_GRAYED : MF_ENABLED);
		AppendMenuA(autoLinkerMenu, flags, item.commandId, item.text.c_str());
		EnableMenuItem(
			autoLinkerMenu,
			item.commandId,
			MF_BYCOMMAND | (disableTranslateText ? MF_GRAYED : MF_ENABLED));
	};

	for (const auto* item : normalItems) {
		if (item != nullptr) {
			appendActionItem(*item);
		}
	}

	if (!functionItems.empty()) {
		if (!normalItems.empty()) {
			AppendMenuA(autoLinkerMenu, MF_SEPARATOR, 0, nullptr);
		}

		std::string functionName = "未知";
		int caretRow = -1;
		int caretCol = -1;
		if (IsAICodeFetchDebugEnabled()) {
			BeginCodeFetchLogSession("ContextMenuFunctionNameScan");
			AppendCodeFetchLogLine("[STEP] enter ContextMenuFunctionNameScan");
		}
		if (GetCaretPosition(caretRow, caretCol)) {
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine(
					"[STEP] menu scan caret row=" + std::to_string(caretRow) +
					" col=" + std::to_string(caretCol) +
					" upLimit=" + std::to_string(kMenuNameScanUpLimit) +
					" failLimit=" + std::to_string(kMenuNameConsecutiveFailLimit));
			}
			int failCount = 0;
			for (int row = caretRow, scan = 0; row >= 0 && scan < kMenuNameScanUpLimit; --row, ++scan) {
				ProgramText rowText = {};
				if (!RunGetPrgText(row, -1, rowText)) {
					++failCount;
					if (IsAICodeFetchDebugEnabled()) {
						AppendCodeFetchLogLine(
							"[STEP] menu scan read failed row=" + std::to_string(row) +
							" scan=" + std::to_string(scan) +
							" failCount=" + std::to_string(failCount));
					}
					if (failCount >= kMenuNameConsecutiveFailLimit) {
						if (IsAICodeFetchDebugEnabled()) {
							AppendCodeFetchLogLine(
								"[STEP] menu scan break on consecutive failures row=" + std::to_string(row) +
								" scan=" + std::to_string(scan) +
								" failCount=" + std::to_string(failCount));
						}
						break;
					}
					continue;
				}
				failCount = 0;
				if (IsAICodeFetchDebugEnabled()) {
					AppendCodeFetchLogLine(
						"[STEP] menu scan row ok row=" + std::to_string(row) +
						" scan=" + std::to_string(scan) +
						" type=" + std::to_string(rowText.type) +
						" isTitle=" + std::to_string(rowText.isTitle ? 1 : 0) +
						" text=\"" + EscapeOneLine(TruncateForPerfLog(rowText.text, 80)) + "\"");
				}
				if (!rowText.isTitle && rowText.type == VT_SUB_NAME) {
					std::string normalized = NormalizeFunctionName(rowText.text);
					if (!normalized.empty()) {
						functionName = normalized;
						if (IsAICodeFetchDebugEnabled()) {
							AppendCodeFetchLogLine(
								"[STEP] menu scan found by upward scan row=" + std::to_string(row) +
								" function=\"" + EscapeOneLine(functionName) + "\"");
						}
					}
					break;
				}
				ProgramText helpText = {};
				if (RunGetPrgHelp(row, -1, helpText) &&
					!helpText.isTitle &&
					helpText.type == VT_SUB_NAME) {
					std::string normalized = NormalizeFunctionName(helpText.text);
					if (!normalized.empty()) {
						functionName = normalized;
						if (IsAICodeFetchDebugEnabled()) {
							AppendCodeFetchLogLine(
								"[STEP] menu scan found by FN_GET_PRG_HELP row=" + std::to_string(row) +
								" function=\"" + EscapeOneLine(functionName) + "\"");
						}
					}
					break;
				}
			}
		}
		else if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine("[STEP] menu scan GetCaretPosition failed");
		}

		if (functionName == "未知" && caretRow >= 0) {
			std::string pageCode;
			std::string pageReason = "not_run";
			if (!GetCurrentPageCode(pageCode)) {
				pageReason = "GetCurrentPageCode_failed";
			}
			else {
				int mappedStartRow = -1;
				int mappedEndRow = -1;
				std::string mappedFunctionName;
				if (!LocateFunctionRangeFromPageTextByCaretRow(
					pageCode,
					caretRow,
					mappedStartRow,
					mappedEndRow,
					&mappedFunctionName,
					&pageReason)) {
					pageReason = "LocateByPageCopy_failed:" + pageReason;
				}
				else if (!TrimAsciiSpace(mappedFunctionName).empty()) {
					functionName = EnsureGbkText(mappedFunctionName);
					pageReason =
						"ok mappedStartRow=" + std::to_string(mappedStartRow) +
						" mappedEndRow=" + std::to_string(mappedEndRow);
				}
				else {
					pageReason = "mapped_function_name_empty";
				}
			}
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine(
					"[STEP] menu scan page-copy fallback done function=\"" + EscapeOneLine(functionName) +
					"\" reason=\"" + EscapeOneLine(pageReason) + "\"");
			}
		}

		if (functionName == "未知" && caretRow >= 0 && caretCol >= 0) {
			const int originalRow = caretRow;
			const int originalCol = caretCol;
			int walkRow = caretRow;
			int walkCol = caretCol;
			int hop = 0;
			std::string stopReason = "max_hop";
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine(
					"[STEP] menu scan parent-walk begin row=" + std::to_string(walkRow) +
					" col=" + std::to_string(walkCol));
			}
			for (; hop < kParentWalkLocateMaxHop; ++hop) {
				ProgramText rowText = {};
				if (RunGetPrgText(walkRow, -1, rowText)) {
					if (!rowText.isTitle && rowText.type == VT_SUB_NAME) {
						const std::string normalized = NormalizeFunctionName(rowText.text);
						if (!normalized.empty()) {
							functionName = normalized;
							stopReason = "found_sub_name";
							if (IsAICodeFetchDebugEnabled()) {
								AppendCodeFetchLogLine(
									"[STEP] menu scan parent-walk found row=" + std::to_string(walkRow) +
									" function=\"" + EscapeOneLine(functionName) + "\"");
							}
							break;
						}
					}
				}

				if (!MoveToParentCommand()) {
					stopReason = "move_parent_failed";
					break;
				}
				int nextRow = -1;
				int nextCol = -1;
				if (!GetCaretPosition(nextRow, nextCol)) {
					stopReason = "get_caret_failed";
					break;
				}
				if (nextRow == walkRow && nextCol == walkCol) {
					stopReason = "caret_not_moved";
					break;
				}
				walkRow = nextRow;
				walkCol = nextCol;
			}
			MoveCaret(originalRow, originalCol);
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine(
					"[STEP] menu scan parent-walk done hop=" + std::to_string(hop) +
					" reason=" + stopReason +
					" function=\"" + EscapeOneLine(functionName) + "\"");
			}
		}

		if (functionName == "未知" && caretRow >= 0 && caretCol >= 0) {
			const int originalRow = caretRow;
			const int originalCol = caretCol;
			int backRow = -1;
			int backCol = -1;
			std::string moveBackReason = "not_attempted";
			bool movedBack = false;
			bool gotBackCaret = false;
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine("[STEP] menu scan move_back_sub begin");
			}
			movedBack = MoveBackSub();
			if (!movedBack) {
				moveBackReason = "move_back_sub_failed";
			}
			else if (!GetCaretPosition(backRow, backCol)) {
				moveBackReason = "get_caret_after_move_back_failed";
			}
			else {
				gotBackCaret = true;
				moveBackReason = "caret_ready";
				ProgramText backText = {};
				if (RunGetPrgText(backRow, -1, backText) &&
					!backText.isTitle &&
					backText.type == VT_SUB_NAME) {
					const std::string normalized = NormalizeFunctionName(backText.text);
					if (!normalized.empty()) {
						functionName = normalized;
						moveBackReason = "found_by_prg_text";
					}
				}
				if (functionName == "未知") {
					ProgramText backHelp = {};
					if (RunGetPrgHelp(backRow, -1, backHelp) &&
						!backHelp.isTitle &&
						backHelp.type == VT_SUB_NAME) {
						const std::string normalized = NormalizeFunctionName(backHelp.text);
						if (!normalized.empty()) {
							functionName = normalized;
							moveBackReason = "found_by_prg_help";
						}
					}
				}
			}
			MoveCaret(originalRow, originalCol);
			if (IsAICodeFetchDebugEnabled()) {
				AppendCodeFetchLogLine(
					"[STEP] menu scan move_back_sub done movedBack=" + std::to_string(movedBack ? 1 : 0) +
					" gotBackCaret=" + std::to_string(gotBackCaret ? 1 : 0) +
					" backRow=" + std::to_string(backRow) +
					" backCol=" + std::to_string(backCol) +
					" reason=" + moveBackReason +
					" function=\"" + EscapeOneLine(functionName) + "\"");
			}
		}

		const std::string currentFunctionItem =
			EnsureGbkText("函数[" + functionName + "]↓↓↓");
		AppendMenuA(autoLinkerMenu, MF_STRING | MF_GRAYED, 0, currentFunctionItem.c_str());
		if (IsAICodeFetchDebugEnabled()) {
			AppendCodeFetchLogLine(
				"[STEP] menu scan result function=\"" + EscapeOneLine(functionName) + "\"");
		}
		for (const auto* item : functionItems) {
			if (item != nullptr) {
				appendActionItem(*item);
			}
		}
	}

	const int menuCount = GetMenuItemCount(popupMenu);
	if (menuCount > 0) {
		const UINT lastState = GetMenuState(popupMenu, static_cast<UINT>(menuCount - 1), MF_BYPOSITION);
		if (lastState != 0xFFFFFFFF && (lastState & MF_SEPARATOR) != MF_SEPARATOR) {
			if (!AppendMenuA(popupMenu, MF_SEPARATOR, 0, nullptr)) {
				DestroyMenu(autoLinkerMenu);
				return false;
			}
		}
	}

	if (!AppendMenuA(
		popupMenu,
		MF_POPUP | MF_STRING,
		reinterpret_cast<UINT_PTR>(autoLinkerMenu),
		"AutoLinker")) {
		DestroyMenu(autoLinkerMenu);
		return false;
	}

	return true;
}

bool IDEFacade::HandleNotifyMessage(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	(void)dwParam2;
	if (nMsg != NL_RIGHT_POPUP_MENU_SHOW || m_contextMenuItems.empty()) {
		return false;
	}
	return InjectContextMenuToPopup(reinterpret_cast<HMENU>(dwParam1));
}

bool IDEFacade::HandleMainWindowCommand(WPARAM wParam)
{
	const UINT commandId = LOWORD(wParam);
	auto it = std::find_if(m_contextMenuItems.begin(), m_contextMenuItems.end(),
		[commandId](const ContextMenuItem& item) { return item.commandId == commandId; });

	if (it == m_contextMenuItems.end()) {
		return false;
	}

	if (it->handler) {
		it->handler();
	}
	return true;
}
