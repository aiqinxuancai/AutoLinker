#include "IDEFacade.h"

#include <fnshare.h>
#include <lib2.h>
#include <algorithm>
#include <cctype>
#include <cstring>
#include <filesystem>
#include <limits>
#include <utility>
#include <vector>

namespace {
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

constexpr int kMaxRowScan = 200000;
constexpr int kMaxFunctionScan = 4096;

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
	std::wstring captionW = ToWide(caption);
	std::wstring toolTipW = ToWide(toolTip);
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
	if (startRow < 0 || endRow < startRow) {
		return false;
	}
	RunBlkClearAllDef();
	return RunBlkAddDef(startRow, endRow);
}

bool IDEFacade::TranslateProgramRowRangeToBlockRange(int startProgramRow, int endProgramRow, int& outStartBlockRow, int& outEndBlockRow) const
{
	outStartBlockRow = -1;
	outEndBlockRow = -1;
	if (startProgramRow < 0 || endProgramRow < startProgramRow) {
		return false;
	}

	int nonTitleRowIndex = -1;
	for (int row = 0; row <= endProgramRow; ++row) {
		ProgramText rowText = {};
		if (!RunGetPrgText(row, -1, rowText)) {
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
		return false;
	}
	return true;
}

bool IDEFacade::ReplaceSelectedRowsText(const std::string& text, bool preCompile) const
{
	ClipboardTextRestoreGuard clipboardGuard;

	const std::string finalText = EnsureTrailingLineBreak(text);
	bool replaced = false;

	// Prefer paste-path since it matches manual replacement behavior in IDE.
	if (SetClipboardTextForPaste(finalText) && RunEditPaste()) {
		replaced = true;
	}
	if (!replaced) {
		if (!RunRemove()) {
			return false;
		}
		const std::string insertPayload = EnsureGbkText(finalText);
		if (!RunInsertText(insertPayload, false)) {
			return false;
		}
	}

	if (!preCompile) {
		return true;
	}

	bool compileOk = false;
	if (!RunPreCompile(compileOk)) {
		return false;
	}
	return compileOk;
}

bool IDEFacade::BuildCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const
{
	outSnapshot = {};

	int caretRow = -1;
	int caretCol = -1;
	if (!GetCaretPosition(caretRow, caretCol)) {
		return false;
	}

	ProgramText currentRowText = {};
	if (!RunGetPrgText(caretRow, -1, currentRowText)) {
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
		return false;
	}

	outSnapshot.firstRow = firstRow;
	outSnapshot.lastRow = lastRow;

	std::vector<ProgramText> rows;
	rows.reserve(static_cast<size_t>(lastRow - firstRow + 1));
	for (int row = firstRow; row <= lastRow; ++row) {
		ProgramText rowText = {};
		if (!RunGetPrgText(row, -1, rowText)) {
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
	if (outDiagnostics != nullptr) {
		outDiagnostics->clear();
	}
	auto setDiag = [&](const std::string& message) {
		if (outDiagnostics != nullptr) {
			*outDiagnostics = message;
		}
	};

	int caretRow = -1;
	int caretCol = -1;
	if (!GetCaretPosition(caretRow, caretCol)) {
		setDiag("GetCaretPosition failed");
		return false;
	}
	if (caretRow < 0) {
		setDiag("invalid caret row");
		return false;
	}

	int startRow = -1;
	ProgramText startRowText = {};
	for (int probeRow = caretRow, scan = 0; probeRow >= 0 && scan < kMaxRowScan; --probeRow, ++scan) {
		ProgramText rowText = {};
		if (!RunGetPrgText(probeRow, -1, rowText)) {
			continue;
		}
		if (!rowText.isTitle && rowText.type == VT_SUB_NAME) {
			startRow = probeRow;
			startRowText = rowText;
			break;
		}
	}
	if (startRow < 0) {
		setDiag("cannot locate VT_SUB_NAME above caretRow=" + std::to_string(caretRow));
		return false;
	}

	// Compatibility fix: some hosts expose the visual ".子程序" header one row above VT_SUB_NAME.
	if (!IsSubHeaderTextLine(startRowText.text) && startRow > 0) {
		ProgramText prevRow = {};
		if (RunGetPrgText(startRow - 1, -1, prevRow) && IsSubHeaderTextLine(prevRow.text)) {
			--startRow;
		}
	}

	int endRow = startRow;
	for (int scan = 0; scan < kMaxRowScan; ++scan) {
		ProgramText nextRow = {};
		if (!RunGetPrgText(endRow + 1, -1, nextRow)) {
			break;
		}
		if (!nextRow.isTitle && nextRow.type == VT_SUB_NAME) {
			break;
		}
		++endRow;
	}

	outStartRow = startRow;
	outEndRow = (std::max)(startRow, endRow);
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

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);

	const DWORD beforeSeq = GetClipboardSequenceNumber();
	RunBlkClearAllDef();
	if (!RunMoveBlkSelAll()) {
		RunBlkClearAllDef();
		return false;
	}

	const bool copied = CopySelection();
	RunBlkClearAllDef();

	if (caretRow >= 0 && caretCol >= 0) {
		MoveCaret(caretRow, caretCol);
	}

	if (!copied) {
		return false;
	}
	if (GetClipboardSequenceNumber() == beforeSeq) {
		return false;
	}
	return ReadClipboardText(outCode) && !outCode.empty();
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
	if (startRow < 0 || endRow < startRow) {
		return false;
	}
	if (TrimAsciiSpace(newText).empty()) {
		return false;
	}

	int caretRow = -1;
	int caretCol = -1;
	GetCaretPosition(caretRow, caretCol);

	if (!MoveCaret(startRow, 0)) {
		return false;
	}
	int blockStartRow = -1;
	int blockEndRow = -1;
	if (!TranslateProgramRowRangeToBlockRange(startRow, endRow, blockStartRow, blockEndRow)) {
		return false;
	}
	if (!SelectRowRange(blockStartRow, blockEndRow)) {
		return false;
	}

	const bool ok = ReplaceSelectedRowsText(newText, preCompile);
	RunBlkClearAllDef();

	if (caretRow >= 0 && caretCol >= 0) {
		MoveCaret(caretRow, caretCol);
	}
	return ok;
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
	if (outDiagnostics != nullptr) {
		outDiagnostics->clear();
	}
	auto setDiag = [&](const std::string& message) {
		if (outDiagnostics != nullptr) {
			*outDiagnostics = message;
		}
	};

	int startRow = -1;
	int endRow = -1;
	std::string locateDiag;
	if (!LocateCurrentFunctionRowRange(startRow, endRow, &locateDiag)) {
		setDiag(locateDiag.empty() ? "LocateCurrentFunctionRowRange failed" : locateDiag);
		return false;
	}

	std::string code;
	for (int row = startRow; row <= endRow; ++row) {
		ProgramText rowText = {};
		if (!RunGetPrgText(row, -1, rowText)) {
			setDiag("RunGetPrgText failed at row=" + std::to_string(row));
			return false;
		}
		AppendLineWithCrLf(code, rowText.text);
	}
	TrimTrailingLineBreaks(code);
	if (TrimAsciiSpace(code).empty()) {
		setDiag("located function block but code is empty after trim");
		return false;
	}

	outCode = std::move(code);
	setDiag("ok: row=" + std::to_string(startRow) + "-" + std::to_string(endRow));
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
	if (TrimAsciiSpace(newFunctionCode).empty()) {
		return false;
	}

	int startRow = -1;
	int endRow = -1;
	if (!LocateCurrentFunctionRowRange(startRow, endRow, nullptr)) {
		return false;
	}

	return ReplaceRowRangeText(startRow, endRow, newFunctionCode, preCompile);
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
	return ReadProgramLikeText(FN_GET_PRG_TEXT, rowIndex, colIndex, outText);
}

bool IDEFacade::RunGetPrgHelp(int rowIndex, int colIndex, ProgramText& outText) const
{
	return ReadProgramLikeText(FN_GET_PRG_HELP, rowIndex, colIndex, outText);
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
	const bool hasSelectedText = IsFnEnabled(FN_EDIT_CUT);

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

		const bool disableTranslateText = (item.text == kAiTranslateTextLabel) && !hasSelectedText;
		EnableMenuItem(
			popupMenu,
			item.commandId,
			MF_BYCOMMAND | (disableTranslateText ? MF_GRAYED : MF_ENABLED));
	}
}

bool IDEFacade::HandleNotifyMessage(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	(void)dwParam2;
	if (nMsg != NL_RIGHT_POPUP_MENU_SHOW || m_contextMenuItems.empty()) {
		return false;
	}

	HMENU popupMenu = reinterpret_cast<HMENU>(dwParam1);
	if (popupMenu == nullptr) {
		return false;
	}

	HMENU autoLinkerMenu = CreatePopupMenu();
	if (autoLinkerMenu == nullptr) {
		return false;
	}

	constexpr const char* kAiTranslateTextLabel = "AI翻译选中文本";
	const bool hasSelectedText = IsFnEnabled(FN_EDIT_CUT);
	for (const auto& item : m_contextMenuItems) {
		const bool disableTranslateText = (item.text == kAiTranslateTextLabel) && !hasSelectedText;
		UINT flags = MF_STRING | (disableTranslateText ? MF_GRAYED : MF_ENABLED);
		if (disableTranslateText) {
			flags = MF_STRING | MF_GRAYED;
		}
		AppendMenuA(autoLinkerMenu, flags, item.commandId, item.text.c_str());
		EnableMenuItem(
			autoLinkerMenu,
			item.commandId,
			MF_BYCOMMAND | (disableTranslateText ? MF_GRAYED : MF_ENABLED));
		if (item.text == kAiTranslateTextLabel) {
			AppendMenuA(autoLinkerMenu, MF_SEPARATOR, 0, nullptr);
		}
	}

	const int menuCount = GetMenuItemCount(popupMenu);
	int insertPos = menuCount >= 0 ? menuCount : 0;
	for (int i = 0; i < menuCount; ++i) {
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
		if (itemTitle.find(L"撤销") != std::wstring::npos ||
			itemTitle.find(L"Undo") != std::wstring::npos ||
			itemTitle.find(L"undo") != std::wstring::npos) {
			insertPos = i;
			break;
		}
	}

	if (!InsertMenuA(
		popupMenu,
		static_cast<UINT>(insertPos),
		MF_BYPOSITION | MF_SEPARATOR,
		0,
		nullptr)) {
		DestroyMenu(autoLinkerMenu);
		return false;
	}

	if (!InsertMenuA(
		popupMenu,
		static_cast<UINT>(insertPos + 1),
		MF_BYPOSITION | MF_POPUP | MF_STRING,
		reinterpret_cast<UINT_PTR>(autoLinkerMenu),
		"AutoLinker")) {
		DestroyMenu(autoLinkerMenu);
		return false;
	}

	return true;
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
