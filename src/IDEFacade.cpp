#include "IDEFacade.h"

#include <fnshare.h>
#include <lib2.h>
#include <algorithm>
#include <cstring>
#include <filesystem>
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
		std::wstring out(static_cast<size_t>(size - 1), L'\0');
		MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, out.data(), size);
		return out;
	}

	std::wstring out(static_cast<size_t>(size - 1), L'\0');
	MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, out.data(), size);
	return out;
}
#endif
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

HWND IDEFacade::GetMainWindow() const
{
	return reinterpret_cast<HWND>(NotifySys(NES_GET_MAIN_HWND, 0, 0));
}

bool IDEFacade::IsFunctionEnabled(INT code) const
{
	BOOL enabled = FALSE;
	if (!RunFunction(FN_IS_FUNC_ENABLED, static_cast<DWORD>(code), PtrToDWORD(&enabled))) {
		return false;
	}
	return enabled == TRUE;
}

IDEFacade::ActiveWindowType IDEFacade::GetActiveWindowType() const
{
	int activeType = 0;
	RunFunction(FN_GET_ACTIVE_WND_TYPE, PtrToDWORD(&activeType), 0);
	return static_cast<ActiveWindowType>(activeType);
}

bool IDEFacade::GetCaretPosition(int& rowIndex, int& colIndex) const
{
	rowIndex = -1;
	colIndex = -1;

	BOOL rowOk = RunFunction(FN_GET_CARET_ROW_INDEX, PtrToDWORD(&rowIndex), 0);
	BOOL colOk = RunFunction(FN_GET_CARET_COL_INDEX, PtrToDWORD(&colIndex), 0);
	return rowOk && colOk;
}

bool IDEFacade::ReadProgramLikeText(INT functionCode, int rowIndex, int colIndex, ProgramText& outText) const
{
	GET_PRG_TEXT_PARAM query = {};
	query.m_nRowIndex = rowIndex;
	query.m_nColIndex = colIndex;
	query.m_pBuf = nullptr;
	query.m_nBufSize = 0;

	if (!RunFunction(functionCode, PtrToDWORD(&query), 0)) {
		return false;
	}

	outText.type = query.m_nType;
	outText.isTitle = query.m_blIsTitle == TRUE;
	outText.text.clear();
	if (query.m_nBufSize <= 0) {
		return true;
	}

	std::vector<char> buffer(static_cast<size_t>(query.m_nBufSize) + 1, '\0');
	query.m_pBuf = buffer.data();
	query.m_nBufSize = static_cast<int>(buffer.size());

	if (!RunFunction(functionCode, PtrToDWORD(&query), 0)) {
		return false;
	}

	outText.type = query.m_nType;
	outText.isTitle = query.m_blIsTitle == TRUE;
	outText.text.assign(buffer.data());
	return true;
}

bool IDEFacade::GetProgramText(int rowIndex, int colIndex, ProgramText& outText) const
{
	return ReadProgramLikeText(FN_GET_PRG_TEXT, rowIndex, colIndex, outText);
}

bool IDEFacade::GetProgramHelp(int rowIndex, int colIndex, ProgramText& outText) const
{
	return ReadProgramLikeText(FN_GET_PRG_HELP, rowIndex, colIndex, outText);
}

bool IDEFacade::InsertText(const std::string& text, bool asKeyboardInput) const
{
	return RunFunction(FN_INSERT_TEXT, PtrToDWORD(text.c_str()), asKeyboardInput ? 1 : 0);
}

bool IDEFacade::SetAndCompileCurrentItemText(const std::string& text, bool preCompile) const
{
	return RunFunction(FN_SET_AND_COMPILE_PRG_ITEM_TEXT, PtrToDWORD(text.c_str()), preCompile ? 1 : 0);
}

bool IDEFacade::ReplaceAll(const std::string& findText, const std::string& replaceText, bool caseSensitive) const
{
	REPLACE_ALL2_PARAM replaceParam = {};
	replaceParam.m_szFind = findText.c_str();
	replaceParam.m_szReplace = replaceText.c_str();
	replaceParam.m_blCase = caseSensitive ? TRUE : FALSE;
	return RunFunction(FN_REPLACE_ALL2, PtrToDWORD(&replaceParam), 0);
}

bool IDEFacade::SelectAll() const
{
	return RunFunction(FN_MOVE_BLK_SEL_ALL);
}

bool IDEFacade::CopySelection() const
{
	return RunFunction(FN_EDIT_COPY);
}

bool IDEFacade::CopyCurrentFunctionCodeToClipboard() const
{
	int originalRow = -1;
	int originalCol = -1;
	if (!GetCaretPosition(originalRow, originalCol)) {
		return false;
	}
	const int initialRow = originalRow;
	const int initialCol = originalCol;

	int workRow = originalRow;
	int workCol = originalCol;
	ProgramText currentText = {};
	if (!GetProgramText(workRow, workCol, currentText)) {
		return false;
	}

	bool movedIntoSub = false;
	if (!IsSubRelatedType(currentText.type)) {
		movedIntoSub = OpenCurrentSub();
		if (!movedIntoSub) {
			return false;
		}
		if (!GetCaretPosition(workRow, workCol) || !GetProgramText(workRow, workCol, currentText)) {
			MoveBackSub();
			return false;
		}
	}

	int startRow = -1;
	for (int row = workRow; row >= 0; --row) {
		ProgramText rowText = {};
		if (!GetProgramText(row, workCol, rowText)) {
			continue;
		}
		if (rowText.isTitle) {
			continue;
		}
		if (rowText.type == VT_SUB_NAME) {
			startRow = row;
			break;
		}
	}

	if (startRow < 0) {
		if (movedIntoSub) {
			MoveBackSub();
		}
		return false;
	}

	int endRow = startRow;
	for (int row = startRow + 1;; ++row) {
		ProgramText rowText = {};
		if (!GetProgramText(row, workCol, rowText)) {
			endRow = row - 1;
			break;
		}
		if (rowText.isTitle) {
			continue;
		}
		if (rowText.type == VT_SUB_NAME) {
			endRow = row - 1;
			break;
		}
		if (!IsSubRelatedType(rowText.type)) {
			endRow = row - 1;
			break;
		}
	}

	if (endRow < startRow) {
		endRow = startRow;
	}

	RunFunction(FN_BLK_CLEAR_ALL_DEF, 0, 0);
	bool selected = RunFunction(FN_BLK_ADD_DEF, static_cast<DWORD>(startRow), static_cast<DWORD>(endRow));
	bool copied = selected && CopySelection();
	RunFunction(FN_BLK_CLEAR_ALL_DEF, 0, 0);

	if (copied) {
		TrimClipboardLastLine();
	}

	if (movedIntoSub) {
		MoveBackSub();
	}
	MoveCaret(initialRow, initialCol);
	return copied;
}

bool IDEFacade::MovePrevUnit() const
{
	return RunFunction(FN_MOVE_PREV_UNIT);
}

bool IDEFacade::MoveNextUnit() const
{
	return RunFunction(FN_MOVE_NEXT_UNIT);
}

bool IDEFacade::MoveToParentCommand() const
{
	return RunFunction(FN_MOVE_TO_PARENT_CMD);
}

bool IDEFacade::MoveCaret(int rowIndex, int colIndex) const
{
	return RunFunction(FN_MOVE_CARET, static_cast<DWORD>(rowIndex), static_cast<DWORD>(colIndex));
}

bool IDEFacade::MoveToReferencedSub() const
{
	return RunFunction(FN_MOVE_SPEC_SUB);
}

bool IDEFacade::OpenCurrentSub() const
{
	return RunFunction(FN_MOVE_OPEN_SPEC_SUB);
}

bool IDEFacade::CloseCurrentSub() const
{
	return RunFunction(FN_MOVE_CLOSE_SPEC_SUB);
}

bool IDEFacade::MoveBackSub() const
{
	return RunFunction(FN_MOVE_BACK_SUB);
}

bool IDEFacade::OpenViewTab(ViewTab tab) const
{
	INT funcNo = FN_NULL;
	switch (tab)
	{
	case ViewTab::DataType:
		funcNo = FN_VIEW_DATA_TYPE_TAB;
		break;
	case ViewTab::GlobalVar:
		funcNo = FN_VIEW_GLOBAL_VAR_TAB;
		break;
	case ViewTab::DllCommand:
		funcNo = FN_VIEW_DLLCMD_TAB;
		break;
	case ViewTab::ConstResource:
		funcNo = FN_VIEW_CONST_TAB;
		break;
	case ViewTab::PictureResource:
		funcNo = FN_VIEW_PIC_TAB;
		break;
	case ViewTab::SoundResource:
		funcNo = FN_VIEW_SOUND_TAB;
		break;
	default:
		return false;
	}

	return RunFunction(funcNo);
}

bool IDEFacade::SaveFile() const
{
	return RunFunction(FN_SAVE_FILE);
}

bool IDEFacade::OpenFile(const std::string& filePath) const
{
	return RunFunction(FN_OPEN_FILE2, PtrToDWORD(filePath.c_str()), 0);
}

bool IDEFacade::Compile() const
{
	return RunFunction(FN_COMPILE);
}

bool IDEFacade::CompileAndRun() const
{
	return RunFunction(FN_COMPILE_AND_RUN);
}

bool IDEFacade::AddOutputTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon) const
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
	return RunFunction(FN_ADD_TAB, PtrToDWORD(&tabInf), 0);
}

bool IDEFacade::GetImportedECOMCount(int& count) const
{
	count = 0;
	return RunFunction(FN_GET_NUM_ECOM, PtrToDWORD(&count), 0);
}

bool IDEFacade::GetImportedECOMPath(int index, std::string& path) const
{
	char buffer[MAX_PATH] = {};
	if (!RunFunction(FN_GET_ECOM_FILE_NAME, static_cast<DWORD>(index), PtrToDWORD(buffer))) {
		return false;
	}
	path.assign(buffer);
	return true;
}

bool IDEFacade::InputECOM(const std::string& filePath, bool useNewAddMethod) const
{
	BOOL success = FALSE;
	INT fnNo = useNewAddMethod ? FN_ADD_NEW_ECOM2 : FN_INPUT_ECOM;
	if (!RunFunction(fnNo, PtrToDWORD(filePath.c_str()), PtrToDWORD(&success))) {
		return false;
	}
	return success == TRUE;
}

bool IDEFacade::RemoveECOM(int index) const
{
	return RunFunction(FN_REMOVE_SPEC_ECOM, static_cast<DWORD>(index), 0);
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

	AppendMenuA(popupMenu, MF_SEPARATOR, 0, nullptr);
	for (const auto& item : m_contextMenuItems) {
		AppendMenuA(popupMenu, MF_STRING | MF_ENABLED, item.commandId, item.text.c_str());
		EnableMenuItem(popupMenu, item.commandId, MF_BYCOMMAND | MF_ENABLED);
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
