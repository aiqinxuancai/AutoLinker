#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#include <PublicIDEFunctions.h>
#include <cstdint>
#include <functional>
#include <string>
#include <vector>

#include "IDEFacadeNoArgFnList.inl"

class IDEFacade {
public:
	// IDE 当前激活页面类型（对应 FN_GET_ACTIVE_WND_TYPE）。
	enum class ActiveWindowType : int {
		None = 0,
		Module = 1,
		UserDataType = 2,
		GlobalVar = 3,
		DllCommand = 4,
		FormDesigner = 5,
		ConstResource = 6,
		PictureResource = 7,
		SoundResource = 8
	};

	// 常用视图页签（用于 OpenViewTab）。
	enum class ViewTab {
		DataType,
		GlobalVar,
		DllCommand,
		ConstResource,
		PictureResource,
		SoundResource
	};

	// 代码文本查询结果（对应 FN_GET_PRG_TEXT / FN_GET_PRG_HELP）。
	struct ProgramText {
		std::string text;
		int type = 0;
		bool isTitle = false;
	};

	// 子程序代码块信息（按当前页扫描得到）。
	struct FunctionBlock {
		std::string name;
		int startRow = -1;
		int endRow = -1;
		int headerCol = -1;
		std::string code;
	};

	// 当前页代码快照：整页文本 + 子程序列表。
	struct PageCodeSnapshot {
		std::string code;
		int firstRow = -1;
		int lastRow = -1;
		std::vector<FunctionBlock> functions;
	};

	using MenuHandler = std::function<void()>;

	static IDEFacade& Instance();

	// 直接调用 IDE 功能号（FN_*），返回原始结果/布尔结果。
	INT RunFunctionRaw(INT code, DWORD p1 = 0, DWORD p2 = 0) const;
	bool RunFunction(INT code, DWORD p1 = 0, DWORD p2 = 0) const;

	// 通用调用与结果辅助。
	bool Invoke(INT fnCode, DWORD p1 = 0, DWORD p2 = 0) const;
	bool IsFnEnabled(INT fnCode) const;
	bool TryGetInt(INT fnCode, int& outValue) const;
	bool TryGetBool(INT fnCode, bool& outValue) const;

	// 无参 FN_* 的一键封装（由宏列表展开）。
#define IDEFACADE_DECLARE_NOARG_METHOD(methodName, fnCode) bool methodName() const;
	IDEFACADE_NOARG_FN_LIST(IDEFACADE_DECLARE_NOARG_METHOD)
#undef IDEFACADE_DECLARE_NOARG_METHOD

	// 常用带参 FN_* 封装。
	bool RunMoveOpenSpecRowArg(int rowIndex) const;
	bool RunMoveCloseSpecRowArg(int rowIndex) const;
	bool RunMoveCaret(int rowIndex, int colIndex) const;
	bool RunScrollSpecHorzPos(int pos) const;
	bool RunScrollSpecVertPos(int pos) const;
	bool RunBlkAddDef(int topRowIndex, int bottomRowIndex) const;
	bool RunBlkRemoveDef(int topRowIndex, int bottomRowIndex) const;
	bool RunInsertText(const std::string& text, bool asKeyboardInput = false) const;
	bool RunPreCompile(bool& success) const;
	bool RunSetAndCompilePrgItemText(const std::string& text, bool preCompile = true) const;
	bool RunReplaceAll2(const std::string& findText, const std::string& replaceText, bool caseSensitive = false) const;
	bool RunInputPrg2(const std::string& filePath) const;
	bool RunAddNewEcom2(const std::string& filePath, bool& success) const;
	bool RunRemoveSpecEcom(int index) const;
	bool RunOpenFile2(const std::string& filePath) const;
	bool RunAddTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon = nullptr) const;
	bool RunGetActiveWndType(int& outType) const;
	bool RunInputEcom(const std::string& filePath, bool& success) const;
	bool RunIsFuncEnabled(INT fnCode, bool& enabled) const;
	bool RunClipGetEprgDataSize(int& size) const;
	bool RunClipGetEprgData(std::vector<uint8_t>& data) const;
	bool RunClipSetEprgData(const std::vector<uint8_t>& data) const;
	bool RunGetCaretRowIndex(int& rowIndex) const;
	bool RunGetCaretColIndex(int& colIndex) const;
	bool RunGetPrgText(int rowIndex, int colIndex, ProgramText& outText) const;
	bool RunGetPrgHelp(int rowIndex, int colIndex, ProgramText& outText) const;
	bool RunGetNumEcom(int& count) const;
	bool RunGetEcomFileName(int index, std::string& path) const;
	bool RunGetNumLib(int& count) const;
	bool RunGetLibInfoText(int index, std::string& text) const;

	// IDE 状态：光标/文本读取。
	HWND GetMainWindow() const;
	bool IsFunctionEnabled(INT code) const;
	ActiveWindowType GetActiveWindowType() const;
	bool GetCaretPosition(int& rowIndex, int& colIndex) const;
	bool GetProgramText(int rowIndex, int colIndex, ProgramText& outText) const;
	bool GetProgramHelp(int rowIndex, int colIndex, ProgramText& outText) const;

	// 编辑操作封装。
	bool InsertText(const std::string& text, bool asKeyboardInput = false) const;
	bool SetAndCompileCurrentItemText(const std::string& text, bool preCompile = true) const;
	bool ReplaceAll(const std::string& findText, const std::string& replaceText, bool caseSensitive = false) const;
	bool SelectAll() const;
	bool CopySelection() const;
	// 获取当前选中文本（通过 IDE 复制后读取剪贴板）。
	bool GetSelectedText(std::string& outText) const;
	// 将文本写入系统剪贴板（同时提供 Unicode/ANSI 格式）。
	bool SetClipboardText(const std::string& text) const;
	bool CopyCurrentFunctionCodeToClipboard() const;

	// 代码导航。
	bool MovePrevUnit() const;
	bool MoveNextUnit() const;
	bool MoveToParentCommand() const;
	bool MoveCaret(int rowIndex, int colIndex) const;
	bool MoveToReferencedSub() const;
	bool OpenCurrentSub() const;
	bool CloseCurrentSub() const;
	bool MoveBackSub() const;

	// 页面切换与编译保存。
	bool OpenViewTab(ViewTab tab) const;
	bool SaveFile() const;
	bool OpenFile(const std::string& filePath) const;
	bool Compile() const;
	bool CompileAndRun() const;
	bool AddOutputTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon = nullptr) const;
	// 向 IDE 输出窗口追加文本/行文本（行文本会自动追加 CRLF）。
	bool AppendOutputWindowText(const std::string& text) const;
	bool AppendOutputWindowLine(const std::string& text) const;

	// ECOM 管理。
	bool AddECOM(const std::string& filePath) const;
	bool AddECOM2(const std::string& filePath) const;
	bool GetImportedECOMCount(int& count) const;
	bool GetImportedECOMPath(int index, std::string& path) const;
	bool InputECOM(const std::string& filePath, bool useNewAddMethod = false) const;
	bool RemoveECOM(int index) const;
	bool RemoveECOM(const std::string& filePath) const;
	int FindECOMIndex(const std::string& filePath) const;
	int FindECOMNameIndex(const std::string& ecomName) const;

	// ===== 强化源码操作接口（当前编辑页范围）====
	// 获取当前页完整代码文本。
	bool GetCurrentPageCode(std::string& outCode) const;
	// 替换当前页全部代码文本。
	bool ReplaceCurrentPageCode(const std::string& newPageCode, bool preCompile = true) const;
	// 替换指定行范围代码（含首尾行）。
	bool ReplaceRowRangeText(int startRow, int endRow, const std::string& newText, bool preCompile = true) const;
	// 获取当前页快照（含函数分块信息）。
	bool GetCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const;
	// 按函数名获取代码（只在当前页查找）。
	bool GetFunctionCodeByName(const std::string& functionName, std::string& outCode) const;
	// 获取光标所在函数代码。
	bool GetCurrentFunctionCode(std::string& outCode, std::string* outDiagnostics = nullptr) const;
	// 按函数名替换函数代码；找不到返回 false。
	bool ReplaceFunctionCodeByName(const std::string& functionName, const std::string& newFunctionCode, bool preCompile = true) const;
	// 替换光标所在函数代码。
	bool ReplaceCurrentFunctionCode(const std::string& newFunctionCode, bool preCompile = true) const;
	// 在指定函数下方插入代码；找不到时可追加到页末。
	bool InsertCodeBelowFunction(const std::string& functionName, const std::string& codeToInsert, bool appendIfNotFound = true, bool preCompile = true) const;
	// 在当前页追加 DLL 声明代码块。
	bool InsertDllDeclaration(const std::string& dllDeclarationCode, bool preCompile = true) const;
	// 在当前页页首插入代码（优先插入到“版本”下一行）。
	bool InsertCodeAtPageTop(const std::string& codeToInsert, bool preCompile = true) const;
	// 在当前页页末插入代码。
	bool InsertCodeAtPageBottom(const std::string& codeToInsert, bool preCompile = true) const;
	// 按模板构建并插入 DLL 声明代码块。
	bool InsertDllDeclarationByTemplate(const std::string& dllName, const std::string& commandName, const std::string& returnType, const std::string& argList, bool preCompile = true) const;
	// 按函数名定位并跳转到函数头（只在当前页查找）。
	bool JumpToFunctionHeaderByName(const std::string& functionName) const;

	// 右键菜单扩展（注册、清理、消息分发）。
	void RegisterContextMenuItem(UINT commandId, const std::string& text, MenuHandler handler);
	void ClearContextMenuItems();
	bool HandleNotifyMessage(INT nMsg, DWORD dwParam1, DWORD dwParam2);
	bool HandleMainWindowCommand(WPARAM wParam);

private:
	IDEFacade() = default;
	// 内部工具：读取代码项文本。
	bool ReadProgramLikeText(INT functionCode, int rowIndex, int colIndex, ProgramText& outText) const;
	// 内部工具：构建当前页快照并拆分函数块。
	bool BuildCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const;
	// 内部工具：按名称查找函数块。
	bool FindFunctionBlockByName(const PageCodeSnapshot& snapshot, const std::string& name, FunctionBlock& outBlock) const;
	// 内部工具：按光标位置查找当前函数块。
	bool FindCurrentFunctionBlock(const PageCodeSnapshot& snapshot, FunctionBlock& outBlock) const;
	// 内部工具：基于当前光标定位函数的起止行（含首尾行）。
	bool LocateCurrentFunctionRowRange(int& outStartRow, int& outEndRow, std::string* outDiagnostics = nullptr) const;
	// 内部工具：选择行范围并替换选中代码。
	bool TranslateProgramRowRangeToBlockRange(int startProgramRow, int endProgramRow, int& outStartBlockRow, int& outEndBlockRow) const;
	bool SelectRowRange(int startRow, int endRow) const;
	bool ReplaceSelectedRowsText(const std::string& text, bool preCompile) const;
	// 内部工具：函数名规范化与文本整理。
	static std::string NormalizeFunctionName(const std::string& name);
	static std::string EnsureTrailingLineBreak(const std::string& text);
	static std::string TrimAsciiSpace(const std::string& s);
	// 内部工具：查找 IDE 输出窗口句柄（非公开控件）。
	HWND FindOutputWindowHandle() const;

	struct ContextMenuItem {
		UINT commandId;
		std::string text;
		MenuHandler handler;
	};

	std::vector<ContextMenuItem> m_contextMenuItems;
};
