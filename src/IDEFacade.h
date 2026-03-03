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
	// IDE 褰撳墠婵€娲婚〉闈㈢被鍨嬶紙瀵瑰簲 FN_GET_ACTIVE_WND_TYPE锛夈€?
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

	// 甯哥敤瑙嗗浘椤电锛堢敤浜?OpenViewTab锛夈€?
	enum class ViewTab {
		DataType,
		GlobalVar,
		DllCommand,
		ConstResource,
		PictureResource,
		SoundResource
	};

	// 浠ｇ爜鏂囨湰鏌ヨ缁撴灉锛堝搴?FN_GET_PRG_TEXT / FN_GET_PRG_HELP锛夈€?
	struct ProgramText {
		std::string text;
		int type = 0;
		bool isTitle = false;
	};

	// 瀛愮▼搴忎唬鐮佸潡淇℃伅锛堟寜褰撳墠椤垫壂鎻忓緱鍒帮級銆?
	struct FunctionBlock {
		std::string name;
		int startRow = -1;
		int endRow = -1;
		int headerCol = -1;
		std::string code;
	};

	// 褰撳墠椤典唬鐮佸揩鐓э細鏁撮〉鏂囨湰 + 瀛愮▼搴忓垪琛ㄣ€?
	struct PageCodeSnapshot {
		std::string code;
		int firstRow = -1;
		int lastRow = -1;
		std::vector<FunctionBlock> functions;
	};

	using MenuHandler = std::function<void()>;

	static IDEFacade& Instance();

	// 鐩存帴璋冪敤 IDE 鍔熻兘鍙凤紙FN_*锛夛紝杩斿洖鍘熷缁撴灉/甯冨皵缁撴灉銆?
	INT RunFunctionRaw(INT code, DWORD p1 = 0, DWORD p2 = 0) const;
	bool RunFunction(INT code, DWORD p1 = 0, DWORD p2 = 0) const;

	// 閫氱敤璋冪敤涓庣粨鏋滆緟鍔┿€?
	bool Invoke(INT fnCode, DWORD p1 = 0, DWORD p2 = 0) const;
	bool IsFnEnabled(INT fnCode) const;
	bool TryGetInt(INT fnCode, int& outValue) const;
	bool TryGetBool(INT fnCode, bool& outValue) const;

	// 鏃犲弬 FN_* 鐨勪竴閿皝瑁咃紙鐢卞畯鍒楄〃灞曞紑锛夈€?
#define IDEFACADE_DECLARE_NOARG_METHOD(methodName, fnCode) bool methodName() const;
	IDEFACADE_NOARG_FN_LIST(IDEFACADE_DECLARE_NOARG_METHOD)
#undef IDEFACADE_DECLARE_NOARG_METHOD

	// 甯哥敤甯﹀弬 FN_* 灏佽銆?
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

	// IDE 鐘舵€?鍏夋爣/鏂囨湰璇诲彇銆?
	HWND GetMainWindow() const;
	bool IsFunctionEnabled(INT code) const;
	ActiveWindowType GetActiveWindowType() const;
	bool GetCaretPosition(int& rowIndex, int& colIndex) const;
	bool GetProgramText(int rowIndex, int colIndex, ProgramText& outText) const;
	bool GetProgramHelp(int rowIndex, int colIndex, ProgramText& outText) const;

	// 缂栬緫鎿嶄綔灏佽銆?
	bool InsertText(const std::string& text, bool asKeyboardInput = false) const;
	bool SetAndCompileCurrentItemText(const std::string& text, bool preCompile = true) const;
	bool ReplaceAll(const std::string& findText, const std::string& replaceText, bool caseSensitive = false) const;
	bool SelectAll() const;
	bool CopySelection() const;
	// 鑾峰彇褰撳墠閫変腑鏂囨湰锛堥€氳繃 IDE 澶嶅埗鍚庤鍙栧壀璐存澘锛夈€?
	bool GetSelectedText(std::string& outText) const;
	bool CopyCurrentFunctionCodeToClipboard() const;

	// 浠ｇ爜瀵艰埅銆?
	bool MovePrevUnit() const;
	bool MoveNextUnit() const;
	bool MoveToParentCommand() const;
	bool MoveCaret(int rowIndex, int colIndex) const;
	bool MoveToReferencedSub() const;
	bool OpenCurrentSub() const;
	bool CloseCurrentSub() const;
	bool MoveBackSub() const;

	// 椤甸潰鍒囨崲涓庣紪璇戜繚瀛樸€?
	bool OpenViewTab(ViewTab tab) const;
	bool SaveFile() const;
	bool OpenFile(const std::string& filePath) const;
	bool Compile() const;
	bool CompileAndRun() const;
	bool AddOutputTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon = nullptr) const;
	// 鍚?IDE 杈撳嚭绐楀彛杩藉姞鏂囨湰/琛屾枃鏈紙琛屾枃鏈細鑷姩杩藉姞 CRLF锛夈€?
	bool AppendOutputWindowText(const std::string& text) const;
	bool AppendOutputWindowLine(const std::string& text) const;

	// ECOM 绠＄悊銆?
	bool AddECOM(const std::string& filePath) const;
	bool AddECOM2(const std::string& filePath) const;
	bool GetImportedECOMCount(int& count) const;
	bool GetImportedECOMPath(int index, std::string& path) const;
	bool InputECOM(const std::string& filePath, bool useNewAddMethod = false) const;
	bool RemoveECOM(int index) const;
	bool RemoveECOM(const std::string& filePath) const;
	int FindECOMIndex(const std::string& filePath) const;
	int FindECOMNameIndex(const std::string& ecomName) const;

	// ===== 寮哄寲婧愮爜鎿嶄綔鎺ュ彛锛堝綋鍓嶇紪杈戦〉鑼冨洿锛?====
	// 鑾峰彇褰撳墠椤靛畬鏁翠唬鐮佹枃鏈€?
	bool GetCurrentPageCode(std::string& outCode) const;
	// 替换当前页全部代码文本。
	bool ReplaceCurrentPageCode(const std::string& newPageCode, bool preCompile = true) const;
	// 替换指定行范围代码（含首尾行）。
	bool ReplaceRowRangeText(int startRow, int endRow, const std::string& newText, bool preCompile = true) const;
	// 获取当前页快照（含函数分块信息）。
	bool GetCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const;
	// 鎸夊嚱鏁板悕鑾峰彇浠ｇ爜锛堝彧鍦ㄥ綋鍓嶉〉鏌ユ壘锛夈€?
	bool GetFunctionCodeByName(const std::string& functionName, std::string& outCode) const;
	// 鑾峰彇鍏夋爣鎵€鍦ㄥ嚱鏁颁唬鐮併€?
	bool GetCurrentFunctionCode(std::string& outCode, std::string* outDiagnostics = nullptr) const;
	// 鎸夊嚱鏁板悕鏇挎崲鍑芥暟浠ｇ爜锛涙壘涓嶅埌杩斿洖 false銆?
	bool ReplaceFunctionCodeByName(const std::string& functionName, const std::string& newFunctionCode, bool preCompile = true) const;
	// 鏇挎崲鍏夋爣鎵€鍦ㄥ嚱鏁颁唬鐮併€?
	bool ReplaceCurrentFunctionCode(const std::string& newFunctionCode, bool preCompile = true) const;
	// 鍦ㄦ寚瀹氬嚱鏁颁笅鏂规彃鍏ヤ唬鐮侊紱鎵句笉鍒版椂鍙拷鍔犲埌椤垫湯銆?
	bool InsertCodeBelowFunction(const std::string& functionName, const std::string& codeToInsert, bool appendIfNotFound = true, bool preCompile = true) const;
	// 鍦ㄥ綋鍓嶉〉杩藉姞 DLL 澹版槑浠ｇ爜鍧椼€?
	bool InsertDllDeclaration(const std::string& dllDeclarationCode, bool preCompile = true) const;
	// 鍦ㄥ綋鍓嶉〉椤甸鎻掑叆浠ｇ爜锛堜紭鍏堟彃鍏ュ埌鈥?鐗堟湰鈥濅笅涓€琛岋級銆?
	bool InsertCodeAtPageTop(const std::string& codeToInsert, bool preCompile = true) const;
	// 鎸夋ā鏉挎瀯寤哄苟鎻掑叆 DLL 澹版槑浠ｇ爜鍧椼€?
	bool InsertDllDeclarationByTemplate(const std::string& dllName, const std::string& commandName, const std::string& returnType, const std::string& argList, bool preCompile = true) const;
	// 鎸夊嚱鏁板悕瀹氫綅骞惰烦杞埌鍑芥暟澶达紙鍙湪褰撳墠椤垫煡鎵撅級銆?
	bool JumpToFunctionHeaderByName(const std::string& functionName) const;

	// 鍙抽敭鑿滃崟鎵╁睍锛堟敞鍐屻€佹竻鐞嗐€佹秷鎭垎鍙戯級銆?
	void RegisterContextMenuItem(UINT commandId, const std::string& text, MenuHandler handler);
	void ClearContextMenuItems();
	bool HandleNotifyMessage(INT nMsg, DWORD dwParam1, DWORD dwParam2);
	bool HandleMainWindowCommand(WPARAM wParam);

private:
	IDEFacade() = default;
	// 鍐呴儴宸ュ叿锛氳鍙栦唬鐮侀」鏂囨湰銆?
	bool ReadProgramLikeText(INT functionCode, int rowIndex, int colIndex, ProgramText& outText) const;
	// 鍐呴儴宸ュ叿锛氭瀯寤哄綋鍓嶉〉蹇収骞舵媶鍒嗗嚱鏁板潡銆?
	bool BuildCurrentPageSnapshot(PageCodeSnapshot& outSnapshot) const;
	// 鍐呴儴宸ュ叿锛氭寜鍚嶇О鏌ユ壘鍑芥暟鍧椼€?
	bool FindFunctionBlockByName(const PageCodeSnapshot& snapshot, const std::string& name, FunctionBlock& outBlock) const;
	// 鍐呴儴宸ュ叿锛氭寜鍏夋爣浣嶇疆鏌ユ壘褰撳墠鍑芥暟鍧椼€?
	bool FindCurrentFunctionBlock(const PageCodeSnapshot& snapshot, FunctionBlock& outBlock) const;
	// 鍐呴儴宸ュ叿锛氶€夋嫨琛岃寖鍥村苟鏇挎崲閫変腑浠ｇ爜銆?
	bool SelectRowRange(int startRow, int endRow) const;
	bool ReplaceSelectedRowsText(const std::string& text, bool preCompile) const;
	// 鍐呴儴宸ュ叿锛氬嚱鏁板悕瑙勮寖鍖栦笌鏂囨湰鏁寸悊銆?
	static std::string NormalizeFunctionName(const std::string& name);
	static std::string EnsureTrailingLineBreak(const std::string& text);
	static std::string TrimAsciiSpace(const std::string& s);
	// 鍐呴儴宸ュ叿锛氭煡鎵?IDE 杈撳嚭绐楀彛鍙ユ焺锛堥潪鍏紑鎺т欢锛夈€?
	HWND FindOutputWindowHandle() const;

	struct ContextMenuItem {
		UINT commandId;
		std::string text;
		MenuHandler handler;
	};

	std::vector<ContextMenuItem> m_contextMenuItems;
};
