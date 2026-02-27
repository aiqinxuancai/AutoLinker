#pragma once

#include <Windows.h>
#include <PublicIDEFunctions.h>
#include <functional>
#include <string>
#include <vector>

class IDEFacade {
public:
	// 当前活动编辑页类型（对应 FN_GET_ACTIVE_WND_TYPE 的返回值）。
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

	// 可由 OpenViewTab 打开的常用编辑页类型。
	enum class ViewTab {
		DataType,
		GlobalVar,
		DllCommand,
		ConstResource,
		PictureResource,
		SoundResource
	};

	// GetProgramText / GetProgramHelp 的返回载体。
	struct ProgramText {
		std::string text;
		int type = 0;
		bool isTitle = false;
	};

	using MenuHandler = std::function<void()>;

	static IDEFacade& Instance();

	// 调用任意 FN_* 功能号，返回 IDE 原始返回值。
	INT RunFunctionRaw(INT code, DWORD p1 = 0, DWORD p2 = 0) const;
	// 调用任意 FN_* 功能号，并将结果转换为 bool。
	bool RunFunction(INT code, DWORD p1 = 0, DWORD p2 = 0) const;

	// 获取 IDE 主窗口句柄（NES_GET_MAIN_HWND）。
	HWND GetMainWindow() const;
	// 检查某个 FN_* 功能当前是否可用。
	bool IsFunctionEnabled(INT code) const;
	// 获取当前活动编辑页类型。
	ActiveWindowType GetActiveWindowType() const;
	// 获取当前光标行列。
	bool GetCaretPosition(int& rowIndex, int& colIndex) const;
	// 读取指定行列处的程序文本，row/col 传 -1 表示当前光标位置。
	bool GetProgramText(int rowIndex, int colIndex, ProgramText& outText) const;
	// 读取指定行列处的帮助文本，row/col 传 -1 表示当前光标位置。
	bool GetProgramHelp(int rowIndex, int colIndex, ProgramText& outText) const;

	// 在当前光标处插入文本。
	bool InsertText(const std::string& text, bool asKeyboardInput = false) const;
	// 设置当前项文本，并可选择立即预编译。
	bool SetAndCompileCurrentItemText(const std::string& text, bool preCompile = true) const;
	// 在当前编辑窗口执行全量替换。
	bool ReplaceAll(const std::string& findText, const std::string& replaceText, bool caseSensitive = false) const;
	// 全选当前代码视图内容。
	bool SelectAll() const;
	// 复制当前选中内容到系统剪贴板。
	bool CopySelection() const;
	// 复制光标所在函数（子程序）代码到系统剪贴板。
	bool CopyCurrentFunctionCodeToClipboard() const;

	// 代码结构导航能力。
	bool MovePrevUnit() const;
	bool MoveNextUnit() const;
	bool MoveToParentCommand() const;
	bool MoveCaret(int rowIndex, int colIndex) const;
	bool MoveToReferencedSub() const;
	bool OpenCurrentSub() const;
	bool CloseCurrentSub() const;
	bool MoveBackSub() const;

	// 打开常用编辑页，以及文件/编译操作。
	bool OpenViewTab(ViewTab tab) const;
	bool SaveFile() const;
	bool OpenFile(const std::string& filePath) const;
	bool Compile() const;
	bool CompileAndRun() const;

	// 在 IDE 输出栏添加自定义页签。
	bool AddOutputTab(HWND hWnd, const std::string& caption, const std::string& toolTip, HICON hIcon = nullptr) const;

	// ECOM 导入与查询能力。
	// 以传统导入方式添加 ECOM（对应 FN_INPUT_ECOM）。
	bool AddECOM(const std::string& filePath) const;
	// 以新导入方式添加 ECOM（对应 FN_ADD_NEW_ECOM2）。
	bool AddECOM2(const std::string& filePath) const;
	bool GetImportedECOMCount(int& count) const;
	bool GetImportedECOMPath(int index, std::string& path) const;
	bool InputECOM(const std::string& filePath, bool useNewAddMethod = false) const;
	bool RemoveECOM(int index) const;
	// 按路径删除 ECOM。
	bool RemoveECOM(const std::string& filePath) const;
	// 按完整路径或按模块名查找 ECOM 索引，找不到返回 -1。
	int FindECOMIndex(const std::string& filePath) const;
	int FindECOMNameIndex(const std::string& ecomName) const;

	// 右键菜单注册与消息分发。
	void RegisterContextMenuItem(UINT commandId, const std::string& text, MenuHandler handler);
	void ClearContextMenuItems();
	bool HandleNotifyMessage(INT nMsg, DWORD dwParam1, DWORD dwParam2);
	bool HandleMainWindowCommand(WPARAM wParam);

private:
	IDEFacade() = default;
	bool ReadProgramLikeText(INT functionCode, int rowIndex, int colIndex, ProgramText& outText) const;

	struct ContextMenuItem {
		UINT commandId;
		std::string text;
		MenuHandler handler;
	};

	std::vector<ContextMenuItem> m_contextMenuItems;
};
