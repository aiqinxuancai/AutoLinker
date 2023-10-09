#include "AutoLinker.h"
#include <vector>
#include <Windows.h>
#include <CommCtrl.h>

#include <format>
#include "ConfigManager.h"
#include <regex>
#include "PathHelper.h"
#include "LinkerManager.h"

#pragma comment(lib, "comctl32.lib")

//当前打开的易源文件路径
std::string g_nowOpenSourceFilePath;

//配置文件，当前路径的e源码对应的编译器名
ConfigManager g_configManager;

//管理当前所有的link文件
LinkerManager g_linkerManager;

//e主窗口句柄
HWND g_hwnd = NULL;

//工具条句柄
HWND g_toolBarHwnd = NULL;

//在窗口上添加的Button句柄
HWND g_buttonHwnd;

static auto originalCreateFileA = CreateFileA;

void OutputStringToELog(std::string szbuf);

HANDLE WINAPI MyCreateFileA(
	LPCSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDisposition,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile


) {
	//OutputStringToELog("MyCreateFileA");
	std::filesystem::path currentPath = GetBasePath();
	std::filesystem::path autoLoaderPath = currentPath / "tools" / "link.ini";
	if (autoLoaderPath.string() == std::string(lpFileName)) {
		auto linkName = g_configManager.getValue(g_nowOpenSourceFilePath);

		if (!linkName.empty() ) {
			auto linkConfig = g_linkerManager.getConfig(linkName);

			if (std::filesystem::exists(linkConfig.path)) {
				//切换路径
				OutputStringToELog("切换为Linker:" + linkConfig.name + " " + linkConfig.path);

				return originalCreateFileA(linkConfig.path.c_str(), dwDesiredAccess, dwShareMode, lpSecurityAttributes,
					dwCreationDisposition,
					dwFlagsAndAttributes,
					hTemplateFile);
			}
			else {
				OutputStringToELog("无法切换Linker，Linker文件不存在#1");
			}

		}
		else {
			OutputStringToELog("未设置此源文件的Linker，使用默认");
		}
	}


	return originalCreateFileA(lpFileName,dwDesiredAccess,dwShareMode,lpSecurityAttributes,
								dwCreationDisposition,
								dwFlagsAndAttributes,
								hTemplateFile);
}

void StartHookCreateFileA() {
	DetourRestoreAfterWith();
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)originalCreateFileA, MyCreateFileA);
	DetourTransactionCommit();
}


BOOL CALLBACK EnumChildProcOutputWindow(HWND hwnd, LPARAM lParam) {
	char buffer[256] = { 0 };
	if (GetDlgCtrlID(hwnd) == 1011) {
		HWND* pResult = reinterpret_cast<HWND*>(lParam);
		*pResult = hwnd;
		return FALSE;
	}
	return TRUE;
}

/// <summary>
/// 查找输出窗口
/// </summary>
/// <param name="hParent"></param>
/// <returns></returns>
HWND FindOutputWindow(HWND hParent) {
	HWND hResult = NULL;
	EnumChildWindows(hParent, EnumChildProcOutputWindow, reinterpret_cast<LPARAM>(&hResult));
	return hResult;
}

/// <summary>
/// 输出文本
/// </summary>
/// <param name="szbuf"></param>
void OutputStringToELog(std::string szbuf) {
	szbuf.insert(0, "[AutoLinker]");

	OutputDebugString(szbuf.c_str());
	HWND hwnd = FindOutputWindow((HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0));
	if (hwnd) {
		SendMessageA(hwnd, 194, 1, (LPARAM)szbuf.c_str());
		SendMessageA(hwnd, 194, 1, (long)"\r\n");
	}
	else {
		//::MessageBox(0, "没有找到", "辅助工具", MB_YESNO);
	}
}

BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam) {
	char buffer[256] = { 0 };
	GetWindowText(hwnd, buffer, sizeof(buffer));
	char bufferClassName[256] = { 0 };
	GetClassName(hwnd, bufferClassName, sizeof(bufferClassName));

	if (std::string(buffer) == "菜单条" && std::string(bufferClassName) == "Afx:400000:b:10003:10:0") {
		HWND* pResult = reinterpret_cast<HWND*>(lParam);
		*pResult = hwnd;
		return FALSE;
	}

	// 继续枚举
	return TRUE;
}

/// <summary>
/// 查找菜单条
/// </summary>
/// <param name="hParent"></param>
/// <returns></returns>
HWND FindChildWindowByTitle(HWND hParent) {
	HWND hResult = NULL;
	EnumChildWindows(hParent, EnumChildProc, reinterpret_cast<LPARAM>(&hResult));
	return hResult;
}

/// <summary>
/// 弹出菜单
/// </summary>
/// <param name="hParent"></param>
void ShowMenu(HWND hParent) {
	if (g_linkerManager.getCount() > 0) {
		HMENU hPopupMenu = CreatePopupMenu();

		auto nowLinkConfigName = g_configManager.getValue(g_nowOpenSourceFilePath);
		//获取当前的所有配置
		for (const auto& [key, value] : g_linkerManager.getMap()) {
			if (!nowLinkConfigName.empty() && nowLinkConfigName == key) {
				AppendMenu(hPopupMenu, MF_STRING | MF_ENABLED | MF_CHECKED, value.id, key.c_str());
			}
			else {
				AppendMenu(hPopupMenu, MF_STRING | MF_ENABLED, value.id, key.c_str());
			}
		}

		POINT pt;
		GetCursorPos(&pt);
		TrackPopupMenu(hPopupMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hParent, NULL);
		DestroyMenu(hPopupMenu);
	}
	else {
		OutputStringToELog("当前没有Linker配置文件，请在e目录的AutoLinker\\Config中添加ini文件");
	}

	
}

/// <summary>
/// 获取当前源文件的路径
/// </summary>
/// <param name="hParent"></param>
std::string GetSourceFilePath(HWND hParent) {
	HWND hWnd = (HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0);

	char buffer[256] = { 0 };
	GetWindowText(hWnd, buffer, sizeof(buffer));
	std::regex path_regex(R"((([A-Za-z]:|\\)\\[^ /:*?"<>|\r\n]+\\?)*)");
	std::smatch match;
	auto title = std::string(buffer);

	//OutputStringToELog("完整路径" + title);

	// 使用迭代器找到所有匹配项
	for (std::sregex_iterator it = std::sregex_iterator(title.begin(), title.end(), path_regex);
		it != std::sregex_iterator(); ++it) {

		if (!it->str().empty()) {
			//OutputStringToELog(it->str());
			return it->str();
		}
		
	}
	return "";

}

void UpdateButton() {
	auto linkName = g_configManager.getValue(g_nowOpenSourceFilePath);
	if (!linkName.empty()) {
		SetWindowTextA(g_buttonHwnd, linkName.c_str());
	}
	else {
		SetWindowTextA(g_buttonHwnd, "默认");
	}
}

/// <summary>
/// 更新当前源文件使用的link文件
/// </summary>
/// <param name="id"></param>
void UpdateCurrentFileLinkerWithId(int id) {
	g_configManager.setValue(g_nowOpenSourceFilePath, g_linkerManager.getConfig(id).name);
	UpdateButton();
}

/// <summary>
/// 创建按钮
/// TODO 和其他插件UI覆盖冲突的问题？
/// </summary>
/// <param name="hParent"></param>
void CreateAndSubclassButton(HWND hParent) {

	int buttonWidth = 90;
	int buttonHeight = 20;


	//获取当前的源文件路径

	HWND hButton = CreateWindow(
		"BUTTON",  // 预定义的按钮类名
		"AutoLinker", // 按钮文本
		WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON | BS_FLAT, // 按钮样式
		500, // x位置
		5, // y位置
		buttonWidth, // 按钮宽度
		buttonHeight, // 按钮高度
		hParent, // 父窗口句柄
		0, // 没有菜单
		(HINSTANCE)GetWindowLong(hParent, GWL_HINSTANCE), // 程序实例句柄
		NULL); // 无附加参数

	//设置字体
	int fontSize = 10;
	int dpi = GetDeviceCaps(GetDC(NULL), LOGPIXELSY);
	int fontHeight = -MulDiv(fontSize, dpi, 72);
	HFONT hFont = CreateFont(fontHeight, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, TEXT("Microsoft YaHei UI"));
	SendMessage(hButton, WM_SETFONT, (WPARAM)hFont, TRUE);

	g_buttonHwnd = hButton;

}

//工具条子类过程
LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	if (uMsg == WM_COMMAND && LOWORD(wParam) == BN_CLICKED) {
		//OutputStringToELog("弹出菜单，选择当前的ini配置");
		//HWND hWnd = (HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0);
		ShowMenu(hWnd);
		return 0;
	}

	//std::string s = std::format("菜单条消息 {0} {1}", (int)hWnd, uMsg);
	//OutputStringToELog(s);


	switch (uMsg)
	{
	case WM_PAINT: {
		//获取当前的文件

		std::string sourceFile = GetSourceFilePath(hWnd);
		//OutputStringToELog(sourceFile);
		g_nowOpenSourceFilePath = sourceFile;
		//TODO 更新按钮文本
		UpdateButton();


		break;
	}

	case WM_INITMENUPOPUP:
		//if (!HIWORD(lParam))
		//{
		//	HMENU hMenu = (HMENU)wParam;
		//	EnableMenuItem(hMenu, ID_MENU_OPTION1, MF_BYCOMMAND | MF_ENABLED);
		//}
		break;
	case WM_COMMAND: {
		auto low = LOWORD(wParam);
		UpdateCurrentFileLinkerWithId(low);
		break;
	}

	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}




INT WINAPI fnAddInFunc(INT nAddInFnIndex) {
	if (nAddInFnIndex == 0) { 

	}
	return 0;
}



BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam)
{
	DWORD lpdwProcessId;
	GetWindowThreadProcessId(hwnd, &lpdwProcessId);
	if (lpdwProcessId == lParam)
	{
		HWND hwndTopLevel = GetAncestor(hwnd, GA_ROOTOWNER);
		g_hwnd = hwndTopLevel;
		return FALSE;
	}
	return TRUE;
}

bool FneInit() {
	OutputStringToELog("开始初始化");
	DWORD processID = GetCurrentProcessId();
	EnumWindows(EnumWindowsProc, processID);

	//此时NotifySys还不可用
	//HWND hWnd = (HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0);


	g_toolBarHwnd = FindChildWindowByTitle(g_hwnd);


	std::string s = std::format("{} {} {}", processID, (int)g_hwnd, (int)g_toolBarHwnd);
	OutputStringToELog(s);

	if (g_hwnd != NULL && g_toolBarHwnd != NULL)
	{
		OutputStringToELog("找到工具条");
		//OutputStringToELog("查找到菜单条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		//std::string s = std::format("菜单条句柄{0}", (int)g_toolBarHwnd);
		//OutputStringToELog(s);
		CreateAndSubclassButton(g_toolBarHwnd);
		//Hook读文件的函数，修改link.ini的路径
		StartHookCreateFileA();
		//std::string sourceFile = GetSourceFilePath(g_hwnd);
		//OutputStringToELog(sourceFile);

		OutputStringToELog("初始化完成");
		return true;
	}
	else
	{
		OutputStringToELog("初始化失败，未找到窗口");



	}
	
	return false;
}

/*-----------------支持库消息处理函数------------------*/

UINT_PTR timerId;

VOID CALLBACK AsyncFneInit(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime) {
	if (!FneInit()) {
		timerId = SetTimer(g_hwnd, 0, 1000, &AsyncFneInit);
	}
	KillTimer(hwnd, idEvent);  // Stop the timer
}

EXTERN_C INT WINAPI AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	//std::string s = std::format("AutoLinker_MessageNotify {0} {1} {2}", (int)nMsg, dwParam1, dwParam2);
	//OutputStringToELog(s);

#ifndef __E_STATIC_LIB
	if (nMsg == NL_GET_CMD_FUNC_NAMES) // 返回所有命令实现函数的的函数名称数组(char*[]), 支持静态编译的动态库必须处理
		return NULL;
	else if (nMsg == NL_GET_NOTIFY_LIB_FUNC_NAME) // 返回处理系统通知的函数名称(PFN_NOTIFY_LIB函数名称), 支持静态编译的动态库必须处理
		return (INT)LIBARAYNAME;
	else if (nMsg == NL_GET_DEPENDENT_LIBS) return (INT)NULL;
	//else if (nMsg == NL_GET_DEPENDENT_LIBS) return (INT)NULL;
	// 返回静态库所依赖的其它静态库文件名列表(格式为\0分隔的文本,结尾两个\0), 支持静态编译的动态库必须处理
	// kernel32.lib user32.lib gdi32.lib 等常用的系统库不需要放在此列表中
	// 返回NULL或NR_ERR表示不指定依赖文件  

	else if (nMsg == NL_SYS_NOTIFY_FUNCTION) {
		if (dwParam1) {
			if (!g_buttonHwnd) {
				//窗口已经建立
				if (!FneInit()) {
					DWORD processID = GetCurrentProcessId();
					EnumWindows(EnumWindowsProc, processID);
					timerId = SetTimer(g_hwnd, 0, 1000, &AsyncFneInit);
				}
			}

		}
	}

#endif
	return ProcessNotifyLib(nMsg, dwParam1, dwParam2);
}
/*定义支持库基本信息*/
#ifndef __E_STATIC_LIB
static LIB_INFO LibInfo =
{
	/* { 库格式号, GUID串号, 主版本号, 次版本号, 构建版本号, 系统主版本号, 系统次版本号, 核心库主版本号, 核心库次版本号,
	支持库名, 支持库语言, 支持库描述, 支持库状态,
	作者姓名, 邮政编码, 通信地址, 电话号码, 传真号码, 电子邮箱, 主页地址, 其它信息,
	类型数量, 类型指针, 类别数量, 命令类别, 命令总数, 命令指针, 命令入口,
	附加功能, 功能描述, 消息指针, 超级模板, 模板描述,
	常量数量, 常量指针, 外部文件} */
	LIB_FORMAT_VER,
	_T(LIB_GUID_STR),
	LIB_MajorVersion,
	LIB_MinorVersion,
	LIB_BuildNumber,
	LIB_SysMajorVer,
	LIB_SysMinorVer,
	LIB_KrnlLibMajorVer,
	LIB_KrnlLibMinorVer,
	_T(LIB_NAME_STR),
	__GBK_LANG_VER,
	_WT(LIB_DESCRIPTION_STR),
	_LIB_OS(__OS_WIN), //#LBS_IDE_PLUGIN  LBS_LIB_INFO2
	_WT(LIB_Author),
	_WT(LIB_ZipCode),
	_WT(LIB_Address),
	_WT(LIB_Phone), 
	_WT(LIB_Fax),
	_WT(LIB_Email),
	_WT(LIB_HomePage),
	_WT(LIB_Other),
	0,
	NULL,
	LIB_TYPE_COUNT,
	_WT(LIB_TYPE_STR),
	0,
	NULL,
	NULL,
	fnAddInFunc,
	_T("辅助功能1\0这是个用作测试的辅助工具功能。\0辅助功能2\0这是个用作测试的辅助工具功能。\0\0") ,
	AutoLinker_MessageNotify,
	NULL,
	NULL,
	0,
	NULL,
	NULL
};

PLIB_INFO WINAPI GetNewInf()
{
	return (&LibInfo);
};
#endif