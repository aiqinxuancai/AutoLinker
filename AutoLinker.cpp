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


//在窗口上添加的Button句柄
HWND _button;

//当前打开的易源文件路径
std::string _nowOpenSourceFilePath;

//配置文件，当前路径的e源码对应的编译器名
ConfigManager _configManager;

//管理当前所有的link文件
LinkerManager _linkerManager;

//e主窗口句柄
HWND g_hwnd = NULL;

HWND g_toolBarHwnd = NULL;

HANDLE hMainThread = NULL;

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
		auto linkName = _configManager.getValue(_nowOpenSourceFilePath);

		if (!linkName.empty() ) {
			auto linkConfig = _linkerManager.getConfig(linkName);

			if (std::filesystem::exists(linkConfig.path)) {
				//切换路径
				OutputStringToELog("[AutoLinker]切换为Linker:" + linkConfig.name + " " + linkConfig.path);

				return originalCreateFileA(linkConfig.path.c_str(), dwDesiredAccess, dwShareMode, lpSecurityAttributes,
					dwCreationDisposition,
					dwFlagsAndAttributes,
					hTemplateFile);
			}
			else {
				OutputStringToELog("[AutoLinker]无法切换Linker，link文件不存在#1");
			}

		}
		else {
			OutputStringToELog("[AutoLinker]无法切换Linker，link文件不存在#2");
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

void ShowMenu(HWND hParent) {
	// 创建一个弹出菜单

	if (_linkerManager.getCount() > 0) {
		HMENU hPopupMenu = CreatePopupMenu();

		auto nowLinkConfigName = _configManager.getValue(_nowOpenSourceFilePath);


		//获取当前的所有配置
		for (const auto& [key, value] : _linkerManager.getMap()) {
			if (!nowLinkConfigName.empty() && nowLinkConfigName == key) {
				AppendMenu(hPopupMenu, MF_STRING | MF_ENABLED | MF_CHECKED, value.id, key.c_str());
			}
			else {
				AppendMenu(hPopupMenu, MF_STRING | MF_ENABLED, value.id, key.c_str());
			}
		}

		// 获取鼠标的位置
		POINT pt;
		GetCursorPos(&pt);

		// 显示弹出菜单
		TrackPopupMenu(hPopupMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hParent, NULL);

		// 删除菜单
		DestroyMenu(hPopupMenu);
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

	OutputStringToELog("完整路径" + title);

	// 使用迭代器找到所有匹配项
	for (std::sregex_iterator it = std::sregex_iterator(title.begin(), title.end(), path_regex);
		it != std::sregex_iterator(); ++it) {

		if (!it->str().empty()) {
			OutputStringToELog(it->str());
			return it->str();
		}
		
	}
	return "";

}

void UpdateButton() {
	auto linkName = _configManager.getValue(_nowOpenSourceFilePath);
	if (!linkName.empty()) {
		SetWindowTextA(_button, linkName.c_str());
	}
	else {
		SetWindowTextA(_button, "默认");
	}
}

/// <summary>
/// 更新当前源文件使用的link文件
/// </summary>
/// <param name="id"></param>
void UpdateCurrentFileLinkerWithId(int id) {
	_configManager.setValue(_nowOpenSourceFilePath, _linkerManager.getConfig(id).name);

	//更新当前的按钮和
	UpdateButton();

}
void CreateAndSubclassButton(HWND hParent) {

	int buttonWidth = 80;
	int buttonHeight = 20;


	//获取当前的源文件路径

	HWND hButton = CreateWindow(
		"BUTTON",  // 预定义的按钮类名
		"AutoLinker", // 按钮文本
		WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, // 按钮样式
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

	_button = hButton;

}

//工具条子类过程
LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	if (uMsg == WM_COMMAND && LOWORD(wParam) == BN_CLICKED) {
		OutputStringToELog("弹出菜单，选择当前的ini配置");
		//HWND hWnd = (HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0);
		ShowMenu(hWnd);
		return 0;
	}

	std::string s = std::format("菜单条消息 {0} {1}", (int)hWnd, uMsg);
	OutputStringToELog(s);


	switch (uMsg)
	{
	case 10601: {
		OutputStringToELog("主窗口需要初始化");
		break;
	}
	case WM_PAINT: {
		//获取当前的文件

		std::string sourceFile = GetSourceFilePath(hWnd);
		//OutputStringToELog(sourceFile);
		_nowOpenSourceFilePath = sourceFile;
		//TODO 更新按钮文本
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

bool FneInitThread() {
	OutputStringToELog("开始初始化");
	DWORD processID = GetCurrentProcessId();
	EnumWindows(EnumWindowsProc, processID);
	//HWND hWnd = (HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0);
	//std::string s = std::format("{} {} {}", processID, (int)g_hwnd, (int)hWnd);
	//OutputStringToELog(s);

	g_toolBarHwnd = FindChildWindowByTitle(g_hwnd);

	if (g_hwnd != NULL && g_toolBarHwnd != NULL)
	{
		OutputStringToELog("查找到菜单条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		if (g_toolBarHwnd) {
			std::string s = std::format("菜单条句柄{0}", (int)g_toolBarHwnd);
			OutputStringToELog(s);
			CreateAndSubclassButton(g_toolBarHwnd);
			//Hook读文件的函数，修改link.ini的路径
			StartHookCreateFileA();
			std::string sourceFile = GetSourceFilePath(g_hwnd);
			OutputStringToELog(sourceFile);

		}
		return true;
	}
	else
	{
		
	}
	
	return false;
}

/*-----------------支持库消息处理函数------------------*/

EXTERN_C INT WINAPI AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	std::string s = std::format("AutoLinker_MessageNotify {0} {1} {2}", (int)nMsg, dwParam1, dwParam2);
	OutputStringToELog(s);

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
			if (!_button) {
				//窗口已经建立
				FneInitThread();
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
	//获取当前进程的主窗口
	//如果没有初始化，在这里初始化
	//std::thread t([]() {
	//	while (true) {
	//		if (!_button) {
	//			if (FneInitThread()) {
	//				OutputStringToELog("AutoLinker Init");
	//				break;
	//			}
	//		}
	//		std::this_thread::sleep_for(std::chrono::milliseconds(500));
	//	}
	//	});

	//t.detach();
	
	return (&LibInfo);
};
#endif