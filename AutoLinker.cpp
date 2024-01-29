#include "AutoLinker.h"
#include <vector>
#include <Windows.h>
#include <CommCtrl.h>
#include <format>
#include "ConfigManager.h"
#include <regex>
#include "PathHelper.h"
#include "LinkerManager.h"
#include "InlineHook.h"
#include "ModelManager.h"
#include "Global.h"
#include "StringHelper.h"
#include <thread>
#include "MouseBack.h"
#include <PublicIDEFunctions.h>
#include "ECOMEx.h"

#pragma comment(lib, "comctl32.lib")

//当前打开的易源文件路径
std::string g_nowOpenSourceFilePath;

//配置文件，当前路径的e源码对应的编译器名
ConfigManager g_configManager;

//管理当前所有的link文件
LinkerManager g_linkerManager;

//管理模块调试版及编译版本的管理器
ModelManager g_modelManager;

//e主窗口句柄
HWND g_hwnd = NULL;

//工具条句柄
HWND g_toolBarHwnd = NULL;

//在窗口上添加的Button句柄
HWND g_buttonHwnd;

//准备开始调试
bool g_preDebugging;

//准备开始编译
bool g_preCompiling;


static auto originalCreateFileA = CreateFileA;
static auto originalGetSaveFileNameA = GetSaveFileNameA;
static auto originalCreateProcessA = CreateProcessA;
static auto originalMessageBoxA = MessageBoxA;


typedef int(__thiscall* OriginalEStartBuildFuncType)(DWORD* thisPtr, int a2);
OriginalEStartBuildFuncType originalEStartBuildFunc = (OriginalEStartBuildFuncType)0x40A9F1;  //int __thiscall sub_40A9F1(_DWORD *this, int a2)


typedef int(__thiscall* OriginalEStartDebugFuncType)(DWORD* thisPtr, int a2, int a3); 
OriginalEStartDebugFuncType originalEStartDebugFunc = (OriginalEStartDebugFuncType)0x40A080; //int __thiscall sub_40A080(int this, int a2, int a3)


//typedef INT(WINAPI* NotifySys3Func)(INT, DWORD, DWORD);
//static NotifySys3Func NotifySys3 ;

int __fastcall MyEStartBuildFunc(DWORD* thisPtr, int dummy, int a2) {
	OutputStringToELog("编译开始#2");
	ChangeVMPModel(true);
	return originalEStartBuildFunc(thisPtr, a2);
}

int __fastcall MyEStartDebugFunc(DWORD* thisPtr, int dummy, int a2, int a3) {
	OutputStringToELog("调试开始#2");
	ChangeVMPModel(false);
	return originalEStartDebugFunc(thisPtr, a2, a3);
}

HANDLE WINAPI MyCreateFileA(
	LPCSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDisposition,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile) {
	if (std::string(lpFileName).find("\\Temp\\e_debug\\") != std::string::npos) {
		OutputStringToELog("结束预编译代码（调试）");
		g_preDebugging = false;
	}

	if (g_preDebugging) {
		
	}

	std::filesystem::path currentPath = GetBasePath();
	std::filesystem::path autoLoaderPath = currentPath / "tools" / "link.ini";
	if (autoLoaderPath.string() == std::string(lpFileName)) {
		g_preCompiling = false;

		//EC编译阶段结束
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

BOOL APIENTRY MyGetSaveFileNameA(LPOPENFILENAMEA item) {
	if (g_preCompiling) {
		OutputStringToELog("结束预编译代码");
		g_preCompiling = false;
	}

	return originalGetSaveFileNameA(item);
}


BOOL WINAPI MyCreateProcessA(
	LPCSTR lpApplicationName,
	LPSTR lpCommandLine,
	LPSECURITY_ATTRIBUTES lpProcessAttributes,
	LPSECURITY_ATTRIBUTES lpThreadAttributes,
	BOOL bInheritHandles,
	DWORD dwCreationFlags,
	LPVOID lpEnvironment,
	LPCSTR lpCurrentDirectory,
	LPSTARTUPINFOA lpStartupInfo,
	LPPROCESS_INFORMATION lpProcessInformation
) {
	std::string commandLine = lpCommandLine;
	auto outFileName = GetLinkerCommandOutFileName(lpCommandLine);

	//std::string s = std::format("进程创建 [{}] {}", lpApplicationName, lpCommandLine);
	//OutputStringToELog(s);

	if (!outFileName.empty()) {
		if (commandLine.find("/pdb:\"build.pdb\"") != std::string::npos) {
			//PDB更名为当前编译的程序的名字+.pdb，看起来更正规

			std::string newPdbCommand = std::format("/pdb:\"{}.pdb\"", outFileName);
			commandLine = replaceSubstring(commandLine, "/pdb:\"build.pdb\"", newPdbCommand);

			std::vector<char> commandLineBuffer(commandLine.begin(), commandLine.end());
			commandLineBuffer.push_back('\0');

			return originalCreateProcessA(lpApplicationName, commandLineBuffer.data(), lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
		}
	}

	return originalCreateProcessA(lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
}

int WINAPI MyMessageBoxA(
	HWND hWnd,
	LPCSTR lpText,
	LPCSTR lpCaption,
	UINT uType) {
	//TODO 自动返回确认编译

	if (std::string(lpCaption).find("链接器输出了大量错误或警告信息") != std::string::npos) {
		return IDNO;
	}
	if (std::string(lpText).find("链接器输出了大量错误或警告信息") != std::string::npos) {
		return IDNO;
	}
	return originalMessageBoxA(hWnd, lpText, lpCaption, uType);
}

void StartHookCreateFileA() {
	DetourRestoreAfterWith();
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)originalCreateFileA, MyCreateFileA);
	DetourAttach(&(PVOID&)originalGetSaveFileNameA, MyGetSaveFileNameA);
	DetourAttach(&(PVOID&)originalCreateProcessA, MyCreateProcessA);
	DetourAttach(&(PVOID&)originalMessageBoxA, MyMessageBoxA);

	
	//用于自动在编译和调试之间切换Lib与Dll模块
	DetourAttach(&(PVOID&)originalEStartBuildFunc, MyEStartBuildFunc);
	DetourAttach(&(PVOID&)originalEStartDebugFunc, MyEStartDebugFunc);


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

	}
}

BOOL CALLBACK FindMenuBarEnumChildProc(HWND hwnd, LPARAM lParam) {
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
HWND FindMenuBar(HWND hParent) {
	HWND hResult = NULL;
	EnumChildWindows(hParent, FindMenuBarEnumChildProc, reinterpret_cast<LPARAM>(&hResult));
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
std::string GetSourceFilePath() {
	HWND hWnd = (HWND)NotifySys(NES_GET_MAIN_HWND, 0, 0);
	char buffer[256] = { 0 };
	GetWindowText(hWnd, buffer, sizeof(buffer));
	auto path = ExtractBetweenDashes(std::string(buffer));
	return path;
}


void UpdateButton() {
	if (g_nowOpenSourceFilePath.empty()) {
		SetWindowTextA(g_buttonHwnd, "默认");
	} else {
		auto linkName = g_configManager.getValue(g_nowOpenSourceFilePath);
		if (!linkName.empty()) {
			SetWindowTextA(g_buttonHwnd, linkName.c_str());
		}
		else {
			SetWindowTextA(g_buttonHwnd, "默认");
		}
	}
}


/// <summary>
/// 更新当前源文件使用的link文件
/// </summary>
/// <param name="id"></param>
void UpdateCurrentFileLinkerWithId(int id) {

	//先更新一下
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		//当前好像没有打开源文件
		OutputStringToELog("当前没有打开源文件，无法切换Linker");
	} else {
		g_configManager.setValue(g_nowOpenSourceFilePath, g_linkerManager.getConfig(id).name);
		UpdateButton();
	}
}

/// <summary>
/// 创建按钮
/// TODO 和其他插件UI覆盖冲突的问题？
/// </summary>
/// <param name="hParent"></param>
void CreateAndSubclassButton(HWND hParent) {

	int buttonWidth = 100;
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

void UpdateCurrentOpenSourceFile() {
	std::string sourceFile = GetSourceFilePath();

	if (sourceFile.empty()) {
		//OutputStringToELog("无法获取源文件路径");
	}

	if (g_nowOpenSourceFilePath != sourceFile) {
		OutputStringToELog(sourceFile);
	}
	g_nowOpenSourceFilePath = sourceFile;
}


//工具条子类过程
LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	if (uMsg == WM_COMMAND && LOWORD(wParam) == BN_CLICKED) {
		ShowMenu(hWnd);
		return 0;
	}

	//std::string s = std::format("菜单条消息 {0} {1}", (int)hWnd, uMsg);
	//OutputStringToELog(s);

	switch (uMsg)
	{
		case WM_PAINT: {
			UpdateCurrentOpenSourceFile();
			UpdateButton();
			break;
		}

		case WM_INITMENUPOPUP:
			break;
		case WM_COMMAND: {
			auto low = LOWORD(wParam);
			UpdateCurrentFileLinkerWithId(low);
			break;
		}
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void PeekAllMessage() {
	MSG msg;
	while (PeekMessage(&msg, 0, 0, 0, 1))
	{
		DispatchMessage(&msg);
		TranslateMessage(&msg);
	}
}

void ChangeVMPModel(bool isLib) {
	if (isLib) {
		
		int sdk = FindEModelNameIndex("VMPSDK");
		if (sdk != -1) {
			RemoveEModel(sdk); //移除
		}
		int sdkLib = FindEModelNameIndex("VMPSDK_LIB");
		if (sdkLib == -1) {
			OutputStringToELog("切换到静态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK_LIB.ec", buffer);

			OutputStringToELog(cmd);
			AddEModel2(cmd);
			PeekAllMessage();
		}
	}
	else {
		int sdk = FindEModelNameIndex("VMPSDK_LIB");
		if (sdk != -1) {
			RemoveEModel(sdk); //移除
		}
		int sdkLib = FindEModelNameIndex("VMPSDK");
		if (sdkLib == -1) {
			OutputStringToELog("切换到动态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK.ec", buffer);
			OutputStringToELog(cmd);
			AddEModel2(cmd);
			PeekAllMessage();
		}
	}
	//
}

//主窗口子类过程
LRESULT CALLBACK MainWindowSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	if (uMsg == 20707) {
		if (wParam) {
			g_preCompiling = true;
			OutputStringToELog("开始编译");
			ChangeVMPModel(true);
		}
		else {
			g_preDebugging = true;
			OutputStringToELog("开始调试");
			ChangeVMPModel(false);
		}
		return 0;
	}

	if (uMsg == 20708) { 
		BOOL result = SetWindowSubclass((HWND)wParam, EditViewSubclassProc, 0, 0);
		return result ? 1 : 0;
	}	

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2) {
	DWORD p[2] = {p1, p2};
	return NotifySys(NES_RUN_FUNC, code, (DWORD) & p);
}





INT WINAPI fnAddInFunc(INT nAddInFnIndex) {

	switch (nAddInFnIndex) {
		case 0: { //TODO 打开项目目录
			std::string cmd = std::format("/select,{}", g_nowOpenSourceFilePath);
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 1: { //Auto 目录
			std::string cmd = std::format("{}\\AutoLoader", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 2: { //E 目录
			std::string cmd = std::format("{}", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 3: { //切换到VMPSDK静态
			ChangeVMPModel(true);
			break;
		}
		case 4: { //切换到VMPSDK动态
			ChangeVMPModel(false);

			break;
		}
			  
		default: {

		}
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

	g_toolBarHwnd = FindMenuBar(g_hwnd);


	std::string s = std::format("{} {} {}", processID, (int)g_hwnd, (int)g_toolBarHwnd);
	OutputStringToELog(s);

	if (g_hwnd != NULL && g_toolBarHwnd != NULL)
	{
		StartEditViewSubclassTask();

		OutputStringToELog("找到工具条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		SetWindowSubclass(g_hwnd, MainWindowSubclassProc, 0, 0);
		CreateAndSubclassButton(g_toolBarHwnd);
		StartHookCreateFileA();
		PostAppMessageA(g_toolBarHwnd, WM_PRINT, 0, 0);
		OutputStringToELog("初始化完成");



		//初始化Lib相关库的状态
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
		//NotifySys3 = (NotifySys3Func)dwParam1;
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
static LIB_INFOX LibInfo =
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
	LBS_IDE_PLUGIN | LBS_LIB_INFO2, //_LIB_OS(__OS_WIN), //#LBS_IDE_PLUGIN  LBS_LIB_INFO2
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
	_T("打开项目目录\0这是个用作测试的辅助工具功能。\0打开AutoLoader配置目录\0这是个用作测试的辅助工具功能。\0打开E语言目录\0这是个用作测试的辅助工具功能。\0切换到VMPSDK_LIB模块\0这是个用作测试的辅助工具功能。\0切换到VMPSDK模块\0这是个用作测试的辅助工具功能。\0\0") ,
	AutoLinker_MessageNotify,
	NULL,
	NULL,
	0,
	NULL,
	NULL,
	//-----------------
	NULL,
	0,
	NULL,
	"You",

};

PLIB_INFOX WINAPI GetNewInf()
{
	//LibInfo.m_szLicenseToUserName = "You"
	return (&LibInfo);
};
#endif