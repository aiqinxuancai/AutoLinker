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
#include "WindowHelper.h"
#include <future>
#include "WinINetUtil.h"
#include "Version.h"
#include "IDEFacade.h"
#include <unordered_map>

#pragma comment(lib, "comctl32.lib")

//当前打开的易源文件路径
std::string g_nowOpenSourceFilePath;

//配置文件，当前路径的e源码对应的编译器名
ConfigManager g_configManager;

//管理当前所有的link文件
LinkerManager g_linkerManager;

//管理模块 调试 -> 编译 的管理器
ModelManager g_modelManager;


//e主窗口句柄
HWND g_hwnd = NULL;

//工具条句柄
HWND g_toolBarHwnd = NULL;

//准备开始调试 废弃
bool g_preDebugging;

//准备开始编译 废弃
bool g_preCompiling;

bool g_initStarted = false;

HMENU g_topLinkerSubMenu = NULL;
std::unordered_map<UINT, std::string> g_topLinkerCommandMap;

void UpdateCurrentOpenSourceFile();
void OutputCurrentSourceLinker();

namespace {
bool g_isContextMenuRegistered = false;
constexpr UINT IDM_AUTOLINKER_CTX_COPY_FUNC = 31001; // Keep < 0x8000 to avoid signed-ID issues in some hosts.
constexpr UINT IDM_AUTOLINKER_LINKER_BASE = 34000;
constexpr UINT IDM_AUTOLINKER_LINKER_MAX = 34999;

void TryCopyCurrentFunctionCode()
{
	if (IDEFacade::Instance().CopyCurrentFunctionCodeToClipboard()) {
		OutputStringToELog("已复制当前函数代码到剪贴板");
	}
	else {
		OutputStringToELog("复制当前函数代码失败，当前位置可能不在代码函数中");
	}
}

void RegisterIDEContextMenu()
{
	if (g_isContextMenuRegistered) {
		return;
	}

	auto& ide = IDEFacade::Instance();
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_COPY_FUNC, "复制当前函数代码", []() {
		TryCopyCurrentFunctionCode();
	});

	g_isContextMenuRegistered = true;
}

std::wstring GetMenuTitleW(HMENU hMenu, UINT item, UINT flags)
{
	wchar_t title[256] = { 0 };
	int len = GetMenuStringW(hMenu, item, title, static_cast<int>(sizeof(title) / sizeof(title[0])), flags);
	if (len <= 0) {
		return L"";
	}
	return std::wstring(title, static_cast<size_t>(len));
}

bool IsCompileOrToolsTopPopup(HMENU hPopupMenu)
{
	if (g_hwnd == NULL || hPopupMenu == NULL) {
		return false;
	}
	if (hPopupMenu == g_topLinkerSubMenu) {
		return false;
	}

	// 右键菜单里会包含该命令，直接排除，避免“链接器切换”混入右键。
	UINT copyState = GetMenuState(hPopupMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND);
	if (copyState != 0xFFFFFFFF) {
		return false;
	}

	auto popupKeywordMatch = [hPopupMenu]() -> bool {
		int keywordHit = 0;
		int itemCount = GetMenuItemCount(hPopupMenu);
		for (int item = 0; item < itemCount; ++item) {
			std::wstring itemTitle = GetMenuTitleW(hPopupMenu, static_cast<UINT>(item), MF_BYPOSITION);
			if (itemTitle.find(L"编译") != std::wstring::npos ||
				itemTitle.find(L"发布") != std::wstring::npos ||
				itemTitle.find(L"静态") != std::wstring::npos ||
				itemTitle.find(L"Build") != std::wstring::npos ||
				itemTitle.find(L"Release") != std::wstring::npos) {
				++keywordHit;
			}
		}
		return keywordHit >= 1;
	};

	HMENU hMainMenu = GetMenu(g_hwnd);
	if (hMainMenu == NULL) {
		return popupKeywordMatch();
	}

	int count = GetMenuItemCount(hMainMenu);
	int popupIndex = -1;
	int compileIndex = -1;
	int toolsIndex = -1;

	for (int i = 0; i < count; ++i) {
		HMENU subMenu = GetSubMenu(hMainMenu, i);
		if (subMenu == NULL) {
			continue;
		}
		if (subMenu == hPopupMenu) {
			popupIndex = i;
		}

		std::wstring title = GetMenuTitleW(hMainMenu, static_cast<UINT>(i), MF_BYPOSITION);
		if (title.find(L"编译") != std::wstring::npos || title.find(L"Build") != std::wstring::npos) {
			if (compileIndex < 0) {
				compileIndex = i;
			}
			continue;
		}
		if (title.find(L"工具") != std::wstring::npos || title.find(L"Tools") != std::wstring::npos) {
			if (toolsIndex < 0) {
				toolsIndex = i;
			}
		}
	}

	if (popupIndex < 0) {
		return popupKeywordMatch();
	}

	int targetIndex = compileIndex >= 0 ? compileIndex : toolsIndex;
	if (targetIndex >= 0) {
		return popupIndex == targetIndex;
	}

	// 顶级标题不可读时，回退到子项关键词判断。
	return popupKeywordMatch();
}

void ClearMenuItemsByPosition(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}

	int count = GetMenuItemCount(hMenu);
	for (int i = count - 1; i >= 0; --i) {
		DeleteMenu(hMenu, static_cast<UINT>(i), MF_BYPOSITION);
	}
}

bool EnsureTopLinkerSubMenu()
{
	if (g_topLinkerSubMenu == NULL || !IsMenu(g_topLinkerSubMenu)) {
		g_topLinkerSubMenu = CreatePopupMenu();
	}
	if (g_topLinkerSubMenu == NULL) {
		return false;
	}

	return true;
}

void EnsureTopLinkerSubMenuAttached(HMENU hTargetMenu)
{
	if (hTargetMenu == NULL) {
		return;
	}
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	bool alreadyAttached = false;
	int attachedIndex = -1;
	for (int i = GetMenuItemCount(hTargetMenu) - 1; i >= 0; --i) {
		MENUITEMINFOW mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_SUBMENU;
		bool removeThis = false;
		if (GetMenuItemInfoW(hTargetMenu, static_cast<UINT>(i), TRUE, &mii) && mii.hSubMenu == g_topLinkerSubMenu) {
			alreadyAttached = true;
			attachedIndex = i;
			continue;
		}
		if (!removeThis) {
			std::wstring title = GetMenuTitleW(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (title.find(L"链接器切换") != std::wstring::npos) {
				removeThis = true;
			}
		}
		if (removeThis) {
			DeleteMenu(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (i > 0) {
				UINT prevState = GetMenuState(hTargetMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				if (prevState != 0xFFFFFFFF && (prevState & MF_SEPARATOR) == MF_SEPARATOR) {
					DeleteMenu(hTargetMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				}
			}
		}
	}

	if (alreadyAttached) {
		if (attachedIndex > 0) {
			UINT prevState = GetMenuState(hTargetMenu, static_cast<UINT>(attachedIndex - 1), MF_BYPOSITION);
			if (prevState == 0xFFFFFFFF || (prevState & MF_SEPARATOR) != MF_SEPARATOR) {
				InsertMenuW(hTargetMenu, static_cast<UINT>(attachedIndex), MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
			}
		}
		return;
	}

	int count = GetMenuItemCount(hTargetMenu);
	if (count > 0) {
		UINT lastState = GetMenuState(hTargetMenu, static_cast<UINT>(count - 1), MF_BYPOSITION);
		if (lastState != 0xFFFFFFFF && (lastState & MF_SEPARATOR) != MF_SEPARATOR) {
			AppendMenuW(hTargetMenu, MF_SEPARATOR, 0, NULL);
		}
	}
	AppendMenuW(hTargetMenu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(g_topLinkerSubMenu), L"链接器切换");
}

void RebuildTopLinkerSubMenu()
{
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	ClearMenuItemsByPosition(g_topLinkerSubMenu);
	g_topLinkerCommandMap.clear();

	if (g_linkerManager.getCount() <= 0) {
		AppendMenuA(g_topLinkerSubMenu, MF_STRING | MF_GRAYED, 0, "（未找到Linker配置）");
		return;
	}

	UpdateCurrentOpenSourceFile();
	std::string nowLinkConfigName = g_configManager.getValue(g_nowOpenSourceFilePath);

	UINT cmd = IDM_AUTOLINKER_LINKER_BASE;
	for (const auto& [key, value] : g_linkerManager.getMap()) {
		if (cmd > IDM_AUTOLINKER_LINKER_MAX) {
			break;
		}

		UINT flags = MF_STRING | MF_ENABLED;
		if (!nowLinkConfigName.empty() && nowLinkConfigName == key) {
			flags |= MF_CHECKED;
		}

		AppendMenuA(g_topLinkerSubMenu, flags, cmd, key.c_str());
		g_topLinkerCommandMap[cmd] = key;
		++cmd;
	}
}

bool HandleTopLinkerMenuCommand(UINT cmd)
{
	auto it = g_topLinkerCommandMap.find(cmd);
	if (it == g_topLinkerCommandMap.end()) {
		return false;
	}

	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		OutputStringToELog("当前没有打开源文件，无法切换Linker");
		return true;
	}

	g_configManager.setValue(g_nowOpenSourceFilePath, it->second);
	OutputCurrentSourceLinker();
	RebuildTopLinkerSubMenu();
	return true;
}

void HandleInitMenuPopup(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}
	if (hMenu == g_topLinkerSubMenu) {
		RebuildTopLinkerSubMenu();
		return;
	}

	if (IsCompileOrToolsTopPopup(hMenu)) {
		EnsureTopLinkerSubMenuAttached(hMenu);
		RebuildTopLinkerSubMenu();
		return;
	}

	UINT state = GetMenuState(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND);
	if (state != 0xFFFFFFFF) {
		EnableMenuItem(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND | MF_ENABLED);
	}
}
}


static auto originalCreateFileA = CreateFileA;
static auto originalGetSaveFileNameA = GetSaveFileNameA;
static auto originalCreateProcessA = CreateProcessA;
static auto originalMessageBoxA = MessageBoxA;

//开始调试 5.71-5.95
typedef int(__thiscall* OriginalEStartDebugFuncType)(DWORD* thisPtr, int a2, int a3);
OriginalEStartDebugFuncType originalEStartDebugFunc = (OriginalEStartDebugFuncType)0x40A080; //int __thiscall sub_40A080(int this, int a2, int a3)

//开始编译 5.71-5.95
typedef int(__thiscall* OriginalEStartCompileFuncType)(DWORD* thisPtr, int a2);
OriginalEStartCompileFuncType originalEStartCompileFunc = (OriginalEStartCompileFuncType)0x40A9F1;  //int __thiscall sub_40A9F1(_DWORD *this, int a2)

int __fastcall MyEStartCompileFunc(DWORD* thisPtr, int dummy, int a2) {
	OutputStringToELog("编译开始#2");
	//ChangeVMProtectModel(true);
	RunChangeECOM(true);
	return originalEStartCompileFunc(thisPtr, a2);
}

int __fastcall MyEStartDebugFunc(DWORD* thisPtr, int dummy, int a2, int a3) {
	OutputStringToELog("调试开始#2");
	//ChangeVMProtectModel(false);
	RunChangeECOM(false);
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
	std::filesystem::path autoLinkerPath = currentPath / "tools" / "link.ini";
	if (autoLinkerPath.string() == std::string(lpFileName)) {
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

	//检查加入提前链接
	auto krnlnPath = GetLinkerCommandKrnlnFileName(lpCommandLine);
	OutputStringToELog(krnlnPath);

	if (!krnlnPath.empty()) {
		//核心库代码优先使用黑月的，然后再使用核心库的
		//std::string bklib = "D:\\git\\KanAutoControls\\Release\\TestCore.lib";
		
		//将把这个Lib插入的链接器的前方，添加/FORCE强制链接，忽略重定义
		std::string libFilePath = std::format("{}\\AutoLinker\\ForceLinkLib.txt", GetBasePath());
		auto libList = ReadFileAndSplitLines(libFilePath);



		if (libList.size() > 0) {
			//和链接器关联
			auto currentLinkerName = g_configManager.getValue(g_nowOpenSourceFilePath);

			OutputStringToELog(std::format("当前指定的链接器：{}", currentLinkerName));

			std::string libCmd;
			for (const auto& line : libList) {
				auto lines = SplitStringTwo(line, '=');
				auto libPath = line;
				std::string linkerName;
				if (lines.size() == 2) {
					linkerName = lines[0];
					libPath = lines[1];
				}

				//OutputStringToELog(std::format("找到设定的强制链接Lib：{} -> {}", linkerName, libPath));

				if (!linkerName.empty()) {
					//要求必须指定Linker才可使用（包含名称既可）
					if (currentLinkerName.find(linkerName) != std::string::npos) {
						//可使用，link名称一致
						if (std::filesystem::exists(libPath)) {
							OutputStringToELog(std::format("强制链接Lib：{} -> {}", linkerName, libPath));
							if (!libCmd.empty()) {
								libCmd += " ";
							}
							libCmd += "\"" + libPath + "\"";
						}
					} else {
						
						OutputStringToELog(std::format("链接器{}不符合当前的链接器{}，不链接：{}", linkerName, currentLinkerName, libPath));
					}

				} else {
					if (std::filesystem::exists(libPath)) {
						OutputStringToELog(std::format("强制链接Lib：{}", libPath));
						if (!libCmd.empty()) {
							libCmd += " ";
						}
						libCmd += "\"" + libPath + "\"";
					}
				}


			}

			std::string newLibs = libCmd + " \"" + krnlnPath + "\" /FORCE";
			commandLine = ReplaceSubstring(commandLine, "\"" + krnlnPath + "\"", newLibs);


		}
	}

	auto outFileName = GetLinkerCommandOutFileName(lpCommandLine);
	if (!outFileName.empty()) {
		if (commandLine.find("/pdb:\"build.pdb\"") != std::string::npos) {
			//PDB更名为当前编译的程序的名字+.pdb，看起来更正规
			std::string newPdbCommand = std::format("/pdb:\"{}.pdb\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/pdb:\"build.pdb\"", newPdbCommand);
		}
		if (commandLine.find("/map:\"build.map\"") != std::string::npos) {
			//Map更名
			std::string newPdbCommand = std::format("/map:\"{}.map\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/map:\"build.map\"", newPdbCommand);
		}
	}
	OutputStringToELog("启动命令行：" + commandLine);
	return originalCreateProcessA(lpApplicationName, (char *)commandLine.c_str(), lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
}

int WINAPI MyMessageBoxA(
	HWND hWnd,
	LPCSTR lpText,
	LPCSTR lpCaption,
	UINT uType) {
	//自动返回确认编译
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


	if (g_debugStartAddress != -1 && g_compileStartAddress != -1) {
		originalEStartCompileFunc = (OriginalEStartCompileFuncType)g_compileStartAddress;
		originalEStartDebugFunc = (OriginalEStartDebugFuncType)g_debugStartAddress;

		//用于自动在编译和调试之间切换Lib与Dll模块
		DetourAttach(&(PVOID&)originalEStartCompileFunc, MyEStartCompileFunc);
		DetourAttach(&(PVOID&)originalEStartDebugFunc, MyEStartDebugFunc);
	}
	else {
		//无法启用
	}
	DetourTransactionCommit();
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

void OutputCurrentSourceLinker()
{
	UpdateCurrentOpenSourceFile();

	std::string linkerName = "默认";
	if (!g_nowOpenSourceFilePath.empty()) {
		std::string configured = g_configManager.getValue(g_nowOpenSourceFilePath);
		if (!configured.empty()) {
			linkerName = configured;
		}
	}

	OutputStringToELog("当前源码链接器：" + linkerName);
}

//工具条子类过程
LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	switch (uMsg)
	{
		case WM_INITMENUPOPUP:
			HandleInitMenuPopup(reinterpret_cast<HMENU>(wParam));
			break;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}


/// <summary>
/// 废弃，使用新的配置来切换了
/// </summary>
/// <param name="isLib"></param>
void ChangeVMProtectModel(bool isLib) {
	if (isLib) {
		
		int sdk = FindECOMNameIndex("VMPSDK");
		if (sdk != -1) {
			RemoveECOM(sdk); //移除
		}
		int sdkLib = FindECOMNameIndex("VMPSDK_LIB");
		if (sdkLib == -1) {
			OutputStringToELog("切换到静态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK_LIB.ec", buffer);

			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	else {
		int sdk = FindECOMNameIndex("VMPSDK_LIB");
		if (sdk != -1) {
			RemoveECOM(sdk); //移除
		}
		int sdkLib = FindECOMNameIndex("VMPSDK");
		if (sdkLib == -1) {
			OutputStringToELog("切换到动态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK.ec", buffer);
			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	//
}

#define WM_AUTOLINKER_INIT (WM_USER + 1000)

bool FneInit();

//主窗口子类过程
LRESULT CALLBACK MainWindowSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	//if (uMsg == 20707) {
	//	//此消息配合AutoLinkerBuild，已废弃！
	//	if (wParam) {
	//		g_preCompiling = true;
	//		OutputStringToELog("开始编译");
	//		//ChangeVMPModel(true);
	//	}
	//	else {
	//		g_preDebugging = true;
	//		OutputStringToELog("开始调试");
	//		//ChangeVMPModel(false);
	//	}
	//	return 0;
	//}

	if (uMsg == WM_COMMAND) {
		UINT cmd = LOWORD(wParam);
		if (HandleTopLinkerMenuCommand(cmd)) {
			return 0;
		}
		if (IDEFacade::Instance().HandleMainWindowCommand(wParam)) {
			return 0;
		}
	}

	if (uMsg == WM_AUTOLINKER_INIT) {
		OutputStringToELog("收到初始化消息，尝试初始化");
		if (FneInit()) {
			OutputStringToELog("初始化成功");
			return 1;
		}
		return 0;
	}

	if (uMsg == 20708) {
		BOOL result = SetWindowSubclass((HWND)wParam, EditViewSubclassProc, 0, 0);
		return result ? 1 : 0;
	}

	if (uMsg == WM_INITMENUPOPUP) {
		LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		HandleInitMenuPopup(reinterpret_cast<HMENU>(wParam));
		return result;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2) {
	return IDEFacade::Instance().RunFunctionRaw(code, p1, p2);
}

INT WINAPI fnAddInFunc(INT nAddInFnIndex) {

	switch (nAddInFnIndex) {
		case 0: { //TODO 打开项目目录
			std::string cmd = std::format("/select,{}", g_nowOpenSourceFilePath);
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 1: { //Auto 目录
			std::string cmd = std::format("{}\\AutoLinker", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 2: { //E 目录
			std::string cmd = std::format("{}", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 3: { // 复制当前函数代码
			TryCopyCurrentFunctionCode();
			break;
		}
		//case 4: { //切换到VMPSDK静态（自用）
		//	ChangeVMProtectModel(true);
		//	break;
		//}
		//case 5: { //切换到VMPSDK动态（自用）
		//	ChangeVMProtectModel(false);
		//	break;
		//}
			  
		default: {

		}
	}

	return 0;
}


void FneCheckNewVersion(void* pParams) {
	Sleep(1000);

	OutputStringToELog("AutoLinker开源下载地址：https://github.com/aiqinxuancai/AutoLinker");
	std::string url = "https://api.github.com/repos/aiqinxuancai/AutoLinker/releases";
	//std::string customHeaders = "user-agent: Mozilla/5.0";
	auto response = PerformGetRequest(url);

	std::string currentVersion = AUTOLINKER_VERSION;


	if (response.second == 200) {
		std::string nowGithubVersion = "0.0.0";
		

		if (strcmp(AUTOLINKER_VERSION, "0.0.0") == 0) {
			//自行编译，无需检查版本更新
			OutputStringToELog(std::format("自编译版本，不检查更新，当前版本：{}", currentVersion));
		} else {
			if (!response.first.empty()) {

				try
				{
					auto releases = json::parse(response.first);
					for (const auto& release : releases) {
						if (!release["prerelease"].get<bool>()) {
							nowGithubVersion = release["tag_name"];
							break;
						}
					}

					Version nowGithubVersionObj(nowGithubVersion);
					Version currentVersionObj(AUTOLINKER_VERSION);

					if (nowGithubVersionObj > currentVersionObj) {
						OutputStringToELog(std::format("有新版本：{}", nowGithubVersion));
					}
					else {

					}
				}
				catch (const std::exception& e) {
					OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误：{}", currentVersion, e.what()));
				}


			}
		}

		
	}
	else {
		//OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误码：{}", currentVersion, response.second));
	}

	//return false;
}

bool FneInit() {
	OutputStringToELog("开始初始化");

	// g_hwnd 已经在外部获取并子类化
	if (g_hwnd == NULL) {
		OutputStringToELog("g_hwnd 为空");
		return false;
	}

	g_toolBarHwnd = FindMenuBar(g_hwnd);

	DWORD processID = GetCurrentProcessId();
	std::string s = std::format("E进程ID{} 主句柄{} 菜单栏句柄{}", processID, (int)g_hwnd, (int)g_toolBarHwnd);
	OutputStringToELog(s);

	if (g_toolBarHwnd != NULL)
	{
		StartEditViewSubclassTask();
		RebuildTopLinkerSubMenu();
		OutputCurrentSourceLinker();

		OutputStringToELog("找到工具条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		StartHookCreateFileA();
		PostAppMessageA(g_toolBarHwnd, WM_PRINT, 0, 0);
		OutputStringToELog("初始化完成");

		//初始化Lib相关库的状态

		//启动版本检查线程
		uintptr_t threadID = _beginthread(FneCheckNewVersion, 0, NULL);

		return true;
	}
	else
	{
		OutputStringToELog(std::format("初始化失败，未找到工具条窗口 {}", (int)g_toolBarHwnd));
	}

	return false;
}

/*-----------------支持库消息处理函数------------------*/

// 初始化重试线程函数
void InitRetryThread(void* pParams) {
	const int MAX_RETRY_COUNT = 5;
	int retryCount = 0;

	while (retryCount < MAX_RETRY_COUNT) {
		Sleep(1000); // 每秒重试一次
		retryCount++;

		OutputStringToELog(std::format("初始化重试第 {}/{} 次", retryCount, MAX_RETRY_COUNT));

		// 发送自定义消息到主窗口进行初始化
		LRESULT result = SendMessage(g_hwnd, WM_AUTOLINKER_INIT, 0, 0);

		if (result == 1) {
			// 初始化成功
			OutputStringToELog("初始化线程完成"); 
			return;
		}
	}

	OutputStringToELog(std::format("初始化失败，已重试 {} 次", MAX_RETRY_COUNT));
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
		if (dwParam1) {
			if (!g_initStarted) {
				// 获取主窗口句柄
				g_hwnd = GetMainWindowByProcessId();

				if (g_hwnd) {
					g_initStarted = true;
					SetWindowSubclass(g_hwnd, MainWindowSubclassProc, 0, 0);
					RegisterIDEContextMenu();
					OutputStringToELog("主窗口子类化完成，启动初始化线程");
					uintptr_t threadID = _beginthread(InitRetryThread, 0, NULL);
				} else {
					OutputStringToELog("无法获取主窗口句柄");
				}
			}

		}
	}
	else if (nMsg == NL_RIGHT_POPUP_MENU_SHOW) {
		IDEFacade::Instance().HandleNotifyMessage(nMsg, dwParam1, dwParam2);
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
	_T("打开项目目录\0这是个用作测试的辅助工具功能。\0打开AutoLinker配置目录\0这是个用作测试的辅助工具功能。\0打开E语言目录\0这是个用作测试的辅助工具功能。\0复制当前函数代码\0复制当前光标所在子程序完整代码到剪贴板。\0\0") ,
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
