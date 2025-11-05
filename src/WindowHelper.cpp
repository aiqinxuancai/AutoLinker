
#include "WindowHelper.h"
#include <lib2.h>
#include "PathHelper.h"
#include <fnshare.h>
#include <format>

BOOL CALLBACK FindMenuBarEnumChildProc(HWND hwnd, LPARAM lParam) {
	char buffer[256] = { 0 };
	GetWindowText(hwnd, buffer, sizeof(buffer));
	char bufferClassName[256] = { 0 };
	GetClassName(hwnd, bufferClassName, sizeof(bufferClassName));
	auto className = std::string(bufferClassName);

	if (std::string(buffer) == "菜单条" && className.starts_with("Afx:400000:b:100") && className.ends_with(":10:0")) {
		HWND* pResult = reinterpret_cast<HWND*>(lParam);
		*pResult = hwnd;
		return FALSE;
	}

	if (std::string(buffer) == "菜单条" && className.starts_with("AfxControlBar42s")) {
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


void PeekAllMessage() {
	MSG msg;
	while (PeekMessage(&msg, 0, 0, 0, 1))
	{
		DispatchMessage(&msg);
		TranslateMessage(&msg);
	}
}

// 用于存储枚举到的主窗口句柄
static HWND g_enumMainHwnd = NULL;

/// <summary>
/// 枚举窗口回调函数
/// </summary>
BOOL CALLBACK EnumWindowsProcForMainWindow(HWND hwnd, LPARAM lParam)
{
	DWORD lpdwProcessId;
	GetWindowThreadProcessId(hwnd, &lpdwProcessId);
	if (lpdwProcessId == lParam)
	{
		char windowTitle[256] = { 0 };
		GetWindowText(hwnd, windowTitle, sizeof(windowTitle));
		char className[256] = { 0 };
		GetClassName(hwnd, className, sizeof(className));

		// 检查窗口是否可见
		BOOL isVisible = IsWindowVisible(hwnd);

		OutputDebugStringA(std::format("枚举到窗口: HWND={}, 标题={}, 类名={}, 可见={}\n",
			(int)hwnd, windowTitle, className, isVisible).c_str());

		HWND hwndTopLevel = GetAncestor(hwnd, GA_ROOTOWNER);

		char topClassName[256] = { 0 };
		if (hwndTopLevel != hwnd) {
			GetWindowText(hwndTopLevel, windowTitle, sizeof(windowTitle));
			GetClassName(hwndTopLevel, topClassName, sizeof(topClassName));
			isVisible = IsWindowVisible(hwndTopLevel);
			OutputDebugStringA(std::format("顶层窗口: HWND={}, 标题={}, 类名={}, 可见={}\n",
				(int)hwndTopLevel, windowTitle, topClassName, isVisible).c_str());
		} else {
			// 如果当前窗口就是顶层窗口，使用当前窗口的类名
			strcpy_s(topClassName, sizeof(topClassName), className);
		}

		// 检查是否是E语言主窗口（类名为 ENewFrame）
		if (hwndTopLevel && std::string(topClassName) == "ENewFrame") {
			OutputDebugStringA(std::format("找到E语言主窗口: HWND={}\n", (int)hwndTopLevel).c_str());
			g_enumMainHwnd = hwndTopLevel;
			return FALSE;
		}
	}
	return TRUE;
}

/// <summary>
/// 获取E主窗口句柄（通过进程ID枚举）
/// </summary>
HWND GetMainWindowByProcessId() {
	g_enumMainHwnd = NULL;
	DWORD processID = GetCurrentProcessId();
	OutputDebugStringA(std::format("开始枚举窗口，进程ID: {}\n", processID).c_str());
	EnumWindows(EnumWindowsProcForMainWindow, processID);

	if (g_enumMainHwnd) {
		char windowTitle[256] = { 0 };
		GetWindowText(g_enumMainHwnd, windowTitle, sizeof(windowTitle));
		char className[256] = { 0 };
		GetClassName(g_enumMainHwnd, className, sizeof(className));
		OutputDebugStringA(std::format("最终返回窗口: HWND={}, 标题={}, 类名={}\n", (int)g_enumMainHwnd, windowTitle, className).c_str());
	} else {
		OutputDebugStringA("未找到主窗口句柄\n");
	}

	return g_enumMainHwnd;
}