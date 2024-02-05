
#include "WindowHelper.h"
#include <lib2.h>
#include "PathHelper.h"
#include <fnshare.h>

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