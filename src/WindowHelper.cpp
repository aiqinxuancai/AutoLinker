
#include "WindowHelper.h"
#include <lib2.h>
#include "PathHelper.h"
#include <fnshare.h>
#include <filesystem>
#include <format>
#include <regex>

namespace {

std::string TrimAsciiCopyLocal(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string StripTrailingBracketAnnotations(const std::string& text)
{
	static const std::regex kTrailingBracketPattern(R"(\s*\[[^\[\]\r\n]*\]\s*$)");
	std::string stripped = text;
	for (;;) {
		const std::string next = std::regex_replace(stripped, kTrailingBracketPattern, std::string());
		if (next == stripped) {
			break;
		}
		stripped = next;
	}
	return TrimAsciiCopyLocal(stripped);
}

bool HasEideSourceFileExtension(const std::string& text)
{
	try {
		std::filesystem::path path(text);
		std::string ext = path.extension().string();
		for (char& ch : ext) {
			ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
		}
		return ext == ".e" || ext == ".ec";
	}
	catch (...) {
		return false;
	}
}

bool LooksLikeAbsolutePathSegment(const std::string& text)
{
	if (text.size() >= 3 &&
		std::isalpha(static_cast<unsigned char>(text[0])) != 0 &&
		text[1] == ':' &&
		(text[2] == '\\' || text[2] == '/')) {
		return true;
	}
	return text.rfind("\\\\", 0) == 0 || text.rfind("//", 0) == 0;
}

std::string ExtractSourcePathFromWindowTitle(const std::string& title)
{
	static constexpr const char* kDelimiter = " - ";
	size_t begin = 0;
	while (begin <= title.size()) {
		size_t end = title.find(kDelimiter, begin);
		if (end == std::string::npos) {
			end = title.size();
		}

		const std::string segment = StripTrailingBracketAnnotations(title.substr(begin, end - begin));
		if (!segment.empty() &&
			LooksLikeAbsolutePathSegment(segment) &&
			HasEideSourceFileExtension(segment)) {
			return segment;
		}

		if (end == title.size()) {
			break;
		}
		begin = end + std::strlen(kDelimiter);
	}
	return std::string();
}

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
	return ExtractSourcePathFromWindowTitle(std::string(buffer));
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
