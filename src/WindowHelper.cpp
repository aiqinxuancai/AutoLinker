
#include "WindowHelper.h"
#include <lib2.h>
#include "PathHelper.h"
#include <fnshare.h>
#include <cctype>
#include <cstring>
#include <filesystem>
#include <format>

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

bool MatchSourceExtensionAt(const std::string& title, size_t position, size_t& extensionEnd)
{
	if (position + 2 > title.size() || title[position] != '.' ||
		std::tolower(static_cast<unsigned char>(title[position + 1])) != 'e') {
		return false;
	}

	extensionEnd = position + 2;
	if (extensionEnd < title.size() &&
		std::tolower(static_cast<unsigned char>(title[extensionEnd])) == 'c') {
		++extensionEnd;
	}
	return true;
}

std::string ExtractSourcePathCandidate(const std::string& title, size_t pathStart)
{
	std::string delimiterTerminatedCandidate;
	for (size_t position = title.find('.', pathStart);
		position != std::string::npos;
		position = title.find('.', position + 1)) {
		size_t extensionEnd = 0;
		if (!MatchSourceExtensionAt(title, position, extensionEnd)) {
			continue;
		}

		const bool reachesTitleEnd = extensionEnd == title.size();
		const bool followedByAnnotation =
			title.compare(extensionEnd, 2, " [") == 0;
		const bool followedByTitleDelimiter =
			title.compare(extensionEnd, 3, " - ") == 0;
		if (!reachesTitleEnd && !followedByAnnotation && !followedByTitleDelimiter) {
			continue;
		}

		const std::string candidate = TrimAsciiCopyLocal(
			title.substr(pathStart, extensionEnd - pathStart));
		if (!LooksLikeAbsolutePathSegment(candidate) ||
			!HasEideSourceFileExtension(candidate)) {
			continue;
		}

		// 注解或标题结束是强边界；遇到 " - " 时继续向后匹配，兼容文件名自身包含该文本。
		if (reachesTitleEnd || followedByAnnotation) {
			return candidate;
		}
		delimiterTerminatedCandidate = candidate;
	}
	return delimiterTerminatedCandidate;
}

std::string ExtractSourcePathFromWindowTitleLocal(const std::string& title)
{
	static constexpr const char* kDelimiter = " - ";
	size_t segmentStart = 0;
	while (segmentStart <= title.size()) {
		while (segmentStart < title.size() &&
			std::isspace(static_cast<unsigned char>(title[segmentStart])) != 0) {
			++segmentStart;
		}

		if (LooksLikeAbsolutePathSegment(title.substr(segmentStart))) {
			const std::string path = ExtractSourcePathCandidate(title, segmentStart);
			if (!path.empty()) {
				return path;
			}
		}

		const size_t delimiterPosition = title.find(kDelimiter, segmentStart);
		if (delimiterPosition == std::string::npos) {
			break;
		}
		segmentStart = delimiterPosition + std::strlen(kDelimiter);
	}
	return std::string();
}

}

std::string ExtractSourcePathFromWindowTitle(const std::string& title)
{
	return ExtractSourcePathFromWindowTitleLocal(title);
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
	if (hWnd == nullptr) {
		return std::string();
	}

	const int titleLength = GetWindowTextLengthA(hWnd);
	if (titleLength <= 0) {
		return std::string();
	}
	std::string title(static_cast<size_t>(titleLength) + 1, '\0');
	const int copiedLength = GetWindowTextA(hWnd, title.data(), titleLength + 1);
	if (copiedLength <= 0) {
		return std::string();
	}
	title.resize(static_cast<size_t>(copiedLength));
	return ExtractSourcePathFromWindowTitle(title);
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
