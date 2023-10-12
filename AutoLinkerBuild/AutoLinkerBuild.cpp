// AutoLinkerBuild.cpp : 定义 DLL 的导出函数。
//

#include "pch.h"
#include "framework.h"
#include "AutoLinkerBuild.h"
#include "CompilerPluginsPublic.h"
#include <tchar.h>
#include "PathHelper.h"
#include <format>

#pragma comment(lib, "comctl32.lib")

HWND g_hwnd;

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
    szbuf.insert(0, "[AutoLinkerBuild]");

    OutputDebugString(szbuf.c_str());
    HWND hwnd = FindOutputWindow(g_hwnd);
    if (hwnd) {
        SendMessageA(hwnd, 194, 1, (LPARAM)szbuf.c_str());
        SendMessageA(hwnd, 194, 1, (long)"\r\n");
    }
    else {
        //::MessageBox(0, "没有找到", "辅助工具", MB_YESNO);
    }
}

AUTOLINKERBUILD_API BOOL WINAPI CompileProcessor(const COMPILE_PROCESSOR_INFO* pInf)
{
    //if (_tcsicmp(pInf->m_szPurePrgFileName, _T("CompilerPluginsSample")) != 0)  // 不为指定的易语言程序名?
    //    return FALSE;  // 返回失败

    DWORD processID = GetCurrentProcessId();
    EnumWindows(EnumWindowsProc, processID);
    auto path = GetBasePath();

    std::string s = std::format("{} {} {} {} {}", (int)g_hwnd, (int)processID, path, (int)pInf->m_blCompileReleaseVersion, pInf->m_szPurePrgFileName);
    OutputStringToELog(s);

    //发一条自定义消息到AutoLinker
    SendMessageA(g_hwnd, 20707, pInf->m_blCompileReleaseVersion, 0);

    //编译调试
    //if (pInf->m_blCompileReleaseVersion) {
    //    //编译 同步通知
    //    MessageBoxA(0, path.c_str(), "编译", 0);
    //}
    //else {
    //    //调试 同步通知
    //    MessageBoxA(0, path.c_str(), "调试", 0);
    //}


    

    return TRUE;  // 返回成功
}