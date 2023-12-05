// global.h
#pragma once
#ifndef GLOBAL_H
#define GLOBAL_H

//编译器地址
extern int g_compilerAddress;

extern HWND g_hwnd;

/// <summary>
/// 输出文本
/// </summary>
/// <param name="szbuf"></param>
void OutputStringToELog(std::string szbuf);


/// <summary>
/// 运行插件通知
/// </summary>
/// <param name="code"></param>
/// <param name="p1"></param>
/// <param name="p2"></param>
/// <returns></returns>
INT NESRUNFUNC(INT code, DWORD p1, DWORD p2);

#endif // GLOBAL_H