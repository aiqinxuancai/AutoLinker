// global.h
#pragma once
#include "ModelManager.h"
#ifndef GLOBAL_H
#define GLOBAL_H

//调试开始地址
extern int g_debugStartAddress;

//编译开始地址
extern int g_compileStartAddress;

extern HWND g_hwnd;


//管理模块调试版及编译版本的管理器
extern ModelManager g_modelManager;

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