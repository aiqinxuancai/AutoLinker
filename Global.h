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

#endif // GLOBAL_H