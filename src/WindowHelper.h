#pragma once

#include "Windows.h"
#include <string>

//查找输出窗口
HWND FindOutputWindow(HWND hParent);


//获取当前源文件的路径 （从标题栏）
std::string GetSourceFilePath();

//从易语言主窗口标题中提取当前源文件路径。
std::string ExtractSourcePathFromWindowTitle(const std::string& title);


//等于"处理事件"
void PeekAllMessage();

//获取E主窗口句柄（通过进程ID枚举）
HWND GetMainWindowByProcessId();
