#pragma once

#include <Windows.h>

// AutoLinker 无头编译启动器：解析 e.exe 命令行并在 IDE 就绪后执行编译。
namespace HeadlessCompileRunner {

// 判断当前进程是否带有 AutoLinker 无头编译请求。
bool HasHeadlessCompileRequest();

// 如请求要求无头运行，则尽早隐藏 IDE 主窗口。
void ApplyInitialWindowState(HWND mainWindow);

// IDE 初始化完成后启动一次无头编译任务。
void StartIfRequested();

} // namespace HeadlessCompileRunner
