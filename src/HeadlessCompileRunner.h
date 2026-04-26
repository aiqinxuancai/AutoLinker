#pragma once

#include <Windows.h>

#include <string>

// AutoLinker 无头编译启动器：解析 e.exe 命令行并在 IDE 就绪后执行编译。
namespace HeadlessCompileRunner {

// 判断当前进程是否带有 AutoLinker 无头编译请求。
bool HasHeadlessCompileRequest();

// 如请求要求无头运行，则尽早隐藏 IDE 主窗口。
void ApplyInitialWindowState(HWND mainWindow);

// IDE 初始化完成后启动一次无头编译任务。
void StartIfRequested();

// 编译相关 Hook 已安装，可以真正发起编译。
void NotifyIdeRuntimeReady();

// 捕获易语言 IDE 弹出的 ANSI/Unicode 消息框。
void ReportIdeMessageBoxA(const std::string& captionLocal, const std::string& textLocal, UINT type);
void ReportIdeMessageBoxW(const std::wstring& caption, const std::wstring& text, UINT type);

// 无头模式下消息框自动返回的按钮值。
int GetMessageBoxAutoResponse(UINT type);

} // namespace HeadlessCompileRunner
