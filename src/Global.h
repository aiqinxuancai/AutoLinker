// global.h
#pragma once

#ifndef GLOBAL_H
#define GLOBAL_H

#include <Windows.h>

#include <cstdint>
#include <string>

#include "ModelManager.h"

extern int g_debugStartAddress;
extern int g_compileStartAddress;
extern std::string g_nowOpenSourceFilePath;
extern HWND g_hwnd;
extern ModelManager g_modelManager;

// 在 IDE 主线程刷新当前打开的易语言源码路径及关联上下文。
void UpdateCurrentOpenSourceFile();

// 返回用于查询项目级配置的源码路径；无头编译时返回原始工程路径。
std::string GetCurrentProjectConfigSourcePath();

void OutputStringToELog(const std::string& szbuf);

uint64_t AllocateAIPerfTraceId();
void SetCurrentAIPerfTraceId(uint64_t traceId);
uint64_t GetCurrentAIPerfTraceId();

bool IsAIPerfLogEnabled();
int GetAIPerfLogThresholdMs();
bool IsAICodeFetchDebugEnabled();
// 是否记录 AI 全链路调试日志（请求、响应、重试和缓存明细）。
bool IsAIDebugLogEnabled();

void LogAIPerfCost(
	uint64_t traceId,
	const std::string& step,
	long long costMs,
	const std::string& extra = std::string(),
	bool force = false);

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2);

bool BeginSilentCompileOutputPathRequest(
	const std::string& outputPath,
	DWORD ownerThreadId = 0,
	std::string* diagnostics = nullptr);

void CancelSilentCompileOutputPathRequest();
bool WasSilentCompileOutputPathRequestConsumed();
bool IsSilentCompileOutputPathRequestActive();

// 获取按当前 IDE 进程隔离的 WebView2 用户数据目录。
std::wstring GetWebView2UserDataFolderPath();

#endif // GLOBAL_H
