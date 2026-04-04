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

void OutputStringToELog(const std::string& szbuf);

uint64_t AllocateAIPerfTraceId();
void SetCurrentAIPerfTraceId(uint64_t traceId);
uint64_t GetCurrentAIPerfTraceId();

bool IsAIPerfLogEnabled();
int GetAIPerfLogThresholdMs();
bool IsAICodeFetchDebugEnabled();

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

bool BeginProjectSnapshotSaveRequest(
	const std::string& sourcePath,
	const std::string& snapshotPath,
	DWORD ownerThreadId = 0,
	std::string* diagnostics = nullptr);

void CancelProjectSnapshotSaveRequest();
bool WasProjectSnapshotSaveRequestConsumed();

// 获取 WebView2 用户数据目录。
std::wstring GetWebView2UserDataFolderPath();

#endif // GLOBAL_H
