// global.h
#pragma once
#include "ModelManager.h"
#include <cstdint>
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
void OutputStringToELog(const std::string& szbuf);



/// <summary>
/// 分配一次 AI 操作的性能跟踪 ID。
/// </summary>
uint64_t AllocateAIPerfTraceId();

/// <summary>
/// 设置/获取当前线程的 AI 性能跟踪 ID。
/// </summary>
void SetCurrentAIPerfTraceId(uint64_t traceId);
uint64_t GetCurrentAIPerfTraceId();

/// <summary>
/// AI 性能日志开关与阈值（来自配置）。
/// ai.perf_log_enabled: 0/1，默认0
/// ai.perf_log_threshold_ms: 默认30
/// </summary>
bool IsAIPerfLogEnabled();
int GetAIPerfLogThresholdMs();
bool IsAICodeFetchDebugEnabled();

/// <summary>
/// 输出统一格式的 AI 性能日志。
/// </summary>
void LogAIPerfCost(
	uint64_t traceId,
	const std::string& step,
	long long costMs,
	const std::string& extra = std::string(),
	bool force = false);
/// <summary>
/// 运行插件通知
/// </summary>
/// <param name="code"></param>
/// <param name="p1"></param>
/// <param name="p2"></param>
/// <returns></returns>
INT NESRUNFUNC(INT code, DWORD p1, DWORD p2);

#endif // GLOBAL_H