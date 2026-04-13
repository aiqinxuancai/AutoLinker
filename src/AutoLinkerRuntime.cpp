#include "AutoLinkerInternal.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <cstring>
#include <filesystem>
#include <format>
#include <string>

#include "Global.h"
#include "IDEFacade.h"
#include "Logger.h"
#include "PageCodeCacheManager.h"
#include "ProjectSourceCacheManager.h"
#include "WindowHelper.h"
#include "EideProjectBinarySerializer.h"

std::string g_nowOpenSourceFilePath;
ConfigManager g_configManager;
LinkerManager g_linkerManager;
ModelManager g_modelManager;
HWND g_hwnd = NULL;
bool g_preDebugging = false;
bool g_preCompiling = false;
bool g_notifySysReady = false;
bool g_uiInitialized = false;

namespace {
std::atomic_uint64_t g_aiPerfTraceSeed = 1;
thread_local uint64_t g_aiPerfTraceIdTLS = 0;
}

std::string TrimAsciiCopy(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string ToLowerAsciiCopy(const std::string& text)
{
	std::string lowered = text;
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
}

std::string TruncateForPerfLog(const std::string& text, size_t maxLen)
{
	if (text.size() <= maxLen) {
		return text;
	}
	return text.substr(0, maxLen) + "...";
}

long long ElapsedMs(const PerfClock::time_point& start)
{
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(PerfClock::now() - start).count());
}

void OutputStringToELog(const std::string& szbuf)
{
	const std::string line = "[AutoLinker]" + szbuf;
	OutputDebugStringA((line + "\n").c_str());
	IDEFacade::Instance().AppendOutputWindowLine(line);
	Logger::Instance().WriteGbk(line);
}

std::string NormalizeSourcePathForRuntime(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		return std::string();
	}

	try {
		std::filesystem::path path(trimmed);
		path = path.lexically_normal();
		if (path.is_relative()) {
			path = std::filesystem::absolute(path);
		}
		path = path.lexically_normal();
		return path.string();
	}
	catch (...) {
		return trimmed;
	}
}

bool IsProjectSourcePathForRuntime(const std::string& sourcePath)
{
	if (sourcePath.empty()) {
		return false;
	}

	try {
		std::filesystem::path path(sourcePath);
		std::string ext = path.extension().string();
		std::transform(ext.begin(), ext.end(), ext.begin(), [](unsigned char ch) {
			return static_cast<char>(std::tolower(ch));
		});
		return ext == ".e";
	}
	catch (...) {
		return false;
	}
}

bool AreSameSourcePathForRuntime(const std::string& left, const std::string& right)
{
	if (left.empty() || right.empty()) {
		return left.empty() && right.empty();
	}
	return _stricmp(left.c_str(), right.c_str()) == 0;
}

std::string DescribeSourcePathForRuntimeLog(const std::string& sourcePath)
{
	return sourcePath.empty() ? "<empty>" : sourcePath;
}

void HandleCurrentSourceFilePathChanged(const std::string& previousPath, const std::string& currentPath)
{
	e571::ProjectBinarySerializer::Instance().ClearVerifiedSerializerContext();
	project_source_cache::ProjectSourceCacheManager::Instance().Clear();
	PageCodeCacheManager::Instance().Clear();

	OutputStringToELog(std::format(
		"[SourceContext] source_changed old={} new={} serializer_context=cleared project_cache=cleared page_code_cache=cleared",
		DescribeSourcePathForRuntimeLog(previousPath),
		DescribeSourcePathForRuntimeLog(currentPath)));

	if (!IsProjectSourcePathForRuntime(currentPath)) {
		OutputStringToELog("[SourceContext] warmup skipped: current source is not an .e file");
		return;
	}

	// 检测同目录的 {stem}.AGENTS.md，存在则提示用户。
	try {
		const std::filesystem::path srcPath(currentPath);
		const std::filesystem::path agentsMdPath =
			srcPath.parent_path() / (srcPath.stem().string() + ".AGENTS.md");
		if (std::filesystem::exists(agentsMdPath)) {
			OutputStringToELog(std::format(
				"[AI] 检测到项目规范文件：{} —— 已作为 AI 对话上下文自动注入",
				agentsMdPath.filename().string()));
		}
	}
	catch (...) {}

	std::string warmupError;
	std::string warmupTrace;
	if (!project_source_cache::ProjectSourceCacheManager::Instance().WarmupCurrentSource(
			&warmupError,
			&warmupTrace)) {
		OutputStringToELog(std::format(
			"[SourceContext] warmup_failed error={} trace={}",
			warmupError.empty() ? "warmup_failed" : warmupError,
			warmupTrace));
	}
}

uint64_t AllocateAIPerfTraceId()
{
	return g_aiPerfTraceSeed.fetch_add(1);
}

void SetCurrentAIPerfTraceId(uint64_t traceId)
{
	g_aiPerfTraceIdTLS = traceId;
}

uint64_t GetCurrentAIPerfTraceId()
{
	return g_aiPerfTraceIdTLS;
}

bool IsAIPerfLogEnabled()
{
	const std::string raw = ToLowerAsciiCopy(TrimAsciiCopy(g_configManager.getValue("ai.perf_log_enabled")));
	if (raw.empty()) {
		return false;
	}
	return raw == "1" || raw == "true" || raw == "on" || raw == "yes";
}

int GetAIPerfLogThresholdMs()
{
	const std::string raw = TrimAsciiCopy(g_configManager.getValue("ai.perf_log_threshold_ms"));
	if (raw.empty()) {
		return 30;
	}
	try {
		const int parsed = std::stoi(raw);
		return (std::max)(0, (std::min)(parsed, 600000));
	}
	catch (...) {
		return 30;
	}
}

bool IsAICodeFetchDebugEnabled()
{
	const std::string raw = ToLowerAsciiCopy(TrimAsciiCopy(g_configManager.getValue("ai.code_fetch_debug_log_enabled")));
	if (raw.empty()) {
		return IsAIPerfLogEnabled();
	}
	return raw == "1" || raw == "true" || raw == "on" || raw == "yes";
}

void LogAIPerfCost(uint64_t traceId, const std::string& step, long long costMs, const std::string& extra, bool force)
{
	if (!IsAIPerfLogEnabled()) {
		return;
	}
	if (traceId == 0) {
		traceId = GetCurrentAIPerfTraceId();
	}
	if (traceId == 0) {
		return;
	}
	if (!force && costMs < static_cast<long long>(GetAIPerfLogThresholdMs())) {
		return;
	}

	std::string line = std::format("[AI-PERF] trace={} step={} cost={}ms", traceId, step, costMs);
	if (!extra.empty()) {
		line += " ";
		line += extra;
	}
	OutputStringToELog(line);
}

void UpdateCurrentOpenSourceFile()
{
	const std::string previousSourceFile = NormalizeSourcePathForRuntime(g_nowOpenSourceFilePath);
	const std::string currentSourceFile = NormalizeSourcePathForRuntime(GetSourceFilePath());
	if (AreSameSourcePathForRuntime(previousSourceFile, currentSourceFile)) {
		g_nowOpenSourceFilePath = currentSourceFile;
		return;
	}

	g_nowOpenSourceFilePath = currentSourceFile;
	HandleCurrentSourceFilePathChanged(previousSourceFile, currentSourceFile);
}

void OutputCurrentSourceLinker()
{
	UpdateCurrentOpenSourceFile();

	std::string linkerName = "默认";
	if (!g_nowOpenSourceFilePath.empty()) {
		std::string configured = g_configManager.getValue(g_nowOpenSourceFilePath);
		if (!configured.empty()) {
			linkerName = configured;
		}
	}

	OutputStringToELog("当前源码链接器：" + linkerName);
}
