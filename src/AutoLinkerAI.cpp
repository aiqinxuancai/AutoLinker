#include "AutoLinkerInternal.h"
#include <Windows.h>
#include <algorithm>
#include <atomic>
#include <cctype>
#include <chrono>
#include <cstdint>
#include <memory>
#include <new>
#include <process.h>
#include <string>
#include <vector>
#include "AIConfigDialog.h"
#include "AIService.h"
#include "Global.h"
#include "IDEFacade.h"

namespace {
std::atomic_bool g_aiTaskInProgress = false;

enum class AIAsyncUiAction {
	ReplaceCurrentFunction,
	OutputTranslation,
	InsertAtPageBottom
};

struct AIAsyncRequest {
	uint64_t traceId = 0;
	AIAsyncUiAction action = AIAsyncUiAction::ReplaceCurrentFunction;
	AITaskKind taskKind = AITaskKind::OptimizeFunction;
	AISettings settings = {};
	std::string inputText;
	std::string displayName;
	std::string pageCodeSnapshot;
	std::string sourceFunctionCode;
	int targetCaretRow = -1;
	int targetCaretCol = -1;
};

struct AIAsyncResult {
	uint64_t traceId = 0;
	AIAsyncUiAction action = AIAsyncUiAction::ReplaceCurrentFunction;
	std::string displayName;
	std::string pageCodeSnapshot;
	std::string sourceFunctionCode;
	AIResult taskResult = {};
	int targetCaretRow = -1;
	int targetCaretCol = -1;
};

struct AIApplyRequest {
	AIAsyncUiAction action = AIAsyncUiAction::ReplaceCurrentFunction;
	std::string text;
	std::string sourceFunctionCode;
	int addedCount = 0;
	int skippedCount = 0;
	int targetCaretRow = -1;
	int targetCaretCol = -1;
};
}

class ScopedAIPerfTrace {
public:
	explicit ScopedAIPerfTrace(uint64_t traceId)
		: m_prev(GetCurrentAIPerfTraceId())
	{
		SetCurrentAIPerfTraceId(traceId);
	}

	~ScopedAIPerfTrace()
	{
		SetCurrentAIPerfTraceId(m_prev);
	}

	ScopedAIPerfTrace(const ScopedAIPerfTrace&) = delete;
	ScopedAIPerfTrace& operator=(const ScopedAIPerfTrace&) = delete;

private:
	uint64_t m_prev = 0;
};

std::vector<std::string> SplitLinesNormalized(const std::string& text)
{
	std::string normalized = text;
	for (size_t pos = 0; (pos = normalized.find("\r\n", pos)) != std::string::npos; ) {
		normalized.replace(pos, 2, "\n");
		++pos;
	}

	std::vector<std::string> lines;
	size_t start = 0;
	while (start <= normalized.size()) {
		size_t end = normalized.find('\n', start);
		if (end == std::string::npos) {
			lines.push_back(normalized.substr(start));
			break;
		}
		lines.push_back(normalized.substr(start, end - start));
		start = end + 1;
	}
	return lines;
}

std::string JoinLinesCrLf(const std::vector<std::string>& lines)
{
	std::string output;
	for (const std::string& line : lines) {
		output += line;
		output += "\r\n";
	}
	return output;
}

std::string NormalizeCodeForEIDE(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}

	std::string expanded = text;
	const bool hasRealLineBreak = expanded.find('\n') != std::string::npos || expanded.find('\r') != std::string::npos;
	if (!hasRealLineBreak && (expanded.find("\\n") != std::string::npos || expanded.find("\\r\\n") != std::string::npos)) {
		std::string unescaped;
		unescaped.reserve(expanded.size());
		for (size_t i = 0; i < expanded.size(); ++i) {
			if (expanded[i] == '\\' && i + 1 < expanded.size()) {
				if (expanded[i + 1] == 'n') {
					unescaped.push_back('\n');
					++i;
					continue;
				}
				if (expanded[i + 1] == 'r') {
					if (i + 3 < expanded.size() && expanded[i + 2] == '\\' && expanded[i + 3] == 'n') {
						unescaped.push_back('\n');
						i += 3;
						continue;
					}
					unescaped.push_back('\n');
					++i;
					continue;
				}
			}
			unescaped.push_back(expanded[i]);
		}
		expanded.swap(unescaped);
	}

	std::string normalized;
	normalized.reserve(expanded.size() + 16);
	for (size_t i = 0; i < expanded.size(); ++i) {
		const char ch = expanded[i];
		if (ch == '\0') {
			continue;
		}
		if (ch == '\r') {
			if (i + 1 < expanded.size() && expanded[i + 1] == '\n') {
				++i;
			}
			normalized += "\r\n";
			continue;
		}
		if (ch == '\n') {
			normalized += "\r\n";
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

bool StartsWithAscii(const std::string& text, const std::string& prefix)
{
	return text.rfind(prefix, 0) == 0;
}

std::string ExtractNameAfterPrefix(const std::string& line, const std::string& prefix)
{
	if (!StartsWithAscii(line, prefix)) {
		return std::string();
	}
	size_t pos = prefix.size();
	while (pos < line.size() && std::isspace(static_cast<unsigned char>(line[pos])) != 0) {
		++pos;
	}
	size_t end = pos;
	while (end < line.size()) {
		const unsigned char c = static_cast<unsigned char>(line[end]);
		if (c == ',' || c == '\t' || c == ' ') {
			break;
		}
		++end;
	}
	return TrimAsciiCopy(line.substr(pos, end - pos));
}

bool IsLikelyChineseText(const std::string& text)
{
	for (unsigned char ch : text) {
		if (ch >= 0x80) {
			return true;
		}
	}
	return false;
}

std::string DescribeActivePageType(IDEFacade::ActiveWindowType type)
{
	switch (type)
	{
	case IDEFacade::ActiveWindowType::Module:
		return "程序集/类页面";
	case IDEFacade::ActiveWindowType::UserDataType:
		return "数据类型页面";
	case IDEFacade::ActiveWindowType::DllCommand:
		return "DLL命令声明页面";
	case IDEFacade::ActiveWindowType::GlobalVar:
		return "全局变量页面";
	case IDEFacade::ActiveWindowType::ConstResource:
		return "常量资源页面";
	case IDEFacade::ActiveWindowType::PictureResource:
		return "图片资源页面";
	case IDEFacade::ActiveWindowType::SoundResource:
		return "声音资源页面";
	default:
		return "未知/通用代码页面";
	}
}

std::string BuildAddByPageTypeOutputRule(IDEFacade::ActiveWindowType type)
{
	switch (type)
	{
	case IDEFacade::ActiveWindowType::Module:
		return "优先新增 .子程序（包含必要 .参数/.局部变量/实现语句），避免重复已有子程序名。";
	case IDEFacade::ActiveWindowType::DllCommand:
		return "优先新增 .DLL命令 + .参数 声明，避免重复已有命令名。";
	case IDEFacade::ActiveWindowType::UserDataType:
		return "优先新增 .数据类型 / .成员 声明，避免重复已有类型名。";
	default:
		return "根据用户需求新增最合适的易语言代码片段，保持可直接追加到页底。";
	}
}

void OutputMultiline(const std::string& title, const std::string& body)
{
	OutputStringToELog(title);
	auto lines = SplitLinesNormalized(body);
	for (const std::string& line : lines) {
		OutputStringToELog(line);
	}
}

bool EnsureAISettingsReady(AISettings& settings)
{
	AIService::LoadSettings(g_aiJsonConfig, &g_configManager, settings);
	std::string missing;
	if (AIService::HasRequiredSettings(settings, missing)) {
		return true;
	}

	OutputStringToELog("AI配置缺失，准备打开配置窗口");
	if (!ShowAIConfigDialog(g_hwnd, settings)) {
		OutputStringToELog("AI配置已取消，本次操作终止");
		return false;
	}
	AIService::SaveSettings(g_aiJsonConfig, settings);
	OutputStringToELog("AI配置已保存");
	return true;
}

bool TryBeginAiTask()
{
	bool expected = false;
	if (!g_aiTaskInProgress.compare_exchange_strong(expected, true)) {
		OutputStringToELog("[AI]已有任务在执行，请稍候");
		return false;
	}
	return true;
}

void EndAiTask()
{
	g_aiTaskInProgress.store(false);
}

void PostAiTaskResult(AIAsyncResult* result)
{
	if (result == nullptr) {
		EndAiTask();
		return;
	}

	if (g_hwnd == NULL || !IsWindow(g_hwnd) || PostMessage(g_hwnd, WM_AUTOLINKER_AI_TASK_DONE, 0, reinterpret_cast<LPARAM>(result)) == FALSE) {
		if (!result->taskResult.ok) {
			OutputStringToELog("[AI]请求失败：" + result->taskResult.error);
		}
		else {
			OutputStringToELog("[AI]任务完成，但回调消息投递失败");
		}
		delete result;
		EndAiTask();
	}
}

void PostAiApplyRequest(AIApplyRequest* request)
{
	if (request == nullptr) {
		return;
	}
	if (g_hwnd == NULL || !IsWindow(g_hwnd) || PostMessage(g_hwnd, WM_AUTOLINKER_AI_APPLY_RESULT, 0, reinterpret_cast<LPARAM>(request)) == FALSE) {
		HandleAiApplyMessage(reinterpret_cast<LPARAM>(request));
	}
}

void RunAiTaskWorker(void* pParams)
{
	std::unique_ptr<AIAsyncRequest> request(reinterpret_cast<AIAsyncRequest*>(pParams));
	if (!request) {
		EndAiTask();
		return;
	}
	ScopedAIPerfTrace traceScope(request->traceId);

	std::unique_ptr<AIAsyncResult> result(new (std::nothrow) AIAsyncResult());
	if (!result) {
		OutputStringToELog("[AI]内存不足，无法执行请求");
		EndAiTask();
		return;
	}

	result->action = request->action;
	result->displayName = request->displayName;
	result->pageCodeSnapshot = request->pageCodeSnapshot;
	result->sourceFunctionCode = request->sourceFunctionCode;
	result->targetCaretRow = request->targetCaretRow;
	result->targetCaretCol = request->targetCaretCol;
	result->traceId = request->traceId;

	try {
		const auto networkStart = PerfClock::now();
		result->taskResult = AIService::ExecuteTask(request->taskKind, request->inputText, request->settings);
		AppendAIRoundtripLogLine(
			"[RESPONSE] ok=" + std::to_string(result->taskResult.ok ? 1 : 0) +
			" httpStatus=" + std::to_string(result->taskResult.httpStatus) +
			" contentBytes=" + std::to_string(result->taskResult.content.size()) +
			" error=\"" + EscapeOneLineForLog(result->taskResult.error) + "\"");
		AppendAIRoundtripLogBlock("ai_raw_response_content", result->taskResult.content);
		LogAIPerfCost(
			request->traceId,
			"RunAiTaskWorker.execute_task_total",
			ElapsedMs(networkStart),
			"ok=" + std::to_string(result->taskResult.ok ? 1 : 0) + " http=" + std::to_string(result->taskResult.httpStatus));
	}
	catch (const std::exception& ex) {
		result->taskResult.ok = false;
		result->taskResult.error = std::string("后台任务异常：") + ex.what();
		AppendAIRoundtripLogLine(
			"[RESPONSE] exception_std what=\"" + EscapeOneLineForLog(ex.what()) + "\"");
	}
	catch (...) {
		result->taskResult.ok = false;
		result->taskResult.error = "后台任务发生未知异常";
		AppendAIRoundtripLogLine("[RESPONSE] exception_unknown");
	}

	PostAiTaskResult(result.release());
}

void RunAiFunctionReplaceTask(AITaskKind kind)
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](const std::string& status) {
		LogAIPerfCost(
			traceId,
			"RunAiFunctionReplaceTask.total",
			ElapsedMs(totalStart),
			"status=" + status + " kind=" + std::to_string(static_cast<int>(kind)));
	};

	try {
		const std::string displayName = AIService::BuildTaskDisplayName(kind);
		OutputStringToELog("[AI]开始执行：" + displayName);
		BeginAIRoundtripLogSession(traceId, "function_replace", displayName);
		AppendAIRoundtripLogLine(
			"[STATE] begin kind=" + std::to_string(static_cast<int>(kind)) +
			" displayName=\"" + EscapeOneLineForLog(displayName) + "\"");
		if (!TryBeginAiTask()) {
			AppendAIRoundtripLogLine("[STATE] abort reason=busy");
			logTotal("busy");
			return;
		}

		IDEFacade& ide = IDEFacade::Instance();
		std::string functionCode;
		std::string functionDiag;
		{
			const auto t0 = PerfClock::now();
			const bool ok = ide.GetCurrentFunctionCode(functionCode, &functionDiag);
			LogAIPerfCost(
				traceId,
				"GetCurrentFunctionCode.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " diag=" + TruncateForPerfLog(functionDiag));
			if (!ok) {
				OutputStringToELog("[AI]无法获取当前函数代码");
				if (!functionDiag.empty()) {
					OutputStringToELog("[AI]函数代码获取诊断：" + functionDiag);
				}
				AppendAIRoundtripLogLine(
					"[STATE] get_function_failed diag=\"" + EscapeOneLineForLog(functionDiag) + "\"");
				EndAiTask();
				logTotal("get_function_failed");
				return;
			}
		}
		if (functionCode.empty()) {
			OutputStringToELog("[AI]无法获取当前函数代码");
			AppendAIRoundtripLogLine("[STATE] get_function_failed reason=empty_function_code");
			EndAiTask();
			logTotal("empty_function");
			return;
		}
		int caretRow = -1;
		int caretCol = -1;
		ide.GetCaretPosition(caretRow, caretCol);
		AppendAIRoundtripLogLine(
			"[STATE] caret row=" + std::to_string(caretRow) +
			" col=" + std::to_string(caretCol) +
			" functionDiag=\"" + EscapeOneLineForLog(functionDiag) + "\"");
		AppendAIRoundtripLogBlock("source_function_code", functionCode);

		AISettings settings = {};
		{
			const auto t0 = PerfClock::now();
			const bool ok = EnsureAISettingsReady(settings);
			LogAIPerfCost(
				traceId,
				"EnsureAISettingsReady.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0));
			if (!ok) {
				AppendAIRoundtripLogLine("[STATE] abort reason=settings_not_ready");
				EndAiTask();
				logTotal("settings_not_ready");
				return;
			}
		}
		AppendAIRoundtripLogLine(
			"[SETTINGS] baseUrl=\"" + EscapeOneLineForLog(settings.baseUrl) +
			"\" model=\"" + EscapeOneLineForLog(settings.model) +
			"\" timeoutMs=" + std::to_string(settings.timeoutMs) +
			" temperature=" + std::format("{:.3f}", settings.temperature));

		const std::string userInput =
			"请处理以下易语言函数代码，并严格返回完整可替换代码：\n```e\n" +
			functionCode + "\n```";
		std::string finalInput = userInput;
		if (kind == AITaskKind::OptimizeFunction) {
			std::string extraRequirement;
			const auto inputStart = PerfClock::now();
			const bool accepted = ShowAITextInputDialog(
				g_hwnd,
				"AI优化函数 - 输入优化要求",
				"请输入你希望优化的方向（可为空）：",
				extraRequirement);
			LogAIPerfCost(
				traceId,
				"ShowAITextInputDialog.modal",
				ElapsedMs(inputStart),
				"ok=" + std::to_string(accepted ? 1 : 0) + " scene=optimize");
			if (!accepted) {
				OutputStringToELog("[AI]已取消输入优化要求");
				AppendAIRoundtripLogLine("[STATE] user_cancel_optimize_requirement");
				EndAiTask();
				logTotal("user_cancel_optimize_requirement");
				return;
			}
			extraRequirement = TrimAsciiCopy(extraRequirement);
			AppendAIRoundtripLogLine(
				"[INPUT] optimize_requirement bytes=" + std::to_string(extraRequirement.size()) +
				" text=\"" + EscapeOneLineForLog(extraRequirement) + "\"");
			if (!extraRequirement.empty()) {
				finalInput = std::string("额外优化要求：\n") + extraRequirement + "\n\n" + userInput;
			}
		}
		AppendAIRoundtripLogBlock("ai_input_prompt", finalInput);

		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			AppendAIRoundtripLogLine("[STATE] abort reason=alloc_request_failed");
			EndAiTask();
			logTotal("alloc_request_failed");
			return;
		}
		request->traceId = traceId;
		request->action = AIAsyncUiAction::ReplaceCurrentFunction;
		request->taskKind = kind;
		request->settings = settings;
		request->inputText = finalInput;
		request->displayName = displayName;
		request->sourceFunctionCode = functionCode;
		request->targetCaretRow = caretRow;
		request->targetCaretCol = caretCol;
		AppendAIRoundtripLogLine(
			"[REQUEST] queued action=ReplaceCurrentFunction trace=" + std::to_string(traceId) +
			" inputBytes=" + std::to_string(request->inputText.size()) +
			" sourceBytes=" + std::to_string(request->sourceFunctionCode.size()));

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			AppendAIRoundtripLogLine("[STATE] abort reason=start_worker_failed");
			EndAiTask();
			logTotal("start_worker_failed");
			return;
		}
		request.release();
		AppendAIRoundtripLogLine("[STATE] worker_started");
		logTotal("queued");
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		AppendAIRoundtripLogLine(
			"[STATE] exception_std what=\"" + EscapeOneLineForLog(ex.what()) + "\"");
		EndAiTask();
		logTotal("exception_std");
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		AppendAIRoundtripLogLine("[STATE] exception_unknown");
		EndAiTask();
		logTotal("exception_unknown");
	}
}

void RunAiTranslateSelectedTextTask()
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](const std::string& status) {
		LogAIPerfCost(
			traceId,
			"RunAiTranslateSelectedTextTask.total",
			ElapsedMs(totalStart),
			"status=" + status);
	};

	try {
		OutputStringToELog("[AI]开始执行：AI翻译选中文本");
		if (!TryBeginAiTask()) {
			logTotal("busy");
			return;
		}
		IDEFacade& ide = IDEFacade::Instance();
		std::string selectedText;
		if (!ide.GetSelectedText(selectedText)) {
			OutputStringToELog("[AI]未检测到有效选中文本");
			EndAiTask();
			logTotal("no_selection");
			return;
		}

		AISettings settings = {};
		{
			const auto t0 = PerfClock::now();
			const bool ok = EnsureAISettingsReady(settings);
			LogAIPerfCost(
				traceId,
				"EnsureAISettingsReady.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " scene=translate_selected");
			if (!ok) {
				EndAiTask();
				logTotal("settings_not_ready");
				return;
			}
		}

		const bool sourceIsChinese = IsLikelyChineseText(selectedText);
		const std::string direction = sourceIsChinese ? "请把以下中文翻译为英文，仅输出翻译结果：" : "请把以下英文翻译为中文，仅输出翻译结果：";
		const std::string input = direction + "\n" + selectedText;
		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			logTotal("alloc_request_failed");
			return;
		}
		request->traceId = traceId;
		request->action = AIAsyncUiAction::OutputTranslation;
		request->taskKind = AITaskKind::TranslateText;
		request->settings = settings;
		request->inputText = input;
		request->displayName = "AI翻译选中文本";

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			logTotal("start_worker_failed");
			return;
		}
		request.release();
		logTotal("queued");
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
		logTotal("exception_std");
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
		logTotal("exception_unknown");
	}
}

void HandleAiTaskCompletionMessage(LPARAM lParam)
{
	struct AiTaskEndGuard {
		~AiTaskEndGuard() { EndAiTask(); }
	} endGuard;

	std::unique_ptr<AIAsyncResult> result(reinterpret_cast<AIAsyncResult*>(lParam));
	if (!result) {
		return;
	}

	try {
		if (!result->taskResult.ok) {
			OutputStringToELog("[AI]请求失败：" + result->taskResult.error);
			AppendAIRoundtripLogLine(
				"[STATE] request_failed error=\"" + EscapeOneLineForLog(result->taskResult.error) + "\"");
			return;
		}

		switch (result->action)
		{
		case AIAsyncUiAction::ReplaceCurrentFunction: {
			std::string generatedCode = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			generatedCode = AIService::Trim(generatedCode);
			if (generatedCode.empty()) {
				OutputStringToELog("[AI]模型返回为空");
				AppendAIRoundtripLogLine("[STATE] empty_generated_code_after_normalize");
				return;
			}
			generatedCode = NormalizeCodeForEIDE(generatedCode);
			AppendAIRoundtripLogBlock("ai_output_normalized_code", generatedCode);

			const std::string title = result->displayName.empty() ? "AI结果预览" : (result->displayName + " - 结果预览");
			const AIPreviewAction previewAction = ShowAIPreviewDialogEx(
				g_hwnd,
				title,
				generatedCode,
				"复制到剪贴板",
				"替换（不稳定）");
			if (previewAction == AIPreviewAction::Cancel) {
				OutputStringToELog("[AI]用户取消应用结果");
				AppendAIRoundtripLogLine("[PREVIEW] action=cancel");
				return;
			}
			if (previewAction == AIPreviewAction::SecondaryConfirm) {
				std::unique_ptr<AIApplyRequest> request(new (std::nothrow) AIApplyRequest());
				if (!request) {
					OutputStringToELog("[AI]内存不足，无法执行替换");
					AppendAIRoundtripLogLine("[PREVIEW] action=replace_unstable alloc_request_failed");
					return;
				}
				request->action = AIAsyncUiAction::ReplaceCurrentFunction;
				request->text = generatedCode;
				request->sourceFunctionCode = result->sourceFunctionCode;
				request->targetCaretRow = result->targetCaretRow;
				request->targetCaretCol = result->targetCaretCol;
				OutputStringToELog("[AI]用户选择：替换（不稳定）");
				AppendAIRoundtripLogLine(
					"[PREVIEW] action=replace_unstable targetCaretRow=" + std::to_string(request->targetCaretRow) +
					" targetCaretCol=" + std::to_string(request->targetCaretCol));
				AppendAIRoundtripLogBlock("replace_request_payload", request->text);
				PostAiApplyRequest(request.release());
				return;
			}

			OutputStringToELog("[AI]用户选择：复制到剪贴板");
			AppendAIRoundtripLogLine("[PREVIEW] action=copy_to_clipboard");
			AppendAIRoundtripLogBlock("clipboard_payload", generatedCode);
			if (!IDEFacade::Instance().SetClipboardText(generatedCode)) {
				OutputStringToELog("[AI]复制到剪贴板失败");
				AppendAIRoundtripLogLine("[STATE] copy_to_clipboard_failed");
				return;
			}
			OutputStringToELog("[AI]已复制到剪贴板，请手动替换");
			AppendAIRoundtripLogLine("[STATE] copy_to_clipboard_done");
			return;
		}
		case AIAsyncUiAction::OutputTranslation: {
			std::string translated = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			translated = AIService::Trim(translated);
			if (translated.empty()) {
				OutputStringToELog("[AI]翻译结果为空");
				return;
			}
			OutputMultiline("[AI]翻译结果：", translated);
			return;
		}
		case AIAsyncUiAction::InsertAtPageBottom: {
			std::string generatedText = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			generatedText = AIService::Trim(generatedText);
			if (generatedText.empty()) {
				OutputStringToELog("[AI]模型未返回可新增内容");
				return;
			}
			generatedText = NormalizeCodeForEIDE(generatedText);

			if (!ShowAIPreviewDialog(g_hwnd, "AI按页类型添加代码 - 结果预览", generatedText, "插入")) {
				OutputStringToELog("[AI]用户取消插入");
				return;
			}

			std::unique_ptr<AIApplyRequest> request(new (std::nothrow) AIApplyRequest());
			if (!request) {
				OutputStringToELog("[AI]内存不足，无法执行插入");
				return;
			}
			request->action = AIAsyncUiAction::InsertAtPageBottom;
			request->text = generatedText;
			PostAiApplyRequest(request.release());
			return;
		}
		default:
			OutputStringToELog("[AI]未知任务结果类型");
			return;
		}
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]处理结果异常：") + ex.what());
	}
	catch (...) {
		OutputStringToELog("[AI]处理结果发生未知异常");
	}
}

void HandleAiApplyMessage(LPARAM lParam)
{
	std::unique_ptr<AIApplyRequest> request(reinterpret_cast<AIApplyRequest*>(lParam));
	if (!request) {
		return;
	}

	try {
		IDEFacade& ide = IDEFacade::Instance();
		constexpr bool kAiApplyPreCompile = false;
		switch (request->action)
		{
		case AIAsyncUiAction::ReplaceCurrentFunction: {
			if (g_hwnd != NULL && IsWindow(g_hwnd)) {
				SetForegroundWindow(g_hwnd);
			}
			if (request->targetCaretRow >= 0 && request->targetCaretCol >= 0) {
				ide.MoveCaret(request->targetCaretRow, request->targetCaretCol);
			}
			AppendAIRoundtripLogLine(
				"[APPLY] begin action=ReplaceCurrentFunction targetCaretRow=" + std::to_string(request->targetCaretRow) +
				" targetCaretCol=" + std::to_string(request->targetCaretCol));
			AppendAIRoundtripLogBlock("apply_request_text_raw", request->text);
			const std::string newFunctionCode = NormalizeCodeForEIDE(request->text);
			AppendAIRoundtripLogBlock("apply_request_text_normalized", newFunctionCode);
			if (ide.ReplaceCurrentFunctionCode(newFunctionCode, kAiApplyPreCompile)) {
				OutputStringToELog("[AI]替换完成");
				AppendAIRoundtripLogLine("[APPLY] result=ok");
				return;
			}
			OutputStringToELog("[AI]替换当前函数失败（当前坐标未定位到可替换子程序）");
			AppendAIRoundtripLogLine("[APPLY] result=failed reason=replace_current_function_failed");
			return;
		}

		case AIAsyncUiAction::InsertAtPageBottom:
			if (g_hwnd != NULL && IsWindow(g_hwnd)) {
				SetForegroundWindow(g_hwnd);
			}
			if (!ide.InsertCodeAtPageBottom(request->text, kAiApplyPreCompile)) {
				OutputStringToELog("[AI]追加到页底失败");
				return;
			}
			OutputStringToELog("[AI]追加到页底完成");
			return;

		default:
			return;
		}
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]执行结果异常：") + ex.what());
	}
	catch (...) {
		OutputStringToELog("[AI]执行结果发生未知异常");
	}
}

void RunAiAddByCurrentPageTypeTask()
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](const std::string& status) {
		LogAIPerfCost(
			traceId,
			"RunAiAddByCurrentPageTypeTask.total",
			ElapsedMs(totalStart),
			"status=" + status);
	};

	try {
		OutputStringToELog("[AI]开始执行：AI按当前页类型添加代码");
		if (!TryBeginAiTask()) {
			logTotal("busy");
			return;
		}

		IDEFacade& ide = IDEFacade::Instance();
		std::string currentPageCode;
		{
			const auto t0 = PerfClock::now();
			const bool ok = ide.GetCurrentPageCode(currentPageCode);
			LogAIPerfCost(
				traceId,
				"GetCurrentPageCode.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " scene=add_by_page");
			if (!ok) {
				OutputStringToELog("[AI]无法获取当前页代码");
				EndAiTask();
				logTotal("get_page_failed");
				return;
			}
		}
		if (currentPageCode.empty()) {
			OutputStringToELog("[AI]无法获取当前页代码");
			EndAiTask();
			logTotal("empty_page_code");
			return;
		}

		const IDEFacade::ActiveWindowType pageType = ide.GetActiveWindowType();
		const std::string pageTypeText = DescribeActivePageType(pageType);
		const std::string outputRule = BuildAddByPageTypeOutputRule(pageType);

		std::string requirement;
		const auto inputStart = PerfClock::now();
		const bool accepted = ShowAITextInputDialog(
			g_hwnd,
			"AI按当前页类型添加代码",
			"请输入添加需求（例如：数量、命名风格、用途、限制）：",
			requirement);
		LogAIPerfCost(
			traceId,
			"ShowAITextInputDialog.modal",
			ElapsedMs(inputStart),
			"ok=" + std::to_string(accepted ? 1 : 0) + " scene=add_by_page");
		if (!accepted) {
			OutputStringToELog("[AI]已取消输入添加需求");
			EndAiTask();
			logTotal("user_cancel_requirement");
			return;
		}
		requirement = TrimAsciiCopy(requirement);
		if (requirement.empty()) {
			OutputStringToELog("[AI]添加需求为空，已取消");
			EndAiTask();
			logTotal("empty_requirement");
			return;
		}

		AISettings settings = {};
		{
			const auto t0 = PerfClock::now();
			const bool ok = EnsureAISettingsReady(settings);
			LogAIPerfCost(
				traceId,
				"EnsureAISettingsReady.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " scene=add_by_page");
			if (!ok) {
				EndAiTask();
				logTotal("settings_not_ready");
				return;
			}
		}

		const std::string input =
			"当前页类型：" + pageTypeText + "\n"
			"按页类型输出规则：" + outputRule + "\n"
			"输出硬性要求：仅返回新增信息，不要返回整页、不解释。\n"
			"用户添加需求：\n" + requirement + "\n\n"
			"当前页完整代码：\n```e\n" + currentPageCode + "\n```";

		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			logTotal("alloc_request_failed");
			return;
		}
		request->traceId = traceId;
		request->action = AIAsyncUiAction::InsertAtPageBottom;
		request->taskKind = AITaskKind::AddByCurrentPageType;
		request->settings = settings;
		request->inputText = input;
		request->displayName = "AI按当前页类型添加代码";
		request->pageCodeSnapshot = currentPageCode;

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			logTotal("start_worker_failed");
			return;
		}
		request.release();
		logTotal("queued");
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
		logTotal("exception_std");
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
		logTotal("exception_unknown");
	}
}

void TryCopyCurrentFunctionCode()
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	const bool ok = IDEFacade::Instance().CopyCurrentFunctionCodeToClipboard();
	LogAIPerfCost(
		traceId,
		"TryCopyCurrentFunctionCode.total",
		ElapsedMs(totalStart),
		"ok=" + std::to_string(ok ? 1 : 0));

	if (ok) {
		OutputStringToELog("已复制当前函数代码到剪贴板");
	}
	else {
		OutputStringToELog("复制当前函数代码失败，当前位置可能不在代码函数中");
	}
}

