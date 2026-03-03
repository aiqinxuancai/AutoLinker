#include "AutoLinker.h"
#include <vector>
#include <Windows.h>
#include <CommCtrl.h>
#include <format>
#include <algorithm>
#include <atomic>
#include <cctype>
#include "ConfigManager.h"
#include <regex>
#include "PathHelper.h"
#include "LinkerManager.h"
#include "InlineHook.h"
#include "ModelManager.h"
#include "Global.h"
#include "StringHelper.h"
#include <thread>
#include "MouseBack.h"
#include <PublicIDEFunctions.h>
#include "ECOMEx.h"
#include "WindowHelper.h"
#include <future>
#include "WinINetUtil.h"
#include "Version.h"
#include "IDEFacade.h"
#include "AIService.h"
#include "AIConfigDialog.h"
#include <memory>
#include <new>
#include <process.h>
#include <unordered_map>
#include <unordered_set>

#pragma comment(lib, "comctl32.lib")

//当前打开的易源文件路径
std::string g_nowOpenSourceFilePath;

//配置文件，当前路径的e源码对应的编译器名
ConfigManager g_configManager;

//管理当前所有的link文件
LinkerManager g_linkerManager;

//管理模块 调试 -> 编译 的管理器
ModelManager g_modelManager;


//e主窗口句柄
HWND g_hwnd = NULL;

//工具条句柄
HWND g_toolBarHwnd = NULL;

//准备开始调试 废弃
bool g_preDebugging;

//准备开始编译 废弃
bool g_preCompiling;

bool g_initStarted = false;

HMENU g_topLinkerSubMenu = NULL;
std::unordered_map<UINT, std::string> g_topLinkerCommandMap;

void UpdateCurrentOpenSourceFile();
void OutputCurrentSourceLinker();

namespace {
bool g_isContextMenuRegistered = false;
constexpr UINT IDM_AUTOLINKER_CTX_COPY_FUNC = 31001; // Keep < 0x8000 to avoid signed-ID issues in some hosts.
constexpr UINT IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC = 31101;
constexpr UINT IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC = 31102;
constexpr UINT IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC = 31103;
constexpr UINT IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT = 31104;
constexpr UINT IDM_AUTOLINKER_CTX_AI_COMPLETE_DECL = 31105;
constexpr UINT IDM_AUTOLINKER_LINKER_BASE = 34000;
constexpr UINT IDM_AUTOLINKER_LINKER_MAX = 34999;
constexpr UINT WM_AUTOLINKER_AI_TASK_DONE = WM_USER + 1001;
constexpr UINT WM_AUTOLINKER_AI_APPLY_RESULT = WM_USER + 1002;

std::atomic_bool g_aiTaskInProgress = false;

enum class AIAsyncUiAction {
	ReplaceCurrentFunction,
	OutputTranslation,
	InsertDeclarations
};

struct AIAsyncRequest {
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

void HandleAiTaskCompletionMessage(LPARAM lParam);
void HandleAiApplyMessage(LPARAM lParam);

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

std::string ToLowerAsciiCopy(const std::string& text)
{
	std::string lowered = text;
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
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

void OutputMultiline(const std::string& title, const std::string& body)
{
	OutputStringToELog(title);
	auto lines = SplitLinesNormalized(body);
	for (const std::string& line : lines) {
		IDEFacade::Instance().AppendOutputWindowLine(line);
	}
}

bool EnsureAISettingsReady(AISettings& settings)
{
	AIService::LoadSettings(g_configManager, settings);
	std::string missing;
	if (AIService::HasRequiredSettings(settings, missing)) {
		return true;
	}

	OutputStringToELog("AI配置缺失，准备打开配置窗口");
	if (!ShowAIConfigDialog(g_hwnd, settings)) {
		OutputStringToELog("AI配置已取消，本次操作终止");
		return false;
	}
	AIService::SaveSettings(g_configManager, settings);
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

	try {
		result->taskResult = AIService::ExecuteTask(request->taskKind, request->inputText, request->settings);
	}
	catch (const std::exception& ex) {
		result->taskResult.ok = false;
		result->taskResult.error = std::string("后台任务异常：") + ex.what();
	}
	catch (...) {
		result->taskResult.ok = false;
		result->taskResult.error = "后台任务发生未知异常";
	}

	PostAiTaskResult(result.release());
}

void RunAiFunctionReplaceTask(AITaskKind kind)
{
	try {
		const std::string displayName = AIService::BuildTaskDisplayName(kind);
		OutputStringToELog("[AI]开始执行：" + displayName);
		if (!TryBeginAiTask()) {
			return;
		}

		IDEFacade& ide = IDEFacade::Instance();
		std::string functionCode;
		std::string functionDiag;
		if (!ide.GetCurrentFunctionCode(functionCode, &functionDiag)) {
			OutputStringToELog("[AI]无法获取当前函数代码");
			if (!functionDiag.empty()) {
				OutputStringToELog("[AI]函数代码获取诊断：" + functionDiag);
			}
			EndAiTask();
			return;
		}
		int caretRow = -1;
		int caretCol = -1;
		ide.GetCaretPosition(caretRow, caretCol);

		AISettings settings = {};
		if (!EnsureAISettingsReady(settings)) {
			EndAiTask();
			return;
		}

		const std::string userInput =
			"请处理以下易语言函数代码，并严格返回完整可替换代码：\n```e\n" +
			functionCode + "\n```";
		std::string finalInput = userInput;
		if (kind == AITaskKind::OptimizeFunction) {
			std::string extraRequirement;
			if (!ShowAITextInputDialog(g_hwnd, "AI优化函数 - 输入优化要求", "请输入你希望优化的方向（可为空）：", extraRequirement)) {
				OutputStringToELog("[AI]已取消输入优化要求");
				EndAiTask();
				return;
			}
			extraRequirement = TrimAsciiCopy(extraRequirement);
			if (!extraRequirement.empty()) {
				finalInput = std::string("额外优化要求：\n") + extraRequirement + "\n\n" + userInput;
			}
		}

		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			return;
		}
		request->action = AIAsyncUiAction::ReplaceCurrentFunction;
		request->taskKind = kind;
		request->settings = settings;
		request->inputText = finalInput;
		request->displayName = displayName;
		request->sourceFunctionCode = functionCode;
		request->targetCaretRow = caretRow;
		request->targetCaretCol = caretCol;

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			return;
		}
		request.release();
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
	}
}

void RunAiTranslateSelectedTextTask()
{
	try {
		OutputStringToELog("[AI]开始执行：AI翻译文本");
		if (!TryBeginAiTask()) {
			return;
		}
		IDEFacade& ide = IDEFacade::Instance();
		std::string selectedText;
		if (!ide.GetSelectedText(selectedText)) {
			OutputStringToELog("[AI]未检测到有效选中文本");
			EndAiTask();
			return;
		}

		AISettings settings = {};
		if (!EnsureAISettingsReady(settings)) {
			EndAiTask();
			return;
		}

		const bool sourceIsChinese = IsLikelyChineseText(selectedText);
		const std::string direction = sourceIsChinese ? "请把以下中文翻译为英文，仅输出翻译结果：" : "请把以下英文翻译为中文，仅输出翻译结果：";
		const std::string input = direction + "\n" + selectedText;
		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			return;
		}
		request->action = AIAsyncUiAction::OutputTranslation;
		request->taskKind = AITaskKind::TranslateText;
		request->settings = settings;
		request->inputText = input;
		request->displayName = "AI翻译文本";

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			return;
		}
		request.release();
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
	}
}

struct DeclarationBlock {
	bool isDllCommand = false;
	std::string name;
	std::vector<std::string> lines;
};

void CollectExistingDeclarationNames(
	const std::string& pageCode,
	std::unordered_set<std::string>& dllCommands,
	std::unordered_set<std::string>& dataTypes)
{
	for (const std::string& rawLine : SplitLinesNormalized(pageCode)) {
		const std::string line = TrimAsciiCopy(rawLine);
		std::string name = ExtractNameAfterPrefix(line, ".DLL命令");
		if (!name.empty()) {
			dllCommands.insert(ToLowerAsciiCopy(name));
			continue;
		}
		name = ExtractNameAfterPrefix(line, ".数据类型");
		if (!name.empty()) {
			dataTypes.insert(ToLowerAsciiCopy(name));
		}
	}
}

std::vector<DeclarationBlock> ParseGeneratedDeclarationBlocks(const std::string& generatedText)
{
	std::vector<DeclarationBlock> blocks;
	DeclarationBlock current = {};
	bool inBlock = false;

	auto flushCurrent = [&]() {
		if (inBlock && !current.name.empty() && !current.lines.empty()) {
			blocks.push_back(current);
		}
		current = {};
		inBlock = false;
	};

	for (const std::string& rawLine : SplitLinesNormalized(generatedText)) {
		const std::string line = TrimAsciiCopy(rawLine);
		if (line.empty()) {
			if (inBlock) {
				current.lines.push_back("");
			}
			continue;
		}

		if (StartsWithAscii(line, ".版本")) {
			continue;
		}

		const std::string dllName = ExtractNameAfterPrefix(line, ".DLL命令");
		if (!dllName.empty()) {
			flushCurrent();
			inBlock = true;
			current.isDllCommand = true;
			current.name = dllName;
			current.lines.push_back(line);
			continue;
		}

		const std::string typeName = ExtractNameAfterPrefix(line, ".数据类型");
		if (!typeName.empty()) {
			flushCurrent();
			inBlock = true;
			current.isDllCommand = false;
			current.name = typeName;
			current.lines.push_back(line);
			continue;
		}

		if (inBlock) {
			current.lines.push_back(line);
		}
	}

	flushCurrent();
	return blocks;
}

std::string BuildUniqueDeclarationInsertText(const std::string& currentPageCode, const std::string& generatedDeclText, int& addedCount, int& skippedCount)
{
	addedCount = 0;
	skippedCount = 0;

	std::unordered_set<std::string> existingDlls;
	std::unordered_set<std::string> existingTypes;
	CollectExistingDeclarationNames(currentPageCode, existingDlls, existingTypes);

	std::vector<DeclarationBlock> generatedBlocks = ParseGeneratedDeclarationBlocks(generatedDeclText);
	std::vector<std::string> outputLines;

	for (const DeclarationBlock& block : generatedBlocks) {
		const std::string key = ToLowerAsciiCopy(block.name);
		bool duplicated = false;
		if (block.isDllCommand) {
			duplicated = existingDlls.contains(key);
			if (!duplicated) {
				existingDlls.insert(key);
			}
		}
		else {
			duplicated = existingTypes.contains(key);
			if (!duplicated) {
				existingTypes.insert(key);
			}
		}

		if (duplicated) {
			++skippedCount;
			continue;
		}

		++addedCount;
		outputLines.insert(outputLines.end(), block.lines.begin(), block.lines.end());
		outputLines.push_back("");
	}

	return JoinLinesCrLf(outputLines);
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
			return;
		}

		switch (result->action)
		{
		case AIAsyncUiAction::ReplaceCurrentFunction: {
			std::string generatedCode = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			generatedCode = AIService::Trim(generatedCode);
			if (generatedCode.empty()) {
				OutputStringToELog("[AI]模型返回为空");
				return;
			}
			generatedCode = NormalizeCodeForEIDE(generatedCode);

			const std::string title = result->displayName.empty() ? "AI结果预览" : (result->displayName + " - 结果预览");
			if (!ShowAIPreviewDialog(g_hwnd, title, generatedCode, "替换")) {
				OutputStringToELog("[AI]用户取消替换");
				return;
			}
			std::unique_ptr<AIApplyRequest> request(new (std::nothrow) AIApplyRequest());
			if (!request) {
				OutputStringToELog("[AI]内存不足，无法执行替换");
				return;
			}
			request->action = AIAsyncUiAction::ReplaceCurrentFunction;
			request->text = generatedCode;
			request->sourceFunctionCode = result->sourceFunctionCode;
			request->targetCaretRow = result->targetCaretRow;
			request->targetCaretCol = result->targetCaretCol;
			PostAiApplyRequest(request.release());
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
		case AIAsyncUiAction::InsertDeclarations: {
			std::string generatedDecl = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			generatedDecl = AIService::Trim(generatedDecl);
			if (generatedDecl.empty()) {
				OutputStringToELog("[AI]模型未返回声明代码");
				return;
			}

			int addedCount = 0;
			int skippedCount = 0;
			std::string insertText = BuildUniqueDeclarationInsertText(result->pageCodeSnapshot, generatedDecl, addedCount, skippedCount);
			insertText = AIService::Trim(insertText);
			if (insertText.empty() || addedCount <= 0) {
				OutputStringToELog("[AI]未发现可新增声明（可能均已存在）");
				return;
			}
			insertText = NormalizeCodeForEIDE(insertText);

			if (!ShowAIPreviewDialog(g_hwnd, "AI补全API声明 - 结果预览", insertText, "插入")) {
				OutputStringToELog("[AI]用户取消插入");
				return;
			}
			std::unique_ptr<AIApplyRequest> request(new (std::nothrow) AIApplyRequest());
			if (!request) {
				OutputStringToELog("[AI]内存不足，无法执行插入");
				return;
			}
			request->action = AIAsyncUiAction::InsertDeclarations;
			request->text = insertText;
			request->addedCount = addedCount;
			request->skippedCount = skippedCount;
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
			const std::string newFunctionCode = NormalizeCodeForEIDE(request->text);
			if (ide.ReplaceCurrentFunctionCode(newFunctionCode, kAiApplyPreCompile)) {
				OutputStringToELog("[AI]替换完成");
				return;
			}
			OutputStringToELog("[AI]替换当前函数失败（当前坐标未定位到可替换子程序）");
			return;
		}

		case AIAsyncUiAction::InsertDeclarations:
			if (!ide.InsertCodeAtPageTop(request->text, kAiApplyPreCompile)) {
				OutputStringToELog("[AI]插入声明失败");
				return;
			}
			OutputStringToELog(std::format("[AI]插入完成：新增 {} 项，跳过 {} 项", request->addedCount, request->skippedCount));
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

void RunAiCompleteDeclarationsTask()
{
	try {
		OutputStringToELog("[AI]开始执行：AI补全API声明");
		if (!TryBeginAiTask()) {
			return;
		}
		IDEFacade& ide = IDEFacade::Instance();

		std::string functionCode;
		std::string functionDiag;
		if (!ide.GetCurrentFunctionCode(functionCode, &functionDiag)) {
			OutputStringToELog("[AI]无法获取当前函数代码");
			if (!functionDiag.empty()) {
				OutputStringToELog("[AI]函数代码获取诊断：" + functionDiag);
			}
			EndAiTask();
			return;
		}

		std::string currentPageCode;
		if (!ide.GetCurrentPageCode(currentPageCode)) {
			OutputStringToELog("[AI]无法获取当前页代码");
			EndAiTask();
			return;
		}

		AISettings settings = {};
		if (!EnsureAISettingsReady(settings)) {
			EndAiTask();
			return;
		}

		const std::string userInput =
			"请根据以下易语言函数，补全缺失的 Windows API 声明和必要结构体声明。\n"
			"只返回 .DLL命令/.参数/.数据类型/.成员 声明代码，不返回函数实现。\n"
			"函数代码：\n```e\n" + functionCode + "\n```";
		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			return;
		}
		request->action = AIAsyncUiAction::InsertDeclarations;
		request->taskKind = AITaskKind::CompleteApiDeclarations;
		request->settings = settings;
		request->inputText = userInput;
		request->displayName = "AI补全API声明";
		request->pageCodeSnapshot = currentPageCode;

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			return;
		}
		request.release();
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
	}
}

void TryCopyCurrentFunctionCode()
{
	if (IDEFacade::Instance().CopyCurrentFunctionCodeToClipboard()) {
		OutputStringToELog("已复制当前函数代码到剪贴板");
	}
	else {
		OutputStringToELog("复制当前函数代码失败，当前位置可能不在代码函数中");
	}
}

void RegisterIDEContextMenu()
{
	if (g_isContextMenuRegistered) {
		return;
	}

	auto& ide = IDEFacade::Instance();
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_COPY_FUNC, "复制当前函数代码", []() {
		TryCopyCurrentFunctionCode();
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC, "AI优化函数", []() {
		RunAiFunctionReplaceTask(AITaskKind::OptimizeFunction);
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC, "AI为当前函数添加注释", []() {
		RunAiFunctionReplaceTask(AITaskKind::AddCommentsToFunction);
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC, "AI翻译当前函数+变量名", []() {
		RunAiFunctionReplaceTask(AITaskKind::TranslateFunctionAndVariables);
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT, "AI翻译文本", []() {
		RunAiTranslateSelectedTextTask();
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_COMPLETE_DECL, "AI为当前函数补全API声明", []() {
		RunAiCompleteDeclarationsTask();
	});

	g_isContextMenuRegistered = true;
}

std::wstring GetMenuTitleW(HMENU hMenu, UINT item, UINT flags)
{
	wchar_t title[256] = { 0 };
	int len = GetMenuStringW(hMenu, item, title, static_cast<int>(sizeof(title) / sizeof(title[0])), flags);
	if (len <= 0) {
		return L"";
	}
	return std::wstring(title, static_cast<size_t>(len));
}

bool IsCompileOrToolsTopPopup(HMENU hPopupMenu)
{
	if (g_hwnd == NULL || hPopupMenu == NULL) {
		return false;
	}
	if (hPopupMenu == g_topLinkerSubMenu) {
		return false;
	}

	// 右键菜单里会包含该命令，直接排除，避免“链接器切换”混入右键。
	UINT copyState = GetMenuState(hPopupMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND);
	if (copyState != 0xFFFFFFFF) {
		return false;
	}

	auto popupKeywordMatch = [hPopupMenu]() -> bool {
		int keywordHit = 0;
		int itemCount = GetMenuItemCount(hPopupMenu);
		for (int item = 0; item < itemCount; ++item) {
			std::wstring itemTitle = GetMenuTitleW(hPopupMenu, static_cast<UINT>(item), MF_BYPOSITION);
			if (itemTitle.find(L"编译") != std::wstring::npos ||
				itemTitle.find(L"发布") != std::wstring::npos ||
				itemTitle.find(L"静态") != std::wstring::npos ||
				itemTitle.find(L"Build") != std::wstring::npos ||
				itemTitle.find(L"Release") != std::wstring::npos) {
				++keywordHit;
			}
		}
		return keywordHit >= 1;
	};

	HMENU hMainMenu = GetMenu(g_hwnd);
	if (hMainMenu == NULL) {
		return popupKeywordMatch();
	}

	int count = GetMenuItemCount(hMainMenu);
	int popupIndex = -1;
	int compileIndex = -1;
	int toolsIndex = -1;

	for (int i = 0; i < count; ++i) {
		HMENU subMenu = GetSubMenu(hMainMenu, i);
		if (subMenu == NULL) {
			continue;
		}
		if (subMenu == hPopupMenu) {
			popupIndex = i;
		}

		std::wstring title = GetMenuTitleW(hMainMenu, static_cast<UINT>(i), MF_BYPOSITION);
		if (title.find(L"编译") != std::wstring::npos || title.find(L"Build") != std::wstring::npos) {
			if (compileIndex < 0) {
				compileIndex = i;
			}
			continue;
		}
		if (title.find(L"工具") != std::wstring::npos || title.find(L"Tools") != std::wstring::npos) {
			if (toolsIndex < 0) {
				toolsIndex = i;
			}
		}
	}

	if (popupIndex < 0) {
		return popupKeywordMatch();
	}

	int targetIndex = compileIndex >= 0 ? compileIndex : toolsIndex;
	if (targetIndex >= 0) {
		return popupIndex == targetIndex;
	}

	// 顶级标题不可读时，回退到子项关键词判断。
	return popupKeywordMatch();
}

void ClearMenuItemsByPosition(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}

	int count = GetMenuItemCount(hMenu);
	for (int i = count - 1; i >= 0; --i) {
		DeleteMenu(hMenu, static_cast<UINT>(i), MF_BYPOSITION);
	}
}

bool EnsureTopLinkerSubMenu()
{
	if (g_topLinkerSubMenu == NULL || !IsMenu(g_topLinkerSubMenu)) {
		g_topLinkerSubMenu = CreatePopupMenu();
	}
	if (g_topLinkerSubMenu == NULL) {
		return false;
	}

	return true;
}

void EnsureTopLinkerSubMenuAttached(HMENU hTargetMenu)
{
	if (hTargetMenu == NULL) {
		return;
	}
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	bool alreadyAttached = false;
	int attachedIndex = -1;
	for (int i = GetMenuItemCount(hTargetMenu) - 1; i >= 0; --i) {
		MENUITEMINFOW mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_SUBMENU;
		bool removeThis = false;
		if (GetMenuItemInfoW(hTargetMenu, static_cast<UINT>(i), TRUE, &mii) && mii.hSubMenu == g_topLinkerSubMenu) {
			alreadyAttached = true;
			attachedIndex = i;
			continue;
		}
		if (!removeThis) {
			std::wstring title = GetMenuTitleW(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (title.find(L"链接器切换") != std::wstring::npos) {
				removeThis = true;
			}
		}
		if (removeThis) {
			DeleteMenu(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (i > 0) {
				UINT prevState = GetMenuState(hTargetMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				if (prevState != 0xFFFFFFFF && (prevState & MF_SEPARATOR) == MF_SEPARATOR) {
					DeleteMenu(hTargetMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				}
			}
		}
	}

	if (alreadyAttached) {
		if (attachedIndex > 0) {
			UINT prevState = GetMenuState(hTargetMenu, static_cast<UINT>(attachedIndex - 1), MF_BYPOSITION);
			if (prevState == 0xFFFFFFFF || (prevState & MF_SEPARATOR) != MF_SEPARATOR) {
				InsertMenuW(hTargetMenu, static_cast<UINT>(attachedIndex), MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
			}
		}
		return;
	}

	int count = GetMenuItemCount(hTargetMenu);
	if (count > 0) {
		UINT lastState = GetMenuState(hTargetMenu, static_cast<UINT>(count - 1), MF_BYPOSITION);
		if (lastState != 0xFFFFFFFF && (lastState & MF_SEPARATOR) != MF_SEPARATOR) {
			AppendMenuW(hTargetMenu, MF_SEPARATOR, 0, NULL);
		}
	}
	AppendMenuW(hTargetMenu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(g_topLinkerSubMenu), L"链接器切换");
}

void RebuildTopLinkerSubMenu()
{
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	ClearMenuItemsByPosition(g_topLinkerSubMenu);
	g_topLinkerCommandMap.clear();

	if (g_linkerManager.getCount() <= 0) {
		AppendMenuA(g_topLinkerSubMenu, MF_STRING | MF_GRAYED, 0, "（未找到Linker配置）");
		return;
	}

	UpdateCurrentOpenSourceFile();
	std::string nowLinkConfigName = g_configManager.getValue(g_nowOpenSourceFilePath);

	UINT cmd = IDM_AUTOLINKER_LINKER_BASE;
	for (const auto& [key, value] : g_linkerManager.getMap()) {
		if (cmd > IDM_AUTOLINKER_LINKER_MAX) {
			break;
		}

		UINT flags = MF_STRING | MF_ENABLED;
		if (!nowLinkConfigName.empty() && nowLinkConfigName == key) {
			flags |= MF_CHECKED;
		}

		AppendMenuA(g_topLinkerSubMenu, flags, cmd, key.c_str());
		g_topLinkerCommandMap[cmd] = key;
		++cmd;
	}
}

bool HandleTopLinkerMenuCommand(UINT cmd)
{
	auto it = g_topLinkerCommandMap.find(cmd);
	if (it == g_topLinkerCommandMap.end()) {
		return false;
	}

	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		OutputStringToELog("当前没有打开源文件，无法切换Linker");
		return true;
	}

	g_configManager.setValue(g_nowOpenSourceFilePath, it->second);
	OutputCurrentSourceLinker();
	RebuildTopLinkerSubMenu();
	return true;
}

void HandleInitMenuPopup(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}
	if (hMenu == g_topLinkerSubMenu) {
		RebuildTopLinkerSubMenu();
		return;
	}

	if (IsCompileOrToolsTopPopup(hMenu)) {
		EnsureTopLinkerSubMenuAttached(hMenu);
		RebuildTopLinkerSubMenu();
		return;
	}

	UINT state = GetMenuState(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND);
	if (state != 0xFFFFFFFF) {
		EnableMenuItem(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND | MF_ENABLED);
	}

	const UINT aiCmdIds[] = {
		IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC,
		IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC,
		IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC,
		IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT,
		IDM_AUTOLINKER_CTX_AI_COMPLETE_DECL
	};
	for (UINT cmdId : aiCmdIds) {
		UINT aiState = GetMenuState(hMenu, cmdId, MF_BYCOMMAND);
		if (aiState != 0xFFFFFFFF) {
			EnableMenuItem(hMenu, cmdId, MF_BYCOMMAND | MF_ENABLED);
		}
	}
}
}


static auto originalCreateFileA = CreateFileA;
static auto originalGetSaveFileNameA = GetSaveFileNameA;
static auto originalCreateProcessA = CreateProcessA;
static auto originalMessageBoxA = MessageBoxA;

//开始调试 5.71-5.95
typedef int(__thiscall* OriginalEStartDebugFuncType)(DWORD* thisPtr, int a2, int a3);
OriginalEStartDebugFuncType originalEStartDebugFunc = (OriginalEStartDebugFuncType)0x40A080; //int __thiscall sub_40A080(int this, int a2, int a3)

//开始编译 5.71-5.95
typedef int(__thiscall* OriginalEStartCompileFuncType)(DWORD* thisPtr, int a2);
OriginalEStartCompileFuncType originalEStartCompileFunc = (OriginalEStartCompileFuncType)0x40A9F1;  //int __thiscall sub_40A9F1(_DWORD *this, int a2)

int __fastcall MyEStartCompileFunc(DWORD* thisPtr, int dummy, int a2) {
	OutputStringToELog("编译开始#2");
	//ChangeVMProtectModel(true);
	RunChangeECOM(true);
	return originalEStartCompileFunc(thisPtr, a2);
}

int __fastcall MyEStartDebugFunc(DWORD* thisPtr, int dummy, int a2, int a3) {
	OutputStringToELog("调试开始#2");
	//ChangeVMProtectModel(false);
	RunChangeECOM(false);
	return originalEStartDebugFunc(thisPtr, a2, a3);
}

HANDLE WINAPI MyCreateFileA(
	LPCSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDisposition,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile) {
	if (std::string(lpFileName).find("\\Temp\\e_debug\\") != std::string::npos) {
		OutputStringToELog("结束预编译代码（调试）");
		g_preDebugging = false;
	}

	if (g_preDebugging) {
		
	}

	std::filesystem::path currentPath = GetBasePath();
	std::filesystem::path autoLinkerPath = currentPath / "tools" / "link.ini";
	if (autoLinkerPath.string() == std::string(lpFileName)) {
		g_preCompiling = false;

		//EC编译阶段结束
		auto linkName = g_configManager.getValue(g_nowOpenSourceFilePath);

		if (!linkName.empty() ) {
			auto linkConfig = g_linkerManager.getConfig(linkName);

			if (std::filesystem::exists(linkConfig.path)) {
				//切换路径
				OutputStringToELog("切换为Linker:" + linkConfig.name + " " + linkConfig.path);

				return originalCreateFileA(linkConfig.path.c_str(), dwDesiredAccess, dwShareMode, lpSecurityAttributes,
					dwCreationDisposition,
					dwFlagsAndAttributes,
					hTemplateFile);
			}
			else {
				OutputStringToELog("无法切换Linker，Linker文件不存在#1");
			}

		}
		else {
			OutputStringToELog("未设置此源文件的Linker，使用默认");
		}
	}


	return originalCreateFileA(lpFileName,dwDesiredAccess,dwShareMode,lpSecurityAttributes,
								dwCreationDisposition,
								dwFlagsAndAttributes,
								hTemplateFile);
}

BOOL APIENTRY MyGetSaveFileNameA(LPOPENFILENAMEA item) {
	if (g_preCompiling) {
		OutputStringToELog("结束预编译代码");
		g_preCompiling = false;
	}

	return originalGetSaveFileNameA(item);
}


BOOL WINAPI MyCreateProcessA(
	LPCSTR lpApplicationName,
	LPSTR lpCommandLine,
	LPSECURITY_ATTRIBUTES lpProcessAttributes,
	LPSECURITY_ATTRIBUTES lpThreadAttributes,
	BOOL bInheritHandles,
	DWORD dwCreationFlags,
	LPVOID lpEnvironment,
	LPCSTR lpCurrentDirectory,
	LPSTARTUPINFOA lpStartupInfo,
	LPPROCESS_INFORMATION lpProcessInformation
) {
	std::string commandLine = lpCommandLine;

	//检查加入提前链接
	auto krnlnPath = GetLinkerCommandKrnlnFileName(lpCommandLine);
	OutputStringToELog(krnlnPath);

	if (!krnlnPath.empty()) {
		//核心库代码优先使用黑月的，然后再使用核心库的
		//std::string bklib = "D:\\git\\KanAutoControls\\Release\\TestCore.lib";
		
		//将把这个Lib插入的链接器的前方，添加/FORCE强制链接，忽略重定义
		std::string libFilePath = std::format("{}\\AutoLinker\\ForceLinkLib.txt", GetBasePath());
		auto libList = ReadFileAndSplitLines(libFilePath);



		if (libList.size() > 0) {
			//和链接器关联
			auto currentLinkerName = g_configManager.getValue(g_nowOpenSourceFilePath);

			OutputStringToELog(std::format("当前指定的链接器：{}", currentLinkerName));

			std::string libCmd;
			for (const auto& line : libList) {
				auto lines = SplitStringTwo(line, '=');
				auto libPath = line;
				std::string linkerName;
				if (lines.size() == 2) {
					linkerName = lines[0];
					libPath = lines[1];
				}

				//OutputStringToELog(std::format("找到设定的强制链接Lib：{} -> {}", linkerName, libPath));

				if (!linkerName.empty()) {
					//要求必须指定Linker才可使用（包含名称既可）
					if (currentLinkerName.find(linkerName) != std::string::npos) {
						//可使用，link名称一致
						if (std::filesystem::exists(libPath)) {
							OutputStringToELog(std::format("强制链接Lib：{} -> {}", linkerName, libPath));
							if (!libCmd.empty()) {
								libCmd += " ";
							}
							libCmd += "\"" + libPath + "\"";
						}
					} else {
						
						OutputStringToELog(std::format("链接器{}不符合当前的链接器{}，不链接：{}", linkerName, currentLinkerName, libPath));
					}

				} else {
					if (std::filesystem::exists(libPath)) {
						OutputStringToELog(std::format("强制链接Lib：{}", libPath));
						if (!libCmd.empty()) {
							libCmd += " ";
						}
						libCmd += "\"" + libPath + "\"";
					}
				}


			}

			std::string newLibs = libCmd + " \"" + krnlnPath + "\" /FORCE";
			commandLine = ReplaceSubstring(commandLine, "\"" + krnlnPath + "\"", newLibs);


		}
	}

	auto outFileName = GetLinkerCommandOutFileName(lpCommandLine);
	if (!outFileName.empty()) {
		if (commandLine.find("/pdb:\"build.pdb\"") != std::string::npos) {
			//PDB更名为当前编译的程序的名字+.pdb，看起来更正规
			std::string newPdbCommand = std::format("/pdb:\"{}.pdb\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/pdb:\"build.pdb\"", newPdbCommand);
		}
		if (commandLine.find("/map:\"build.map\"") != std::string::npos) {
			//Map更名
			std::string newPdbCommand = std::format("/map:\"{}.map\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/map:\"build.map\"", newPdbCommand);
		}
	}
	OutputStringToELog("启动命令行：" + commandLine);
	return originalCreateProcessA(lpApplicationName, (char *)commandLine.c_str(), lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
}

int WINAPI MyMessageBoxA(
	HWND hWnd,
	LPCSTR lpText,
	LPCSTR lpCaption,
	UINT uType) {
	//自动返回确认编译
	if (std::string(lpCaption).find("链接器输出了大量错误或警告信息") != std::string::npos) {
		return IDNO;
	}
	if (std::string(lpText).find("链接器输出了大量错误或警告信息") != std::string::npos) {
		return IDNO;
	}
	return originalMessageBoxA(hWnd, lpText, lpCaption, uType);
}

void StartHookCreateFileA() {
	DetourRestoreAfterWith();
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)originalCreateFileA, MyCreateFileA);
	DetourAttach(&(PVOID&)originalGetSaveFileNameA, MyGetSaveFileNameA);
	DetourAttach(&(PVOID&)originalCreateProcessA, MyCreateProcessA);
	DetourAttach(&(PVOID&)originalMessageBoxA, MyMessageBoxA);


	if (g_debugStartAddress != -1 && g_compileStartAddress != -1) {
		originalEStartCompileFunc = (OriginalEStartCompileFuncType)g_compileStartAddress;
		originalEStartDebugFunc = (OriginalEStartDebugFuncType)g_debugStartAddress;

		//用于自动在编译和调试之间切换Lib与Dll模块
		DetourAttach(&(PVOID&)originalEStartCompileFunc, MyEStartCompileFunc);
		DetourAttach(&(PVOID&)originalEStartDebugFunc, MyEStartDebugFunc);
	}
	else {
		//无法启用
	}
	DetourTransactionCommit();
}



/// <summary>
/// 输出文本
/// </summary>
/// <param name="szbuf"></param>
void OutputStringToELog(const std::string& szbuf) {
	const std::string line = "[AutoLinker]" + szbuf;
	OutputDebugStringA((line + "\n").c_str());
	IDEFacade::Instance().AppendOutputWindowLine(line);
}


void UpdateCurrentOpenSourceFile() {
	std::string sourceFile = GetSourceFilePath();

	if (sourceFile.empty()) {
		//OutputStringToELog("无法获取源文件路径");
	}

	if (g_nowOpenSourceFilePath != sourceFile) {
		OutputStringToELog(sourceFile);
	}
	g_nowOpenSourceFilePath = sourceFile;
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

//工具条子类过程
LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	switch (uMsg)
	{
		case WM_INITMENUPOPUP:
			HandleInitMenuPopup(reinterpret_cast<HMENU>(wParam));
			break;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}


/// <summary>
/// 废弃，使用新的配置来切换了
/// </summary>
/// <param name="isLib"></param>
void ChangeVMProtectModel(bool isLib) {
	if (isLib) {
		
		int sdk = FindECOMNameIndex("VMPSDK");
		if (sdk != -1) {
			RemoveECOM(sdk); //移除
		}
		int sdkLib = FindECOMNameIndex("VMPSDK_LIB");
		if (sdkLib == -1) {
			OutputStringToELog("切换到静态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK_LIB.ec", buffer);

			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	else {
		int sdk = FindECOMNameIndex("VMPSDK_LIB");
		if (sdk != -1) {
			RemoveECOM(sdk); //移除
		}
		int sdkLib = FindECOMNameIndex("VMPSDK");
		if (sdkLib == -1) {
			OutputStringToELog("切换到动态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK.ec", buffer);
			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	//
}

#define WM_AUTOLINKER_INIT (WM_USER + 1000)

bool FneInit();

//主窗口子类过程
LRESULT CALLBACK MainWindowSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	//if (uMsg == 20707) {
	//	//此消息配合AutoLinkerBuild，已废弃！
	//	if (wParam) {
	//		g_preCompiling = true;
	//		OutputStringToELog("开始编译");
	//		//ChangeVMPModel(true);
	//	}
	//	else {
	//		g_preDebugging = true;
	//		OutputStringToELog("开始调试");
	//		//ChangeVMPModel(false);
	//	}
	//	return 0;
	//}

	if (uMsg == WM_AUTOLINKER_AI_TASK_DONE) {
		HandleAiTaskCompletionMessage(lParam);
		return 0;
	}
	if (uMsg == WM_AUTOLINKER_AI_APPLY_RESULT) {
		HandleAiApplyMessage(lParam);
		return 0;
	}

	if (uMsg == WM_COMMAND) {
		UINT cmd = LOWORD(wParam);
		if (HandleTopLinkerMenuCommand(cmd)) {
			return 0;
		}
		if (IDEFacade::Instance().HandleMainWindowCommand(wParam)) {
			return 0;
		}
	}

	if (uMsg == WM_AUTOLINKER_INIT) {
		OutputStringToELog("收到初始化消息，尝试初始化");
		if (FneInit()) {
			OutputStringToELog("初始化成功");
			return 1;
		}
		return 0;
	}

	if (uMsg == 20708) {
		BOOL result = SetWindowSubclass((HWND)wParam, EditViewSubclassProc, 0, 0);
		return result ? 1 : 0;
	}

	if (uMsg == WM_INITMENUPOPUP) {
		LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		HandleInitMenuPopup(reinterpret_cast<HMENU>(wParam));
		return result;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2) {
	return IDEFacade::Instance().RunFunctionRaw(code, p1, p2);
}

INT WINAPI fnAddInFunc(INT nAddInFnIndex) {

	switch (nAddInFnIndex) {
		case 0: { //TODO 打开项目目录
			std::string cmd = std::format("/select,{}", g_nowOpenSourceFilePath);
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 1: { //Auto 目录
			std::string cmd = std::format("{}\\AutoLinker", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 2: { //E 目录
			std::string cmd = std::format("{}", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 3: { // 复制当前函数代码
			TryCopyCurrentFunctionCode();
			break;
		}
		case 4: { // AutoLinker AI接口设置
			AISettings settings = {};
			AIService::LoadSettings(g_configManager, settings);
			if (!ShowAIConfigDialog(g_hwnd, settings)) {
				OutputStringToELog("AI配置已取消");
				break;
			}
			AIService::SaveSettings(g_configManager, settings);
			OutputStringToELog("AI配置已保存");
			break;
		}
		//case 4: { //切换到VMPSDK静态（自用）
		//	ChangeVMProtectModel(true);
		//	break;
		//}
		//case 5: { //切换到VMPSDK动态（自用）
		//	ChangeVMProtectModel(false);
		//	break;
		//}
			  
		default: {

		}
	}

	return 0;
}


void FneCheckNewVersion(void* pParams) {
	Sleep(1000);

	OutputStringToELog("AutoLinker开源下载地址：https://github.com/aiqinxuancai/AutoLinker");
	std::string url = "https://api.github.com/repos/aiqinxuancai/AutoLinker/releases";
	//std::string customHeaders = "user-agent: Mozilla/5.0";
	auto response = PerformGetRequest(url);

	std::string currentVersion = AUTOLINKER_VERSION;


	if (response.second == 200) {
		std::string nowGithubVersion = "0.0.0";
		

		if (strcmp(AUTOLINKER_VERSION, "0.0.0") == 0) {
			//自行编译，无需检查版本更新
			OutputStringToELog(std::format("自编译版本，不检查更新，当前版本：{}", currentVersion));
		} else {
			if (!response.first.empty()) {

				try
				{
					auto releases = json::parse(response.first);
					for (const auto& release : releases) {
						if (!release["prerelease"].get<bool>()) {
							nowGithubVersion = release["tag_name"];
							break;
						}
					}

					Version nowGithubVersionObj(nowGithubVersion);
					Version currentVersionObj(AUTOLINKER_VERSION);

					if (nowGithubVersionObj > currentVersionObj) {
						OutputStringToELog(std::format("有新版本：{}", nowGithubVersion));
					}
					else {

					}
				}
				catch (const std::exception& e) {
					OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误：{}", currentVersion, e.what()));
				}


			}
		}

		
	}
	else {
		//OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误码：{}", currentVersion, response.second));
	}

	//return false;
}

bool FneInit() {
	OutputStringToELog("开始初始化");

	// g_hwnd 已经在外部获取并子类化
	if (g_hwnd == NULL) {
		OutputStringToELog("g_hwnd 为空");
		return false;
	}

	g_toolBarHwnd = FindMenuBar(g_hwnd);

	DWORD processID = GetCurrentProcessId();
	std::string s = std::format("E进程ID{} 主句柄{} 菜单栏句柄{}", processID, (int)g_hwnd, (int)g_toolBarHwnd);
	OutputStringToELog(s);

	if (g_toolBarHwnd != NULL)
	{
		StartEditViewSubclassTask();
		RebuildTopLinkerSubMenu();
		OutputCurrentSourceLinker();

		OutputStringToELog("找到工具条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		StartHookCreateFileA();
		PostAppMessageA(g_toolBarHwnd, WM_PRINT, 0, 0);
		OutputStringToELog("初始化完成");

		//初始化Lib相关库的状态

		//启动版本检查线程
		uintptr_t threadID = _beginthread(FneCheckNewVersion, 0, NULL);

		return true;
	}
	else
	{
		OutputStringToELog(std::format("初始化失败，未找到工具条窗口 {}", (int)g_toolBarHwnd));
	}

	return false;
}

/*-----------------支持库消息处理函数------------------*/

// 初始化重试线程函数
void InitRetryThread(void* pParams) {
	const int MAX_RETRY_COUNT = 5;
	int retryCount = 0;

	while (retryCount < MAX_RETRY_COUNT) {
		Sleep(1000); // 每秒重试一次
		retryCount++;

		OutputStringToELog(std::format("初始化重试第 {}/{} 次", retryCount, MAX_RETRY_COUNT));

		// 发送自定义消息到主窗口进行初始化
		LRESULT result = SendMessage(g_hwnd, WM_AUTOLINKER_INIT, 0, 0);

		if (result == 1) {
			// 初始化成功
			OutputStringToELog("初始化线程完成"); 
			return;
		}
	}

	OutputStringToELog(std::format("初始化失败，已重试 {} 次", MAX_RETRY_COUNT));
}

EXTERN_C INT WINAPI AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	std::string s = std::format("AutoLinker_MessageNotify {0} {1} {2}", (int)nMsg, dwParam1, dwParam2);
	OutputStringToELog(s);

#ifndef __E_STATIC_LIB
	if (nMsg == NL_GET_CMD_FUNC_NAMES) // 返回所有命令实现函数的的函数名称数组(char*[]), 支持静态编译的动态库必须处理
		return NULL;
	else if (nMsg == NL_GET_NOTIFY_LIB_FUNC_NAME) // 返回处理系统通知的函数名称(PFN_NOTIFY_LIB函数名称), 支持静态编译的动态库必须处理
		return (INT)LIBARAYNAME;
	else if (nMsg == NL_GET_DEPENDENT_LIBS) return (INT)NULL;
	//else if (nMsg == NL_GET_DEPENDENT_LIBS) return (INT)NULL;
	// 返回静态库所依赖的其它静态库文件名列表(格式为\0分隔的文本,结尾两个\0), 支持静态编译的动态库必须处理
	// kernel32.lib user32.lib gdi32.lib 等常用的系统库不需要放在此列表中
	// 返回NULL或NR_ERR表示不指定依赖文件  

	else if (nMsg == NL_SYS_NOTIFY_FUNCTION) {
		if (dwParam1) {
			if (!g_initStarted) {
				// 获取主窗口句柄
				g_hwnd = GetMainWindowByProcessId();

				if (g_hwnd) {
					g_initStarted = true;
					SetWindowSubclass(g_hwnd, MainWindowSubclassProc, 0, 0);
					RegisterIDEContextMenu();
					OutputStringToELog("主窗口子类化完成，启动初始化线程");
					uintptr_t threadID = _beginthread(InitRetryThread, 0, NULL);
				} else {
					OutputStringToELog("无法获取主窗口句柄");
				}
			}

		}
	}
	else if (nMsg == NL_RIGHT_POPUP_MENU_SHOW) {
		IDEFacade::Instance().HandleNotifyMessage(nMsg, dwParam1, dwParam2);
	}

#endif
	return ProcessNotifyLib(nMsg, dwParam1, dwParam2);
}
/*定义支持库基本信息*/
#ifndef __E_STATIC_LIB
static LIB_INFOX LibInfo =
{
	/* { 库格式号, GUID串号, 主版本号, 次版本号, 构建版本号, 系统主版本号, 系统次版本号, 核心库主版本号, 核心库次版本号,
	支持库名, 支持库语言, 支持库描述, 支持库状态,
	作者姓名, 邮政编码, 通信地址, 电话号码, 传真号码, 电子邮箱, 主页地址, 其它信息,
	类型数量, 类型指针, 类别数量, 命令类别, 命令总数, 命令指针, 命令入口,
	附加功能, 功能描述, 消息指针, 超级模板, 模板描述,
	常量数量, 常量指针, 外部文件} */
	LIB_FORMAT_VER,
	_T(LIB_GUID_STR),
	LIB_MajorVersion,
	LIB_MinorVersion,
	LIB_BuildNumber,
	LIB_SysMajorVer,
	LIB_SysMinorVer,
	LIB_KrnlLibMajorVer,
	LIB_KrnlLibMinorVer,
	_T(LIB_NAME_STR),
	__GBK_LANG_VER,
	_WT(LIB_DESCRIPTION_STR),
	LBS_IDE_PLUGIN | LBS_LIB_INFO2, //_LIB_OS(__OS_WIN), //#LBS_IDE_PLUGIN  LBS_LIB_INFO2
	_WT(LIB_Author),
	_WT(LIB_ZipCode),
	_WT(LIB_Address),
	_WT(LIB_Phone), 
	_WT(LIB_Fax),
	_WT(LIB_Email),
	_WT(LIB_HomePage),
	_WT(LIB_Other),
	0,
	NULL,
	LIB_TYPE_COUNT,
	_WT(LIB_TYPE_STR),
	0,
	NULL,
	NULL,
	fnAddInFunc,
	_T("打开项目目录\0这是个用作测试的辅助工具功能。\0打开AutoLinker配置目录\0这是个用作测试的辅助工具功能。\0打开E语言目录\0这是个用作测试的辅助工具功能。\0复制当前函数代码\0复制当前光标所在子程序完整代码到剪贴板。\0AutoLinker AI接口设置\0编辑AI接口地址、API Key、模型和提示词等配置。\0\0") ,
	AutoLinker_MessageNotify,
	NULL,
	NULL,
	0,
	NULL,
	NULL,
	//-----------------
	NULL,
	0,
	NULL,
	"You",

};

PLIB_INFOX WINAPI GetNewInf()
{
	//LibInfo.m_szLicenseToUserName = "You"
	return (&LibInfo);
};
#endif
