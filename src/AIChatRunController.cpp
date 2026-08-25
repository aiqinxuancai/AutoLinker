#include "AIChatRunController.h"

#include <algorithm>
#include <sstream>
#include <utility>

#include "..\\thirdparty\\json.hpp"

namespace {

constexpr int kAutoCompactPercent = 90;
constexpr int kAutoCompactSamplingRounds = 256;
constexpr size_t kAutoCompactContextBytes = 900000;
constexpr int kRecoveryHintFailureCount = 3;
constexpr int kStalledFailureCount = 8;
constexpr int kRepeatedWriteRecoveryCount = 3;
constexpr int kRepeatedWriteStalledCount = 5;
constexpr size_t kMaxFallbackEventCount = 32;

size_t MessageBytes(const AIChatMessage& message)
{
	return message.role.size() +
		message.content.size() +
		message.reasoningContent.size() +
		message.rawMessageJsonUtf8.size();
}

std::string TruncateText(const std::string& text, size_t limit)
{
	if (text.size() <= limit) {
		return text;
	}
	return text.substr(0, limit) + "...";
}

std::string ToLowerAsciiCopy(std::string text)
{
	for (char& ch : text) {
		if (ch >= 'A' && ch <= 'Z') {
			ch = static_cast<char>(ch - 'A' + 'a');
		}
	}
	return text;
}

bool IsSourceWriteTool(const std::string& toolName)
{
	const std::string normalized = ToLowerAsciiCopy(toolName);
	return normalized == "edit_file" ||
		normalized == "multi_edit_file" ||
		normalized == "write_file" ||
		normalized == "add_new_file" ||
		normalized == "restore_file_snapshot";
}

std::string GetJsonStringField(const std::string& jsonText, const char* key)
{
	if (jsonText.empty() || key == nullptr) {
		return std::string();
	}
	try {
		const nlohmann::json value = nlohmann::json::parse(jsonText);
		if (value.is_object() && value.contains(key) && value[key].is_string()) {
			return value[key].get<std::string>();
		}
	}
	catch (...) {
	}
	return std::string();
}

std::string BuildWriteTargetKey(
	const std::string& argumentsJson,
	std::string& outDisplayTarget)
{
	outDisplayTarget = GetJsonStringField(argumentsJson, "file_path");
	if (outDisplayTarget.empty()) {
		outDisplayTarget = GetJsonStringField(argumentsJson, "name");
	}
	if (outDisplayTarget.empty()) {
		outDisplayTarget = "<unknown>";
	}
	return ToLowerAsciiCopy(outDisplayTarget);
}

std::string BuildFailureDetail(const std::string& resultJson)
{
	std::string detail = GetJsonStringField(resultJson, "error");
	if (detail.empty()) {
		detail = GetJsonStringField(resultJson, "reason");
	}
	if (detail.empty()) {
		detail = TruncateText(resultJson, 400);
	}
	return detail.empty() ? "<unknown>" : detail;
}

std::string BuildFailureSignature(const std::string& resultJson)
{
	return ToLowerAsciiCopy(BuildFailureDetail(resultJson));
}

} // namespace

AIChatRunController::AIChatRunController(
	const AISettings& settings,
	const std::vector<AIChatMessage>& initialContext,
	const AIChatRunOptions& options)
	: m_protocolType(settings.protocolType),
	  m_model(settings.model),
	  m_contextWindowTokens(AIService::ResolveContextWindowTokens(settings)),
	  m_options(options),
	  m_contextMessages(initialContext)
{
	if (options.resumeCheckpoint != nullptr) {
		ApplyResumeCheckpoint(*options.resumeCheckpoint);
	}
}

const std::vector<AIChatMessage>& AIChatRunController::ContextMessages() const
{
	return m_contextMessages;
}

void AIChatRunController::AppendContextMessage(AIChatMessage message)
{
	m_contextBytesAfterUsage += MessageBytes(message);
	m_contextMessages.push_back(std::move(message));
}

void AIChatRunController::ReplaceContextWithSummary(const std::string& summaryLocal)
{
	m_summary = summaryLocal;
	m_contextMessages.clear();
	m_contextMessages.push_back(AIChatMessage{
		"system",
		"长期任务压缩检查点：\n" + summaryLocal,
		"",
		""
	});
	m_contextMessages.push_back(AIChatMessage{
		"user",
		"请从上述检查点继续执行原任务。先确认未完成步骤，再继续调用必要工具。",
		"",
		""
	});
	m_contextBytesAfterUsage = 0;
	m_promptTokens = 0;
	m_totalTokens = 0;
	m_hasUsage = false;
	ResetForNewContextWindow();
}

void AIChatRunController::BeginSampling()
{
	++m_samplingRounds;
}

void AIChatRunController::ReportActivity(const std::string& line) const
{
	if (m_options.activityCallback && !line.empty()) {
		m_options.activityCallback(line);
	}
}

void AIChatRunController::RecordUsage(int promptTokens, int totalTokens, bool hasUsage)
{
	if (!hasUsage) {
		return;
	}
	m_hasUsage = true;
	m_promptTokens = (std::max)(0, promptTokens);
	m_totalTokens = (std::max)(m_promptTokens, totalTokens);
	m_accumulatedInputTokens += static_cast<long long>(m_promptTokens);
	m_accumulatedOutputTokens += static_cast<long long>(m_totalTokens - m_promptTokens);
	m_contextBytesAfterUsage = 0;
}

void AIChatRunController::RecordModelRound()
{
	++m_completedModelRounds;
}

bool AIChatRunController::ShouldCompact() const
{
	if (m_contextWindowTokens <= 0) {
		return false;
	}
	if (m_samplingRounds - m_samplingRoundsAtLastCompaction >= kAutoCompactSamplingRounds) {
		return true;
	}
	size_t contextBytes = 0;
	for (const AIChatMessage& message : m_contextMessages) {
		contextBytes += MessageBytes(message);
	}
	if (contextBytes >= kAutoCompactContextBytes) {
		return true;
	}
	const size_t limit = static_cast<size_t>(m_contextWindowTokens) * kAutoCompactPercent / 100;
	const size_t predictedTokens = m_hasUsage
		? static_cast<size_t>(m_promptTokens) + (m_contextBytesAfterUsage + 3) / 4
		: EstimateContextTokens(m_contextMessages);
	return predictedTokens >= limit;
}

void AIChatRunController::BeginToolBatch(std::vector<AIChatCheckpointToolCall> calls)
{
	m_toolCalls = std::move(calls);
	PublishCheckpoint();
}

void AIChatRunController::CompleteToolCall(
	size_t index,
	const std::string& resultJsonLocal,
	bool ok)
{
	CompleteToolCallInternal(index, resultJsonLocal, ok, nullptr);
}

void AIChatRunController::CompleteToolCall(
	size_t index,
	const std::string& resultJsonLocal,
	bool ok,
	AIChatMessage contextMessage)
{
	CompleteToolCallInternal(index, resultJsonLocal, ok, &contextMessage);
}

void AIChatRunController::CompleteToolCallInternal(
	size_t index,
	const std::string& resultJsonLocal,
	bool ok,
	AIChatMessage* contextMessage)
{
	std::string toolName;
	std::string argumentsJson;
	if (index < m_toolCalls.size()) {
		auto& call = m_toolCalls[index];
		toolName = call.name;
		argumentsJson = call.argumentsJson;
		call.resultJson = resultJsonLocal;
		call.completed = true;
		call.ok = ok;
	}
	if (contextMessage != nullptr) {
		AppendContextMessage(std::move(*contextMessage));
	}
	if (ok) {
		m_consecutiveFailures = 0;
		m_recoveryHint.clear();
		m_lastFailedToolName.clear();
		m_lastFailureDetail.clear();
		if (IsSourceWriteTool(toolName)) {
			std::string displayTarget;
			m_writeFailures.erase(BuildWriteTargetKey(argumentsJson, displayTarget));
			if (m_repeatedWriteFailureTarget == displayTarget) {
				m_repeatedWriteFailureStalled = false;
				m_repeatedWriteFailureTarget.clear();
			}
		}
	}
	else {
		++m_consecutiveFailures;
		m_lastFailedToolName = toolName.empty() ? "<unknown>" : toolName;
		m_lastFailureDetail = BuildFailureDetail(resultJsonLocal);
		if (m_consecutiveFailures == kRecoveryHintFailureCount) {
			m_recoveryHint =
				"连续工具调用失败 " + std::to_string(m_consecutiveFailures) +
				" 次。最近失败工具：" + m_lastFailedToolName +
				"；最近错误：" + TruncateText(m_lastFailureDetail, 300) +
				"。请先按错误结果修正根因，禁止继续猜测字段或无变化重试；无法修正时改用其他验证或实现路径。";
			if (ToLowerAsciiCopy(m_lastFailedToolName) == "request_user_input") {
				m_recoveryHint +=
					"该工具的失败结果包含 expected_arguments，必须按其结构重新生成：根节点只允许 questions，"
					"header 位于每个问题内，options 是包含 label 和 description 的对象数组。";
			}
		}
		if (IsSourceWriteTool(toolName)) {
			std::string displayTarget;
			const std::string targetKey = BuildWriteTargetKey(argumentsJson, displayTarget);
			const std::string errorSignature = BuildFailureSignature(resultJsonLocal);
			auto& state = m_writeFailures[targetKey];
			if (state.errorSignature == errorSignature) {
				++state.count;
			}
			else {
				state.errorSignature = errorSignature;
				state.count = 1;
			}
			if (state.count == kRepeatedWriteRecoveryCount) {
				m_recoveryHint =
					std::string("对源码目标 ") + displayTarget +
					" 的写入已重复 " + std::to_string(state.count) +
					" 次出现相同错误：" + TruncateText(errorSignature, 300) +
					"。成功读取或预览不代表该写入问题已解决；请先修正根因并采用不同的写入内容或方案。";
			}
			if (state.count >= kRepeatedWriteStalledCount) {
				m_repeatedWriteFailureStalled = true;
				m_repeatedWriteFailureTarget = displayTarget;
			}
		}
	}
	PublishCheckpoint();
}

std::string AIChatRunController::TakeRecoveryHint()
{
	return std::exchange(m_recoveryHint, std::string());
}

bool AIChatRunController::IsStalled() const
{
	return m_consecutiveFailures >= kStalledFailureCount || m_repeatedWriteFailureStalled;
}

std::string AIChatRunController::StallReason() const
{
	if (m_repeatedWriteFailureStalled) {
		return "source write stalled after " + std::to_string(kRepeatedWriteStalledCount) +
			" repeated failures for " +
			(m_repeatedWriteFailureTarget.empty() ? "<unknown>" : m_repeatedWriteFailureTarget);
	}
	return "tool execution stalled after " + std::to_string(kStalledFailureCount) +
		" consecutive failures; last tool: " +
		(m_lastFailedToolName.empty() ? "<unknown>" : m_lastFailedToolName) +
		"; last error: " + TruncateText(
			m_lastFailureDetail.empty() ? "<unknown>" : m_lastFailureDetail,
			400);
}

void AIChatRunController::ResetForNewContextWindow()
{
	m_toolCalls.clear();
	m_consecutiveFailures = 0;
	m_recoveryHint.clear();
	m_lastFailedToolName.clear();
	m_lastFailureDetail.clear();
}

void AIChatRunController::RecordCompaction(const std::string& summaryLocal)
{
	++m_compactionCount;
	ReplaceContextWithSummary(summaryLocal);
	m_samplingRoundsAtLastCompaction = m_samplingRounds;
	PublishCheckpoint();
}

void AIChatRunController::PublishCheckpoint(const std::string& state)
{
	if (m_options.checkpointCallback) {
		m_options.checkpointCallback(BuildCheckpoint(state));
	}
}

AIChatRunCheckpoint AIChatRunController::BuildCheckpoint(const std::string& state) const
{
	AIChatRunCheckpoint checkpoint;
	checkpoint.protocolType = m_protocolType;
	checkpoint.model = m_model;
	checkpoint.state = state;
	checkpoint.summary = m_summary;
	checkpoint.samplingRounds = m_samplingRounds;
	checkpoint.compactionCount = m_compactionCount;
	checkpoint.promptTokens = m_promptTokens;
	checkpoint.totalTokens = m_totalTokens;
	checkpoint.accumulatedInputTokens = m_accumulatedInputTokens;
	checkpoint.accumulatedOutputTokens = m_accumulatedOutputTokens;
	checkpoint.completedModelRounds = m_completedModelRounds;
	checkpoint.hasUsage = m_hasUsage;
	checkpoint.contextMessages = m_contextMessages;
	checkpoint.toolCalls = m_toolCalls;
	return checkpoint;
}

std::string AIChatRunController::BuildLocalFallbackSummary(
	const std::vector<AIChatToolEvent>& events) const
{
	std::ostringstream out;
	out << "目标与历史上下文：\n";
	for (const AIChatMessage& message : m_contextMessages) {
		if (message.role == "user" || message.role == "system") {
			out << "[" << message.role << "] " << TruncateText(message.content, 800) << "\n";
		}
	}
	out << "\n最近工具执行：\n";
	const size_t begin = events.size() > kMaxFallbackEventCount
		? events.size() - kMaxFallbackEventCount
		: 0;
	for (size_t i = begin; i < events.size(); ++i) {
		const AIChatToolEvent& event = events[i];
		out << "- " << event.name << (event.ok ? " (ok)" : " (failed)")
			<< " args=" << TruncateText(event.argumentsJson, 300)
			<< " result=" << TruncateText(event.resultJson, 600) << "\n";
	}
	out << "\n继续要求：复核当前工程状态，完成剩余修改并执行必要测试。";
	return out.str();
}

int AIChatRunController::SamplingRounds() const
{
	return m_samplingRounds;
}

int AIChatRunController::CompactionCount() const
{
	return m_compactionCount;
}

int AIChatRunController::PromptTokens() const
{
	return m_promptTokens;
}

int AIChatRunController::TotalTokens() const
{
	return m_totalTokens;
}

long long AIChatRunController::AccumulatedInputTokens() const
{
	return m_accumulatedInputTokens;
}

long long AIChatRunController::AccumulatedOutputTokens() const
{
	return m_accumulatedOutputTokens;
}

int AIChatRunController::CompletedModelRounds() const
{
	return m_completedModelRounds;
}

bool AIChatRunController::HasUsage() const
{
	return m_hasUsage;
}

size_t AIChatRunController::EstimateContextTokens(const std::vector<AIChatMessage>& messages)
{
	size_t bytes = 0;
	for (const AIChatMessage& message : messages) {
		bytes += MessageBytes(message);
	}
	return (bytes + 3) / 4;
}

void AIChatRunController::ApplyResumeCheckpoint(const AIChatRunCheckpoint& checkpoint)
{
	m_samplingRounds = (std::max)(0, checkpoint.samplingRounds);
	m_compactionCount = (std::max)(0, checkpoint.compactionCount);
	m_promptTokens = (std::max)(0, checkpoint.promptTokens);
	m_totalTokens = (std::max)(m_promptTokens, checkpoint.totalTokens);
	m_accumulatedInputTokens = (std::max)(0LL, checkpoint.accumulatedInputTokens);
	m_accumulatedOutputTokens = (std::max)(0LL, checkpoint.accumulatedOutputTokens);
	m_completedModelRounds = (std::max)(0, checkpoint.completedModelRounds);
	m_hasUsage = checkpoint.hasUsage;
	m_summary = checkpoint.summary;

	const bool hasIncompleteToolCall = std::any_of(
		checkpoint.toolCalls.begin(),
		checkpoint.toolCalls.end(),
		[](const AIChatCheckpointToolCall& call) {
			return !call.completed;
		});
	const bool exactResume = checkpoint.protocolType == m_protocolType &&
		checkpoint.model == m_model &&
		!checkpoint.contextMessages.empty() &&
		!hasIncompleteToolCall;
	if (exactResume) {
		m_contextMessages = checkpoint.contextMessages;
		m_toolCalls = checkpoint.toolCalls;
		if (ShouldCompact()) {
			++m_compactionCount;
			ReplaceContextWithSummary(BuildResumeFallbackSummary(checkpoint));
		}
		return;
	}

	ReplaceContextWithSummary(
		checkpoint.summary.empty()
			? BuildResumeFallbackSummary(checkpoint)
			: checkpoint.summary);
}

std::string AIChatRunController::BuildResumeFallbackSummary(
	const AIChatRunCheckpoint& checkpoint) const
{
	std::ostringstream out;
	out << "恢复来源：协议=" << static_cast<int>(checkpoint.protocolType)
		<< "，模型=" << checkpoint.model << "。\n";
	out << "任务上下文：\n";
	for (const AIChatMessage& message : checkpoint.contextMessages) {
		if (message.content.empty()) {
			continue;
		}
		out << "[" << message.role << "] " << TruncateText(message.content, 1200) << "\n";
	}
	if (!checkpoint.toolCalls.empty()) {
		out << "\n最近工具调用：\n";
		for (const AIChatCheckpointToolCall& call : checkpoint.toolCalls) {
			out << "- " << (call.name.empty() ? "<unknown>" : call.name)
				<< (call.completed ? (call.ok ? " (ok)" : " (failed)") : " (interrupted)")
				<< " args=" << TruncateText(call.argumentsJson, 400);
			if (call.completed && !call.resultJson.empty()) {
				out << " result=" << TruncateText(call.resultJson, 800);
			}
			out << "\n";
		}
	}
	out << "\n恢复要求：工具调用可能已产生外部副作用。先检查工程当前状态，"
		"不要盲目重放已中断调用，再完成剩余步骤和验证。";
	return out.str();
}
