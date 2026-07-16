#include "AIChatRunController.h"

#include <algorithm>
#include <format>
#include <sstream>
#include <utility>

namespace {

constexpr int kAutoCompactPercent = 90;
constexpr int kRecoveryHintFailureCount = 3;
constexpr int kStalledFailureCount = 8;
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

void AIChatRunController::RecordUsage(int promptTokens, int totalTokens, bool hasUsage)
{
	if (!hasUsage) {
		return;
	}
	m_hasUsage = true;
	m_promptTokens = (std::max)(0, promptTokens);
	m_totalTokens = (std::max)(m_promptTokens, totalTokens);
	m_contextBytesAfterUsage = 0;
}

bool AIChatRunController::ShouldCompact() const
{
	if (m_contextWindowTokens <= 0) {
		return false;
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
	bool ok,
	AIChatMessage contextMessage)
{
	if (index < m_toolCalls.size()) {
		auto& call = m_toolCalls[index];
		call.resultJson = resultJsonLocal;
		call.completed = true;
		call.ok = ok;
	}
	AppendContextMessage(std::move(contextMessage));
	if (ok) {
		m_consecutiveFailures = 0;
		m_recoveryHintPending = false;
	}
	else {
		++m_consecutiveFailures;
		if (m_consecutiveFailures == kRecoveryHintFailureCount) {
			m_recoveryHintPending = true;
		}
	}
	PublishCheckpoint();
}

bool AIChatRunController::ShouldInjectRecoveryHint()
{
	return std::exchange(m_recoveryHintPending, false);
}

bool AIChatRunController::IsStalled() const
{
	return m_consecutiveFailures >= kStalledFailureCount;
}

void AIChatRunController::ResetForNewContextWindow()
{
	m_toolCalls.clear();
	m_consecutiveFailures = 0;
	m_recoveryHintPending = false;
}

void AIChatRunController::RecordCompaction(const std::string& summaryLocal)
{
	++m_compactionCount;
	ReplaceContextWithSummary(summaryLocal);
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
