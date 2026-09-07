#pragma once

#include <cstddef>
#include <string>
#include <unordered_map>
#include <vector>

#include "AIService.h"

// 管理 AI 长任务的 token 预算、停滞检测和可恢复检查点。
class AIChatRunController {
public:
	AIChatRunController(
		const AISettings& settings,
		const std::vector<AIChatMessage>& initialContext,
		const AIChatRunOptions& options);

	const std::vector<AIChatMessage>& ContextMessages() const;
	void AppendContextMessage(AIChatMessage message);
	void ReplaceContextWithSummary(const std::string& summaryLocal);
	// 按 Codex 压缩规则保留最近 user 消息，并追加摘要 user 片段。
	void ReplaceContextWithCompaction(const std::string& summaryLocal);

	void BeginSampling();
	// 向当前 AI 对话报告阶段状态，不改变检查点状态。
	void ReportActivity(const std::string& line) const;
	void RecordUsage(
		int promptTokens,
		int totalTokens,
		bool hasUsage,
		int cachedInputTokens = 0,
		int cacheWriteInputTokens = 0);
	void RecordModelRound();
	bool ShouldCompact() const;
	// 返回自动压缩判定所使用的上下文字节数和预测 token，供诊断日志使用。
	size_t ContextBytes() const;
	size_t PredictedContextTokens() const;
	int ContextWindowTokens() const;

	void BeginToolBatch(std::vector<AIChatCheckpointToolCall> calls);
	void CompleteToolCall(
		size_t index,
		const std::string& resultJsonLocal,
		bool ok);
	void CompleteToolCall(
		size_t index,
		const std::string& resultJsonLocal,
		bool ok,
		AIChatMessage contextMessage);
	std::string TakeRecoveryHint();
	bool IsStalled() const;
	std::string StallReason() const;
	void ResetForNewContextWindow();

	void RecordCompaction(const std::string& summaryLocal);
	void PublishCheckpoint(const std::string& state = "running");
	AIChatRunCheckpoint BuildCheckpoint(const std::string& state = "running") const;
	std::string BuildLocalFallbackSummary(const std::vector<AIChatToolEvent>& events) const;

	int SamplingRounds() const;
	int CompactionCount() const;
	int PromptTokens() const;
	int TotalTokens() const;
	int CachedInputTokens() const;
	int CacheWriteInputTokens() const;
	long long AccumulatedInputTokens() const;
	long long AccumulatedOutputTokens() const;
	int CompletedModelRounds() const;
	bool HasUsage() const;

	static size_t EstimateContextTokens(const std::vector<AIChatMessage>& messages);

private:
	struct RepeatedWriteFailureState {
		std::string errorSignature;
		int count = 0;
	};

	void ApplyResumeCheckpoint(const AIChatRunCheckpoint& checkpoint);
	std::string BuildResumeFallbackSummary(const AIChatRunCheckpoint& checkpoint) const;
	void CompleteToolCallInternal(
		size_t index,
		const std::string& resultJsonLocal,
		bool ok,
		AIChatMessage* contextMessage);

	AIProtocolType m_protocolType;
	std::string m_model;
	int m_contextWindowTokens;
	AIChatRunOptions m_options;
	std::vector<AIChatMessage> m_contextMessages;
	std::vector<AIChatCheckpointToolCall> m_toolCalls;
	std::string m_summary;
	int m_samplingRounds = 0;
	int m_compactionCount = 0;
	int m_promptTokens = 0;
	int m_totalTokens = 0;
	int m_cachedInputTokens = 0;
	int m_cacheWriteInputTokens = 0;
	long long m_accumulatedInputTokens = 0;
	long long m_accumulatedOutputTokens = 0;
	int m_completedModelRounds = 0;
	bool m_hasUsage = false;
	size_t m_contextBytesAfterUsage = 0;
	int m_consecutiveFailures = 0;
	std::string m_recoveryHint;
	std::string m_lastFailedToolName;
	std::string m_lastFailureDetail;
	std::unordered_map<std::string, RepeatedWriteFailureState> m_writeFailures;
	bool m_repeatedWriteFailureStalled = false;
	std::string m_repeatedWriteFailureTarget;
};
