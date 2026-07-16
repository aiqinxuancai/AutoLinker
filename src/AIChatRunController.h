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

	void BeginSampling();
	void RecordUsage(int promptTokens, int totalTokens, bool hasUsage);
	bool ShouldCompact() const;

	void BeginToolBatch(std::vector<AIChatCheckpointToolCall> calls);
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
	bool HasUsage() const;

	static size_t EstimateContextTokens(const std::vector<AIChatMessage>& messages);

private:
	struct RepeatedWriteFailureState {
		std::string errorSignature;
		int count = 0;
	};

	void ApplyResumeCheckpoint(const AIChatRunCheckpoint& checkpoint);
	std::string BuildResumeFallbackSummary(const AIChatRunCheckpoint& checkpoint) const;

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
	bool m_hasUsage = false;
	size_t m_contextBytesAfterUsage = 0;
	int m_consecutiveFailures = 0;
	std::string m_recoveryHint;
	std::unordered_map<std::string, RepeatedWriteFailureState> m_writeFailures;
	bool m_repeatedWriteFailureStalled = false;
	std::string m_repeatedWriteFailureTarget;
};
