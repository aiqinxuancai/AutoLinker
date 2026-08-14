#pragma once

#include <string>

// AI 对话 Goal 的持久化状态和状态转换管理器。
enum class AIChatGoalStatus {
	None,
	Active,
	Paused,
	Blocked,
	Complete
};

struct AIChatGoalState {
	AIChatGoalStatus status = AIChatGoalStatus::None;
	std::string objectiveLocal;
	long long tokensUsed = 0;
	long long elapsedMs = 0;
	long long activeStartedAtUnixMs = 0;
	long long createdAtUnixMs = 0;
	long long updatedAtUnixMs = 0;
};

class AIChatGoalManager {
public:
	static const char* StatusToString(AIChatGoalStatus status);
	static AIChatGoalStatus StatusFromString(const std::string& value);
	static bool HasGoal(const AIChatGoalState& goal);
	static bool IsActive(const AIChatGoalState& goal);
	static bool CanAdvance(const AIChatGoalState& goal, bool planModeActive);

	static bool Create(AIChatGoalState& goal, const std::string& objectiveLocal, long long nowUnixMs);
	static bool UpdateObjective(AIChatGoalState& goal, const std::string& objectiveLocal, long long nowUnixMs);
	static bool Pause(AIChatGoalState& goal, long long nowUnixMs);
	static bool Resume(AIChatGoalState& goal, long long nowUnixMs);
	static bool SuspendActiveTiming(AIChatGoalState& goal, long long nowUnixMs);
	static bool ResumeActiveTiming(AIChatGoalState& goal, long long nowUnixMs);
	static bool Complete(AIChatGoalState& goal, long long nowUnixMs);
	static bool Block(AIChatGoalState& goal, long long nowUnixMs);
	static void Clear(AIChatGoalState& goal);
	static void AddUsage(AIChatGoalState& goal, long long tokens, long long nowUnixMs);
	static long long CurrentElapsedMs(const AIChatGoalState& goal, long long nowUnixMs);
	static std::string BuildSelfTestReportJson();

private:
	static void FinishActiveTiming(AIChatGoalState& goal, long long nowUnixMs);
};
