#include "AIChatGoalManager.h"

#include <algorithm>
#include <format>

const char* AIChatGoalManager::StatusToString(AIChatGoalStatus status)
{
	switch (status) {
	case AIChatGoalStatus::Active:
		return "active";
	case AIChatGoalStatus::Paused:
		return "paused";
	case AIChatGoalStatus::Blocked:
		return "blocked";
	case AIChatGoalStatus::Complete:
		return "complete";
	case AIChatGoalStatus::None:
	default:
		return "none";
	}
}

AIChatGoalStatus AIChatGoalManager::StatusFromString(const std::string& value)
{
	if (value == "active") return AIChatGoalStatus::Active;
	if (value == "paused") return AIChatGoalStatus::Paused;
	if (value == "blocked") return AIChatGoalStatus::Blocked;
	if (value == "complete") return AIChatGoalStatus::Complete;
	return AIChatGoalStatus::None;
}

bool AIChatGoalManager::HasGoal(const AIChatGoalState& goal)
{
	return goal.status != AIChatGoalStatus::None && !goal.objectiveLocal.empty();
}

bool AIChatGoalManager::IsActive(const AIChatGoalState& goal)
{
	return HasGoal(goal) && goal.status == AIChatGoalStatus::Active;
}

bool AIChatGoalManager::CanAdvance(const AIChatGoalState& goal, bool planModeActive)
{
	return IsActive(goal) && !planModeActive;
}

bool AIChatGoalManager::Create(AIChatGoalState& goal, const std::string& objectiveLocal, long long nowUnixMs)
{
	if (HasGoal(goal) || objectiveLocal.empty()) return false;
	goal = {};
	goal.status = AIChatGoalStatus::Active;
	goal.objectiveLocal = objectiveLocal;
	goal.activeStartedAtUnixMs = nowUnixMs;
	goal.createdAtUnixMs = nowUnixMs;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::UpdateObjective(AIChatGoalState& goal, const std::string& objectiveLocal, long long nowUnixMs)
{
	if (!HasGoal(goal) || objectiveLocal.empty()) return false;
	goal.objectiveLocal = objectiveLocal;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::Pause(AIChatGoalState& goal, long long nowUnixMs)
{
	if (!IsActive(goal)) return false;
	FinishActiveTiming(goal, nowUnixMs);
	goal.status = AIChatGoalStatus::Paused;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::Resume(AIChatGoalState& goal, long long nowUnixMs)
{
	if (!HasGoal(goal) || (goal.status != AIChatGoalStatus::Paused && goal.status != AIChatGoalStatus::Blocked)) return false;
	goal.status = AIChatGoalStatus::Active;
	goal.activeStartedAtUnixMs = nowUnixMs;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::SuspendActiveTiming(AIChatGoalState& goal, long long nowUnixMs)
{
	if (!IsActive(goal) || goal.activeStartedAtUnixMs <= 0) return false;
	FinishActiveTiming(goal, nowUnixMs);
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::ResumeActiveTiming(AIChatGoalState& goal, long long nowUnixMs)
{
	if (!IsActive(goal) || goal.activeStartedAtUnixMs > 0) return false;
	goal.activeStartedAtUnixMs = nowUnixMs;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::Complete(AIChatGoalState& goal, long long nowUnixMs)
{
	if (!HasGoal(goal) || goal.status == AIChatGoalStatus::Complete) return false;
	FinishActiveTiming(goal, nowUnixMs);
	goal.status = AIChatGoalStatus::Complete;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

bool AIChatGoalManager::Block(AIChatGoalState& goal, long long nowUnixMs)
{
	if (!IsActive(goal)) return false;
	FinishActiveTiming(goal, nowUnixMs);
	goal.status = AIChatGoalStatus::Blocked;
	goal.updatedAtUnixMs = nowUnixMs;
	return true;
}

void AIChatGoalManager::Clear(AIChatGoalState& goal)
{
	goal = {};
}

void AIChatGoalManager::AddUsage(AIChatGoalState& goal, long long tokens, long long nowUnixMs)
{
	if (!HasGoal(goal)) return;
	goal.tokensUsed += (std::max)(0LL, tokens);
	goal.updatedAtUnixMs = nowUnixMs;
}

long long AIChatGoalManager::CurrentElapsedMs(const AIChatGoalState& goal, long long nowUnixMs)
{
	long long elapsed = (std::max)(0LL, goal.elapsedMs);
	if (goal.status == AIChatGoalStatus::Active && goal.activeStartedAtUnixMs > 0 && nowUnixMs > goal.activeStartedAtUnixMs) {
		elapsed += nowUnixMs - goal.activeStartedAtUnixMs;
	}
	return elapsed;
}

void AIChatGoalManager::FinishActiveTiming(AIChatGoalState& goal, long long nowUnixMs)
{
	if (goal.status == AIChatGoalStatus::Active && goal.activeStartedAtUnixMs > 0 && nowUnixMs > goal.activeStartedAtUnixMs) {
		goal.elapsedMs += nowUnixMs - goal.activeStartedAtUnixMs;
	}
	goal.activeStartedAtUnixMs = 0;
}

std::string AIChatGoalManager::BuildSelfTestReportJson()
{
	AIChatGoalState goal;
	const bool created = Create(goal, "ship goal mode", 1000);
	const bool duplicateRejected = !Create(goal, "replacement", 1100);
	const bool activeElapsed = CurrentElapsedMs(goal, 1500) == 500;
	AddUsage(goal, 42, 1500);
	const bool objectiveUpdated = UpdateObjective(goal, "ship updated goal mode", 1550) &&
		goal.objectiveLocal == "ship updated goal mode" &&
		goal.status == AIChatGoalStatus::Active &&
		goal.tokensUsed == 42 &&
		CurrentElapsedMs(goal, 1550) == 550;
	const bool paused = Pause(goal, 1600) && goal.status == AIChatGoalStatus::Paused && goal.elapsedMs == 600;
	const bool invalidPauseRejected = !Pause(goal, 1700);
	const bool resumed = Resume(goal, 2000) && goal.status == AIChatGoalStatus::Active;
	const bool blocked = Block(goal, 2400) && goal.status == AIChatGoalStatus::Blocked && goal.elapsedMs == 1000;
	const bool resumedFromBlocked = Resume(goal, 2500) && goal.status == AIChatGoalStatus::Active;
	const bool completed = Complete(goal, 3000) && goal.status == AIChatGoalStatus::Complete && goal.elapsedMs == 1500;
	const bool usageTracked = goal.tokensUsed == 42;
	const bool statusRoundTrip = StatusFromString(StatusToString(goal.status)) == AIChatGoalStatus::Complete;
	Clear(goal);
	const bool cleared = !HasGoal(goal) && goal.status == AIChatGoalStatus::None;

	AIChatGoalState planGoal;
	const bool planGoalCreated = Create(planGoal, "plan isolation", 4000);
	const bool timingSuspended = SuspendActiveTiming(planGoal, 4300) &&
		planGoal.status == AIChatGoalStatus::Active &&
		CurrentElapsedMs(planGoal, 5000) == 300;
	const bool planBlocksAdvance = !CanAdvance(planGoal, true);
	const bool timingResumed = ResumeActiveTiming(planGoal, 5200) &&
		CurrentElapsedMs(planGoal, 5500) == 600;
	const bool defaultAllowsAdvance = CanAdvance(planGoal, false);
	const bool ok = created && duplicateRejected && activeElapsed && objectiveUpdated && paused && invalidPauseRejected &&
		resumed && blocked && resumedFromBlocked && completed && usageTracked && statusRoundTrip && cleared &&
		planGoalCreated && timingSuspended && planBlocksAdvance && timingResumed && defaultAllowsAdvance;
	const auto flag = [](bool value) { return value ? "true" : "false"; };
	return std::format(
		R"({{"name":"goal-mode-state-machine","ok":{},"created":{},"duplicate_rejected":{},"active_elapsed":{},"objective_updated":{},"paused":{},"invalid_pause_rejected":{},"resumed":{},"blocked":{},"resumed_from_blocked":{},"completed":{},"usage_tracked":{},"status_round_trip":{},"cleared":{},"plan_goal_created":{},"timing_suspended":{},"plan_blocks_advance":{},"timing_resumed":{},"default_allows_advance":{}}})",
		flag(ok), flag(created), flag(duplicateRejected), flag(activeElapsed), flag(objectiveUpdated), flag(paused),
		flag(invalidPauseRejected), flag(resumed), flag(blocked), flag(resumedFromBlocked),
		flag(completed), flag(usageTracked), flag(statusRoundTrip), flag(cleared),
		flag(planGoalCreated), flag(timingSuspended), flag(planBlocksAdvance),
		flag(timingResumed), flag(defaultAllowsAdvance));
}
