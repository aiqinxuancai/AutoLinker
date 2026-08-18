#include "GameAnalyticsEventBatcher.h"

#include <algorithm>
#include <utility>

namespace {

GameAnalyticsEventBatcher::Policy NormalizePolicy(GameAnalyticsEventBatcher::Policy policy)
{
	policy.flushThreshold = (std::max)(size_t{1}, policy.flushThreshold);
	policy.maxPendingEvents = (std::max)(size_t{1}, policy.maxPendingEvents);
	policy.maxSubmitFailures = (std::max)(1, policy.maxSubmitFailures);
	policy.flushInterval = (std::max)(GameAnalyticsEventBatcher::Duration::zero(), policy.flushInterval);
	policy.retryBaseDelay = (std::max)(GameAnalyticsEventBatcher::Duration::zero(), policy.retryBaseDelay);
	return policy;
}

} // namespace

GameAnalyticsEventBatcher::GameAnalyticsEventBatcher()
	: GameAnalyticsEventBatcher(Policy())
{
}

GameAnalyticsEventBatcher::GameAnalyticsEventBatcher(Policy policy)
	: m_policy(NormalizePolicy(std::move(policy)))
{
}

GameAnalyticsEventBatcher::EnqueueResult GameAnalyticsEventBatcher::Enqueue(
	nlohmann::json event,
	TimePoint now)
{
	EnqueueResult result;
	if (m_pendingEvents.size() >= m_policy.maxPendingEvents) {
		m_pendingEvents.erase(m_pendingEvents.begin());
		result.droppedOldestEvent = true;
	}
	if (m_pendingEvents.empty()) {
		m_pendingDeadline = now + m_policy.flushInterval;
	}
	m_pendingEvents.push_back(std::move(event));
	result.readyForImmediateFlush = !m_retryBatch.has_value() &&
		m_pendingEvents.size() >= m_policy.flushThreshold;
	return result;
}

bool GameAnalyticsEventBatcher::IsReady(TimePoint now) const
{
	if (m_retryBatch.has_value()) {
		return m_retryDeadline.has_value() && now >= *m_retryDeadline;
	}
	if (m_pendingEvents.empty()) {
		return false;
	}
	return m_pendingEvents.size() >= m_policy.flushThreshold ||
		(m_pendingDeadline.has_value() && now >= *m_pendingDeadline);
}

std::optional<GameAnalyticsEventBatcher::TimePoint> GameAnalyticsEventBatcher::NextDeadline() const
{
	if (m_retryBatch.has_value()) {
		return m_retryDeadline;
	}
	return m_pendingDeadline;
}

std::optional<GameAnalyticsEventBatcher::Batch> GameAnalyticsEventBatcher::TakeBatch(
	TimePoint now,
	bool forceRetry,
	bool forcePending)
{
	if (m_retryBatch.has_value()) {
		if (!forceRetry && (!m_retryDeadline.has_value() || now < *m_retryDeadline)) {
			return std::nullopt;
		}
		Batch batch = std::move(*m_retryBatch);
		m_retryBatch.reset();
		m_retryDeadline.reset();
		return batch;
	}
	if (m_pendingEvents.empty()) {
		return std::nullopt;
	}
	if (!forcePending &&
		m_pendingEvents.size() < m_policy.flushThreshold &&
		(!m_pendingDeadline.has_value() || now < *m_pendingDeadline)) {
		return std::nullopt;
	}

	Batch batch;
	batch.events = std::move(m_pendingEvents);
	m_pendingEvents.clear();
	m_pendingDeadline.reset();
	return batch;
}

GameAnalyticsEventBatcher::FailureResult GameAnalyticsEventBatcher::CompleteFailure(
	Batch batch,
	TimePoint now)
{
	FailureResult result;
	result.failureCount = batch.failureCount + 1;
	if (result.failureCount >= m_policy.maxSubmitFailures) {
		result.disposition = FailureDisposition::Dropped;
		return result;
	}

	batch.failureCount = result.failureCount;
	result.disposition = FailureDisposition::RetryScheduled;
	result.retryDelay = RetryDelayForFailureCount(m_policy, result.failureCount);
	m_retryDeadline = now + result.retryDelay;
	m_retryBatch = std::move(batch);
	return result;
}

void GameAnalyticsEventBatcher::Clear()
{
	m_pendingEvents.clear();
	m_pendingDeadline.reset();
	m_retryBatch.reset();
	m_retryDeadline.reset();
}

bool GameAnalyticsEventBatcher::Empty() const noexcept
{
	return m_pendingEvents.empty() && !m_retryBatch.has_value();
}

size_t GameAnalyticsEventBatcher::PendingCount() const noexcept
{
	return m_pendingEvents.size();
}

bool GameAnalyticsEventBatcher::HasRetryBatch() const noexcept
{
	return m_retryBatch.has_value();
}

int GameAnalyticsEventBatcher::RetryFailureCount() const noexcept
{
	return m_retryBatch.has_value() ? m_retryBatch->failureCount : 0;
}

const GameAnalyticsEventBatcher::Policy& GameAnalyticsEventBatcher::GetPolicy() const noexcept
{
	return m_policy;
}

GameAnalyticsEventBatcher::Duration GameAnalyticsEventBatcher::RetryDelayForFailureCount(
	const Policy& policy,
	int failureCount)
{
	const int exponent = (std::clamp)(failureCount - 1, 0, 4);
	const long long multiplier = 1LL << exponent;
	return policy.retryBaseDelay * multiplier;
}
