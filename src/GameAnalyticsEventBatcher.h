#pragma once

// GameAnalytics 事件批次调度器：负责数量阈值、时间阈值和失败批次退避。

#include <chrono>
#include <cstddef>
#include <optional>
#include <vector>

#include "..\thirdparty\json.hpp"

class GameAnalyticsEventBatcher final {
public:
	using Clock = std::chrono::steady_clock;
	using TimePoint = Clock::time_point;
	using Duration = std::chrono::milliseconds;

	struct Policy {
		size_t flushThreshold = 20;
		size_t maxPendingEvents = 128;
		int maxSubmitFailures = 3;
		Duration flushInterval = std::chrono::seconds(30);
		Duration retryBaseDelay = std::chrono::seconds(1);
	};

	struct Batch {
		std::vector<nlohmann::json> events;
		int failureCount = 0;
	};

	struct EnqueueResult {
		bool readyForImmediateFlush = false;
		bool droppedOldestEvent = false;
	};

	enum class FailureDisposition {
		RetryScheduled,
		Dropped
	};

	struct FailureResult {
		FailureDisposition disposition = FailureDisposition::Dropped;
		int failureCount = 0;
		Duration retryDelay = Duration::zero();
	};

	GameAnalyticsEventBatcher();
	explicit GameAnalyticsEventBatcher(Policy policy);

	EnqueueResult Enqueue(nlohmann::json event, TimePoint now);
	bool IsReady(TimePoint now) const;
	std::optional<TimePoint> NextDeadline() const;
	std::optional<Batch> TakeBatch(
		TimePoint now,
		bool forceRetry,
		bool forcePending);
	FailureResult CompleteFailure(Batch batch, TimePoint now);

	void Clear();
	bool Empty() const noexcept;
	size_t PendingCount() const noexcept;
	bool HasRetryBatch() const noexcept;
	int RetryFailureCount() const noexcept;
	const Policy& GetPolicy() const noexcept;

	static Duration RetryDelayForFailureCount(const Policy& policy, int failureCount);

private:
	Policy m_policy;
	std::vector<nlohmann::json> m_pendingEvents;
	std::optional<TimePoint> m_pendingDeadline;
	std::optional<Batch> m_retryBatch;
	std::optional<TimePoint> m_retryDeadline;
};
