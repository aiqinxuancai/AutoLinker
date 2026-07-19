#include "IdeLogStore.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <deque>
#include <limits>
#include <mutex>
#include <string>
#include <utility>

#include "..\\thirdparty\\json.hpp"

namespace IdeLogStore {
namespace {

constexpr std::size_t kMaxEntries = 40000;
constexpr std::size_t kMaxBufferedBytes = 16 * 1024 * 1024;
constexpr std::size_t kMaxEntryBytes = 256 * 1024;
constexpr std::uint64_t kWindowsToUnixEpoch100ns = 116444736000000000ULL;

std::uint64_t GetUnixTimeMilliseconds()
{
	FILETIME fileTime = {};
	GetSystemTimeAsFileTime(&fileTime);
	ULARGE_INTEGER value = {};
	value.LowPart = fileTime.dwLowDateTime;
	value.HighPart = fileTime.dwHighDateTime;
	if (value.QuadPart <= kWindowsToUnixEpoch100ns) {
		return 0;
	}
	return (value.QuadPart - kWindowsToUnixEpoch100ns) / 10000ULL;
}

class LogRing final {
public:
	LogRing(std::size_t maxEntries, std::size_t maxBytes, std::size_t maxEntryBytes)
		: m_maxEntries((std::max)(static_cast<std::size_t>(1), maxEntries)),
		m_maxBytes((std::max)(static_cast<std::size_t>(1), maxBytes)),
		m_maxEntryBytes((std::max)(static_cast<std::size_t>(1), maxEntryBytes))
	{
	}

	void Append(
		const char* text,
		std::size_t length,
		EntrySource source,
		const std::atomic_bool* enabledGate = nullptr)
	{
		if (text == nullptr || length == 0 ||
			(enabledGate != nullptr && !enabledGate->load(std::memory_order_acquire))) {
			return;
		}

		const std::size_t copyLength = (std::min)(length, m_maxEntryBytes);
		Entry entry;
		entry.unixTimeMs = GetUnixTimeMilliseconds();
		entry.threadId = GetCurrentThreadId();
		entry.source = source;
		entry.localText.assign(text, copyLength);
		entry.truncated = copyLength < length;

		std::lock_guard<std::mutex> lock(m_mutex);
		if (enabledGate != nullptr && !enabledGate->load(std::memory_order_acquire)) {
			return;
		}
		entry.sequence = ++m_latestSequence;
		m_bufferedBytes += entry.localText.size();
		m_entries.push_back(std::move(entry));
		TrimLocked();
	}

	Batch ReadAfter(
		std::uint64_t afterSequence,
		std::uint64_t knownGeneration,
		std::size_t maxEntries,
		std::size_t maxTextBytes)
	{
		Batch result;
		std::lock_guard<std::mutex> lock(m_mutex);
		result.generation = m_generation;
		result.latestSequence = m_latestSequence;
		result.droppedEntries = m_droppedEntries;
		result.resetRequired = knownGeneration != m_generation;

		if (!m_entries.empty() &&
			afterSequence != 0 &&
			afterSequence < m_entries.front().sequence - 1) {
			result.resetRequired = true;
		}

		const std::size_t limit = (std::max)(static_cast<std::size_t>(1), maxEntries);
		const std::size_t textByteLimit = (std::max)(static_cast<std::size_t>(1), maxTextBytes);
		std::size_t copiedTextBytes = 0;
		for (const Entry& entry : m_entries) {
			if (!result.resetRequired && entry.sequence <= afterSequence) {
				continue;
			}
			if (!result.entries.empty() &&
				entry.localText.size() > textByteLimit - (std::min)(copiedTextBytes, textByteLimit)) {
				break;
			}
			result.entries.push_back(entry);
			copiedTextBytes += entry.localText.size();
			if (result.entries.size() >= limit) {
				break;
			}
		}
		return result;
	}

	std::vector<Entry> SnapshotAll()
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		return {m_entries.begin(), m_entries.end()};
	}

	std::uint64_t LatestSequence()
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		return m_latestSequence;
	}

	std::uint64_t Generation()
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		return m_generation;
	}

	void Clear(bool resetDroppedEntries = false)
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_entries.clear();
		m_bufferedBytes = 0;
		if (resetDroppedEntries) {
			m_droppedEntries = 0;
		}
		++m_generation;
		if (m_generation == 0) {
			++m_generation;
		}
	}

private:
	void TrimLocked()
	{
		while (!m_entries.empty() &&
			(m_entries.size() > m_maxEntries || m_bufferedBytes > m_maxBytes)) {
			m_bufferedBytes -= m_entries.front().localText.size();
			m_entries.pop_front();
			++m_droppedEntries;
		}
	}

	const std::size_t m_maxEntries;
	const std::size_t m_maxBytes;
	const std::size_t m_maxEntryBytes;
	std::mutex m_mutex;
	std::deque<Entry> m_entries;
	std::size_t m_bufferedBytes = 0;
	std::uint64_t m_latestSequence = 0;
	std::uint64_t m_droppedEntries = 0;
	std::uint64_t m_generation = 1;
};

LogRing g_logRing(kMaxEntries, kMaxBufferedBytes, kMaxEntryBytes);
std::atomic_bool g_recordingEnabled = false;

} // namespace

void AppendLocal(const char* text, std::size_t length, EntrySource source) noexcept
{
	if (!g_recordingEnabled.load(std::memory_order_acquire)) {
		return;
	}
	try {
		g_logRing.Append(text, length, source, &g_recordingEnabled);
	}
	catch (...) {
		// 日志捕获绝不能影响 IDE 原始输出调用。
	}
}

void BeginRecording() noexcept
{
	g_recordingEnabled.store(false, std::memory_order_release);
	try {
		g_logRing.Clear(true);
		g_recordingEnabled.store(true, std::memory_order_release);
	}
	catch (...) {
		g_recordingEnabled.store(false, std::memory_order_release);
	}
}

void EndRecording() noexcept
{
	if (!g_recordingEnabled.exchange(false, std::memory_order_acq_rel)) {
		return;
	}
	try {
		g_logRing.Clear(true);
	}
	catch (...) {
	}
}

bool IsRecordingEnabled() noexcept
{
	return g_recordingEnabled.load(std::memory_order_acquire);
}

Batch ReadAfter(
	std::uint64_t afterSequence,
	std::uint64_t knownGeneration,
	std::size_t maxEntries,
	std::size_t maxTextBytes)
{
	return g_logRing.ReadAfter(afterSequence, knownGeneration, maxEntries, maxTextBytes);
}

std::vector<Entry> SnapshotAll()
{
	return g_logRing.SnapshotAll();
}

std::uint64_t GetLatestSequence()
{
	return g_logRing.LatestSequence();
}

std::uint64_t GetGeneration()
{
	return g_logRing.Generation();
}

void Clear()
{
	g_logRing.Clear();
}

std::string BuildSelfTestJson()
{
	LogRing ring(3, 12, 5);
	ring.Append("one", 3, EntrySource::IdeHook);
	ring.Append("two", 3, EntrySource::IdeHook);
	const Batch ordered = ring.ReadAfter(0, ring.Generation(), 10, 100);
	const bool orderedRead = ordered.entries.size() == 2 &&
		ordered.entries[0].sequence < ordered.entries[1].sequence &&
		ordered.entries[0].localText == "one" &&
		ordered.entries[1].localText == "two";

	ring.Append("123456789", 9, EntrySource::ExistingOutputSnapshot);
	const auto afterTruncate = ring.SnapshotAll();
	const bool entryTruncated = !afterTruncate.empty() &&
		afterTruncate.back().localText == "12345" &&
		afterTruncate.back().truncated;

	ring.Append("four", 4, EntrySource::IdeHook);
	ring.Append("five", 4, EntrySource::IdeHook);
	const Batch dropped = ring.ReadAfter(1, ring.Generation(), 10, 100);
	const bool droppedReset = dropped.resetRequired && dropped.droppedEntries > 0 &&
		dropped.entries.size() <= 3;

	const std::uint64_t previousGeneration = ring.Generation();
	ring.Clear();
	const Batch cleared = ring.ReadAfter(0, previousGeneration, 10, 100);
	const bool clearGenerationReset = cleared.resetRequired &&
		cleared.generation != previousGeneration && cleared.entries.empty();
	LogRing payloadRing(10, 100, 20);
	payloadRing.Append("1234", 4, EntrySource::IdeHook);
	payloadRing.Append("5678", 4, EntrySource::IdeHook);
	const Batch payloadLimited = payloadRing.ReadAfter(0, payloadRing.Generation(), 10, 5);
	const bool payloadByteBudgetPassed = payloadLimited.entries.size() == 1;
	LogRing sourceRing(2, 100, 50);
	constexpr char kDirectText[] = "[AutoLinker]direct";
	sourceRing.Append(kDirectText, sizeof(kDirectText) - 1, EntrySource::AutoLinkerDirect);
	const Batch directSourceBatch = sourceRing.ReadAfter(0, sourceRing.Generation(), 2, 100);
	const bool directSourcePreserved = directSourceBatch.entries.size() == 1 &&
		directSourceBatch.entries.front().source == EntrySource::AutoLinkerDirect &&
		directSourceBatch.entries.front().localText == "[AutoLinker]direct";
	std::atomic_bool recordingGate = false;
	LogRing gatedRing(2, 100, 50);
	gatedRing.Append("closed", 6, EntrySource::IdeHook, &recordingGate);
	const bool closedRecordingSkipped = gatedRing.SnapshotAll().empty();
	recordingGate.store(true, std::memory_order_release);
	gatedRing.Append("open", 4, EntrySource::IdeHook, &recordingGate);
	const auto gatedEntries = gatedRing.SnapshotAll();
	const bool openRecordingAccepted = gatedEntries.size() == 1 && gatedEntries.front().localText == "open";
	const bool recordingGatePassed = closedRecordingSkipped && openRecordingAccepted;

	const bool ok = orderedRead && entryTruncated && droppedReset && clearGenerationReset &&
		payloadByteBudgetPassed && directSourcePreserved && recordingGatePassed;
	return nlohmann::json({
		{"name", "ide-log-store"},
		{"ok", ok},
		{"ordered_read", orderedRead},
		{"entry_truncated", entryTruncated},
		{"dropped_range_resets", droppedReset},
		{"clear_generation_resets", clearGenerationReset},
		{"payload_byte_budget_passed", payloadByteBudgetPassed},
		{"autolinker_direct_source_preserved", directSourcePreserved},
		{"recording_gate_passed", recordingGatePassed}
	}).dump();
}

} // namespace IdeLogStore
