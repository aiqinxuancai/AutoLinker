#pragma once

// IDE 全量日志环形缓冲：供日志控件采集器写入，并由日志查看器按序增量读取。

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

namespace IdeLogStore {

enum class EntrySource {
	IdeHook,
	AutoLinkerDirect,
	ExistingOutputSnapshot,
	OutputControlSubclass,
	OutputControlFallback,
};

struct Entry {
	std::uint64_t sequence = 0;
	std::uint64_t unixTimeMs = 0;
	std::uint32_t threadId = 0;
	EntrySource source = EntrySource::OutputControlSubclass;
	std::string localText;
	bool truncated = false;
};

struct Batch {
	std::uint64_t generation = 0;
	std::uint64_t latestSequence = 0;
	std::uint64_t droppedEntries = 0;
	bool resetRequired = false;
	std::vector<Entry> entries;
};

// 采集安全边界：内部吞掉分配异常，且对单条与总缓冲大小设有上限。
void AppendLocal(
	const char* text,
	std::size_t length,
	EntrySource source = EntrySource::OutputControlSubclass) noexcept;

// 仅日志中心打开期间启用 UI 环形缓冲；编译会话 Hook 缓冲不受此开关影响。
void BeginRecording() noexcept;
void EndRecording() noexcept;
bool IsRecordingEnabled() noexcept;

Batch ReadAfter(
	std::uint64_t afterSequence,
	std::uint64_t knownGeneration,
	std::size_t maxEntries,
	std::size_t maxTextBytes = static_cast<std::size_t>(-1));
std::vector<Entry> SnapshotAll();

std::uint64_t GetLatestSequence();
std::uint64_t GetGeneration();
void Clear();

// 无需 IDE 的顺序、丢弃、截断和清空代际自检。
std::string BuildSelfTestJson();

} // namespace IdeLogStore
