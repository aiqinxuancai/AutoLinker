#include "PageCodeCacheManager.h"

#include <algorithm>
#include <cctype>
#include <format>
#include <limits>

namespace {

constexpr size_t kMaxSnapshotsPerPage = 20;
constexpr size_t kMaxPageCodeCacheBytes = 128 * 1024 * 1024;

size_t EstimateSnapshotBytes(const PageCodeSnapshotEntry& snapshot)
{
	return snapshot.snapshotId.size() + snapshot.code.size() + snapshot.note.size();
}

size_t EstimateEntryBytes(const PageCodeCacheEntry& entry)
{
	size_t bytes = entry.pageName.size() + entry.kind.size() + entry.pageTypeKey.size() +
		entry.pageTypeName.size() + entry.code.size() + entry.codeHash.size();
	for (const PageCodeSnapshotEntry& snapshot : entry.snapshots) {
		bytes += EstimateSnapshotBytes(snapshot);
	}
	return bytes;
}

std::string TrimAsciiCopyForCache(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string ToLowerAsciiCopyForCache(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string BuildSnapshotId(const std::string& code, size_t sequence)
{
	return std::format(
		"snapshot-{}-{}-{}",
		GetTickCount64(),
		sequence,
		BuildStableTextHashForRealCode(code).substr(0, 8));
}

}  // namespace

PageCodeCacheManager& PageCodeCacheManager::Instance()
{
	static PageCodeCacheManager instance;
	return instance;
}

bool PageCodeCacheManager::Get(const std::string& pageName, const std::string& kind, PageCodeCacheEntry& outEntry) const
{
	std::lock_guard<std::mutex> lock(m_mutex);
	const auto it = m_entries.find(BuildKey(pageName, kind));
	if (it == m_entries.end()) {
		return false;
	}
	outEntry = it->second;
	return true;
}

void PageCodeCacheManager::Put(const PageCodeCacheEntry& entry)
{
	PageCodeCacheEntry saved = entry;
	saved.pageName = TrimAsciiCopyForCache(saved.pageName);
	saved.kind = ToLowerAsciiCopyForCache(TrimAsciiCopyForCache(saved.kind));
	saved.code = NormalizeRealCodeLineBreaksToCrLf(saved.code);
	saved.codeHash = BuildStableTextHashForRealCode(saved.code);
	saved.updatedAtTick = GetTickCount64();

	std::lock_guard<std::mutex> lock(m_mutex);
	const std::string key = BuildKey(saved.pageName, saved.kind);
	const auto existing = m_entries.find(key);
	if (existing != m_entries.end() && saved.snapshots.empty()) {
		saved.snapshots = existing->second.snapshots;
	}
	if (saved.snapshots.size() > kMaxSnapshotsPerPage) {
		saved.snapshots.erase(saved.snapshots.begin(), saved.snapshots.end() - static_cast<std::ptrdiff_t>(kMaxSnapshotsPerPage));
	}
	m_entries[key] = std::move(saved);
	TrimToBudgetLocked();
}

void PageCodeCacheManager::Remove(const std::string& pageName, const std::string& kind)
{
	std::lock_guard<std::mutex> lock(m_mutex);
	m_entries.erase(BuildKey(pageName, kind));
}

void PageCodeCacheManager::Clear()
{
	std::lock_guard<std::mutex> lock(m_mutex);
	m_entries.clear();
}

bool PageCodeCacheManager::AddSnapshot(
	const std::string& pageName,
	const std::string& kind,
	const std::string& code,
	const std::string& note,
	PageCodeSnapshotEntry& outSnapshot)
{
	outSnapshot = {};

	std::lock_guard<std::mutex> lock(m_mutex);
	const std::string key = BuildKey(pageName, kind);
	PageCodeCacheEntry& entry = m_entries[key];
	entry.pageName = TrimAsciiCopyForCache(pageName);
	entry.kind = ToLowerAsciiCopyForCache(TrimAsciiCopyForCache(kind));
	if (entry.code.empty()) {
		entry.code = NormalizeRealCodeLineBreaksToCrLf(code);
	}
	if (entry.codeHash.empty()) {
		entry.codeHash = BuildStableTextHashForRealCode(entry.code);
	}

	PageCodeSnapshotEntry snapshot;
	snapshot.code = NormalizeRealCodeLineBreaksToCrLf(code);
	snapshot.note = note;
	snapshot.createdAtTick = GetTickCount64();
	snapshot.snapshotId = BuildSnapshotId(snapshot.code, entry.snapshots.size() + 1);
	entry.snapshots.push_back(snapshot);
	if (entry.snapshots.size() > kMaxSnapshotsPerPage) {
		entry.snapshots.erase(entry.snapshots.begin(), entry.snapshots.end() - static_cast<std::ptrdiff_t>(kMaxSnapshotsPerPage));
	}
	outSnapshot = snapshot;
	TrimToBudgetLocked(snapshot.snapshotId);
	return true;
}

bool PageCodeCacheManager::GetLatestSnapshot(
	const std::string& pageName,
	const std::string& kind,
	PageCodeSnapshotEntry& outSnapshot) const
{
	outSnapshot = {};

	std::lock_guard<std::mutex> lock(m_mutex);
	const auto it = m_entries.find(BuildKey(pageName, kind));
	if (it == m_entries.end() || it->second.snapshots.empty()) {
		return false;
	}
	outSnapshot = it->second.snapshots.back();
	return true;
}

bool PageCodeCacheManager::GetSnapshot(
	const std::string& pageName,
	const std::string& kind,
	const std::string& snapshotId,
	PageCodeSnapshotEntry& outSnapshot) const
{
	outSnapshot = {};

	std::lock_guard<std::mutex> lock(m_mutex);
	const auto it = m_entries.find(BuildKey(pageName, kind));
	if (it == m_entries.end()) {
		return false;
	}

	for (const auto& snapshot : it->second.snapshots) {
		if (snapshot.snapshotId == snapshotId) {
			outSnapshot = snapshot;
			return true;
		}
	}
	return false;
}

bool PageCodeCacheManager::ListSnapshots(
	const std::string& pageName,
	const std::string& kind,
	std::vector<PageCodeSnapshotEntry>& outSnapshots) const
{
	outSnapshots.clear();

	std::lock_guard<std::mutex> lock(m_mutex);
	const auto it = m_entries.find(BuildKey(pageName, kind));
	if (it == m_entries.end()) {
		return false;
	}

	outSnapshots = it->second.snapshots;
	return true;
}

std::string PageCodeCacheManager::BuildKey(const std::string& pageName, const std::string& kind)
{
	return ToLowerAsciiCopyForCache(TrimAsciiCopyForCache(kind)) + "\n" + TrimAsciiCopyForCache(pageName);
}

void PageCodeCacheManager::TrimToBudgetLocked(const std::string& protectedSnapshotId)
{
	const auto totalBytes = [this]() {
		size_t total = 0;
		for (const auto& [key, entry] : m_entries) {
			(void)key;
			total += EstimateEntryBytes(entry);
		}
		return total;
	};

	size_t currentBytes = totalBytes();
	while (currentBytes > kMaxPageCodeCacheBytes) {
		auto oldestEntry = m_entries.end();
		size_t oldestIndex = 0;
		ULONGLONG oldestTick = (std::numeric_limits<ULONGLONG>::max)();
		for (auto entryIt = m_entries.begin(); entryIt != m_entries.end(); ++entryIt) {
			for (size_t i = 0; i < entryIt->second.snapshots.size(); ++i) {
				const PageCodeSnapshotEntry& snapshot = entryIt->second.snapshots[i];
				if (snapshot.snapshotId == protectedSnapshotId) {
					continue;
				}
				if (snapshot.createdAtTick < oldestTick) {
					oldestTick = snapshot.createdAtTick;
					oldestEntry = entryIt;
					oldestIndex = i;
				}
			}
		}
		if (oldestEntry == m_entries.end()) {
			break;
		}
		currentBytes -= EstimateSnapshotBytes(oldestEntry->second.snapshots[oldestIndex]);
		oldestEntry->second.snapshots.erase(
			oldestEntry->second.snapshots.begin() + static_cast<std::ptrdiff_t>(oldestIndex));
	}

	while (currentBytes > kMaxPageCodeCacheBytes) {
		auto oldestEntry = m_entries.end();
		ULONGLONG oldestTick = (std::numeric_limits<ULONGLONG>::max)();
		for (auto entryIt = m_entries.begin(); entryIt != m_entries.end(); ++entryIt) {
			const bool protectedEntry = std::any_of(
				entryIt->second.snapshots.begin(),
				entryIt->second.snapshots.end(),
				[&protectedSnapshotId](const PageCodeSnapshotEntry& snapshot) {
					return !protectedSnapshotId.empty() && snapshot.snapshotId == protectedSnapshotId;
				});
			if (!protectedEntry && entryIt->second.updatedAtTick < oldestTick) {
				oldestTick = entryIt->second.updatedAtTick;
				oldestEntry = entryIt;
			}
		}
		if (oldestEntry == m_entries.end()) {
			break;
		}
		currentBytes -= EstimateEntryBytes(oldestEntry->second);
		m_entries.erase(oldestEntry);
	}

	if (currentBytes > kMaxPageCodeCacheBytes && !protectedSnapshotId.empty()) {
		for (auto& [key, entry] : m_entries) {
			(void)key;
			const bool protectedEntry = std::any_of(
				entry.snapshots.begin(),
				entry.snapshots.end(),
				[&protectedSnapshotId](const PageCodeSnapshotEntry& snapshot) {
					return snapshot.snapshotId == protectedSnapshotId;
				});
			if (protectedEntry && !entry.code.empty()) {
				currentBytes -= entry.code.size();
				entry.code.clear();
				entry.code.shrink_to_fit();
				entry.codeHash.clear();
				break;
			}
		}
	}
}
