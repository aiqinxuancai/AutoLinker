#pragma once

#include <Windows.h>

#include <mutex>
#include <string>
#include <unordered_map>
#include <vector>

#include "RealPageCodeToolSupport.h"

// 真实页面源码快照项。
struct PageCodeSnapshotEntry {
	std::string snapshotId;
	std::string code;
	std::string note;
	ULONGLONG createdAtTick = 0;
};

// 真实页面源码缓存项。
struct PageCodeCacheEntry {
	std::string pageName;
	std::string kind;
	std::string pageTypeKey;
	std::string pageTypeName;
	unsigned int itemData = 0;
	std::string code;
	std::string codeHash;
	ULONGLONG updatedAtTick = 0;
	std::vector<PageCodeSnapshotEntry> snapshots;
};

// 当前进程内的真实页面源码缓存管理器。
class PageCodeCacheManager {
public:
	static PageCodeCacheManager& Instance();

	bool Get(const std::string& pageName, const std::string& kind, PageCodeCacheEntry& outEntry) const;
	void Put(const PageCodeCacheEntry& entry);
	void Remove(const std::string& pageName, const std::string& kind);
	void Clear();
	bool AddSnapshot(
		const std::string& pageName,
		const std::string& kind,
		const std::string& code,
		const std::string& note,
		PageCodeSnapshotEntry& outSnapshot);
	bool GetLatestSnapshot(const std::string& pageName, const std::string& kind, PageCodeSnapshotEntry& outSnapshot) const;
	bool GetSnapshot(
		const std::string& pageName,
		const std::string& kind,
		const std::string& snapshotId,
		PageCodeSnapshotEntry& outSnapshot) const;
	bool ListSnapshots(
		const std::string& pageName,
		const std::string& kind,
		std::vector<PageCodeSnapshotEntry>& outSnapshots) const;

private:
	PageCodeCacheManager() = default;

	static std::string BuildKey(const std::string& pageName, const std::string& kind);

	mutable std::mutex m_mutex;
	std::unordered_map<std::string, PageCodeCacheEntry> m_entries;
};
