#pragma once

#include <Windows.h>

#include <mutex>
#include <string>
#include <unordered_map>

// 真实页面源码缓存项。
struct PageCodeCacheEntry {
	std::string pageName;
	std::string kind;
	std::string pageTypeKey;
	std::string pageTypeName;
	unsigned int itemData = 0;
	std::string code;
	ULONGLONG updatedAtTick = 0;
};

// 当前进程内的真实页面源码缓存管理器。
class PageCodeCacheManager {
public:
	static PageCodeCacheManager& Instance();

	bool Get(const std::string& pageName, const std::string& kind, PageCodeCacheEntry& outEntry) const;
	void Put(const PageCodeCacheEntry& entry);
	void Remove(const std::string& pageName, const std::string& kind);
	void Clear();

private:
	PageCodeCacheManager() = default;

	static std::string BuildKey(const std::string& pageName, const std::string& kind);

	mutable std::mutex m_mutex;
	std::unordered_map<std::string, PageCodeCacheEntry> m_entries;
};
