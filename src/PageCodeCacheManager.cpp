#include "PageCodeCacheManager.h"

#include <algorithm>
#include <cctype>

namespace {

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

}

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
	saved.updatedAtTick = GetTickCount64();
	std::lock_guard<std::mutex> lock(m_mutex);
	m_entries[BuildKey(saved.pageName, saved.kind)] = std::move(saved);
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

std::string PageCodeCacheManager::BuildKey(const std::string& pageName, const std::string& kind)
{
	return ToLowerAsciiCopyForCache(TrimAsciiCopyForCache(kind)) + "\n" + TrimAsciiCopyForCache(pageName);
}
