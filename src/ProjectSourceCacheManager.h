#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace project_source_cache {

// 工程源码页面缓存项。
struct Page {
	int pageIndex = -1;
	std::string typeKey;
	std::string typeName;
	std::string name;
	std::vector<std::string> lines;
};

// 工程源码解析快照。
struct Snapshot {
	bool ok = false;
	uint64_t revision = 0;
	std::string sourcePath;
	std::string snapshotPath;
	std::string parsedInputPath;
	std::string parsedInputKind;
	std::string projectName;
	std::string versionText;
	std::vector<Page> pages;
};

// 工程源码命中令牌。
struct HitToken {
	std::string sourcePath;
	uint64_t revision = 0;
	int pageIndex = -1;
	int lineNumber = -1;
	std::string pageName;
	std::string pageTypeKey;
	std::string pageTypeName;
};

// 当前工程源码解析缓存管理器。
class ProjectSourceCacheManager {
public:
	static ProjectSourceCacheManager& Instance();

	// 清空当前进程内已缓存的工程源码快照。
	void Clear();

	bool WarmupCurrentSource(std::string* outError = nullptr, std::string* outTrace = nullptr);

	bool EnsureCurrentSourceLatest(
		Snapshot& outSnapshot,
		bool forceRefresh,
		bool* outRefreshed = nullptr,
		std::string* outError = nullptr,
		std::string* outTrace = nullptr);

	bool ResolveSnapshotForHit(
		const HitToken& token,
		bool refreshIfStale,
		Snapshot& outSnapshot,
		bool* outRefreshed = nullptr,
		std::string* outError = nullptr,
		std::string* outTrace = nullptr);

	bool GetSnapshotCopy(Snapshot& outSnapshot) const;

	std::string BuildHitToken(const HitToken& token) const;
	bool ParseHitToken(const std::string& text, HitToken& outToken, std::string* outError = nullptr) const;

private:
	ProjectSourceCacheManager() = default;
	ProjectSourceCacheManager(const ProjectSourceCacheManager&) = delete;
	ProjectSourceCacheManager& operator=(const ProjectSourceCacheManager&) = delete;
};

} // namespace project_source_cache
