#pragma once

#include <condition_variable>
#include <filesystem>
#include <mutex>
#include <string>
#include <vector>

// 依赖目录缓存：扫描当前易语言进程目录下的 ecom/lib，并为 AI 工具提供可搜索索引。
class DependencyCatalogCache {
public:
	enum class RefreshState {
		Idle,
		Running,
		Ready,
		Failed
	};

	struct Status {
		RefreshState state = RefreshState::Idle;
		bool hasSnapshot = false;
		bool running = false;
		unsigned int refreshGeneration = 0;
		long long lastDurationMs = 0;
		size_t moduleCount = 0;
		size_t libraryCount = 0;
		size_t decodedModuleCount = 0;
		size_t decodedLibraryCount = 0;
		std::string cacheRootLocal;
		std::string ecomRootLocal;
		std::string libRootLocal;
		std::string lastErrorLocal;
	};

	struct ModuleEntry {
		std::string moduleNameLocal;
		std::string fileNameLocal;
		std::string pathLocal;
		std::string cachePathLocal;
		bool decoded = false;
		std::string decodeErrorLocal;
		std::string searchTextUtf8;
	};

	struct LibraryEntry {
		std::string libraryNameLocal;
		std::string fileNameLocal;
		std::string pathLocal;
		std::string cachePathLocal;
		bool decoded = false;
		std::string decodeErrorLocal;
		std::string infoTextLocal;
		std::string searchTextUtf8;
	};

	struct SearchOptions {
		std::string queryUtf8;
		int limit = 50;
		bool includeSnippets = true;
	};

	struct ModuleSearchResult {
		ModuleEntry entry;
		std::string snippetUtf8;
		int score = 0;
	};

	struct LibrarySearchResult {
		LibraryEntry entry;
		std::string snippetUtf8;
		int score = 0;
	};

	static DependencyCatalogCache& Instance();

	// 若当前没有刷新任务，启动后台刷新；force=true 时会重建缓存。
	void StartAsyncRefreshIfNeeded(bool force);

	// 同步刷新，供显式 MCP 工具和测试使用。
	bool RefreshNow(bool force, std::string& outErrorLocal);

	// 等待正在运行的刷新任务结束。
	bool WaitForIdle(unsigned int timeoutMs);

	Status GetStatus() const;
	std::vector<ModuleEntry> SnapshotModules() const;
	std::vector<LibraryEntry> SnapshotLibraries() const;

	std::vector<ModuleSearchResult> SearchModules(const SearchOptions& options);
	std::vector<LibrarySearchResult> SearchLibraries(const SearchOptions& options);

	bool ResolveModulePath(
		const std::string& moduleNameLocal,
		const std::string& modulePathLocal,
		std::string& outPathLocal,
		std::vector<ModuleEntry>* outCandidates,
		std::string& outErrorLocal) const;

	bool ResolveLibrary(
		const std::string& libraryNameLocal,
		const std::string& libraryPathLocal,
		LibraryEntry& outEntry,
		std::vector<LibraryEntry>* outCandidates,
		std::string& outErrorLocal) const;

private:
	DependencyCatalogCache() = default;

	void RefreshWorker(bool force);
	bool RefreshInternal(bool force, std::string& outErrorLocal);

	mutable std::mutex m_mutex;
	std::condition_variable m_cv;
	RefreshState m_state = RefreshState::Idle;
	bool m_hasSnapshot = false;
	bool m_workerRunning = false;
	unsigned int m_refreshGeneration = 0;
	long long m_lastDurationMs = 0;
	size_t m_decodedModuleCount = 0;
	size_t m_decodedLibraryCount = 0;
	std::string m_cacheRootLocal;
	std::string m_ecomRootLocal;
	std::string m_libRootLocal;
	std::string m_lastErrorLocal;
	std::vector<ModuleEntry> m_modules;
	std::vector<LibraryEntry> m_libraries;
};
