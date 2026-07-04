#include "DependencyCatalogCache.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#include <lib2.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <fstream>
#include <sstream>
#include <thread>
#include <unordered_set>

#include "EPackagerIntegration.h"
#include "Logger.h"
#include "PathHelper.h"

namespace {

constexpr int kDefaultSearchLimit = 50;
constexpr int kMaxSearchLimit = 200;
constexpr size_t kMaxIndexedTextBytes = 512 * 1024;

using Clock = std::chrono::steady_clock;

std::string TrimAsciiCopy(std::string text)
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

std::string ToLowerAsciiCopy(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

bool EqualsInsensitiveAscii(const std::string& left, const std::string& right)
{
	return ToLowerAsciiCopy(left) == ToLowerAsciiCopy(right);
}

std::wstring WideFromCodePage(const std::string& text, UINT codePage, DWORD flags = 0)
{
	if (text.empty()) {
		return {};
	}
	const int wideLen = MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLen <= 0) {
		return {};
	}
	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), wide.data(), wideLen) <= 0) {
		return {};
	}
	return wide;
}

std::string StringFromWideCodePage(const std::wstring& text, UINT codePage)
{
	if (text.empty()) {
		return {};
	}
	const int outLen = WideCharToMultiByte(codePage, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (outLen <= 0) {
		return {};
	}
	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(codePage, 0, text.data(), static_cast<int>(text.size()), out.data(), outLen, nullptr, nullptr) <= 0) {
		return {};
	}
	return out;
}

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
	if (text.empty()) {
		return {};
	}
	const std::wstring wide = WideFromCodePage(text, fromCodePage, fromFlags);
	if (wide.empty()) {
		return text;
	}
	const std::string converted = StringFromWideCodePage(wide, toCodePage);
	return converted.empty() ? text : converted;
}

std::string LocalToUtf8(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	if (IsValidUtf8(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8);
}

std::string Utf8ToLocal(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	if (!IsValidUtf8(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string BytesToUtf8Text(std::string bytes)
{
	if (bytes.size() >= 3 &&
		static_cast<unsigned char>(bytes[0]) == 0xEF &&
		static_cast<unsigned char>(bytes[1]) == 0xBB &&
		static_cast<unsigned char>(bytes[2]) == 0xBF) {
		bytes.erase(0, 3);
	}
	if (IsValidUtf8(bytes)) {
		return bytes;
	}
	return LocalToUtf8(bytes);
}

bool ReadFileBytesLimited(const std::filesystem::path& path, std::string& outBytes)
{
	outBytes.clear();
	std::ifstream file(path, std::ios::binary);
	if (!file.is_open()) {
		return false;
	}

	std::ostringstream ss;
	char buffer[8192] = {};
	size_t total = 0;
	while (file.good() && total < kMaxIndexedTextBytes) {
		const size_t want = (std::min)(sizeof(buffer), kMaxIndexedTextBytes - total);
		file.read(buffer, static_cast<std::streamsize>(want));
		const std::streamsize got = file.gcount();
		if (got <= 0) {
			break;
		}
		ss.write(buffer, got);
		total += static_cast<size_t>(got);
	}
	outBytes = ss.str();
	return true;
}

bool WriteTextUtf8Bom(const std::filesystem::path& path, const std::string& textUtf8, std::string& outError)
{
	outError.clear();
	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);
	if (ec) {
		outError = "create directory failed: " + ec.message();
		return false;
	}

	std::ofstream file(path, std::ios::binary | std::ios::trunc);
	if (!file.is_open()) {
		outError = "open output file failed";
		return false;
	}
	static constexpr unsigned char kBom[] = { 0xEF, 0xBB, 0xBF };
	file.write(reinterpret_cast<const char*>(kBom), sizeof(kBom));
	file.write(textUtf8.data(), static_cast<std::streamsize>(textUtf8.size()));
	return file.good();
}

std::string PathStemLocal(const std::filesystem::path& path)
{
	return path.stem().string();
}

std::string PathFileNameLocal(const std::filesystem::path& path)
{
	return path.filename().string();
}

uint64_t Fnv1a64(const std::string& text)
{
	uint64_t hash = 1469598103934665603ull;
	for (unsigned char ch : text) {
		hash ^= ch;
		hash *= 1099511628211ull;
	}
	return hash;
}

std::string ShortHashHex(const std::string& text)
{
	std::ostringstream ss;
	ss << std::hex << std::uppercase << (Fnv1a64(text) & 0xFFFFFFFFull);
	return ss.str();
}

std::filesystem::path BuildUniqueCachePath(
	const std::filesystem::path& root,
	const std::string& displayNameLocal,
	const std::string& sourcePathLocal,
	const std::string& extension)
{
	const std::string base = SanitizePathComponentForStorage(displayNameLocal.empty() ? "Unnamed" : displayNameLocal);
	return root / (base + "-" + ShortHashHex(sourcePathLocal) + extension);
}

bool IsTextIndexCandidate(const std::filesystem::path& path)
{
	const std::string ext = ToLowerAsciiCopy(path.extension().string());
	return ext == ".txt" || ext == ".json" || ext == ".md" || ext == ".csv" || ext == ".ini" || ext == ".xml";
}

std::string CollectTextFilesUtf8(const std::filesystem::path& root)
{
	std::string merged;
	std::error_code ec;
	if (!std::filesystem::exists(root, ec)) {
		return merged;
	}
	const auto options = std::filesystem::directory_options::skip_permission_denied;
	for (std::filesystem::recursive_directory_iterator it(root, options, ec), end; it != end; it.increment(ec)) {
		if (ec) {
			ec.clear();
			continue;
		}
		const auto& entry = *it;
		if (!entry.is_regular_file(ec) || ec) {
			ec.clear();
			continue;
		}
		if (!IsTextIndexCandidate(entry.path())) {
			continue;
		}
		std::string bytes;
		if (!ReadFileBytesLimited(entry.path(), bytes)) {
			continue;
		}
		merged += "\n# ";
		merged += LocalToUtf8(entry.path().filename().string());
		merged += "\n";
		merged += BytesToUtf8Text(std::move(bytes));
		if (merged.size() >= kMaxIndexedTextBytes) {
			merged.resize(kMaxIndexedTextBytes);
			break;
		}
	}
	return merged;
}

void RemoveDirectoryContents(const std::filesystem::path& dir)
{
	std::error_code ec;
	if (!std::filesystem::exists(dir, ec) || !std::filesystem::is_directory(dir, ec)) {
		return;
	}
	for (std::filesystem::directory_iterator it(dir, ec), end; it != end && !ec; it.increment(ec)) {
		std::filesystem::remove_all(it->path(), ec);
		ec.clear();
	}
}

std::vector<std::filesystem::path> ListFilesByExtension(const std::filesystem::path& root, const char* extension)
{
	std::vector<std::filesystem::path> paths;
	std::error_code ec;
	if (!std::filesystem::exists(root, ec) || !std::filesystem::is_directory(root, ec)) {
		return paths;
	}
	const std::string loweredExt = ToLowerAsciiCopy(extension == nullptr ? std::string() : std::string(extension));
	const auto options = std::filesystem::directory_options::skip_permission_denied;
	for (std::filesystem::recursive_directory_iterator it(root, options, ec), end; it != end; it.increment(ec)) {
		if (ec) {
			ec.clear();
			continue;
		}
		const auto& entry = *it;
		if (!entry.is_regular_file(ec) || ec) {
			ec.clear();
			continue;
		}
		if (ToLowerAsciiCopy(entry.path().extension().string()) == loweredExt) {
			paths.push_back(entry.path());
		}
	}
	std::sort(paths.begin(), paths.end(), [](const auto& left, const auto& right) {
		return ToLowerAsciiCopy(left.string()) < ToLowerAsciiCopy(right.string());
	});
	return paths;
}

std::string SafeString(const char* text)
{
	return text == nullptr ? std::string() : std::string(text);
}

PLIB_INFO SafeCallGetLibInfo(PFN_GET_LIB_INFO fn)
{
	if (fn == nullptr) {
		return nullptr;
	}
	__try {
		return fn();
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return nullptr;
	}
}

std::string BuildLibraryInfoTextLocal(PLIB_INFO info, const std::filesystem::path& path)
{
	std::ostringstream out;
	const std::string name = SafeString(info != nullptr ? info->m_szName : nullptr);
	out << "文件: " << path.string() << "\r\n";
	out << "支持库名: " << (name.empty() ? PathStemLocal(path) : name) << "\r\n";
	if (info == nullptr) {
		out << "解码失败: GetNewInf returned null\r\n";
		return out.str();
	}

	out << "版本: " << info->m_nMajorVersion << "." << info->m_nMinorVersion << "." << info->m_nBuildNumber << "\r\n";
	const std::string explain = SafeString(info->m_szExplain);
	if (!explain.empty()) {
		out << "说明: " << explain << "\r\n";
	}
	const std::string author = SafeString(info->m_szAuthor);
	if (!author.empty()) {
		out << "作者: " << author << "\r\n";
	}
	const std::string homepage = SafeString(info->m_szHomePage);
	if (!homepage.empty()) {
		out << "主页: " << homepage << "\r\n";
	}
	out << "命令数: " << info->m_nCmdCount << "\r\n";
	for (int i = 0; i < info->m_nCmdCount && info->m_pBeginCmdInfo != nullptr && i < 2000; ++i) {
		const CMD_INFO& cmd = info->m_pBeginCmdInfo[i];
		out << "命令[" << i << "]: " << SafeString(cmd.m_szName);
		const std::string egName = SafeString(cmd.m_szEgName);
		if (!egName.empty()) {
			out << " / " << egName;
		}
		const std::string cmdExplain = SafeString(cmd.m_szExplain);
		if (!cmdExplain.empty()) {
			out << " - " << cmdExplain;
		}
		out << "\r\n";
	}

	out << "数据类型数: " << info->m_nDataTypeCount << "\r\n";
	for (int i = 0; i < info->m_nDataTypeCount && info->m_pDataType != nullptr && i < 1000; ++i) {
		const LIB_DATA_TYPE_INFO& type = info->m_pDataType[i];
		out << "数据类型[" << i << "]: " << SafeString(type.m_szName);
		const std::string egName = SafeString(type.m_szEgName);
		if (!egName.empty()) {
			out << " / " << egName;
		}
		const std::string typeExplain = SafeString(type.m_szExplain);
		if (!typeExplain.empty()) {
			out << " - " << typeExplain;
		}
		out << "\r\n";
	}

	out << "常量数: " << info->m_nLibConstCount << "\r\n";
	for (int i = 0; i < info->m_nLibConstCount && info->m_pLibConst != nullptr && i < 1000; ++i) {
		const LIB_CONST_INFO& constant = info->m_pLibConst[i];
		out << "常量[" << i << "]: " << SafeString(constant.m_szName);
		const std::string constExplain = SafeString(constant.m_szExplain);
		if (!constExplain.empty()) {
			out << " - " << constExplain;
		}
		out << "\r\n";
	}
	return out.str();
}

bool DecodeLibraryInfo(
	const std::filesystem::path& path,
	std::string& outNameLocal,
	std::string& outInfoLocal,
	std::string& outErrorLocal)
{
	outNameLocal = PathStemLocal(path);
	outInfoLocal.clear();
	outErrorLocal.clear();

	HMODULE module = LoadLibraryExW(path.wstring().c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
	if (module == nullptr) {
		outErrorLocal = "LoadLibraryEx failed: " + std::to_string(GetLastError());
		return false;
	}

	auto fn = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
	if (fn == nullptr) {
		fn = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, "GetLibInf"));
	}
	PLIB_INFO info = SafeCallGetLibInfo(fn);
	if (info == nullptr) {
		outErrorLocal = "GetNewInf failed";
		outInfoLocal = BuildLibraryInfoTextLocal(nullptr, path);
		FreeLibrary(module);
		return false;
	}

	const std::string name = SafeString(info->m_szName);
	if (!TrimAsciiCopy(name).empty()) {
		outNameLocal = name;
	}
	outInfoLocal = BuildLibraryInfoTextLocal(info, path);
	FreeLibrary(module);
	return true;
}

int ClampLimit(int limit)
{
	if (limit <= 0) {
		return kDefaultSearchLimit;
	}
	return (std::min)(limit, kMaxSearchLimit);
}

std::string NormalizeSearchText(std::string text)
{
	std::replace(text.begin(), text.end(), '\r', '\n');
	while (text.find("\n\n") != std::string::npos) {
		text.replace(text.find("\n\n"), 2, "\n");
	}
	return text;
}

bool ContainsUtf8Search(const std::string& haystackUtf8, const std::string& needleUtf8)
{
	if (needleUtf8.empty()) {
		return true;
	}
	if (haystackUtf8.find(needleUtf8) != std::string::npos) {
		return true;
	}
	return ToLowerAsciiCopy(haystackUtf8).find(ToLowerAsciiCopy(needleUtf8)) != std::string::npos;
}

std::string BuildSnippetUtf8(const std::string& haystackUtf8, const std::string& needleUtf8)
{
	if (haystackUtf8.empty()) {
		return {};
	}
	size_t pos = needleUtf8.empty() ? 0 : haystackUtf8.find(needleUtf8);
	if (pos == std::string::npos) {
		const std::string lowerHaystack = ToLowerAsciiCopy(haystackUtf8);
		const std::string lowerNeedle = ToLowerAsciiCopy(needleUtf8);
		pos = lowerNeedle.empty() ? 0 : lowerHaystack.find(lowerNeedle);
	}
	if (pos == std::string::npos) {
		pos = 0;
	}
	auto isContinuationByte = [](unsigned char ch) {
		return (ch & 0xC0) == 0x80;
	};
	auto adjustBegin = [&](size_t value) {
		while (value < haystackUtf8.size() && isContinuationByte(static_cast<unsigned char>(haystackUtf8[value]))) {
			++value;
		}
		return value;
	};
	auto adjustEnd = [&](size_t value) {
		if (value >= haystackUtf8.size()) {
			return haystackUtf8.size();
		}
		while (value > 0 && isContinuationByte(static_cast<unsigned char>(haystackUtf8[value]))) {
			--value;
		}
		return value;
	};
	const size_t begin = adjustBegin(pos > 120 ? pos - 120 : 0);
	const size_t end = adjustEnd((std::min)(haystackUtf8.size(), pos + (std::max<size_t>)(needleUtf8.size(), 1) + 240));
	std::string snippet = haystackUtf8.substr(begin, end - begin);
	if (begin > 0) {
		snippet.insert(0, "...");
	}
	if (end < haystackUtf8.size()) {
		snippet += "...";
	}
	return NormalizeSearchText(std::move(snippet));
}

int ScoreMatch(const std::string& nameUtf8, const std::string& fileUtf8, const std::string& pathUtf8, const std::string& textUtf8, const std::string& queryUtf8)
{
	if (queryUtf8.empty()) {
		return 1;
	}
	const std::string lowerQuery = ToLowerAsciiCopy(queryUtf8);
	if (ToLowerAsciiCopy(nameUtf8) == lowerQuery || ToLowerAsciiCopy(fileUtf8) == lowerQuery) {
		return 1000;
	}
	if (ContainsUtf8Search(nameUtf8, queryUtf8)) {
		return 800;
	}
	if (ContainsUtf8Search(fileUtf8, queryUtf8)) {
		return 700;
	}
	if (ContainsUtf8Search(pathUtf8, queryUtf8)) {
		return 500;
	}
	if (ContainsUtf8Search(textUtf8, queryUtf8)) {
		return 200;
	}
	return 0;
}

bool SamePathLoose(const std::string& left, const std::string& right)
{
	if (left.empty() || right.empty()) {
		return false;
	}
	try {
		return EqualsInsensitiveAscii(
			std::filesystem::path(left).lexically_normal().string(),
			std::filesystem::path(right).lexically_normal().string());
	}
	catch (...) {
		return EqualsInsensitiveAscii(left, right);
	}
}

bool IsPathInsideOrSame(const std::filesystem::path& path, const std::filesystem::path& root)
{
	try {
		const std::filesystem::path normalizedPath = path.lexically_normal();
		const std::filesystem::path normalizedRoot = root.lexically_normal();
		auto pit = normalizedPath.begin();
		auto rit = normalizedRoot.begin();
		for (; rit != normalizedRoot.end(); ++rit, ++pit) {
			if (pit == normalizedPath.end()) {
				return false;
			}
			if (!EqualsInsensitiveAscii((*pit).string(), (*rit).string())) {
				return false;
			}
		}
		return true;
	}
	catch (...) {
		return false;
	}
}

std::string RefreshStateToLogText(DependencyCatalogCache::RefreshState state)
{
	switch (state) {
	case DependencyCatalogCache::RefreshState::Idle:
		return "idle";
	case DependencyCatalogCache::RefreshState::Running:
		return "running";
	case DependencyCatalogCache::RefreshState::Ready:
		return "ready";
	case DependencyCatalogCache::RefreshState::Failed:
		return "failed";
	default:
		return "unknown";
	}
}

} // namespace

DependencyCatalogCache& DependencyCatalogCache::Instance()
{
	static DependencyCatalogCache instance;
	return instance;
}

void DependencyCatalogCache::StartAsyncRefreshIfNeeded(bool force)
{
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (m_workerRunning || (!force && m_hasSnapshot)) {
			return;
		}
		m_workerRunning = true;
		m_state = RefreshState::Running;
		m_lastErrorLocal.clear();
		++m_refreshGeneration;
	}

	std::thread([this, force]() {
		RefreshWorker(force);
	}).detach();
}

bool DependencyCatalogCache::RefreshNow(bool force, std::string& outErrorLocal)
{
	{
		std::unique_lock<std::mutex> lock(m_mutex);
		if (m_workerRunning) {
			m_cv.wait(lock, [this]() { return !m_workerRunning; });
		}
		m_workerRunning = true;
		m_state = RefreshState::Running;
		m_lastErrorLocal.clear();
		++m_refreshGeneration;
	}

	RefreshWorker(force);
	const Status status = GetStatus();
	outErrorLocal = status.lastErrorLocal;
	return status.state == RefreshState::Ready;
}

bool DependencyCatalogCache::WaitForIdle(unsigned int timeoutMs)
{
	std::unique_lock<std::mutex> lock(m_mutex);
	if (!m_workerRunning) {
		return true;
	}
	if (timeoutMs == 0) {
		m_cv.wait(lock, [this]() { return !m_workerRunning; });
		return true;
	}
	return m_cv.wait_for(lock, std::chrono::milliseconds(timeoutMs), [this]() { return !m_workerRunning; });
}

DependencyCatalogCache::Status DependencyCatalogCache::GetStatus() const
{
	std::lock_guard<std::mutex> lock(m_mutex);
	Status status;
	status.state = m_state;
	status.hasSnapshot = m_hasSnapshot;
	status.running = m_workerRunning;
	status.refreshGeneration = m_refreshGeneration;
	status.lastDurationMs = m_lastDurationMs;
	status.moduleCount = m_modules.size();
	status.libraryCount = m_libraries.size();
	status.decodedModuleCount = m_decodedModuleCount;
	status.decodedLibraryCount = m_decodedLibraryCount;
	status.cacheRootLocal = m_cacheRootLocal;
	status.ecomRootLocal = m_ecomRootLocal;
	status.libRootLocal = m_libRootLocal;
	status.lastErrorLocal = m_lastErrorLocal;
	return status;
}

std::vector<DependencyCatalogCache::ModuleEntry> DependencyCatalogCache::SnapshotModules() const
{
	std::lock_guard<std::mutex> lock(m_mutex);
	return m_modules;
}

std::vector<DependencyCatalogCache::LibraryEntry> DependencyCatalogCache::SnapshotLibraries() const
{
	std::lock_guard<std::mutex> lock(m_mutex);
	return m_libraries;
}

std::vector<DependencyCatalogCache::ModuleSearchResult> DependencyCatalogCache::SearchModules(const SearchOptions& options)
{
	StartAsyncRefreshIfNeeded(false);
	const std::vector<ModuleEntry> entries = SnapshotModules();
	const std::string query = TrimAsciiCopy(options.queryUtf8);
	std::vector<ModuleSearchResult> results;
	for (const auto& entry : entries) {
		const std::string nameUtf8 = LocalToUtf8(entry.moduleNameLocal);
		const std::string fileUtf8 = LocalToUtf8(entry.fileNameLocal);
		const std::string pathUtf8 = LocalToUtf8(entry.pathLocal);
		const int score = ScoreMatch(nameUtf8, fileUtf8, pathUtf8, entry.searchTextUtf8, query);
		if (score <= 0) {
			continue;
		}
		ModuleSearchResult row;
		row.entry = entry;
		row.score = score;
		if (options.includeSnippets) {
			row.snippetUtf8 = BuildSnippetUtf8(entry.searchTextUtf8.empty() ? pathUtf8 : entry.searchTextUtf8, query);
		}
		results.push_back(std::move(row));
	}
	std::sort(results.begin(), results.end(), [](const auto& left, const auto& right) {
		if (left.score != right.score) {
			return left.score > right.score;
		}
		return ToLowerAsciiCopy(left.entry.pathLocal) < ToLowerAsciiCopy(right.entry.pathLocal);
	});
	const int limit = ClampLimit(options.limit);
	if (results.size() > static_cast<size_t>(limit)) {
		results.resize(static_cast<size_t>(limit));
	}
	return results;
}

std::vector<DependencyCatalogCache::LibrarySearchResult> DependencyCatalogCache::SearchLibraries(const SearchOptions& options)
{
	StartAsyncRefreshIfNeeded(false);
	const std::vector<LibraryEntry> entries = SnapshotLibraries();
	const std::string query = TrimAsciiCopy(options.queryUtf8);
	std::vector<LibrarySearchResult> results;
	for (const auto& entry : entries) {
		const std::string nameUtf8 = LocalToUtf8(entry.libraryNameLocal);
		const std::string fileUtf8 = LocalToUtf8(entry.fileNameLocal);
		const std::string pathUtf8 = LocalToUtf8(entry.pathLocal);
		const int score = ScoreMatch(nameUtf8, fileUtf8, pathUtf8, entry.searchTextUtf8, query);
		if (score <= 0) {
			continue;
		}
		LibrarySearchResult row;
		row.entry = entry;
		row.score = score;
		if (options.includeSnippets) {
			row.snippetUtf8 = BuildSnippetUtf8(entry.searchTextUtf8.empty() ? pathUtf8 : entry.searchTextUtf8, query);
		}
		results.push_back(std::move(row));
	}
	std::sort(results.begin(), results.end(), [](const auto& left, const auto& right) {
		if (left.score != right.score) {
			return left.score > right.score;
		}
		return ToLowerAsciiCopy(left.entry.pathLocal) < ToLowerAsciiCopy(right.entry.pathLocal);
	});
	const int limit = ClampLimit(options.limit);
	if (results.size() > static_cast<size_t>(limit)) {
		results.resize(static_cast<size_t>(limit));
	}
	return results;
}

bool DependencyCatalogCache::ResolveModulePath(
	const std::string& moduleNameLocal,
	const std::string& modulePathLocal,
	std::string& outPathLocal,
	std::vector<ModuleEntry>* outCandidates,
	std::string& outErrorLocal) const
{
	outPathLocal.clear();
	outErrorLocal.clear();
	if (outCandidates != nullptr) {
		outCandidates->clear();
	}

	const Status status = GetStatus();
	const std::filesystem::path ecomRoot(status.ecomRootLocal.empty() ? (std::filesystem::path(GetBasePath()) / "ecom") : std::filesystem::path(status.ecomRootLocal));
	const std::string trimmedPath = TrimAsciiCopy(modulePathLocal);
	if (!trimmedPath.empty()) {
		std::filesystem::path candidate(trimmedPath);
		if (candidate.is_relative()) {
			candidate = ecomRoot / candidate;
		}
		std::error_code ec;
		if (std::filesystem::exists(candidate, ec) &&
			std::filesystem::is_regular_file(candidate, ec) &&
			EqualsInsensitiveAscii(candidate.extension().string(), ".ec")) {
			outPathLocal = candidate.lexically_normal().string();
			return true;
		}
		outErrorLocal = "module file not found";
		return false;
	}

	const std::string name = TrimAsciiCopy(moduleNameLocal);
	if (name.empty()) {
		outErrorLocal = "module_name or module_path is required";
		return false;
	}

	const std::string loweredName = ToLowerAsciiCopy(name);
	std::vector<ModuleEntry> exactMatched;
	std::vector<ModuleEntry> fuzzyMatched;
	for (const auto& entry : SnapshotModules()) {
		const std::string stem = entry.moduleNameLocal;
		const std::string file = entry.fileNameLocal;
		if (EqualsInsensitiveAscii(stem, name) ||
			EqualsInsensitiveAscii(file, name) ||
			EqualsInsensitiveAscii(entry.pathLocal, name)) {
			exactMatched.push_back(entry);
			continue;
		}
		if (ToLowerAsciiCopy(stem).find(loweredName) != std::string::npos ||
			ToLowerAsciiCopy(file).find(loweredName) != std::string::npos) {
			fuzzyMatched.push_back(entry);
		}
	}
	const std::vector<ModuleEntry>& matched = exactMatched.empty() ? fuzzyMatched : exactMatched;
	if (outCandidates != nullptr) {
		*outCandidates = matched;
	}
	if (matched.empty()) {
		outErrorLocal = "module not found in dependency catalog";
		return false;
	}
	if (matched.size() > 1) {
		outErrorLocal = "module name is ambiguous";
		return false;
	}
	outPathLocal = matched.front().pathLocal;
	return true;
}

bool DependencyCatalogCache::ResolveLibrary(
	const std::string& libraryNameLocal,
	const std::string& libraryPathLocal,
	LibraryEntry& outEntry,
	std::vector<LibraryEntry>* outCandidates,
	std::string& outErrorLocal) const
{
	outEntry = {};
	outErrorLocal.clear();
	if (outCandidates != nullptr) {
		outCandidates->clear();
	}

	const Status status = GetStatus();
	const std::filesystem::path libRoot(status.libRootLocal.empty() ? (std::filesystem::path(GetBasePath()) / "lib") : std::filesystem::path(status.libRootLocal));
	const std::string trimmedPath = TrimAsciiCopy(libraryPathLocal);
	if (!trimmedPath.empty()) {
		std::filesystem::path candidate(trimmedPath);
		if (candidate.is_relative()) {
			candidate = libRoot / candidate;
		}
		std::error_code ec;
		if (!std::filesystem::exists(candidate, ec) || !std::filesystem::is_regular_file(candidate, ec)) {
			outErrorLocal = "support library file not found";
			return false;
		}
		const std::string ext = ToLowerAsciiCopy(candidate.extension().string());
		if (ext != ".fne" && ext != ".fnr") {
			outErrorLocal = "support library file must be .fne or .fnr";
			return false;
		}
		for (const auto& entry : SnapshotLibraries()) {
			if (SamePathLoose(entry.pathLocal, candidate.string())) {
				outEntry = entry;
				return true;
			}
		}
		outEntry.fileNameLocal = PathFileNameLocal(candidate);
		outEntry.libraryNameLocal = PathStemLocal(candidate);
		outEntry.pathLocal = candidate.lexically_normal().string();
		return true;
	}

	const std::string name = TrimAsciiCopy(libraryNameLocal);
	if (name.empty()) {
		outErrorLocal = "library_name or library_path is required";
		return false;
	}

	const std::string loweredName = ToLowerAsciiCopy(name);
	std::vector<LibraryEntry> exactMatched;
	std::vector<LibraryEntry> fuzzyMatched;
	for (const auto& entry : SnapshotLibraries()) {
		if (EqualsInsensitiveAscii(entry.libraryNameLocal, name) ||
			EqualsInsensitiveAscii(entry.fileNameLocal, name) ||
			EqualsInsensitiveAscii(entry.pathLocal, name)) {
			exactMatched.push_back(entry);
			continue;
		}
		if (ToLowerAsciiCopy(entry.libraryNameLocal).find(loweredName) != std::string::npos ||
			ToLowerAsciiCopy(entry.fileNameLocal).find(loweredName) != std::string::npos) {
			fuzzyMatched.push_back(entry);
		}
	}
	const std::vector<LibraryEntry>& matched = exactMatched.empty() ? fuzzyMatched : exactMatched;
	if (outCandidates != nullptr) {
		*outCandidates = matched;
	}
	if (matched.empty()) {
		outErrorLocal = "support library not found in dependency catalog";
		return false;
	}
	if (matched.size() > 1) {
		outErrorLocal = "support library name is ambiguous";
		return false;
	}
	outEntry = matched.front();
	return true;
}

void DependencyCatalogCache::RefreshWorker(bool force)
{
	std::string error;
	const bool ok = RefreshInternal(force, error);
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_state = ok ? RefreshState::Ready : RefreshState::Failed;
		m_lastErrorLocal = error;
		m_workerRunning = false;
	}
	m_cv.notify_all();
}

bool DependencyCatalogCache::RefreshInternal(bool force, std::string& outErrorLocal)
{
	outErrorLocal.clear();
	const auto start = Clock::now();
	auto elapsedMs = [&]() -> long long {
		return static_cast<long long>(std::chrono::duration_cast<std::chrono::milliseconds>(Clock::now() - start).count());
	};
	const std::filesystem::path basePath(GetBasePath());
	const std::filesystem::path ecomRoot = basePath / "ecom";
	const std::filesystem::path libRoot = basePath / "lib";
	const std::filesystem::path cacheRoot = GetAutoLinkerDirectoryPath() / "Cache";
	const std::filesystem::path ecomCacheRoot = cacheRoot / "EcomInfo";
	const std::filesystem::path libCacheRoot = cacheRoot / "LibInfo";

	Logger::Instance().WriteAndIde("DependencyCatalog", "开始刷新模块/支持库缓存");

	std::error_code ec;
	std::filesystem::create_directories(ecomCacheRoot, ec);
	if (ec) {
		outErrorLocal = "create EcomInfo cache directory failed: " + ec.message();
		return false;
	}
	std::filesystem::create_directories(libCacheRoot, ec);
	if (ec) {
		outErrorLocal = "create LibInfo cache directory failed: " + ec.message();
		return false;
	}
	if (force) {
		RemoveDirectoryContents(ecomCacheRoot);
		RemoveDirectoryContents(libCacheRoot);
	}

	const std::vector<std::filesystem::path> modulePaths = ListFilesByExtension(ecomRoot, ".ec");
	const std::vector<std::filesystem::path> libraryPaths = ListFilesByExtension(libRoot, ".fne");

	std::vector<ModuleEntry> modules;
	modules.reserve(modulePaths.size());
	for (const std::filesystem::path& path : modulePaths) {
		ModuleEntry entry;
		entry.fileNameLocal = PathFileNameLocal(path);
		entry.moduleNameLocal = PathStemLocal(path);
		entry.pathLocal = path.lexically_normal().string();
		entry.cachePathLocal = BuildUniqueCachePath(ecomCacheRoot, entry.moduleNameLocal, entry.pathLocal, "").string();
		entry.searchTextUtf8 =
			LocalToUtf8(entry.moduleNameLocal) + "\n" +
			LocalToUtf8(entry.fileNameLocal) + "\n" +
			LocalToUtf8(entry.pathLocal);
		modules.push_back(std::move(entry));
	}

	std::vector<LibraryEntry> libraries;
	libraries.reserve(libraryPaths.size());
	for (const std::filesystem::path& path : libraryPaths) {
		LibraryEntry entry;
		entry.fileNameLocal = PathFileNameLocal(path);
		entry.libraryNameLocal = PathStemLocal(path);
		entry.pathLocal = path.lexically_normal().string();
		entry.cachePathLocal = BuildUniqueCachePath(libCacheRoot, entry.libraryNameLocal, entry.pathLocal, ".txt").string();
		entry.searchTextUtf8 =
			LocalToUtf8(entry.libraryNameLocal) + "\n" +
			LocalToUtf8(entry.fileNameLocal) + "\n" +
			LocalToUtf8(entry.pathLocal);
		libraries.push_back(std::move(entry));
	}

	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_modules = modules;
		m_libraries = libraries;
		m_decodedModuleCount = 0;
		m_decodedLibraryCount = 0;
		m_lastDurationMs = elapsedMs();
		m_cacheRootLocal = cacheRoot.string();
		m_ecomRootLocal = ecomRoot.string();
		m_libRootLocal = libRoot.string();
		m_hasSnapshot = true;
	}

	Logger::Instance().WriteAndIde(
		"DependencyCatalog",
		"已发布初始依赖快照 modules=" + std::to_string(modules.size()) +
			" libs=" + std::to_string(libraries.size()) +
			" elapsed_ms=" + std::to_string(elapsedMs()));

	size_t decodedModules = 0;
	size_t decodedLibraries = 0;

	for (const std::filesystem::path& path : libraryPaths) {
		LibraryEntry entry;
		entry.fileNameLocal = PathFileNameLocal(path);
		entry.libraryNameLocal = PathStemLocal(path);
		entry.pathLocal = path.lexically_normal().string();

		std::string infoLocal;
		std::string decodedName;
		std::string decodeError;
		if (DecodeLibraryInfo(path, decodedName, infoLocal, decodeError)) {
			entry.decoded = true;
			entry.libraryNameLocal = decodedName.empty() ? entry.libraryNameLocal : decodedName;
			entry.infoTextLocal = infoLocal;
			++decodedLibraries;
		}
		else {
			entry.decodeErrorLocal = decodeError;
			entry.infoTextLocal = infoLocal.empty()
				? ("文件: " + entry.pathLocal + "\r\n支持库名: " + entry.libraryNameLocal + "\r\n解码失败: " + decodeError + "\r\n")
				: infoLocal;
		}

		const std::filesystem::path cachePath = BuildUniqueCachePath(libCacheRoot, entry.libraryNameLocal, entry.pathLocal, ".txt");
		entry.cachePathLocal = cachePath.string();
		std::string writeError;
		if (!WriteTextUtf8Bom(cachePath, LocalToUtf8(entry.infoTextLocal), writeError) && entry.decodeErrorLocal.empty()) {
			entry.decodeErrorLocal = "write cache failed: " + writeError;
		}
		entry.searchTextUtf8 =
			LocalToUtf8(entry.libraryNameLocal) + "\n" +
			LocalToUtf8(entry.fileNameLocal) + "\n" +
			LocalToUtf8(entry.pathLocal) + "\n" +
			LocalToUtf8(entry.infoTextLocal);

		{
			std::lock_guard<std::mutex> lock(m_mutex);
			bool replaced = false;
			for (LibraryEntry& existing : m_libraries) {
				if (SamePathLoose(existing.pathLocal, entry.pathLocal)) {
					existing = entry;
					replaced = true;
					break;
				}
			}
			if (!replaced) {
				m_libraries.push_back(entry);
			}
			m_decodedLibraryCount = decodedLibraries;
			m_lastDurationMs = elapsedMs();
		}
	}

	std::filesystem::path toolPath;
	std::string toolError;
	const bool toolReady = EPackagerIntegration::EnsureToolReady(toolPath, toolError);

	for (const std::filesystem::path& path : modulePaths) {
		ModuleEntry entry;
		entry.fileNameLocal = PathFileNameLocal(path);
		entry.moduleNameLocal = PathStemLocal(path);
		entry.pathLocal = path.lexically_normal().string();
		const std::filesystem::path cachePath = BuildUniqueCachePath(ecomCacheRoot, entry.moduleNameLocal, entry.pathLocal, "");
		entry.cachePathLocal = cachePath.string();

		if (!toolReady) {
			entry.decodeErrorLocal = toolError.empty() ? "e-packager unavailable" : toolError;
		}
		else {
			std::filesystem::create_directories(cachePath, ec);
			ec.clear();
			const EPackagerIntegration::ProcessRunResult result = EPackagerIntegration::RunProcessAndCapture(
				toolPath,
				{ L"unpack", path.wstring(), cachePath.wstring(), L"--main-only" },
				toolPath.parent_path());
			if (result.ok) {
				entry.searchTextUtf8 = CollectTextFilesUtf8(cachePath);
				entry.decoded = true;
				++decodedModules;
			}
			else {
				entry.decodeErrorLocal = "e-packager unpack failed";
				if (!result.error.empty()) {
					entry.decodeErrorLocal += ": " + result.error;
				}
				if (!result.stdErrBytes.empty()) {
					entry.decodeErrorLocal += " stderr=" + BytesToUtf8Text(result.stdErrBytes);
				}
			}
		}
		if (entry.searchTextUtf8.empty()) {
			entry.searchTextUtf8 =
				LocalToUtf8(entry.moduleNameLocal) + "\n" +
				LocalToUtf8(entry.fileNameLocal) + "\n" +
				LocalToUtf8(entry.pathLocal);
		}

		{
			std::lock_guard<std::mutex> lock(m_mutex);
			bool replaced = false;
			for (ModuleEntry& existing : m_modules) {
				if (SamePathLoose(existing.pathLocal, entry.pathLocal)) {
					existing = entry;
					replaced = true;
					break;
				}
			}
			if (!replaced) {
				m_modules.push_back(entry);
			}
			m_decodedModuleCount = decodedModules;
			m_lastDurationMs = elapsedMs();
		}
	}

	const long long durationMs = elapsedMs();
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_decodedModuleCount = decodedModules;
		m_decodedLibraryCount = decodedLibraries;
		m_lastDurationMs = durationMs;
		m_cacheRootLocal = cacheRoot.string();
		m_ecomRootLocal = ecomRoot.string();
		m_libRootLocal = libRoot.string();
		m_hasSnapshot = true;
	}

	Logger::Instance().WriteAndIde(
		"DependencyCatalog",
		"刷新完成 state=" + RefreshStateToLogText(RefreshState::Ready) +
			" modules=" + std::to_string(SnapshotModules().size()) +
			" libs=" + std::to_string(SnapshotLibraries().size()) +
			" decoded_modules=" + std::to_string(decodedModules) +
			" decoded_libs=" + std::to_string(decodedLibraries) +
			" elapsed_ms=" + std::to_string(durationMs));
	return true;
}
