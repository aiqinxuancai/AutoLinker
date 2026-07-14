#include "LocalMcpInstanceRegistry.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <mutex>
#include <string>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <vector>

#include "..\\thirdparty\\json.hpp"

namespace {

constexpr const char* kLegacyRegistryFileName = "local_mcp_instances.json";
constexpr const char* kRegistryDirectoryName = "local_mcp_instances";
constexpr std::uint64_t kInstanceTtlMs = 15000;
constexpr std::uint64_t kDeadInstanceCleanupTtlMs = 60000;
constexpr std::uint64_t kHardArtifactCleanupTtlMs = 24ULL * 60ULL * 60ULL * 1000ULL;
constexpr auto kLocalOperationLockTimeout = std::chrono::milliseconds(250);

std::timed_mutex g_currentInstanceMutex;
std::atomic_ullong g_tempFileCounter = 1;

void SetError(std::string* outError, const std::string& message)
{
	if (outError != nullptr) {
		*outError = message;
	}
}

std::uint64_t GetUnixTimeMilliseconds()
{
	return static_cast<std::uint64_t>(
		std::chrono::duration_cast<std::chrono::milliseconds>(
			std::chrono::system_clock::now().time_since_epoch()).count());
}

std::filesystem::path GetRegistryBaseDirectoryObject()
{
	std::error_code ec;
	std::filesystem::path dir = std::filesystem::temp_directory_path(ec);
	if (ec || dir.empty()) {
		dir = std::filesystem::current_path(ec);
	}
	if (dir.empty()) {
		dir = ".";
	}
	dir /= "AutoLinker";
	std::filesystem::create_directories(dir, ec);
	return dir;
}

std::filesystem::path GetLegacyRegistryPathObject(const std::filesystem::path& baseDirectory)
{
	return baseDirectory / kLegacyRegistryFileName;
}

std::filesystem::path GetRegistryDirectoryPathObject(const std::filesystem::path& baseDirectory)
{
	return baseDirectory / kRegistryDirectoryName;
}

bool IsSafeInstanceId(const std::string& instanceId)
{
	if (instanceId.empty() || instanceId.size() > 96 || instanceId == "." || instanceId == "..") {
		return false;
	}
	for (const unsigned char ch : instanceId) {
		const bool alphaNumeric =
			(ch >= 'a' && ch <= 'z') ||
			(ch >= 'A' && ch <= 'Z') ||
			(ch >= '0' && ch <= '9');
		if (!alphaNumeric && ch != '-' && ch != '_' && ch != '.') {
			return false;
		}
	}
	return true;
}

std::filesystem::path GetInstanceFilePath(
	const std::filesystem::path& baseDirectory,
	const std::string& instanceId)
{
	return GetRegistryDirectoryPathObject(baseDirectory) / (instanceId + ".json");
}

bool IsProcessLikelyAlive(unsigned long processId)
{
	if (processId == 0) {
		return false;
	}

	HANDLE process = OpenProcess(SYNCHRONIZE, FALSE, processId);
	if (process == nullptr) {
		return false;
	}

	const DWORD waitResult = WaitForSingleObject(process, 0);
	CloseHandle(process);
	return waitResult == WAIT_TIMEOUT;
}

bool IsValidUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return text;
	}

	const int outLen = WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		out.data(),
		outLen,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return out;
}

std::string LocalToUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocalText(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

nlohmann::json BuildRecordJson(const LocalMcpInstanceRegistry::InstanceRecord& record)
{
	return {
		{"instance_id", LocalToUtf8Text(record.instanceId)},
		{"process_id", record.processId},
		{"process_path", LocalToUtf8Text(record.processPath)},
		{"process_name", LocalToUtf8Text(record.processName)},
		{"port", record.port},
		{"endpoint", LocalToUtf8Text(record.endpoint)},
		{"source_file_path_hint", LocalToUtf8Text(record.sourceFilePathHint)},
		{"page_name_hint", LocalToUtf8Text(record.pageNameHint)},
		{"page_type_hint", LocalToUtf8Text(record.pageTypeHint)},
		{"last_seen_unix_ms", record.lastSeenUnixMs}
	};
}

LocalMcpInstanceRegistry::InstanceRecord ParseRecordJson(const nlohmann::json& value)
{
	LocalMcpInstanceRegistry::InstanceRecord record;
	if (!value.is_object()) {
		return record;
	}

	record.instanceId = Utf8ToLocalText(value.value("instance_id", std::string()));
	record.processId = value.value("process_id", 0UL);
	record.processPath = Utf8ToLocalText(value.value("process_path", std::string()));
	record.processName = Utf8ToLocalText(value.value("process_name", std::string()));
	record.port = value.value("port", 0);
	record.endpoint = Utf8ToLocalText(value.value("endpoint", std::string()));
	record.sourceFilePathHint = Utf8ToLocalText(value.value("source_file_path_hint", std::string()));
	record.pageNameHint = Utf8ToLocalText(value.value("page_name_hint", std::string()));
	record.pageTypeHint = Utf8ToLocalText(value.value("page_type_hint", std::string()));
	record.lastSeenUnixMs = value.value("last_seen_unix_ms", static_cast<std::uint64_t>(0));
	return record;
}

bool IsRecordValid(const LocalMcpInstanceRegistry::InstanceRecord& record)
{
	return IsSafeInstanceId(record.instanceId) &&
		record.port > 0 &&
		!record.endpoint.empty() &&
		record.lastSeenUnixMs > 0;
}

std::uint64_t GetRecordAgeMs(
	const LocalMcpInstanceRegistry::InstanceRecord& record,
	std::uint64_t nowMs)
{
	return nowMs >= record.lastSeenUnixMs ? nowMs - record.lastSeenUnixMs : 0;
}

bool IsRecordFresh(
	const LocalMcpInstanceRegistry::InstanceRecord& record,
	std::uint64_t nowMs)
{
	return IsRecordValid(record) && GetRecordAgeMs(record, nowMs) <= kInstanceTtlMs;
}

bool TryLoadJsonFile(
	const std::filesystem::path& path,
	nlohmann::json& outRoot,
	std::string* outError)
{
	outRoot = nlohmann::json::object();
	if (outError != nullptr) {
		outError->clear();
	}

	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		SetError(outError, "open registry file failed");
		return false;
	}

	std::string text((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		text.erase(0, 3);
	}

	try {
		outRoot = text.empty() ? nlohmann::json::object() : nlohmann::json::parse(text);
		if (!outRoot.is_object()) {
			SetError(outError, "registry json root is not an object");
			return false;
		}
	}
	catch (const std::exception& ex) {
		SetError(outError, std::string("parse registry json failed: ") + ex.what());
		return false;
	}
	return true;
}

std::filesystem::path BuildUniqueTempPath(const std::filesystem::path& finalPath)
{
	std::wstring tempPath = finalPath.wstring();
	tempPath += L".";
	tempPath += std::to_wstring(GetCurrentProcessId());
	tempPath += L".";
	tempPath += std::to_wstring(GetCurrentThreadId());
	tempPath += L".";
	tempPath += std::to_wstring(g_tempFileCounter.fetch_add(1));
	tempPath += L".tmp";
	return std::filesystem::path(std::move(tempPath));
}

bool SaveJsonFileAtomically(
	const std::filesystem::path& path,
	const nlohmann::json& root,
	std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}

	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);
	if (ec) {
		SetError(outError, "create registry directory failed, error=" + std::to_string(ec.value()));
		return false;
	}

	const std::filesystem::path tempPath = BuildUniqueTempPath(path);
	std::ofstream out(tempPath, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		SetError(outError, "open registry temp file failed");
		return false;
	}

	static constexpr unsigned char kUtf8Bom[] = { 0xEF, 0xBB, 0xBF };
	out.write(reinterpret_cast<const char*>(kUtf8Bom), sizeof(kUtf8Bom));
	const std::string text = root.dump(2);
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	out.close();
	if (!out) {
		SetError(outError, "write registry temp file failed");
		std::filesystem::remove(tempPath, ec);
		return false;
	}

	if (MoveFileExW(
		tempPath.wstring().c_str(),
		path.wstring().c_str(),
		MOVEFILE_REPLACE_EXISTING) == FALSE) {
		const DWORD error = GetLastError();
		SetError(outError, "replace registry file failed, win32_error=" + std::to_string(error));
		std::filesystem::remove(tempPath, ec);
		return false;
	}
	return true;
}

nlohmann::json BuildInstanceFileJson(const LocalMcpInstanceRegistry::InstanceRecord& record)
{
	return {
		{"version", 2},
		{"instance", BuildRecordJson(record)}
	};
}

LocalMcpInstanceRegistry::InstanceRecord ParseInstanceFileJson(const nlohmann::json& root)
{
	if (root.contains("instance")) {
		return ParseRecordJson(root["instance"]);
	}
	return ParseRecordJson(root);
}

bool UpsertInstanceAtBaseDirectory(
	const std::filesystem::path& baseDirectory,
	const LocalMcpInstanceRegistry::InstanceRecord& record,
	std::string* outError)
{
	if (!IsSafeInstanceId(record.instanceId)) {
		SetError(outError, "instance_id is invalid");
		return false;
	}
	if (record.port <= 0) {
		SetError(outError, "port is invalid");
		return false;
	}
	if (record.endpoint.empty()) {
		SetError(outError, "endpoint is empty");
		return false;
	}
	if (record.lastSeenUnixMs == 0) {
		SetError(outError, "last_seen_unix_ms is invalid");
		return false;
	}

	return SaveJsonFileAtomically(
		GetInstanceFilePath(baseDirectory, record.instanceId),
		BuildInstanceFileJson(record),
		outError);
}

bool RemoveInstanceAtBaseDirectory(
	const std::filesystem::path& baseDirectory,
	const std::string& instanceId,
	std::string* outError)
{
	if (!IsSafeInstanceId(instanceId)) {
		SetError(outError, "instance_id is invalid");
		return false;
	}

	std::error_code ec;
	std::filesystem::remove(GetInstanceFilePath(baseDirectory, instanceId), ec);
	if (ec) {
		SetError(outError, "remove instance registry file failed, error=" + std::to_string(ec.value()));
		return false;
	}
	return true;
}

void MergeRecord(
	std::unordered_map<std::string, LocalMcpInstanceRegistry::InstanceRecord>& records,
	LocalMcpInstanceRegistry::InstanceRecord record,
	std::uint64_t nowMs)
{
	if (!IsRecordFresh(record, nowMs)) {
		return;
	}
	auto current = records.find(record.instanceId);
	if (current == records.end() || current->second.lastSeenUnixMs < record.lastSeenUnixMs) {
		records[record.instanceId] = std::move(record);
	}
}

bool IsArtifactOlderThan(const std::filesystem::path& path, std::chrono::milliseconds age)
{
	std::error_code ec;
	const auto writeTime = std::filesystem::last_write_time(path, ec);
	if (ec) {
		return false;
	}
	return std::filesystem::file_time_type::clock::now() - writeTime > age;
}

void RemovePathsBestEffort(const std::vector<std::filesystem::path>& paths)
{
	std::error_code ec;
	for (const auto& path : paths) {
		std::filesystem::remove(path, ec);
		ec.clear();
	}
}

void LoadLegacyRecords(
	const std::filesystem::path& baseDirectory,
	std::unordered_map<std::string, LocalMcpInstanceRegistry::InstanceRecord>& records,
	std::uint64_t nowMs)
{
	const std::filesystem::path path = GetLegacyRegistryPathObject(baseDirectory);
	std::error_code ec;
	if (!std::filesystem::exists(path, ec) || ec) {
		return;
	}

	nlohmann::json root;
	if (!TryLoadJsonFile(path, root, nullptr) ||
		!root.contains("instances") ||
		!root["instances"].is_array()) {
		return;
	}
	for (const auto& item : root["instances"]) {
		MergeRecord(records, ParseRecordJson(item), nowMs);
	}
}

bool LoadInstancesAtBaseDirectory(
	const std::filesystem::path& baseDirectory,
	std::vector<LocalMcpInstanceRegistry::InstanceRecord>& outRecords,
	std::string* outError)
{
	outRecords.clear();
	if (outError != nullptr) {
		outError->clear();
	}

	const std::uint64_t nowMs = GetUnixTimeMilliseconds();
	std::unordered_map<std::string, LocalMcpInstanceRegistry::InstanceRecord> records;
	LoadLegacyRecords(baseDirectory, records, nowMs);

	const std::filesystem::path directory = GetRegistryDirectoryPathObject(baseDirectory);
	std::error_code ec;
	const bool directoryExists = std::filesystem::exists(directory, ec);
	if (ec) {
		SetError(outError, "query registry directory failed, error=" + std::to_string(ec.value()));
		return false;
	}

	std::vector<std::filesystem::path> cleanupPaths;
	if (directoryExists) {
		std::filesystem::directory_iterator iterator(directory, ec);
		if (ec) {
			SetError(outError, "enumerate registry directory failed, error=" + std::to_string(ec.value()));
			return false;
		}
		for (const auto& entry : iterator) {
			if (!entry.is_regular_file(ec)) {
				ec.clear();
				continue;
			}
			const std::filesystem::path path = entry.path();
			if (path.extension() == ".tmp") {
				if (IsArtifactOlderThan(path, std::chrono::milliseconds(kDeadInstanceCleanupTtlMs))) {
					cleanupPaths.push_back(path);
				}
				continue;
			}
			if (path.extension() != ".json") {
				continue;
			}

			nlohmann::json root;
			if (!TryLoadJsonFile(path, root, nullptr)) {
				if (IsArtifactOlderThan(path, std::chrono::milliseconds(kDeadInstanceCleanupTtlMs))) {
					cleanupPaths.push_back(path);
				}
				continue;
			}

			LocalMcpInstanceRegistry::InstanceRecord record = ParseInstanceFileJson(root);
			if (!IsRecordValid(record)) {
				if (IsArtifactOlderThan(path, std::chrono::milliseconds(kDeadInstanceCleanupTtlMs))) {
					cleanupPaths.push_back(path);
				}
				continue;
			}

			const std::uint64_t ageMs = GetRecordAgeMs(record, nowMs);
			if (ageMs > kHardArtifactCleanupTtlMs ||
				(ageMs > kDeadInstanceCleanupTtlMs && !IsProcessLikelyAlive(record.processId))) {
				cleanupPaths.push_back(path);
			}
			MergeRecord(records, std::move(record), nowMs);
		}
	}
	RemovePathsBestEffort(cleanupPaths);

	outRecords.reserve(records.size());
	for (auto& [instanceId, record] : records) {
		(void)instanceId;
		outRecords.push_back(std::move(record));
	}
	std::sort(outRecords.begin(), outRecords.end(), [](const auto& left, const auto& right) {
		if (left.port != right.port) {
			return left.port < right.port;
		}
		return left.instanceId < right.instanceId;
	});
	return true;
}

LocalMcpInstanceRegistry::InstanceRecord BuildSelfTestRecord(
	const std::string& instanceId,
	int port,
	std::uint64_t lastSeenUnixMs)
{
	LocalMcpInstanceRegistry::InstanceRecord record;
	record.instanceId = instanceId;
	record.processId = GetCurrentProcessId();
	record.processPath = "AutoLinkerTest.exe";
	record.processName = "AutoLinkerTest.exe";
	record.port = port;
	record.endpoint = "http://127.0.0.1:" + std::to_string(port) + "/mcp";
	record.sourceFilePathHint = "test.e";
	record.pageNameHint = "test";
	record.pageTypeHint = "assembly";
	record.lastSeenUnixMs = lastSeenUnixMs;
	return record;
}

} // namespace

namespace LocalMcpInstanceRegistry {

std::string GetRegistryFilePath()
{
	return GetLegacyRegistryPathObject(GetRegistryBaseDirectoryObject()).string();
}

std::string GetRegistryDirectoryPath()
{
	return GetRegistryDirectoryPathObject(GetRegistryBaseDirectoryObject()).string();
}

bool UpsertCurrentInstance(const InstanceRecord& record, std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}
	std::unique_lock<std::timed_mutex> lock(g_currentInstanceMutex, std::defer_lock);
	if (!lock.try_lock_for(kLocalOperationLockTimeout)) {
		SetError(outError, "local instance registry update is still in progress");
		return false;
	}
	return UpsertInstanceAtBaseDirectory(GetRegistryBaseDirectoryObject(), record, outError);
}

bool RemoveCurrentInstance(const std::string& instanceId, std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}
	if (instanceId.empty()) {
		return true;
	}
	std::unique_lock<std::timed_mutex> lock(g_currentInstanceMutex, std::defer_lock);
	if (!lock.try_lock_for(kLocalOperationLockTimeout)) {
		SetError(outError, "local instance registry update is still in progress");
		return false;
	}
	return RemoveInstanceAtBaseDirectory(GetRegistryBaseDirectoryObject(), instanceId, outError);
}

bool LoadInstances(std::vector<InstanceRecord>& outRecords, std::string* outError)
{
	return LoadInstancesAtBaseDirectory(GetRegistryBaseDirectoryObject(), outRecords, outError);
}

std::string BuildSelfTestReportJson()
{
	nlohmann::json report = {
		{"name", "local-mcp-instance-registry-isolation"},
		{"ok", false}
	};

	std::error_code ec;
	const std::filesystem::path baseDirectory =
		std::filesystem::temp_directory_path(ec) /
		("AutoLinkerRegistrySelfTest-" + std::to_string(GetCurrentProcessId()) + "-" + std::to_string(GetTickCount64()));
	if (ec) {
		report["error"] = "resolve temp directory failed";
		return report.dump();
	}
	std::filesystem::remove_all(baseDirectory, ec);

	try {
		constexpr int kConcurrentInstanceCount = 8;
		constexpr int kUpdatesPerInstance = 12;
		const std::uint64_t nowMs = GetUnixTimeMilliseconds();
		std::atomic_int writeFailures = 0;
		std::vector<std::thread> writers;
		writers.reserve(kConcurrentInstanceCount);
		for (int index = 0; index < kConcurrentInstanceCount; ++index) {
			writers.emplace_back([&, index]() {
				for (int update = 0; update < kUpdatesPerInstance; ++update) {
					InstanceRecord record = BuildSelfTestRecord(
						"self-instance-" + std::to_string(index),
						20000 + index,
						nowMs + static_cast<std::uint64_t>(update));
					std::string error;
					if (!UpsertInstanceAtBaseDirectory(baseDirectory, record, &error)) {
						writeFailures.fetch_add(1);
					}
				}
			});
		}
		for (auto& writer : writers) {
			writer.join();
		}

		InstanceRecord legacyRecord = BuildSelfTestRecord("legacy-instance", 21000, nowMs);
		const nlohmann::json legacyRoot = {
			{"version", 1},
			{"instances", nlohmann::json::array({BuildRecordJson(legacyRecord)})}
		};
		std::string legacyWriteError;
		const bool legacyWriteOk = SaveJsonFileAtomically(
			GetLegacyRegistryPathObject(baseDirectory),
			legacyRoot,
			&legacyWriteError);

		InstanceRecord staleRecord = BuildSelfTestRecord(
			"stale-instance",
			22000,
			nowMs - kDeadInstanceCleanupTtlMs - 1000);
		staleRecord.processId = 0;
		std::string staleWriteError;
		const bool staleWriteOk = UpsertInstanceAtBaseDirectory(baseDirectory, staleRecord, &staleWriteError);

		std::vector<InstanceRecord> records;
		std::string loadError;
		const bool loadOk = LoadInstancesAtBaseDirectory(baseDirectory, records, &loadError);
		const bool concurrentRecordsOk = records.size() == kConcurrentInstanceCount + 1;
		const bool staleIgnored = std::none_of(records.begin(), records.end(), [](const auto& record) {
			return record.instanceId == "stale-instance";
		});

		std::string removeError;
		const bool removeOk = RemoveInstanceAtBaseDirectory(baseDirectory, "self-instance-0", &removeError);
		std::vector<InstanceRecord> recordsAfterRemove;
		std::string reloadError;
		const bool reloadOk = LoadInstancesAtBaseDirectory(baseDirectory, recordsAfterRemove, &reloadError);
		const bool removedRecordAbsent = std::none_of(
			recordsAfterRemove.begin(),
			recordsAfterRemove.end(),
			[](const auto& record) { return record.instanceId == "self-instance-0"; });
		const bool removeCountOk = recordsAfterRemove.size() == kConcurrentInstanceCount;

		const bool ok =
			writeFailures.load() == 0 &&
			legacyWriteOk &&
			staleWriteOk &&
			loadOk &&
			concurrentRecordsOk &&
			staleIgnored &&
			removeOk &&
			reloadOk &&
			removedRecordAbsent &&
			removeCountOk;
		report["ok"] = ok;
		report["write_failures"] = writeFailures.load();
		report["concurrent_record_count"] = records.size();
		report["legacy_compatibility"] = legacyWriteOk;
		report["stale_record_ignored"] = staleIgnored;
		report["remove_isolated"] = removeOk && removedRecordAbsent && removeCountOk;
		if (!ok) {
			report["legacy_write_error"] = legacyWriteError;
			report["stale_write_error"] = staleWriteError;
			report["load_error"] = loadError;
			report["remove_error"] = removeError;
			report["reload_error"] = reloadError;
		}
	}
	catch (const std::exception& ex) {
		report["error"] = ex.what();
	}
	catch (...) {
		report["error"] = "unknown exception";
	}

	std::filesystem::remove_all(baseDirectory, ec);
	return report.dump();
}

} // namespace LocalMcpInstanceRegistry
