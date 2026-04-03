#include "LocalMcpInstanceRegistry.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <string>
#include <system_error>
#include <vector>

#include "..\\thirdparty\\json.hpp"

namespace {

constexpr const char* kRegistryMutexName = "Local\\AutoLinker.LocalMcpInstanceRegistry";
constexpr const char* kRegistryFileName = "local_mcp_instances.json";
constexpr std::uint64_t kInstanceTtlMs = 15000;
constexpr std::uint64_t kFallbackDeadProcessTtlMs = 60000;

class NamedMutexGuard {
public:
	explicit NamedMutexGuard(const char* mutexName)
	{
		if (mutexName == nullptr || mutexName[0] == '\0') {
			return;
		}
		m_handle = CreateMutexA(nullptr, FALSE, mutexName);
		if (m_handle == nullptr) {
			return;
		}
		const DWORD waitResult = WaitForSingleObject(m_handle, 5000);
		if (waitResult == WAIT_OBJECT_0 || waitResult == WAIT_ABANDONED) {
			m_locked = true;
		}
	}

	~NamedMutexGuard()
	{
		if (m_locked && m_handle != nullptr) {
			ReleaseMutex(m_handle);
		}
		if (m_handle != nullptr) {
			CloseHandle(m_handle);
		}
	}

	NamedMutexGuard(const NamedMutexGuard&) = delete;
	NamedMutexGuard& operator=(const NamedMutexGuard&) = delete;

	bool IsLocked() const
	{
		return m_locked;
	}

private:
	HANDLE m_handle = nullptr;
	bool m_locked = false;
};

std::uint64_t GetUnixTimeMilliseconds()
{
	return static_cast<std::uint64_t>(
		std::chrono::duration_cast<std::chrono::milliseconds>(
			std::chrono::system_clock::now().time_since_epoch()).count());
}

std::filesystem::path GetRegistryPathObject()
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
	return dir / kRegistryFileName;
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

bool TryLoadRegistryJson(
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
		outRoot["version"] = 1;
		outRoot["instances"] = nlohmann::json::array();
		return true;
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
			outRoot = nlohmann::json::object();
		}
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) {
			*outError = std::string("parse registry json failed: ") + ex.what();
		}
		outRoot = nlohmann::json::object();
	}

	if (!outRoot.contains("version")) {
		outRoot["version"] = 1;
	}
	if (!outRoot.contains("instances") || !outRoot["instances"].is_array()) {
		outRoot["instances"] = nlohmann::json::array();
	}
	return true;
}

bool SaveRegistryJson(
	const std::filesystem::path& path,
	const nlohmann::json& root,
	std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}

	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);

	const std::filesystem::path tempPath = path.string() + ".tmp";
	std::ofstream out(tempPath, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		if (outError != nullptr) {
			*outError = "open registry temp file failed";
		}
		return false;
	}

	static constexpr unsigned char kUtf8Bom[] = { 0xEF, 0xBB, 0xBF };
	out.write(reinterpret_cast<const char*>(kUtf8Bom), sizeof(kUtf8Bom));
	const std::string text = root.dump(2);
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	out.close();
	if (!out) {
		if (outError != nullptr) {
			*outError = "write registry temp file failed";
		}
		return false;
	}

	if (MoveFileExW(
			tempPath.wstring().c_str(),
			path.wstring().c_str(),
			MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == FALSE) {
		if (outError != nullptr) {
			*outError = "replace registry file failed";
		}
		std::filesystem::remove(tempPath, ec);
		return false;
	}

	return true;
}

std::vector<LocalMcpInstanceRegistry::InstanceRecord> ParseAndCleanupRecords(
	const nlohmann::json& root,
	bool& outChanged)
{
	outChanged = false;
	std::vector<LocalMcpInstanceRegistry::InstanceRecord> records;
	const std::uint64_t nowMs = GetUnixTimeMilliseconds();

	if (!root.contains("instances") || !root["instances"].is_array()) {
		outChanged = true;
		return records;
	}

	for (const auto& item : root["instances"]) {
		LocalMcpInstanceRegistry::InstanceRecord record = ParseRecordJson(item);
		if (record.instanceId.empty() || record.port <= 0 || record.endpoint.empty()) {
			outChanged = true;
			continue;
		}

		const std::uint64_t ageMs = nowMs >= record.lastSeenUnixMs
			? (nowMs - record.lastSeenUnixMs)
			: 0;
		const bool alive = IsProcessLikelyAlive(record.processId);
		if ((!alive && ageMs > kFallbackDeadProcessTtlMs) || ageMs > kInstanceTtlMs) {
			outChanged = true;
			continue;
		}

		records.push_back(std::move(record));
	}

	std::sort(records.begin(), records.end(), [](const auto& left, const auto& right) {
		if (left.port != right.port) {
			return left.port < right.port;
		}
		return left.instanceId < right.instanceId;
	});
	return records;
}

nlohmann::json BuildRootJsonFromRecords(const std::vector<LocalMcpInstanceRegistry::InstanceRecord>& records)
{
	nlohmann::json instances = nlohmann::json::array();
	for (const auto& record : records) {
		instances.push_back(BuildRecordJson(record));
	}

	return {
		{"version", 1},
		{"instances", std::move(instances)}
	};
}

} // namespace

namespace LocalMcpInstanceRegistry {

std::string GetRegistryFilePath()
{
	return GetRegistryPathObject().string();
}

bool UpsertCurrentInstance(const InstanceRecord& record, std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}
	if (record.instanceId.empty()) {
		if (outError != nullptr) {
			*outError = "instance_id is empty";
		}
		return false;
	}
	if (record.port <= 0) {
		if (outError != nullptr) {
			*outError = "port is invalid";
		}
		return false;
	}
	if (record.endpoint.empty()) {
		if (outError != nullptr) {
			*outError = "endpoint is empty";
		}
		return false;
	}

	NamedMutexGuard guard(kRegistryMutexName);
	if (!guard.IsLocked()) {
		if (outError != nullptr) {
			*outError = "lock registry mutex failed";
		}
		return false;
	}

	const std::filesystem::path path = GetRegistryPathObject();
	nlohmann::json root;
	TryLoadRegistryJson(path, root, nullptr);

	bool changed = false;
	std::vector<InstanceRecord> records = ParseAndCleanupRecords(root, changed);

	bool updated = false;
	for (auto& current : records) {
		if (current.instanceId == record.instanceId) {
			current = record;
			updated = true;
			break;
		}
	}
	if (!updated) {
		records.push_back(record);
	}

	std::sort(records.begin(), records.end(), [](const auto& left, const auto& right) {
		if (left.port != right.port) {
			return left.port < right.port;
		}
		return left.instanceId < right.instanceId;
	});

	return SaveRegistryJson(path, BuildRootJsonFromRecords(records), outError);
}

bool RemoveCurrentInstance(const std::string& instanceId, std::string* outError)
{
	if (outError != nullptr) {
		outError->clear();
	}
	if (instanceId.empty()) {
		return true;
	}

	NamedMutexGuard guard(kRegistryMutexName);
	if (!guard.IsLocked()) {
		if (outError != nullptr) {
			*outError = "lock registry mutex failed";
		}
		return false;
	}

	const std::filesystem::path path = GetRegistryPathObject();
	nlohmann::json root;
	TryLoadRegistryJson(path, root, nullptr);

	bool changed = false;
	std::vector<InstanceRecord> records = ParseAndCleanupRecords(root, changed);
	const size_t oldSize = records.size();
	records.erase(
		std::remove_if(records.begin(), records.end(), [&instanceId](const auto& record) {
			return record.instanceId == instanceId;
		}),
		records.end());

	if (records.size() == oldSize && !changed) {
		return true;
	}

	return SaveRegistryJson(path, BuildRootJsonFromRecords(records), outError);
}

bool LoadInstances(std::vector<InstanceRecord>& outRecords, std::string* outError)
{
	outRecords.clear();
	if (outError != nullptr) {
		outError->clear();
	}

	NamedMutexGuard guard(kRegistryMutexName);
	if (!guard.IsLocked()) {
		if (outError != nullptr) {
			*outError = "lock registry mutex failed";
		}
		return false;
	}

	const std::filesystem::path path = GetRegistryPathObject();
	nlohmann::json root;
	if (!TryLoadRegistryJson(path, root, outError)) {
		return false;
	}

	bool changed = false;
	outRecords = ParseAndCleanupRecords(root, changed);
	if (changed) {
		SaveRegistryJson(path, BuildRootJsonFromRecords(outRecords), nullptr);
	}
	return true;
}

} // namespace LocalMcpInstanceRegistry
