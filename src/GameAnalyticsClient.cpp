#include "GameAnalyticsClient.h"

#include <Windows.h>
#include <bcrypt.h>
#ifdef min
#undef min
#endif
#ifdef max
#undef max
#endif

#include <algorithm>
#include <chrono>
#include <condition_variable>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <deque>
#include <format>
#include <functional>
#include <mutex>
#include <optional>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AutoLinkerVersion.h"
#include "ConfigManager.h"
#include "GameAnalyticsEventBatcher.h"
#include "Logger.h"
#include "WinINetUtil.h"

#pragma comment(lib, "bcrypt.lib")

namespace {

constexpr const char* kGameKey = "38499c977724e8cf1c4286d4dc987864";
constexpr const char* kSecretKey = "eae352d617a567f5bae193cf7c7661593a1db67b";
constexpr const char* kApiBaseUrl = "https://api.gameanalytics.com";
constexpr const char* kSdkVersion = "rest api v2";
constexpr const char* kPlatform = "windows";
constexpr const char* kManufacturer = "microsoft";
constexpr const char* kDevice = "windows";
constexpr int kHttpTimeoutMs = 8000;
constexpr size_t kFlushEventThreshold = 20;
constexpr size_t kMaxPendingEvents = 128;
constexpr int kMaxEventSubmitFailures = 3;
constexpr auto kEventFlushInterval = std::chrono::seconds(30);

constexpr const char* kConfigEnabled = "gameanalytics.enabled";
constexpr const char* kConfigUserId = "gameanalytics.user_id";
constexpr const char* kConfigDeviceId = "gameanalytics.device_id";
constexpr const char* kConfigSessionNum = "gameanalytics.session_num";
constexpr const char* kStartShowIdeConfigKey = "START_SHOW_IDE";

enum class WorkerTask {
	Init,
	FlushEvents,
	RefreshRemoteConfigs
};

GameAnalyticsEventBatcher::Policy BuildProductionEventBatchPolicy()
{
	return {
		kFlushEventThreshold,
		kMaxPendingEvents,
		kMaxEventSubmitFailures,
		std::chrono::duration_cast<GameAnalyticsEventBatcher::Duration>(kEventFlushInterval),
		std::chrono::seconds(1)
	};
}

struct ClientState {
	ConfigManager* configManager = nullptr;
	bool enabled = true;
	bool initialized = false;
	bool shuttingDown = false;
	bool collectorReady = false;
	bool collectorEnabled = true;
	bool startupEventQueued = false;
	bool hasServerTsOffset = false;
	long long serverTsOffset = 0;
	long long sessionStartUnix = 0;
	std::string userId;
	std::string deviceId;
	std::string sessionId;
	int sessionNum = 1;
	std::string osVersion = "windows 10.0.0";
	GameAnalyticsEventBatcher eventBatcher{BuildProductionEventBatchPolicy()};
	bool flushTaskQueued = false;
	std::deque<WorkerTask> tasks;
	GameAnalyticsClient::RemoteConfigSnapshot remoteConfigs;
	bool startShowIdeValueSeen = false;
	std::string startShowIdeLastValue;
};

std::mutex g_mutex;
std::condition_variable g_cv;
std::thread g_worker;
ClientState g_state;
using EventSubmitOverride = std::function<std::pair<std::string, int>(
	const nlohmann::json&,
	const std::string&,
	const std::string&)>;
EventSubmitOverride g_eventSubmitOverride;

std::string ToLowerAscii(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string TrimAscii(std::string text)
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

bool IsExplicitFalse(const std::string& value)
{
	const std::string normalized = ToLowerAscii(TrimAscii(value));
	return normalized == "0" ||
		normalized == "false" ||
		normalized == "off" ||
		normalized == "no" ||
		normalized == "disabled";
}

long long UnixTimeSeconds()
{
	return std::chrono::duration_cast<std::chrono::seconds>(
		std::chrono::system_clock::now().time_since_epoch()).count();
}

std::string Base64Encode(const unsigned char* data, size_t size)
{
	static constexpr char kTable[] =
		"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	std::string out;
	out.reserve(((size + 2) / 3) * 4);
	for (size_t i = 0; i < size; i += 3) {
		const unsigned int a = data[i];
		const unsigned int b = (i + 1 < size) ? data[i + 1] : 0;
		const unsigned int c = (i + 2 < size) ? data[i + 2] : 0;
		const unsigned int n = (a << 16) | (b << 8) | c;

		out.push_back(kTable[(n >> 18) & 0x3F]);
		out.push_back(kTable[(n >> 12) & 0x3F]);
		out.push_back(i + 1 < size ? kTable[(n >> 6) & 0x3F] : '=');
		out.push_back(i + 2 < size ? kTable[n & 0x3F] : '=');
	}
	return out;
}

std::string HmacSha256Base64(const std::string& body, const std::string& secret)
{
	BCRYPT_ALG_HANDLE algorithm = nullptr;
	BCRYPT_HASH_HANDLE hash = nullptr;
	std::vector<unsigned char> hashObject;
	std::vector<unsigned char> hashValue;
	DWORD objectLength = 0;
	DWORD hashLength = 0;
	DWORD resultLength = 0;

	if (BCryptOpenAlgorithmProvider(
			&algorithm,
			BCRYPT_SHA256_ALGORITHM,
			nullptr,
			BCRYPT_ALG_HANDLE_HMAC_FLAG) < 0) {
		return std::string();
	}

	auto cleanup = [&]() {
		if (hash != nullptr) {
			BCryptDestroyHash(hash);
			hash = nullptr;
		}
		if (algorithm != nullptr) {
			BCryptCloseAlgorithmProvider(algorithm, 0);
			algorithm = nullptr;
		}
	};

	if (BCryptGetProperty(
			algorithm,
			BCRYPT_OBJECT_LENGTH,
			reinterpret_cast<PUCHAR>(&objectLength),
			sizeof(objectLength),
			&resultLength,
			0) < 0 ||
		BCryptGetProperty(
			algorithm,
			BCRYPT_HASH_LENGTH,
			reinterpret_cast<PUCHAR>(&hashLength),
			sizeof(hashLength),
			&resultLength,
			0) < 0) {
		cleanup();
		return std::string();
	}

	hashObject.resize(objectLength);
	hashValue.resize(hashLength);
	if (BCryptCreateHash(
			algorithm,
			&hash,
			hashObject.data(),
			static_cast<ULONG>(hashObject.size()),
			reinterpret_cast<PUCHAR>(const_cast<char*>(secret.data())),
			static_cast<ULONG>(secret.size()),
			0) < 0 ||
		BCryptHashData(
			hash,
			reinterpret_cast<PUCHAR>(const_cast<char*>(body.data())),
			static_cast<ULONG>(body.size()),
			0) < 0 ||
		BCryptFinishHash(hash, hashValue.data(), static_cast<ULONG>(hashValue.size()), 0) < 0) {
		cleanup();
		return std::string();
	}

	cleanup();
	return Base64Encode(hashValue.data(), hashValue.size());
}

uint32_t Crc32(const unsigned char* data, size_t size)
{
	static uint32_t table[256] = {};
	static bool initialized = false;
	if (!initialized) {
		for (uint32_t i = 0; i < 256; ++i) {
			uint32_t c = i;
			for (int bit = 0; bit < 8; ++bit) {
				c = (c & 1U) ? (0xEDB88320U ^ (c >> 1U)) : (c >> 1U);
			}
			table[i] = c;
		}
		initialized = true;
	}

	uint32_t c = 0xFFFFFFFFU;
	for (size_t i = 0; i < size; ++i) {
		c = table[(c ^ data[i]) & 0xFFU] ^ (c >> 8U);
	}
	return c ^ 0xFFFFFFFFU;
}

void AppendUInt16Le(std::string& out, uint16_t value)
{
	out.push_back(static_cast<char>(value & 0xFFU));
	out.push_back(static_cast<char>((value >> 8U) & 0xFFU));
}

void AppendUInt32Le(std::string& out, uint32_t value)
{
	out.push_back(static_cast<char>(value & 0xFFU));
	out.push_back(static_cast<char>((value >> 8U) & 0xFFU));
	out.push_back(static_cast<char>((value >> 16U) & 0xFFU));
	out.push_back(static_cast<char>((value >> 24U) & 0xFFU));
}

std::string GzipStoredDeflate(const std::string& input)
{
	std::string out;
	out.reserve(input.size() + 64);
	out.push_back(static_cast<char>(0x1F));
	out.push_back(static_cast<char>(0x8B));
	out.push_back(static_cast<char>(0x08));
	out.push_back(static_cast<char>(0x00));
	AppendUInt32Le(out, 0);
	out.push_back(static_cast<char>(0x00));
	out.push_back(static_cast<char>(0xFF));

	size_t offset = 0;
	do {
		const size_t remaining = input.size() - offset;
		const uint16_t blockLen = static_cast<uint16_t>((std::min<size_t>)(remaining, 65535));
		const bool finalBlock = offset + blockLen >= input.size();
		out.push_back(static_cast<char>(finalBlock ? 0x01 : 0x00));
		AppendUInt16Le(out, blockLen);
		AppendUInt16Le(out, static_cast<uint16_t>(~blockLen));
		if (blockLen > 0) {
			out.append(input.data() + offset, blockLen);
		}
		offset += blockLen;
	} while (offset < input.size());

	AppendUInt32Le(out, Crc32(reinterpret_cast<const unsigned char*>(input.data()), input.size()));
	AppendUInt32Le(out, static_cast<uint32_t>(input.size() & 0xFFFFFFFFU));
	return out;
}

bool FillRandomBytes(unsigned char* data, size_t size)
{
	return BCryptGenRandom(
		nullptr,
		data,
		static_cast<ULONG>(size),
		BCRYPT_USE_SYSTEM_PREFERRED_RNG) >= 0;
}

std::string GenerateUuidLower()
{
	unsigned char bytes[16] = {};
	if (!FillRandomBytes(bytes, sizeof(bytes))) {
		const auto now = static_cast<uint64_t>(GetTickCount64());
		for (size_t i = 0; i < sizeof(bytes); ++i) {
			bytes[i] = static_cast<unsigned char>((now >> ((i % 8) * 8)) & 0xFFU);
		}
	}

	bytes[6] = static_cast<unsigned char>((bytes[6] & 0x0FU) | 0x40U);
	bytes[8] = static_cast<unsigned char>((bytes[8] & 0x3FU) | 0x80U);

	static constexpr char kHex[] = "0123456789abcdef";
	std::string out;
	out.reserve(36);
	for (int i = 0; i < 16; ++i) {
		if (i == 4 || i == 6 || i == 8 || i == 10) {
			out.push_back('-');
		}
		out.push_back(kHex[(bytes[i] >> 4U) & 0x0FU]);
		out.push_back(kHex[bytes[i] & 0x0FU]);
	}
	return out;
}

int GenerateRemoteConfigSalt()
{
	unsigned char bytes[4] = {};
	if (!FillRandomBytes(bytes, sizeof(bytes))) {
		return static_cast<int>(GetTickCount() & 0x7FFFFFFF);
	}
	uint32_t value = 0;
	for (unsigned char byte : bytes) {
		value = (value << 8U) | byte;
	}
	return static_cast<int>(value & 0x7FFFFFFFU);
}

std::string GetWindowsVersionString()
{
	OSVERSIONINFOEXW info = {};
	info.dwOSVersionInfoSize = sizeof(info);
#pragma warning(push)
#pragma warning(disable:4996)
	const BOOL versionOk = GetVersionExW(reinterpret_cast<OSVERSIONINFOW*>(&info));
#pragma warning(pop)
	if (versionOk == FALSE) {
		return "windows 10.0.0";
	}
	return std::format("windows {}.{}.{}", info.dwMajorVersion, info.dwMinorVersion, info.dwBuildNumber);
}

std::string BuildVersionString()
{
	std::string version = AUTOLINKER_VERSION;
	if (version.empty() || version.size() > 32) {
		version = "0.0.0";
	}
	return version;
}

std::string BuildAuthorizationHeader(
	const std::string& body,
	const std::string& secretKey)
{
	const std::string auth = HmacSha256Base64(body, secretKey);
	if (auth.empty()) {
		return std::string();
	}
	return "Authorization: " + auth + "\r\n";
}

std::string BuildJsonHeaders(
	const std::string& body,
	const std::string& secretKey = kSecretKey)
{
	std::string headers = "Content-Type: application/json\r\n";
	headers += BuildAuthorizationHeader(body, secretKey);
	return headers;
}

std::string BuildGzipJsonHeaders(
	const std::string& gzippedBody,
	const std::string& secretKey = kSecretKey)
{
	std::string headers = "Content-Type: application/json\r\nContent-Encoding: gzip\r\n";
	headers += BuildAuthorizationHeader(gzippedBody, secretKey);
	return headers;
}

void LogGA(const std::string& message)
{
	Logger::Instance().Write("GameAnalytics", message);
}

std::string BuildRemoteConfigRequestError(int httpStatus, const std::string& responseBody)
{
	const std::string responsePreview = responseBody.substr(0, 600);
	if (httpStatus == 0) {
		return responsePreview.empty()
			? "remote config transport error without details"
			: responsePreview;
	}

	std::string error = std::format("remote config HTTP {}", httpStatus);
	if (!responsePreview.empty()) {
		error += ": ";
		error += responsePreview;
	}
	return error;
}

std::vector<std::string> ParseStartShowIdeMessages(const std::string& rawValue, std::string& outError)
{
	outError.clear();
	std::vector<std::string> messages;
	const nlohmann::json parsed = nlohmann::json::parse(rawValue, nullptr, false);
	if (parsed.is_discarded()) {
		outError = "START_SHOW_IDE is not valid JSON";
		return messages;
	}
	if (!parsed.is_array()) {
		outError = "START_SHOW_IDE must be a JSON array";
		return messages;
	}
	for (const auto& item : parsed) {
		if (!item.is_string()) {
			outError = "START_SHOW_IDE array items must be strings";
			messages.clear();
			return messages;
		}
		messages.push_back(item.get<std::string>());
	}
	return messages;
}

bool MarkStartShowIdeValueForOutput(const std::string& rawValue)
{
	std::lock_guard<std::mutex> lock(g_mutex);
	if (g_state.startShowIdeValueSeen && g_state.startShowIdeLastValue == rawValue) {
		return false;
	}
	g_state.startShowIdeValueSeen = true;
	g_state.startShowIdeLastValue = rawValue;
	return true;
}

void OutputStartShowIdeRemoteConfig(const GameAnalyticsClient::RemoteConfigSnapshot& snapshot)
{
	const auto it = snapshot.values.find(kStartShowIdeConfigKey);
	if (it == snapshot.values.end()) {
		return;
	}
	if (!MarkStartShowIdeValueForOutput(it->second)) {
		return;
	}

	std::string parseError;
	const std::vector<std::string> messages = ParseStartShowIdeMessages(it->second, parseError);
	if (!parseError.empty()) {
		LogGA(parseError);
		return;
	}
	for (const std::string& message : messages) {
		Logger::Instance().WriteSplit("", "[GameAnalytics] START_SHOW_IDE: " + message, message);
	}
}

nlohmann::json BuildSharedAnnotationsLocked()
{
	const long long now = UnixTimeSeconds();
	const long long clientTs = g_state.hasServerTsOffset ? now + g_state.serverTsOffset : now;
	nlohmann::json event = {
		{"v", 2},
		{"event_uuid", GenerateUuidLower()},
		{"user_id", g_state.userId},
		{"client_ts", clientTs},
		{"sdk_version", kSdkVersion},
		{"os_version", g_state.osVersion},
		{"manufacturer", kManufacturer},
		{"device", g_state.deviceId.empty() ? kDevice : g_state.deviceId},
		{"platform", kPlatform},
		{"session_id", g_state.sessionId},
		{"session_num", g_state.sessionNum},
		{"build", BuildVersionString()}
	};

	if (!g_state.remoteConfigs.values.empty()) {
		nlohmann::json configurations = nlohmann::json::object();
		for (const auto& [key, value] : g_state.remoteConfigs.values) {
			configurations[key] = value;
		}
		event["configurations"] = std::move(configurations);
	}
	if (!g_state.remoteConfigs.abId.empty()) {
		event["ab_id"] = g_state.remoteConfigs.abId;
	}
	if (!g_state.remoteConfigs.abVariantId.empty()) {
		event["ab_variant_id"] = g_state.remoteConfigs.abVariantId;
	}
	return event;
}

nlohmann::json BuildStartupEventLocked()
{
	nlohmann::json event = BuildSharedAnnotationsLocked();
	event["category"] = "user";
	return event;
}

nlohmann::json BuildSessionEndEventLocked()
{
	nlohmann::json event = BuildSharedAnnotationsLocked();
	const long long length = (std::max<long long>)(0, UnixTimeSeconds() - g_state.sessionStartUnix);
	event["category"] = "session_end";
	event["length"] = static_cast<int>((std::min<long long>)(length, 172800));
	return event;
}

nlohmann::json BuildDesignEventLocked(const std::string& eventId, double value)
{
	nlohmann::json event = BuildSharedAnnotationsLocked();
	event["category"] = "design";
	event["event_id"] = eventId;
	event["value"] = value;
	return event;
}

void PushEventLocked(nlohmann::json event)
{
	if (!g_state.initialized || !g_state.enabled || g_state.shuttingDown) {
		return;
	}
	const auto enqueueResult = g_state.eventBatcher.Enqueue(
		std::move(event),
		GameAnalyticsEventBatcher::Clock::now());
	if (enqueueResult.droppedOldestEvent) {
		LogGA("pending events reached capacity; dropped oldest event");
	}
}

bool QueueTaskLocked(WorkerTask task)
{
	if (!g_worker.joinable() || g_state.shuttingDown) {
		return false;
	}
	g_state.tasks.push_back(task);
	g_cv.notify_one();
	return true;
}

void ScheduleEventFlushLocked()
{
	if (g_state.eventBatcher.Empty()) {
		return;
	}
	if (g_state.eventBatcher.IsReady(GameAnalyticsEventBatcher::Clock::now()) &&
		!g_state.flushTaskQueued &&
		QueueTaskLocked(WorkerTask::FlushEvents)) {
		g_state.flushTaskQueued = true;
	}
	g_cv.notify_one();
}

bool UpdateServerOffsetFromJson(const nlohmann::json& body)
{
	if (!body.is_object() || !body.contains("server_ts") || !body["server_ts"].is_number_integer()) {
		return false;
	}
	const long long serverTs = body["server_ts"].get<long long>();
	std::lock_guard<std::mutex> lock(g_mutex);
	g_state.hasServerTsOffset = true;
	g_state.serverTsOffset = serverTs - UnixTimeSeconds();
	return true;
}

void UpdateRemoteSnapshotRequestStarted()
{
	std::lock_guard<std::mutex> lock(g_mutex);
	g_state.remoteConfigs.ready = false;
	g_state.remoteConfigs.requestInFlight = true;
	g_state.remoteConfigs.httpStatus = 0;
	g_state.remoteConfigs.error.clear();
}

void UpdateRemoteSnapshotFailure(int status, const std::string& error)
{
	std::lock_guard<std::mutex> lock(g_mutex);
	g_state.remoteConfigs.ready = false;
	g_state.remoteConfigs.requestInFlight = false;
	g_state.remoteConfigs.httpStatus = status;
	g_state.remoteConfigs.error = error;
}

void UpdateRemoteSnapshotSuccess(int status, const nlohmann::json& body)
{
	GameAnalyticsClient::RemoteConfigSnapshot snapshot;
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		snapshot = g_state.remoteConfigs;
	}

	snapshot.ready = true;
	snapshot.requestInFlight = false;
	snapshot.httpStatus = status;
	snapshot.error.clear();
	if (body.is_object()) {
		if (body.contains("server_ts") && body["server_ts"].is_number_integer()) {
			snapshot.serverTs = body["server_ts"].get<long long>();
		}
		if (body.contains("configs_hash") && body["configs_hash"].is_string()) {
			snapshot.configsHash = body["configs_hash"].get<std::string>();
		}
		if (body.contains("ab_id") && body["ab_id"].is_string()) {
			snapshot.abId = body["ab_id"].get<std::string>();
		}
		if (body.contains("ab_variant_id") && body["ab_variant_id"].is_string()) {
			snapshot.abVariantId = body["ab_variant_id"].get<std::string>();
		}
		if (body.contains("configs") && body["configs"].is_array()) {
			snapshot.values.clear();
			snapshot.valueTypes.clear();
			for (const auto& item : body["configs"]) {
				if (!item.is_object() ||
					!item.contains("key") ||
					!item["key"].is_string() ||
					!item.contains("value")) {
					continue;
				}
				const std::string key = item["key"].get<std::string>();
				snapshot.values[key] = item["value"].is_string()
					? item["value"].get<std::string>()
					: item["value"].dump();
				snapshot.valueTypes.push_back(
					item.contains("type") && item["type"].is_string()
						? item["type"].get<std::string>()
						: std::string());
			}
		}
	}

	std::lock_guard<std::mutex> lock(g_mutex);
	g_state.remoteConfigs = std::move(snapshot);
}

void PerformCollectorInit()
{
	std::string osVersion;
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		osVersion = g_state.osVersion;
	}

	const nlohmann::json bodyJson = {
		{"platform", kPlatform},
		{"os_version", osVersion},
		{"sdk_version", kSdkVersion}
	};
	const std::string body = bodyJson.dump();
	const std::string headers = BuildJsonHeaders(body);
	if (headers.find("Authorization: ") == std::string::npos) {
		LogGA("collector init skipped: hmac failed");
		return;
	}

	const auto response = PerformPostRequest(
		std::format("{}/v2/{}/init", kApiBaseUrl, kGameKey),
		body,
		headers,
		kHttpTimeoutMs,
		true,
		true);

	bool enabled = true;
	bool ready = response.second == 200;
	try {
		const nlohmann::json parsed = response.first.empty()
			? nlohmann::json::object()
			: nlohmann::json::parse(response.first);
		UpdateServerOffsetFromJson(parsed);
		if (parsed.contains("enabled") && parsed["enabled"].is_boolean()) {
			enabled = parsed["enabled"].get<bool>();
		}
	}
	catch (...) {
	}

	{
		std::lock_guard<std::mutex> lock(g_mutex);
		g_state.collectorReady = ready;
		g_state.collectorEnabled = enabled;
	}
	LogGA(std::format("collector init status={} enabled={}", response.second, enabled ? 1 : 0));
}

void PerformRemoteConfigRefresh()
{
	UpdateRemoteSnapshotRequestStarted();

	std::string userId;
	std::string deviceId;
	std::string osVersion;
	std::string configsHash;
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		userId = g_state.userId;
		deviceId = g_state.deviceId;
		osVersion = g_state.osVersion;
		configsHash = g_state.remoteConfigs.configsHash;
	}

	const nlohmann::json bodyJson = {
		{"user_id", userId},
		{"platform", kPlatform},
		{"os_version", osVersion},
		{"sdk_version", kSdkVersion},
		{"manufacturer", kManufacturer},
		{"device", deviceId.empty() ? kDevice : deviceId},
		{"build", BuildVersionString()},
		{"random_salt", GenerateRemoteConfigSalt()}
	};
	const std::string body = bodyJson.dump();
	const std::string headers = BuildJsonHeaders(body);
	if (headers.find("Authorization: ") == std::string::npos) {
		UpdateRemoteSnapshotFailure(0, "hmac failed");
		LogGA("remote config skipped: hmac failed");
		return;
	}

	const auto response = PerformPostRequest(
		std::format(
			"{}/remote_configs/v1/init?game_key={}&interval_seconds=0&configs_hash={}&config_vsn_supported=3",
			kApiBaseUrl,
			kGameKey,
			configsHash),
		body,
		headers,
		kHttpTimeoutMs,
		true,
		true);

	if (response.second != 200 && response.second != 201) {
		const std::string error = BuildRemoteConfigRequestError(response.second, response.first);
		UpdateRemoteSnapshotFailure(response.second, error);
		LogGA(std::format(
			"remote config request failed status={} error={}",
			response.second,
			error));
		return;
	}

	try {
		const nlohmann::json parsed = response.first.empty()
			? nlohmann::json::object()
			: nlohmann::json::parse(response.first);
		UpdateServerOffsetFromJson(parsed);
		UpdateRemoteSnapshotSuccess(response.second, parsed);
		const auto snapshot = GameAnalyticsClient::GetRemoteConfigs();
		OutputStartShowIdeRemoteConfig(snapshot);
		LogGA(std::format(
			"remote config success status={} hash={} keys={} response={}",
			response.second,
			snapshot.configsHash,
			snapshot.values.size(),
			response.first.substr(0, 600)));
	}
	catch (const std::exception& ex) {
		UpdateRemoteSnapshotFailure(response.second, ex.what());
		LogGA(std::format(
			"remote config parse failed status={} error={} response={}",
			response.second,
			ex.what(),
			response.first.substr(0, 600)));
	}
}

void FlushEvents(bool forceRetry, bool forcePending)
{
	GameAnalyticsEventBatcher::Batch batch;
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		if (!g_state.enabled || !g_state.collectorEnabled) {
			g_state.eventBatcher.Clear();
			return;
		}
		const auto batchMaybe = g_state.eventBatcher.TakeBatch(
			GameAnalyticsEventBatcher::Clock::now(),
			forceRetry,
			forcePending);
		if (!batchMaybe.has_value()) {
			return;
		}
		batch = std::move(*batchMaybe);
	}
	if (batch.events.empty()) {
		return;
	}

	nlohmann::json eventArray = nlohmann::json::array();
	for (const auto& event : batch.events) {
		eventArray.push_back(event);
	}

	const std::string body = eventArray.dump();
	const std::string gzippedBody = GzipStoredDeflate(body);
	const std::string headers = BuildGzipJsonHeaders(gzippedBody);
	std::string responseBody;
	int responseStatus = 0;
	if (headers.find("Authorization: ") == std::string::npos) {
		responseBody = "hmac failed";
	}
	else {
		const auto response = g_eventSubmitOverride
			? g_eventSubmitOverride(eventArray, gzippedBody, headers)
			: PerformPostRequest(
				std::format("{}/v2/{}/events", kApiBaseUrl, kGameKey),
				gzippedBody,
				headers,
				kHttpTimeoutMs,
				true,
				true);
		responseBody = response.first;
		responseStatus = response.second;
	}
	if (responseStatus == 200) {
		LogGA(std::format(
			"events submitted count={} retry_failures={}",
			eventArray.size(),
			batch.failureCount));
		std::lock_guard<std::mutex> lock(g_mutex);
		ScheduleEventFlushLocked();
	}
	else {
		GameAnalyticsEventBatcher::FailureResult failureResult;
		{
			std::lock_guard<std::mutex> lock(g_mutex);
			failureResult = g_state.eventBatcher.CompleteFailure(
				std::move(batch),
				GameAnalyticsEventBatcher::Clock::now());
		}
		if (failureResult.disposition == GameAnalyticsEventBatcher::FailureDisposition::Dropped) {
			LogGA(std::format(
				"events dropped count={} failures={} status={} response={}",
				eventArray.size(),
				failureResult.failureCount,
				responseStatus,
				responseBody.substr(0, 300)));
			std::lock_guard<std::mutex> lock(g_mutex);
			ScheduleEventFlushLocked();
		}
		else {
			LogGA(std::format(
				"events submit failed count={} failure={} retry_in={}s status={} response={}",
				eventArray.size(),
				failureResult.failureCount,
				std::chrono::duration_cast<std::chrono::seconds>(failureResult.retryDelay).count(),
				responseStatus,
				responseBody.substr(0, 300)));
			std::lock_guard<std::mutex> lock(g_mutex);
			g_cv.notify_one();
		}
	}
}

void QueueStartupEvents()
{
	std::lock_guard<std::mutex> lock(g_mutex);
	if (!g_state.initialized || !g_state.enabled || !g_state.collectorEnabled) {
		return;
	}
	if (g_state.startupEventQueued) {
		return;
	}
	PushEventLocked(BuildStartupEventLocked());
	g_state.startupEventQueued = true;
}

void WorkerMain()
{
	for (;;) {
		WorkerTask task = WorkerTask::FlushEvents;
		bool hasTask = false;
		bool forceFlush = false;
		{
			std::unique_lock<std::mutex> lock(g_mutex);
			for (;;) {
				if (!g_state.tasks.empty()) {
					task = g_state.tasks.front();
					g_state.tasks.pop_front();
					if (task == WorkerTask::FlushEvents) {
						g_state.flushTaskQueued = false;
					}
					forceFlush = g_state.shuttingDown;
					hasTask = true;
					break;
				}
				if (g_state.shuttingDown) {
					break;
				}

				const auto now = GameAnalyticsEventBatcher::Clock::now();
				if (g_state.eventBatcher.IsReady(now)) {
					task = WorkerTask::FlushEvents;
					hasTask = true;
					break;
				}

				const auto deadline = g_state.eventBatcher.NextDeadline();
				if (deadline.has_value()) {
					g_cv.wait_until(lock, *deadline);
				}
				else {
					g_cv.wait(lock, []() {
						return g_state.shuttingDown ||
							!g_state.tasks.empty() ||
							!g_state.eventBatcher.Empty();
					});
				}
			}
		}
		if (!hasTask) {
			break;
		}

		try {
			if (task == WorkerTask::Init) {
				PerformCollectorInit();
				PerformRemoteConfigRefresh();
				QueueStartupEvents();
				FlushEvents(false, true);
			}
			else if (task == WorkerTask::RefreshRemoteConfigs) {
				PerformRemoteConfigRefresh();
				FlushEvents(false, true);
			}
			else if (task == WorkerTask::FlushEvents) {
				FlushEvents(forceFlush, forceFlush);
			}
		}
		catch (const std::exception& ex) {
			LogGA(std::format("worker task exception: {}", ex.what()));
		}
		catch (...) {
			LogGA("worker task exception: unknown");
		}
	}

	for (;;) {
		{
			std::lock_guard<std::mutex> lock(g_mutex);
			if (g_state.eventBatcher.Empty()) {
				break;
			}
		}
		FlushEvents(true, true);
	}
	LogGA("worker stopped");
}

int LoadAndIncrementSessionNum(ConfigManager* configManager)
{
	if (configManager == nullptr) {
		return 1;
	}

	int sessionNum = 0;
	try {
		const std::string raw = TrimAscii(configManager->getValue(kConfigSessionNum));
		if (!raw.empty()) {
			sessionNum = std::stoi(raw);
		}
	}
	catch (...) {
		sessionNum = 0;
	}
	sessionNum = (std::max)(1, sessionNum + 1);
	configManager->setValue(kConfigSessionNum, std::to_string(sessionNum));
	return sessionNum;
}

std::string LoadOrCreateUserId(ConfigManager* configManager)
{
	if (configManager == nullptr) {
		return GenerateUuidLower();
	}

	std::string userId = TrimAscii(configManager->getValue(kConfigUserId));
	if (userId.empty()) {
		userId = TrimAscii(configManager->getValue(kConfigDeviceId));
	}
	if (userId.empty()) {
		userId = GenerateUuidLower();
	}
	configManager->setValue(kConfigUserId, userId);
	if (TrimAscii(configManager->getValue(kConfigDeviceId)).empty()) {
		configManager->setValue(kConfigDeviceId, userId);
	}
	return userId;
}

std::string LoadOrCreateDeviceId(ConfigManager* configManager, const std::string& fallbackUserId)
{
	if (configManager == nullptr) {
		return fallbackUserId.empty() ? GenerateUuidLower() : fallbackUserId;
	}

	std::string deviceId = TrimAscii(configManager->getValue(kConfigDeviceId));
	if (deviceId.empty()) {
		deviceId = fallbackUserId.empty() ? GenerateUuidLower() : fallbackUserId;
		configManager->setValue(kConfigDeviceId, deviceId);
	}
	if (TrimAscii(configManager->getValue(kConfigUserId)).empty()) {
		configManager->setValue(kConfigUserId, deviceId);
	}
	return deviceId;
}

std::string BuildDeviceLabel(const std::string& deviceId)
{
	if (deviceId.empty()) {
		return kDevice;
	}
	std::string label = kDevice;
	label += "-";
	label += deviceId.substr(0, (std::min<size_t>)(deviceId.size(), 36));
	return label;
}

bool ParseRemoteConfigExampleOk()
{
	const nlohmann::json sample = {
		{"server_ts", 1546300000},
		{"configs_hash", "124595093888900"},
		{"ab_id", "aaabc"},
		{"ab_variant_id", "1"},
		{"configs", nlohmann::json::array({
			{{"key", "feature_enabled"}, {"value", "true"}, {"type", "string"}, {"end_ts", nullptr}},
			{{"key", "sample_rate"}, {"value", "50"}, {"type", "number"}},
			{{"key", "NEWS-LINK"}, {"value", R"([{"title":"AutoLinker","url":"https://github.com/aiqinxuancai/AutoLinker"}])"}, {"type", "json"}}
		})}
	};
	UpdateRemoteSnapshotSuccess(201, sample);
	const auto snapshot = GameAnalyticsClient::GetRemoteConfigs();
	return snapshot.ready &&
		snapshot.httpStatus == 201 &&
		snapshot.configsHash == "124595093888900" &&
		snapshot.abId == "aaabc" &&
		snapshot.abVariantId == "1" &&
		snapshot.values.size() == 3 &&
		snapshot.values.at("feature_enabled") == "true" &&
		snapshot.values.at("sample_rate") == "50" &&
		snapshot.values.at("NEWS-LINK").find("AutoLinker") != std::string::npos;
}

bool RemoteConfigRequestErrorPreservedOk()
{
	const std::string transportError =
		"HttpSendRequest failed, wininet_error=12029 ERROR_INTERNET_CANNOT_CONNECT";
	return BuildRemoteConfigRequestError(0, transportError) == transportError &&
		BuildRemoteConfigRequestError(503, "Service Unavailable") ==
		"remote config HTTP 503: Service Unavailable";
}

bool StartShowIdeRemoteConfigParseOk()
{
	std::string error;
	const std::vector<std::string> messages = ParseStartShowIdeMessages(
		R"(["\u5185\u5bb91","\u5185\u5bb92"])",
		error);
	if (!error.empty() || messages.size() != 2 || messages[0].empty() || messages[1].empty()) {
		return false;
	}

	const std::vector<std::string> invalid = ParseStartShowIdeMessages(R"({"text":"bad"})", error);
	return invalid.empty() && !error.empty();
}

nlohmann::json BuildEventBatchStateMachineSelfTest()
{
	using Batcher = GameAnalyticsEventBatcher;
	const Batcher::Policy policy = BuildProductionEventBatchPolicy();
	const Batcher::TimePoint start(Batcher::Duration(100000));
	const auto makeEvent = [](int sequence) {
		return nlohmann::json{{"sequence", sequence}};
	};

	nlohmann::json report;
	report["policy"] = {
		{"flush_threshold", policy.flushThreshold},
		{"max_pending_events", policy.maxPendingEvents},
		{"max_submit_failures", policy.maxSubmitFailures},
		{"flush_interval_ms", policy.flushInterval.count()},
		{"retry_1_ms", Batcher::RetryDelayForFailureCount(policy, 1).count()},
		{"retry_2_ms", Batcher::RetryDelayForFailureCount(policy, 2).count()}
	};
	const bool policyOk =
		policy.flushThreshold == 20 &&
		policy.maxPendingEvents == 128 &&
		policy.maxSubmitFailures == 3 &&
		policy.flushInterval == std::chrono::seconds(30) &&
		Batcher::RetryDelayForFailureCount(policy, 1) == std::chrono::seconds(1) &&
		Batcher::RetryDelayForFailureCount(policy, 2) == std::chrono::seconds(2);

	Batcher thresholdBatcher(policy);
	for (int i = 0; i < 19; ++i) {
		thresholdBatcher.Enqueue(makeEvent(i), start);
	}
	const bool belowThresholdWaits = !thresholdBatcher.IsReady(start) &&
		thresholdBatcher.NextDeadline() == start + std::chrono::seconds(30);
	const auto thresholdEnqueue = thresholdBatcher.Enqueue(makeEvent(19), start);
	const auto thresholdBatch = thresholdBatcher.TakeBatch(start, false, false);
	const bool countThresholdOk = thresholdEnqueue.readyForImmediateFlush &&
		thresholdBatch.has_value() &&
		thresholdBatch->events.size() == 20 &&
		thresholdBatcher.Empty();

	Batcher timerBatcher(policy);
	timerBatcher.Enqueue(makeEvent(1), start);
	timerBatcher.Enqueue(makeEvent(2), start + std::chrono::seconds(5));
	const bool timerNotEarly =
		!timerBatcher.IsReady(start + std::chrono::seconds(29)) &&
		!timerBatcher.TakeBatch(start + std::chrono::seconds(29), false, false).has_value();
	const auto timerBatch = timerBatcher.TakeBatch(start + std::chrono::seconds(30), false, false);
	const bool timeThresholdOk = timerNotEarly &&
		timerBatch.has_value() &&
		timerBatch->events.size() == 2;

	Batcher retrySuccessBatcher(policy);
	for (int i = 0; i < 20; ++i) {
		retrySuccessBatcher.Enqueue(makeEvent(i), start);
	}
	auto retryBatch = retrySuccessBatcher.TakeBatch(start, false, false);
	bool retryThenSuccessOk = retryBatch.has_value();
	if (retryBatch.has_value()) {
		const auto firstFailure = retrySuccessBatcher.CompleteFailure(std::move(*retryBatch), start);
		retryThenSuccessOk = retryThenSuccessOk &&
			firstFailure.disposition == Batcher::FailureDisposition::RetryScheduled &&
			firstFailure.failureCount == 1 &&
			firstFailure.retryDelay == std::chrono::seconds(1) &&
			!retrySuccessBatcher.IsReady(start + std::chrono::milliseconds(999));
		retryBatch = retrySuccessBatcher.TakeBatch(start + std::chrono::seconds(1), false, false);
		retryThenSuccessOk = retryThenSuccessOk &&
			retryBatch.has_value() &&
			retryBatch->failureCount == 1 &&
			retryBatch->events.size() == 20;
		if (retryBatch.has_value()) {
			const auto secondFailure = retrySuccessBatcher.CompleteFailure(
				std::move(*retryBatch),
				start + std::chrono::seconds(1));
			retryThenSuccessOk = retryThenSuccessOk &&
				secondFailure.disposition == Batcher::FailureDisposition::RetryScheduled &&
				secondFailure.failureCount == 2 &&
				secondFailure.retryDelay == std::chrono::seconds(2) &&
				!retrySuccessBatcher.IsReady(start + std::chrono::milliseconds(2999));
			retryBatch = retrySuccessBatcher.TakeBatch(start + std::chrono::seconds(3), false, false);
			retryThenSuccessOk = retryThenSuccessOk &&
				retryBatch.has_value() &&
				retryBatch->failureCount == 2 &&
				retrySuccessBatcher.Empty();
		}
	}

	Batcher dropBatcher(policy);
	for (int i = 0; i < 20; ++i) {
		dropBatcher.Enqueue(makeEvent(i), start);
	}
	auto droppedBatch = dropBatcher.TakeBatch(start, false, false);
	bool thirdFailureDropsOk = droppedBatch.has_value();
	for (int failure = 1; failure <= 3 && droppedBatch.has_value(); ++failure) {
		const auto failureResult = dropBatcher.CompleteFailure(
			std::move(*droppedBatch),
			start + std::chrono::seconds(failure == 1 ? 0 : failure == 2 ? 1 : 3));
		if (failure < 3) {
			thirdFailureDropsOk = thirdFailureDropsOk &&
				failureResult.disposition == Batcher::FailureDisposition::RetryScheduled;
			droppedBatch = dropBatcher.TakeBatch(
				start + std::chrono::seconds(failure == 1 ? 1 : 3),
				false,
				false);
		}
		else {
			thirdFailureDropsOk = thirdFailureDropsOk &&
				failureResult.disposition == Batcher::FailureDisposition::Dropped &&
				failureResult.failureCount == 3 &&
				dropBatcher.Empty();
		}
	}

	Batcher isolationBatcher(policy);
	for (int i = 0; i < 20; ++i) {
		isolationBatcher.Enqueue(makeEvent(i), start);
	}
	auto originalBatch = isolationBatcher.TakeBatch(start, false, false);
	const auto newEventTime = start + std::chrono::milliseconds(100);
	isolationBatcher.Enqueue(makeEvent(100), newEventTime);
	bool isolatedRetryBatchOk = originalBatch.has_value();
	if (originalBatch.has_value()) {
		isolationBatcher.CompleteFailure(
			std::move(*originalBatch),
			start + std::chrono::milliseconds(200));
		auto isolatedRetry = isolationBatcher.TakeBatch(
			start + std::chrono::milliseconds(1200),
			false,
			false);
		isolatedRetryBatchOk = isolatedRetryBatchOk &&
			isolatedRetry.has_value() &&
			isolatedRetry->events.size() == 20 &&
			isolationBatcher.PendingCount() == 1 &&
			isolationBatcher.NextDeadline() == newEventTime + std::chrono::seconds(30) &&
			!isolationBatcher.IsReady(newEventTime + std::chrono::seconds(29));
		const auto isolatedNewBatch = isolationBatcher.TakeBatch(
			newEventTime + std::chrono::seconds(30),
			false,
			false);
		isolatedRetryBatchOk = isolatedRetryBatchOk &&
			isolatedNewBatch.has_value() &&
			isolatedNewBatch->events.size() == 1 &&
			isolatedNewBatch->events[0].value("sequence", -1) == 100;
	}

	Batcher capacityBatcher(policy);
	bool capacityDroppedOldest = false;
	for (int i = 0; i < 129; ++i) {
		capacityDroppedOldest = capacityBatcher.Enqueue(makeEvent(i), start).droppedOldestEvent ||
			capacityDroppedOldest;
	}
	const auto capacityBatch = capacityBatcher.TakeBatch(start, false, true);
	const bool capacityLimitOk = capacityDroppedOldest &&
		capacityBatch.has_value() &&
		capacityBatch->events.size() == 128 &&
		capacityBatch->events.front().value("sequence", -1) == 1 &&
		capacityBatch->events.back().value("sequence", -1) == 128;

	Batcher shutdownBatcher(policy);
	shutdownBatcher.Enqueue(makeEvent(7), start);
	const auto shutdownBatch = shutdownBatcher.TakeBatch(start, true, true);
	const bool shutdownForceFlushOk = shutdownBatch.has_value() &&
		shutdownBatch->events.size() == 1 &&
		shutdownBatch->events[0].value("sequence", -1) == 7;

	report["policy_ok"] = policyOk;
	report["below_threshold_waits"] = belowThresholdWaits;
	report["count_threshold_flush"] = countThresholdOk;
	report["time_threshold_flush"] = timeThresholdOk;
	report["retry_twice_then_success"] = retryThenSuccessOk;
	report["third_failure_drops"] = thirdFailureDropsOk;
	report["retry_batch_isolation"] = isolatedRetryBatchOk;
	report["capacity_drops_oldest"] = capacityLimitOk;
	report["shutdown_force_flush"] = shutdownForceFlushOk;
	report["ok"] = policyOk &&
		belowThresholdWaits &&
		countThresholdOk &&
		timeThresholdOk &&
		retryThenSuccessOk &&
		thirdFailureDropsOk &&
		isolatedRetryBatchOk &&
		capacityLimitOk &&
		shutdownForceFlushOk;
	return report;
}

bool LifecycleEventFormatOk()
{
	std::lock_guard<std::mutex> lock(g_mutex);
	const ClientState saved = g_state;
	g_state = ClientState();
	g_state.userId = "self-test-user";
	g_state.deviceId = "windows-self-test-device";
	g_state.sessionId = GenerateUuidLower();
	g_state.sessionNum = 1;
	g_state.osVersion = "windows 10.0.0";
	g_state.sessionStartUnix = UnixTimeSeconds() - 3;
	g_state.remoteConfigs.values["NEWS-LINK"] = R"([{"title":"AutoLinker"}])";

	const nlohmann::json startup = BuildStartupEventLocked();
	const nlohmann::json sessionEnd = BuildSessionEndEventLocked();
	const nlohmann::json design = BuildDesignEventLocked(
		"AIChat:RunCompleted:TotalTokens",
		18640.0);
	g_state = saved;

	return startup.value("category", "") == "user" &&
		startup.contains("event_uuid") &&
		!startup.contains("event_id") &&
		sessionEnd.value("category", "") == "session_end" &&
		sessionEnd.contains("length") &&
		sessionEnd["length"].is_number_integer() &&
		!sessionEnd.contains("event_id") &&
		design.value("category", "") == "design" &&
		design.value("event_id", "") == "AIChat:RunCompleted:TotalTokens" &&
		design.value("value", 0.0) == 18640.0;
}

bool ValidateStoredGzipForSelfTest(const std::string& original, const std::string& gzip)
{
	if (gzip.size() < 23 ||
		static_cast<unsigned char>(gzip[0]) != 0x1F ||
		static_cast<unsigned char>(gzip[1]) != 0x8B ||
		static_cast<unsigned char>(gzip[2]) != 0x08) {
		return false;
	}

	size_t offset = 10;
	std::string restored;
	for (;;) {
		if (offset + 5 > gzip.size()) {
			return false;
		}
		const unsigned char header = static_cast<unsigned char>(gzip[offset++]);
		if ((header & 0x06U) != 0) {
			return false;
		}
		const bool finalBlock = (header & 0x01U) != 0;
		const uint16_t len =
			static_cast<uint16_t>(static_cast<unsigned char>(gzip[offset]) |
				(static_cast<unsigned char>(gzip[offset + 1]) << 8U));
		offset += 2;
		const uint16_t nlen =
			static_cast<uint16_t>(static_cast<unsigned char>(gzip[offset]) |
				(static_cast<unsigned char>(gzip[offset + 1]) << 8U));
		offset += 2;
		if (static_cast<uint16_t>(~len) != nlen || offset + len > gzip.size()) {
			return false;
		}
		restored.append(gzip.data() + offset, len);
		offset += len;
		if (finalBlock) {
			break;
		}
	}
	if (offset + 8 != gzip.size()) {
		return false;
	}
	const uint32_t crc =
		static_cast<uint32_t>(static_cast<unsigned char>(gzip[offset]) |
			(static_cast<unsigned char>(gzip[offset + 1]) << 8U) |
			(static_cast<unsigned char>(gzip[offset + 2]) << 16U) |
			(static_cast<unsigned char>(gzip[offset + 3]) << 24U));
	const uint32_t size =
		static_cast<uint32_t>(static_cast<unsigned char>(gzip[offset + 4]) |
			(static_cast<unsigned char>(gzip[offset + 5]) << 8U) |
			(static_cast<unsigned char>(gzip[offset + 6]) << 16U) |
			(static_cast<unsigned char>(gzip[offset + 7]) << 24U));
	return restored == original &&
		crc == Crc32(reinterpret_cast<const unsigned char*>(original.data()), original.size()) &&
		size == static_cast<uint32_t>(original.size());
}

} // namespace

namespace GameAnalyticsClient {

void Initialize(ConfigManager* configManager)
{
	std::lock_guard<std::mutex> lock(g_mutex);
	if (g_worker.joinable()) {
		return;
	}

	const bool enabled = configManager == nullptr || !IsExplicitFalse(configManager->getValue(kConfigEnabled));
	g_state = ClientState();
	g_state.configManager = configManager;
	g_state.enabled = enabled;
	if (!enabled) {
		LogGA("disabled by config");
		return;
	}

	g_state.userId = LoadOrCreateUserId(configManager);
	g_state.deviceId = BuildDeviceLabel(LoadOrCreateDeviceId(configManager, g_state.userId));
	g_state.sessionNum = LoadAndIncrementSessionNum(configManager);
	g_state.sessionId = GenerateUuidLower();
	g_state.osVersion = GetWindowsVersionString();
	g_state.sessionStartUnix = UnixTimeSeconds();
	g_state.initialized = true;
	g_state.tasks.push_back(WorkerTask::Init);

	g_worker = std::thread(WorkerMain);
	g_cv.notify_one();
	LogGA("initialized");
}

void Shutdown()
{
	std::thread worker;
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		if (!g_worker.joinable()) {
			return;
		}
		if (g_state.initialized && g_state.enabled && !g_state.shuttingDown) {
			if (!g_state.startupEventQueued) {
				PushEventLocked(BuildStartupEventLocked());
				g_state.startupEventQueued = true;
			}
			PushEventLocked(BuildSessionEndEventLocked());
			g_state.tasks.push_back(WorkerTask::FlushEvents);
		}
		g_state.shuttingDown = true;
		g_cv.notify_one();
		worker = std::move(g_worker);
	}
	if (worker.joinable()) {
		worker.join();
	}
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		g_state = ClientState();
	}
}

bool IsRunning()
{
	std::lock_guard<std::mutex> lock(g_mutex);
	return g_worker.joinable() && g_state.initialized && !g_state.shuttingDown;
}

void QueueDesignEvent(const std::string& eventId, double value)
{
	if (eventId.empty() || !std::isfinite(value)) {
		return;
	}
	std::lock_guard<std::mutex> lock(g_mutex);
	if (!g_state.initialized || !g_state.enabled || g_state.shuttingDown) {
		return;
	}
	PushEventLocked(BuildDesignEventLocked(eventId, value));
	ScheduleEventFlushLocked();
}

void RefreshRemoteConfigsAsync()
{
	std::lock_guard<std::mutex> lock(g_mutex);
	if (!g_state.initialized || !g_state.enabled || g_state.shuttingDown) {
		return;
	}
	QueueTaskLocked(WorkerTask::RefreshRemoteConfigs);
}

RemoteConfigSnapshot GetRemoteConfigs()
{
	std::lock_guard<std::mutex> lock(g_mutex);
	return g_state.remoteConfigs;
}

std::optional<std::string> GetRemoteConfigValue(const std::string& key)
{
	std::lock_guard<std::mutex> lock(g_mutex);
	const auto it = g_state.remoteConfigs.values.find(key);
	if (it == g_state.remoteConfigs.values.end()) {
		return std::nullopt;
	}
	return it->second;
}

nlohmann::json BuildWorkerIntegrationSelfTest()
{
	using Batcher = GameAnalyticsEventBatcher;
	nlohmann::json report = {{"ok", false}};
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		if (g_worker.joinable()) {
			report["error"] = "GameAnalytics worker is already running";
			return report;
		}
	}

	std::mutex transportMutex;
	std::condition_variable transportCv;
	std::vector<nlohmann::json> submittedBatches;
	std::vector<Batcher::TimePoint> attemptTimes;
	const std::vector<int> responseStatuses = {200, 200, 500, 500, 500, 200, 200};
	bool allPayloadsValid = true;
	g_eventSubmitOverride = [&](const nlohmann::json& events,
		const std::string& gzip,
		const std::string& headers) {
		std::lock_guard<std::mutex> lock(transportMutex);
		const size_t attempt = submittedBatches.size();
		submittedBatches.push_back(events);
		attemptTimes.push_back(Batcher::Clock::now());
		allPayloadsValid = allPayloadsValid &&
			events.is_array() &&
			ValidateStoredGzipForSelfTest(events.dump(), gzip) &&
			headers.find("Authorization: ") != std::string::npos &&
			headers.find("Content-Encoding: gzip") != std::string::npos;
		transportCv.notify_all();
		return std::make_pair(
			std::string("{}"),
			attempt < responseStatuses.size() ? responseStatuses[attempt] : 200);
	};

	const Batcher::Policy testPolicy = {
		3,
		32,
		3,
		std::chrono::milliseconds(120),
		std::chrono::milliseconds(50)
	};
	{
		std::lock_guard<std::mutex> lock(g_mutex);
		g_state = ClientState();
		g_state.enabled = true;
		g_state.initialized = true;
		g_state.collectorReady = true;
		g_state.collectorEnabled = true;
		g_state.userId = "worker-self-test-user";
		g_state.deviceId = "worker-self-test-device";
		g_state.sessionId = GenerateUuidLower();
		g_state.sessionNum = 1;
		g_state.osVersion = "windows 10.0.0";
		g_state.sessionStartUnix = UnixTimeSeconds();
		g_state.eventBatcher = Batcher(testPolicy);
		g_worker = std::thread(WorkerMain);
	}

	const auto waitForAttempts = [&](size_t count, std::chrono::milliseconds timeout) {
		std::unique_lock<std::mutex> lock(transportMutex);
		return transportCv.wait_for(lock, timeout, [&]() {
			return submittedBatches.size() >= count;
		});
	};
	const auto containsEventId = [](const nlohmann::json& batch, const std::string& eventId) {
		if (!batch.is_array()) {
			return false;
		}
		return std::any_of(batch.begin(), batch.end(), [&](const nlohmann::json& event) {
			return event.is_object() && event.value("event_id", std::string()) == eventId;
		});
	};

	QueueDesignEvent("SelfTest:Threshold:1", 1.0);
	QueueDesignEvent("SelfTest:Threshold:2", 2.0);
	QueueDesignEvent("SelfTest:Threshold:3", 3.0);
	const bool thresholdAttemptArrived = waitForAttempts(1, std::chrono::milliseconds(600));

	const auto timerQueuedAt = Batcher::Clock::now();
	QueueDesignEvent("SelfTest:Timer", 4.0);
	const bool timerDidNotFireEarly = !waitForAttempts(2, std::chrono::milliseconds(60));
	const bool timerAttemptArrived = waitForAttempts(2, std::chrono::milliseconds(600));

	QueueDesignEvent("SelfTest:Retry:1", 5.0);
	QueueDesignEvent("SelfTest:Retry:2", 6.0);
	QueueDesignEvent("SelfTest:Retry:3", 7.0);
	const bool firstFailureArrived = waitForAttempts(3, std::chrono::milliseconds(600));
	const auto isolatedQueuedAt = Batcher::Clock::now();
	QueueDesignEvent("SelfTest:Isolated", 8.0);
	const bool retryAndIsolationArrived = waitForAttempts(6, std::chrono::milliseconds(1200));

	QueueDesignEvent("SelfTest:Shutdown", 9.0);
	Shutdown();
	const bool shutdownAttemptArrived = waitForAttempts(7, std::chrono::milliseconds(200));
	g_eventSubmitOverride = {};

	std::vector<nlohmann::json> batches;
	std::vector<Batcher::TimePoint> times;
	bool payloadsValid = false;
	{
		std::lock_guard<std::mutex> lock(transportMutex);
		batches = submittedBatches;
		times = attemptTimes;
		payloadsValid = allPayloadsValid;
	}
	const auto elapsedMs = [](Batcher::TimePoint from, Batcher::TimePoint to) {
		return std::chrono::duration_cast<std::chrono::milliseconds>(to - from).count();
	};
	const long long timerDelayMs = times.size() > 1 ? elapsedMs(timerQueuedAt, times[1]) : -1;
	const long long retryOneDelayMs = times.size() > 3 ? elapsedMs(times[2], times[3]) : -1;
	const long long retryTwoDelayMs = times.size() > 4 ? elapsedMs(times[3], times[4]) : -1;
	const long long isolatedDelayMs = times.size() > 5 ? elapsedMs(isolatedQueuedAt, times[5]) : -1;

	const bool thresholdWorkerOk = thresholdAttemptArrived &&
		batches.size() > 0 &&
		batches[0].size() == 3 &&
		containsEventId(batches[0], "SelfTest:Threshold:1") &&
		containsEventId(batches[0], "SelfTest:Threshold:3");
	const bool timerWorkerOk = timerDidNotFireEarly &&
		timerAttemptArrived &&
		batches.size() > 1 &&
		batches[1].size() == 1 &&
		containsEventId(batches[1], "SelfTest:Timer") &&
		timerDelayMs >= 90;
	const bool retryWorkerOk = firstFailureArrived &&
		retryAndIsolationArrived &&
		batches.size() > 5 &&
		batches[2].size() == 3 &&
		batches[2] == batches[3] &&
		batches[3] == batches[4] &&
		retryOneDelayMs >= 35 &&
		retryTwoDelayMs >= 80;
	const bool thirdFailureDropWorkerOk = retryWorkerOk &&
		batches[5].size() == 1 &&
		containsEventId(batches[5], "SelfTest:Isolated") &&
		!containsEventId(batches[5], "SelfTest:Retry:1");
	const bool isolationWorkerOk = thirdFailureDropWorkerOk &&
		isolatedDelayMs >= 100 &&
		isolatedDelayMs <= 260;
	const bool shutdownWorkerOk = shutdownAttemptArrived &&
		batches.size() == 7 &&
		batches[6].size() == 3 &&
		containsEventId(batches[6], "SelfTest:Shutdown") &&
		std::any_of(batches[6].begin(), batches[6].end(), [](const nlohmann::json& event) {
			return event.value("category", std::string()) == "session_end";
		});

	report["threshold_flush"] = thresholdWorkerOk;
	report["timer_flush"] = timerWorkerOk;
	report["retry_backoff"] = retryWorkerOk;
	report["third_failure_drops"] = thirdFailureDropWorkerOk;
	report["retry_batch_isolation"] = isolationWorkerOk;
	report["shutdown_force_flush"] = shutdownWorkerOk;
	report["payload_gzip_hmac"] = payloadsValid;
	report["attempt_count"] = batches.size();
	report["timer_delay_ms"] = timerDelayMs;
	report["retry_1_delay_ms"] = retryOneDelayMs;
	report["retry_2_delay_ms"] = retryTwoDelayMs;
	report["isolated_event_delay_ms"] = isolatedDelayMs;
	report["ok"] = thresholdWorkerOk &&
		timerWorkerOk &&
		retryWorkerOk &&
		thirdFailureDropWorkerOk &&
		isolationWorkerOk &&
		shutdownWorkerOk &&
		payloadsValid;
	return report;
}

std::string BuildSelfTestReportJson()
{
	nlohmann::json report;
	const std::string hmac = HmacSha256Base64(
		R"({"test": "test"})",
		"16813a12f718bc5c620f56944e1abc3ea13ccbac");
	report["hmac_sha256_base64_ok"] = hmac == "slnR8CKJtKtFDaESSrqnqQeUvp5FaVV7d5XHxt50N5A=";
	report["hmac_sha256_base64"] = hmac;

	const std::string gzipOriginal = R"([{"category":"user"}])";
	const std::string gzip = GzipStoredDeflate(gzipOriginal);
	report["gzip_stored_deflate_ok"] = ValidateStoredGzipForSelfTest(gzipOriginal, gzip);
	report["gzip_payload_size"] = gzip.size();

	{
		std::lock_guard<std::mutex> lock(g_mutex);
		g_state.remoteConfigs = RemoteConfigSnapshot();
	}
	report["remote_config_parse_ok"] = ParseRemoteConfigExampleOk();
	report["remote_config_request_error_preserved_ok"] = RemoteConfigRequestErrorPreservedOk();
	report["start_show_ide_parse_ok"] = StartShowIdeRemoteConfigParseOk();
	report["lifecycle_event_format_ok"] = LifecycleEventFormatOk();
	report["event_batch_state_machine"] = BuildEventBatchStateMachineSelfTest();
	report["event_batch_policy_ok"] = report["event_batch_state_machine"]["ok"];
	report["event_worker_integration"] = BuildWorkerIntegrationSelfTest();
	report["ok"] =
		report["hmac_sha256_base64_ok"].get<bool>() &&
		report["gzip_stored_deflate_ok"].get<bool>() &&
		report["remote_config_parse_ok"].get<bool>() &&
		report["remote_config_request_error_preserved_ok"].get<bool>() &&
		report["start_show_ide_parse_ok"].get<bool>() &&
		report["lifecycle_event_format_ok"].get<bool>() &&
		report["event_batch_policy_ok"].get<bool>() &&
		report["event_worker_integration"]["ok"].get<bool>();
	return report.dump(2);
}

std::string BuildLiveBatchTestReportJson(
	const std::string& gameKey,
	const std::string& secretKey)
{
	nlohmann::json report = {
		{"ok", false},
		{"step", "validate_credentials"}
	};
	const bool gameKeyValid = gameKey.size() >= 8 && gameKey.size() <= 128 &&
		std::all_of(gameKey.begin(), gameKey.end(), [](unsigned char ch) {
			return std::isalnum(ch) != 0 || ch == '-' || ch == '_';
		});
	if (!gameKeyValid || secretKey.empty()) {
		report["error"] = "invalid GameAnalytics test credentials";
		return report.dump(2);
	}

	report["step"] = "collector_init";
	const std::string osVersion = GetWindowsVersionString();
	const nlohmann::json initBodyJson = {
		{"platform", kPlatform},
		{"os_version", osVersion},
		{"sdk_version", kSdkVersion}
	};
	const std::string initBody = initBodyJson.dump();
	const std::string initHeaders = BuildJsonHeaders(initBody, secretKey);
	const auto initResponse = PerformPostRequest(
		std::format("{}/v2/{}/init", kApiBaseUrl, gameKey),
		initBody,
		initHeaders,
		kHttpTimeoutMs,
		true,
		true);
	report["collector_init"] = {
		{"http_status", initResponse.second},
		{"authorization_present", initHeaders.find("Authorization: ") != std::string::npos},
		{"response_size", initResponse.first.size()}
	};
	if (initResponse.second != 200) {
		report["error"] = "GameAnalytics collector init failed";
		return report.dump(2);
	}

	long long clientTs = UnixTimeSeconds();
	bool collectorEnabled = true;
	try {
		const nlohmann::json parsed = nlohmann::json::parse(initResponse.first);
		if (parsed.contains("server_ts") && parsed["server_ts"].is_number_integer()) {
			clientTs = parsed["server_ts"].get<long long>();
		}
		collectorEnabled = parsed.value("enabled", true);
	}
	catch (...) {
	}
	report["collector_init"]["enabled"] = collectorEnabled;
	if (!collectorEnabled) {
		report["error"] = "GameAnalytics collector disabled this game";
		return report.dump(2);
	}

	report["step"] = "batch_events";
	const std::string sessionId = GenerateUuidLower();
	const std::string eventId = "AutoLinkerTest:BatchSmoke:" + GenerateUuidLower();
	nlohmann::json events = nlohmann::json::array();
	for (int i = 0; i < static_cast<int>(kFlushEventThreshold); ++i) {
		events.push_back({
			{"v", 2},
			{"event_uuid", GenerateUuidLower()},
			{"user_id", "autolinker-live-test"},
			{"client_ts", clientTs},
			{"sdk_version", kSdkVersion},
			{"os_version", osVersion},
			{"manufacturer", kManufacturer},
			{"device", "windows-live-test"},
			{"platform", kPlatform},
			{"session_id", sessionId},
			{"session_num", 1},
			{"build", BuildVersionString()},
			{"category", "design"},
			{"event_id", eventId},
			{"value", i + 1}
		});
	}
	const std::string eventsBody = events.dump();
	const std::string gzippedBody = GzipStoredDeflate(eventsBody);
	const std::string eventHeaders = BuildGzipJsonHeaders(gzippedBody, secretKey);
	const auto eventResponse = PerformPostRequest(
		std::format("{}/v2/{}/events", kApiBaseUrl, gameKey),
		gzippedBody,
		eventHeaders,
		kHttpTimeoutMs,
		true,
		true);
	report["batch_events"] = {
		{"http_status", eventResponse.second},
		{"event_count", events.size()},
		{"event_id", eventId},
		{"json_bytes", eventsBody.size()},
		{"gzip_bytes", gzippedBody.size()},
		{"gzip_round_trip", ValidateStoredGzipForSelfTest(eventsBody, gzippedBody)},
		{"authorization_present", eventHeaders.find("Authorization: ") != std::string::npos},
		{"response_size", eventResponse.first.size()},
		{"response_preview", eventResponse.first.substr(0, 300)}
	};
	report["step"] = "completed";
	report["ok"] = eventResponse.second == 200 &&
		events.size() == kFlushEventThreshold &&
		report["batch_events"]["gzip_round_trip"].get<bool>() &&
		report["batch_events"]["authorization_present"].get<bool>();
	if (!report["ok"].get<bool>()) {
		report["error"] = "GameAnalytics batch event submission failed";
	}
	return report.dump(2);
}

} // namespace GameAnalyticsClient
