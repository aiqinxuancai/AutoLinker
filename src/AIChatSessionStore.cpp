#include "AIChatSessionStore.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <ctime>
#include <format>
#include <fstream>
#include <system_error>

#include "..\\thirdparty\\json.hpp"

#include "PathHelper.h"

namespace {

bool IsValidUtf8TextForSessionStore(const std::string& text)
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

std::string ConvertCodePageForSessionStore(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
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

std::string LocalToUtf8TextForSessionStore(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8TextForSessionStore(text)) {
		return text;
	}
	return ConvertCodePageForSessionStore(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocalTextForSessionStore(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8TextForSessionStore(text)) {
		return text;
	}
	return ConvertCodePageForSessionStore(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string GetJsonStringAsLocalText(const nlohmann::json& row, const char* key)
{
	if (!row.is_object() || key == nullptr || !row.contains(key) || !row[key].is_string()) {
		return std::string();
	}
	return Utf8ToLocalTextForSessionStore(row[key].get<std::string>());
}

std::string GetJsonStringUtf8(const nlohmann::json& row, const char* key)
{
	if (!row.is_object() || key == nullptr || !row.contains(key) || !row[key].is_string()) {
		return std::string();
	}
	return row[key].get<std::string>();
}

bool GetJsonBool(const nlohmann::json& row, const char* key, bool defaultValue)
{
	if (!row.is_object() || key == nullptr || !row.contains(key) || !row[key].is_boolean()) {
		return defaultValue;
	}
	return row[key].get<bool>();
}

long long GetJsonInt64(const nlohmann::json& row, const char* key, long long defaultValue)
{
	if (!row.is_object() || key == nullptr || !row.contains(key) || !row[key].is_number_integer()) {
		return defaultValue;
	}
	return row[key].get<long long>();
}

std::string BuildTimestampDisplayLocal(long long unixMs)
{
	if (unixMs <= 0) {
		return std::string();
	}

	const time_t unixSeconds = static_cast<time_t>(unixMs / 1000);
	std::tm localTm = {};
#if defined(_MSC_VER)
	if (localtime_s(&localTm, &unixSeconds) != 0) {
		return std::string();
	}
#else
	const std::tm* ptr = std::localtime(&unixSeconds);
	if (ptr == nullptr) {
		return std::string();
	}
	localTm = *ptr;
#endif

	char buffer[64] = {};
	if (std::strftime(buffer, sizeof(buffer), "%Y-%m-%d %H:%M:%S", &localTm) == 0) {
		return std::string();
	}
	return std::string(buffer);
}

long long GetCurrentUnixTimeMs()
{
	const auto now = std::chrono::system_clock::now();
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count());
}

std::string SanitizeSessionIdFileName(const std::string& sessionId)
{
	std::string sanitized;
	sanitized.reserve(sessionId.size());
	for (char ch : sessionId) {
		const unsigned char uch = static_cast<unsigned char>(ch);
		if ((uch >= '0' && uch <= '9') ||
			(uch >= 'A' && uch <= 'Z') ||
			(uch >= 'a' && uch <= 'z') ||
			ch == '-' || ch == '_') {
			sanitized.push_back(ch);
		}
	}
	return sanitized.empty() ? std::string("session") : sanitized;
}

std::string BuildSessionTitleLocal(const AIChatStoredSession& session)
{
	for (auto it = session.messages.rbegin(); it != session.messages.rend(); ++it) {
		if (_stricmp(it->role.c_str(), "user") != 0) {
			continue;
		}
		std::string title = it->contentLocal;
		for (char& ch : title) {
			if (ch == '\r' || ch == '\n' || ch == '\t') {
				ch = ' ';
			}
		}
		while (!title.empty() && title.front() == ' ') {
			title.erase(title.begin());
		}
		while (!title.empty() && title.back() == ' ') {
			title.pop_back();
		}
		if (title.size() > 80) {
			title.resize(80);
			title += "...";
		}
		if (!title.empty()) {
			return title;
		}
	}
	return session.sourceFileNameLocal.empty()
		? std::string("未命名会话")
		: ("[" + session.sourceFileNameLocal + "] 会话");
}

bool SerializeSession(const AIChatStoredSession& session, nlohmann::json& outJson, std::string& outError)
{
	try {
		outJson = nlohmann::json::object();
		outJson["schema_version"] = session.schemaVersion;
		outJson["session_id"] = session.sessionId;
		outJson["source_file_name"] = LocalToUtf8TextForSessionStore(session.sourceFileNameLocal);
		outJson["source_file_path_hint"] = LocalToUtf8TextForSessionStore(session.sourceFilePathHintLocal);
		outJson["created_at_unix_ms"] = session.createdAtUnixMs;
		outJson["updated_at_unix_ms"] = session.updatedAtUnixMs;
		outJson["elapsed_ms"] = session.elapsedMs;
		outJson["created_at_display"] = LocalToUtf8TextForSessionStore(session.createdAtDisplayLocal);
		outJson["updated_at_display"] = LocalToUtf8TextForSessionStore(session.updatedAtDisplayLocal);
		outJson["rolling_summary"] = LocalToUtf8TextForSessionStore(session.rollingSummaryLocal);
		outJson["plan_mode_state"] = session.planModeState;
		outJson["pending_plan"] = LocalToUtf8TextForSessionStore(session.pendingPlanLocal);
		outJson["auto_allow_writes"] = session.autoAllowWrites;
		outJson["messages"] = nlohmann::json::array();

		for (const auto& message : session.messages) {
			nlohmann::json row = nlohmann::json::object();
			row["role"] = message.role;
			row["content"] = LocalToUtf8TextForSessionStore(message.contentLocal);
			row["include_in_context"] = message.includeInContext;
			row["visible_in_history"] = message.visibleInHistory;
			row["reasoning_content"] = message.reasoningContentUtf8;
			row["raw_message_json_utf8"] = message.rawMessageJsonUtf8;
			outJson["messages"].push_back(std::move(row));
		}
		return true;
	}
	catch (const std::exception& ex) {
		outError = ex.what();
		return false;
	}
}

bool DeserializeSession(const nlohmann::json& jsonValue, AIChatStoredSession& outSession, std::string& outError)
{
	if (!jsonValue.is_object()) {
		outError = "session json root is not object";
		return false;
	}

	outSession = {};
	outSession.schemaVersion = static_cast<int>(GetJsonInt64(jsonValue, "schema_version", 1));
	outSession.sessionId = GetJsonStringUtf8(jsonValue, "session_id");
	outSession.sourceFileNameLocal = GetJsonStringAsLocalText(jsonValue, "source_file_name");
	outSession.sourceFilePathHintLocal = GetJsonStringAsLocalText(jsonValue, "source_file_path_hint");
	outSession.createdAtUnixMs = GetJsonInt64(jsonValue, "created_at_unix_ms", 0);
	outSession.updatedAtUnixMs = GetJsonInt64(jsonValue, "updated_at_unix_ms", 0);
	outSession.elapsedMs = GetJsonInt64(jsonValue, "elapsed_ms", 0);
	if (outSession.elapsedMs < 0) {
		outSession.elapsedMs = 0;
	}
	outSession.createdAtDisplayLocal = GetJsonStringAsLocalText(jsonValue, "created_at_display");
	outSession.updatedAtDisplayLocal = GetJsonStringAsLocalText(jsonValue, "updated_at_display");
	outSession.rollingSummaryLocal = GetJsonStringAsLocalText(jsonValue, "rolling_summary");
	outSession.planModeState = GetJsonStringUtf8(jsonValue, "plan_mode_state");
	outSession.pendingPlanLocal = GetJsonStringAsLocalText(jsonValue, "pending_plan");
	outSession.autoAllowWrites = GetJsonBool(jsonValue, "auto_allow_writes", false);

	if (outSession.createdAtDisplayLocal.empty()) {
		outSession.createdAtDisplayLocal = BuildTimestampDisplayLocal(outSession.createdAtUnixMs);
	}
	if (outSession.updatedAtDisplayLocal.empty()) {
		outSession.updatedAtDisplayLocal = BuildTimestampDisplayLocal(outSession.updatedAtUnixMs);
	}

	if (!jsonValue.contains("messages") || !jsonValue["messages"].is_array()) {
		return true;
	}

	for (const auto& row : jsonValue["messages"]) {
		if (!row.is_object()) {
			continue;
		}
		AIChatStoredMessage message = {};
		message.role = GetJsonStringUtf8(row, "role");
		message.contentLocal = GetJsonStringAsLocalText(row, "content");
		message.includeInContext = GetJsonBool(row, "include_in_context", true);
		message.visibleInHistory = GetJsonBool(row, "visible_in_history", true);
		message.reasoningContentUtf8 = GetJsonStringUtf8(row, "reasoning_content");
		message.rawMessageJsonUtf8 = GetJsonStringUtf8(row, "raw_message_json_utf8");
		outSession.messages.push_back(std::move(message));
	}
	return true;
}

} // namespace

std::string CreateAIChatSessionId()
{
	SYSTEMTIME st = {};
	GetLocalTime(&st);
	return std::format(
		"{:04}{:02}{:02}_{:02}{:02}{:02}_{:03}_{}_{}",
		static_cast<int>(st.wYear),
		static_cast<int>(st.wMonth),
		static_cast<int>(st.wDay),
		static_cast<int>(st.wHour),
		static_cast<int>(st.wMinute),
		static_cast<int>(st.wSecond),
		static_cast<int>(st.wMilliseconds),
		static_cast<unsigned long long>(GetCurrentProcessId()),
		static_cast<unsigned long long>(GetTickCount64()));
}

std::filesystem::path GetAIChatSessionDirectoryPathForSourceFile(const std::string& sourceFilePathLocal)
{
	std::filesystem::path sourcePath(sourceFilePathLocal);
	std::string sourceName = sourcePath.filename().string();
	if (sourceName.empty()) {
		sourceName = "UnknownSource";
	}
	return GetAutoLinkerSessionRootDirectoryPath() / SanitizePathComponentForStorage(sourceName);
}

std::filesystem::path ResolveAIChatSessionFilePath(
	const std::string& sourceFilePathLocal,
	const std::string& sessionId)
{
	return GetAIChatSessionDirectoryPathForSourceFile(sourceFilePathLocal) /
		(SanitizeSessionIdFileName(sessionId) + ".json");
}

bool SaveAIChatStoredSession(const AIChatStoredSession& session, std::string* outError)
{
	if (session.sessionId.empty()) {
		if (outError != nullptr) {
			*outError = "session id is empty";
		}
		return false;
	}
	if (session.sessionFilePath.empty()) {
		if (outError != nullptr) {
			*outError = "session file path is empty";
		}
		return false;
	}

	nlohmann::json jsonValue;
	std::string error;
	if (!SerializeSession(session, jsonValue, error)) {
		if (outError != nullptr) {
			*outError = error;
		}
		return false;
	}

	std::error_code createEc;
	std::filesystem::create_directories(session.sessionFilePath.parent_path(), createEc);
	if (createEc) {
		if (outError != nullptr) {
			*outError = createEc.message();
		}
		return false;
	}

	std::ofstream out(session.sessionFilePath, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		if (outError != nullptr) {
			*outError = "failed to open session file";
		}
		return false;
	}

	std::string text;
	try {
		text = jsonValue.dump(2, ' ', false, nlohmann::json::error_handler_t::replace);
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) {
			*outError = ex.what();
		}
		return false;
	}

	out.write("\xEF\xBB\xBF", 3);
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	out.flush();
	if (!out.good()) {
		if (outError != nullptr) {
			*outError = "failed to write session file";
		}
		return false;
	}
	return true;
}

bool LoadAIChatStoredSession(
	const std::filesystem::path& sessionFilePath,
	AIChatStoredSession& outSession,
	std::string* outError)
{
	outSession = {};
	std::ifstream in(sessionFilePath, std::ios::binary);
	if (!in.is_open()) {
		if (outError != nullptr) {
			*outError = "failed to open session file";
		}
		return false;
	}

	std::string text((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		text.erase(0, 3);
	}

	nlohmann::json jsonValue;
	try {
		jsonValue = nlohmann::json::parse(text);
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) {
			*outError = ex.what();
		}
		return false;
	}

	std::string error;
	if (!DeserializeSession(jsonValue, outSession, error)) {
		if (outError != nullptr) {
			*outError = error;
		}
		return false;
	}
	outSession.sessionFilePath = sessionFilePath;
	return true;
}

std::vector<AIChatStoredSessionListEntry> ListRecentAIChatStoredSessions(
	const std::string& sourceFilePathLocal,
	size_t limit)
{
	std::vector<AIChatStoredSessionListEntry> out;
	if (limit == 0) {
		return out;
	}

	const std::filesystem::path dir = GetAIChatSessionDirectoryPathForSourceFile(sourceFilePathLocal);
	std::error_code existsEc;
	if (!std::filesystem::exists(dir, existsEc) || existsEc) {
		return out;
	}

	std::vector<AIChatStoredSessionListEntry> loaded;
	std::error_code iterEc;
	for (std::filesystem::directory_iterator it(dir, iterEc), end; it != end && !iterEc; it.increment(iterEc)) {
		if (!it->is_regular_file()) {
			continue;
		}
		AIChatStoredSession session;
		if (!LoadAIChatStoredSession(it->path(), session, nullptr)) {
			continue;
		}
		AIChatStoredSessionListEntry row = {};
		row.sessionId = session.sessionId;
		row.sessionFilePath = it->path();
		row.updatedAtUnixMs = session.updatedAtUnixMs;
		row.updatedAtDisplayLocal = session.updatedAtDisplayLocal;
		if (row.updatedAtDisplayLocal.empty()) {
			row.updatedAtDisplayLocal = BuildTimestampDisplayLocal(row.updatedAtUnixMs);
		}
		row.titleLocal = BuildSessionTitleLocal(session);
		loaded.push_back(std::move(row));
	}

	std::sort(loaded.begin(), loaded.end(), [](const AIChatStoredSessionListEntry& left, const AIChatStoredSessionListEntry& right) {
		if (left.updatedAtUnixMs != right.updatedAtUnixMs) {
			return left.updatedAtUnixMs > right.updatedAtUnixMs;
		}
		return _stricmp(left.sessionId.c_str(), right.sessionId.c_str()) < 0;
	});

	if (loaded.size() > limit) {
		loaded.resize(limit);
	}
	return loaded;
}
