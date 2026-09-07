#include "AIChatSessionStore.h"

#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <chrono>
#include <ctime>
#include <format>
#include <fstream>
#include <system_error>

#include "..\\thirdparty\\json.hpp"

#include "PathHelper.h"
#include "UnicodeTextCodec.h"

namespace {

std::string LocalToUtf8TextForSessionStore(const std::string& text)
{
	return UnicodeTextCodec::LocalToUtf8RestoringUnicode(text);
}

std::string Utf8ToLocalTextForSessionStore(const std::string& text)
{
	return UnicodeTextCodec::Utf8ToLocalPreservingUnicode(text);
}

std::filesystem::path LocalTextToPathForSessionStore(const std::string& text)
{
	const std::string utf8 = LocalToUtf8TextForSessionStore(text);
	const int length = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		utf8.data(),
		static_cast<int>(utf8.size()),
		nullptr,
		0);
	if (length <= 0) {
		return std::filesystem::path(text);
	}
	std::wstring wide(static_cast<size_t>(length), L'\0');
	if (MultiByteToWideChar(
			CP_UTF8,
			MB_ERR_INVALID_CHARS,
			utf8.data(),
			static_cast<int>(utf8.size()),
			wide.data(),
			length) <= 0) {
		return std::filesystem::path(text);
	}
	return std::filesystem::path(wide);
}

std::string PathToLocalTextForSessionStore(const std::filesystem::path& path)
{
	const std::wstring wide = path.wstring();
	const int length = WideCharToMultiByte(
		CP_UTF8,
		0,
		wide.data(),
		static_cast<int>(wide.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (length <= 0) {
		return path.string();
	}
	std::string utf8(static_cast<size_t>(length), '\0');
	if (WideCharToMultiByte(
			CP_UTF8,
			0,
			wide.data(),
			static_cast<int>(wide.size()),
			utf8.data(),
			length,
			nullptr,
			nullptr) <= 0) {
		return path.string();
	}
	return Utf8ToLocalTextForSessionStore(utf8);
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
		if (!it->visibleInHistory) {
			continue;
		}
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
	if (AIChatGoalManager::HasGoal(session.goal)) {
		std::string title = session.goal.objectiveLocal;
		for (char& ch : title) {
			if (ch == '\r' || ch == '\n' || ch == '\t') ch = ' ';
		}
		if (title.size() > 80) {
			title.resize(80);
			title += "...";
		}
		if (!title.empty()) return title;
	}
	return session.sourceFileNameLocal.empty()
		? std::string("未命名会话")
		: ("[" + session.sourceFileNameLocal + "] 会话");
}

nlohmann::json SerializeImageAttachments(
	const std::vector<AIImageAttachment>& attachments,
	const std::filesystem::path& sessionFilePath)
{
	nlohmann::json values = nlohmann::json::array();
	const std::filesystem::path sessionDirectory = sessionFilePath.parent_path();
	for (const AIImageAttachment& attachment : attachments) {
		std::filesystem::path storedPath = LocalTextToPathForSessionStore(attachment.assetPathLocal);
		if (!sessionDirectory.empty()) {
			std::error_code ec;
			const std::filesystem::path relative = std::filesystem::relative(storedPath, sessionDirectory, ec);
			if (!ec && !relative.empty()) {
				storedPath = relative;
			}
		}
		values.push_back({
			{"id", attachment.id},
			{"file_name", LocalToUtf8TextForSessionStore(attachment.fileNameLocal)},
			{"mime_type", attachment.mimeType},
			{"asset_path", LocalToUtf8TextForSessionStore(PathToLocalTextForSessionStore(storedPath))},
			{"source_path", LocalToUtf8TextForSessionStore(attachment.sourcePathLocal)},
			{"detail", attachment.detail},
			{"byte_size", attachment.byteSize},
			{"width", attachment.width},
			{"height", attachment.height}
		});
	}
	return values;
}

std::vector<AIImageAttachment> DeserializeImageAttachments(
	const nlohmann::json& owner,
	const std::filesystem::path& sessionFilePath)
{
	std::vector<AIImageAttachment> attachments;
	if (!owner.is_object() || !owner.contains("attachments") || !owner["attachments"].is_array()) {
		return attachments;
	}
	for (const auto& value : owner["attachments"]) {
		if (!value.is_object()) {
			continue;
		}
		AIImageAttachment attachment;
		attachment.id = GetJsonStringUtf8(value, "id");
		attachment.fileNameLocal = GetJsonStringAsLocalText(value, "file_name");
		attachment.mimeType = GetJsonStringUtf8(value, "mime_type");
		attachment.assetPathLocal = GetJsonStringAsLocalText(value, "asset_path");
		attachment.sourcePathLocal = GetJsonStringAsLocalText(value, "source_path");
		attachment.detail = GetJsonStringUtf8(value, "detail");
		std::transform(attachment.detail.begin(), attachment.detail.end(), attachment.detail.begin(), [](unsigned char ch) {
			return static_cast<char>(std::tolower(ch));
		});
		attachment.byteSize = static_cast<std::uint64_t>((std::max)(0LL, GetJsonInt64(value, "byte_size", 0)));
		attachment.width = static_cast<unsigned int>((std::max)(0LL, GetJsonInt64(value, "width", 0)));
		attachment.height = static_cast<unsigned int>((std::max)(0LL, GetJsonInt64(value, "height", 0)));
		if (attachment.detail != "low" && attachment.detail != "high") {
			attachment.detail = "auto";
		}
		if (!attachment.assetPathLocal.empty()) {
			std::filesystem::path assetPath = LocalTextToPathForSessionStore(attachment.assetPathLocal);
			if (assetPath.is_relative() && !sessionFilePath.empty()) {
				assetPath = (sessionFilePath.parent_path() / assetPath).lexically_normal();
				attachment.assetPathLocal = PathToLocalTextForSessionStore(assetPath);
			}
		}
		attachments.push_back(std::move(attachment));
	}
	return attachments;
}

nlohmann::json SerializeRunCheckpoint(
	const AIChatRunCheckpoint& checkpoint,
	const std::filesystem::path& sessionFilePath)
{
	nlohmann::json value = {
		{"schema_version", checkpoint.schemaVersion},
		{"protocol_type", static_cast<int>(checkpoint.protocolType)},
		{"model", LocalToUtf8TextForSessionStore(checkpoint.model)},
		{"state", checkpoint.state},
		{"summary", LocalToUtf8TextForSessionStore(checkpoint.summary)},
		{"sampling_rounds", checkpoint.samplingRounds},
		{"compaction_count", checkpoint.compactionCount},
		{"prompt_tokens", checkpoint.promptTokens},
		{"total_tokens", checkpoint.totalTokens},
		{"cached_input_tokens", checkpoint.cachedInputTokens},
		{"cache_write_input_tokens", checkpoint.cacheWriteInputTokens},
		{"accumulated_input_tokens", checkpoint.accumulatedInputTokens},
		{"accumulated_output_tokens", checkpoint.accumulatedOutputTokens},
		{"completed_model_rounds", checkpoint.completedModelRounds},
		{"has_usage", checkpoint.hasUsage},
		{"context_messages", nlohmann::json::array()},
		{"tool_calls", nlohmann::json::array()}
	};
	for (const AIChatMessage& message : checkpoint.contextMessages) {
		nlohmann::json row = {
			{"role", message.role},
			{"content", LocalToUtf8TextForSessionStore(message.content)},
			{"reasoning_content", message.reasoningContent},
			{"raw_message_json_utf8", message.rawMessageJsonUtf8}
		};
		row["attachments"] = SerializeImageAttachments(message.attachments, sessionFilePath);
		value["context_messages"].push_back(std::move(row));
	}
	for (const AIChatCheckpointToolCall& call : checkpoint.toolCalls) {
		value["tool_calls"].push_back({
			{"call_id", call.callId},
			{"name", call.name},
			{"arguments", LocalToUtf8TextForSessionStore(call.argumentsJson)},
			{"result", LocalToUtf8TextForSessionStore(call.resultJson)},
			{"completed", call.completed},
			{"ok", call.ok}
		});
	}
	return value;
}

bool DeserializeRunCheckpoint(const nlohmann::json& value, AIChatRunCheckpoint& checkpoint)
{
	if (!value.is_object()) {
		return false;
	}
	checkpoint = {};
	checkpoint.schemaVersion = static_cast<int>(GetJsonInt64(value, "schema_version", 1));
	checkpoint.protocolType = static_cast<AIProtocolType>(GetJsonInt64(value, "protocol_type", 0));
	checkpoint.model = GetJsonStringAsLocalText(value, "model");
	checkpoint.state = GetJsonStringUtf8(value, "state");
	checkpoint.summary = GetJsonStringAsLocalText(value, "summary");
	checkpoint.samplingRounds = static_cast<int>(GetJsonInt64(value, "sampling_rounds", 0));
	checkpoint.compactionCount = static_cast<int>(GetJsonInt64(value, "compaction_count", 0));
	checkpoint.promptTokens = static_cast<int>(GetJsonInt64(value, "prompt_tokens", 0));
	checkpoint.totalTokens = static_cast<int>(GetJsonInt64(value, "total_tokens", 0));
	checkpoint.cachedInputTokens = static_cast<int>(GetJsonInt64(value, "cached_input_tokens", 0));
	checkpoint.cacheWriteInputTokens = static_cast<int>(GetJsonInt64(value, "cache_write_input_tokens", 0));
	checkpoint.accumulatedInputTokens = GetJsonInt64(value, "accumulated_input_tokens", 0);
	checkpoint.accumulatedOutputTokens = GetJsonInt64(value, "accumulated_output_tokens", 0);
	checkpoint.completedModelRounds = static_cast<int>(GetJsonInt64(value, "completed_model_rounds", 0));
	checkpoint.hasUsage = GetJsonBool(value, "has_usage", false);
	if (!value.contains("accumulated_input_tokens") && checkpoint.hasUsage) {
		checkpoint.accumulatedInputTokens = (std::max)(0, checkpoint.promptTokens);
		checkpoint.accumulatedOutputTokens =
			(std::max)(0, checkpoint.totalTokens - checkpoint.promptTokens);
	}
	if (value.contains("context_messages") && value["context_messages"].is_array()) {
		for (const auto& row : value["context_messages"]) {
			if (!row.is_object()) {
				continue;
			}
			AIChatMessage message{
				GetJsonStringUtf8(row, "role"),
				GetJsonStringAsLocalText(row, "content"),
				GetJsonStringUtf8(row, "reasoning_content"),
				GetJsonStringUtf8(row, "raw_message_json_utf8")
			};
			message.attachments = DeserializeImageAttachments(row, {});
			checkpoint.contextMessages.push_back(std::move(message));
		}
	}
	if (value.contains("tool_calls") && value["tool_calls"].is_array()) {
		for (const auto& row : value["tool_calls"]) {
			if (!row.is_object()) {
				continue;
			}
			checkpoint.toolCalls.push_back(AIChatCheckpointToolCall{
				GetJsonStringUtf8(row, "call_id"),
				GetJsonStringUtf8(row, "name"),
				GetJsonStringAsLocalText(row, "arguments"),
				GetJsonStringAsLocalText(row, "result"),
				GetJsonBool(row, "completed", false),
				GetJsonBool(row, "ok", false)
			});
		}
	}
	return !checkpoint.contextMessages.empty();
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
		if (AIChatGoalManager::HasGoal(session.goal)) {
			outJson["goal"] = {
				{"objective", LocalToUtf8TextForSessionStore(session.goal.objectiveLocal)},
				{"status", AIChatGoalManager::StatusToString(session.goal.status)},
				{"tokens_used", session.goal.tokensUsed},
				{"elapsed_ms", session.goal.elapsedMs},
				{"created_at_unix_ms", session.goal.createdAtUnixMs},
				{"updated_at_unix_ms", session.goal.updatedAtUnixMs}
			};
		}
		if (session.hasRunCheckpoint) {
			outJson["run_checkpoint"] = SerializeRunCheckpoint(session.runCheckpoint, session.sessionFilePath);
		}
		outJson["pending_inputs"] = nlohmann::json::array();
		for (const auto& pending : session.pendingInputs) {
			nlohmann::json row = {
				{"id", pending.id},
				{"content", LocalToUtf8TextForSessionStore(pending.contentLocal)},
				{"queued_at_unix_ms", pending.queuedAtUnixMs}
			};
			row["attachments"] = SerializeImageAttachments(pending.attachments, session.sessionFilePath);
			outJson["pending_inputs"].push_back(std::move(row));
		}
		outJson["messages"] = nlohmann::json::array();

		for (const auto& message : session.messages) {
			nlohmann::json row = nlohmann::json::object();
			row["role"] = message.role;
			row["content"] = LocalToUtf8TextForSessionStore(message.contentLocal);
			row["include_in_context"] = message.includeInContext;
			row["visible_in_history"] = message.visibleInHistory;
			row["reasoning_content"] = message.reasoningContentUtf8;
			row["raw_message_json_utf8"] = message.rawMessageJsonUtf8;
			row["pending_input_id"] = message.pendingInputId;
			row["attachments"] = SerializeImageAttachments(message.attachments, session.sessionFilePath);
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
	if (jsonValue.contains("goal") && jsonValue["goal"].is_object()) {
		const nlohmann::json& goalValue = jsonValue["goal"];
		outSession.goal.objectiveLocal = GetJsonStringAsLocalText(goalValue, "objective");
		outSession.goal.status = AIChatGoalManager::StatusFromString(GetJsonStringUtf8(goalValue, "status"));
		outSession.goal.tokensUsed = (std::max)(0LL, GetJsonInt64(goalValue, "tokens_used", 0));
		outSession.goal.elapsedMs = (std::max)(0LL, GetJsonInt64(goalValue, "elapsed_ms", 0));
		outSession.goal.createdAtUnixMs = GetJsonInt64(goalValue, "created_at_unix_ms", 0);
		outSession.goal.updatedAtUnixMs = GetJsonInt64(goalValue, "updated_at_unix_ms", 0);
		if (outSession.goal.status == AIChatGoalStatus::Active) {
			outSession.goal.status = AIChatGoalStatus::Paused;
		}
		if (!AIChatGoalManager::HasGoal(outSession.goal)) {
			outSession.goal = {};
		}
	}
	if (jsonValue.contains("run_checkpoint")) {
		outSession.hasRunCheckpoint = DeserializeRunCheckpoint(
			jsonValue["run_checkpoint"],
			outSession.runCheckpoint);
		if (outSession.hasRunCheckpoint) {
			if (outSession.runCheckpoint.state == "completed") {
				outSession.hasRunCheckpoint = false;
				outSession.runCheckpoint = {};
			}
			else {
				outSession.runCheckpoint.state = "paused";
			}
		}
	}
	if (jsonValue.contains("pending_inputs") && jsonValue["pending_inputs"].is_array()) {
		for (const auto& row : jsonValue["pending_inputs"]) {
			if (!row.is_object()) {
				continue;
			}
			AIChatStoredPendingInput pending = {};
			const long long storedId = GetJsonInt64(row, "id", 0);
			pending.id = storedId > 0 ? static_cast<unsigned long long>(storedId) : 0;
			pending.contentLocal = GetJsonStringAsLocalText(row, "content");
			pending.queuedAtUnixMs = GetJsonInt64(row, "queued_at_unix_ms", 0);
			pending.attachments = DeserializeImageAttachments(row, {});
			if (!pending.contentLocal.empty() || !pending.attachments.empty()) {
				outSession.pendingInputs.push_back(std::move(pending));
			}
		}
	}

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
		const long long pendingInputId = GetJsonInt64(row, "pending_input_id", 0);
		message.pendingInputId = pendingInputId > 0
			? static_cast<unsigned long long>(pendingInputId)
			: 0;
		message.attachments = DeserializeImageAttachments(row, {});
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

namespace {

std::filesystem::path NormalizeSessionSourcePath(const std::string& sourceFilePathLocal)
{
	if (sourceFilePathLocal.empty()) {
		return {};
	}

	std::filesystem::path path = LocalTextToPathForSessionStore(sourceFilePathLocal);
	if (path.empty()) {
		return {};
	}

	std::error_code ec;
	path = std::filesystem::weakly_canonical(path, ec);
	if (ec || path.empty()) {
		ec.clear();
		path = std::filesystem::absolute(path, ec);
	}
	if (ec || path.empty()) {
		return {};
	}
	return path.lexically_normal();
}

} // namespace

bool AreAIChatSessionSourcePathsEquivalent(
	const std::string& leftSourceFilePathLocal,
	const std::string& rightSourceFilePathLocal)
{
	const std::filesystem::path left = NormalizeSessionSourcePath(leftSourceFilePathLocal);
	const std::filesystem::path right = NormalizeSessionSourcePath(rightSourceFilePathLocal);
	if (left.empty() || right.empty()) {
		return false;
	}
	return _wcsicmp(left.wstring().c_str(), right.wstring().c_str()) == 0;
}

std::filesystem::path ResolveAIChatSessionFilePath(
	const std::string& sourceFilePathLocal,
	const std::string& sessionId)
{
	return GetAIChatSessionDirectoryPathForSourceFile(sourceFilePathLocal) /
		(SanitizeSessionIdFileName(sessionId) + ".json");
}

std::filesystem::path GetAIChatSessionAssetDirectoryPath(const std::filesystem::path& sessionFilePath)
{
	if (sessionFilePath.empty()) {
		return {};
	}
	return sessionFilePath.parent_path() / (sessionFilePath.stem().wstring() + L".assets");
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

	std::filesystem::path temporaryPath = session.sessionFilePath;
	temporaryPath += std::format(
		L".tmp.{}.{}",
		static_cast<unsigned long>(GetCurrentProcessId()),
		static_cast<unsigned long>(GetCurrentThreadId()));
	std::ofstream out(temporaryPath, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		if (outError != nullptr) {
			*outError = "failed to open session file";
		}
		return false;
	}
	out.write("\xEF\xBB\xBF", 3);
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	out.flush();
	if (!out.good()) {
		out.close();
		DeleteFileW(temporaryPath.c_str());
		if (outError != nullptr) {
			*outError = "failed to write session file";
		}
		return false;
	}
	out.close();
	if (MoveFileExW(
			temporaryPath.c_str(),
			session.sessionFilePath.c_str(),
			MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == FALSE) {
		const DWORD errorCode = GetLastError();
		DeleteFileW(temporaryPath.c_str());
		if (outError != nullptr) {
			*outError = std::format("failed to replace session file: win32={}", errorCode);
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
	for (auto& message : outSession.messages) {
		for (auto& attachment : message.attachments) {
			std::filesystem::path path = LocalTextToPathForSessionStore(attachment.assetPathLocal);
			if (path.is_relative()) {
				attachment.assetPathLocal = PathToLocalTextForSessionStore(
					(sessionFilePath.parent_path() / path).lexically_normal());
			}
		}
	}
	for (auto& pending : outSession.pendingInputs) {
		for (auto& attachment : pending.attachments) {
			std::filesystem::path path = LocalTextToPathForSessionStore(attachment.assetPathLocal);
			if (path.is_relative()) {
				attachment.assetPathLocal = PathToLocalTextForSessionStore(
					(sessionFilePath.parent_path() / path).lexically_normal());
			}
		}
	}
	for (auto& message : outSession.runCheckpoint.contextMessages) {
		for (auto& attachment : message.attachments) {
			std::filesystem::path path = LocalTextToPathForSessionStore(attachment.assetPathLocal);
			if (path.is_relative()) {
				attachment.assetPathLocal = PathToLocalTextForSessionStore(
					(sessionFilePath.parent_path() / path).lexically_normal());
			}
		}
	}
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
	if (NormalizeSessionSourcePath(sourceFilePathLocal).empty()) {
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
		if (!AreAIChatSessionSourcePathsEquivalent(
			session.sourceFilePathHintLocal,
			sourceFilePathLocal)) {
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
