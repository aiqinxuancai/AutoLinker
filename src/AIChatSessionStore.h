#pragma once

#include <filesystem>
#include <string>
#include <vector>

// AI 对话会话存储结构。
struct AIChatStoredMessage {
	std::string role;
	std::string contentLocal;
	bool includeInContext = true;
	bool visibleInHistory = true;
	std::string reasoningContentUtf8;
	std::string rawMessageJsonUtf8;
};

// AI 对话会话存储数据。
struct AIChatStoredSession {
	int schemaVersion = 1;
	std::string sessionId;
	std::string sourceFileNameLocal;
	std::string sourceFilePathHintLocal;
	long long createdAtUnixMs = 0;
	long long updatedAtUnixMs = 0;
	long long elapsedMs = 0;  // 会话累计实际用时（毫秒）。
	std::string createdAtDisplayLocal;
	std::string updatedAtDisplayLocal;
	std::string rollingSummaryLocal;
	std::string planModeState;      // 计划模式状态：normal / planning / awaiting_approval / approved。
	std::string pendingPlanLocal;   // 待批准的计划正文。
	bool autoAllowWrites = false;   // 自动允许写入模式。
	std::vector<AIChatStoredMessage> messages;
	std::filesystem::path sessionFilePath;
};

// AI 对话最近会话列表项。
struct AIChatStoredSessionListEntry {
	std::string sessionId;
	std::filesystem::path sessionFilePath;
	long long updatedAtUnixMs = 0;
	std::string updatedAtDisplayLocal;
	std::string titleLocal;
};

// 生成新的 AI 对话会话标识。
std::string CreateAIChatSessionId();

// 获取当前源码对应的 AI 对话会话目录。
std::filesystem::path GetAIChatSessionDirectoryPathForSourceFile(const std::string& sourceFilePathLocal);

// 解析当前源码下某个会话文件路径。
std::filesystem::path ResolveAIChatSessionFilePath(
	const std::string& sourceFilePathLocal,
	const std::string& sessionId);

// 保存 AI 对话会话到磁盘。
bool SaveAIChatStoredSession(const AIChatStoredSession& session, std::string* outError = nullptr);

// 从磁盘加载 AI 对话会话。
bool LoadAIChatStoredSession(
	const std::filesystem::path& sessionFilePath,
	AIChatStoredSession& outSession,
	std::string* outError = nullptr);

// 列出当前源码最近的 AI 对话会话。
std::vector<AIChatStoredSessionListEntry> ListRecentAIChatStoredSessions(
	const std::string& sourceFilePathLocal,
	size_t limit);
