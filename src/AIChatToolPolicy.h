#pragma once

#include <cstddef>
#include <memory>
#include <string>

// 内置 AI 对话的工具调用预算、重复读取与写后收敛策略。
namespace AIChatToolPolicy {

inline constexpr int kExplorationReminderInterval = 8;
inline constexpr std::size_t kReadFileContextBytes = 24 * 1024;
inline constexpr std::size_t kReadFilesPerFileContextBytes = 8 * 1024;
inline constexpr std::size_t kReadRealFileCodeContextBytes = 48 * 1024;

struct Decision {
	bool allowed = true;
	std::string reason;
	std::string resultJsonUtf8;
};

class Session {
public:
	Session();
	~Session();
	Session(Session&&) noexcept;
	Session& operator=(Session&&) noexcept;
	Session(const Session&) = delete;
	Session& operator=(const Session&) = delete;

	Decision BeforeToolCall(const std::string& toolName, const std::string& argumentsJsonUtf8);
	std::string AfterToolCall(
		const std::string& toolName,
		const std::string& argumentsJsonUtf8,
		const std::string& resultJsonUtf8,
		bool ok);

	int ExplorationCalls() const;
	bool PreferLowThinkingForNextRound() const;
	void StartNewContextWindow();

private:
	struct Impl;
	std::unique_ptr<Impl> m_impl;
};

} // namespace AIChatToolPolicy
