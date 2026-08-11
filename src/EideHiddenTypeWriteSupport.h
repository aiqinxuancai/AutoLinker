#pragma once

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>

namespace e571 {

// 检测源码声明的类型字段中是否包含 IDE 隐藏的“通用型”。
bool ContainsHiddenGenericTypeDeclaration(const std::string& pageCode);

// 控制 MCP 整页写入期间是否临时解析隐藏的“通用型”。
void SetMcpWriteHiddenGenericTypeEnabled(bool enabled) noexcept;
bool IsMcpWriteHiddenGenericTypeEnabled() noexcept;

// 控制是否在 IDE 全部类型解析路径中启用隐藏的“通用型”。
void SetFullHiddenGenericTypeEnabled(bool enabled) noexcept;
bool IsFullHiddenGenericTypeEnabled() noexcept;
bool IsFullHiddenGenericTypeHookInstalled() noexcept;

// 在 AutoLinker 的总 Detours 事务中附加/完成常驻类型解析 Hook。
bool AttachFullHiddenGenericTypeHookToCurrentDetourTransaction(
	std::uintptr_t moduleBase,
	std::string* outTrace = nullptr);
void CompleteFullHiddenGenericTypeHookInstallation(
	bool transactionCommitted,
	std::string* outTrace = nullptr);

// 在 MCP 文本解析期间按需复用常驻 Hook，或安装短生命周期 Hook。
class ScopedHiddenTypeResolverHook {
public:
	ScopedHiddenTypeResolverHook(
		std::uintptr_t moduleBase,
		const std::string& pageCode,
		std::string* outTrace = nullptr);
	~ScopedHiddenTypeResolverHook();

	ScopedHiddenTypeResolverHook(const ScopedHiddenTypeResolverHook&) = delete;
	ScopedHiddenTypeResolverHook& operator=(const ScopedHiddenTypeResolverHook&) = delete;

	bool IsRequired() const noexcept;
	bool IsReady() const noexcept;
	size_t ReplacementCount() const noexcept;

private:
	struct State;
	std::unique_ptr<State> m_state;
};

} // namespace e571
