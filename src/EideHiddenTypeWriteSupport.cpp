#include "EideHiddenTypeWriteSupport.h"

#include <Windows.h>

#include <detours.h>

#include <array>
#include <atomic>
#include <cstring>
#include <mutex>
#include <string_view>

namespace e571 {
namespace {

constexpr std::uintptr_t kImageBase = 0x400000;
constexpr std::uintptr_t kE595ResolveDataTypeAddress = 0x462890;
constexpr unsigned int kHiddenGenericTypeId = 0x80000000u;
constexpr char kGenericTypeNameGbk[] = "\xCD\xA8\xD3\xC3\xD0\xCD";
constexpr char kGenericTypeNameUtf8[] = "\xE9\x80\x9A\xE7\x94\xA8\xE5\x9E\x8B";
constexpr std::array<unsigned char, 16> kE595ResolveDataTypeSignature = {
	0x83, 0xEC, 0x0C, 0x53, 0x8B, 0x5C, 0x24, 0x14,
	0x55, 0x56, 0x57, 0x8B, 0xE9, 0x53, 0x89, 0x6C
};

struct HiddenTypeResolutionContext {
	size_t replacementCount = 0;
};

std::mutex g_hookMutex;
thread_local HiddenTypeResolutionContext* g_resolutionContext = nullptr;
std::atomic<bool> g_mcpWriteHiddenGenericTypeEnabled{true};
std::atomic<bool> g_fullHiddenGenericTypeEnabled{false};
std::atomic<bool> g_fullHiddenGenericTypeHookInstalled{false};
bool g_fullHiddenGenericTypeHookAttachQueued = false;

#if defined(_M_IX86)
using FnResolveDataType = int(__thiscall*)(
	void*,
	unsigned char*,
	int,
	int,
	int,
	int);

FnResolveDataType g_originalResolveDataType = nullptr;
#endif

std::string_view TrimAscii(std::string_view text)
{
	while (!text.empty() && (text.front() == ' ' || text.front() == '\t')) {
		text.remove_prefix(1);
	}
	while (!text.empty() && (text.back() == ' ' || text.back() == '\t')) {
		text.remove_suffix(1);
	}
	return text;
}

bool IsHiddenGenericTypeName(std::string_view text)
{
	return text == kGenericTypeNameGbk || text == kGenericTypeNameUtf8;
}

bool IsHiddenGenericTypeName(const unsigned char* text)
{
	if (text == nullptr) {
		return false;
	}
	const auto* bytes = reinterpret_cast<const char*>(text);
	return std::strcmp(bytes, kGenericTypeNameGbk) == 0 ||
		std::strcmp(bytes, kGenericTypeNameUtf8) == 0;
}

bool MatchSignature(const void* address)
{
	if (address == nullptr) {
		return false;
	}
	__try {
		return std::memcmp(
			address,
			kE595ResolveDataTypeSignature.data(),
			kE595ResolveDataTypeSignature.size()) == 0;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

#if defined(_M_IX86)
int __fastcall HiddenTypeResolverHook(
	void* thisPtr,
	void*,
	unsigned char* typeName,
	int excludedType,
	int searchGlobalItems,
	int ownerId,
	int bypassOwnerFilter)
{
	if ((g_fullHiddenGenericTypeEnabled.load(std::memory_order_acquire) ||
			g_resolutionContext != nullptr) &&
		IsHiddenGenericTypeName(typeName)) {
		if (g_resolutionContext != nullptr) {
			++g_resolutionContext->replacementCount;
		}
		return static_cast<int>(kHiddenGenericTypeId);
	}
	return g_originalResolveDataType(
		thisPtr,
		typeName,
		excludedType,
		searchGlobalItems,
		ownerId,
		bypassOwnerFilter);
}
#endif

} // namespace

struct ScopedHiddenTypeResolverHook::State {
	bool required = false;
	bool active = false;
	bool ownsHook = false;
	HiddenTypeResolutionContext* previousContext = nullptr;
	HiddenTypeResolutionContext context;
	std::unique_lock<std::mutex> hookLock;
};

void SetMcpWriteHiddenGenericTypeEnabled(bool enabled) noexcept
{
	g_mcpWriteHiddenGenericTypeEnabled.store(enabled, std::memory_order_release);
}

bool IsMcpWriteHiddenGenericTypeEnabled() noexcept
{
	return g_mcpWriteHiddenGenericTypeEnabled.load(std::memory_order_acquire);
}

void SetFullHiddenGenericTypeEnabled(bool enabled) noexcept
{
	g_fullHiddenGenericTypeEnabled.store(enabled, std::memory_order_release);
}

bool IsFullHiddenGenericTypeEnabled() noexcept
{
	return g_fullHiddenGenericTypeEnabled.load(std::memory_order_acquire);
}

bool IsFullHiddenGenericTypeHookInstalled() noexcept
{
	return g_fullHiddenGenericTypeHookInstalled.load(std::memory_order_acquire);
}

bool AttachFullHiddenGenericTypeHookToCurrentDetourTransaction(
	std::uintptr_t moduleBase,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
#if !defined(_M_IX86)
	if (outTrace != nullptr) {
		*outTrace = "hidden_type_resolver_unsupported_architecture";
	}
	return false;
#else
	if (!IsFullHiddenGenericTypeEnabled()) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_full_disabled";
		}
		return false;
	}
	if (IsFullHiddenGenericTypeHookInstalled() || g_fullHiddenGenericTypeHookAttachQueued) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_full_already_attached";
		}
		return true;
	}
	if (moduleBase == 0 || kE595ResolveDataTypeAddress < kImageBase) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_invalid_module";
		}
		return false;
	}

	std::lock_guard<std::mutex> lock(g_hookMutex);
	void* const target = reinterpret_cast<void*>(
		moduleBase + (kE595ResolveDataTypeAddress - kImageBase));
	if (!MatchSignature(target)) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_signature_mismatch";
		}
		return false;
	}

	g_originalResolveDataType = reinterpret_cast<FnResolveDataType>(target);
	const LONG error = DetourAttach(
		reinterpret_cast<PVOID*>(&g_originalResolveDataType),
		HiddenTypeResolverHook);
	if (error != NO_ERROR) {
		g_originalResolveDataType = nullptr;
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_hook_failed|error=" + std::to_string(error);
		}
		return false;
	}
	g_fullHiddenGenericTypeHookAttachQueued = true;
	if (outTrace != nullptr) {
		*outTrace = "hidden_type_resolver_full_attach_queued|generic_type_id=0x80000000";
	}
	return true;
#endif
}

void CompleteFullHiddenGenericTypeHookInstallation(
	bool transactionCommitted,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
#if defined(_M_IX86)
	if (!g_fullHiddenGenericTypeHookAttachQueued) {
		if (outTrace != nullptr) {
			*outTrace = IsFullHiddenGenericTypeHookInstalled()
				? "hidden_type_resolver_full_installed"
				: "hidden_type_resolver_full_not_queued";
		}
		return;
	}
	g_fullHiddenGenericTypeHookAttachQueued = false;
	if (transactionCommitted) {
		g_fullHiddenGenericTypeHookInstalled.store(true, std::memory_order_release);
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_full_installed|generic_type_id=0x80000000";
		}
	}
	else {
		g_originalResolveDataType = nullptr;
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_full_transaction_failed";
		}
	}
#else
	(void)transactionCommitted;
	if (outTrace != nullptr) {
		*outTrace = "hidden_type_resolver_unsupported_architecture";
	}
#endif
}

bool ContainsHiddenGenericTypeDeclaration(const std::string& pageCode)
{
	std::string_view remaining(pageCode);
	while (!remaining.empty()) {
		const size_t lineEnd = remaining.find_first_of("\r\n");
		const std::string_view line = TrimAscii(remaining.substr(0, lineEnd));
		if (!line.empty() && line.front() == '.') {
			const size_t firstComma = line.find(',');
			if (firstComma != std::string_view::npos) {
				const size_t secondComma = line.find(',', firstComma + 1);
				const std::string_view typeName = TrimAscii(line.substr(
					firstComma + 1,
					secondComma == std::string_view::npos
						? std::string_view::npos
						: secondComma - firstComma - 1));
				if (IsHiddenGenericTypeName(typeName)) {
					return true;
				}
			}
		}

		if (lineEnd == std::string_view::npos) {
			break;
		}
		size_t nextLine = lineEnd + 1;
		if (remaining[lineEnd] == '\r' && nextLine < remaining.size() && remaining[nextLine] == '\n') {
			++nextLine;
		}
		remaining.remove_prefix(nextLine);
	}
	return false;
}

ScopedHiddenTypeResolverHook::ScopedHiddenTypeResolverHook(
	std::uintptr_t moduleBase,
	const std::string& pageCode,
	std::string* outTrace)
	: m_state(std::make_unique<State>())
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	m_state->required = ContainsHiddenGenericTypeDeclaration(pageCode);
	if (!m_state->required) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_generic_type_not_present";
		}
		return;
	}
	const bool mcpWriteEnabled = IsMcpWriteHiddenGenericTypeEnabled();
	const bool fullHookInstalled = IsFullHiddenGenericTypeHookInstalled();
	const bool fullHookActive = fullHookInstalled && IsFullHiddenGenericTypeEnabled();
	if (!mcpWriteEnabled && !fullHookActive) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_mcp_write_disabled";
		}
		return;
	}
#if !defined(_M_IX86)
	if (outTrace != nullptr) {
		*outTrace = "hidden_type_resolver_unsupported_architecture";
	}
	return;
#else
	if (fullHookInstalled) {
		m_state->active = true;
		m_state->previousContext = g_resolutionContext;
		g_resolutionContext = &m_state->context;
		if (outTrace != nullptr) {
			*outTrace = std::string(
				fullHookActive
					? "hidden_type_resolver_full_hook_active"
					: "hidden_type_resolver_installed_hook_reused") +
				"|generic_type_id=0x80000000";
		}
		return;
	}
	if (moduleBase == 0 || kE595ResolveDataTypeAddress < kImageBase) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_invalid_module";
		}
		return;
	}

	m_state->hookLock = std::unique_lock<std::mutex>(g_hookMutex);
	void* const target = reinterpret_cast<void*>(
		moduleBase + (kE595ResolveDataTypeAddress - kImageBase));
	if (!MatchSignature(target)) {
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_signature_mismatch";
		}
		return;
	}

	g_originalResolveDataType = reinterpret_cast<FnResolveDataType>(target);
	LONG error = DetourTransactionBegin();
	if (error == NO_ERROR) {
		error = DetourUpdateThread(GetCurrentThread());
	}
	if (error == NO_ERROR) {
		error = DetourAttach(
			reinterpret_cast<PVOID*>(&g_originalResolveDataType),
			HiddenTypeResolverHook);
	}
	if (error == NO_ERROR) {
		error = DetourTransactionCommit();
	}
	else {
		DetourTransactionAbort();
	}
	if (error != NO_ERROR) {
		g_originalResolveDataType = nullptr;
		if (outTrace != nullptr) {
			*outTrace = "hidden_type_resolver_hook_failed|error=" + std::to_string(error);
		}
		return;
	}

	m_state->active = true;
	m_state->ownsHook = true;
	m_state->previousContext = g_resolutionContext;
	g_resolutionContext = &m_state->context;
	if (outTrace != nullptr) {
		*outTrace = "hidden_type_resolver_hook_installed|generic_type_id=0x80000000";
	}
#endif
}

ScopedHiddenTypeResolverHook::~ScopedHiddenTypeResolverHook()
{
	if (m_state == nullptr || !m_state->active) {
		return;
	}
#if defined(_M_IX86)
	g_resolutionContext = m_state->previousContext;
	if (!m_state->ownsHook) {
		return;
	}
	LONG error = DetourTransactionBegin();
	if (error == NO_ERROR) {
		error = DetourUpdateThread(GetCurrentThread());
	}
	if (error == NO_ERROR) {
		error = DetourDetach(
			reinterpret_cast<PVOID*>(&g_originalResolveDataType),
			HiddenTypeResolverHook);
	}
	if (error == NO_ERROR) {
		error = DetourTransactionCommit();
	}
	else {
		DetourTransactionAbort();
	}
	if (error == NO_ERROR) {
		g_originalResolveDataType = nullptr;
	}
#endif
}

bool ScopedHiddenTypeResolverHook::IsRequired() const noexcept
{
	return m_state != nullptr && m_state->required;
}

bool ScopedHiddenTypeResolverHook::IsReady() const noexcept
{
	return m_state != nullptr && (!m_state->required || m_state->active);
}

size_t ScopedHiddenTypeResolverHook::ReplacementCount() const noexcept
{
	return m_state != nullptr ? m_state->context.replacementCount : 0;
}

} // namespace e571
