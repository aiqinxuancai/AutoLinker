#include "IdeCompileDialogGuard.h"

#include <Windows.h>

#include <atomic>
#include <chrono>
#include <cstdint>
#include <cwctype>
#include <format>
#include <iterator>
#include <thread>
#include <string>

#include "..\\thirdparty\\json.hpp"
#include "Logger.h"

namespace IdeCompileDialogGuard {
namespace {

std::atomic_bool g_compileSessionActive = false;
std::atomic_bool g_dependencyDialogDismissed = false;
std::atomic_uint g_nameConflictDialogsDismissed = 0;
std::atomic_uint64_t g_compileSessionGeneration = 0;

std::wstring GetWindowTextCopy(HWND window)
{
	if (window == nullptr || !IsWindow(window)) {
		return {};
	}
	const int length = GetWindowTextLengthW(window);
	if (length <= 0) {
		return {};
	}
	std::wstring text(static_cast<size_t>(length) + 1, L'\0');
	const int copied = GetWindowTextW(window, text.data(), length + 1);
	if (copied <= 0) {
		return {};
	}
	text.resize(static_cast<size_t>(copied));
	return text;
}

std::wstring GetWindowClassCopy(HWND window)
{
	wchar_t buffer[128] = {};
	const int length = GetClassNameW(window, buffer, static_cast<int>(std::size(buffer)));
	return length > 0 ? std::wstring(buffer, static_cast<size_t>(length)) : std::wstring();
}

std::wstring NormalizeVisibleText(std::wstring text)
{
	std::wstring normalized;
	normalized.reserve(text.size());
	for (const wchar_t ch : text) {
		if (ch == L'&' || std::iswspace(ch)) {
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

bool IsDependencyWritePrompt(const std::wstring& text)
{
	const std::wstring normalized = NormalizeVisibleText(text);
	return normalized.find(L"此程序所使用到的相关依赖文件") != std::wstring::npos &&
		normalized.find(L"写出到同一目录") != std::wstring::npos;
}

bool IsDoNotWriteButton(const std::wstring& text)
{
	return NormalizeVisibleText(text).find(L"不写出") != std::wstring::npos;
}

bool IsNameConflictPrompt(const std::wstring& text)
{
	const std::wstring normalized = NormalizeVisibleText(text);
	return normalized.find(L"现有多个名称与指定拼音输入字") != std::wstring::npos &&
		normalized.find(L"相对应") != std::wstring::npos &&
		normalized.find(L"请选择") != std::wstring::npos;
}

bool IsConfirmButton(const std::wstring& text)
{
	const std::wstring normalized = NormalizeVisibleText(text);
	return normalized.find(L"确定") != std::wstring::npos;
}

struct ChildSearchContext {
	bool dependencyPromptMatched = false;
	bool nameConflictPromptMatched = false;
	HWND doNotWriteButton = nullptr;
	HWND confirmButton = nullptr;
};

BOOL CALLBACK EnumDialogChild(HWND child, LPARAM param)
{
	auto* context = reinterpret_cast<ChildSearchContext*>(param);
	if (context == nullptr || !IsWindowVisible(child)) {
		return TRUE;
	}

	const std::wstring className = GetWindowClassCopy(child);
	const std::wstring text = GetWindowTextCopy(child);
	if (className == L"Static" && IsDependencyWritePrompt(text)) {
		context->dependencyPromptMatched = true;
	}
	else if (className == L"Button" && IsDoNotWriteButton(text)) {
		context->doNotWriteButton = child;
	}
	else if (className == L"Static" && IsNameConflictPrompt(text)) {
		context->nameConflictPromptMatched = true;
	}
	else if (className == L"Button" && IsConfirmButton(text)) {
		context->confirmButton = child;
	}
	return TRUE;
}

struct DialogSearchContext {
	DWORD processId = 0;
	bool nameConflictOnly = false;
	HWND dialog = nullptr;
	HWND doNotWriteButton = nullptr;
	HWND confirmButton = nullptr;
};

BOOL CALLBACK EnumProcessWindow(HWND window, LPARAM param)
{
	auto* context = reinterpret_cast<DialogSearchContext*>(param);
	if (context == nullptr || context->dialog != nullptr || !IsWindowVisible(window)) {
		return context != nullptr && context->dialog == nullptr;
	}

	DWORD processId = 0;
	GetWindowThreadProcessId(window, &processId);
	if (processId != context->processId || GetWindowClassCopy(window) != L"#32770") {
		return TRUE;
	}

	ChildSearchContext childContext;
	EnumChildWindows(window, EnumDialogChild, reinterpret_cast<LPARAM>(&childContext));
	if (!context->nameConflictOnly && childContext.dependencyPromptMatched &&
		childContext.doNotWriteButton != nullptr && IsWindowEnabled(childContext.doNotWriteButton)) {
		context->dialog = window;
		context->doNotWriteButton = childContext.doNotWriteButton;
		return FALSE;
	}
	if (!context->nameConflictOnly || !childContext.nameConflictPromptMatched ||
		childContext.confirmButton == nullptr || !IsWindowEnabled(childContext.confirmButton)) {
		return TRUE;
	}

	context->dialog = window;
	context->confirmButton = childContext.confirmButton;
	return FALSE;
}

void CompileDialogWatcherMain(std::uint64_t sessionGeneration)
{
	while (g_compileSessionActive.load(std::memory_order_acquire) &&
		g_compileSessionGeneration.load(std::memory_order_acquire) == sessionGeneration) {
		TryDismissDependencyWriteDialog();
		TryDismissNameConflictDialog();
		std::this_thread::sleep_for(std::chrono::milliseconds(50));
	}
}

} // namespace

void BeginCompileSession()
{
	g_dependencyDialogDismissed.store(false, std::memory_order_release);
	g_nameConflictDialogsDismissed.store(0, std::memory_order_release);
	g_compileSessionActive.store(true, std::memory_order_release);
	// 编译入口可能在 IDE 主线程中同步阻塞，使用独立监视线程保证该线程仍可处理编译。
	const std::uint64_t sessionGeneration =
		g_compileSessionGeneration.fetch_add(1, std::memory_order_acq_rel) + 1;
	std::thread(CompileDialogWatcherMain, sessionGeneration).detach();
}

void EndCompileSession()
{
	g_compileSessionActive.store(false, std::memory_order_release);
	g_compileSessionGeneration.fetch_add(1, std::memory_order_acq_rel);
}

bool IsCompileSessionActive()
{
	return g_compileSessionActive.load(std::memory_order_acquire);
}

bool TryDismissDependencyWriteDialog()
{
	if (!g_compileSessionActive.load(std::memory_order_acquire) ||
		g_dependencyDialogDismissed.load(std::memory_order_acquire)) {
		return false;
	}

	DialogSearchContext context;
	context.processId = GetCurrentProcessId();
	EnumWindows(EnumProcessWindow, reinterpret_cast<LPARAM>(&context));
	if (context.dialog == nullptr || context.doNotWriteButton == nullptr) {
		return false;
	}

	if (!PostMessageW(context.doNotWriteButton, BM_CLICK, 0, 0)) {
		return false;
	}
	g_dependencyDialogDismissed.store(true, std::memory_order_release);
	Logger::Instance().Write(
		"SilentCompile",
		std::format(
			"dismissed dependency write dialog with do-not-write button dialog=0x{:X} button_id={}",
			reinterpret_cast<std::uintptr_t>(context.dialog),
			GetDlgCtrlID(context.doNotWriteButton)));
	return true;
}

bool WasDependencyWriteDialogDismissed()
{
	return g_dependencyDialogDismissed.load(std::memory_order_acquire);
}

bool TryDismissNameConflictDialog()
{
	if (!g_compileSessionActive.load(std::memory_order_acquire)) {
		return false;
	}

	DialogSearchContext context;
	context.processId = GetCurrentProcessId();
	context.nameConflictOnly = true;
	EnumWindows(EnumProcessWindow, reinterpret_cast<LPARAM>(&context));
	if (context.dialog == nullptr || context.confirmButton == nullptr) {
		return false;
	}

	if (!PostMessageW(context.confirmButton, BM_CLICK, 0, 0)) {
		return false;
	}
	g_nameConflictDialogsDismissed.fetch_add(1, std::memory_order_acq_rel);
	Logger::Instance().Write(
		"SilentCompile",
		std::format(
			"dismissed name conflict dialog with confirm button dialog=0x{:X} button_id={}",
			reinterpret_cast<std::uintptr_t>(context.dialog),
			GetDlgCtrlID(context.confirmButton)));
	return true;
}

bool WasNameConflictDialogDismissed()
{
	return g_nameConflictDialogsDismissed.load(std::memory_order_acquire) != 0;
}

std::string BuildSelfTestJson()
{
	const bool observedPromptAccepted = IsDependencyWritePrompt(
		L"请问需要将此程序所使用到的相关依赖文件写出到同一目录中去吗?");
	const bool compactPromptAccepted = IsDependencyWritePrompt(
		L"询问需要将此程序所使用到的相关依赖文件写出到同一目录中去吗？");
	const bool unrelatedPromptRejected = !IsDependencyWritePrompt(
		L"是否覆盖已经存在的目标文件？");
	const bool doNotWriteAccepted = IsDoNotWriteButton(L"不写出(&C)");
	const bool writeRejected = !IsDoNotWriteButton(L"写出(&W)");
	const bool nameConflictAccepted = IsNameConflictPrompt(
		L"现有多个名称与指定拼音输入字“位置”相对应，请选择其一:");
	const bool nameConflictRejected = !IsNameConflictPrompt(L"请选择要覆盖的目标文件:");
	const bool confirmAccepted = IsConfirmButton(L"确定(O)");
	const bool cancelRejected = !IsConfirmButton(L"取消(C)");
	const bool ok = observedPromptAccepted && compactPromptAccepted && unrelatedPromptRejected &&
		doNotWriteAccepted && writeRejected && nameConflictAccepted && nameConflictRejected &&
		confirmAccepted && cancelRejected;
	return nlohmann::json({
		{"name", "ide-compile-dialog-guard"},
		{"ok", ok},
		{"observed_prompt_accepted", observedPromptAccepted},
		{"compact_prompt_accepted", compactPromptAccepted},
		{"unrelated_prompt_rejected", unrelatedPromptRejected},
		{"do_not_write_button_accepted", doNotWriteAccepted},
		{"write_button_rejected", writeRejected},
		{"name_conflict_prompt_accepted", nameConflictAccepted},
		{"name_conflict_prompt_rejected", nameConflictRejected},
		{"confirm_button_accepted", confirmAccepted},
		{"cancel_button_rejected", cancelRejected}
	}).dump();
}

} // namespace IdeCompileDialogGuard
