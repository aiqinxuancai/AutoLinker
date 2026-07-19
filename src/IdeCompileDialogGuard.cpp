#include "IdeCompileDialogGuard.h"

#include <Windows.h>

#include <atomic>
#include <cstdint>
#include <cwctype>
#include <format>
#include <iterator>
#include <string>

#include "..\\thirdparty\\json.hpp"
#include "Logger.h"

namespace IdeCompileDialogGuard {
namespace {

std::atomic_bool g_compileSessionActive = false;
std::atomic_bool g_dependencyDialogDismissed = false;

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

struct ChildSearchContext {
	bool promptMatched = false;
	HWND doNotWriteButton = nullptr;
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
		context->promptMatched = true;
	}
	else if (className == L"Button" && IsDoNotWriteButton(text)) {
		context->doNotWriteButton = child;
	}
	return TRUE;
}

struct DialogSearchContext {
	DWORD processId = 0;
	HWND dialog = nullptr;
	HWND doNotWriteButton = nullptr;
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
	if (!childContext.promptMatched ||
		childContext.doNotWriteButton == nullptr ||
		!IsWindowEnabled(childContext.doNotWriteButton)) {
		return TRUE;
	}

	context->dialog = window;
	context->doNotWriteButton = childContext.doNotWriteButton;
	return FALSE;
}

} // namespace

void BeginCompileSession()
{
	g_dependencyDialogDismissed.store(false, std::memory_order_release);
	g_compileSessionActive.store(true, std::memory_order_release);
}

void EndCompileSession()
{
	g_compileSessionActive.store(false, std::memory_order_release);
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
	const bool ok = observedPromptAccepted && compactPromptAccepted && unrelatedPromptRejected &&
		doNotWriteAccepted && writeRejected;
	return nlohmann::json({
		{"name", "ide-compile-dialog-guard"},
		{"ok", ok},
		{"observed_prompt_accepted", observedPromptAccepted},
		{"compact_prompt_accepted", compactPromptAccepted},
		{"unrelated_prompt_rejected", unrelatedPromptRejected},
		{"do_not_write_button_accepted", doNotWriteAccepted},
		{"write_button_rejected", writeRejected}
	}).dump();
}

} // namespace IdeCompileDialogGuard
