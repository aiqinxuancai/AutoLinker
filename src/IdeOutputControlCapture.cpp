#include "IdeOutputControlCapture.h"

#include <CommCtrl.h>

#include <cstddef>
#include <string>

#include "IdeLogStore.h"

#pragma comment(lib, "comctl32.lib")

namespace IdeOutputControlCapture {
namespace {

constexpr UINT_PTR kSubclassId = 0xA110;
constexpr std::size_t kMaxMessageChars = 256 * 1024;

HWND g_outputWindow = nullptr;

std::size_t SafeAnsiLength(const char* text, std::size_t maxLength) noexcept
{
	if (text == nullptr || maxLength == 0) {
		return 0;
	}
	__try {
		std::size_t length = 0;
		while (length < maxLength && text[length] != '\0') {
			++length;
		}
		return length;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

std::size_t SafeWideLength(const wchar_t* text, std::size_t maxLength) noexcept
{
	if (text == nullptr || maxLength == 0) {
		return 0;
	}
	__try {
		std::size_t length = 0;
		while (length < maxLength && text[length] != L'\0') {
			++length;
		}
		return length;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

bool HasVisibleText(const char* text, std::size_t length) noexcept
{
	if (text == nullptr) {
		return false;
	}
	__try {
		for (std::size_t index = 0; index < length; ++index) {
			if (text[index] != '\r' && text[index] != '\n') {
				return true;
			}
		}
		return false;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

void AppendAnsiText(const char* text) noexcept
{
	const std::size_t length = SafeAnsiLength(text, kMaxMessageChars + 1);
	if (length == 0 || !HasVisibleText(text, length)) {
		return;
	}
	IdeLogStore::AppendLocal(
		text,
		length,
		IdeLogStore::EntrySource::OutputControlSubclass);
}

void AppendWideText(const wchar_t* text) noexcept
{
	const std::size_t length = SafeWideLength(text, kMaxMessageChars + 1);
	if (length == 0) {
		return;
	}
	try {
		const int localLength = WideCharToMultiByte(
			CP_ACP,
			0,
			text,
			static_cast<int>(length),
			nullptr,
			0,
			nullptr,
			nullptr);
		if (localLength <= 0) {
			return;
		}
		std::string localText(static_cast<std::size_t>(localLength), '\0');
		if (WideCharToMultiByte(
			CP_ACP,
			0,
			text,
			static_cast<int>(length),
			localText.data(),
			localLength,
			nullptr,
			nullptr) <= 0 ||
			!HasVisibleText(localText.data(), localText.size())) {
			return;
		}
		IdeLogStore::AppendLocal(
			localText.data(),
			localText.size(),
			IdeLogStore::EntrySource::OutputControlSubclass);
	}
	catch (...) {
		// 控件消息采集失败不能影响 IDE 原始窗口过程。
	}
}

void AppendMessageText(HWND window, LPARAM lParam) noexcept
{
	if (lParam == 0) {
		return;
	}
	if (IsWindowUnicode(window)) {
		AppendWideText(reinterpret_cast<const wchar_t*>(lParam));
	}
	else {
		AppendAnsiText(reinterpret_cast<const char*>(lParam));
	}
}

LRESULT CALLBACK OutputControlSubclassProc(
	HWND window,
	UINT message,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR /*subclassId*/,
	DWORD_PTR /*referenceData*/)
{
	const bool capturesText = message == EM_REPLACESEL || message == WM_SETTEXT;
	const LRESULT result = DefSubclassProc(window, message, wParam, lParam);
	if (capturesText && (message != WM_SETTEXT || result != FALSE)) {
		AppendMessageText(window, lParam);
	}
	if (message == WM_NCDESTROY && g_outputWindow == window) {
		g_outputWindow = nullptr;
	}
	return result;
}

} // namespace

bool Attach(HWND outputWindow) noexcept
{
	if (outputWindow == nullptr || !IsWindow(outputWindow)) {
		return false;
	}
	if (IsAttachedTo(outputWindow)) {
		return true;
	}

	Detach();
	if (!SetWindowSubclass(
		outputWindow,
		OutputControlSubclassProc,
		kSubclassId,
		0)) {
		return false;
	}
	g_outputWindow = outputWindow;
	return true;
}

void Detach() noexcept
{
	const HWND outputWindow = g_outputWindow;
	g_outputWindow = nullptr;
	if (outputWindow != nullptr && IsWindow(outputWindow)) {
		RemoveWindowSubclass(
			outputWindow,
			OutputControlSubclassProc,
			kSubclassId);
	}
}

bool IsAttachedTo(HWND outputWindow) noexcept
{
	return outputWindow != nullptr &&
		g_outputWindow == outputWindow &&
		IsWindow(outputWindow);
}

} // namespace IdeOutputControlCapture
