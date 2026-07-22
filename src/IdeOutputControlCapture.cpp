#include "IdeOutputControlCapture.h"

#include <CommCtrl.h>

#include <algorithm>
#include <cstddef>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"
#include "IdeLogStore.h"

#pragma comment(lib, "comctl32.lib")

namespace IdeOutputControlCapture {
namespace {

constexpr UINT_PTR kSubclassId = 0xA110;
constexpr std::size_t kMaxMessageChars = 256 * 1024;
constexpr int kMaxControlSnapshotChars = 16 * 1024 * 1024;
constexpr std::size_t kMinimumRollingOverlap = 64;

HWND g_outputWindow = nullptr;
HWND g_observerWindow = nullptr;
UINT g_layoutChangedMessage = 0;
std::string g_lastControlText;

void NotifyLayoutChanged(HWND outputWindow) noexcept
{
	const HWND observerWindow = g_observerWindow;
	const UINT message = g_layoutChangedMessage;
	if (observerWindow != nullptr && IsWindow(observerWindow) && message != 0) {
		PostMessageW(observerWindow, message, reinterpret_cast<WPARAM>(outputWindow), 0);
	}
}

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

std::string CopyAnsiText(const char* text) noexcept
{
	const std::size_t length = SafeAnsiLength(text, kMaxMessageChars + 1);
	if (length == 0 || !HasVisibleText(text, length)) {
		return {};
	}
	try {
		return std::string(text, length);
	}
	catch (...) {
		return {};
	}
}

bool ConvertWideToLocal(const wchar_t* text, std::size_t length, std::string& outText) noexcept
{
	outText.clear();
	if (text == nullptr || length == 0) {
		return true;
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
			return false;
		}
		outText.resize(static_cast<std::size_t>(localLength));
		return WideCharToMultiByte(
			CP_ACP,
			0,
			text,
			static_cast<int>(length),
			outText.data(),
			localLength,
			nullptr,
			nullptr) > 0;
	}
	catch (...) {
		outText.clear();
		return false;
	}
}

std::string CopyWideText(const wchar_t* text) noexcept
{
	const std::size_t length = SafeWideLength(text, kMaxMessageChars + 1);
	if (length == 0) {
		return {};
	}
	std::string localText;
	if (!ConvertWideToLocal(text, length, localText) ||
		!HasVisibleText(localText.data(), localText.size())) {
		return {};
	}
	return localText;
}

std::string CopyMessageText(HWND window, LPARAM lParam) noexcept
{
	if (lParam == 0) {
		return {};
	}
	if (IsWindowUnicode(window)) {
		return CopyWideText(reinterpret_cast<const wchar_t*>(lParam));
	}
	return CopyAnsiText(reinterpret_cast<const char*>(lParam));
}

bool ReadControlTextLocal(HWND window, std::string& outText) noexcept
{
	outText.clear();
	try {
		if (IsWindowUnicode(window)) {
			const int length = GetWindowTextLengthW(window);
			if (length < 0 || length > kMaxControlSnapshotChars) {
				return false;
			}
			std::wstring wideText(static_cast<std::size_t>(length) + 1, L'\0');
			const int copied = GetWindowTextW(window, wideText.data(), length + 1);
			if (copied < 0) {
				return false;
			}
			return ConvertWideToLocal(
				wideText.data(),
				static_cast<std::size_t>(copied),
				outText);
		}

		const int length = GetWindowTextLengthA(window);
		if (length < 0 || length > kMaxControlSnapshotChars) {
			return false;
		}
		outText.resize(static_cast<std::size_t>(length) + 1);
		const int copied = GetWindowTextA(window, outText.data(), length + 1);
		if (copied < 0) {
			outText.clear();
			return false;
		}
		outText.resize(static_cast<std::size_t>(copied));
		return true;
	}
	catch (...) {
		outText.clear();
		return false;
	}
}

std::size_t FindRollingOverlap(
	const std::string& previousText,
	const std::string& currentText)
{
	if (previousText.empty() || currentText.empty()) {
		return 0;
	}
	std::vector<std::size_t> prefix(currentText.size(), 0);
	for (std::size_t index = 1; index < currentText.size(); ++index) {
		std::size_t matched = prefix[index - 1];
		while (matched > 0 && currentText[index] != currentText[matched]) {
			matched = prefix[matched - 1];
		}
		if (currentText[index] == currentText[matched]) {
			++matched;
		}
		prefix[index] = matched;
	}

	std::size_t matched = 0;
	for (std::size_t index = 0; index < previousText.size(); ++index) {
		while (matched > 0 && previousText[index] != currentText[matched]) {
			matched = prefix[matched - 1];
		}
		if (previousText[index] == currentText[matched]) {
			++matched;
		}
		if (matched == currentText.size() && index + 1 < previousText.size()) {
			matched = prefix[matched - 1];
		}
	}
	return matched;
}

std::string ExtractAddedText(
	const std::string& previousText,
	const std::string& currentText)
{
	if (currentText.empty() || currentText == previousText) {
		return {};
	}
	if (previousText.empty()) {
		return currentText;
	}
	if (currentText.size() >= previousText.size() &&
		currentText.compare(0, previousText.size(), previousText) == 0) {
		return currentText.substr(previousText.size());
	}
	if (previousText.size() >= currentText.size() &&
		previousText.compare(0, currentText.size(), currentText) == 0) {
		return {};
	}

	const std::size_t overlap = FindRollingOverlap(previousText, currentText);
	const std::size_t shorterLength = (std::min)(previousText.size(), currentText.size());
	if (overlap > 0 &&
		(overlap >= kMinimumRollingOverlap || overlap * 2 >= shorterLength)) {
		return currentText.substr(overlap);
	}

	std::size_t commonPrefix = 0;
	while (commonPrefix < shorterLength &&
		previousText[commonPrefix] == currentText[commonPrefix]) {
		++commonPrefix;
	}
	std::size_t commonSuffix = 0;
	while (commonSuffix < previousText.size() - commonPrefix &&
		commonSuffix < currentText.size() - commonPrefix &&
		previousText[previousText.size() - commonSuffix - 1] ==
			currentText[currentText.size() - commonSuffix - 1]) {
		++commonSuffix;
	}
	return currentText.substr(
		commonPrefix,
		currentText.size() - commonPrefix - commonSuffix);
}

void AppendDeltaLines(const std::string& delta) noexcept
{
	try {
		std::size_t start = 0;
		while (start < delta.size()) {
			const std::size_t lineEnd = delta.find_first_of("\r\n", start);
			std::size_t next = lineEnd;
			if (lineEnd == std::string::npos) {
				next = delta.size();
			}
			else {
				next = lineEnd + 1;
				if (delta[lineEnd] == '\r' && next < delta.size() && delta[next] == '\n') {
					++next;
				}
			}
			const std::size_t length = next - start;
			if (length > 0 && HasVisibleText(delta.data() + start, length)) {
				IdeLogStore::AppendLocal(
					delta.data() + start,
					length,
					IdeLogStore::EntrySource::OutputControlSubclass);
			}
			start = next;
		}
	}
	catch (...) {
	}
}

void CaptureControlChange(HWND window, const std::string& fallbackText) noexcept
{
	try {
		std::string currentText;
		if (!ReadControlTextLocal(window, currentText)) {
			AppendDeltaLines(fallbackText);
			return;
		}
		const std::string delta = ExtractAddedText(g_lastControlText, currentText);
		g_lastControlText = std::move(currentText);
		AppendDeltaLines(delta);
	}
	catch (...) {
		AppendDeltaLines(fallbackText);
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
	const bool capturesText = message == EM_REPLACESEL || message == WM_SETTEXT ||
		message == WM_CLEAR || message == WM_CUT || message == WM_UNDO;
	// lParam 仅作读取控件失败时的兜底；日志边界以原输出控件处理后的实际文本变化为准。
	const std::string capturedText = capturesText ? CopyMessageText(window, lParam) : std::string();
	const LRESULT result = DefSubclassProc(window, message, wParam, lParam);
	if (capturesText && (message != WM_SETTEXT || result != FALSE)) {
		CaptureControlChange(window, capturedText);
	}
	if (message == WM_WINDOWPOSCHANGED || message == WM_SHOWWINDOW || message == WM_STYLECHANGED) {
		NotifyLayoutChanged(window);
	}
	if (message == WM_NCDESTROY && g_outputWindow == window) {
		NotifyLayoutChanged(window);
		g_outputWindow = nullptr;
		g_observerWindow = nullptr;
		g_layoutChangedMessage = 0;
		g_lastControlText.clear();
	}
	return result;
}

} // namespace

bool Attach(HWND outputWindow, HWND observerWindow, UINT layoutChangedMessage) noexcept
{
	if (outputWindow == nullptr || !IsWindow(outputWindow)) {
		return false;
	}
	if (IsAttachedTo(outputWindow)) {
		g_observerWindow = observerWindow;
		g_layoutChangedMessage = layoutChangedMessage;
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
	g_observerWindow = observerWindow;
	g_layoutChangedMessage = layoutChangedMessage;
	ReadControlTextLocal(outputWindow, g_lastControlText);
	return true;
}

void Detach() noexcept
{
	const HWND outputWindow = g_outputWindow;
	g_outputWindow = nullptr;
	g_observerWindow = nullptr;
	g_layoutChangedMessage = 0;
	g_lastControlText.clear();
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

std::string BuildSelfTestJson()
{
	char ansiBuffer[] = {'f', 'i', 'r', 's', 't', '\0', 's', 'e', 'c', 'o', 'n', 'd', '\0'};
	const std::string ansiCopy = CopyAnsiText(ansiBuffer);
	ansiBuffer[5] = '-';
	const bool ansiBufferReusePassed = ansiCopy == "first";

	wchar_t wideBuffer[] = {L'1', L'2', L'3', L'\0', L'4', L'5', L'6', L'\0'};
	const std::string wideCopy = CopyWideText(wideBuffer);
	wideBuffer[3] = L'-';
	const bool wideBufferReusePassed = wideCopy == "123";

	const bool appendDeltaPassed = ExtractAddedText(
		"* 1\r\n* 2\r\n",
		"* 1\r\n* 2\r\n* 3\r\n") == "* 3\r\n";
	const std::string rollingPrevious =
		"prefix-padding-that-keeps-the-overlap-above-the-minimum-threshold\r\n"
		"* 100\r\n* 101\r\n* 102\r\n";
	const std::string rollingCurrent =
		"* 100\r\n* 101\r\n* 102\r\n* 103\r\n";
	const bool rollingDeltaPassed = ExtractAddedText(
		rollingPrevious,
		rollingCurrent) == "* 103\r\n";
	const bool replacementDeltaPassed = ExtractAddedText(
		"old-prefix\r\nkept-line\r\n",
		"trim-marker\r\nkept-line\r\n") == "trim-marker";
	const bool deletionSkipped = ExtractAddedText("abcdef", "abc").empty();

	return nlohmann::json({
		{"name", "ide-output-control-capture"},
		{"ok", ansiBufferReusePassed && wideBufferReusePassed &&
			appendDeltaPassed && rollingDeltaPassed && replacementDeltaPassed && deletionSkipped},
		{"ansi_buffer_reuse_safe", ansiBufferReusePassed},
		{"wide_buffer_reuse_safe", wideBufferReusePassed},
		{"append_delta", appendDeltaPassed},
		{"rolling_snapshot_delta", rollingDeltaPassed},
		{"replacement_delta", replacementDeltaPassed},
		{"deletion_skipped", deletionSkipped}
	}).dump();
}

} // namespace IdeOutputControlCapture
