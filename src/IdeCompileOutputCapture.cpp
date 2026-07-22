#include "IdeCompileOutputCapture.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <format>
#include <mutex>
#include <sstream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include <detours.h>
#include <intrin.h>

#include "..\\thirdparty\\json.hpp"
#include "Global.h"
#include "Logger.h"
#include "MemFind.h"

namespace IdeCompileOutputCapture {
namespace {

constexpr size_t kMaxCaptureBytes = 2 * 1024 * 1024;
constexpr size_t kMaxTextProbeBytes = kMaxCaptureBytes + 1;
constexpr size_t kSemanticDistanceLimit = 0x300;
constexpr size_t kReturnSearchBytes = 0x100;
constexpr size_t kDebugOutputQueueMaxBytes = 16 * 1024 * 1024;
constexpr size_t kDebugOutputUiBatchBytes = 8 * 1024;
constexpr int kDebugOutputNativeRollingThreshold = 8 * 1024;
constexpr UINT kDebugOutputCoalesceMs = 40;
constexpr UINT kDebugOutputContinueMs = 1;
constexpr wchar_t kDebugOutputFlushMessageName[] = L"AutoLinker.IdeDebugOutputFlush.v1";
constexpr wchar_t kDebugOutputTimerWindowClass[] = L"AutoLinker.DebugOutputTimer.v1";
constexpr UINT_PTR kDebugOutputTimerId = 1;

constexpr const char* kOutputFunctionEntryPattern =
	"55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 "
	"50 64 89 25 00 00 00 00 83 EC 08 53 56 57 8B F9 "
	"B8 01 00 00 00 39 87 ?? ?? 00 00 0F 84 ?? ?? ?? ?? "
	"8B B7 ?? ?? 00 00 8B 5D 08 81 C6 ?? ?? 00 00 85 DB";

constexpr const char* kSehFunctionProloguePattern =
	"55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 "
	"50 64 89 25 00 00 00 00";

constexpr const char* kOutputWriteSemanticPattern =
	"8B 56 1C 57 57 8B 3D ?? ?? ?? ?? 68 B1 00 00 00 "
	"52 FF D7 8B 46 1C 6A 00 6A 00 68 B7 00 00 00 50 "
	"FF D7 8B 4D 08 8B 56 1C 51 6A 00 68 C2 00 00 00 "
	"52 FF D7";

constexpr const char* kThiscallReturnPattern = "C2 08 00";

// OUTPUT_DEBUG_STRING_EVENT 分支：验证 "* " 前缀后调用输出函数，再释放远程文本副本。
constexpr const char* kDebugOutputCallSitePattern =
	"80 3E 2A 75 ?? 80 7E 01 20 75 ?? 6A 01 56 "
	"B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 56 E8 ?? ?? ?? ??";
constexpr size_t kDebugOutputCallInstructionOffset = 19;
constexpr size_t kDebugOutputCallReturnOffset = 24;

struct CodeRange {
	const byte* data = nullptr;
	size_t size = 0;
	std::uintptr_t runtimeBase = 0;
};

struct PatternMatch {
	size_t rangeIndex = 0;
	size_t offset = 0;
};

struct ResolveResult {
	std::uintptr_t address = 0;
	std::string method;
	std::string diagnostics;
};

size_t PatternLength(const char* pattern)
{
	if (pattern == nullptr) {
		return 0;
	}
	std::istringstream input(pattern);
	std::string token;
	size_t length = 0;
	while (input >> token) {
		++length;
	}
	return length;
}

std::vector<size_t> FindOffsetsInRange(
	const CodeRange& range,
	size_t begin,
	size_t end,
	const char* pattern)
{
	std::vector<size_t> result;
	if (range.data == nullptr || begin >= range.size || begin >= end) {
		return result;
	}
	end = (std::min)(end, range.size);
	const auto relative = FindPatternOffsets(range.data + begin, end - begin, pattern);
	result.reserve(relative.size());
	for (const size_t offset : relative) {
		result.push_back(begin + offset);
	}
	return result;
}

std::vector<PatternMatch> FindAllMatches(
	const std::vector<CodeRange>& ranges,
	const char* pattern)
{
	std::vector<PatternMatch> matches;
	for (size_t rangeIndex = 0; rangeIndex < ranges.size(); ++rangeIndex) {
		const CodeRange& range = ranges[rangeIndex];
		const auto offsets = FindOffsetsInRange(range, 0, range.size, pattern);
		for (const size_t offset : offsets) {
			matches.push_back({rangeIndex, offset});
		}
	}
	return matches;
}

bool ValidateCandidate(const CodeRange& range, size_t candidateOffset)
{
	const size_t semanticLength = PatternLength(kOutputWriteSemanticPattern);
	if (semanticLength == 0 || candidateOffset >= range.size) {
		return false;
	}

	const size_t semanticSearchEnd = (std::min)(
		range.size,
		candidateOffset + kSemanticDistanceLimit + semanticLength);
	const auto semanticMatches = FindOffsetsInRange(
		range,
		candidateOffset,
		semanticSearchEnd,
		kOutputWriteSemanticPattern);
	if (semanticMatches.size() != 1 ||
		semanticMatches.front() < candidateOffset ||
		semanticMatches.front() - candidateOffset > kSemanticDistanceLimit) {
		return false;
	}

	const size_t returnSearchBegin = semanticMatches.front() + semanticLength;
	const size_t returnSearchEnd = (std::min)(
		range.size,
		returnSearchBegin + kReturnSearchBytes);
	return !FindOffsetsInRange(
		range,
		returnSearchBegin,
		returnSearchEnd,
		kThiscallReturnPattern).empty();
}

ResolveResult ResolveFromCodeRanges(const std::vector<CodeRange>& ranges)
{
	ResolveResult result;
	const auto entryMatches = FindAllMatches(ranges, kOutputFunctionEntryPattern);
	if (entryMatches.size() == 1) {
		const PatternMatch& match = entryMatches.front();
		if (ValidateCandidate(ranges[match.rangeIndex], match.offset)) {
			result.address = ranges[match.rangeIndex].runtimeBase + match.offset;
			result.method = "entry_pattern";
			result.diagnostics = "validated unique entry pattern";
			return result;
		}
	}

	// 未知版本可能只改变入口附近的寄存器或分支布局。此时以唯一写入语义为锚点，
	// 最多向前 0x300 字节寻找唯一标准 SEH 函数序言，再执行同样的完整校验。
	const auto semanticMatches = FindAllMatches(ranges, kOutputWriteSemanticPattern);
	if (semanticMatches.size() == 1) {
		const PatternMatch& semanticMatch = semanticMatches.front();
		const CodeRange& range = ranges[semanticMatch.rangeIndex];
		const size_t searchBegin = semanticMatch.offset > kSemanticDistanceLimit
			? semanticMatch.offset - kSemanticDistanceLimit
			: 0;
		const auto prologueMatches = FindOffsetsInRange(
			range,
			searchBegin,
			semanticMatch.offset,
			kSehFunctionProloguePattern);

		std::vector<size_t> validatedPrologues;
		for (const size_t offset : prologueMatches) {
			if (offset < semanticMatch.offset && ValidateCandidate(range, offset)) {
				validatedPrologues.push_back(offset);
			}
		}
		if (validatedPrologues.size() == 1) {
			result.address = range.runtimeBase + validatedPrologues.front();
			result.method = "semantic_backtrack";
			result.diagnostics = "validated unique output semantic and SEH prologue";
			return result;
		}
	}

	result.diagnostics = std::format(
		"unavailable entry_matches={} semantic_matches={}",
		entryMatches.size(),
		semanticMatches.size());
	return result;
}

std::uintptr_t ResolveDebugOutputCallReturnAddress(
	const std::vector<CodeRange>& ranges,
	std::uintptr_t outputFunctionAddress)
{
	if (outputFunctionAddress == 0) {
		return 0;
	}

	std::vector<std::uintptr_t> validatedReturns;
	for (size_t rangeIndex = 0; rangeIndex < ranges.size(); ++rangeIndex) {
		const CodeRange& range = ranges[rangeIndex];
		const auto matches = FindOffsetsInRange(range, 0, range.size, kDebugOutputCallSitePattern);
		for (const size_t offset : matches) {
			const size_t callOffset = offset + kDebugOutputCallInstructionOffset;
			if (range.data == nullptr ||
				callOffset + 5 > range.size ||
				range.data[callOffset] != 0xE8) {
				continue;
			}

			std::int32_t displacement = 0;
			std::memcpy(&displacement, range.data + callOffset + 1, sizeof(displacement));
			const std::uintptr_t returnAddress = range.runtimeBase + callOffset + 5;
			const std::intptr_t callTarget = static_cast<std::intptr_t>(returnAddress) + displacement;
			if (callTarget == static_cast<std::intptr_t>(outputFunctionAddress)) {
				validatedReturns.push_back(range.runtimeBase + offset + kDebugOutputCallReturnOffset);
			}
		}
	}

	return validatedReturns.size() == 1 ? validatedReturns.front() : 0;
}

std::vector<CodeRange> GetMainExecutableCodeRanges()
{
	std::vector<CodeRange> ranges;
#if defined(_M_IX86)
	const HMODULE module = GetModuleHandleW(nullptr);
	if (module == nullptr) {
		return ranges;
	}

	const auto moduleBase = reinterpret_cast<const byte*>(module);
	const auto dosHeader = reinterpret_cast<const IMAGE_DOS_HEADER*>(moduleBase);
	if (dosHeader->e_magic != IMAGE_DOS_SIGNATURE || dosHeader->e_lfanew <= 0) {
		return ranges;
	}
	const auto ntHeaders = reinterpret_cast<const IMAGE_NT_HEADERS*>(
		moduleBase + static_cast<size_t>(dosHeader->e_lfanew));
	if (ntHeaders->Signature != IMAGE_NT_SIGNATURE ||
		ntHeaders->OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR32_MAGIC ||
		ntHeaders->OptionalHeader.SizeOfImage == 0 ||
		ntHeaders->FileHeader.NumberOfSections == 0 ||
		ntHeaders->FileHeader.NumberOfSections > 96) {
		return ranges;
	}

	const size_t imageSize = ntHeaders->OptionalHeader.SizeOfImage;
	const IMAGE_SECTION_HEADER* section = IMAGE_FIRST_SECTION(ntHeaders);
	for (WORD index = 0; index < ntHeaders->FileHeader.NumberOfSections; ++index, ++section) {
		if ((section->Characteristics & IMAGE_SCN_MEM_EXECUTE) == 0 ||
			section->VirtualAddress >= imageSize) {
			continue;
		}
		const size_t requestedSize = (std::max)(
			static_cast<size_t>(section->Misc.VirtualSize),
			static_cast<size_t>(section->SizeOfRawData));
		const size_t availableSize = imageSize - section->VirtualAddress;
		const size_t sectionSize = (std::min)(requestedSize, availableSize);
		if (sectionSize == 0) {
			continue;
		}
		const byte* sectionData = moduleBase + section->VirtualAddress;
		ranges.push_back({
			sectionData,
			sectionSize,
			reinterpret_cast<std::uintptr_t>(sectionData)
		});
	}
#endif
	return ranges;
}

class CaptureStore final {
public:
	SessionId Begin()
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		++m_generation;
		if (m_generation == 0) {
			++m_generation;
		}
		m_activeSession = m_generation;
		m_text.clear();
		m_truncated = false;
		return m_activeSession;
	}

	void Append(const char* text, size_t length)
	{
		if (text == nullptr || length == 0) {
			return;
		}
		std::lock_guard<std::mutex> lock(m_mutex);
		if (m_activeSession == 0) {
			return;
		}
		const size_t remaining = kMaxCaptureBytes - m_text.size();
		const size_t copyLength = (std::min)(remaining, length);
		if (copyLength > 0) {
			m_text.append(text, copyLength);
		}
		if (copyLength < length) {
			m_truncated = true;
		}
	}

	CaptureSnapshot Snapshot(SessionId sessionId)
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (sessionId == 0 || sessionId != m_activeSession) {
			return {};
		}
		return {m_text, m_truncated};
	}

	CaptureSnapshot End(SessionId sessionId)
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (sessionId == 0 || sessionId != m_activeSession) {
			return {};
		}
		CaptureSnapshot result{std::move(m_text), m_truncated};
		m_text.clear();
		m_truncated = false;
		m_activeSession = 0;
		return result;
	}

	void Cancel(SessionId sessionId)
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (sessionId == 0 || sessionId != m_activeSession) {
			return;
		}
		m_text.clear();
		m_truncated = false;
		m_activeSession = 0;
	}

private:
	std::mutex m_mutex;
	SessionId m_generation = 0;
	SessionId m_activeSession = 0;
	std::string m_text;
	bool m_truncated = false;
};

struct PendingDebugOutputBatch {
	void* thisPtr = nullptr;
	std::string text;
};

class DebugOutputQueue {
public:
	explicit DebugOutputQueue(size_t maxBytes = kDebugOutputQueueMaxBytes)
		: m_maxBytes(maxBytes)
	{
	}

	bool Enqueue(void* thisPtr, const char* text, size_t length, int appendMode)
	{
		if (thisPtr == nullptr || text == nullptr || length == 0) {
			return false;
		}

		std::string normalized(text, length);
		if (appendMode != 0 && normalized.back() != '\n') {
			normalized.append("\r\n");
		}
		if (normalized.size() > m_maxBytes || m_bytes > m_maxBytes - normalized.size()) {
			return false;
		}
		const size_t normalizedSize = normalized.size();

		if (!m_entries.empty() &&
			m_entries.back().thisPtr == thisPtr &&
			m_entries.back().text.size() + normalized.size() <= kDebugOutputUiBatchBytes) {
			m_entries.back().text.append(normalized);
		}
		else {
			m_entries.push_back({thisPtr, std::move(normalized)});
		}
		m_bytes += normalizedSize;
		return true;
	}

	bool MatchesActiveContext(void* thisPtr) const noexcept
	{
		return thisPtr != nullptr && !m_entries.empty() && m_entries.front().thisPtr == thisPtr;
	}

	PendingDebugOutputBatch TakeBatch() noexcept
	{
		if (m_entries.empty()) {
			return {};
		}

		PendingDebugOutputBatch batch = std::move(m_entries.front());
		m_bytes -= batch.text.size();
		m_entries.pop_front();
		return batch;
	}

	bool Empty() const noexcept
	{
		return m_entries.empty();
	}

	size_t SizeBytes() const noexcept
	{
		return m_bytes;
	}

	void Clear() noexcept
	{
		m_entries.clear();
		m_bytes = 0;
	}

private:
	size_t m_maxBytes = 0;
	size_t m_bytes = 0;
	std::deque<PendingDebugOutputBatch> m_entries;
};

CaptureStore g_captureStore;
std::atomic_bool g_hookAvailable = false;
std::atomic_bool g_captureHookEnabled = false;
std::atomic_bool g_debugOutputOptimizationEnabled = false;
bool g_hookAttachQueued = false;
std::uintptr_t g_resolvedAddress = 0;
std::uintptr_t g_debugOutputCallReturnAddress = 0;
std::string g_resolutionMethod;

std::mutex g_debugOutputQueueMutex;
DebugOutputQueue g_debugOutputQueue;
bool g_debugOutputFlushScheduled = false;
HWND g_debugOutputTimerWindow = nullptr;
std::atomic_uint64_t g_debugOutputQueuedEntries = 0;
std::atomic_uint64_t g_debugOutputQueuedBytes = 0;
std::atomic_uint64_t g_debugOutputFlushedBatches = 0;
std::atomic_uint64_t g_debugOutputLastEnqueueTick = 0;
std::atomic_uint g_debugOutputFlushMessage = 0;

UINT GetDebugOutputFlushMessage() noexcept
{
	UINT message = g_debugOutputFlushMessage.load(std::memory_order_acquire);
	if (message != 0) {
		return message;
	}

	message = RegisterWindowMessageW(kDebugOutputFlushMessageName);
	if (message != 0) {
		g_debugOutputFlushMessage.store(message, std::memory_order_release);
	}
	return message;
}

void ClearPendingDebugOutput() noexcept
{
	std::lock_guard<std::mutex> lock(g_debugOutputQueueMutex);
	g_debugOutputQueue.Clear();
	g_debugOutputFlushScheduled = false;
	g_debugOutputLastEnqueueTick.store(0, std::memory_order_relaxed);
}

#if defined(_M_IX86)
using OriginalOutputFunction = void(__thiscall*)(void* thisPtr, const char* text, int appendMode);
OriginalOutputFunction g_originalOutputFunction = nullptr;

size_t SafeBoundedStringLength(const char* text, size_t maxLength) noexcept
{
	if (text == nullptr || maxLength == 0) {
		return 0;
	}
	__try {
		size_t length = 0;
		while (length < maxLength && text[length] != '\0') {
			++length;
		}
		return length;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

bool TryQueueDebugOutput(
	void* thisPtr,
	const char* text,
	size_t length,
	int appendMode,
	const void* returnAddress) noexcept
{
	if (!g_debugOutputOptimizationEnabled.load(std::memory_order_acquire) ||
		g_debugOutputCallReturnAddress == 0 ||
		reinterpret_cast<std::uintptr_t>(returnAddress) != g_debugOutputCallReturnAddress ||
		length < 2 || text[0] != '*' || text[1] != ' ' ||
		g_hwnd == nullptr || !IsWindow(g_hwnd)) {
		return false;
	}

	const UINT flushMessage = GetDebugOutputFlushMessage();
	if (flushMessage == 0) {
		return false;
	}

	bool shouldPost = false;
	const size_t queuedBytes = length +
		(appendMode != 0 && text[length - 1] != '\n' ? 2 : 0);
	try {
		{
			std::lock_guard<std::mutex> lock(g_debugOutputQueueMutex);
			if (!g_debugOutputOptimizationEnabled.load(std::memory_order_relaxed) ||
				!g_debugOutputQueue.Enqueue(thisPtr, text, length, appendMode)) {
				return false;
			}
			if (!g_debugOutputFlushScheduled) {
				g_debugOutputFlushScheduled = true;
				shouldPost = true;
				g_debugOutputQueuedEntries.store(1, std::memory_order_relaxed);
				g_debugOutputQueuedBytes.store(queuedBytes, std::memory_order_relaxed);
				g_debugOutputFlushedBatches.store(0, std::memory_order_relaxed);
			}
			else {
				g_debugOutputQueuedEntries.fetch_add(1, std::memory_order_relaxed);
				g_debugOutputQueuedBytes.fetch_add(queuedBytes, std::memory_order_relaxed);
			}
			g_debugOutputLastEnqueueTick.store(GetTickCount64(), std::memory_order_release);
		}
	}
	catch (...) {
		return false;
	}

	if (shouldPost && !PostMessageW(g_hwnd, flushMessage, 0, 0)) {
		ClearPendingDebugOutput();
		return false;
	}
	return true;
}

bool TryQueueFollowingDebugOutput(
	void* thisPtr,
	const char* text,
	size_t length,
	int appendMode) noexcept
{
	if (thisPtr == nullptr || text == nullptr || length == 0) {
		return false;
	}

	const size_t queuedBytes = length +
		(appendMode != 0 && text[length - 1] != '\n' ? 2 : 0);
	try {
		std::lock_guard<std::mutex> lock(g_debugOutputQueueMutex);
		if (!g_debugOutputOptimizationEnabled.load(std::memory_order_relaxed) ||
			!g_debugOutputFlushScheduled ||
			!g_debugOutputQueue.MatchesActiveContext(thisPtr) ||
			!g_debugOutputQueue.Enqueue(thisPtr, text, length, appendMode)) {
			return false;
		}

		g_debugOutputQueuedEntries.fetch_add(1, std::memory_order_relaxed);
		g_debugOutputQueuedBytes.fetch_add(queuedBytes, std::memory_order_relaxed);
		g_debugOutputLastEnqueueTick.store(GetTickCount64(), std::memory_order_release);
		return true;
	}
	catch (...) {
		return false;
	}
}

BOOL CALLBACK FindNativeOutputControl(HWND window, LPARAM parameter) noexcept
{
	if (GetDlgCtrlID(window) != 1011) {
		return TRUE;
	}
	auto* outputWindow = reinterpret_cast<HWND*>(parameter);
	*outputWindow = window;
	return FALSE;
}

HWND GetNativeOutputControl() noexcept
{
	if (g_hwnd == nullptr || !IsWindow(g_hwnd)) {
		return nullptr;
	}
	HWND outputWindow = nullptr;
	EnumChildWindows(
		g_hwnd,
		FindNativeOutputControl,
		reinterpret_cast<LPARAM>(&outputWindow));
	return outputWindow;
}

bool AppendDebugOutputToNativeControl(const std::string& text) noexcept
{
	if (text.empty()) {
		return true;
	}
	const HWND outputWindow = GetNativeOutputControl();
	if (outputWindow == nullptr || !IsWindow(outputWindow)) {
		return false;
	}

	const int originalLength = GetWindowTextLengthA(outputWindow);
	if (originalLength < 0) {
		return false;
	}
	const bool trimOldLines = originalLength >= kDebugOutputNativeRollingThreshold;
	if (trimOldLines) {
		SendMessageA(outputWindow, WM_SETREDRAW, FALSE, 0);
		const int preferredTrimEnd = static_cast<int>((std::min)(
			static_cast<size_t>(originalLength),
			text.size()));
		int trimEnd = originalLength;
		if (preferredTrimEnd < originalLength) {
			const LRESULT line = SendMessageA(
				outputWindow,
				EM_LINEFROMCHAR,
				static_cast<WPARAM>(preferredTrimEnd),
				0);
			const LRESULT nextLineStart = line >= 0
				? SendMessageA(outputWindow, EM_LINEINDEX, static_cast<WPARAM>(line + 1), 0)
				: -1;
			if (nextLineStart > 0 && nextLineStart <= originalLength) {
				trimEnd = static_cast<int>(nextLineStart);
			}
		}
		SendMessageA(outputWindow, EM_SETSEL, 0, static_cast<LPARAM>(trimEnd));
		SendMessageA(outputWindow, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(""));
	}

	const int appendPosition = GetWindowTextLengthA(outputWindow);
	SendMessageA(
		outputWindow,
		EM_SETSEL,
		static_cast<WPARAM>(appendPosition),
		static_cast<LPARAM>(appendPosition));
	SendMessageA(
		outputWindow,
		EM_REPLACESEL,
		FALSE,
		reinterpret_cast<LPARAM>(text.c_str()));
	SendMessageA(outputWindow, EM_EMPTYUNDOBUFFER, 0, 0);

	const int finalLength = GetWindowTextLengthA(outputWindow);
	if (trimOldLines) {
		SendMessageA(outputWindow, WM_SETREDRAW, TRUE, 0);
		InvalidateRect(outputWindow, nullptr, TRUE);
	}
	if (finalLength <= appendPosition) {
		return false;
	}
	SendMessageA(
		outputWindow,
		EM_SETSEL,
		static_cast<WPARAM>(finalLength),
		static_cast<LPARAM>(finalLength));
	SendMessageA(outputWindow, EM_SCROLLCARET, 0, 0);
	return true;
}

bool FlushPendingDebugOutput() noexcept
{
	PendingDebugOutputBatch batch;
	{
		std::lock_guard<std::mutex> lock(g_debugOutputQueueMutex);
		batch = g_debugOutputQueue.TakeBatch();
		if (batch.text.empty()) {
			g_debugOutputFlushScheduled = false;
			return false;
		}
	}

	try {
		if (!AppendDebugOutputToNativeControl(batch.text)) {
			Logger::Instance().Write(
				"CompileOutputCapture",
				std::format("failed to append debug output batch bytes={}", batch.text.size()));
		}
		g_debugOutputFlushedBatches.fetch_add(1, std::memory_order_relaxed);
	}
	catch (...) {
	}

	bool hasPending = false;
	{
		std::lock_guard<std::mutex> lock(g_debugOutputQueueMutex);
		hasPending = !g_debugOutputQueue.Empty();
		if (!hasPending) {
			g_debugOutputFlushScheduled = false;
		}
	}
	if (!hasPending) {
		try {
			const HWND outputWindow = GetNativeOutputControl();
			Logger::Instance().Write(
				"CompileOutputCapture",
				std::format(
					"debug output flushed entries={} bytes={} batches={} final_chars={}",
					g_debugOutputQueuedEntries.load(std::memory_order_relaxed),
					g_debugOutputQueuedBytes.load(std::memory_order_relaxed),
					g_debugOutputFlushedBatches.load(std::memory_order_relaxed),
					outputWindow != nullptr ? GetWindowTextLengthA(outputWindow) : -1));
		}
		catch (...) {
		}
	}
	return hasPending;
}

void DrainPendingDebugOutput() noexcept
{
	while (FlushPendingDebugOutput()) {
	}
}

LRESULT CALLBACK DebugOutputTimerWindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_TIMER && wParam == kDebugOutputTimerId) {
		KillTimer(window, kDebugOutputTimerId);
		const ULONGLONG now = GetTickCount64();
		const ULONGLONG lastEnqueue =
			g_debugOutputLastEnqueueTick.load(std::memory_order_acquire);
		const ULONGLONG idleMs = now >= lastEnqueue ? now - lastEnqueue : 0;
		if (lastEnqueue != 0 && idleMs < kDebugOutputCoalesceMs) {
			const UINT remainingMs = static_cast<UINT>(kDebugOutputCoalesceMs - idleMs);
			if (SetTimer(
				window,
				kDebugOutputTimerId,
				(std::max)(remainingMs, 1u),
				nullptr) == 0) {
				DrainPendingDebugOutput();
			}
			return 0;
		}
		if (FlushPendingDebugOutput() &&
			SetTimer(window, kDebugOutputTimerId, kDebugOutputContinueMs, nullptr) == 0) {
			// 极端资源不足时仍保证日志完整，不让已经从调试线程接收的内容滞留或丢失。
			DrainPendingDebugOutput();
		}
		return 0;
	}
	if (message == WM_NCDESTROY && g_debugOutputTimerWindow == window) {
		g_debugOutputTimerWindow = nullptr;
	}
	return DefWindowProcW(window, message, wParam, lParam);
}

HMODULE GetCurrentModuleHandle() noexcept
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
			GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandle),
		&module);
	return module;
}

bool EnsureDebugOutputTimerWindow() noexcept
{
	if (g_debugOutputTimerWindow != nullptr && IsWindow(g_debugOutputTimerWindow)) {
		return true;
	}

	const HMODULE module = GetCurrentModuleHandle();
	if (module == nullptr) {
		return false;
	}
	WNDCLASSEXW windowClass = {};
	windowClass.cbSize = sizeof(windowClass);
	windowClass.lpfnWndProc = DebugOutputTimerWindowProc;
	windowClass.hInstance = module;
	windowClass.lpszClassName = kDebugOutputTimerWindowClass;
	if (RegisterClassExW(&windowClass) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
		return false;
	}

	g_debugOutputTimerWindow = CreateWindowExW(
		0,
		kDebugOutputTimerWindowClass,
		L"",
		0,
		0,
		0,
		0,
		0,
		HWND_MESSAGE,
		nullptr,
		module,
		nullptr);
	return g_debugOutputTimerWindow != nullptr;
}

bool ScheduleDebugOutputFlush(UINT delayMs) noexcept
{
	return EnsureDebugOutputTimerWindow() &&
		SetTimer(g_debugOutputTimerWindow, kDebugOutputTimerId, delayMs, nullptr) != 0;
}

void __fastcall HookOutputFunction(
	void* thisPtr,
	void* /*dummy*/,
	const char* text,
	int appendMode)
{
	if (!g_captureHookEnabled.load(std::memory_order_acquire)) {
		g_originalOutputFunction(thisPtr, text, appendMode);
		return;
	}

	const size_t length = SafeBoundedStringLength(text, kMaxTextProbeBytes);
	if (length > 0) {
		// 内部 Hook 仅服务编译会话；日志中心统一从实际输出控件采集，避免重复记录。
		try {
			g_captureStore.Append(text, length);
		}
		catch (...) {
		}
	}
	if (length > 0 &&
		g_debugOutputOptimizationEnabled.load(std::memory_order_acquire)) {
		const void* const returnAddress = _ReturnAddress();
		if (TryQueueDebugOutput(
				thisPtr,
				text,
				length,
				appendMode,
				returnAddress) ||
			TryQueueFollowingDebugOutput(thisPtr, text, length, appendMode)) {
			return;
		}
	}
	g_originalOutputFunction(thisPtr, text, appendMode);
}
#endif

std::vector<byte> MaterializePattern(const char* pattern, byte wildcardValue)
{
	std::vector<byte> bytes;
	std::istringstream input(pattern != nullptr ? pattern : "");
	std::string token;
	while (input >> token) {
		if (token == "??") {
			bytes.push_back(wildcardValue);
			continue;
		}
		unsigned int value = 0;
		std::istringstream hex(token);
		hex >> std::hex >> value;
		bytes.push_back(static_cast<byte>(value & 0xFF));
	}
	return bytes;
}

void WriteBytes(std::vector<byte>& target, size_t offset, const std::vector<byte>& source)
{
	if (offset > target.size() || source.size() > target.size() - offset) {
		return;
	}
	std::copy(source.begin(), source.end(), target.begin() + offset);
}

void WriteValidatedFunction(
	std::vector<byte>& target,
	size_t entryOffset,
	bool fullEntry,
	bool includeReturn)
{
	const auto entry = MaterializePattern(
		fullEntry ? kOutputFunctionEntryPattern : kSehFunctionProloguePattern,
		0x11);
	const auto semantic = MaterializePattern(kOutputWriteSemanticPattern, 0x22);
	const auto functionReturn = MaterializePattern(kThiscallReturnPattern, 0x00);
	WriteBytes(target, entryOffset, entry);
	const size_t semanticOffset = entryOffset + 0x1B1;
	WriteBytes(target, semanticOffset, semantic);
	if (includeReturn) {
		WriteBytes(target, semanticOffset + semantic.size() + 8, functionReturn);
	}
}

void WriteDebugOutputCallSite(
	std::vector<byte>& target,
	size_t offset,
	std::uintptr_t runtimeBase,
	std::uintptr_t outputFunctionAddress)
{
	const auto callSite = MaterializePattern(kDebugOutputCallSitePattern, 0x11);
	WriteBytes(target, offset, callSite);
	const size_t callOffset = offset + kDebugOutputCallInstructionOffset;
	if (callOffset + 5 > target.size()) {
		return;
	}
	const std::uintptr_t returnAddress = runtimeBase + callOffset + 5;
	const std::intptr_t displacementWide =
		static_cast<std::intptr_t>(outputFunctionAddress) - static_cast<std::intptr_t>(returnAddress);
	const std::int32_t displacement = static_cast<std::int32_t>(displacementWide);
	std::memcpy(target.data() + callOffset + 1, &displacement, sizeof(displacement));
}

} // namespace

bool AttachToCurrentDetourTransaction()
{
	ClearPendingDebugOutput();
	g_hookAvailable.store(false, std::memory_order_release);
	g_hookAttachQueued = false;
	g_resolvedAddress = 0;
	g_debugOutputCallReturnAddress = 0;
	g_resolutionMethod.clear();
	if (!g_captureHookEnabled.load(std::memory_order_acquire)) {
		Logger::Instance().Write("CompileOutputCapture", "hook disabled by configuration; installation skipped");
		return false;
	}

#if !defined(_M_IX86)
	Logger::Instance().Write("CompileOutputCapture", "unsupported architecture; control fallback enabled");
	return false;
#else
	const std::vector<CodeRange> codeRanges = GetMainExecutableCodeRanges();
	const ResolveResult resolved = ResolveFromCodeRanges(codeRanges);
	if (resolved.address == 0) {
		Logger::Instance().Write(
			"CompileOutputCapture",
			"IDE output hook unavailable: " + resolved.diagnostics + "; control fallback enabled");
		return false;
	}

	g_resolvedAddress = resolved.address;
	g_debugOutputCallReturnAddress = ResolveDebugOutputCallReturnAddress(codeRanges, resolved.address);
	g_resolutionMethod = resolved.method;
	g_originalOutputFunction = reinterpret_cast<OriginalOutputFunction>(resolved.address);
	const LONG error = DetourAttach(
		reinterpret_cast<PVOID*>(&g_originalOutputFunction),
		HookOutputFunction);
	if (error != NO_ERROR) {
		Logger::Instance().Write(
			"CompileOutputCapture",
			std::format(
				"DetourAttach failed error={} address=0x{:X}; control fallback enabled",
				error,
				resolved.address));
		g_originalOutputFunction = nullptr;
		g_resolvedAddress = 0;
		g_debugOutputCallReturnAddress = 0;
		g_resolutionMethod.clear();
		return false;
	}

	g_hookAttachQueued = true;
	return true;
#endif
}

void CompleteHookInstallation(bool transactionCommitted)
{
	const bool installed = transactionCommitted && g_hookAttachQueued;
	g_hookAvailable.store(installed, std::memory_order_release);
	if (installed) {
		const std::uintptr_t moduleBase = reinterpret_cast<std::uintptr_t>(GetModuleHandleW(nullptr));
		const std::uintptr_t rva = moduleBase != 0 && g_resolvedAddress >= moduleBase
			? g_resolvedAddress - moduleBase
			: 0;
		const std::uintptr_t debugReturnRva = moduleBase != 0 && g_debugOutputCallReturnAddress >= moduleBase
			? g_debugOutputCallReturnAddress - moduleBase
			: 0;
		Logger::Instance().Write(
			"CompileOutputCapture",
			std::format(
				"IDE output hook installed method={} rva=0x{:X} fast_debug={} debug_return_rva=0x{:X}",
				g_resolutionMethod,
				rva,
				debugReturnRva != 0 ? 1 : 0,
				debugReturnRva));
	}
	else if (g_hookAttachQueued) {
		Logger::Instance().Write(
			"CompileOutputCapture",
			"Detours transaction failed; control fallback enabled");
	}
	g_hookAttachQueued = false;
}

bool IsHookAvailable()
{
	return g_captureHookEnabled.load(std::memory_order_acquire) &&
		g_hookAvailable.load(std::memory_order_acquire);
}

void SetCaptureHookEnabled(bool enabled) noexcept
{
	const bool wasEnabled = g_captureHookEnabled.exchange(enabled, std::memory_order_acq_rel);
	if (!enabled && wasEnabled) {
		SetDebugOutputOptimizationEnabled(false);
		ClearPendingDebugOutput();
	}
}

bool IsCaptureHookEnabled() noexcept
{
	return g_captureHookEnabled.load(std::memory_order_acquire);
}

void SetDebugOutputOptimizationEnabled(bool enabled) noexcept
{
	enabled = enabled && g_captureHookEnabled.load(std::memory_order_acquire);
	const bool wasEnabled = g_debugOutputOptimizationEnabled.exchange(
		enabled,
		std::memory_order_acq_rel);
#if defined(_M_IX86)
	if (!enabled && wasEnabled) {
		// 先关闭入队门，再在 IDE 主线程排空既有内容，保证切换时不丢日志。
		DrainPendingDebugOutput();
		if (g_debugOutputTimerWindow != nullptr && IsWindow(g_debugOutputTimerWindow)) {
			KillTimer(g_debugOutputTimerWindow, kDebugOutputTimerId);
			DestroyWindow(g_debugOutputTimerWindow);
		}
		g_debugOutputTimerWindow = nullptr;
	}
#else
	(void)wasEnabled;
#endif
}

bool IsDebugOutputOptimizationEnabled() noexcept
{
	return g_debugOutputOptimizationEnabled.load(std::memory_order_acquire);
}

SessionId BeginCapture()
{
	if (!IsHookAvailable()) {
		return 0;
	}
	return g_captureStore.Begin();
}

CaptureSnapshot SnapshotCapture(SessionId sessionId)
{
	return g_captureStore.Snapshot(sessionId);
}

CaptureSnapshot EndCapture(SessionId sessionId)
{
	return g_captureStore.End(sessionId);
}

void CancelCapture(SessionId sessionId)
{
	g_captureStore.Cancel(sessionId);
}

bool HandleMainWindowMessage(HWND window, UINT message, WPARAM wParam, LPARAM /*lParam*/) noexcept
{
#if !defined(_M_IX86)
	(void)window;
	(void)message;
	(void)wParam;
	return false;
#else
	if (!IsDebugOutputOptimizationEnabled()) {
		const UINT registeredMessage = g_debugOutputFlushMessage.load(std::memory_order_acquire);
		return registeredMessage != 0 && message == registeredMessage;
	}
	const UINT flushMessage = GetDebugOutputFlushMessage();
	if (flushMessage != 0 && message == flushMessage) {
		if (!ScheduleDebugOutputFlush(kDebugOutputCoalesceMs)) {
			DrainPendingDebugOutput();
		}
		return true;
	}
	(void)window;
	(void)wParam;
	return false;
#endif
}

void Shutdown(HWND window) noexcept
{
	(void)window;
	g_debugOutputOptimizationEnabled.store(false, std::memory_order_release);
	if (g_debugOutputTimerWindow != nullptr && IsWindow(g_debugOutputTimerWindow)) {
		KillTimer(g_debugOutputTimerWindow, kDebugOutputTimerId);
		DestroyWindow(g_debugOutputTimerWindow);
	}
	g_debugOutputTimerWindow = nullptr;
	ClearPendingDebugOutput();
}

std::string BuildSelfTestJson()
{
	const auto runResolve = [](std::vector<byte>& code, std::uintptr_t base) {
		return ResolveFromCodeRanges({CodeRange{code.data(), code.size(), base}});
	};

	constexpr std::uintptr_t primaryBase = 0x500000;
	constexpr size_t primaryOutputOffset = 0x20;
	constexpr size_t primaryDebugCallOffset = 0x3C0;
	std::vector<byte> primaryCode(0x500, 0x90);
	WriteValidatedFunction(primaryCode, primaryOutputOffset, true, true);
	WriteDebugOutputCallSite(
		primaryCode,
		primaryDebugCallOffset,
		primaryBase,
		primaryBase + primaryOutputOffset);
	const ResolveResult primary = runResolve(primaryCode, primaryBase);
	const bool primaryResolved = primary.address == primaryBase + primaryOutputOffset &&
		primary.method == "entry_pattern";
	const std::uintptr_t debugCallReturn = ResolveDebugOutputCallReturnAddress(
		{CodeRange{primaryCode.data(), primaryCode.size(), primaryBase}},
		primary.address);
	const bool debugCallSiteResolved =
		debugCallReturn == primaryBase + primaryDebugCallOffset + kDebugOutputCallReturnOffset;

	std::vector<byte> wrongCallTargetCode(0x100, 0x90);
	WriteDebugOutputCallSite(
		wrongCallTargetCode,
		0x20,
		0x510000,
		0x510080);
	const bool wrongDebugCallTargetRejected = ResolveDebugOutputCallReturnAddress(
		{CodeRange{wrongCallTargetCode.data(), wrongCallTargetCode.size(), 0x510000}},
		primary.address) == 0;

	std::vector<byte> fallbackCode(0x500, 0x90);
	WriteValidatedFunction(fallbackCode, 0x40, false, true);
	const ResolveResult fallback = runResolve(fallbackCode, 0x600000);
	const bool fallbackResolved = fallback.address == 0x600040 && fallback.method == "semantic_backtrack";

	std::vector<byte> ambiguousCode(0x900, 0x90);
	WriteValidatedFunction(ambiguousCode, 0x20, true, true);
	WriteValidatedFunction(ambiguousCode, 0x480, true, true);
	const bool ambiguousRejected = runResolve(ambiguousCode, 0x700000).address == 0;

	std::vector<byte> invalidReturnCode(0x500, 0x90);
	WriteValidatedFunction(invalidReturnCode, 0x20, true, false);
	const bool invalidReturnRejected = runResolve(invalidReturnCode, 0x800000).address == 0;

	CaptureStore store;
	const SessionId session = store.Begin();
	store.Append("first", 5);
	const bool wrongSessionRejected = store.End(session + 1).text.empty();
	const CaptureSnapshot firstSnapshot = store.Snapshot(session);
	store.Append("-second", 7);
	const CaptureSnapshot finished = store.End(session);
	const bool sessionIsolationPassed = wrongSessionRejected &&
		firstSnapshot.text == "first" && finished.text == "first-second";

	const SessionId cappedSession = store.Begin();
	const std::string oversized(kMaxCaptureBytes + 1, 'x');
	store.Append(oversized.data(), oversized.size());
	const CaptureSnapshot capped = store.End(cappedSession);
	const bool capPassed = capped.truncated && capped.text.size() == kMaxCaptureBytes;

	DebugOutputQueue debugQueue(64);
	void* const firstThisPtr = reinterpret_cast<void*>(static_cast<std::uintptr_t>(1));
	void* const secondThisPtr = reinterpret_cast<void*>(static_cast<std::uintptr_t>(2));
	const bool firstQueued = debugQueue.Enqueue(firstThisPtr, "* 1", 3, 1);
	const bool secondQueued = debugQueue.Enqueue(firstThisPtr, "* 2\n", 4, 1);
	const bool statusQueued = debugQueue.Enqueue(firstThisPtr, "finished", 8, 1);
	const bool activeContextMatched = debugQueue.MatchesActiveContext(firstThisPtr) &&
		!debugQueue.MatchesActiveContext(secondThisPtr);
	const bool otherContextQueued = debugQueue.Enqueue(secondThisPtr, "other", 5, 0);
	const PendingDebugOutputBatch firstBatch = debugQueue.TakeBatch();
	const bool nextContextMatched = debugQueue.MatchesActiveContext(secondThisPtr);
	const PendingDebugOutputBatch secondBatch = debugQueue.TakeBatch();
	const bool debugBatchingPassed = firstQueued && secondQueued && statusQueued &&
		activeContextMatched && otherContextQueued && nextContextMatched &&
		firstBatch.thisPtr == firstThisPtr && firstBatch.text == "* 1\r\n* 2\nfinished\r\n" &&
		secondBatch.thisPtr == secondThisPtr && secondBatch.text == "other" &&
		debugQueue.Empty() && debugQueue.SizeBytes() == 0;
	DebugOutputQueue cappedDebugQueue(4);
	const bool debugQueueCapPassed = !cappedDebugQueue.Enqueue(firstThisPtr, "1234", 4, 1);

	const bool ok = primaryResolved && debugCallSiteResolved && wrongDebugCallTargetRejected &&
		fallbackResolved && ambiguousRejected && invalidReturnRejected &&
		sessionIsolationPassed && capPassed && debugBatchingPassed && debugQueueCapPassed;
	return nlohmann::json({
		{"name", "ide-compile-output-capture"},
		{"ok", ok},
		{"primary_pattern_resolved", primaryResolved},
		{"debug_call_site_resolved", debugCallSiteResolved},
		{"wrong_debug_call_target_rejected", wrongDebugCallTargetRejected},
		{"semantic_backtrack_resolved", fallbackResolved},
		{"ambiguous_pattern_rejected", ambiguousRejected},
		{"missing_thiscall_return_rejected", invalidReturnRejected},
		{"session_isolation_passed", sessionIsolationPassed},
		{"capture_cap_passed", capPassed},
		{"debug_batching_passed", debugBatchingPassed},
		{"debug_queue_cap_passed", debugQueueCapPassed}
	}).dump();
}

} // namespace IdeCompileOutputCapture
