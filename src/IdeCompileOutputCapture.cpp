#include "IdeCompileOutputCapture.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <format>
#include <mutex>
#include <sstream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include <detours.h>

#include "..\\thirdparty\\json.hpp"
#include "Logger.h"
#include "MemFind.h"

namespace IdeCompileOutputCapture {
namespace {

constexpr size_t kMaxCaptureBytes = 2 * 1024 * 1024;
constexpr size_t kMaxTextProbeBytes = kMaxCaptureBytes + 1;
constexpr size_t kSemanticDistanceLimit = 0x300;
constexpr size_t kReturnSearchBytes = 0x100;

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

CaptureStore g_captureStore;
std::atomic_bool g_hookAvailable = false;
bool g_hookAttachQueued = false;
std::uintptr_t g_resolvedAddress = 0;
std::string g_resolutionMethod;

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

void __fastcall HookOutputFunction(
	void* thisPtr,
	void* /*dummy*/,
	const char* text,
	int appendMode)
{
	const size_t length = SafeBoundedStringLength(text, kMaxTextProbeBytes);
	if (length > 0) {
		// Hook 边界绝不能把分配异常传播到 IDE；捕获失败时仍必须调用原函数。
		try {
			g_captureStore.Append(text, length);
		}
		catch (...) {
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

} // namespace

bool AttachToCurrentDetourTransaction()
{
	g_hookAvailable.store(false, std::memory_order_release);
	g_hookAttachQueued = false;
	g_resolvedAddress = 0;
	g_resolutionMethod.clear();

#if !defined(_M_IX86)
	Logger::Instance().Write("CompileOutputCapture", "unsupported architecture; control fallback enabled");
	return false;
#else
	const ResolveResult resolved = ResolveFromCodeRanges(GetMainExecutableCodeRanges());
	if (resolved.address == 0) {
		Logger::Instance().Write(
			"CompileOutputCapture",
			"IDE output hook unavailable: " + resolved.diagnostics + "; control fallback enabled");
		return false;
	}

	g_resolvedAddress = resolved.address;
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
		Logger::Instance().Write(
			"CompileOutputCapture",
			std::format(
				"IDE output hook installed method={} rva=0x{:X}",
				g_resolutionMethod,
				rva));
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
	return g_hookAvailable.load(std::memory_order_acquire);
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

std::string BuildSelfTestJson()
{
	const auto runResolve = [](std::vector<byte>& code, std::uintptr_t base) {
		return ResolveFromCodeRanges({CodeRange{code.data(), code.size(), base}});
	};

	std::vector<byte> primaryCode(0x500, 0x90);
	WriteValidatedFunction(primaryCode, 0x20, true, true);
	const ResolveResult primary = runResolve(primaryCode, 0x500000);
	const bool primaryResolved = primary.address == 0x500020 && primary.method == "entry_pattern";

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

	const bool ok = primaryResolved && fallbackResolved && ambiguousRejected &&
		invalidReturnRejected && sessionIsolationPassed && capPassed;
	return nlohmann::json({
		{"name", "ide-compile-output-capture"},
		{"ok", ok},
		{"primary_pattern_resolved", primaryResolved},
		{"semantic_backtrack_resolved", fallbackResolved},
		{"ambiguous_pattern_rejected", ambiguousRejected},
		{"missing_thiscall_return_rejected", invalidReturnRejected},
		{"session_isolation_passed", sessionIsolationPassed},
		{"capture_cap_passed", capPassed}
	}).dump();
}

} // namespace IdeCompileOutputCapture
