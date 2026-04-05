#include "EideInternalTextBridge.h"

#include <afxwin.h>
#include <Windows.h>

#include <detours.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cstring>
#include <fstream>
#include <initializer_list>
#include <mutex>
#include <string>
#include <unordered_map>
#include <vector>

#include "Global.h"
#include "MemFind.h"
#include "WindowHelper.h"
#include "direct_global_search_debug.hpp"

namespace e571 {
namespace {

constexpr std::uintptr_t kImageBase = 0x400000;
constexpr int kEditorCmdSelectAll = 0x01010023;
constexpr int kEditorCmdDeleteSelection = 0x02020003;
constexpr int kEditorCmdPaste = 0x02020005;
constexpr int kEditorCmdCopy = 0x02030001;
constexpr int kEditorCmdInsertRawText = 0x0202006D;
constexpr unsigned int kEditorUiCmdSelectAll = 0x2041;
constexpr unsigned int kEditorUiCmdCopy = 0x2043;
constexpr unsigned int kEditorUiCmdPaste = 0x2056;
constexpr unsigned int kEditorUiCmdDelete = VK_DELETE;
constexpr size_t kEditorDispatchVtableIndex = 56;
constexpr int kEditorDispatchDefaultFlags = 1;
constexpr std::uintptr_t kFallbackEditorDispatchCommandRva = 0x4B6D70;
constexpr std::uintptr_t kKnownEditorFormatRangeTextRva = 0x48EA10;
constexpr std::uintptr_t kKnownEditorGetRangeCountRva = 0x4C05B0;
constexpr std::uintptr_t kKnownClipboardTextToObjectWrapperRva = 0x4C9220;
constexpr std::uintptr_t kKnownClipboardTextToObjectDirectRva = 0x4CB160;
constexpr std::uintptr_t kKnownClipboardInsertObjectRva = 0x48B720;
constexpr std::uintptr_t kKnownClipboardDeserializeRva = 0x4535F0;
constexpr std::uintptr_t kKnownClipboardSerializeRva = 0x45B640;
constexpr std::uintptr_t kKnownClipboardValidateA00Rva = 0x452A00;
constexpr std::uintptr_t kKnownClipboardValidate560Rva = 0x452560;
constexpr std::uintptr_t kKnownClipboardValidate230Rva = 0x452230;
constexpr std::uintptr_t kKnownClipboardCollectionFirstInvalidRva = 0x4521D0;
constexpr std::uintptr_t kFallbackEditorPasteHandlerRva = 0x4BC830;
constexpr std::uintptr_t kKnownGlobalSupportLibraryArrayRva = 0x5CB028;
constexpr std::uintptr_t kKnownSupportLibraryArrayAddUniqueRva = 0x4F4830;
constexpr std::uintptr_t kKnownClipboardMergeParsedRangeRva = 0x4DABA0;
constexpr std::uintptr_t kKnownParsedItemDestroyRva = 0x40F330;
constexpr std::uintptr_t kKnownPtrArrayRemoveAtRva = 0x53ACC0;
constexpr std::uintptr_t kKnownClipboardObjectDestroyRva = 0x401ED0;
constexpr std::uintptr_t kKnownGenericArrayInitRva = 0x486060;
constexpr std::uintptr_t kKnownGenericArrayAssignRva = 0x486920;
constexpr std::uintptr_t kKnownGenericArrayFinalizeRva = 0x4863A0;
constexpr std::uintptr_t kKnownGenericArrayDestroyRva = 0x486260;
constexpr std::uintptr_t kKnownClipboardObjectCtorRva = 0x40E400;
constexpr std::uintptr_t kKnownGenericArrayVtableRva = 0x574900;
constexpr std::uintptr_t kKnownCollectionGetValueByIndexRva = 0x4E7EA0;
constexpr size_t kInternalClipboardObjectSize = 0x340;
constexpr size_t kInternalGenericArraySize = 0x14;
constexpr size_t kClipboardObjectSupportLibrariesOffset = 0x94;
constexpr size_t kClipboardObjectGuidFlagOffset = 0x10C;
constexpr size_t kClipboardObjectGuidOffset = 0x110;

using FnEditorDispatchCommand = int(__thiscall*)(void*, int, int, char*, char*);
using FnEditorFormatRangeText = void(__thiscall*)(void*, int, int, void*, int);
using FnEditorUiCommand = int(__thiscall*)(void*, unsigned int);
using FnEditorPasteHandler = void(__thiscall*)(void*, int);
using FnClipboardTextToObjectWrapper = int(__thiscall*)(void*, void*);
using FnClipboardTextToObjectDirect = int(__thiscall*)(void*, void*);
using FnEditorInsertClipboardObject = int(__thiscall*)(void*, void*, int, int, int);
using FnThiscallVoid = void(__thiscall*)(void*);
using FnThiscallVoidInt = void(__thiscall*)(void*, int);
using FnThiscallVoidIntInt = void(__thiscall*)(void*, int, int);
using FnThiscallInt = int(__thiscall*)(void*);
using FnStdcallIntPtr = int(__stdcall*)(void*);
using FnThiscallBool = BOOL(__thiscall*)(void*);
using FnThiscallHandle = HANDLE(__thiscall*)(void*);
using FnThiscallGenericArrayAssign = void(__thiscall*)(void*, const void*, size_t);
using FnThiscallCollectionGetValueByIndex = int(__thiscall*)(void*, int, void**);
using FnCdeclCStringArrayAddUnique = int(__cdecl*)(CStringArray*, unsigned char*, int);
using FnThiscallMergeParsedRange = void(__thiscall*)(void*, void*, int, int);
using FnThiscallPtrArrayRemoveAt = void*(__thiscall*)(void*, int, int);
using FnCdeclFillMemory = void(__cdecl*)(void*, int);

struct NativeEditorCommandAddresses {
	bool initialized = false;
	bool ok = false;
	std::uintptr_t moduleBase = 0;
	std::uintptr_t editorDispatchCommand = 0;
	std::uintptr_t editorPasteHandler = 0;
};

struct EditorDispatchTargetInfo {
	std::uintptr_t rawObject = 0;
	std::uintptr_t innerObject = 0;
	unsigned int pageType = 0;
};

struct FakeClipboardContext {
	bool opened = false;
	std::unordered_map<UINT, HANDLE> formats;
	size_t openCalls = 0;
	size_t emptyCalls = 0;
	size_t setCalls = 0;
	size_t setTextCalls = 0;
	size_t getCalls = 0;
	size_t formatQueryCalls = 0;
	size_t closeCalls = 0;

	~FakeClipboardContext()
	{
		Clear();
	}

	void Clear()
	{
		for (auto& [format, handle] : formats) {
			(void)format;
			if (handle != nullptr) {
				GlobalFree(handle);
			}
		}
		formats.clear();
		opened = false;
	}
};

std::mutex g_nativeEditorCommandAddressMutex;
NativeEditorCommandAddresses g_nativeEditorCommandAddresses;
std::mutex g_pageEditTraceMutex;

std::mutex g_fakeClipboardHookMutex;
std::mutex g_fakeClipboardDataMutex;
bool g_fakeClipboardHooksInstalled = false;
FakeClipboardContext* g_activeFakeClipboard = nullptr;

decltype(&::OpenClipboard) g_originalOpenClipboard = ::OpenClipboard;
decltype(&::EmptyClipboard) g_originalEmptyClipboard = ::EmptyClipboard;
decltype(&::SetClipboardData) g_originalSetClipboardData = ::SetClipboardData;
decltype(&::GetClipboardData) g_originalGetClipboardData = ::GetClipboardData;
decltype(&::IsClipboardFormatAvailable) g_originalIsClipboardFormatAvailable = ::IsClipboardFormatAvailable;
decltype(&::CloseClipboard) g_originalCloseClipboard = ::CloseClipboard;

std::string NormalizeLineBreakToCrLf(const std::string& text);

std::uintptr_t NormalizeRuntimeAddress(std::uintptr_t runtimeAddress, std::uintptr_t moduleBase)
{
	if (runtimeAddress == 0 || moduleBase == 0 || runtimeAddress < moduleBase) {
		return 0;
	}
	return runtimeAddress - moduleBase + kImageBase;
}

std::uintptr_t ResolveUniqueCodeAddress(const char* pattern, std::uintptr_t moduleBase)
{
	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.size() != 1) {
		return 0;
	}
	return NormalizeRuntimeAddress(static_cast<std::uintptr_t>(matches.front()), moduleBase);
}

std::uintptr_t ResolveUniqueCodeAddressFromPatterns(
	const std::initializer_list<const char*> patterns,
	std::uintptr_t moduleBase)
{
	for (const char* pattern : patterns) {
		if (pattern == nullptr || *pattern == '\0') {
			continue;
		}

		const std::uintptr_t resolvedAddress = ResolveUniqueCodeAddress(pattern, moduleBase);
		if (resolvedAddress != 0) {
			return resolvedAddress;
		}
	}
	return 0;
}

bool PopulateNativeEditorCommandAddresses(NativeEditorCommandAddresses& addrs, std::uintptr_t moduleBase)
{
	addrs = {};
	addrs.moduleBase = moduleBase;
	addrs.editorDispatchCommand = ResolveUniqueCodeAddressFromPatterns(
		{
			"A1 ?? ?? ?? ?? 53 56 8B F1 85 C0 74 0A 5E B8 01 00 00 00 5B C2 10 00 8B 5C 24 10 F6 C3 10 74 0A 5E B8 01 00 00 00 5B C2 10 00 8B 4C 24 0C 33 C0 8B D1 81 E2 00 00 FF 7F",
			"8B 0D ?? ?? ?? ?? 53 56 8B F1 85 C9 74 0A 5E B8 01 00 00 00 5B C2 10 00 8B 5C 24 10 F6 C3 10 74 0A 5E B8 01 00 00 00 5B C2 10 00 8B 4C 24 0C 33 C0 8B D1 81 E2 00 00 FF 7F",
		},
		moduleBase);
	if (addrs.editorDispatchCommand == 0) {
		addrs.editorDispatchCommand = kFallbackEditorDispatchCommandRva;
	}
	addrs.editorPasteHandler = ResolveUniqueCodeAddressFromPatterns(
		{
			"6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC ?? 03 00 00 53 55 56 57 8B F9 8D 4C 24 28 89 7C 24 24 E8 ?? ?? ?? ?? 8D 4C 24 64 E8 ?? ?? ?? ?? 33 ED 8D 8C 24 BC 00 00 00 89 AC 24 ?? 03 00 00 E8 ?? ?? ?? ?? 8D 8C 24 D0 00 00 00 C6 84 24 ?? 03 00 00 01 E8 ?? ?? ?? ??",
		},
		moduleBase);
	if (addrs.editorPasteHandler == 0) {
		addrs.editorPasteHandler = kFallbackEditorPasteHandlerRva;
	}
	addrs.ok = addrs.editorDispatchCommand != 0;
	addrs.initialized = true;
	return addrs.ok;
}

const NativeEditorCommandAddresses& GetNativeEditorCommandAddresses(std::uintptr_t moduleBase)
{
	std::lock_guard<std::mutex> lock(g_nativeEditorCommandAddressMutex);
	if (!g_nativeEditorCommandAddresses.initialized || g_nativeEditorCommandAddresses.moduleBase != moduleBase) {
		PopulateNativeEditorCommandAddresses(g_nativeEditorCommandAddresses, moduleBase);
	}
	return g_nativeEditorCommandAddresses;
}

template <typename T>
T ResolveInternalAddress(std::uintptr_t moduleBase, std::uintptr_t rva)
{
	if (moduleBase == 0 || rva < kImageBase) {
		return nullptr;
	}
	return reinterpret_cast<T>(moduleBase + (rva - kImageBase));
}

std::uintptr_t ResolveInternalPointer(std::uintptr_t moduleBase, std::uintptr_t rva)
{
	if (moduleBase == 0 || rva < kImageBase) {
		return 0;
	}
	return moduleBase + (rva - kImageBase);
}

void AppendPageEditTraceLine(const std::string& text)
{
	std::lock_guard<std::mutex> lock(g_pageEditTraceMutex);

	char tempPath[MAX_PATH + 1] = {};
	DWORD pathLen = ::GetTempPathA(MAX_PATH, tempPath);
	if (pathLen == 0 || pathLen > MAX_PATH) {
		return;
	}

	SYSTEMTIME st{};
	::GetLocalTime(&st);

	std::ofstream out(
		std::string(tempPath) + "AutoLinker_page_edit_trace.log",
		std::ios::out | std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}

	out
		<< st.wYear << '-'
		<< st.wMonth << '-'
		<< st.wDay << ' '
		<< st.wHour << ':'
		<< st.wMinute << ':'
		<< st.wSecond << '.'
		<< st.wMilliseconds
		<< " | "
		<< text
		<< "\r\n";
	out.flush();
}

bool DuplicateGlobalHandle(HANDLE sourceHandle, HANDLE* outHandle)
{
	if (outHandle != nullptr) {
		*outHandle = nullptr;
	}
	if (sourceHandle == nullptr || outHandle == nullptr) {
		return false;
	}

	const SIZE_T bytes = GlobalSize(sourceHandle);
	if (bytes == 0) {
		return false;
	}

	const void* sourceData = GlobalLock(sourceHandle);
	if (sourceData == nullptr) {
		return false;
	}

	HGLOBAL duplicate = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes);
	if (duplicate == nullptr) {
		GlobalUnlock(sourceHandle);
		return false;
	}

	void* duplicateData = GlobalLock(duplicate);
	if (duplicateData == nullptr) {
		GlobalUnlock(sourceHandle);
		GlobalFree(duplicate);
		return false;
	}

	std::memcpy(duplicateData, sourceData, bytes);
	GlobalUnlock(duplicate);
	GlobalUnlock(sourceHandle);
	*outHandle = duplicate;
	return true;
}

bool CopyGlobalHandleBytesNoFree(HANDLE handle, std::vector<unsigned char>* outBytes)
{
	if (outBytes != nullptr) {
		outBytes->clear();
	}
	if (handle == nullptr || outBytes == nullptr) {
		return false;
	}

	const SIZE_T size = GlobalSize(handle);
	if (size == 0) {
		return false;
	}

	const void* data = GlobalLock(handle);
	if (data == nullptr) {
		return false;
	}

	const auto* byteData = static_cast<const unsigned char*>(data);
	outBytes->assign(byteData, byteData + size);
	GlobalUnlock(handle);
	return !outBytes->empty();
}

bool CallGenericThiscallCommandSafe(
	std::uintptr_t targetObject,
	std::uintptr_t functionAddress,
	std::uintptr_t arg1,
	std::uintptr_t arg2,
	std::uintptr_t arg3,
	std::uintptr_t arg4,
	int* outResult,
	bool* outThrew)
{
	if (outResult != nullptr) {
		*outResult = 0;
	}
	if (outThrew != nullptr) {
		*outThrew = false;
	}
	if (targetObject == 0 || functionAddress == 0) {
		return false;
	}

	using FnGenericThiscall = int(__thiscall*)(void*, std::uintptr_t, std::uintptr_t, std::uintptr_t, std::uintptr_t);
	const auto fn = reinterpret_cast<FnGenericThiscall>(functionAddress);
	__try {
		const int result = fn(
			reinterpret_cast<void*>(targetObject),
			arg1,
			arg2,
			arg3,
			arg4);
		if (outResult != nullptr) {
			*outResult = result;
		}
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		if (outThrew != nullptr) {
			*outThrew = true;
		}
		return false;
	}
}

struct CustomClipboardPayload {
	UINT format = 0;
	HANDLE handle = nullptr;

	~CustomClipboardPayload()
	{
		Reset();
	}

	CustomClipboardPayload() = default;
	CustomClipboardPayload(const CustomClipboardPayload&) = delete;
	CustomClipboardPayload& operator=(const CustomClipboardPayload&) = delete;

	CustomClipboardPayload(CustomClipboardPayload&& other) noexcept
	{
		MoveFrom(std::move(other));
	}

	CustomClipboardPayload& operator=(CustomClipboardPayload&& other) noexcept
	{
		if (this != &other) {
			Reset();
			MoveFrom(std::move(other));
		}
		return *this;
	}

	void Reset()
	{
		if (handle != nullptr) {
			GlobalFree(handle);
			handle = nullptr;
		}
		format = 0;
	}

	bool IsValid() const
	{
		return format != 0 && handle != nullptr;
	}

	HANDLE DuplicateHandle() const
	{
		HANDLE duplicate = nullptr;
		return DuplicateGlobalHandle(handle, &duplicate) ? duplicate : nullptr;
	}

	size_t ByteSize() const
	{
		return handle == nullptr ? 0 : static_cast<size_t>(GlobalSize(handle));
	}

private:
	void MoveFrom(CustomClipboardPayload&& other)
	{
		format = other.format;
		handle = other.handle;
		other.format = 0;
		other.handle = nullptr;
	}
};

BOOL WINAPI FakeClipboard_OpenClipboard(HWND hWndNewOwner)
{
	FakeClipboardContext* ctx = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		ctx = g_activeFakeClipboard;
	}
	if (ctx == nullptr) {
		return g_originalOpenClipboard(hWndNewOwner);
	}
	std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
	++ctx->openCalls;
	ctx->opened = true;
	return TRUE;
}

BOOL WINAPI FakeClipboard_EmptyClipboard()
{
	FakeClipboardContext* ctx = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		ctx = g_activeFakeClipboard;
	}
	if (ctx == nullptr) {
		return g_originalEmptyClipboard();
	}
	std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
	++ctx->emptyCalls;
	ctx->Clear();
	ctx->opened = true;
	return TRUE;
}

HANDLE WINAPI FakeClipboard_SetClipboardData(UINT uFormat, HANDLE hMem)
{
	FakeClipboardContext* ctx = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		ctx = g_activeFakeClipboard;
	}
	if (ctx == nullptr) {
		return g_originalSetClipboardData(uFormat, hMem);
	}
	std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
	++ctx->setCalls;
	if (uFormat == CF_TEXT) {
		++ctx->setTextCalls;
	}
	auto it = ctx->formats.find(uFormat);
	if (it != ctx->formats.end() && it->second != hMem && it->second != nullptr) {
		GlobalFree(it->second);
	}
	ctx->formats[uFormat] = hMem;
	return hMem;
}

HANDLE WINAPI FakeClipboard_GetClipboardData(UINT uFormat)
{
	FakeClipboardContext* ctx = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		ctx = g_activeFakeClipboard;
	}
	if (ctx == nullptr) {
		return g_originalGetClipboardData(uFormat);
	}
	std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
	++ctx->getCalls;
	const auto it = ctx->formats.find(uFormat);
	return it == ctx->formats.end() ? nullptr : it->second;
}

BOOL WINAPI FakeClipboard_CloseClipboard()
{
	FakeClipboardContext* ctx = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		ctx = g_activeFakeClipboard;
	}
	if (ctx == nullptr) {
		return g_originalCloseClipboard();
	}
	std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
	++ctx->closeCalls;
	ctx->opened = false;
	return TRUE;
}

BOOL WINAPI FakeClipboard_IsClipboardFormatAvailable(UINT format)
{
	FakeClipboardContext* ctx = nullptr;
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		ctx = g_activeFakeClipboard;
	}
	if (ctx == nullptr) {
		return g_originalIsClipboardFormatAvailable(format);
	}
	std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
	++ctx->formatQueryCalls;
	const auto it = ctx->formats.find(format);
	return (it != ctx->formats.end() && it->second != nullptr) ? TRUE : FALSE;
}

bool EnsureFakeClipboardHooksInstalled()
{
	std::lock_guard<std::mutex> lock(g_fakeClipboardHookMutex);
	if (g_fakeClipboardHooksInstalled) {
		return true;
	}

	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)g_originalOpenClipboard, FakeClipboard_OpenClipboard);
	DetourAttach(&(PVOID&)g_originalEmptyClipboard, FakeClipboard_EmptyClipboard);
	DetourAttach(&(PVOID&)g_originalSetClipboardData, FakeClipboard_SetClipboardData);
	DetourAttach(&(PVOID&)g_originalGetClipboardData, FakeClipboard_GetClipboardData);
	DetourAttach(&(PVOID&)g_originalIsClipboardFormatAvailable, FakeClipboard_IsClipboardFormatAvailable);
	DetourAttach(&(PVOID&)g_originalCloseClipboard, FakeClipboard_CloseClipboard);
	g_fakeClipboardHooksInstalled = (DetourTransactionCommit() == NO_ERROR);
	return g_fakeClipboardHooksInstalled;
}

std::vector<std::string> ParseSupportLibraryNamesFromPageCode(const std::string& pageCode);

class ScopedFakeClipboard {
public:
	ScopedFakeClipboard()
	{
		m_active = EnsureFakeClipboardHooksInstalled();
		if (m_active) {
			std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
			g_activeFakeClipboard = &m_context;
		}
	}

	~ScopedFakeClipboard()
	{
		if (m_active) {
			std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
			if (g_activeFakeClipboard == &m_context) {
				g_activeFakeClipboard = nullptr;
			}
		}
	}

	bool IsActive() const
	{
		return m_active;
	}

	bool SetFormatHandle(UINT format, HANDLE handle)
	{
		if (!m_active || handle == nullptr) {
			return false;
		}
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		m_context.formats[format] = handle;
		return true;
	}

	HANDLE GetFormatHandle(UINT format) const
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		const auto it = m_context.formats.find(format);
		return it == m_context.formats.end() ? nullptr : it->second;
	}

	bool GetFirstCustomFormat(UINT* outFormat, HANDLE* outHandle) const
	{
		if (outFormat != nullptr) {
			*outFormat = 0;
		}
		if (outHandle != nullptr) {
			*outHandle = nullptr;
		}

		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		for (const auto& [format, handle] : m_context.formats) {
			if (format != CF_TEXT && handle != nullptr) {
				if (outFormat != nullptr) {
					*outFormat = format;
				}
				if (outHandle != nullptr) {
					*outHandle = handle;
				}
				return true;
			}
		}
		return false;
	}

	bool HasCustomFormat() const
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		for (const auto& [format, handle] : m_context.formats) {
			if (handle != nullptr && format != CF_TEXT) {
				return true;
			}
		}
		return false;
	}

	bool HasClipboardTraffic() const
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		return
			m_context.openCalls != 0 ||
			m_context.emptyCalls != 0 ||
			m_context.setCalls != 0 ||
			m_context.setTextCalls != 0 ||
			m_context.getCalls != 0 ||
			m_context.formatQueryCalls != 0 ||
			m_context.closeCalls != 0;
	}

	std::string BuildStatsText() const
	{
		std::lock_guard<std::mutex> lock(g_fakeClipboardDataMutex);
		return
			"open=" + std::to_string(m_context.openCalls) +
			"|empty=" + std::to_string(m_context.emptyCalls) +
			"|set=" + std::to_string(m_context.setCalls) +
			"|set_cf_text=" + std::to_string(m_context.setTextCalls) +
			"|get=" + std::to_string(m_context.getCalls) +
			"|fmt_query=" + std::to_string(m_context.formatQueryCalls) +
			"|close=" + std::to_string(m_context.closeCalls);
	}

private:
	bool m_active = false;
	FakeClipboardContext m_context;
};

std::string NormalizeLineBreakToCrLf(const std::string& text);
std::string NormalizeLineBreakToLf(const std::string& text);
std::string NormalizePageCodeForLooseCompare(const std::string& text);
std::string NormalizePageCodeLineForStructuralCompare(const std::string& line);
std::vector<std::string> BuildPageCodeStructuralFingerprint(const std::string& text);
std::string BuildStructuralMismatchSummary(
	const std::vector<std::string>& expectedLines,
	const std::vector<std::string>& actualLines);
bool VerifyRealPageCodeMatches(
	const std::string& expectedCode,
	const std::string& actualCode,
	std::string* outMode,
	std::string* outSummary);

std::string EnsureTextUsesGbkCodePage(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, nullptr, 0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, wide.data(), wideLen) <= 0) {
		return text;
	}

	const int outLen = WideCharToMultiByte(936, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(936, 0, wide.c_str(), -1, out.data(), outLen, nullptr, nullptr) <= 0) {
		return text;
	}
	if (!out.empty() && out.back() == '\0') {
		out.pop_back();
	}
	return out;
}

std::string BuildTextMismatchSummary(const std::string& expected, const std::string& actual)
{
	const auto buildAsciiSafeSnippet = [](const std::string& text, size_t pos) {
		const size_t begin = pos > 32 ? pos - 32 : 0;
		const size_t len = (std::min<size_t>)(64, text.size() - begin);
		const std::string snippet = text.substr(begin, len);
		std::string safe;
		safe.reserve(snippet.size() * 4);
		static constexpr char kHex[] = "0123456789ABCDEF";
		for (unsigned char ch : snippet) {
			if (ch == '\r' || ch == '\n') {
				safe.push_back(' ');
				continue;
			}
			if (ch >= 0x20 && ch <= 0x7E && ch != '\\') {
				safe.push_back(static_cast<char>(ch));
				continue;
			}
			safe.push_back('\\');
			safe.push_back('x');
			safe.push_back(kHex[(ch >> 4) & 0x0F]);
			safe.push_back(kHex[ch & 0x0F]);
		}
		return safe;
	};

	const size_t mismatchPos = [&]() -> size_t {
		const size_t common = (std::min)(expected.size(), actual.size());
		for (size_t i = 0; i < common; ++i) {
			if (expected[i] != actual[i]) {
				return i;
			}
		}
		return common;
	}();

	return
		"mismatch_at=" + std::to_string(mismatchPos) +
		"|expected_bytes=" + std::to_string(expected.size()) +
		"|actual_bytes=" + std::to_string(actual.size()) +
		"|expected_snippet=" + buildAsciiSafeSnippet(expected, mismatchPos) +
		"|actual_snippet=" + buildAsciiSafeSnippet(actual, mismatchPos);
}

std::string TrimAsciiCopyLocal(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string BuildDirectWholePageHeaderText()
{
	return ".版本 2\r\n\r\n";
}

std::vector<std::string> ParseSupportLibraryNamesFromPageCode(const std::string& pageCode)
{
	std::vector<std::string> names;
	const std::string normalized = NormalizeLineBreakToCrLf(pageCode);
	size_t start = 0;
	while (start < normalized.size()) {
		size_t end = normalized.find("\r\n", start);
		if (end == std::string::npos) {
			end = normalized.size();
		}

		const std::string trimmedLine = TrimAsciiCopyLocal(normalized.substr(start, end - start));
		static constexpr char kPrefix[] = ".支持库";
		if (trimmedLine.rfind(kPrefix, 0) == 0) {
			const std::string name = TrimAsciiCopyLocal(trimmedLine.substr(sizeof(kPrefix) - 1));
			if (!name.empty()) {
				names.push_back(name);
			}
		}

		if (end == normalized.size()) {
			break;
		}
		start = end + 2;
	}
	return names;
}

bool CallClipboardBoolSafe(FnThiscallBool fn, void* thisPtr)
{
	if (fn == nullptr || thisPtr == nullptr) {
		return false;
	}
	__try {
		return fn(thisPtr) != FALSE;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CallEditorFormatRangeTextSafe(
	FnEditorFormatRangeText fn,
	void* editorObject,
	int startIndex,
	int endIndex,
	void* outBuffer,
	int formatMode)
{
	if (fn == nullptr || editorObject == nullptr || outBuffer == nullptr) {
		return false;
	}

	__try {
		fn(editorObject, startIndex, endIndex, outBuffer, formatMode);
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CallClipboardTextToObjectSafe(FnClipboardTextToObjectDirect fn, void* thisPtr, void* textBuffer)
{
	if (fn == nullptr || thisPtr == nullptr || textBuffer == nullptr) {
		return false;
	}
	__try {
		return fn(thisPtr, textBuffer) != 0;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CallClipboardTextToObjectWrapperSafe(FnClipboardTextToObjectWrapper fn, void* thisPtr, void* textBuffer)
{
	if (fn == nullptr || thisPtr == nullptr || textBuffer == nullptr) {
		return false;
	}
	__try {
		return fn(thisPtr, textBuffer) != 0;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

HANDLE CallClipboardSerializeSafe(FnThiscallHandle fn, void* thisPtr)
{
	if (fn == nullptr || thisPtr == nullptr) {
		return nullptr;
	}
	__try {
		return fn(thisPtr);
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return nullptr;
	}
}

int CallThiscallIntSafe(FnThiscallInt fn, void* thisPtr)
{
	if (fn == nullptr || thisPtr == nullptr) {
		return 0;
	}
	__try {
		return fn(thisPtr);
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

int CallStdcallIntPtrSafe(FnStdcallIntPtr fn, void* ptr)
{
	if (fn == nullptr || ptr == nullptr) {
		return 0;
	}
	__try {
		return fn(ptr);
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

bool CallCollectionGetValueByIndexSafe(
	FnThiscallCollectionGetValueByIndex fn,
	void* collection,
	int index,
	void** outValue)
{
	if (outValue != nullptr) {
		*outValue = nullptr;
	}
	if (fn == nullptr || collection == nullptr || index < 0) {
		return false;
	}
	__try {
		return fn(collection, index, outValue) != 0;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CopyMemoryToStringSafe(const void* source, size_t bytes, std::string* outText)
{
	if (source == nullptr || bytes == 0 || outText == nullptr) {
		return false;
	}

	char* buffer = static_cast<char*>(std::malloc(bytes));
	if (buffer == nullptr) {
		return false;
	}

	bool copied = false;
	__try {
		std::memcpy(buffer, source, bytes);
		copied = true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		copied = false;
	}

	if (!copied) {
		std::free(buffer);
		return false;
	}

	outText->assign(buffer, bytes);
	std::free(buffer);
	return true;
}

bool CallVirtualVoidAtOffsetSafe(void* thisPtr, size_t byteOffset)
{
	if (thisPtr == nullptr) {
		return false;
	}
	__try {
		auto* objectBytes = reinterpret_cast<std::byte*>(thisPtr) + byteOffset;
		auto** vtable = *reinterpret_cast<void***>(objectBytes);
		if (vtable == nullptr || vtable[2] == nullptr) {
			return false;
		}
		reinterpret_cast<void(__thiscall*)(void*)>(vtable[2])(objectBytes);
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

int GetStructuredCollectionCount(const void* collection)
{
	if (collection == nullptr) {
		return 0;
	}
	__try {
		return *reinterpret_cast<const int*>(reinterpret_cast<const std::byte*>(collection) + 0x18) >> 3;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

bool CopyCStringFieldSafe(void* destinationField, const void* sourceField)
{
	if (destinationField == nullptr || sourceField == nullptr) {
		return false;
	}
	__try {
		*reinterpret_cast<CString*>(destinationField) = *reinterpret_cast<const CString*>(sourceField);
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CopyInternalGenericArrayField(
	std::uintptr_t moduleBase,
	void* destinationField,
	const void* sourceField)
{
	if (moduleBase == 0 || destinationField == nullptr || sourceField == nullptr) {
		return false;
	}

	const auto destroyFn = ResolveInternalAddress<FnThiscallVoid>(
		moduleBase,
		kKnownGenericArrayDestroyRva);
	const auto assignFn = ResolveInternalAddress<FnThiscallGenericArrayAssign>(
		moduleBase,
		kKnownGenericArrayAssignRva);
	const auto finalizeFn = ResolveInternalAddress<FnThiscallVoidInt>(
		moduleBase,
		kKnownGenericArrayFinalizeRva);
	if (destroyFn == nullptr || assignFn == nullptr || finalizeFn == nullptr) {
		return false;
	}

	const auto* sourceBytes = reinterpret_cast<const std::byte*>(sourceField);
	const int byteCount = *reinterpret_cast<const int*>(sourceBytes + 0x10);
	const void* sourceData = *reinterpret_cast<void* const*>(sourceBytes + 0x08);

	__try {
		destroyFn(destinationField);
		if (byteCount > 0 && sourceData != nullptr) {
			assignFn(destinationField, sourceData, static_cast<size_t>(byteCount));
			finalizeFn(destinationField, 0);
		}
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CopyType10TextFields(
	std::uintptr_t moduleBase,
	void* destinationItem,
	const void* sourceItem)
{
	if (destinationItem == nullptr || sourceItem == nullptr) {
		return false;
	}

	auto* destinationBytes = reinterpret_cast<std::byte*>(destinationItem);
	const auto* sourceBytes = reinterpret_cast<const std::byte*>(sourceItem);
	return CopyInternalGenericArrayField(moduleBase, destinationBytes + 0xBC, sourceBytes + 0xBC);
}

bool CopyType13TextFields(void* destinationItem, const void* sourceItem)
{
	if (destinationItem == nullptr || sourceItem == nullptr) {
		return false;
	}

	auto* destinationBytes = reinterpret_cast<std::byte*>(destinationItem);
	const auto* sourceBytes = reinterpret_cast<const std::byte*>(sourceItem);
	bool ok = true;
	ok = CopyCStringFieldSafe(destinationBytes + 0x00, sourceBytes + 0x00) && ok;
	ok = CopyCStringFieldSafe(destinationBytes + 0x04, sourceBytes + 0x04) && ok;
	__try {
		*reinterpret_cast<DWORD*>(destinationBytes + 0x08) =
			*reinterpret_cast<const DWORD*>(sourceBytes + 0x08);
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		ok = false;
	}
	return ok;
}

bool CopyType9TextFields(void* destinationItem, const void* sourceItem)
{
	if (destinationItem == nullptr || sourceItem == nullptr) {
		return false;
	}

	auto* destinationBytes = reinterpret_cast<std::byte*>(destinationItem);
	const auto* sourceBytes = reinterpret_cast<const std::byte*>(sourceItem);
	bool ok = true;
	ok = CopyCStringFieldSafe(destinationBytes + 0x00, sourceBytes + 0x00) && ok;
	ok = CopyCStringFieldSafe(destinationBytes + 0x04, sourceBytes + 0x04) && ok;
	ok = CopyCStringFieldSafe(destinationBytes + 0x30, sourceBytes + 0x30) && ok;
	ok = CopyCStringFieldSafe(destinationBytes + 0x34, sourceBytes + 0x34) && ok;
	__try {
		*reinterpret_cast<DWORD*>(destinationBytes + 0x08) =
			*reinterpret_cast<const DWORD*>(sourceBytes + 0x08);
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		ok = false;
	}
	return ok;
}

bool FinalizeParsedClipboardObjectRaw(
	void* rawObject,
	FnThiscallMergeParsedRange mergeRangeFn,
	FnThiscallVoid destroyParsedItemFn,
	FnThiscallPtrArrayRemoveAt removePtrAtFn,
	size_t* outMergedItems,
	size_t* outRemovedItems)
{
	if (outMergedItems != nullptr) {
		*outMergedItems = 0;
	}
	if (outRemovedItems != nullptr) {
		*outRemovedItems = 0;
	}
	if (rawObject == nullptr ||
		mergeRangeFn == nullptr ||
		destroyParsedItemFn == nullptr ||
		removePtrAtFn == nullptr) {
		return false;
	}

	__try {
		auto* base = reinterpret_cast<std::byte*>(rawObject);
		auto* parsedItemArray = base + 0x1C8;
		auto* mergedRangeArray = base + 0x218;
		auto** items = *reinterpret_cast<void***>(parsedItemArray + 4);
		int count = *reinterpret_cast<int*>(parsedItemArray + 8);
		for (int index = count - 1; index >= 0; --index) {
			if (items == nullptr) {
				break;
			}

			auto* item = items[index];
			if (item == nullptr) {
				continue;
			}

			auto* itemBytes = reinterpret_cast<std::byte*>(item);
			mergeRangeFn(
				mergedRangeArray,
				itemBytes + 0x38,
				0,
				*reinterpret_cast<int*>(itemBytes + 0x3C));
			mergeRangeFn(
				mergedRangeArray,
				itemBytes + 0x18,
				0,
				*reinterpret_cast<int*>(itemBytes + 0x1C));
			if (outMergedItems != nullptr) {
				++(*outMergedItems);
			}

			if (*reinterpret_cast<int*>(itemBytes + 0xCC) != 0) {
				(void)CallVirtualVoidAtOffsetSafe(item, 0x38);
				(void)CallVirtualVoidAtOffsetSafe(item, 0x18);
				continue;
			}

			destroyParsedItemFn(item);
			operator delete(item);
			removePtrAtFn(parsedItemArray, index, 1);
			if (outRemovedItems != nullptr) {
				++(*outRemovedItems);
			}
			items = *reinterpret_cast<void***>(parsedItemArray + 4);
		}
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}

	return true;
}

bool CallEditorInsertClipboardObjectSafe(
	FnEditorInsertClipboardObject fn,
	void* editorObject,
	void* clipboardObject,
	int pasteMode,
	int insertFlags,
	int parsedTextMode)
{
	if (fn == nullptr || editorObject == nullptr || clipboardObject == nullptr) {
		return false;
	}
	__try {
		fn(editorObject, clipboardObject, pasteMode, insertFlags, parsedTextMode);
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool CallEditorPasteHandlerSafe(FnEditorPasteHandler fn, void* editorObject, int pasteMode)
{
	if (fn == nullptr || editorObject == nullptr) {
		return false;
	}
	__try {
		fn(editorObject, pasteMode);
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

class InternalClipboardObject;

bool FinalizeParsedClipboardObject(
	std::uintptr_t moduleBase,
	InternalClipboardObject* object,
	const std::vector<std::string>& supportLibraries,
	std::string* outTrace);

bool EnsureGlobalSupportLibraries(
	std::uintptr_t moduleBase,
	const std::vector<std::string>& supportLibraries,
	std::string* outTrace = nullptr);

class InternalGenericArrayBuffer {
public:
	explicit InternalGenericArrayBuffer(std::uintptr_t moduleBase)
		: m_moduleBase(moduleBase)
	{
		Initialize();
	}

	~InternalGenericArrayBuffer()
	{
		Destroy();
	}

	bool Assign(const void* data, size_t bytes)
	{
		if (!m_initialized || data == nullptr || bytes == 0) {
			return false;
		}

		const auto assignFn = ResolveInternalAddress<FnThiscallGenericArrayAssign>(
			m_moduleBase,
			kKnownGenericArrayAssignRva);
		const auto finalizeFn = ResolveInternalAddress<FnThiscallVoidInt>(
			m_moduleBase,
			kKnownGenericArrayFinalizeRva);
		if (assignFn == nullptr || finalizeFn == nullptr) {
			return false;
		}

		__try {
			assignFn(m_storage.data(), data, bytes);
			finalizeFn(m_storage.data(), 0);
			return true;
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
			return false;
		}
	}

	void* Data()
	{
		return m_storage.data();
	}

	const void* ByteData() const
	{
		if (!m_initialized) {
			return nullptr;
		}
		__try {
			return *reinterpret_cast<void* const*>(m_storage.data() + 0x08);
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
			return nullptr;
		}
	}

	int ByteSize() const
	{
		if (!m_initialized) {
			return 0;
		}
		__try {
			return *reinterpret_cast<const int*>(m_storage.data() + 0x10);
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
			return 0;
		}
	}

	bool CopyTextTo(std::string* outText) const
	{
		if (outText == nullptr) {
			return false;
		}
		outText->clear();

		const int byteCount = ByteSize();
		const auto* data = static_cast<const char*>(ByteData());
		if (byteCount <= 0 || data == nullptr) {
			return false;
		}

		if (!CopyMemoryToStringSafe(data, static_cast<size_t>(byteCount), outText)) {
			outText->clear();
			return false;
		}

		while (!outText->empty() && outText->back() == '\0') {
			outText->pop_back();
		}
		*outText = NormalizeLineBreakToCrLf(*outText);
		return !outText->empty();
	}

private:
	void Initialize()
	{
		std::memset(m_storage.data(), 0, m_storage.size());
		const auto initFn = ResolveInternalAddress<FnThiscallVoid>(
			m_moduleBase,
			kKnownGenericArrayInitRva);
		if (initFn != nullptr) {
			__try {
				initFn(m_storage.data());
				m_initialized = true;
			}
			__except (EXCEPTION_EXECUTE_HANDLER) {
				m_initialized = false;
			}
		}
	}

	void Destroy()
	{
		if (!m_initialized) {
			return;
		}

		const auto destroyFn = ResolveInternalAddress<FnThiscallVoid>(
			m_moduleBase,
			kKnownGenericArrayDestroyRva);
		if (destroyFn != nullptr) {
			__try {
				*reinterpret_cast<std::uintptr_t*>(m_storage.data()) =
					ResolveInternalPointer(m_moduleBase, kKnownGenericArrayVtableRva);
				destroyFn(m_storage.data());
			}
			__except (EXCEPTION_EXECUTE_HANDLER) {
			}
		}
		m_initialized = false;
	}

	std::uintptr_t m_moduleBase = 0;
	bool m_initialized = false;
	std::array<std::byte, kInternalGenericArraySize> m_storage{};
};

class InternalClipboardObject {
public:
	explicit InternalClipboardObject(std::uintptr_t moduleBase)
		: m_moduleBase(moduleBase)
	{
		Initialize();
	}

	~InternalClipboardObject()
	{
		Destroy();
	}

	InternalClipboardObject(const InternalClipboardObject&) = delete;
	InternalClipboardObject& operator=(const InternalClipboardObject&) = delete;

	bool DeserializeFromClipboard() const
	{
		if (!m_initialized) {
			return false;
		}
		const auto deserializeFn = ResolveInternalAddress<FnThiscallBool>(
			m_moduleBase,
			kKnownClipboardDeserializeRva);
		if (deserializeFn == nullptr) {
			return false;
		}
		return CallClipboardBoolSafe(deserializeFn, Base());
	}

	bool ParseText(const std::string& text) const
	{
		if (!m_initialized) {
			return false;
		}
		const auto parseWrapperFn = ResolveInternalAddress<FnClipboardTextToObjectWrapper>(
			m_moduleBase,
			kKnownClipboardTextToObjectWrapperRva);
		const auto parseFn = ResolveInternalAddress<FnClipboardTextToObjectDirect>(
			m_moduleBase,
			kKnownClipboardTextToObjectDirectRva);
		if (parseWrapperFn == nullptr && parseFn == nullptr) {
			return false;
		}

		const std::string normalizedText = NormalizeLineBreakToCrLf(text);
		InternalGenericArrayBuffer buffer(m_moduleBase);
		if (!buffer.Assign(normalizedText.c_str(), normalizedText.size() + 1)) {
			return false;
		}

		if (parseWrapperFn != nullptr) {
			return CallClipboardTextToObjectWrapperSafe(parseWrapperFn, Base(), buffer.Data());
		}
		return CallClipboardTextToObjectSafe(parseFn, Base(), buffer.Data());
	}

	HANDLE SerializeToClipboardHandle() const
	{
		if (!m_initialized) {
			return nullptr;
		}
		const auto serializeFn = ResolveInternalAddress<FnThiscallHandle>(
			m_moduleBase,
			kKnownClipboardSerializeRva);
		if (serializeFn == nullptr) {
			return nullptr;
		}
		return CallClipboardSerializeSafe(serializeFn, Base());
	}

	void CopyAssemblyPageMetadataFrom(const InternalClipboardObject& source) const
	{
		if (!m_initialized || !source.m_initialized) {
			return;
		}
		__try {
			SupportLibraries()->Copy(*source.SupportLibraries());
			std::memcpy(
				MutableBytes(kClipboardObjectGuidFlagOffset),
				source.Bytes(kClipboardObjectGuidFlagOffset),
				sizeof(DWORD) + sizeof(GUID));
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
		}
	}

	void CopySerializationMetadataFrom(const InternalClipboardObject& source) const
	{
		if (!m_initialized || !source.m_initialized) {
			return;
		}
		__try {
			std::memcpy(
				MutableBytes(kClipboardObjectGuidFlagOffset),
				source.Bytes(kClipboardObjectGuidFlagOffset),
				sizeof(DWORD) + sizeof(GUID));
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
		}
	}

	size_t SupportLibraryCount() const
	{
		return static_cast<size_t>(SupportLibraries()->GetSize());
	}

	void* RawObject() const
	{
		return Base();
	}

	void SetSupportLibraries(const std::vector<std::string>& names) const
	{
		if (!m_initialized) {
			return;
		}

		__try {
			SupportLibraries()->RemoveAll();
			for (const std::string& name : names) {
				SupportLibraries()->Add(name.c_str());
			}
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
		}
	}

private:
	void Initialize()
	{
		std::memset(m_storage.data(), 0, m_storage.size());
		const auto ctorFn = ResolveInternalAddress<int(__thiscall*)(void*)>(
			m_moduleBase,
			kKnownClipboardObjectCtorRva);
		if (ctorFn == nullptr) {
			return;
		}

		__try {
			(void)ctorFn(Base());
			m_initialized = true;
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
			m_initialized = false;
		}
	}

	void Destroy()
	{
		if (!m_initialized) {
			return;
		}

		const auto destroyFn = ResolveInternalAddress<FnThiscallVoid>(
			m_moduleBase,
			kKnownClipboardObjectDestroyRva);
		if (destroyFn != nullptr) {
			__try {
				destroyFn(Base());
			}
			__except (EXCEPTION_EXECUTE_HANDLER) {
			}
		}
		m_initialized = false;
	}

	void* Base() const
	{
		return const_cast<std::byte*>(m_storage.data());
	}

	void* MutableBytes(size_t offset) const
	{
		return const_cast<std::byte*>(m_storage.data() + offset);
	}

	const void* Bytes(size_t offset) const
	{
		return m_storage.data() + offset;
	}

	CStringArray* SupportLibraries() const
	{
		return reinterpret_cast<CStringArray*>(MutableBytes(kClipboardObjectSupportLibrariesOffset));
	}

	std::uintptr_t m_moduleBase = 0;
	bool m_initialized = false;
	mutable std::array<std::byte, kInternalClipboardObjectSize> m_storage{};
};

struct ClipboardObjectValidationStats {
	int validateA00 = 0;
	int validate560 = 0;
	int validate230 = 0;
	int count1FC = 0;
	int count218 = 0;
	int count238 = 0;
	int count254 = 0;
	int count270 = 0;
	int invalid1FC = 0;
	int invalid218 = 0;
	int invalid238 = 0;
	int invalid254 = 0;
	int invalid270 = 0;
};

ClipboardObjectValidationStats GatherClipboardObjectValidationStats(
	std::uintptr_t moduleBase,
	void* rawObject)
{
	ClipboardObjectValidationStats stats{};
	if (moduleBase == 0 || rawObject == nullptr) {
		return stats;
	}

	const auto validateA00Fn = ResolveInternalAddress<FnThiscallInt>(
		moduleBase,
		kKnownClipboardValidateA00Rva);
	const auto validate560Fn = ResolveInternalAddress<FnThiscallInt>(
		moduleBase,
		kKnownClipboardValidate560Rva);
	const auto validate230Fn = ResolveInternalAddress<FnThiscallInt>(
		moduleBase,
		kKnownClipboardValidate230Rva);
	const auto firstInvalidFn = ResolveInternalAddress<FnStdcallIntPtr>(
		moduleBase,
		kKnownClipboardCollectionFirstInvalidRva);
	auto* rawObjectBytes = reinterpret_cast<std::byte*>(rawObject);

	stats.count1FC = GetStructuredCollectionCount(rawObjectBytes + 0x1FC);
	stats.count218 = GetStructuredCollectionCount(rawObjectBytes + 0x218);
	stats.count238 = GetStructuredCollectionCount(rawObjectBytes + 0x238);
	stats.count254 = GetStructuredCollectionCount(rawObjectBytes + 0x254);
	stats.count270 = GetStructuredCollectionCount(rawObjectBytes + 0x270);
	stats.invalid1FC = CallStdcallIntPtrSafe(firstInvalidFn, rawObjectBytes + 0x1FC);
	stats.invalid218 = CallStdcallIntPtrSafe(firstInvalidFn, rawObjectBytes + 0x218);
	stats.invalid238 = CallStdcallIntPtrSafe(firstInvalidFn, rawObjectBytes + 0x238);
	stats.invalid254 = CallStdcallIntPtrSafe(firstInvalidFn, rawObjectBytes + 0x254);
	stats.invalid270 = CallStdcallIntPtrSafe(firstInvalidFn, rawObjectBytes + 0x270);
	stats.validateA00 = CallThiscallIntSafe(validateA00Fn, rawObject);
	stats.validate560 = CallThiscallIntSafe(validate560Fn, rawObject);
	stats.validate230 = CallThiscallIntSafe(validate230Fn, rawObject);
	return stats;
}

std::string BuildClipboardObjectValidationTrace(const ClipboardObjectValidationStats& stats)
{
	return
		"validate_a00=" + std::to_string(stats.validateA00) +
		"|validate_560=" + std::to_string(stats.validate560) +
		"|validate_230=" + std::to_string(stats.validate230) +
		"|count_1fc=" + std::to_string(stats.count1FC) +
		"|count_218=" + std::to_string(stats.count218) +
		"|count_238=" + std::to_string(stats.count238) +
		"|count_254=" + std::to_string(stats.count254) +
		"|count_270=" + std::to_string(stats.count270) +
		"|invalid_1fc=" + std::to_string(stats.invalid1FC) +
		"|invalid_218=" + std::to_string(stats.invalid218) +
		"|invalid_238=" + std::to_string(stats.invalid238) +
		"|invalid_254=" + std::to_string(stats.invalid254) +
		"|invalid_270=" + std::to_string(stats.invalid270);
}

bool TransplantParsedTextOntoTemplateObject(
	std::uintptr_t moduleBase,
	InternalClipboardObject* templateObject,
	const InternalClipboardObject& parsedObject,
	const std::vector<std::string>& supportLibraries,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (moduleBase == 0 || templateObject == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "transplant_invalid_argument";
		}
		return false;
	}

	const auto getValueByIndexFn = ResolveInternalAddress<FnThiscallCollectionGetValueByIndex>(
		moduleBase,
		kKnownCollectionGetValueByIndexRva);
	if (getValueByIndexFn == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "transplant_resolve_failed";
		}
		return false;
	}

	auto* templateBytes = reinterpret_cast<std::byte*>(templateObject->RawObject());
	auto* parsedBytes = reinterpret_cast<std::byte*>(parsedObject.RawObject());
	const int templateCount238 = GetStructuredCollectionCount(templateBytes + 0x238);
	const int parsedCount238 = GetStructuredCollectionCount(parsedBytes + 0x238);
	const int templateCount254 = GetStructuredCollectionCount(templateBytes + 0x254);
	const int parsedCount254 = GetStructuredCollectionCount(parsedBytes + 0x254);
	const int templateCount270 = GetStructuredCollectionCount(templateBytes + 0x270);
	const int parsedCount270 = GetStructuredCollectionCount(parsedBytes + 0x270);
	if (templateCount238 != parsedCount238) {
		if (outTrace != nullptr) {
			*outTrace =
				"transplant_count_mismatch" +
				std::string("|template_238=") + std::to_string(templateCount238) +
				"|parsed_238=" + std::to_string(parsedCount238) +
				"|template_254=" + std::to_string(templateCount254) +
				"|parsed_254=" + std::to_string(parsedCount254) +
				"|template_270=" + std::to_string(templateCount270) +
				"|parsed_270=" + std::to_string(parsedCount270);
		}
		return false;
	}

	templateObject->SetSupportLibraries(supportLibraries);

	size_t copied254 = 0;
	if (templateCount254 == parsedCount254) {
		for (int index = 0; index < templateCount254; ++index) {
			void* templateItem = nullptr;
			void* parsedItem = nullptr;
			if (!CallCollectionGetValueByIndexSafe(getValueByIndexFn, templateBytes + 0x254, index, &templateItem) ||
				!CallCollectionGetValueByIndexSafe(getValueByIndexFn, parsedBytes + 0x254, index, &parsedItem) ||
				!CopyType13TextFields(templateItem, parsedItem)) {
				if (outTrace != nullptr) {
					*outTrace = "transplant_type13_failed|index=" + std::to_string(index);
				}
				return false;
			}
			++copied254;
		}
	}

	size_t copied270 = 0;
	if (templateCount270 == parsedCount270) {
		for (int index = 0; index < templateCount270; ++index) {
			void* templateItem = nullptr;
			void* parsedItem = nullptr;
			if (!CallCollectionGetValueByIndexSafe(getValueByIndexFn, templateBytes + 0x270, index, &templateItem) ||
				!CallCollectionGetValueByIndexSafe(getValueByIndexFn, parsedBytes + 0x270, index, &parsedItem) ||
				!CopyType9TextFields(templateItem, parsedItem)) {
				if (outTrace != nullptr) {
					*outTrace = "transplant_type9_failed|index=" + std::to_string(index);
				}
				return false;
			}
			++copied270;
		}
	}

	size_t copied238 = 0;
	for (int index = 0; index < templateCount238; ++index) {
		void* templateItem = nullptr;
		void* parsedItem = nullptr;
		if (!CallCollectionGetValueByIndexSafe(getValueByIndexFn, templateBytes + 0x238, index, &templateItem) ||
			!CallCollectionGetValueByIndexSafe(getValueByIndexFn, parsedBytes + 0x238, index, &parsedItem) ||
			!CopyType10TextFields(moduleBase, templateItem, parsedItem)) {
			if (outTrace != nullptr) {
				*outTrace = "transplant_type10_failed|index=" + std::to_string(index);
			}
			return false;
		}
		++copied238;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"transplant_ok|copied_254=" + std::to_string(copied254) +
			"|copied_270=" + std::to_string(copied270) +
			"|copied_238=" + std::to_string(copied238) +
			"|template_254=" + std::to_string(templateCount254) +
			"|parsed_254=" + std::to_string(parsedCount254) +
			"|template_270=" + std::to_string(templateCount270) +
			"|parsed_270=" + std::to_string(parsedCount270) +
			"|support_libs=" + std::to_string(supportLibraries.size());
	}
	return true;
}

bool EnsureGlobalSupportLibraries(
	std::uintptr_t moduleBase,
	const std::vector<std::string>& supportLibraries,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (moduleBase == 0) {
		if (outTrace != nullptr) {
			*outTrace = "support_library_invalid_module";
		}
		return false;
	}

	const auto addSupportLibraryFn = ResolveInternalAddress<FnCdeclCStringArrayAddUnique>(
		moduleBase,
		kKnownSupportLibraryArrayAddUniqueRva);
	auto* globalSupportLibraryArray = reinterpret_cast<CStringArray*>(
		ResolveInternalPointer(moduleBase, kKnownGlobalSupportLibraryArrayRva));
	if (addSupportLibraryFn == nullptr || globalSupportLibraryArray == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "support_library_resolve_failed";
		}
		return false;
	}

	size_t addedCount = 0;
	for (const std::string& name : supportLibraries) {
		if (name.empty()) {
			continue;
		}
		addSupportLibraryFn(
			globalSupportLibraryArray,
			reinterpret_cast<unsigned char*>(const_cast<char*>(name.c_str())),
			0);
		++addedCount;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"support_library_global_ok|requested=" +
			std::to_string(supportLibraries.size()) +
			"|added=" +
			std::to_string(addedCount);
	}
	return true;
}

bool FinalizeParsedClipboardObject(
	std::uintptr_t moduleBase,
	InternalClipboardObject* object,
	const std::vector<std::string>& supportLibraries,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (moduleBase == 0 || object == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "finalize_invalid_argument";
		}
		return false;
	}

	const auto mergeRangeFn = ResolveInternalAddress<FnThiscallMergeParsedRange>(
		moduleBase,
		kKnownClipboardMergeParsedRangeRva);
	const auto destroyParsedItemFn = ResolveInternalAddress<FnThiscallVoid>(
		moduleBase,
		kKnownParsedItemDestroyRva);
	const auto removePtrAtFn = ResolveInternalAddress<FnThiscallPtrArrayRemoveAt>(
		moduleBase,
		kKnownPtrArrayRemoveAtRva);
	if (mergeRangeFn == nullptr ||
		destroyParsedItemFn == nullptr ||
		removePtrAtFn == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "finalize_resolve_failed";
		}
		return false;
	}

	size_t mergedItems = 0;
	size_t removedItems = 0;
	std::string supportLibraryTrace;
	if (!EnsureGlobalSupportLibraries(moduleBase, supportLibraries, &supportLibraryTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "finalize_support_libraries_failed|" + supportLibraryTrace;
		}
		return false;
	}

	if (!FinalizeParsedClipboardObjectRaw(
			object->RawObject(),
			mergeRangeFn,
			destroyParsedItemFn,
			removePtrAtFn,
			&mergedItems,
			&removedItems)) {
		if (outTrace != nullptr) {
			*outTrace = "finalize_exception";
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"finalize_ok|support_libs=" + std::to_string(supportLibraries.size()) +
			"|merged_items=" + std::to_string(mergedItems) +
			"|removed_items=" + std::to_string(removedItems) +
			"|" +
			supportLibraryTrace;
	}
	return true;
}

bool ReadClipboardAnsiTextHandle(HANDLE handle, std::string* outText);

void PumpPendingMessages()
{
	MSG msg = {};
	while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
		if (!AfxPreTranslateMessage(&msg)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
}

bool WaitForFakeClipboardText(
	ScopedFakeClipboard& fakeClipboard,
	std::string* outText,
	std::string* outTrace)
{
	if (outText != nullptr) {
		outText->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	for (int attempt = 0; attempt < 80; ++attempt) {
		PumpPendingMessages();
		HANDLE textHandle = fakeClipboard.GetFormatHandle(CF_TEXT);
		if (ReadClipboardAnsiTextHandle(textHandle, outText) &&
			outText != nullptr &&
			!outText->empty()) {
			if (outTrace != nullptr) {
				*outTrace = "fake_clipboard_text_ok|" + fakeClipboard.BuildStatsText();
			}
			return true;
		}
		Sleep(25);
	}

	if (outTrace != nullptr) {
		*outTrace = "fake_clipboard_text_missing|" + fakeClipboard.BuildStatsText();
	}
	return false;
}

std::string NormalizeLineBreakToCrLf(const std::string& text)
{
	std::string normalized;
	normalized.reserve(text.size() + 16);
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			normalized += "\r\n";
			continue;
		}
		if (ch == '\n') {
			normalized += "\r\n";
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

std::string NormalizeLineBreakToLf(const std::string& text)
{
	std::string normalized;
	normalized.reserve(text.size());
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			normalized.push_back('\n');
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

std::string NormalizePageCodeForLooseCompare(const std::string& text)
{
	const std::string normalized = NormalizeLineBreakToLf(text);
	std::string result;
	result.reserve(normalized.size());
	bool lastKeptLineBlank = false;

	size_t start = 0;
	while (start <= normalized.size()) {
		size_t end = normalized.find('\n', start);
		if (end == std::string::npos) {
			end = normalized.size();
		}

		const std::string line = normalized.substr(start, end - start);
		const std::string trimmedLine = TrimAsciiCopyLocal(line);
		if (trimmedLine.rfind(".支持库", 0) != 0) {
			const bool isBlankLine = trimmedLine.empty();
			if (isBlankLine && lastKeptLineBlank) {
				if (end == normalized.size()) {
					break;
				}
				start = end + 1;
				continue;
			}
			result.append(line);
			if (end != normalized.size()) {
				result.push_back('\n');
			}
			lastKeptLineBlank = isBlankLine;
		}

		if (end == normalized.size()) {
			break;
		}
		start = end + 1;
	}

	while (!result.empty() && result.back() == '\n') {
		result.pop_back();
	}

	return result;
}

std::string NormalizePageCodeLineForStructuralCompare(const std::string& line)
{
	const std::string trimmedLine = TrimAsciiCopyLocal(line);
	if (trimmedLine.empty()) {
		return std::string();
	}

	std::string normalized;
	normalized.reserve(trimmedLine.size());
	bool inString = false;

	for (size_t i = 0; i < trimmedLine.size(); ++i) {
		const char ch = trimmedLine[i];
		if (!inString && std::isspace(static_cast<unsigned char>(ch)) != 0) {
			continue;
		}

		normalized.push_back(ch);
		if (ch != '"') {
			continue;
		}

		if (inString && i + 1 < trimmedLine.size() && trimmedLine[i + 1] == '"') {
			normalized.push_back(trimmedLine[i + 1]);
			++i;
			continue;
		}

		inString = !inString;
	}

	return normalized;
}

std::vector<std::string> BuildPageCodeStructuralFingerprint(const std::string& text)
{
	const std::string normalized = NormalizeLineBreakToLf(text);
	std::vector<std::string> lines;

	size_t start = 0;
	while (start <= normalized.size()) {
		size_t end = normalized.find('\n', start);
		if (end == std::string::npos) {
			end = normalized.size();
		}

		const std::string line = normalized.substr(start, end - start);
		const std::string trimmedLine = TrimAsciiCopyLocal(line);
		if (!trimmedLine.empty() && trimmedLine.rfind(".支持库", 0) != 0) {
			const std::string normalizedLine = NormalizePageCodeLineForStructuralCompare(trimmedLine);
			if (!normalizedLine.empty()) {
				lines.push_back(normalizedLine);
			}
		}

		if (end == normalized.size()) {
			break;
		}
		start = end + 1;
	}

	return lines;
}

std::string BuildStructuralMismatchSummary(
	const std::vector<std::string>& expectedLines,
	const std::vector<std::string>& actualLines)
{
	const auto buildAsciiSafeSnippet = [](const std::string& text) {
		const size_t len = (std::min<size_t>)(80, text.size());
		const std::string snippet = text.substr(0, len);
		std::string safe;
		safe.reserve(snippet.size() * 4);
		static constexpr char kHex[] = "0123456789ABCDEF";
		for (unsigned char ch : snippet) {
			if (ch >= 0x20 && ch <= 0x7E && ch != '\\') {
				safe.push_back(static_cast<char>(ch));
				continue;
			}
			safe.push_back('\\');
			safe.push_back('x');
			safe.push_back(kHex[(ch >> 4) & 0x0F]);
			safe.push_back(kHex[ch & 0x0F]);
		}
		return safe;
	};

	const size_t common = (std::min)(expectedLines.size(), actualLines.size());
	size_t mismatchLine = common;
	for (size_t i = 0; i < common; ++i) {
		if (expectedLines[i] != actualLines[i]) {
			mismatchLine = i;
			break;
		}
	}

	const std::string expectedLine = mismatchLine < expectedLines.size()
		? expectedLines[mismatchLine]
		: std::string();
	const std::string actualLine = mismatchLine < actualLines.size()
		? actualLines[mismatchLine]
		: std::string();

	return
		"structural_line=" + std::to_string(mismatchLine + 1) +
		"|expected_lines=" + std::to_string(expectedLines.size()) +
		"|actual_lines=" + std::to_string(actualLines.size()) +
		"|expected_line=" + buildAsciiSafeSnippet(expectedLine) +
		"|actual_line=" + buildAsciiSafeSnippet(actualLine);
}

bool VerifyRealPageCodeMatches(
	const std::string& expectedCode,
	const std::string& actualCode,
	std::string* outMode,
	std::string* outSummary)
{
	if (outMode != nullptr) {
		outMode->clear();
	}
	if (outSummary != nullptr) {
		outSummary->clear();
	}

	const std::string normalizedExpectedCode = NormalizePageCodeForLooseCompare(expectedCode);
	const std::string normalizedActualCode = NormalizePageCodeForLooseCompare(actualCode);
	if (normalizedExpectedCode == normalizedActualCode) {
		if (outMode != nullptr) {
			*outMode = "loose_text";
		}
		return true;
	}

	const std::vector<std::string> expectedFingerprint = BuildPageCodeStructuralFingerprint(expectedCode);
	const std::vector<std::string> actualFingerprint = BuildPageCodeStructuralFingerprint(actualCode);
	if (expectedFingerprint == actualFingerprint) {
		if (outMode != nullptr) {
			*outMode = "structural";
		}
		if (outSummary != nullptr) {
			*outSummary =
				"loose_text_mismatch|" +
				BuildTextMismatchSummary(normalizedExpectedCode, normalizedActualCode) +
				"|structural_lines=" +
				std::to_string(expectedFingerprint.size());
		}
		return true;
	}

	if (outSummary != nullptr) {
		*outSummary =
			"loose_text_mismatch|" +
			BuildTextMismatchSummary(normalizedExpectedCode, normalizedActualCode) +
			"|structural_mismatch|" +
			BuildStructuralMismatchSummary(expectedFingerprint, actualFingerprint);
	}
	return false;
}

HANDLE CreateClipboardAnsiTextHandle(const std::string& text)
{
	const std::string normalized = NormalizeLineBreakToCrLf(EnsureTextUsesGbkCodePage(text));
	const SIZE_T bytes = static_cast<SIZE_T>(normalized.size() + 1);
	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes);
	if (hMem == nullptr) {
		return nullptr;
	}

	void* mem = GlobalLock(hMem);
	if (mem == nullptr) {
		GlobalFree(hMem);
		return nullptr;
	}

	if (!normalized.empty()) {
		std::memcpy(mem, normalized.data(), normalized.size());
	}
	static_cast<char*>(mem)[normalized.size()] = '\0';
	GlobalUnlock(hMem);
	return hMem;
}

bool ReadClipboardAnsiTextHandle(HANDLE handle, std::string* outText)
{
	if (outText != nullptr) {
		outText->clear();
	}
	if (handle == nullptr || outText == nullptr) {
		return false;
	}

	const SIZE_T size = GlobalSize(handle);
	if (size == 0) {
		return false;
	}

	const char* mem = static_cast<const char*>(GlobalLock(handle));
	if (mem == nullptr) {
		return false;
	}

	size_t textLen = 0;
	while (textLen < size && mem[textLen] != '\0') {
		++textLen;
	}
	outText->assign(mem, textLen);
	GlobalUnlock(handle);
	return true;
}

bool ReadRealClipboardText(std::string* outText)
{
	if (outText != nullptr) {
		outText->clear();
	}
	if (outText == nullptr) {
		return false;
	}

	for (int attempt = 0; attempt < 20; ++attempt) {
		if (!OpenClipboard(nullptr)) {
			Sleep(25);
			continue;
		}

		bool ok = false;
		HANDLE unicodeHandle = GetClipboardData(CF_UNICODETEXT);
		if (unicodeHandle != nullptr) {
			const wchar_t* wideText = static_cast<const wchar_t*>(GlobalLock(unicodeHandle));
			if (wideText != nullptr) {
				const int bytes = WideCharToMultiByte(CP_ACP, 0, wideText, -1, nullptr, 0, nullptr, nullptr);
				if (bytes > 0) {
					std::string localText(static_cast<size_t>(bytes), '\0');
					if (WideCharToMultiByte(CP_ACP, 0, wideText, -1, localText.data(), bytes, nullptr, nullptr) > 0) {
						if (!localText.empty() && localText.back() == '\0') {
							localText.pop_back();
						}
						*outText = std::move(localText);
						ok = !outText->empty();
					}
				}
				GlobalUnlock(unicodeHandle);
			}
		}

		if (!ok) {
			HANDLE ansiHandle = GetClipboardData(CF_TEXT);
			ok = ReadClipboardAnsiTextHandle(ansiHandle, outText) && !outText->empty();
		}

		CloseClipboard();
		if (ok) {
			return true;
		}
		Sleep(25);
	}
	return false;
}

HANDLE CreateClipboardUnicodeTextHandle(const std::string& text)
{
	const int wideChars = MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, nullptr, 0);
	if (wideChars <= 0) {
		return nullptr;
	}

	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, static_cast<SIZE_T>(wideChars) * sizeof(wchar_t));
	if (hMem == nullptr) {
		return nullptr;
	}

	wchar_t* wideBuffer = static_cast<wchar_t*>(GlobalLock(hMem));
	if (wideBuffer == nullptr) {
		GlobalFree(hMem);
		return nullptr;
	}
	if (MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, wideBuffer, wideChars) <= 0) {
		GlobalUnlock(hMem);
		GlobalFree(hMem);
		return nullptr;
	}
	GlobalUnlock(hMem);
	return hMem;
}

bool SetRealClipboardText(const std::string& text)
{
	for (int attempt = 0; attempt < 20; ++attempt) {
		if (!OpenClipboard(nullptr)) {
			Sleep(25);
			continue;
		}

		EmptyClipboard();
		HANDLE ansiHandle = CreateClipboardAnsiTextHandle(text);
		if (ansiHandle != nullptr) {
			SetClipboardData(CF_TEXT, ansiHandle);
		}
		HANDLE unicodeHandle = CreateClipboardUnicodeTextHandle(text);
		if (unicodeHandle != nullptr) {
			SetClipboardData(CF_UNICODETEXT, unicodeHandle);
		}
		CloseClipboard();
		return ansiHandle != nullptr || unicodeHandle != nullptr;
	}
	return false;
}

class ScopedRealClipboardTextRestore {
public:
	ScopedRealClipboardTextRestore()
	{
		m_hasText = ReadRealClipboardText(&m_text);
	}

	~ScopedRealClipboardTextRestore()
	{
		if (m_hasText) {
			(void)SetRealClipboardText(m_text);
		}
	}

private:
	bool m_hasText = false;
	std::string m_text;
};

bool TryResolveInnerEditorObject(std::uintptr_t rawObject, EditorDispatchTargetInfo* outInfo);
void DispatchSyntheticKey(HWND editorHwnd, WORD vk, bool ctrlDown);
void SettleEditorAfterCommand(DWORD waitMs);

std::string ToLowerAsciiCopyLocalBridge(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string WindowClassToString(HWND hWnd)
{
	char className[128] = {};
	if (hWnd != nullptr && GetClassNameA(hWnd, className, static_cast<int>(sizeof(className))) > 0) {
		return className;
	}
	return std::string();
}

std::string WindowTextToString(HWND hWnd)
{
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return std::string();
	}
	const int length = GetWindowTextLengthA(hWnd);
	if (length <= 0) {
		return std::string();
	}
	std::string text(static_cast<size_t>(length), '\0');
	if (GetWindowTextA(hWnd, text.data(), length + 1) <= 0) {
		return std::string();
	}
	return text;
}

std::vector<HWND> CollectChildWindowsByClassName(HWND rootHwnd, const char* className)
{
	std::vector<HWND> windows;
	if (rootHwnd == nullptr || !IsWindow(rootHwnd) || className == nullptr || className[0] == '\0') {
		return windows;
	}

	struct EnumContext {
		const char* className = nullptr;
		std::vector<HWND>* windows = nullptr;
	};

	EnumContext context{};
	context.className = className;
	context.windows = &windows;
	EnumChildWindows(
		rootHwnd,
		[](HWND hWnd, LPARAM lParam) -> BOOL {
			auto* ctx = reinterpret_cast<EnumContext*>(lParam);
			if (ctx == nullptr || ctx->windows == nullptr || ctx->className == nullptr) {
				return TRUE;
			}

			char currentClassName[128] = {};
			if (GetClassNameA(hWnd, currentClassName, static_cast<int>(sizeof(currentClassName))) > 0 &&
				std::strcmp(currentClassName, ctx->className) == 0) {
				ctx->windows->push_back(hWnd);
			}
			return TRUE;
		},
		reinterpret_cast<LPARAM>(&context));
	return windows;
}

HWND ResolveMainIdeWindow()
{
	if (g_hwnd != nullptr && IsWindow(g_hwnd)) {
		return g_hwnd;
	}

	CWnd* const mainWnd = AfxGetMainWnd();
	if (mainWnd != nullptr) {
		const HWND hWnd = mainWnd->GetSafeHwnd();
		if (hWnd != nullptr && IsWindow(hWnd)) {
			g_hwnd = hWnd;
			return hWnd;
		}
	}

	const HWND enumeratedHwnd = GetMainWindowByProcessId();
	if (enumeratedHwnd != nullptr && IsWindow(enumeratedHwnd)) {
		g_hwnd = enumeratedHwnd;
		return enumeratedHwnd;
	}
	return nullptr;
}

struct WindowDepthInfo {
	HWND hWnd = nullptr;
	int depth = 0;
};

struct EditorWindowCandidateInfo {
	HWND hWnd = nullptr;
	int depth = 0;
	int score = 0;
	std::string className;
	std::string title;
};

bool IsSameOrChildWindow(HWND rootHwnd, HWND targetHwnd)
{
	return rootHwnd != nullptr &&
		targetHwnd != nullptr &&
		(rootHwnd == targetHwnd || IsChild(rootHwnd, targetHwnd) != FALSE);
}

void CollectDescendantWindowsWithDepth(HWND rootHwnd, int depth, std::vector<WindowDepthInfo>& outWindows)
{
	if (rootHwnd == nullptr || !IsWindow(rootHwnd)) {
		return;
	}

	for (HWND childHwnd = GetWindow(rootHwnd, GW_CHILD);
		 childHwnd != nullptr;
		 childHwnd = GetWindow(childHwnd, GW_HWNDNEXT)) {
		if (!IsWindow(childHwnd)) {
			continue;
		}
		outWindows.push_back({ childHwnd, depth });
		CollectDescendantWindowsWithDepth(childHwnd, depth + 1, outWindows);
	}
}

bool IsIgnoredEditorCandidateClass(const std::string& className)
{
	const std::string lowerClassName = ToLowerAsciiCopyLocalBridge(className);
	if (lowerClassName.empty()) {
		return false;
	}

	if (lowerClassName == "scrollbar" ||
		lowerClassName == "toolbarwindow32" ||
		lowerClassName == "ccustomtabctrl" ||
		lowerClassName == "mdiclient" ||
		lowerClassName == "msctls_statusbar32" ||
		lowerClassName == "systabcontrol32" ||
		lowerClassName == "button" ||
		lowerClassName == "combobox" ||
		lowerClassName == "msctls_progress32" ||
		lowerClassName == "tooltips_class32" ||
		lowerClassName == "systreeview32" ||
		lowerClassName == "#32770") {
		return true;
	}

	return lowerClassName.rfind("afxcontrolbar", 0) == 0;
}

bool IsPreferredEditorCandidateClass(const std::string& className)
{
	return className.rfind("Afx:", 0) == 0 || className.rfind("AfxMDIFrame", 0) == 0;
}

bool IsReadableCommittedProtection(DWORD protect)
{
	if ((protect & PAGE_GUARD) != 0) {
		return false;
	}

	switch (protect & 0xFF) {
	case PAGE_READONLY:
	case PAGE_READWRITE:
	case PAGE_WRITECOPY:
	case PAGE_EXECUTE_READ:
	case PAGE_EXECUTE_READWRITE:
	case PAGE_EXECUTE_WRITECOPY:
		return true;
	default:
		return false;
	}
}

bool IsReadableAddressRange(std::uintptr_t address, size_t bytes)
{
	if (address == 0 || bytes == 0) {
		return false;
	}

	MEMORY_BASIC_INFORMATION mbi = {};
	if (VirtualQuery(reinterpret_cast<LPCVOID>(address), &mbi, sizeof(mbi)) != sizeof(mbi)) {
		return false;
	}
	if (mbi.State != MEM_COMMIT || !IsReadableCommittedProtection(mbi.Protect)) {
		return false;
	}

	const std::uintptr_t regionStart = reinterpret_cast<std::uintptr_t>(mbi.BaseAddress);
	const std::uintptr_t regionEnd = regionStart + mbi.RegionSize;
	if (address < regionStart || address > regionEnd) {
		return false;
	}
	return bytes <= (regionEnd - address);
}

bool TryReadUint32ArraySafe(std::uintptr_t address, std::uint32_t* outValues, size_t count)
{
	if (outValues == nullptr || count == 0) {
		return false;
	}
	const size_t totalBytes = count * sizeof(std::uint32_t);
	if (!IsReadableAddressRange(address, totalBytes)) {
		return false;
	}

	__try {
		std::memcpy(outValues, reinterpret_cast<const void*>(address), totalBytes);
		return true;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool TryReadMemoryBytesSafe(std::uintptr_t address, void* buffer, size_t bytes)
{
	if (buffer == nullptr || bytes == 0) {
		return false;
	}
	if (!IsReadableAddressRange(address, bytes)) {
		return false;
	}

	SIZE_T bytesRead = 0;
	return ReadProcessMemory(
			   GetCurrentProcess(),
			   reinterpret_cast<LPCVOID>(address),
			   buffer,
			   bytes,
			   &bytesRead) != FALSE &&
		bytesRead == bytes;
}

std::uintptr_t GetModuleImageEnd(std::uintptr_t moduleBase)
{
	if (moduleBase == 0) {
		return 0;
	}

	__try {
		const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(moduleBase);
		if (dos->e_magic != IMAGE_DOS_SIGNATURE) {
			return 0;
		}
		const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS*>(moduleBase + dos->e_lfanew);
		if (nt->Signature != IMAGE_NT_SIGNATURE) {
			return 0;
		}
		return moduleBase + nt->OptionalHeader.SizeOfImage;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return 0;
	}
}

bool IsRecognizedEditorPageType(unsigned int pageType)
{
	switch (pageType) {
	case 1:
	case 2:
	case 3:
	case 4:
	case 6:
	case 7:
	case 8:
		return true;
	default:
		return false;
	}
}

std::uintptr_t ResolveInnerEditorObjectFromFields(const std::uint32_t* fields, unsigned int* outPageType = nullptr)
{
	if (outPageType != nullptr) {
		*outPageType = 0;
	}
	if (fields == nullptr) {
		return 0;
	}

	const unsigned int pageType = fields[15];
	if (outPageType != nullptr) {
		*outPageType = pageType;
	}

	switch (pageType) {
	case 1:
		return static_cast<std::uintptr_t>(fields[21]);
	case 2:
		return static_cast<std::uintptr_t>(fields[16]);
	case 3:
		return static_cast<std::uintptr_t>(fields[17]);
	case 4:
		return static_cast<std::uintptr_t>(fields[18]);
	case 6:
	case 7:
	case 8:
		return static_cast<std::uintptr_t>(fields[20]);
	default:
		return 0;
	}
}

std::vector<EditorWindowCandidateInfo> CollectEditorWindowCandidatesFromMdiChild(HWND mdiChildHwnd)
{
	std::vector<EditorWindowCandidateInfo> candidates;
	if (mdiChildHwnd == nullptr || !IsWindow(mdiChildHwnd)) {
		return candidates;
	}

	const auto appendCandidate = [&candidates](HWND hWnd, int depth) {
		if (hWnd == nullptr || !IsWindow(hWnd) || !IsWindowVisible(hWnd)) {
			return;
		}

		const std::string className = WindowClassToString(hWnd);
		if (IsIgnoredEditorCandidateClass(className)) {
			return;
		}

		const std::string title = WindowTextToString(hWnd);
		int score = depth * 100;
		if (IsPreferredEditorCandidateClass(className)) {
			score += 1000;
		}
		if (title.empty()) {
			score += 100;
		}
		if (GetWindow(hWnd, GW_CHILD) == nullptr) {
			score += 20;
		}
		if (GetFocus() == hWnd) {
			score += 2000;
		}

		for (auto& existing : candidates) {
			if (existing.hWnd == hWnd) {
				if (score > existing.score) {
					existing.depth = depth;
					existing.score = score;
					existing.className = className;
					existing.title = title;
				}
				return;
			}
		}

		candidates.push_back({ hWnd, depth, score, className, title });
	};

	GUITHREADINFO guiInfo = {};
	guiInfo.cbSize = sizeof(guiInfo);
	if (GetGUIThreadInfo(GetCurrentThreadId(), &guiInfo)) {
		if (guiInfo.hwndFocus != nullptr && IsSameOrChildWindow(mdiChildHwnd, guiInfo.hwndFocus)) {
			appendCandidate(guiInfo.hwndFocus, 0);
		}
		if (guiInfo.hwndCaret != nullptr && IsSameOrChildWindow(mdiChildHwnd, guiInfo.hwndCaret)) {
			appendCandidate(guiInfo.hwndCaret, 0);
		}
	}

	std::vector<WindowDepthInfo> descendants;
	CollectDescendantWindowsWithDepth(mdiChildHwnd, 1, descendants);
	for (const auto& info : descendants) {
		appendCandidate(info.hWnd, info.depth);
	}

	std::sort(
		candidates.begin(),
		candidates.end(),
		[](const EditorWindowCandidateInfo& lhs, const EditorWindowCandidateInfo& rhs) {
			if (lhs.score != rhs.score) {
				return lhs.score > rhs.score;
			}
			if (lhs.depth != rhs.depth) {
				return lhs.depth > rhs.depth;
			}
			return reinterpret_cast<std::uintptr_t>(lhs.hWnd) < reinterpret_cast<std::uintptr_t>(rhs.hWnd);
		});
	return candidates;
}

bool TryResolveEditorObjectFromWindowHandleHeuristic(
	HWND targetHwnd,
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	std::string* outTrace)
{
	if (outEditorObject != nullptr) {
		*outEditorObject = 0;
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (targetHwnd == nullptr || !IsWindow(targetHwnd) || moduleBase == 0) {
		if (outTrace != nullptr) {
			*outTrace = "hwnd_heuristic_invalid_argument";
		}
		return false;
	}

	const std::uintptr_t moduleEnd = GetModuleImageEnd(moduleBase);
	if (moduleEnd <= moduleBase) {
		if (outTrace != nullptr) {
			*outTrace = "hwnd_heuristic_module_range_invalid";
		}
		return false;
	}

	const std::uint32_t hwndValue = static_cast<std::uint32_t>(
		reinterpret_cast<std::uintptr_t>(targetHwnd));
	const size_t kMaxScannableRegionBytes = 64u * 1024u * 1024u;
	std::uintptr_t bestObject = 0;
	unsigned int bestPageType = 0;
	std::uintptr_t bestInnerObject = 0;
	std::uintptr_t bestVtable = 0;
	int bestScore = -1;
	size_t hitCount = 0;
	size_t scannedRegionCount = 0;
	size_t candidateCount = 0;

	SYSTEM_INFO systemInfo = {};
	GetSystemInfo(&systemInfo);
	std::uintptr_t cursor = reinterpret_cast<std::uintptr_t>(systemInfo.lpMinimumApplicationAddress);
	const std::uintptr_t limit = reinterpret_cast<std::uintptr_t>(systemInfo.lpMaximumApplicationAddress);
	while (cursor < limit) {
		MEMORY_BASIC_INFORMATION mbi = {};
		if (VirtualQuery(reinterpret_cast<LPCVOID>(cursor), &mbi, sizeof(mbi)) != sizeof(mbi)) {
			break;
		}

		const std::uintptr_t regionBase = reinterpret_cast<std::uintptr_t>(mbi.BaseAddress);
		const std::uintptr_t regionSize = static_cast<std::uintptr_t>(mbi.RegionSize);
		const std::uintptr_t nextCursor = regionBase + regionSize;
		if (nextCursor <= cursor) {
			break;
		}

		if (mbi.State == MEM_COMMIT &&
			IsReadableCommittedProtection(mbi.Protect) &&
			regionSize >= sizeof(std::uint32_t) &&
			regionSize <= kMaxScannableRegionBytes) {
			++scannedRegionCount;
			const size_t wordCount = static_cast<size_t>(regionSize / sizeof(std::uint32_t));
			std::vector<std::uint32_t> words(wordCount, 0);
			if (!TryReadMemoryBytesSafe(
					regionBase,
					words.data(),
					wordCount * sizeof(std::uint32_t))) {
				cursor = nextCursor;
				continue;
			}

			for (size_t index = 0; index < wordCount; ++index) {
				if (words[index] != hwndValue) {
					continue;
				}

				++hitCount;
				const std::uintptr_t hitAddress = regionBase + index * sizeof(std::uint32_t);
				if (hitAddress < 0x1C) {
					continue;
				}

				const std::uintptr_t candidateObject = hitAddress - 0x1C;
				std::uint32_t fields[24] = {};
				if (!TryReadUint32ArraySafe(candidateObject, fields, std::size(fields)) ||
					fields[7] != hwndValue) {
					continue;
				}

				++candidateCount;
				const std::uintptr_t vtable = static_cast<std::uintptr_t>(fields[0]);
				const bool vtableInModule = vtable >= moduleBase && vtable < moduleEnd;
				unsigned int pageType = 0;
				const std::uintptr_t innerObject = ResolveInnerEditorObjectFromFields(fields, &pageType);
				const bool pageTypeOk = IsRecognizedEditorPageType(pageType);
				const bool innerReadable =
					innerObject != 0 &&
					innerObject > 0x10000 &&
					IsReadableAddressRange(innerObject, sizeof(std::uint32_t) * 8);

				int score = 0;
				if (fields[7] == hwndValue) {
					score += 100;
				}
				if (vtableInModule) {
					score += 200;
				}
				if (pageTypeOk) {
					score += 1000;
				}
				if (innerObject != 0) {
					score += 80;
				}
				if (innerReadable) {
					score += 120;
				}
				if (fields[1] == 1) {
					score += 10;
				}
				if (fields[5] == 1) {
					score += 10;
				}

				if (score > bestScore) {
					bestScore = score;
					bestObject = candidateObject;
					bestPageType = pageType;
					bestInnerObject = innerObject;
					bestVtable = vtable;
				}
			}
		}

		cursor = nextCursor;
	}

	if (bestObject == 0) {
		if (outTrace != nullptr) {
			*outTrace =
				"hwnd_heuristic_failed"
				"|hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(targetHwnd)) +
				"|hits=" + std::to_string(hitCount) +
				"|candidates=" + std::to_string(candidateCount) +
				"|regions=" + std::to_string(scannedRegionCount);
		}
		return false;
	}

	if (outEditorObject != nullptr) {
		*outEditorObject = bestObject;
	}
	if (outTrace != nullptr) {
		*outTrace =
			"hwnd_heuristic_ok"
			"|hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(targetHwnd)) +
			"|object=" + std::to_string(bestObject) +
			"|page_type=" + std::to_string(bestPageType) +
			"|inner=" + std::to_string(bestInnerObject) +
			"|vtable=" + std::to_string(bestVtable) +
			"|score=" + std::to_string(bestScore) +
			"|hits=" + std::to_string(hitCount) +
			"|candidates=" + std::to_string(candidateCount) +
			"|regions=" + std::to_string(scannedRegionCount);
	}
	return true;
}

bool TryResolveEditorObjectFromMdiChildHeuristic(
	HWND mdiChildHwnd,
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	HWND* outMatchedHwnd,
	std::string* outTrace)
{
	if (outEditorObject != nullptr) {
		*outEditorObject = 0;
	}
	if (outMatchedHwnd != nullptr) {
		*outMatchedHwnd = nullptr;
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (mdiChildHwnd == nullptr || !IsWindow(mdiChildHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "mdi_child_invalid";
		}
		return false;
	}

	const auto candidates = CollectEditorWindowCandidatesFromMdiChild(mdiChildHwnd);
	if (candidates.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "editor_hwnd_candidates_missing";
		}
		return false;
	}

	std::string attemptsSummary;
	for (const auto& candidate : candidates) {
		std::uintptr_t editorObject = 0;
		std::string candidateTrace;
		if (TryResolveEditorObjectFromWindowHandleHeuristic(
				candidate.hWnd,
				moduleBase,
				&editorObject,
				&candidateTrace) &&
			editorObject != 0) {
			if (outEditorObject != nullptr) {
				*outEditorObject = editorObject;
			}
			if (outMatchedHwnd != nullptr) {
				*outMatchedHwnd = candidate.hWnd;
			}
			if (outTrace != nullptr) {
				*outTrace =
					"mdi_hwnd_heuristic_ok"
					"|candidate_count=" + std::to_string(candidates.size()) +
					"|matched_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(candidate.hWnd)) +
					"|matched_class=" + candidate.className +
					"|matched_depth=" + std::to_string(candidate.depth) +
					"|matched_score=" + std::to_string(candidate.score) +
					"|" +
					candidateTrace;
			}
			return true;
		}

		if (attemptsSummary.size() < 320) {
			if (!attemptsSummary.empty()) {
				attemptsSummary += ";";
			}
			attemptsSummary +=
				std::to_string(reinterpret_cast<std::uintptr_t>(candidate.hWnd)) +
				"/" +
				candidate.className +
				"/" +
				candidateTrace;
		}
	}

	if (outTrace != nullptr) {
		*outTrace =
			"mdi_hwnd_heuristic_failed"
			"|candidate_count=" + std::to_string(candidates.size());
		if (!attemptsSummary.empty()) {
			*outTrace += "|attempts=" + attemptsSummary;
		}
	}
	return false;
}

bool TryResolveEditorWindowFromMdiChild(
	HWND mdiChildHwnd,
	HWND* outEditorHwnd,
	std::string* outTrace)
{
	if (outEditorHwnd != nullptr) {
		*outEditorHwnd = nullptr;
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (mdiChildHwnd == nullptr || !IsWindow(mdiChildHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "mdi_child_invalid";
		}
		return false;
	}

	GUITHREADINFO guiInfo = {};
	guiInfo.cbSize = sizeof(guiInfo);
	if (GetGUIThreadInfo(GetCurrentThreadId(), &guiInfo) &&
		guiInfo.hwndFocus != nullptr &&
		IsSameOrChildWindow(mdiChildHwnd, guiInfo.hwndFocus)) {
		const std::string focusClass = WindowClassToString(guiInfo.hwndFocus);
		if (!IsIgnoredEditorCandidateClass(focusClass)) {
			if (outEditorHwnd != nullptr) {
				*outEditorHwnd = guiInfo.hwndFocus;
			}
			if (outTrace != nullptr) {
				*outTrace =
					"editor_hwnd_focus_ok"
					"|class=" + focusClass +
					"|title=" + WindowTextToString(guiInfo.hwndFocus);
			}
			return true;
		}
	}

	std::vector<WindowDepthInfo> descendants;
	CollectDescendantWindowsWithDepth(mdiChildHwnd, 1, descendants);

	HWND bestHwnd = nullptr;
	int bestScore = -2147483647;
	int bestDepth = 0;
	std::string bestClass;
	std::string bestTitle;
	for (const auto& info : descendants) {
		if (info.hWnd == nullptr || !IsWindow(info.hWnd) || !IsWindowVisible(info.hWnd)) {
			continue;
		}

		const std::string className = WindowClassToString(info.hWnd);
		if (IsIgnoredEditorCandidateClass(className)) {
			continue;
		}

		const std::string title = WindowTextToString(info.hWnd);
		int score = info.depth * 100;
		if (IsPreferredEditorCandidateClass(className)) {
			score += 1000;
		}
		if (title.empty()) {
			score += 100;
		}
		if (GetWindow(info.hWnd, GW_CHILD) == nullptr) {
			score += 20;
		}

		if (bestHwnd == nullptr || score > bestScore) {
			bestHwnd = info.hWnd;
			bestScore = score;
			bestDepth = info.depth;
			bestClass = className;
			bestTitle = title;
		}
	}

	if (bestHwnd == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				"editor_hwnd_candidate_missing"
				"|mdi_class=" + WindowClassToString(mdiChildHwnd) +
				"|mdi_title=" + WindowTextToString(mdiChildHwnd);
		}
		return false;
	}

	if (outEditorHwnd != nullptr) {
		*outEditorHwnd = bestHwnd;
	}
	if (outTrace != nullptr) {
		*outTrace =
			"editor_hwnd_candidate_ok"
			"|class=" + bestClass +
			"|title=" + bestTitle +
			"|depth=" + std::to_string(bestDepth) +
			"|count=" + std::to_string(descendants.size());
	}
	return true;
}

HTREEITEM GetTreeNextItemInternal(HWND treeHwnd, HTREEITEM item, UINT code)
{
	return reinterpret_cast<HTREEITEM>(
		SendMessageA(treeHwnd, TVM_GETNEXTITEM, static_cast<WPARAM>(code), reinterpret_cast<LPARAM>(item)));
}

bool QueryTreeItemInfoInternal(
	HWND treeHwnd,
	HTREEITEM item,
	std::string* outText,
	unsigned int* outItemData,
	int* outChildCount)
{
	if (outText != nullptr) {
		outText->clear();
	}
	if (outItemData != nullptr) {
		*outItemData = 0;
	}
	if (outChildCount != nullptr) {
		*outChildCount = 0;
	}
	if (treeHwnd == nullptr || !IsWindow(treeHwnd) || item == nullptr) {
		return false;
	}

	char textBuf[512] = {};
	TVITEMA tvItem = {};
	tvItem.mask = TVIF_HANDLE | TVIF_TEXT | TVIF_PARAM | TVIF_CHILDREN;
	tvItem.hItem = item;
	tvItem.pszText = textBuf;
	tvItem.cchTextMax = static_cast<int>(sizeof(textBuf));
	if (SendMessageA(treeHwnd, TVM_GETITEMA, 0, reinterpret_cast<LPARAM>(&tvItem)) == FALSE) {
		return false;
	}

	if (outText != nullptr) {
		*outText = textBuf;
	}
	if (outItemData != nullptr) {
		*outItemData = static_cast<unsigned int>(tvItem.lParam);
	}
	if (outChildCount != nullptr) {
		*outChildCount = tvItem.cChildren;
	}
	return true;
}

HWND FindProgramDataTreeViewWindow(std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	const HWND mainHwnd = ResolveMainIdeWindow();
	if (mainHwnd == nullptr || !IsWindow(mainHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "main_window_invalid";
		}
		return nullptr;
	}

	const auto treeWindows = CollectChildWindowsByClassName(mainHwnd, WC_TREEVIEWA);
	for (HWND treeHwnd : treeWindows) {
		const HTREEITEM rootItem = GetTreeNextItemInternal(treeHwnd, nullptr, TVGN_ROOT);
		std::string rootText;
		if (!QueryTreeItemInfoInternal(treeHwnd, rootItem, &rootText, nullptr, nullptr)) {
			continue;
		}
		if (rootText == "程序数据") {
			if (outTrace != nullptr) {
				*outTrace =
					"tree_ok|hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(treeHwnd)) +
					"|root=" + rootText;
			}
			return treeHwnd;
		}
	}

	if (outTrace != nullptr) {
		*outTrace = "program_tree_not_found";
	}
	return nullptr;
}

HTREEITEM FindProgramTreeItemByDataRecursive(
	HWND treeHwnd,
	HTREEITEM firstItem,
	unsigned int targetItemData,
	int depth,
	int maxDepth,
	std::string* outItemText)
{
	if (treeHwnd == nullptr || firstItem == nullptr || depth > maxDepth) {
		return nullptr;
	}

	for (HTREEITEM item = firstItem; item != nullptr; item = GetTreeNextItemInternal(treeHwnd, item, TVGN_NEXT)) {
		unsigned int itemData = 0;
		std::string itemText;
		int childCount = 0;
		if (!QueryTreeItemInfoInternal(treeHwnd, item, &itemText, &itemData, &childCount)) {
			continue;
		}

		if (itemData == targetItemData) {
			if (outItemText != nullptr) {
				*outItemText = itemText;
			}
			return item;
		}

		if (depth < maxDepth && childCount > 0) {
			if (HTREEITEM foundItem = FindProgramTreeItemByDataRecursive(
					treeHwnd,
					GetTreeNextItemInternal(treeHwnd, item, TVGN_CHILD),
					targetItemData,
					depth + 1,
					maxDepth,
					outItemText)) {
				return foundItem;
			}
		}
	}
	return nullptr;
}

HWND GetActiveMdiChildWindow(HWND mainHwnd)
{
	for (HWND mdiHwnd : CollectChildWindowsByClassName(mainHwnd, "MDIClient")) {
		const HWND activeChild = reinterpret_cast<HWND>(SendMessageA(mdiHwnd, WM_MDIGETACTIVE, 0, 0));
		if (activeChild != nullptr && IsWindow(activeChild)) {
			return activeChild;
		}
	}
	return nullptr;
}

bool TriggerProgramTreeItemOpenByUi(
	HWND treeHwnd,
	HTREEITEM item,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (treeHwnd == nullptr || !IsWindow(treeHwnd) || item == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "tree_open_invalid_argument";
		}
		return false;
	}

	SendMessageA(treeHwnd, TVM_ENSUREVISIBLE, 0, reinterpret_cast<LPARAM>(item));
	SendMessageA(treeHwnd, TVM_SELECTITEM, TVGN_CARET, reinterpret_cast<LPARAM>(item));
	SetFocus(treeHwnd);
	PumpPendingMessages();
	Sleep(40);
	DispatchSyntheticKey(treeHwnd, VK_RETURN, false);
	PumpPendingMessages();
	Sleep(80);

	if (outTrace != nullptr) {
		*outTrace = "tree_open_triggered";
	}
	return true;
}

bool TryResolveProgramTreeItemEditorWindowByUi(
	unsigned int itemData,
	HWND* outEditorHwnd,
	std::string* outTrace)
{
	if (outEditorHwnd != nullptr) {
		*outEditorHwnd = nullptr;
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	std::string treeTrace;
	const HWND mainHwnd = ResolveMainIdeWindow();
	const HWND treeHwnd = FindProgramDataTreeViewWindow(&treeTrace);
	if (mainHwnd == nullptr || treeHwnd == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace.empty() ? "program_tree_unavailable" : treeTrace;
		}
		return false;
	}

	const HTREEITEM rootItem = GetTreeNextItemInternal(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM firstChild = GetTreeNextItemInternal(treeHwnd, rootItem, TVGN_CHILD);
	std::string itemText;
	const HTREEITEM item = FindProgramTreeItemByDataRecursive(
		treeHwnd,
		firstChild,
		itemData,
		0,
		8,
		&itemText);
	if (item == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				treeTrace +
				"|tree_item_not_found|item_data=" + std::to_string(itemData);
		}
		return false;
	}

	const HWND activeBefore = GetActiveMdiChildWindow(mainHwnd);
	std::string openTrace;
	if (!TriggerProgramTreeItemOpenByUi(treeHwnd, item, &openTrace)) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace + "|" + openTrace;
		}
		return false;
	}

	for (int attempt = 0; attempt < 60; ++attempt) {
		PumpPendingMessages();
		const HWND activeChild = GetActiveMdiChildWindow(mainHwnd);
		if (activeChild != nullptr && IsWindow(activeChild)) {
			const std::string activeTitle = WindowTextToString(activeChild);
			const bool titleMatches =
				!itemText.empty() &&
				activeTitle.find(itemText) != std::string::npos;
			if (activeChild == activeBefore && !titleMatches) {
				Sleep(25);
				continue;
			}

			HWND editorHwnd = nullptr;
			std::string editorTrace;
			if (TryResolveEditorWindowFromMdiChild(activeChild, &editorHwnd, &editorTrace) &&
				editorHwnd != nullptr &&
				IsWindow(editorHwnd)) {
				if (outEditorHwnd != nullptr) {
					*outEditorHwnd = editorHwnd;
				}
				if (outTrace != nullptr) {
					*outTrace =
						treeTrace +
						"|" +
						openTrace +
						"|resolve_editor_hwnd_ui_ok" +
						"|item=" + itemText +
						"|active_changed=" + std::to_string(activeChild != activeBefore ? 1 : 0) +
						"|active_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(activeChild)) +
						"|active_class=" + WindowClassToString(activeChild) +
						"|active_text=" + activeTitle +
						"|" +
						editorTrace;
				}
				return true;
			}
		}
		Sleep(25);
	}

	if (outTrace != nullptr) {
		*outTrace =
			treeTrace +
			"|" +
			openTrace +
			"|resolve_editor_hwnd_ui_failed|item=" + itemText;
	}
	return false;
}

bool TryResolveProgramTreeItemEditorObjectByUi(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	std::string* outTrace)
{
	if (outEditorObject != nullptr) {
		*outEditorObject = 0;
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	std::string treeTrace;
	const HWND mainHwnd = ResolveMainIdeWindow();
	const HWND treeHwnd = FindProgramDataTreeViewWindow(&treeTrace);
	if (mainHwnd == nullptr || treeHwnd == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace.empty() ? "program_tree_unavailable" : treeTrace;
		}
		return false;
	}

	const HTREEITEM rootItem = GetTreeNextItemInternal(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM firstChild = GetTreeNextItemInternal(treeHwnd, rootItem, TVGN_CHILD);
	std::string itemText;
	const HTREEITEM item = FindProgramTreeItemByDataRecursive(
		treeHwnd,
		firstChild,
		itemData,
		0,
		8,
		&itemText);
	if (item == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				treeTrace +
				"|tree_item_not_found|item_data=" + std::to_string(itemData);
		}
		return false;
	}

	const HWND activeBefore = GetActiveMdiChildWindow(mainHwnd);
	std::string openTrace;
	if (!TriggerProgramTreeItemOpenByUi(treeHwnd, item, &openTrace)) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace + "|" + openTrace;
		}
		return false;
	}

	for (int attempt = 0; attempt < 60; ++attempt) {
		PumpPendingMessages();
		const HWND activeChild = GetActiveMdiChildWindow(mainHwnd);
		if (activeChild != nullptr && IsWindow(activeChild)) {
			const std::string activeTitle = WindowTextToString(activeChild);
			const bool titleMatches =
				!itemText.empty() &&
				activeTitle.find(itemText) != std::string::npos;
			if (activeChild == activeBefore && !titleMatches) {
				Sleep(25);
				continue;
			}

			std::uintptr_t editorObject = 0;
			HWND matchedEditorHwnd = nullptr;
			std::string heuristicTrace;
			if (TryResolveEditorObjectFromMdiChildHeuristic(
					activeChild,
					moduleBase,
					&editorObject,
					&matchedEditorHwnd,
					&heuristicTrace) &&
				editorObject != 0) {
				if (outEditorObject != nullptr) {
					*outEditorObject = editorObject;
				}
				if (outTrace != nullptr) {
					*outTrace =
						treeTrace +
						"|" +
						openTrace +
						"|resolve_editor_ui_ok" +
						"|item=" + itemText +
						"|active_changed=" + std::to_string(activeChild != activeBefore ? 1 : 0) +
						"|active_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(activeChild)) +
						"|active_class=" + WindowClassToString(activeChild) +
						"|active_text=" + activeTitle +
						"|matched_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(matchedEditorHwnd)) +
						"|" +
						heuristicTrace;
				}
				return true;
			}
		}
		Sleep(25);
	}

	if (outTrace != nullptr) {
		*outTrace =
			treeTrace +
			"|" +
			openTrace +
			"|resolve_editor_ui_failed|item=" + itemText;
	}
	return false;
}

HWND FindMdiChildWindow(HWND hWnd)
{
	HWND current = hWnd;
	while (current != nullptr && IsWindow(current)) {
		const HWND parent = GetParent(current);
		if (parent != nullptr && WindowClassToString(parent) == "MDIClient") {
			return current;
		}
		current = parent;
	}
	return nullptr;
}

std::vector<HWND> CollectWindowRouteChain(HWND hWnd)
{
	std::vector<HWND> chain;
	HWND current = hWnd;
	while (current != nullptr && IsWindow(current)) {
		if (std::find(chain.begin(), chain.end(), current) == chain.end()) {
			chain.push_back(current);
		}
		current = GetParent(current);
	}
	const HWND root = GetAncestor(hWnd, GA_ROOT);
	if (root != nullptr &&
		IsWindow(root) &&
		std::find(chain.begin(), chain.end(), root) == chain.end()) {
		chain.push_back(root);
	}
	return chain;
}

bool TryReadEditorWindowHandleFromObject(std::uintptr_t editorObject, HWND* outHwnd)
{
	if (outHwnd != nullptr) {
		*outHwnd = nullptr;
	}
	if (editorObject == 0) {
		return false;
	}

	__try {
		const auto* fields = reinterpret_cast<const std::uint32_t*>(editorObject);
		const HWND hWnd = reinterpret_cast<HWND>(static_cast<UINT_PTR>(fields[7]));
		if (hWnd != nullptr && IsWindow(hWnd)) {
			if (outHwnd != nullptr) {
				*outHwnd = hWnd;
			}
			return true;
		}
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
	}
	return false;
}

HWND ResolveEditorInputWindow(std::uintptr_t editorObject, std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	EditorDispatchTargetInfo targetInfo{};
	if (TryResolveInnerEditorObject(editorObject, &targetInfo) &&
		targetInfo.innerObject != 0 &&
		targetInfo.innerObject != editorObject) {
		HWND innerHwnd = nullptr;
		if (TryReadEditorWindowHandleFromObject(targetInfo.innerObject, &innerHwnd)) {
			if (outTrace != nullptr) {
				*outTrace =
					"resolve_hwnd=inner"
					"|type=" + std::to_string(targetInfo.pageType) +
					"|class=" + WindowClassToString(innerHwnd);
			}
			return innerHwnd;
		}
	}

	HWND rawHwnd = nullptr;
	if (TryReadEditorWindowHandleFromObject(editorObject, &rawHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "resolve_hwnd=raw|class=" + WindowClassToString(rawHwnd);
		}
		return rawHwnd;
	}

	GUITHREADINFO guiInfo = {};
	guiInfo.cbSize = sizeof(guiInfo);
	if (GetGUIThreadInfo(GetCurrentThreadId(), &guiInfo)) {
		HWND focusHwnd = guiInfo.hwndCaret != nullptr ? guiInfo.hwndCaret : guiInfo.hwndFocus;
		if (focusHwnd != nullptr && IsWindow(focusHwnd)) {
			if (outTrace != nullptr) {
				*outTrace = "resolve_hwnd=current_focus|class=" + WindowClassToString(focusHwnd);
			}
			return focusHwnd;
		}
	}

	if (outTrace != nullptr) {
		*outTrace = "resolve_hwnd_failed";
	}
	return nullptr;
}

void SendVirtualKey(WORD vk)
{
	INPUT inputs[2] = {};
	inputs[0].type = INPUT_KEYBOARD;
	inputs[0].ki.wVk = vk;
	inputs[1].type = INPUT_KEYBOARD;
	inputs[1].ki.wVk = vk;
	inputs[1].ki.dwFlags = KEYEVENTF_KEYUP;
	(void)SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT));
}

void SendCtrlChord(WORD vk)
{
	INPUT inputs[4] = {};
	inputs[0].type = INPUT_KEYBOARD;
	inputs[0].ki.wVk = VK_CONTROL;
	inputs[1].type = INPUT_KEYBOARD;
	inputs[1].ki.wVk = vk;
	inputs[2].type = INPUT_KEYBOARD;
	inputs[2].ki.wVk = vk;
	inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;
	inputs[3].type = INPUT_KEYBOARD;
	inputs[3].ki.wVk = VK_CONTROL;
	inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;
	(void)SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT));
}

enum class EditorInputSequence {
	SelectAllCopy,
	SelectAllDeletePaste,
};

struct EditorInputWorkerContext {
	EditorInputSequence sequence = EditorInputSequence::SelectAllCopy;
	std::wstring windowTitle;
};

LPARAM BuildKeyLParam(WORD vk, bool keyUp)
{
	const UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
	LPARAM lParam = 1 | (static_cast<LPARAM>(scanCode) << 16);
	if (keyUp) {
		lParam |= 0xC0000000;
	}
	return lParam;
}

void DispatchSyntheticKey(HWND editorHwnd, WORD vk, bool ctrlDown)
{
	BYTE originalKeyboardState[256] = {};
	if (!GetKeyboardState(originalKeyboardState)) {
		std::memset(originalKeyboardState, 0, sizeof(originalKeyboardState));
	}

	BYTE modifiedKeyboardState[256] = {};
	std::memcpy(modifiedKeyboardState, originalKeyboardState, sizeof(modifiedKeyboardState));
	if (ctrlDown) {
		modifiedKeyboardState[VK_CONTROL] |= 0x80;
	}
	SetKeyboardState(modifiedKeyboardState);

	MSG keyDown = {};
	keyDown.hwnd = editorHwnd;
	keyDown.message = WM_KEYDOWN;
	keyDown.wParam = vk;
	keyDown.lParam = BuildKeyLParam(vk, false);
	if (!AfxPreTranslateMessage(&keyDown)) {
		TranslateMessage(&keyDown);
		DispatchMessage(&keyDown);
	}

	MSG keyUp = {};
	keyUp.hwnd = editorHwnd;
	keyUp.message = WM_KEYUP;
	keyUp.wParam = vk;
	keyUp.lParam = BuildKeyLParam(vk, true);
	if (!AfxPreTranslateMessage(&keyUp)) {
		TranslateMessage(&keyUp);
		DispatchMessage(&keyUp);
	}

	SetKeyboardState(originalKeyboardState);
}

void ActivateEditorWindowByClick(HWND editorHwnd)
{
	if (editorHwnd == nullptr || !IsWindow(editorHwnd)) {
		return;
	}

	RECT clientRect = {};
	GetClientRect(editorHwnd, &clientRect);
	const int x = clientRect.right > 8 ? 8 : (clientRect.right > 0 ? clientRect.right / 2 : 1);
	const int y = clientRect.bottom > 8 ? 8 : (clientRect.bottom > 0 ? clientRect.bottom / 2 : 1);
	const LPARAM clickPos = MAKELPARAM(x, y);
	SendMessageA(editorHwnd, WM_MOUSEMOVE, 0, clickPos);
	SendMessageA(editorHwnd, WM_LBUTTONDOWN, MK_LBUTTON, clickPos);
	SendMessageA(editorHwnd, WM_LBUTTONUP, 0, clickPos);
	PumpPendingMessages();
}

bool InvokeDispatchMethod(
	IDispatch* dispatch,
	const wchar_t* methodName,
	VARIANT* args,
	UINT argCount)
{
	if (dispatch == nullptr || methodName == nullptr || methodName[0] == L'\0') {
		return false;
	}

	DISPID dispid = DISPID_UNKNOWN;
	LPOLESTR nameBuffer = const_cast<LPOLESTR>(methodName);
	if (FAILED(dispatch->GetIDsOfNames(IID_NULL, &nameBuffer, 1, LOCALE_USER_DEFAULT, &dispid))) {
		return false;
	}

	DISPPARAMS params = {};
	params.rgvarg = args;
	params.cArgs = argCount;
	return SUCCEEDED(dispatch->Invoke(
		dispid,
		IID_NULL,
		LOCALE_USER_DEFAULT,
		DISPATCH_METHOD,
		&params,
		nullptr,
		nullptr,
		nullptr));
}

bool WshSendKeys(
	IDispatch* dispatch,
	const wchar_t* keys)
{
	VARIANT args[2] = {};
	VariantInit(&args[0]);
	VariantInit(&args[1]);
	args[0].vt = VT_BOOL;
	args[0].boolVal = VARIANT_TRUE;
	args[1].vt = VT_BSTR;
	args[1].bstrVal = SysAllocString(keys);
	const bool ok = InvokeDispatchMethod(dispatch, L"SendKeys", args, 2);
	VariantClear(&args[1]);
	VariantClear(&args[0]);
	return ok;
}

bool PumpMessagesWhileWorkerRuns(HANDLE workerHandle, DWORD timeoutMs);

std::wstring EscapePowerShellSingleQuoted(const std::wstring& text)
{
	std::wstring escaped;
	escaped.reserve(text.size() + 8);
	for (wchar_t ch : text) {
		if (ch == L'\'') {
			escaped += L"''";
		}
		else {
			escaped.push_back(ch);
		}
	}
	return escaped;
}

bool RunExternalSendKeysHelper(
	const std::wstring& windowTitle,
	EditorInputSequence sequence)
{
	if (windowTitle.empty()) {
		return false;
	}

	std::wstring activateTitle = windowTitle;
	const size_t titleSplitPos = activateTitle.find(L" - ");
	if (titleSplitPos != std::wstring::npos && titleSplitPos > 0) {
		activateTitle.resize(titleSplitPos);
	}
	const std::wstring escapedTitle = EscapePowerShellSingleQuoted(activateTitle);
	std::wstring script =
		L"$ws=New-Object -ComObject WScript.Shell; "
		L"$null=$ws.AppActivate('" + escapedTitle + L"'); "
		L"Start-Sleep -Milliseconds 120; "
		L"$ws.SendKeys('^a',$true); ";
	if (sequence == EditorInputSequence::SelectAllCopy) {
		script += L"Start-Sleep -Milliseconds 120; $ws.SendKeys('^c',$true);";
	}
	else if (sequence == EditorInputSequence::SelectAllDeletePaste) {
		script += L"Start-Sleep -Milliseconds 120; $ws.SendKeys('{DEL}',$true); Start-Sleep -Milliseconds 120; $ws.SendKeys('^v',$true);";
	}
	else {
		return false;
	}

	std::wstring commandLine =
		L"powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -Command \"" +
		script +
		L"\"";
	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');

	STARTUPINFOW startupInfo = {};
	startupInfo.cb = sizeof(startupInfo);
	startupInfo.dwFlags = STARTF_USESHOWWINDOW;
	startupInfo.wShowWindow = SW_HIDE;
	PROCESS_INFORMATION processInfo = {};
	const BOOL createOk = CreateProcessW(
		nullptr,
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		FALSE,
		CREATE_NO_WINDOW,
		nullptr,
		nullptr,
		&startupInfo,
		&processInfo);
	if (!createOk) {
		return false;
	}

	CloseHandle(processInfo.hThread);
	const bool waitOk = PumpMessagesWhileWorkerRuns(processInfo.hProcess, 4000);
	CloseHandle(processInfo.hProcess);
	return waitOk;
}

bool RunInProcessWshSendKeys(
	const std::wstring& windowTitle,
	EditorInputSequence sequence)
{
	if (windowTitle.empty()) {
		return false;
	}

	std::wstring activateTitle = windowTitle;
	const size_t titleSplitPos = activateTitle.find(L" - ");
	if (titleSplitPos != std::wstring::npos && titleSplitPos > 0) {
		activateTitle.resize(titleSplitPos);
	}

	const HRESULT initHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (FAILED(initHr)) {
		return false;
	}

	CLSID clsid = {};
	if (FAILED(CLSIDFromProgID(L"WScript.Shell", &clsid))) {
		CoUninitialize();
		return false;
	}

	IDispatch* dispatch = nullptr;
	if (FAILED(CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_IDispatch, reinterpret_cast<void**>(&dispatch))) ||
		dispatch == nullptr) {
		CoUninitialize();
		return false;
	}

	VARIANT activateArg = {};
	VariantInit(&activateArg);
	activateArg.vt = VT_BSTR;
	activateArg.bstrVal = SysAllocString(activateTitle.c_str());
	const bool activateOk = InvokeDispatchMethod(dispatch, L"AppActivate", &activateArg, 1);
	VariantClear(&activateArg);

	Sleep(120);
	bool sendOk = false;
	switch (sequence) {
	case EditorInputSequence::SelectAllCopy:
		sendOk = WshSendKeys(dispatch, L"^a");
		Sleep(120);
		sendOk = WshSendKeys(dispatch, L"^c") && sendOk;
		break;
	case EditorInputSequence::SelectAllDeletePaste:
		sendOk = WshSendKeys(dispatch, L"^a");
		Sleep(120);
		sendOk = WshSendKeys(dispatch, L"{DEL}") && sendOk;
		Sleep(120);
		sendOk = WshSendKeys(dispatch, L"^v") && sendOk;
		break;
	default:
		sendOk = false;
		break;
	}

	dispatch->Release();
	CoUninitialize();
	return activateOk && sendOk;
}

DWORD WINAPI EditorInputWorkerProc(LPVOID parameter)
{
	auto* context = reinterpret_cast<EditorInputWorkerContext*>(parameter);
	if (context == nullptr) {
		return 0;
	}
	return RunInProcessWshSendKeys(context->windowTitle, context->sequence) ? 1u : 0u;
}

bool RunWorkerSendKeysHelper(
	const std::wstring& windowTitle,
	EditorInputSequence sequence,
	bool* outUsedExternal)
{
	if (outUsedExternal != nullptr) {
		*outUsedExternal = false;
	}
	if (windowTitle.empty()) {
		return false;
	}

	if (RunExternalSendKeysHelper(windowTitle, sequence)) {
		if (outUsedExternal != nullptr) {
			*outUsedExternal = true;
		}
		return true;
	}

	EditorInputWorkerContext context{};
	context.sequence = sequence;
	context.windowTitle = windowTitle;

	HANDLE workerHandle = CreateThread(nullptr, 0, EditorInputWorkerProc, &context, 0, nullptr);
	if (workerHandle != nullptr) {
		const bool pumpOk = PumpMessagesWhileWorkerRuns(workerHandle, 4000);
		DWORD exitCode = 0;
		const bool exitOk = GetExitCodeThread(workerHandle, &exitCode) != FALSE;
		CloseHandle(workerHandle);
		if (pumpOk && exitOk && exitCode == 1u) {
			return true;
		}
	}
	return false;
}

bool PumpMessagesWhileWorkerRuns(HANDLE workerHandle, DWORD timeoutMs)
{
	if (workerHandle == nullptr) {
		return false;
	}

	const DWORD startTick = GetTickCount();
	for (;;) {
		MSG msg = {};
		while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
			if (!AfxPreTranslateMessage(&msg)) {
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}

		const DWORD elapsed = GetTickCount() - startTick;
		if (elapsed >= timeoutMs) {
			return WaitForSingleObject(workerHandle, 0) == WAIT_OBJECT_0;
		}

		const DWORD waitMs = (std::min)(static_cast<DWORD>(20), timeoutMs - elapsed);
		const DWORD waitResult = MsgWaitForMultipleObjects(1, &workerHandle, FALSE, waitMs, QS_ALLINPUT);
		if (waitResult == WAIT_OBJECT_0) {
			break;
		}
		if (waitResult != WAIT_TIMEOUT && waitResult != WAIT_OBJECT_0 + 1) {
			return false;
		}
	}

	const DWORD settleUntil = GetTickCount() + 120;
	while (GetTickCount() < settleUntil) {
		MSG msg = {};
		bool hadMessage = false;
		while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
			hadMessage = true;
			if (!AfxPreTranslateMessage(&msg)) {
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
		if (!hadMessage) {
			Sleep(10);
		}
	}
	return true;
}

bool RunEditorInputSequence(
	HWND editorHwnd,
	EditorInputSequence sequence,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "editor_hwnd_invalid";
		}
		return false;
	}

	HWND rootHwnd = GetAncestor(editorHwnd, GA_ROOT);
	const HWND mdiChildHwnd = FindMdiChildWindow(editorHwnd);
	const HWND foregroundBefore = GetForegroundWindow();
	if (rootHwnd != nullptr && IsWindow(rootHwnd)) {
		const DWORD currentThreadId = GetCurrentThreadId();
		const DWORD foregroundThreadId = foregroundBefore != nullptr
			? GetWindowThreadProcessId(foregroundBefore, nullptr)
			: 0;
		const bool attachedInput =
			foregroundThreadId != 0 &&
			foregroundThreadId != currentThreadId &&
			AttachThreadInput(currentThreadId, foregroundThreadId, TRUE) != FALSE;
		if (IsIconic(rootHwnd)) {
			ShowWindow(rootHwnd, SW_RESTORE);
		}
		BringWindowToTop(rootHwnd);
		SetForegroundWindow(rootHwnd);
		SetActiveWindow(rootHwnd);
		if (attachedInput) {
			AttachThreadInput(currentThreadId, foregroundThreadId, FALSE);
		}
	}
	const HWND previousFocus = GetFocus();
	SetFocus(editorHwnd);
	const HWND focusedAfterSetFocus = GetFocus();
	const HWND foregroundAfterActivate = GetForegroundWindow();
	UpdateWindow(editorHwnd);

	bool pumpOk = false;
	std::string inputMode = "synthetic";
	if (rootHwnd != nullptr && IsWindow(rootHwnd)) {
		const int titleLength = GetWindowTextLengthW(rootHwnd);
		if (titleLength > 0) {
			std::wstring windowTitle;
			windowTitle.resize(static_cast<size_t>(titleLength) + 1);
			const int copied = GetWindowTextW(rootHwnd, windowTitle.data(), titleLength + 1);
			if (copied >= 0) {
				windowTitle.resize(static_cast<size_t>(copied));
			}
			bool usedExternal = false;
			if (RunWorkerSendKeysHelper(windowTitle, sequence, &usedExternal)) {
				pumpOk = true;
				inputMode = usedExternal ? "wsh_external" : "wsh_worker";
			}
		}
	}

	if (!pumpOk) {
		pumpOk = true;
		switch (sequence) {
		case EditorInputSequence::SelectAllCopy:
			DispatchSyntheticKey(editorHwnd, 'A', true);
			Sleep(50);
			DispatchSyntheticKey(editorHwnd, 'C', true);
			break;
		case EditorInputSequence::SelectAllDeletePaste:
			DispatchSyntheticKey(editorHwnd, 'A', true);
			Sleep(50);
			DispatchSyntheticKey(editorHwnd, VK_DELETE, false);
			Sleep(50);
			DispatchSyntheticKey(editorHwnd, 'V', true);
			break;
		default:
			pumpOk = false;
			break;
		}
		PumpPendingMessages();
	}

	if (previousFocus != nullptr && IsWindow(previousFocus)) {
		SetFocus(previousFocus);
	}

	if (outTrace != nullptr) {
		*outTrace =
			std::string("input_sequence_") +
			(pumpOk ? "ok" : "failed") +
			"|class=" + WindowClassToString(editorHwnd) +
			"|mdi_child_class=" + WindowClassToString(mdiChildHwnd) +
			"|mdi_child_title=" + WindowTextToString(mdiChildHwnd) +
			"|focus_after_set=" + WindowClassToString(focusedAfterSetFocus) +
			"|foreground_after=" + WindowClassToString(foregroundAfterActivate) +
			"|mode=" + inputMode;
	}
	return pumpOk;
}

bool RunEditorInputSequenceSyntheticOnly(
	HWND editorHwnd,
	EditorInputSequence sequence,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "editor_hwnd_invalid";
		}
		return false;
	}

	const HWND rootHwnd = GetAncestor(editorHwnd, GA_ROOT);
	const HWND mdiChildHwnd = FindMdiChildWindow(editorHwnd);
	const HWND foregroundBefore = GetForegroundWindow();
	if (rootHwnd != nullptr && IsWindow(rootHwnd)) {
		const DWORD currentThreadId = GetCurrentThreadId();
		const DWORD foregroundThreadId = foregroundBefore != nullptr
			? GetWindowThreadProcessId(foregroundBefore, nullptr)
			: 0;
		const bool attachedInput =
			foregroundThreadId != 0 &&
			foregroundThreadId != currentThreadId &&
			AttachThreadInput(currentThreadId, foregroundThreadId, TRUE) != FALSE;
		if (IsIconic(rootHwnd)) {
			ShowWindow(rootHwnd, SW_RESTORE);
		}
		BringWindowToTop(rootHwnd);
		SetForegroundWindow(rootHwnd);
		SetActiveWindow(rootHwnd);
		if (attachedInput) {
			AttachThreadInput(currentThreadId, foregroundThreadId, FALSE);
		}
	}

	const HWND previousFocus = GetFocus();
	SetFocus(editorHwnd);
	const HWND focusedAfterSetFocus = GetFocus();
	const HWND foregroundAfterActivate = GetForegroundWindow();
	UpdateWindow(editorHwnd);
	ActivateEditorWindowByClick(editorHwnd);

	bool pumpOk = true;
	switch (sequence) {
	case EditorInputSequence::SelectAllCopy:
		DispatchSyntheticKey(editorHwnd, 'A', true);
		Sleep(50);
		DispatchSyntheticKey(editorHwnd, 'C', true);
		break;
	case EditorInputSequence::SelectAllDeletePaste:
		DispatchSyntheticKey(editorHwnd, 'A', true);
		Sleep(50);
		DispatchSyntheticKey(editorHwnd, VK_DELETE, false);
		Sleep(50);
		DispatchSyntheticKey(editorHwnd, 'V', true);
		break;
	default:
		pumpOk = false;
		break;
	}
	PumpPendingMessages();

	if (previousFocus != nullptr && IsWindow(previousFocus)) {
		SetFocus(previousFocus);
	}

	if (outTrace != nullptr) {
		*outTrace =
			std::string("input_sequence_") +
			(pumpOk ? "ok" : "failed") +
			"|class=" + WindowClassToString(editorHwnd) +
			"|mdi_child_class=" + WindowClassToString(mdiChildHwnd) +
			"|mdi_child_title=" + WindowTextToString(mdiChildHwnd) +
			"|focus_after_set=" + WindowClassToString(focusedAfterSetFocus) +
			"|foreground_after=" + WindowClassToString(foregroundAfterActivate) +
			"|mode=synthetic_only";
	}
	return pumpOk;
}

bool SendWindowCommandToRoute(HWND editorHwnd, UINT commandId, std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd) || commandId == 0) {
		if (outTrace != nullptr) {
			*outTrace = "window_command_invalid";
		}
		return false;
	}

	const HWND rootHwnd = GetAncestor(editorHwnd, GA_ROOT);
	if (rootHwnd != nullptr && IsWindow(rootHwnd)) {
		BringWindowToTop(rootHwnd);
		SetForegroundWindow(rootHwnd);
		SetActiveWindow(rootHwnd);
	}
	SetFocus(editorHwnd);
	ActivateEditorWindowByClick(editorHwnd);

	const auto chain = CollectWindowRouteChain(editorHwnd);
	if (chain.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "window_command_chain_empty";
		}
		return false;
	}

	for (auto it = chain.rbegin(); it != chain.rend(); ++it) {
		const HWND targetHwnd = *it;
		if (targetHwnd == nullptr || !IsWindow(targetHwnd)) {
			continue;
		}
		SendMessageA(targetHwnd, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
		PumpPendingMessages();
	}

	if (outTrace != nullptr) {
		*outTrace =
			"window_command_sent"
			"|id=" + std::to_string(commandId) +
			"|targets=" + std::to_string(chain.size());
	}
	return true;
}

bool RunEditorInputSequenceByWindowRoute(
	HWND editorHwnd,
	EditorInputSequence sequence,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "editor_hwnd_invalid";
		}
		return false;
	}

	std::string selectTrace;
	if (!SendWindowCommandToRoute(editorHwnd, kEditorUiCmdSelectAll, &selectTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace.empty() ? "route_select_all_failed" : selectTrace;
		}
		return false;
	}
	Sleep(50);

	std::string actionTrace;
	switch (sequence) {
	case EditorInputSequence::SelectAllCopy:
		if (!SendWindowCommandToRoute(editorHwnd, kEditorUiCmdCopy, &actionTrace)) {
			if (outTrace != nullptr) {
				*outTrace = selectTrace + "|" + actionTrace;
			}
			return false;
		}
		break;
	case EditorInputSequence::SelectAllDeletePaste:
		DispatchSyntheticKey(editorHwnd, VK_DELETE, false);
		Sleep(50);
		if (!SendWindowCommandToRoute(editorHwnd, kEditorUiCmdPaste, &actionTrace)) {
			if (outTrace != nullptr) {
				*outTrace = selectTrace + "|" + actionTrace;
			}
			return false;
		}
		break;
	default:
		if (outTrace != nullptr) {
			*outTrace = "route_sequence_unsupported";
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace = selectTrace + "|" + actionTrace + "|mode=window_route";
	}
	return true;
}

bool CopyWholePageTextByEditorWindowInput(
	HWND editorHwnd,
	std::string* outCode,
	NativeRealPageAccessResult* outResult)
{
	if (outCode != nullptr) {
		outCode->clear();
	}
	if (outResult != nullptr) {
		*outResult = {};
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd) || outCode == nullptr) {
		if (outResult != nullptr) {
			outResult->trace = "editor_hwnd_invalid";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outResult != nullptr) {
			outResult->trace = "fake_clipboard_install_failed";
		}
		return false;
	}

	std::string inputTrace;
	if (!RunEditorInputSequenceByWindowRoute(editorHwnd, EditorInputSequence::SelectAllCopy, &inputTrace)) {
		if (outResult != nullptr) {
			outResult->trace = inputTrace.empty() ? "input_sequence_failed" : inputTrace;
		}
		return false;
	}

	std::string clipboardTrace;
	if ((!WaitForFakeClipboardText(fakeClipboard, outCode, &clipboardTrace) || outCode->empty()) &&
		!fakeClipboard.HasClipboardTraffic()) {
		std::string syntheticTrace;
		if (!RunEditorInputSequenceSyntheticOnly(editorHwnd, EditorInputSequence::SelectAllCopy, &syntheticTrace)) {
			if (outResult != nullptr) {
				outResult->trace =
					inputTrace +
					"|copy_text_format_missing|" +
					clipboardTrace +
					"|synthetic_failed|" +
					syntheticTrace;
			}
			return false;
		}
		inputTrace += "|fallback_synthetic|" + syntheticTrace;
	}

	if (!WaitForFakeClipboardText(fakeClipboard, outCode, &clipboardTrace) ||
		outCode->empty()) {
		if (outResult != nullptr) {
			outResult->trace =
				inputTrace +
				"|copy_text_format_missing|" +
				clipboardTrace;
		}
		return false;
	}

	if (outResult != nullptr) {
		outResult->ok = true;
		outResult->usedClipboardEmulation = true;
		outResult->capturedCustomFormat = fakeClipboard.HasCustomFormat();
		outResult->textBytes = outCode->size();
		outResult->trace =
			"copy_ok|window_input|" +
			inputTrace +
			"|" +
			clipboardTrace;
	}
	return true;
}

bool ReplaceWholePageTextByEditorWindowInput(
	HWND editorHwnd,
	const std::string& newPageCode,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd) || newPageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "editor_hwnd_invalid";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "fake_clipboard_install_failed";
		}
		return false;
	}

	HANDLE textHandle = CreateClipboardAnsiTextHandle(newPageCode);
	if (textHandle == nullptr || !fakeClipboard.SetFormatHandle(CF_TEXT, textHandle)) {
		if (textHandle != nullptr) {
			GlobalFree(textHandle);
		}
		if (outTrace != nullptr) {
			*outTrace = "paste_text_handle_failed";
		}
		return false;
	}

	std::string inputTrace;
	if (!RunEditorInputSequenceByWindowRoute(editorHwnd, EditorInputSequence::SelectAllDeletePaste, &inputTrace)) {
		if (outTrace != nullptr) {
			*outTrace = inputTrace.empty() ? "input_sequence_failed" : inputTrace;
		}
		return false;
	}

	if (!fakeClipboard.HasClipboardTraffic()) {
		std::string syntheticTrace;
		if (!RunEditorInputSequenceSyntheticOnly(editorHwnd, EditorInputSequence::SelectAllDeletePaste, &syntheticTrace)) {
			if (outTrace != nullptr) {
				*outTrace =
					inputTrace +
					"|synthetic_failed|" +
					syntheticTrace;
			}
			return false;
		}
		inputTrace += "|fallback_synthetic|" + syntheticTrace;
	}

	SettleEditorAfterCommand(120);
	if (outTrace != nullptr) {
		*outTrace =
			"replace_ok|window_input|" +
			inputTrace +
			"|" +
			fakeClipboard.BuildStatsText();
	}
	return true;
}

bool InvokeMfcCommandRoute(HWND editorHwnd, UINT commandId, std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd)) {
		if (outTrace != nullptr) {
			*outTrace = "command_route_invalid_hwnd";
		}
		return false;
	}

	for (HWND targetHwnd : CollectWindowRouteChain(editorHwnd)) {
		if (targetHwnd == nullptr || !IsWindow(targetHwnd)) {
			continue;
		}
		CWnd* targetWnd = CWnd::FromHandlePermanent(targetHwnd);
		if (targetWnd == nullptr) {
			continue;
		}
		if (targetWnd->OnCmdMsg(commandId, CN_COMMAND, nullptr, nullptr)) {
			if (outTrace != nullptr) {
				*outTrace =
					"command_route_ok"
					"|class=" + WindowClassToString(targetHwnd) +
					"|title=" + WindowTextToString(targetHwnd) +
					"|id=" + std::to_string(commandId);
			}
			return true;
		}
	}

	if (outTrace != nullptr) {
		*outTrace = "command_route_unhandled|id=" + std::to_string(commandId);
	}
	return false;
}

bool CopyWholePageTextByUiInput(
	std::uintptr_t editorObject,
	std::string* outCode,
	NativeRealPageAccessResult* outResult)
{
	if (outCode != nullptr) {
		outCode->clear();
	}
	ScopedRealClipboardTextRestore clipboardRestore;

	std::string resolveTrace;
	const HWND editorHwnd = ResolveEditorInputWindow(editorObject, &resolveTrace);
	if (editorHwnd == nullptr) {
		if (outResult != nullptr) {
			outResult->trace = resolveTrace.empty() ? "resolve_editor_hwnd_failed" : resolveTrace;
		}
		return false;
	}

	std::string inputTrace;
	if (!RunEditorInputSequence(editorHwnd, EditorInputSequence::SelectAllCopy, &inputTrace)) {
		if (outResult != nullptr) {
			outResult->trace = resolveTrace + "|" + inputTrace;
		}
		return false;
	}

	if (!ReadRealClipboardText(outCode) || outCode->empty()) {
		if (outResult != nullptr) {
			outResult->trace = resolveTrace + "|" + inputTrace + "|read_real_clipboard_failed";
		}
		return false;
	}

	if (outResult != nullptr) {
		outResult->usedClipboardEmulation = false;
		outResult->textBytes = outCode->size();
		outResult->trace = resolveTrace + "|" + inputTrace + "|copy_ok";
	}
	return true;
}

bool ReplaceWholePageTextByUiInput(
	std::uintptr_t editorObject,
	const std::string& newPageCode,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	ScopedRealClipboardTextRestore clipboardRestore;
	if (!SetRealClipboardText(newPageCode)) {
		if (outTrace != nullptr) {
			*outTrace = "set_real_clipboard_text_failed";
		}
		return false;
	}

	std::string resolveTrace;
	const HWND editorHwnd = ResolveEditorInputWindow(editorObject, &resolveTrace);
	if (editorHwnd == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = resolveTrace.empty() ? "resolve_editor_hwnd_failed" : resolveTrace;
		}
		return false;
	}

	std::string inputTrace;
	if (!RunEditorInputSequence(editorHwnd, EditorInputSequence::SelectAllDeletePaste, &inputTrace)) {
		if (outTrace != nullptr) {
			*outTrace = resolveTrace + "|" + inputTrace;
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace = resolveTrace + "|" + inputTrace + "|replace_ok";
	}
	return true;
}

bool TryResolveInnerEditorObject(std::uintptr_t rawObject, EditorDispatchTargetInfo* outInfo)
{
	if (outInfo != nullptr) {
		*outInfo = {};
		outInfo->rawObject = rawObject;
	}
	if (rawObject == 0) {
		return false;
	}

	__try {
		const auto* fields = reinterpret_cast<const std::uint32_t*>(rawObject);
		unsigned int pageType = 0;
		const std::uintptr_t innerObject = ResolveInnerEditorObjectFromFields(fields, &pageType);

		if (outInfo != nullptr) {
			outInfo->pageType = pageType;
			outInfo->innerObject = innerObject;
		}
		return innerObject != 0;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}
}

bool TryMapEditorCommandToUiCommand(int command, unsigned int* outUiCommand)
{
	if (outUiCommand != nullptr) {
		*outUiCommand = 0;
	}

	unsigned int uiCommand = 0;
	switch (command) {
	case kEditorCmdSelectAll:
		uiCommand = kEditorUiCmdSelectAll;
		break;
	case kEditorCmdCopy:
		uiCommand = kEditorUiCmdCopy;
		break;
	case kEditorCmdPaste:
		uiCommand = kEditorUiCmdPaste;
		break;
	case kEditorCmdDeleteSelection:
		uiCommand = kEditorUiCmdDelete;
		break;
	default:
		return false;
	}

	if (outUiCommand != nullptr) {
		*outUiCommand = uiCommand;
	}
	return true;
}

bool InvokeEditorCommandDirect(
	std::uintptr_t commandTarget,
	std::uintptr_t moduleBase,
	int command,
	int arg2,
	char* arg3,
	char* arg4,
	bool* outThrew = nullptr,
	int* outReturnValue = nullptr,
	std::string* outInvokeMode = nullptr)
{
	if (outThrew != nullptr) {
		*outThrew = false;
	}
	if (outReturnValue != nullptr) {
		*outReturnValue = 0;
	}
	if (outInvokeMode != nullptr) {
		outInvokeMode->clear();
	}
	if (commandTarget == 0) {
		return false;
	}

	const auto& addrs = GetNativeEditorCommandAddresses(moduleBase);
	if (addrs.ok) {
		const auto fn = reinterpret_cast<FnEditorDispatchCommand>(
			addrs.editorDispatchCommand - kImageBase + moduleBase);
		__try {
			const int result = fn(reinterpret_cast<void*>(commandTarget), command, arg2, arg3, arg4);
			if (outReturnValue != nullptr) {
				*outReturnValue = result;
			}
			if (outInvokeMode != nullptr) {
				*outInvokeMode = "fixed";
			}
			if (result != 0) {
				return true;
			}
		}
		__except (EXCEPTION_EXECUTE_HANDLER) {
			if (outThrew != nullptr) {
				*outThrew = true;
			}
			if (outInvokeMode != nullptr) {
				*outInvokeMode = "fixed_exc";
			}
		}
	}

	__try {
		const auto vtable = *reinterpret_cast<std::uintptr_t*>(commandTarget);
		if (vtable != 0) {
			const auto fn = reinterpret_cast<FnEditorDispatchCommand>(
				*reinterpret_cast<std::uintptr_t*>(
					vtable + kEditorDispatchVtableIndex * sizeof(std::uintptr_t)));
			if (fn != nullptr) {
				const int result = fn(reinterpret_cast<void*>(commandTarget), command, arg2, arg3, arg4);
				if (outReturnValue != nullptr) {
					*outReturnValue = result;
				}
				if (outInvokeMode != nullptr) {
					*outInvokeMode = "vtbl56";
				}
				return result != 0;
			}
		}
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		if (outThrew != nullptr) {
			*outThrew = true;
		}
		if (outInvokeMode != nullptr) {
			*outInvokeMode = "vtbl56_exc";
		}
	}
	return false;
}

bool InvokeEditorCommand(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	int command,
	int arg2,
	char* arg3,
	char* arg4,
	std::string* outTrace = nullptr)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	bool rawThrew = false;
	int rawReturn = 0;
	std::string rawMode;
	if (InvokeEditorCommandDirect(
			editorObject,
			moduleBase,
			command,
			arg2,
			arg3,
			arg4,
			&rawThrew,
			&rawReturn,
			&rawMode)) {
		if (outTrace != nullptr) {
			*outTrace =
				"dispatch_raw_ok"
				"|mode=" + rawMode +
				"|ret=" + std::to_string(rawReturn);
		}
		return true;
	}

	EditorDispatchTargetInfo targetInfo{};
	const bool hasInnerObject = TryResolveInnerEditorObject(editorObject, &targetInfo);
	if (hasInnerObject && targetInfo.innerObject != 0 && targetInfo.innerObject != editorObject) {
		bool innerThrew = false;
		int innerReturn = 0;
		std::string innerMode;
		if (InvokeEditorCommandDirect(
				targetInfo.innerObject,
				moduleBase,
				command,
				arg2,
				arg3,
				arg4,
				&innerThrew,
				&innerReturn,
				&innerMode)) {
			if (outTrace != nullptr) {
				*outTrace =
					"dispatch_inner_ok_after_raw"
					"|page_type=" + std::to_string(targetInfo.pageType) +
					"|raw_mode=" + rawMode +
					"|raw_ret=" + std::to_string(rawReturn) +
					"|raw_exc=" + std::to_string(rawThrew ? 1 : 0) +
					"|inner_mode=" + innerMode +
					"|inner_ret=" + std::to_string(innerReturn);
			}
			return true;
		}
		if (outTrace != nullptr) {
			*outTrace =
				"dispatch_failed"
				"|page_type=" + std::to_string(targetInfo.pageType) +
				"|raw_mode=" + rawMode +
				"|raw_ret=" + std::to_string(rawReturn) +
				"|raw_exc=" + std::to_string(rawThrew ? 1 : 0) +
				"|inner_mode=" + innerMode +
				"|inner_ret=" + std::to_string(innerReturn) +
				"|inner_exc=" + std::to_string(innerThrew ? 1 : 0);
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"dispatch_failed"
			"|raw_mode=" + rawMode +
			"|raw_ret=" + std::to_string(rawReturn) +
			"|raw_exc=" + std::to_string(rawThrew ? 1 : 0) +
			"|inner_unavailable";
	}
	return false;
}

const char* DescribeEditorCommand(int command)
{
	switch (command) {
	case kEditorCmdSelectAll:
		return "select_all";
	case kEditorCmdDeleteSelection:
		return "delete_selection";
	case kEditorCmdPaste:
		return "paste";
	case kEditorCmdCopy:
		return "copy";
	case kEditorCmdInsertRawText:
		return "insert_raw_text";
	default:
		return "unknown";
	}
}

void SettleEditorAfterCommand(DWORD waitMs = 60)
{
	const DWORD deadline = GetTickCount() + waitMs;
	do {
		PumpPendingMessages();
		Sleep(10);
	} while (GetTickCount() < deadline);
	PumpPendingMessages();
}

bool InvokeEditorCommandByRoute(
	std::uintptr_t editorObject,
	int command,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	if (command == kEditorCmdDeleteSelection) {
		if (outTrace != nullptr) {
			*outTrace = "route_delete_unsupported";
		}
		return false;
	}

	unsigned int uiCommand = 0;
	if (!TryMapEditorCommandToUiCommand(command, &uiCommand) || uiCommand == 0 || uiCommand == kEditorUiCmdDelete) {
		if (outTrace != nullptr) {
			*outTrace = "route_command_unmapped";
		}
		return false;
	}

	std::string resolveTrace;
	const HWND editorHwnd = ResolveEditorInputWindow(editorObject, &resolveTrace);
	if (editorHwnd == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = resolveTrace.empty() ? "route_resolve_hwnd_failed" : resolveTrace;
		}
		return false;
	}

	std::string routeTrace;
	if (!InvokeMfcCommandRoute(editorHwnd, uiCommand, &routeTrace)) {
		if (outTrace != nullptr) {
			*outTrace = resolveTrace + "|" + routeTrace;
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			resolveTrace +
			"|ui_id=" + std::to_string(uiCommand) +
			"|" + routeTrace;
	}
	return true;
}

bool InvokeEditorCommandWithFallback(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	int command,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	std::string directTrace;
	if (InvokeEditorCommand(
			editorObject,
			moduleBase,
			command,
			kEditorDispatchDefaultFlags,
			nullptr,
			nullptr,
			&directTrace)) {
		SettleEditorAfterCommand();
		if (outTrace != nullptr) {
			*outTrace = std::string(DescribeEditorCommand(command)) + "|direct|" + directTrace;
		}
		return true;
	}

	std::string routeTrace;
	if (InvokeEditorCommandByRoute(editorObject, command, &routeTrace)) {
		SettleEditorAfterCommand();
		if (outTrace != nullptr) {
			*outTrace =
				std::string(DescribeEditorCommand(command)) +
				"|route|" +
				routeTrace +
				"|direct_fail=" + directTrace;
		}
		return true;
	}

	if (outTrace != nullptr) {
		*outTrace =
			std::string(DescribeEditorCommand(command)) +
			"|failed"
			"|direct=" + directTrace +
			"|route=" + routeTrace;
	}
	return false;
}

bool CaptureCurrentSelectionCustomClipboardPayloadByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	CustomClipboardPayload* outPayload,
	std::string* outText,
	std::string* outTrace)
{
	if (outPayload != nullptr) {
		outPayload->Reset();
	}
	if (outText != nullptr) {
		outText->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "fake_clipboard_install_failed";
		}
		return false;
	}

	std::string copyTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdCopy,
			&copyTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "editor_copy_failed|" + copyTrace;
		}
		return false;
	}

	std::string clipboardTrace;
	if (!WaitForFakeClipboardText(fakeClipboard, outText, &clipboardTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "copy_text_format_missing|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	UINT customFormat = 0;
	HANDLE customHandle = nullptr;
	if (!fakeClipboard.GetFirstCustomFormat(&customFormat, &customHandle) ||
		customFormat == 0 ||
		customHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "copy_custom_format_missing|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	HANDLE duplicatedHandle = nullptr;
	if (!DuplicateGlobalHandle(customHandle, &duplicatedHandle) || duplicatedHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "copy_custom_format_duplicate_failed|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	if (outPayload != nullptr) {
		outPayload->Reset();
		outPayload->format = customFormat;
		outPayload->handle = duplicatedHandle;
	}
	else {
		GlobalFree(duplicatedHandle);
	}

	if (outTrace != nullptr) {
		*outTrace =
			"copy_custom_ok|fmt=" + std::to_string(customFormat) +
			"|bytes=" + std::to_string(duplicatedHandle == nullptr ? 0 : static_cast<size_t>(GlobalSize(duplicatedHandle))) +
			"|" + copyTrace +
			"|" + clipboardTrace;
	}
	return true;
}

bool CaptureCurrentSelectionClipboardPayloadAndDeserializeByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	CustomClipboardPayload* outPayload,
	std::string* outText,
	InternalClipboardObject* outObject,
	std::string* outTrace)
{
	if (outPayload != nullptr) {
		outPayload->Reset();
	}
	if (outText != nullptr) {
		outText->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (outObject == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "capture_deserialize_invalid_argument";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "fake_clipboard_install_failed";
		}
		return false;
	}

	std::string copyTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdCopy,
			&copyTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "editor_copy_failed|" + copyTrace;
		}
		return false;
	}

	std::string clipboardTrace;
	if (!WaitForFakeClipboardText(fakeClipboard, outText, &clipboardTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "copy_text_format_missing|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	if (!outObject->DeserializeFromClipboard()) {
		if (outTrace != nullptr) {
			*outTrace = "deserialize_in_capture_failed|" + copyTrace + "|" + clipboardTrace + "|" + fakeClipboard.BuildStatsText();
		}
		return false;
	}

	UINT customFormat = 0;
	HANDLE customHandle = nullptr;
	if (!fakeClipboard.GetFirstCustomFormat(&customFormat, &customHandle) ||
		customFormat == 0 ||
		customHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "copy_custom_format_missing|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	HANDLE duplicatedHandle = nullptr;
	if (!DuplicateGlobalHandle(customHandle, &duplicatedHandle) || duplicatedHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "copy_custom_format_duplicate_failed|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	if (outPayload != nullptr) {
		outPayload->format = customFormat;
		outPayload->handle = duplicatedHandle;
	}
	else {
		GlobalFree(duplicatedHandle);
	}

	if (outTrace != nullptr) {
		*outTrace =
			"copy_and_deserialize_ok|fmt=" + std::to_string(customFormat) +
			"|bytes=" + std::to_string(duplicatedHandle == nullptr ? 0 : static_cast<size_t>(GlobalSize(duplicatedHandle))) +
			"|" + copyTrace +
			"|" + clipboardTrace +
			"|" + fakeClipboard.BuildStatsText();
	}
	return true;
}

bool DeserializeClipboardPayloadToObject(
	const CustomClipboardPayload& payload,
	std::uintptr_t moduleBase,
	InternalClipboardObject* outObject,
	std::string* outTrace)
{
	(void)moduleBase;
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (!payload.IsValid() || outObject == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "deserialize_invalid_argument";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "fake_clipboard_install_failed";
		}
		return false;
	}

	HANDLE duplicatedHandle = payload.DuplicateHandle();
	if (duplicatedHandle == nullptr || !fakeClipboard.SetFormatHandle(payload.format, duplicatedHandle)) {
		if (duplicatedHandle != nullptr) {
			GlobalFree(duplicatedHandle);
		}
		if (outTrace != nullptr) {
			*outTrace = "deserialize_payload_duplicate_failed";
		}
		return false;
	}

	if (!outObject->DeserializeFromClipboard()) {
		if (outTrace != nullptr) {
			*outTrace = "deserialize_from_clipboard_failed|" + fakeClipboard.BuildStatsText();
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"deserialize_ok|fmt=" + std::to_string(payload.format) +
			"|bytes=" + std::to_string(payload.ByteSize()) +
			"|" + fakeClipboard.BuildStatsText();
	}
	return true;
}

bool BuildCustomClipboardPayloadFromPageText(
	std::uintptr_t moduleBase,
	UINT customFormat,
	const std::string& newPageCode,
	const CustomClipboardPayload* templatePayload,
	const InternalClipboardObject* templateObject,
	CustomClipboardPayload* outPayload,
	std::string* outTrace)
{
	if (outPayload != nullptr) {
		outPayload->Reset();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (outPayload == nullptr || customFormat == 0 || newPageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "build_payload_invalid_argument";
		}
		return false;
	}

	InternalClipboardObject replacementObject(moduleBase);
	if (!replacementObject.ParseText(newPageCode)) {
		if (outTrace != nullptr) {
			*outTrace = "parse_text_to_object_failed";
		}
		return false;
	}

	if (templateObject != nullptr) {
		replacementObject.CopySerializationMetadataFrom(*templateObject);
	}

	const std::vector<std::string> supportLibraries = ParseSupportLibraryNamesFromPageCode(newPageCode);
	replacementObject.SetSupportLibraries(supportLibraries);
	std::string finalizeTrace;
	if (!FinalizeParsedClipboardObject(
			moduleBase,
			&replacementObject,
			supportLibraries,
			&finalizeTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "finalize_parsed_object_failed|" + finalizeTrace;
		}
		return false;
	}

	const ClipboardObjectValidationStats replacementStats =
		GatherClipboardObjectValidationStats(moduleBase, replacementObject.RawObject());
	const std::string replacementStatsTrace =
		BuildClipboardObjectValidationTrace(replacementStats);

	HANDLE serializedHandle = replacementObject.SerializeToClipboardHandle();
	if (serializedHandle != nullptr) {
		outPayload->Reset();
		outPayload->format = customFormat;
		outPayload->handle = serializedHandle;
		if (outTrace != nullptr) {
			*outTrace =
				"build_payload_ok|support_libs=" + std::to_string(supportLibraries.size()) +
				"|template=" + std::to_string(templateObject != nullptr ? 1 : 0) +
				"|" + finalizeTrace +
				"|" + replacementStatsTrace +
				"|bytes=" + std::to_string(outPayload->ByteSize());
		}
		return true;
	}

	const std::string primaryFailureTrace =
		"serialize_object_failed|" + finalizeTrace + "|" + replacementStatsTrace;
	if (templatePayload == nullptr || !templatePayload->IsValid() || templateObject == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = primaryFailureTrace;
		}
		return false;
	}

	InternalClipboardObject hybridObject(moduleBase);
	std::string deserializeTemplateTrace;
	if (!DeserializeClipboardPayloadToObject(
			*templatePayload,
			moduleBase,
			&hybridObject,
			&deserializeTemplateTrace)) {
		if (outTrace != nullptr) {
			*outTrace =
				primaryFailureTrace +
				"|hybrid_deserialize_failed|" +
				deserializeTemplateTrace;
		}
		return false;
	}

	std::string transplantTrace;
	if (!TransplantParsedTextOntoTemplateObject(
			moduleBase,
			&hybridObject,
			replacementObject,
			supportLibraries,
			&transplantTrace)) {
		if (outTrace != nullptr) {
			*outTrace =
				primaryFailureTrace +
				"|hybrid_deserialize_ok|" +
				deserializeTemplateTrace +
				"|hybrid_transplant_failed|" +
				transplantTrace;
		}
		return false;
	}

	const ClipboardObjectValidationStats hybridStats =
		GatherClipboardObjectValidationStats(moduleBase, hybridObject.RawObject());
	const std::string hybridStatsTrace =
		BuildClipboardObjectValidationTrace(hybridStats);
	serializedHandle = hybridObject.SerializeToClipboardHandle();
	if (serializedHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				primaryFailureTrace +
				"|hybrid_deserialize_ok|" +
				deserializeTemplateTrace +
				"|" +
				transplantTrace +
				"|hybrid_serialize_failed|" +
				hybridStatsTrace;
		}
		return false;
	}

	outPayload->Reset();
	outPayload->format = customFormat;
	outPayload->handle = serializedHandle;
	if (outTrace != nullptr) {
		*outTrace =
			"build_payload_hybrid_ok|support_libs=" + std::to_string(supportLibraries.size()) +
			"|template=1|" +
			primaryFailureTrace +
			"|hybrid_deserialize_ok|" +
			deserializeTemplateTrace +
			"|" +
			transplantTrace +
			"|" +
			hybridStatsTrace +
			"|bytes=" + std::to_string(outPayload->ByteSize());
	}
	return true;
}

bool PasteParsedTextObjectByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& pageCode,
	const InternalClipboardObject* templateObject,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || pageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "paste_parsed_invalid_argument";
		}
		return false;
	}

	InternalClipboardObject parsedObject(moduleBase);
	if (!parsedObject.ParseText(pageCode)) {
		if (outTrace != nullptr) {
			*outTrace = "parse_text_to_object_failed";
		}
		return false;
	}

	const std::vector<std::string> supportLibraries = ParseSupportLibraryNamesFromPageCode(pageCode);
	if (templateObject != nullptr) {
		parsedObject.CopyAssemblyPageMetadataFrom(*templateObject);
	}
	parsedObject.SetSupportLibraries(supportLibraries);

	std::string finalizeTrace;
	if (!FinalizeParsedClipboardObject(
			moduleBase,
			&parsedObject,
			supportLibraries,
			&finalizeTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "finalize_parsed_object_failed|" + finalizeTrace;
		}
		return false;
	}

	const auto insertFn = ResolveInternalAddress<FnEditorInsertClipboardObject>(
		moduleBase,
		kKnownClipboardInsertObjectRva);
	if (insertFn == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "insert_object_resolve_failed";
		}
		return false;
	}

	if (!CallEditorInsertClipboardObjectSafe(
			insertFn,
			reinterpret_cast<void*>(editorObject),
			parsedObject.RawObject(),
			0,
			0,
			0)) {
		if (outTrace != nullptr) {
			*outTrace = "insert_object_failed|" + finalizeTrace;
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"paste_parsed_ok|support_libs=" + std::to_string(supportLibraries.size()) +
			"|" +
			finalizeTrace;
	}
	return true;
}

bool PastePlainTextByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& pageCode,
	std::string* outTrace)
{
	AppendPageEditTraceLine(
		"PastePlainTextByEditor.begin|editor=" +
		std::to_string(editorObject) +
		"|bytes=" +
		std::to_string(pageCode.size()));
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || pageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "paste_text_invalid_argument";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "fake_clipboard_install_failed";
		}
		return false;
	}

	HANDLE textHandle = CreateClipboardAnsiTextHandle(pageCode);
	if (textHandle == nullptr || !fakeClipboard.SetFormatHandle(CF_TEXT, textHandle)) {
		if (textHandle != nullptr) {
			GlobalFree(textHandle);
		}
		if (outTrace != nullptr) {
			*outTrace = "paste_text_handle_failed";
		}
		return false;
	}
	AppendPageEditTraceLine(
		"PastePlainTextByEditor.clipboard_ready|" +
		fakeClipboard.BuildStatsText());

	const auto& addrs = GetNativeEditorCommandAddresses(moduleBase);
	const auto pasteHandlerFn = ResolveInternalAddress<FnEditorPasteHandler>(
		moduleBase,
		addrs.editorPasteHandler);
	if (pasteHandlerFn == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				"paste_handler_resolve_failed|module_base=" + std::to_string(moduleBase) +
				"|normalized=" + std::to_string(addrs.editorPasteHandler);
		}
		return false;
	}

	std::string pasteTrace;
	AppendPageEditTraceLine("PastePlainTextByEditor.before_paste_command");
	if (!CallEditorPasteHandlerSafe(
			pasteHandlerFn,
			reinterpret_cast<void*>(editorObject),
			0)) {
		AppendPageEditTraceLine(
			"PastePlainTextByEditor.paste_handler_failed|addr=" +
			std::to_string(reinterpret_cast<std::uintptr_t>(pasteHandlerFn)) +
			"|" +
			fakeClipboard.BuildStatsText());
		if (outTrace != nullptr) {
			*outTrace =
				"paste_handler_failed|addr=" +
				std::to_string(reinterpret_cast<std::uintptr_t>(pasteHandlerFn)) +
				"|" + fakeClipboard.BuildStatsText();
		}
		return false;
	}
	pasteTrace =
		"paste_handler_ok|addr=" +
		std::to_string(reinterpret_cast<std::uintptr_t>(pasteHandlerFn));
	AppendPageEditTraceLine(
		"PastePlainTextByEditor.after_paste_command|" +
		pasteTrace +
		"|" +
		fakeClipboard.BuildStatsText());
	SettleEditorAfterCommand();
	AppendPageEditTraceLine("PastePlainTextByEditor.after_settle");

	if (outTrace != nullptr) {
		*outTrace =
			"paste_text_ok|bytes=" + std::to_string(pageCode.size()) +
			"|" +
			pasteTrace +
			"|" + fakeClipboard.BuildStatsText();
	}
	return true;
}

bool PasteCustomClipboardPayloadByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const CustomClipboardPayload& payload,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (!payload.IsValid()) {
		if (outTrace != nullptr) {
			*outTrace = "paste_payload_invalid";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "fake_clipboard_install_failed";
		}
		return false;
	}

	HANDLE duplicatedHandle = payload.DuplicateHandle();
	if (duplicatedHandle == nullptr || !fakeClipboard.SetFormatHandle(payload.format, duplicatedHandle)) {
		if (duplicatedHandle != nullptr) {
			GlobalFree(duplicatedHandle);
		}
		if (outTrace != nullptr) {
			*outTrace = "paste_payload_duplicate_failed";
		}
		return false;
	}

	std::string pasteTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdPaste,
			&pasteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "editor_paste_failed|" + pasteTrace + "|" + fakeClipboard.BuildStatsText();
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"paste_custom_ok|fmt=" + std::to_string(payload.format) +
			"|bytes=" + std::to_string(payload.ByteSize()) +
			"|" + pasteTrace +
			"|" + fakeClipboard.BuildStatsText();
	}
	return true;
}

bool CopyCurrentSelectionByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	std::string* outCode,
	NativeRealPageAccessResult* outResult)
{
	if (outCode != nullptr) {
		outCode->clear();
	}
	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outResult != nullptr) {
			outResult->trace = "fake_clipboard_install_failed";
		}
		return false;
	}

	std::string copyTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdCopy,
			&copyTrace)) {
		if (outResult != nullptr) {
			outResult->trace = "editor_copy_failed|" + copyTrace;
		}
		return false;
	}

	std::string clipboardTrace;
	if (!WaitForFakeClipboardText(fakeClipboard, outCode, &clipboardTrace) ||
		outCode == nullptr ||
		outCode->empty()) {
		if (outResult != nullptr) {
			outResult->trace = "copy_text_format_missing|" + copyTrace + "|" + clipboardTrace;
		}
		return false;
	}

	if (outResult != nullptr) {
		outResult->usedClipboardEmulation = true;
		outResult->capturedCustomFormat = fakeClipboard.HasCustomFormat();
		outResult->textBytes = outCode->size();
		outResult->trace = "copy_ok|" + copyTrace + "|" + clipboardTrace;
	}
	return true;
}

bool TryFormatWholePageTextDirectByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	std::string* outCode,
	std::string* outTrace)
{
	if (outCode != nullptr) {
		outCode->clear();
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || moduleBase == 0 || outCode == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "invalid_argument";
		}
		return false;
	}

	const auto getRangeCountFn = ResolveInternalAddress<FnThiscallInt>(
		moduleBase,
		kKnownEditorGetRangeCountRva);
	const auto formatRangeTextFn = ResolveInternalAddress<FnEditorFormatRangeText>(
		moduleBase,
		kKnownEditorFormatRangeTextRva);
	if (getRangeCountFn == nullptr || formatRangeTextFn == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				"direct_formatter_resolve_failed|count=" +
				std::to_string(reinterpret_cast<std::uintptr_t>(getRangeCountFn)) +
				"|format=" +
				std::to_string(reinterpret_cast<std::uintptr_t>(formatRangeTextFn));
		}
		return false;
	}

	const int rangeCount = CallThiscallIntSafe(getRangeCountFn, reinterpret_cast<void*>(editorObject));
	if (rangeCount <= 0) {
		if (outTrace != nullptr) {
			*outTrace = "direct_formatter_range_count_invalid|count=" + std::to_string(rangeCount);
		}
		return false;
	}

	InternalGenericArrayBuffer buffer(moduleBase);
	if (!CallEditorFormatRangeTextSafe(
			formatRangeTextFn,
			reinterpret_cast<void*>(editorObject),
			0,
			rangeCount - 1,
			buffer.Data(),
			0)) {
		if (outTrace != nullptr) {
			*outTrace =
				"direct_formatter_call_failed|count=" +
				std::to_string(rangeCount) +
				"|editor=" + std::to_string(editorObject);
		}
		return false;
	}

	if (!buffer.CopyTextTo(outCode)) {
		if (outTrace != nullptr) {
			*outTrace =
				"direct_formatter_empty|count=" +
				std::to_string(rangeCount) +
				"|buffer_bytes=" + std::to_string(buffer.ByteSize());
		}
		return false;
	}

	bool rebuiltHeader = false;
	if (!outCode->empty()) {
		const std::string trimmed = TrimAsciiCopyLocal(*outCode);
		if (trimmed.rfind(".版本", 0) != 0) {
			*outCode = BuildDirectWholePageHeaderText() + *outCode;
			rebuiltHeader = true;
		}
	}

	if (outTrace != nullptr) {
		*outTrace =
			"direct_formatter_ok|count=" +
			std::to_string(rangeCount) +
			"|buffer_bytes=" + std::to_string(buffer.ByteSize()) +
			"|text_bytes=" + std::to_string(outCode->size()) +
			"|rebuilt_header=" + std::to_string(rebuiltHeader ? 1 : 0) +
			"|support_lib_omitted=1";
	}
	return true;
}

bool CopyWholePageTextByEditor(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	std::string* outCode,
	NativeRealPageAccessResult* outResult)
{
	if (outResult != nullptr) {
		*outResult = {};
		outResult->editorObject = editorObject;
	}
	if (outCode != nullptr) {
		outCode->clear();
	}
	if (editorObject == 0 || outCode == nullptr) {
		if (outResult != nullptr) {
			outResult->trace = "invalid_argument";
		}
		return false;
	}

	std::string selectTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdSelectAll,
			&selectTrace)) {
		if (outResult != nullptr) {
			outResult->trace = "select_all_failed|" + selectTrace;
		}
		return false;
	}

	NativeRealPageAccessResult copyResult{};
	if (!CopyCurrentSelectionByEditor(editorObject, moduleBase, outCode, &copyResult)) {
		if (outResult != nullptr) {
			*outResult = copyResult;
			outResult->editorObject = editorObject;
			outResult->trace = selectTrace + "|" + copyResult.trace;
		}
		return false;
	}

	if (outResult != nullptr) {
		*outResult = copyResult;
		outResult->editorObject = editorObject;
		outResult->ok = true;
		outResult->trace = selectTrace + "|" + copyResult.trace;
	}

	if (outCode != nullptr && !outCode->empty()) {
		std::string directCode;
		std::string directTrace;
		if (TryFormatWholePageTextDirectByEditor(editorObject, moduleBase, &directCode, &directTrace)) {
			std::string matchMode;
			std::string matchSummary;
			bool matched = false;
			if (NormalizeLineBreakToCrLf(*outCode) == NormalizeLineBreakToCrLf(directCode)) {
				matchMode = "exact";
				matched = true;
			}
			else {
				matched = VerifyRealPageCodeMatches(*outCode, directCode, &matchMode, &matchSummary);
			}
			if (outResult != nullptr) {
				outResult->trace +=
					"|direct_compare_" +
					std::string(matched ? "match" : "mismatch") +
					"|mode=" + (matchMode.empty() ? std::string("none") : matchMode) +
					"|" + directTrace;
				if (!matchSummary.empty()) {
					outResult->trace += "|" + matchSummary;
				}
			}
		}
		else if (outResult != nullptr) {
			outResult->trace += "|direct_compare_failed|" + directTrace;
		}
	}
	return true;
}

bool ReplaceWholePageByCustomClipboardPayload(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const CustomClipboardPayload& payload,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || !payload.IsValid()) {
		if (outTrace != nullptr) {
			*outTrace = "invalid_argument";
		}
		return false;
	}

	std::string selectTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdSelectAll,
			&selectTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "select_all_failed|" + selectTrace;
		}
		return false;
	}

	std::string deleteTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdDeleteSelection,
			&deleteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|delete_failed|" + deleteTrace;
		}
		return false;
	}

	std::string pasteTrace;
	if (!PasteCustomClipboardPayloadByEditor(editorObject, moduleBase, payload, &pasteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|" + deleteTrace + "|" + pasteTrace;
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace = selectTrace + "|" + deleteTrace + "|" + pasteTrace;
	}
	return true;
}

bool ReplaceWholePageByParsedTextObject(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& pageCode,
	const InternalClipboardObject* templateObject,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || pageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "invalid_argument";
		}
		return false;
	}

	std::string selectTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdSelectAll,
			&selectTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "select_all_failed|" + selectTrace;
		}
		return false;
	}

	std::string deleteTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdDeleteSelection,
			&deleteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|delete_failed|" + deleteTrace;
		}
		return false;
	}

	std::string pasteTrace;
	if (!PasteParsedTextObjectByEditor(editorObject, moduleBase, pageCode, templateObject, &pasteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|" + deleteTrace + "|" + pasteTrace;
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace = selectTrace + "|" + deleteTrace + "|" + pasteTrace;
	}
	return true;
}

bool ReplaceWholePageByTextPaste(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& pageCode,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || pageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "invalid_argument";
		}
		return false;
	}

	std::string selectTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdSelectAll,
			&selectTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "select_all_failed|" + selectTrace;
		}
		return false;
	}

	std::string deleteTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdDeleteSelection,
			&deleteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|delete_failed|" + deleteTrace;
		}
		return false;
	}

	std::string pasteTrace;
	if (!PastePlainTextByEditor(editorObject, moduleBase, pageCode, &pasteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|" + deleteTrace + "|" + pasteTrace;
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace = selectTrace + "|" + deleteTrace + "|" + pasteTrace;
	}
	return true;
}

bool ReplaceWholePageByRawInsertCommand(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& pageCode,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (editorObject == 0 || pageCode.empty()) {
		if (outTrace != nullptr) {
			*outTrace = "invalid_argument";
		}
		return false;
	}

	std::string selectTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdSelectAll,
			&selectTrace)) {
		if (outTrace != nullptr) {
			*outTrace = "select_all_failed|" + selectTrace;
		}
		return false;
	}

	std::string deleteTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdDeleteSelection,
			&deleteTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|delete_failed|" + deleteTrace;
		}
		return false;
	}

	std::string insertTrace;
	if (!InvokeEditorCommand(
			editorObject,
			moduleBase,
			kEditorCmdInsertRawText,
			1,
			const_cast<char*>(pageCode.c_str()),
			reinterpret_cast<char*>(1),
			&insertTrace)) {
		if (outTrace != nullptr) {
			*outTrace = selectTrace + "|" + deleteTrace + "|insert_failed|" + insertTrace;
		}
		return false;
	}
	SettleEditorAfterCommand();

	if (outTrace != nullptr) {
		*outTrace = selectTrace + "|" + deleteTrace + "|" + insertTrace;
	}
	return true;
}

bool TryRollbackRealPageCode(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const CustomClipboardPayload* rollbackPayload,
	const std::string* expectedPageCode,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	const bool hasExpectedPageCode =
		expectedPageCode != nullptr &&
		!expectedPageCode->empty();
	if (!hasExpectedPageCode &&
		(rollbackPayload == nullptr || !rollbackPayload->IsValid())) {
		if (outTrace != nullptr) {
			*outTrace = "rollback_skipped";
		}
		return false;
	}

	std::string rollbackWriteTrace;
	if (hasExpectedPageCode) {
		if (!ReplaceWholePageByTextPaste(
				editorObject,
				moduleBase,
				*expectedPageCode,
				&rollbackWriteTrace)) {
			if (outTrace != nullptr) {
				*outTrace = "rollback_text_paste_failed|" + rollbackWriteTrace;
			}
			return false;
		}
		rollbackWriteTrace = "rollback_text_paste|" + rollbackWriteTrace;
	}
	else {
		if (!ReplaceWholePageByCustomClipboardPayload(editorObject, moduleBase, *rollbackPayload, &rollbackWriteTrace)) {
			if (outTrace != nullptr) {
				*outTrace = "rollback_write_failed|" + rollbackWriteTrace;
			}
			return false;
		}
	}

	std::string verifyCode;
	NativeRealPageAccessResult verifyResult{};
	if (!CopyWholePageTextByEditor(editorObject, moduleBase, &verifyCode, &verifyResult)) {
		if (outTrace != nullptr) {
			*outTrace = rollbackWriteTrace + "|rollback_verify_read_failed|" + verifyResult.trace;
		}
		return false;
	}

	if (expectedPageCode != nullptr &&
		!expectedPageCode->empty()) {
		std::string verifyMode;
		std::string verifySummary;
		if (!VerifyRealPageCodeMatches(*expectedPageCode, verifyCode, &verifyMode, &verifySummary)) {
			if (outTrace != nullptr) {
				*outTrace = rollbackWriteTrace + "|rollback_verify_mismatch|" + verifySummary;
			}
			return false;
		}
		if (outTrace != nullptr) {
			*outTrace =
				"rollback_ok|" +
				rollbackWriteTrace +
				"|rollback_verify_" +
				verifyMode +
				"|" +
				verifyResult.trace;
		}
		return true;
	}

	if (outTrace != nullptr) {
		*outTrace = "rollback_ok|" + rollbackWriteTrace + "|" + verifyResult.trace;
	}
	return true;
}

bool ReplaceRealPageCodeByEditorObjectInternal(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& newPageCode,
	const std::string* rollbackPageCode,
	NativeRealPageAccessResult* outResult)
{
	(void)rollbackPageCode;
	AppendPageEditTraceLine(
		"ReplaceRealPageCode.begin|editor=" +
		std::to_string(editorObject) +
		"|bytes=" +
		std::to_string(newPageCode.size()));
	if (outResult != nullptr) {
		*outResult = {};
		outResult->editorObject = editorObject;
	}
	if (editorObject == 0 || newPageCode.empty()) {
		if (outResult != nullptr) {
			outResult->trace = "invalid_argument";
		}
		return false;
	}

	const std::string normalizedNewPageCode = NormalizeLineBreakToCrLf(newPageCode);
	std::string selectTrace;
	if (!InvokeEditorCommandWithFallback(
			editorObject,
			moduleBase,
			kEditorCmdSelectAll,
			&selectTrace)) {
		AppendPageEditTraceLine("ReplaceRealPageCode.select_all_failed|" + selectTrace);
		if (outResult != nullptr) {
			outResult->trace = "select_all_failed|" + selectTrace;
		}
		return false;
	}
	AppendPageEditTraceLine("ReplaceRealPageCode.after_select_all|" + selectTrace);

	std::string replaceTrace;
	const std::string writeStrategyTrace = "write_by_real_paste_handler_text";
	AppendPageEditTraceLine("ReplaceRealPageCode.before_replace|" + writeStrategyTrace);
	const bool replaceOk = ReplaceWholePageByTextPaste(
		editorObject,
		moduleBase,
		normalizedNewPageCode,
		&replaceTrace);
	AppendPageEditTraceLine(
		std::string("ReplaceRealPageCode.after_replace|ok=") +
		(replaceOk ? "1" : "0") +
		"|" +
		replaceTrace);

	if (!replaceOk) {
		if (outResult != nullptr) {
			outResult->trace =
				selectTrace +
				"|replace_failed|" +
				replaceTrace;
		}
		return false;
	}

	std::string verifyCode;
	NativeRealPageAccessResult verifyResult{};
	AppendPageEditTraceLine("ReplaceRealPageCode.before_verify_read");
	if (!CopyWholePageTextByEditor(editorObject, moduleBase, &verifyCode, &verifyResult)) {
		AppendPageEditTraceLine("ReplaceRealPageCode.verify_read_failed|" + verifyResult.trace);
		if (outResult != nullptr) {
			outResult->trace =
				selectTrace +
				"|" +
				writeStrategyTrace +
				"|" +
				replaceTrace +
				"|verify_read_failed|" +
				verifyResult.trace;
		}
		return false;
	}
	AppendPageEditTraceLine(
		"ReplaceRealPageCode.after_verify_read|bytes=" +
		std::to_string(verifyCode.size()) +
		"|" +
		verifyResult.trace);

	std::string verifyMode;
	std::string verifySummary;
	if (!VerifyRealPageCodeMatches(normalizedNewPageCode, verifyCode, &verifyMode, &verifySummary)) {
		AppendPageEditTraceLine("ReplaceRealPageCode.verify_mismatch|" + verifySummary);
		if (outResult != nullptr) {
			outResult->trace =
				selectTrace +
				"|" +
				writeStrategyTrace +
				"|" +
				replaceTrace +
				"|verify_mismatch|" +
				verifySummary;
		}
		return false;
	}
	if (verifyMode == "structural") {
		AppendPageEditTraceLine("ReplaceRealPageCode.verify_structural_ok|" + verifySummary);
	}
	AppendPageEditTraceLine("ReplaceRealPageCode.success");

	if (outResult != nullptr) {
		outResult->ok = true;
		outResult->usedClipboardEmulation = true;
		outResult->textBytes = verifyCode.size();
		outResult->trace =
			selectTrace +
			"|" +
			writeStrategyTrace +
			"|" +
			replaceTrace +
			"|verify_ok_" +
			verifyMode +
			"|" +
			verifyResult.trace;
	}
	return true;
}

bool ReplaceRealPageCodeByEditorWindowInternal(
	HWND editorHwnd,
	const std::string& newPageCode,
	NativeRealPageAccessResult* outResult)
{
	AppendPageEditTraceLine(
		"ReplaceRealPageCodeByWindow.begin|hwnd=" +
		std::to_string(reinterpret_cast<std::uintptr_t>(editorHwnd)) +
		"|bytes=" +
		std::to_string(newPageCode.size()));
	if (outResult != nullptr) {
		*outResult = {};
	}
	if (editorHwnd == nullptr || !IsWindow(editorHwnd) || newPageCode.empty()) {
		if (outResult != nullptr) {
			outResult->trace = "invalid_argument";
		}
		return false;
	}

	const std::string normalizedNewPageCode = NormalizeLineBreakToCrLf(newPageCode);
	std::string replaceTrace;
	if (!ReplaceWholePageTextByEditorWindowInput(editorHwnd, normalizedNewPageCode, &replaceTrace)) {
		AppendPageEditTraceLine("ReplaceRealPageCodeByWindow.replace_failed|" + replaceTrace);
		if (outResult != nullptr) {
			outResult->trace = "replace_failed|" + replaceTrace;
		}
		return false;
	}
	AppendPageEditTraceLine("ReplaceRealPageCodeByWindow.after_replace|" + replaceTrace);

	std::string verifyCode;
	NativeRealPageAccessResult verifyResult{};
	if (!CopyWholePageTextByEditorWindowInput(editorHwnd, &verifyCode, &verifyResult)) {
		AppendPageEditTraceLine("ReplaceRealPageCodeByWindow.verify_read_failed|" + verifyResult.trace);
		if (outResult != nullptr) {
			outResult->trace =
				replaceTrace +
				"|verify_read_failed|" +
				verifyResult.trace;
		}
		return false;
	}
	AppendPageEditTraceLine(
		"ReplaceRealPageCodeByWindow.after_verify_read|bytes=" +
		std::to_string(verifyCode.size()) +
		"|" +
		verifyResult.trace);

	std::string verifyMode;
	std::string verifySummary;
	if (!VerifyRealPageCodeMatches(normalizedNewPageCode, verifyCode, &verifyMode, &verifySummary)) {
		AppendPageEditTraceLine("ReplaceRealPageCodeByWindow.verify_mismatch|" + verifySummary);
		if (outResult != nullptr) {
			outResult->trace =
				replaceTrace +
				"|verify_mismatch|" +
				verifySummary;
		}
		return false;
	}
	if (verifyMode == "structural") {
		AppendPageEditTraceLine("ReplaceRealPageCodeByWindow.verify_structural_ok|" + verifySummary);
	}

	if (outResult != nullptr) {
		*outResult = verifyResult;
		outResult->ok = true;
		outResult->trace =
			"write_by_window_input|" +
			replaceTrace +
			"|verify_ok_" +
			verifyMode +
			"|" +
			verifyResult.trace;
	}
	return true;
}

bool ResolveEditorObjectByProgramTreeItemDataInternal(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	std::string* outTrace)
{
	if (outEditorObject != nullptr) {
		*outEditorObject = 0;
	}
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	std::uintptr_t editorObject = 0;
	std::string uiTrace;
	if (TryResolveProgramTreeItemEditorObjectByUi(itemData, moduleBase, &editorObject, &uiTrace) &&
		editorObject != 0) {
		if (outEditorObject != nullptr) {
			*outEditorObject = editorObject;
		}
		if (outTrace != nullptr) {
			*outTrace = uiTrace;
		}
		return true;
	}

	std::string debugTrace;
	if (!DebugResolveEditorObjectByProgramTreeItemData(
			itemData,
			moduleBase,
			&editorObject,
			nullptr,
			nullptr,
			nullptr,
			&debugTrace) ||
		editorObject == 0) {
		if (outTrace != nullptr) {
			if (!uiTrace.empty() && !debugTrace.empty()) {
				*outTrace = uiTrace + "|fallback_debug_failed|" + debugTrace;
			}
			else {
				*outTrace = !uiTrace.empty() ? uiTrace : debugTrace;
			}
		}
		return false;
	}

	if (outEditorObject != nullptr) {
		*outEditorObject = editorObject;
	}
	if (outTrace != nullptr) {
		*outTrace =
			(!uiTrace.empty() ? (uiTrace + "|fallback_debug_ok|") : "fallback_debug_ok|") +
			debugTrace;
	}
	return true;
}

bool TryActivateProgramTreeItemPageByEditorObjectFallback(
	unsigned int itemData,
	const std::string& itemText,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	const std::uintptr_t moduleBase =
		reinterpret_cast<std::uintptr_t>(GetModuleHandleW(nullptr));
	if (moduleBase == 0) {
		if (outTrace != nullptr) {
			*outTrace = "fallback_activate_module_base_invalid";
		}
		return false;
	}

	std::uintptr_t editorObject = 0;
	int resolvedType = 0;
	int resolvedIndex = -1;
	int bucketData = 0;
	std::string resolveTrace;
	if (!DebugResolveEditorObjectByProgramTreeItemData(
			itemData,
			moduleBase,
			&editorObject,
			&resolvedType,
			&resolvedIndex,
			&bucketData,
			&resolveTrace) ||
		editorObject == 0) {
		if (outTrace != nullptr) {
			*outTrace = "fallback_activate_resolve_failed|" + resolveTrace;
		}
		return false;
	}

	std::string editorHwndTrace;
	HWND editorHwnd = ResolveEditorInputWindow(editorObject, &editorHwndTrace);
	if ((editorHwnd == nullptr || !IsWindow(editorHwnd)) &&
		!TryReadEditorWindowHandleFromObject(editorObject, &editorHwnd)) {
		if (outTrace != nullptr) {
			*outTrace =
				"fallback_activate_editor_hwnd_failed|" +
				resolveTrace +
				"|" +
				editorHwndTrace;
		}
		return false;
	}

	const HWND mdiChildHwnd = FindMdiChildWindow(editorHwnd);
	if (mdiChildHwnd == nullptr || !IsWindow(mdiChildHwnd)) {
		if (outTrace != nullptr) {
			*outTrace =
				"fallback_activate_mdi_child_failed|" +
				resolveTrace +
				"|" +
				editorHwndTrace +
				"|editor_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(editorHwnd));
		}
		return false;
	}

	const HWND mdiClientHwnd = GetParent(mdiChildHwnd);
	const HWND rootHwnd = GetAncestor(mdiChildHwnd, GA_ROOT);
	if (rootHwnd != nullptr && IsWindow(rootHwnd)) {
		if (IsIconic(rootHwnd)) {
			ShowWindow(rootHwnd, SW_RESTORE);
		}
		BringWindowToTop(rootHwnd);
		SetForegroundWindow(rootHwnd);
		SetActiveWindow(rootHwnd);
	}
	if (mdiClientHwnd != nullptr && IsWindow(mdiClientHwnd)) {
		SendMessageA(mdiClientHwnd, WM_MDIACTIVATE, reinterpret_cast<WPARAM>(mdiChildHwnd), 0);
	}
	SetFocus(editorHwnd);
	UpdateWindow(editorHwnd);
	SendMessageA(editorHwnd, WM_CANCELMODE, 0, 0);
	ReleaseCapture();
	PumpPendingMessages();

	const HWND mainHwnd = ResolveMainIdeWindow();
	for (int attempt = 0; attempt < 60; ++attempt) {
		PumpPendingMessages();
		const HWND activeChild = GetActiveMdiChildWindow(mainHwnd);
		if (activeChild != nullptr && IsWindow(activeChild)) {
			const std::string activeTitle = WindowTextToString(activeChild);
			const bool titleMatches =
				!itemText.empty() &&
				activeTitle.find(itemText) != std::string::npos;
			if (activeChild == mdiChildHwnd || titleMatches) {
				if (outTrace != nullptr) {
					*outTrace =
						"fallback_activate_ok"
						"|item=" + itemText +
						"|editor_object=" + std::to_string(editorObject) +
						"|resolved_type=" + std::to_string(resolvedType) +
						"|resolved_index=" + std::to_string(resolvedIndex) +
						"|bucket_data=" + std::to_string(bucketData) +
						"|mdi_child=" + std::to_string(reinterpret_cast<std::uintptr_t>(mdiChildHwnd)) +
						"|active_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(activeChild)) +
						"|active_class=" + WindowClassToString(activeChild) +
						"|active_text=" + activeTitle +
						"|" +
						resolveTrace +
						"|" +
						editorHwndTrace;
				}
				return true;
			}
		}
		Sleep(25);
	}

	if (outTrace != nullptr) {
		*outTrace =
			"fallback_activate_timeout"
			"|item=" + itemText +
			"|editor_object=" + std::to_string(editorObject) +
			"|mdi_child=" + std::to_string(reinterpret_cast<std::uintptr_t>(mdiChildHwnd)) +
			"|" +
			resolveTrace +
			"|" +
			editorHwndTrace;
	}
	return false;
}

}

bool OpenProgramTreeItemPageByData(
	unsigned int itemData,
	std::string* outTrace)
{
	if (outTrace != nullptr) {
		outTrace->clear();
	}

	std::string treeTrace;
	const HWND mainHwnd = ResolveMainIdeWindow();
	const HWND treeHwnd = FindProgramDataTreeViewWindow(&treeTrace);
	if (mainHwnd == nullptr || treeHwnd == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace.empty() ? "program_tree_unavailable" : treeTrace;
		}
		return false;
	}

	const HTREEITEM rootItem = GetTreeNextItemInternal(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM firstChild = GetTreeNextItemInternal(treeHwnd, rootItem, TVGN_CHILD);
	std::string itemText;
	const HTREEITEM item = FindProgramTreeItemByDataRecursive(
		treeHwnd,
		firstChild,
		itemData,
		0,
		8,
		&itemText);
	if (item == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				treeTrace +
				"|tree_item_not_found|item_data=" + std::to_string(itemData);
		}
		return false;
	}

	const HWND activeBefore = GetActiveMdiChildWindow(mainHwnd);
	std::string openTrace;
	if (!TriggerProgramTreeItemOpenByUi(treeHwnd, item, &openTrace)) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace + "|" + openTrace;
		}
		return false;
	}

	for (int attempt = 0; attempt < 60; ++attempt) {
		PumpPendingMessages();
		const HWND activeChild = GetActiveMdiChildWindow(mainHwnd);
		if (activeChild != nullptr && IsWindow(activeChild)) {
			const std::string activeTitle = WindowTextToString(activeChild);
			const bool titleMatches =
				!itemText.empty() &&
				activeTitle.find(itemText) != std::string::npos;
			if (activeChild != activeBefore || titleMatches) {
				if (outTrace != nullptr) {
					*outTrace =
						treeTrace +
						"|" +
						openTrace +
						"|open_page_ok" +
						"|item=" + itemText +
						"|active_changed=" + std::to_string(activeChild != activeBefore ? 1 : 0) +
						"|active_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(activeChild)) +
						"|active_class=" + WindowClassToString(activeChild) +
						"|active_text=" + activeTitle;
				}
				return true;
			}
		}
		Sleep(25);
	}

	std::string fallbackTrace;
	if (TryActivateProgramTreeItemPageByEditorObjectFallback(itemData, itemText, &fallbackTrace)) {
		if (outTrace != nullptr) {
			*outTrace = treeTrace + "|" + openTrace + "|" + fallbackTrace;
		}
		return true;
	}

	if (outTrace != nullptr) {
		*outTrace =
			treeTrace +
			"|" +
			openTrace +
			"|open_page_timeout|item=" + itemText +
			"|" +
			fallbackTrace;
	}
	return false;
}

bool GetRealPageCodeByEditorObject(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	std::string* outCode,
	NativeRealPageAccessResult* outResult)
{
	if (outResult != nullptr) {
		*outResult = {};
		outResult->editorObject = editorObject;
	}
	if (outCode != nullptr) {
		outCode->clear();
	}
	if (editorObject == 0 || outCode == nullptr) {
		if (outResult != nullptr) {
			outResult->trace = "invalid_argument";
		}
		return false;
	}

	if (CopyWholePageTextByEditor(editorObject, moduleBase, outCode, outResult)) {
		if (outResult != nullptr) {
			outResult->ok = true;
		}
		return true;
	}

	return false;
}

bool GetRealPageCodeByProgramTreeItemData(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	std::string* outCode,
	NativeRealPageAccessResult* outResult)
{
	if (outCode != nullptr) {
		outCode->clear();
	}
	if (outResult != nullptr) {
		*outResult = {};
	}

	std::uintptr_t hiddenEditorObject = 0;
	std::uintptr_t hiddenOriginalEditorObject = 0;
	const HWND hiddenMainHwnd = ResolveMainIdeWindow();
	const HWND hiddenOriginalMdiChildHwnd = GetActiveMdiChildWindow(hiddenMainHwnd);
	const std::string hiddenOriginalMdiTitle = WindowTextToString(hiddenOriginalMdiChildHwnd);
	std::string hiddenResolveTrace;
	std::string hiddenAttemptTrace;
	std::string hiddenOriginalTrace;
	std::string hiddenSwitchTrace;
	std::string hiddenRestoreTrace;
	bool hiddenNeedRestore = false;
	const bool hiddenOriginalCaptured =
		DebugGetMainEditorActiveEditorObject(moduleBase, &hiddenOriginalEditorObject, &hiddenOriginalTrace) &&
		hiddenOriginalEditorObject != 0;
	const auto restoreHiddenMdiChild = [&]() {
		if (hiddenMainHwnd == nullptr || !IsWindow(hiddenMainHwnd) ||
			hiddenOriginalMdiChildHwnd == nullptr || !IsWindow(hiddenOriginalMdiChildHwnd)) {
			return;
		}
		const HWND mdiClientHwnd = GetParent(hiddenOriginalMdiChildHwnd);
		if (mdiClientHwnd != nullptr && IsWindow(mdiClientHwnd)) {
			SendMessageA(
				mdiClientHwnd,
				WM_MDIACTIVATE,
				reinterpret_cast<WPARAM>(hiddenOriginalMdiChildHwnd),
				0);
			PumpPendingMessages();

			bool mdiRestoreOk = false;
			for (int attempt = 0; attempt < 20; ++attempt) {
				const HWND activeChild = GetActiveMdiChildWindow(hiddenMainHwnd);
				if (activeChild != nullptr && IsWindow(activeChild)) {
					const std::string activeTitle = WindowTextToString(activeChild);
					if (activeChild == hiddenOriginalMdiChildHwnd ||
						(!hiddenOriginalMdiTitle.empty() &&
						 activeTitle.find(hiddenOriginalMdiTitle) != std::string::npos)) {
						hiddenRestoreTrace +=
							(hiddenRestoreTrace.empty() ? std::string() : std::string("|")) +
							"mdi_restore_ok"
							"|active_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(activeChild)) +
							"|active_text=" + activeTitle;
						mdiRestoreOk = true;
						break;
					}
				}
				Sleep(10);
				PumpPendingMessages();
			}

			if (!mdiRestoreOk) {
				hiddenRestoreTrace +=
					(hiddenRestoreTrace.empty() ? std::string() : std::string("|")) +
					"mdi_restore_timeout"
					"|requested_hwnd=" + std::to_string(reinterpret_cast<std::uintptr_t>(hiddenOriginalMdiChildHwnd)) +
					"|requested_text=" + hiddenOriginalMdiTitle;
			}
		}
		else {
			hiddenRestoreTrace +=
				(hiddenRestoreTrace.empty() ? std::string() : std::string("|")) +
				"mdi_restore_client_missing";
		}
	};
	const auto restoreHiddenActiveEditor = [&]() {
		if (!hiddenNeedRestore || hiddenOriginalEditorObject == 0) {
			restoreHiddenMdiChild();
			return;
		}
		std::string restoreTrace;
		if (DebugSetMainEditorActiveEditorObject(
				moduleBase,
				hiddenOriginalEditorObject,
				1,
				nullptr,
				&restoreTrace)) {
			hiddenRestoreTrace = restoreTrace;
		}
		else {
			hiddenRestoreTrace =
				restoreTrace.empty() ? "restore_active_editor_failed" : restoreTrace;
		}
		restoreHiddenMdiChild();
		hiddenNeedRestore = false;
	};
	if (DebugResolveEditorObjectByProgramTreeItemDataNoActivate(
			itemData,
			moduleBase,
			&hiddenEditorObject,
			nullptr,
			nullptr,
			nullptr,
			&hiddenResolveTrace) &&
		hiddenEditorObject != 0) {
		hiddenAttemptTrace =
			(hiddenOriginalTrace.empty() ? std::string() : ("active_before|" + hiddenOriginalTrace + "|")) +
			(hiddenResolveTrace.empty() ? std::string("no_activate") : hiddenResolveTrace);

		NativeRealPageAccessResult hiddenDirectResult{};
		if (GetRealPageCodeByEditorObject(hiddenEditorObject, moduleBase, outCode, &hiddenDirectResult)) {
			restoreHiddenMdiChild();
			hiddenDirectResult.editorObject = hiddenEditorObject;
			hiddenDirectResult.trace =
				hiddenAttemptTrace +
				(hiddenRestoreTrace.empty() ? std::string() : ("|post_read_restore|" + hiddenRestoreTrace)) +
				"|no_host_swap|" +
				hiddenDirectResult.trace;
			if (outResult != nullptr) {
				*outResult = std::move(hiddenDirectResult);
			}
			return true;
		}
		hiddenAttemptTrace += "|no_host_swap_read_failed|" + hiddenDirectResult.trace;

		std::string directNoSwapCode;
		std::string directNoSwapTrace;
		if (TryFormatWholePageTextDirectByEditor(
				hiddenEditorObject,
				moduleBase,
				&directNoSwapCode,
				&directNoSwapTrace) &&
			!directNoSwapCode.empty()) {
			restoreHiddenMdiChild();
			if (outCode != nullptr) {
				*outCode = std::move(directNoSwapCode);
			}
			if (outResult != nullptr) {
				outResult->ok = true;
				outResult->editorObject = hiddenEditorObject;
				outResult->usedClipboardEmulation = false;
				outResult->capturedCustomFormat = false;
				outResult->textBytes = outCode == nullptr ? 0 : outCode->size();
				outResult->trace =
					hiddenAttemptTrace +
					(hiddenRestoreTrace.empty() ? std::string() : ("|post_read_restore|" + hiddenRestoreTrace)) +
					"|no_host_swap|direct_only|" +
					directNoSwapTrace;
			}
			return true;
		}
		hiddenAttemptTrace += "|no_host_swap_direct_failed|" + directNoSwapTrace;

		hiddenNeedRestore =
			hiddenOriginalCaptured &&
			hiddenOriginalEditorObject != 0 &&
			hiddenOriginalEditorObject != hiddenEditorObject;
		if (DebugSetMainEditorActiveEditorObject(
				moduleBase,
				hiddenEditorObject,
				1,
				nullptr,
				&hiddenSwitchTrace)) {
			hiddenAttemptTrace += "|host_active_swap|" + hiddenSwitchTrace;
		}
		else {
			hiddenAttemptTrace +=
				"|host_active_swap_failed|" +
				(hiddenSwitchTrace.empty() ? std::string("set_active_editor_failed") : hiddenSwitchTrace);
		}

		NativeRealPageAccessResult hiddenResult{};
		if (GetRealPageCodeByEditorObject(hiddenEditorObject, moduleBase, outCode, &hiddenResult)) {
			restoreHiddenActiveEditor();
			hiddenResult.editorObject = hiddenEditorObject;
			hiddenResult.trace =
				hiddenAttemptTrace +
				(hiddenRestoreTrace.empty() ? std::string() : ("|host_active_restore|" + hiddenRestoreTrace)) +
				"|no_activate|" +
				hiddenResult.trace;
			if (outResult != nullptr) {
				*outResult = std::move(hiddenResult);
			}
			return true;
		}
		hiddenAttemptTrace += "|read_failed|" + hiddenResult.trace;

		std::string directCode;
		std::string directTrace;
		if (TryFormatWholePageTextDirectByEditor(
				hiddenEditorObject,
				moduleBase,
				&directCode,
				&directTrace) &&
			!directCode.empty()) {
			restoreHiddenActiveEditor();
			if (outCode != nullptr) {
				*outCode = std::move(directCode);
			}
			if (outResult != nullptr) {
				outResult->ok = true;
				outResult->editorObject = hiddenEditorObject;
				outResult->usedClipboardEmulation = false;
				outResult->capturedCustomFormat = false;
				outResult->textBytes = outCode == nullptr ? 0 : outCode->size();
				outResult->trace =
					hiddenAttemptTrace +
					(hiddenRestoreTrace.empty() ? std::string() : ("|host_active_restore|" + hiddenRestoreTrace)) +
					"|no_activate|direct_only|" +
					directTrace;
			}
			return true;
		}
		hiddenAttemptTrace += "|direct_failed|" + directTrace;
		restoreHiddenActiveEditor();
		if (!hiddenRestoreTrace.empty()) {
			hiddenAttemptTrace += "|host_active_restore|" + hiddenRestoreTrace;
		}
	}
	else if (!hiddenResolveTrace.empty()) {
		hiddenAttemptTrace =
			(hiddenOriginalTrace.empty() ? std::string() : ("active_before|" + hiddenOriginalTrace + "|")) +
			hiddenResolveTrace;
	}

	std::uintptr_t editorObject = 0;
	std::string resolveTrace;
	if (!ResolveEditorObjectByProgramTreeItemDataInternal(
			itemData,
			moduleBase,
			&editorObject,
			&resolveTrace)) {
		if (outResult != nullptr) {
			if (!hiddenAttemptTrace.empty()) {
				outResult->trace =
					"no_activate_attempt|" +
					hiddenAttemptTrace +
					"|fallback_resolve_failed|" +
					(resolveTrace.empty() ? std::string("resolve_editor_failed") : resolveTrace);
			}
			else {
				outResult->trace = resolveTrace.empty() ? "resolve_editor_failed" : resolveTrace;
			}
		}
		return false;
	}

	NativeRealPageAccessResult localResult{};
	const bool ok = GetRealPageCodeByEditorObject(editorObject, moduleBase, outCode, &localResult);
	localResult.editorObject = editorObject;
	if (!hiddenAttemptTrace.empty()) {
		localResult.trace =
			"no_activate_attempt|" +
			hiddenAttemptTrace +
			"|fallback|" +
			(resolveTrace.empty() ? localResult.trace : (resolveTrace + "|" + localResult.trace));
	}
	else {
		localResult.trace = resolveTrace.empty() ? localResult.trace : (resolveTrace + "|" + localResult.trace);
	}
	if (outResult != nullptr) {
		*outResult = std::move(localResult);
	}
	return ok;
}

bool ReplaceRealPageCodeByEditorObject(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& newPageCode,
	const std::string* rollbackPageCode,
	NativeRealPageAccessResult* outResult)
{
	return ReplaceRealPageCodeByEditorObjectInternal(
		editorObject,
		moduleBase,
		newPageCode,
		rollbackPageCode,
		outResult);
}

bool ReplaceRealPageCodeByProgramTreeItemData(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	const std::string& newPageCode,
	const std::string* rollbackPageCode,
	NativeRealPageAccessResult* outResult)
{
	if (outResult != nullptr) {
		*outResult = {};
	}

	std::uintptr_t editorObject = 0;
	std::string resolveTrace;
	if (!ResolveEditorObjectByProgramTreeItemDataInternal(
			itemData,
			moduleBase,
			&editorObject,
			&resolveTrace)) {
		if (outResult != nullptr) {
			outResult->trace = resolveTrace.empty() ? "resolve_editor_failed" : resolveTrace;
		}
		return false;
	}

	NativeRealPageAccessResult localResult{};
	const bool ok = ReplaceRealPageCodeByEditorObjectInternal(
		editorObject,
		moduleBase,
		newPageCode,
		rollbackPageCode,
		&localResult);
	localResult.editorObject = editorObject;
	localResult.trace = resolveTrace.empty() ? localResult.trace : (resolveTrace + "|" + localResult.trace);
	if (outResult != nullptr) {
		*outResult = std::move(localResult);
	}
	return ok;
}

bool ResolveCurrentActiveEditorObject(
	std::uintptr_t moduleBase,
	ActiveEditorObjectInfo* outInfo)
{
	if (outInfo != nullptr) {
		*outInfo = {};
	}
	if (moduleBase == 0) {
		if (outInfo != nullptr) {
			outInfo->trace = "module_base_invalid";
		}
		return false;
	}

	const HWND mainHwnd = ResolveMainIdeWindow();
	if (mainHwnd == nullptr || !IsWindow(mainHwnd)) {
		if (outInfo != nullptr) {
			outInfo->trace = "main_ide_window_missing";
		}
		return false;
	}

	const HWND mdiChildHwnd = GetActiveMdiChildWindow(mainHwnd);
	if (mdiChildHwnd == nullptr || !IsWindow(mdiChildHwnd)) {
		if (outInfo != nullptr) {
			outInfo->trace = "active_mdi_child_missing";
		}
		return false;
	}

	std::uintptr_t rawEditorObject = 0;
	std::string resolveTrace;
	if (!TryResolveEditorObjectFromMdiChildHeuristic(
			mdiChildHwnd,
			moduleBase,
			&rawEditorObject,
			nullptr,
			&resolveTrace) ||
		rawEditorObject == 0) {
		if (outInfo != nullptr) {
			outInfo->trace =
				"resolve_active_editor_failed"
				"|mdi_child=" + std::to_string(reinterpret_cast<std::uintptr_t>(mdiChildHwnd)) +
				"|mdi_title=" + WindowTextToString(mdiChildHwnd) +
				"|" +
				resolveTrace;
		}
		return false;
	}

	ActiveEditorObjectInfo info{};
	info.ok = true;
	info.rawEditorObject = rawEditorObject;
	info.trace =
		"resolve_active_editor_ok"
		"|mdi_child=" + std::to_string(reinterpret_cast<std::uintptr_t>(mdiChildHwnd)) +
		"|mdi_title=" + WindowTextToString(mdiChildHwnd) +
		"|" +
		resolveTrace;

	EditorDispatchTargetInfo targetInfo{};
	if (TryResolveInnerEditorObject(rawEditorObject, &targetInfo)) {
		info.innerEditorObject = targetInfo.innerObject;
		info.pageType = targetInfo.pageType;
		info.trace +=
			"|inner=" + std::to_string(targetInfo.innerObject) +
			"|page_type=" + std::to_string(targetInfo.pageType);
	}

	if (outInfo != nullptr) {
		*outInfo = std::move(info);
	}
	return true;
}

bool CaptureCustomClipboardPayloadByThiscall(
	const InternalThiscallCommandSpec& spec,
	std::vector<unsigned char>& outBytes,
	std::string* outTrace)
{
	outBytes.clear();
	if (outTrace != nullptr) {
		outTrace->clear();
	}
	if (spec.targetObject == 0 || spec.functionAddress == 0) {
		if (outTrace != nullptr) {
			*outTrace = "capture_invalid_argument";
		}
		return false;
	}

	ScopedFakeClipboard fakeClipboard;
	if (!fakeClipboard.IsActive()) {
		if (outTrace != nullptr) {
			*outTrace = "capture_fake_clipboard_unavailable";
		}
		return false;
	}

	int invokeResult = 0;
	bool invokeThrew = false;
	(void)CallGenericThiscallCommandSafe(
		spec.targetObject,
		spec.functionAddress,
		spec.arg1,
		spec.arg2,
		spec.arg3,
		spec.arg4,
		&invokeResult,
		&invokeThrew);

	UINT capturedFormat = 0;
	HANDLE capturedHandle = nullptr;
	if (!fakeClipboard.GetFirstCustomFormat(&capturedFormat, &capturedHandle) || capturedHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				"capture_custom_format_missing"
				"|target=" + std::to_string(spec.targetObject) +
				"|function=" + std::to_string(spec.functionAddress) +
				"|invoke_ret=" + std::to_string(invokeResult) +
				"|invoke_exc=" + std::to_string(invokeThrew ? 1 : 0) +
				"|" +
				fakeClipboard.BuildStatsText();
		}
		return false;
	}

	HANDLE duplicatedHandle = nullptr;
	if (!DuplicateGlobalHandle(capturedHandle, &duplicatedHandle) || duplicatedHandle == nullptr) {
		if (outTrace != nullptr) {
			*outTrace =
				"capture_duplicate_handle_failed"
				"|format=" + std::to_string(capturedFormat) +
				"|target=" + std::to_string(spec.targetObject) +
				"|function=" + std::to_string(spec.functionAddress) +
				"|invoke_ret=" + std::to_string(invokeResult) +
				"|invoke_exc=" + std::to_string(invokeThrew ? 1 : 0) +
				"|" +
				fakeClipboard.BuildStatsText();
		}
		return false;
	}

	const bool copyOk = CopyGlobalHandleBytesNoFree(duplicatedHandle, &outBytes);
	GlobalFree(duplicatedHandle);
	if (!copyOk || outBytes.empty()) {
		if (outTrace != nullptr) {
			*outTrace =
				"capture_copy_bytes_failed"
				"|format=" + std::to_string(capturedFormat) +
				"|target=" + std::to_string(spec.targetObject) +
				"|function=" + std::to_string(spec.functionAddress) +
				"|invoke_ret=" + std::to_string(invokeResult) +
				"|invoke_exc=" + std::to_string(invokeThrew ? 1 : 0) +
				"|" +
				fakeClipboard.BuildStatsText();
		}
		return false;
	}

	if (outTrace != nullptr) {
		*outTrace =
			"capture_ok"
			"|format=" + std::to_string(capturedFormat) +
			"|target=" + std::to_string(spec.targetObject) +
			"|function=" + std::to_string(spec.functionAddress) +
			"|invoke_ret=" + std::to_string(invokeResult) +
			"|invoke_exc=" + std::to_string(invokeThrew ? 1 : 0) +
			"|bytes=" + std::to_string(outBytes.size()) +
			"|" +
			fakeClipboard.BuildStatsText();
	}
	return true;
}

}
