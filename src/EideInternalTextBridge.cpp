#include "EideInternalTextBridge.h"

#include <afxwin.h>
#include <Windows.h>

#include <detours.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <cctype>
#include <cstring>
#include <fstream>
#include <mutex>
#include <string>
#include <unordered_map>
#include <vector>

#include "MemFind.h"
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
constexpr std::uintptr_t kKnownEditorDispatchCommandRva = 0x4B6D70;
constexpr std::uintptr_t kKnownClipboardTextToObjectWrapperRva = 0x4C9220;
constexpr std::uintptr_t kKnownClipboardTextToObjectDirectRva = 0x4CB160;
constexpr std::uintptr_t kKnownClipboardInsertObjectRva = 0x48B720;
constexpr std::uintptr_t kKnownClipboardDeserializeRva = 0x4535F0;
constexpr std::uintptr_t kKnownClipboardSerializeRva = 0x45B640;
constexpr std::uintptr_t kKnownClipboardValidateA00Rva = 0x452A00;
constexpr std::uintptr_t kKnownClipboardValidate560Rva = 0x452560;
constexpr std::uintptr_t kKnownClipboardValidate230Rva = 0x452230;
constexpr std::uintptr_t kKnownClipboardCollectionFirstInvalidRva = 0x4521D0;
constexpr std::uintptr_t kKnownEditorPasteHandlerRva = 0x4BC830;
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
constexpr std::uintptr_t kKnownClipboardObjectInit00Rva = 0x44E2D0;
constexpr std::uintptr_t kKnownClipboardObjectCtorRva = 0x40E400;
constexpr std::uintptr_t kKnownClipboardObjectInit3CRva = 0x40EC50;
constexpr std::uintptr_t kKnownClipboardObjectInit94Rva = 0x53A56A;
constexpr std::uintptr_t kKnownClipboardObjectInitA8Rva = 0x53A9C8;
constexpr std::uintptr_t kKnownClipboardObjectInitBCRva = 0x40F9A0;
constexpr std::uintptr_t kKnownClipboardObjectInit188Rva = 0x40F520;
constexpr std::uintptr_t kKnownClipboardObjectInit1DCRva = 0x40E810;
constexpr std::uintptr_t kKnownClipboardObjectInit1FCRva = 0x40EB60;
constexpr std::uintptr_t kKnownClipboardObjectInit214Rva = 0x40EA70;
constexpr std::uintptr_t kKnownClipboardObjectInit230Rva = 0x40EBC0;
constexpr std::uintptr_t kKnownClipboardObjectInit24CRva = 0x40EB90;
constexpr std::uintptr_t kKnownClipboardObjectInit268Rva = 0x40EBF0;
constexpr std::uintptr_t kKnownClipboardObjectInit284Rva = 0x40E930;
constexpr std::uintptr_t kKnownClipboardObjectInit2A0Rva = 0x40E990;
constexpr std::uintptr_t kKnownClipboardObjectInit2BCRva = 0x40EC20;
constexpr std::uintptr_t kKnownClipboardObjectArraySetSizeRva = 0x53AAEE;
constexpr std::uintptr_t kKnownClipboardObjectFillMemoryRva = 0x486F90;
constexpr std::uintptr_t kKnownGenericArrayVtableRva = 0x574900;
constexpr std::uintptr_t kKnownEmptyStringRva = 0x5C72A4;
constexpr std::uintptr_t kKnownCollectionGetValueByIndexRva = 0x4E7EA0;
constexpr size_t kInternalClipboardObjectSize = 0x340;
constexpr size_t kInternalGenericArraySize = 0x14;
constexpr size_t kClipboardObjectSupportLibrariesOffset = 0x94;
constexpr size_t kClipboardObjectGuidFlagOffset = 0x10C;
constexpr size_t kClipboardObjectGuidOffset = 0x110;

using FnEditorDispatchCommand = int(__thiscall*)(void*, int, int, char*, char*);
using FnEditorUiCommand = int(__thiscall*)(void*, unsigned int);
using FnEditorPasteHandler = void(__thiscall*)(void*, int);
using FnClipboardTextToObjectHook = int(__fastcall*)(void*, void*, void*);
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
std::mutex g_clipboardTextParseHookMutex;
std::mutex g_clipboardTextPostProcessMutex;
bool g_clipboardTextParseHookInstalled = false;
std::atomic<bool> g_forceFullTextClipboardParse = false;
bool g_applySupportLibrariesAfterParse = false;
std::string g_parseSupportLibrarySourceText;
FnClipboardTextToObjectHook g_originalClipboardTextToObject = nullptr;
FnClipboardTextToObjectDirect g_directClipboardTextToObject = nullptr;

decltype(&::OpenClipboard) g_originalOpenClipboard = ::OpenClipboard;
decltype(&::EmptyClipboard) g_originalEmptyClipboard = ::EmptyClipboard;
decltype(&::SetClipboardData) g_originalSetClipboardData = ::SetClipboardData;
decltype(&::GetClipboardData) g_originalGetClipboardData = ::GetClipboardData;
decltype(&::IsClipboardFormatAvailable) g_originalIsClipboardFormatAvailable = ::IsClipboardFormatAvailable;
decltype(&::CloseClipboard) g_originalCloseClipboard = ::CloseClipboard;

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

bool PopulateNativeEditorCommandAddresses(NativeEditorCommandAddresses& addrs, std::uintptr_t moduleBase)
{
	addrs = {};
	addrs.moduleBase = moduleBase;
	addrs.editorDispatchCommand = ResolveUniqueCodeAddress(
		"A1 ?? ?? ?? ?? 53 56 8B F1 85 C0 74 0A 5E B8 01 00 00 00 5B C2 10 00 8A 5C 24 10 F6 C3 10 74 09",
		moduleBase);
	if (addrs.editorDispatchCommand == 0) {
		addrs.editorDispatchCommand = kKnownEditorDispatchCommandRva;
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

int __fastcall HookedClipboardTextToObject(void* thisPtr, void* edx, void* textObject)
{
	(void)edx;
	int result = 0;
	if (g_forceFullTextClipboardParse.load(std::memory_order_relaxed) &&
		g_directClipboardTextToObject != nullptr) {
		result = g_directClipboardTextToObject(thisPtr, textObject);
	}
	else {
		result = g_originalClipboardTextToObject != nullptr
			? g_originalClipboardTextToObject(thisPtr, nullptr, textObject)
			: 0;
	}

	if (result != 0 && thisPtr != nullptr) {
		std::string sourceText;
		{
			std::lock_guard<std::mutex> lock(g_clipboardTextPostProcessMutex);
			if (g_applySupportLibrariesAfterParse) {
				sourceText = g_parseSupportLibrarySourceText;
			}
		}
		if (!sourceText.empty()) {
			const std::vector<std::string> supportLibraries = ParseSupportLibraryNamesFromPageCode(sourceText);
			if (!supportLibraries.empty()) {
				auto* supportLibraryArray = reinterpret_cast<CStringArray*>(
					reinterpret_cast<std::byte*>(thisPtr) + kClipboardObjectSupportLibrariesOffset);
				supportLibraryArray->RemoveAll();
				for (const std::string& name : supportLibraries) {
					supportLibraryArray->Add(name.c_str());
				}
			}
		}
	}

	return result;
}

bool EnsureClipboardTextParseHookInstalled(std::uintptr_t moduleBase)
{
	std::lock_guard<std::mutex> lock(g_clipboardTextParseHookMutex);
	if (g_clipboardTextParseHookInstalled) {
		return true;
	}
	if (moduleBase == 0) {
		return false;
	}

	g_originalClipboardTextToObject = reinterpret_cast<FnClipboardTextToObjectHook>(
		moduleBase + (kKnownClipboardTextToObjectWrapperRva - kImageBase));
	g_directClipboardTextToObject = reinterpret_cast<FnClipboardTextToObjectDirect>(
		moduleBase + (kKnownClipboardTextToObjectDirectRva - kImageBase));
	if (g_originalClipboardTextToObject == nullptr || g_directClipboardTextToObject == nullptr) {
		return false;
	}

	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)g_originalClipboardTextToObject, HookedClipboardTextToObject);
	g_clipboardTextParseHookInstalled = (DetourTransactionCommit() == NO_ERROR);
	return g_clipboardTextParseHookInstalled;
}

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

class ScopedFullTextClipboardParseOverride {
public:
	explicit ScopedFullTextClipboardParseOverride(std::uintptr_t moduleBase)
	{
		m_active = EnsureClipboardTextParseHookInstalled(moduleBase);
		if (m_active) {
			g_forceFullTextClipboardParse.store(true, std::memory_order_relaxed);
		}
	}

	~ScopedFullTextClipboardParseOverride()
	{
		if (m_active) {
			g_forceFullTextClipboardParse.store(false, std::memory_order_relaxed);
		}
	}

	bool IsActive() const
	{
		return m_active;
	}

private:
	bool m_active = false;
};

class ScopedClipboardParseSupportLibraryOverride {
public:
	ScopedClipboardParseSupportLibraryOverride(std::uintptr_t moduleBase, const std::string& pageCode)
	{
		m_active = EnsureClipboardTextParseHookInstalled(moduleBase);
		if (!m_active) {
			return;
		}

		std::lock_guard<std::mutex> lock(g_clipboardTextPostProcessMutex);
		g_applySupportLibrariesAfterParse = true;
		g_parseSupportLibrarySourceText = pageCode;
	}

	~ScopedClipboardParseSupportLibraryOverride()
	{
		if (!m_active) {
			return;
		}

		std::lock_guard<std::mutex> lock(g_clipboardTextPostProcessMutex);
		g_applySupportLibrariesAfterParse = false;
		g_parseSupportLibrarySourceText.clear();
	}

	bool IsActive() const
	{
		return m_active;
	}

private:
	bool m_active = false;
};

std::string NormalizeLineBreakToCrLf(const std::string& text);
std::string NormalizeLineBreakToLf(const std::string& text);
std::string NormalizePageCodeForLooseCompare(const std::string& text);

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
		const unsigned int pageType = fields[15];
		std::uintptr_t innerObject = 0;
		switch (pageType) {
		case 1:
			innerObject = static_cast<std::uintptr_t>(fields[21]);
			break;
		case 2:
			innerObject = static_cast<std::uintptr_t>(fields[16]);
			break;
		case 3:
			innerObject = static_cast<std::uintptr_t>(fields[17]);
			break;
		case 4:
			innerObject = static_cast<std::uintptr_t>(fields[18]);
			break;
		case 6:
		case 7:
		case 8:
			innerObject = static_cast<std::uintptr_t>(fields[20]);
			break;
		default:
			innerObject = 0;
			break;
		}

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

	const auto pasteHandlerFn = moduleBase == 0
		? nullptr
		: reinterpret_cast<FnEditorPasteHandler>(
			moduleBase + (kKnownEditorPasteHandlerRva - kImageBase));
	if (pasteHandlerFn == nullptr) {
		if (outTrace != nullptr) {
			*outTrace = "paste_handler_resolve_failed|module_base=" + std::to_string(moduleBase);
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
		!expectedPageCode->empty() &&
		NormalizeLineBreakToLf(verifyCode) != NormalizeLineBreakToLf(*expectedPageCode)) {
		if (outTrace != nullptr) {
			*outTrace = rollbackWriteTrace + "|rollback_verify_mismatch";
		}
		return false;
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

	const std::string normalizedVerifyCode = NormalizePageCodeForLooseCompare(verifyCode);
	const std::string normalizedExpectedCode = NormalizePageCodeForLooseCompare(normalizedNewPageCode);
	if (normalizedVerifyCode != normalizedExpectedCode) {
		const std::string mismatchSummary = BuildTextMismatchSummary(
			normalizedExpectedCode,
			normalizedVerifyCode);
		AppendPageEditTraceLine("ReplaceRealPageCode.verify_mismatch|" + mismatchSummary);
		if (outResult != nullptr) {
			outResult->trace =
				selectTrace +
				"|" +
				writeStrategyTrace +
				"|" +
				replaceTrace +
				"|verify_mismatch|" +
				mismatchSummary;
		}
		return false;
	}
	AppendPageEditTraceLine("ReplaceRealPageCode.success");

	if (outResult != nullptr) {
		outResult->ok = true;
		outResult->usedClipboardEmulation = true;
		outResult->textBytes = verifyCode.size();
		outResult->trace =
			selectTrace +
			writeStrategyTrace +
			"|" +
			replaceTrace +
			"|verify_ok|" +
			verifyResult.trace;
	}
	return true;
}

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

	std::uintptr_t editorObject = 0;
	std::string openTrace;
	if (!DebugResolveEditorObjectByProgramTreeItemData(
			itemData,
			moduleBase,
			&editorObject,
			nullptr,
			nullptr,
			nullptr,
			&openTrace)) {
		if (outResult != nullptr) {
			outResult->trace = openTrace.empty() ? "resolve_editor_failed" : openTrace;
		}
		return false;
	}

	NativeRealPageAccessResult localResult{};
	const bool ok = GetRealPageCodeByEditorObject(editorObject, moduleBase, outCode, &localResult);
	localResult.editorObject = editorObject;
	localResult.trace = openTrace.empty() ? localResult.trace : (openTrace + "|" + localResult.trace);
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
	std::string openTrace;
	if (!DebugResolveEditorObjectByProgramTreeItemData(
			itemData,
			moduleBase,
			&editorObject,
			nullptr,
			nullptr,
			nullptr,
			&openTrace)) {
		if (outResult != nullptr) {
			outResult->trace = openTrace.empty() ? "resolve_editor_failed" : openTrace;
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
	localResult.trace = openTrace.empty() ? localResult.trace : (openTrace + "|" + localResult.trace);
	if (outResult != nullptr) {
		*outResult = std::move(localResult);
	}
	return ok;
}

}
