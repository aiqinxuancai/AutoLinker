#include "direct_global_search.hpp"
#include "direct_global_search_debug.hpp"

#include <detours.h>

#include <array>
#include <format>
#include <mbstring.h>

#include <cstring>
#include <mutex>

#include "MemFind.h"

extern HWND g_hwnd;
void OutputStringToELog(const std::string& szbuf);

namespace e571 {

namespace {

struct ExeCStringA;

using FnInitContext = DirectGlobalSearch::SearchContext*(__thiscall*)(DirectGlobalSearch::SearchContext*);
using FnGetOuterCount = int(__thiscall*)(DirectGlobalSearch::SearchContext*);
using FnGetInnerCount = int(__thiscall*)(DirectGlobalSearch::SearchContext*, int);
using FnFetchSearchText = int(__thiscall*)(
    DirectGlobalSearch::SearchContext*,
    int,
    int,
    CStringA*,
    int,
    int,
    CStringA*,
    int,
    int*);
using FnContainerGetAt = int(__thiscall*)(void*, int, int*);
using FnContainerGetId = int(__thiscall*)(void*, int);
using FnResolveBucketIndex = int(__thiscall*)(void*, int, int*, int*);
using FnOpenCodeTarget = HWND(__thiscall*)(void*, int, int, int, int, int, int, int);
using FnMoveToLine = int(__thiscall*)(HWND, int, int, int, int, int);
using FnEnsureVisible = int(__thiscall*)(HWND, int, int, int);
using FnMoveCaretToOffset = int(__thiscall*)(void*, int, int, void*, int);
using FnActivateWindow = int(__thiscall*)(HWND);
using FnNotifyOpenFailure = int(__thiscall*)(void*, int);
using FnCStringDestroy = void(__thiscall*)(void*);
using FnAppendBytes = int(__thiscall*)(void*, void*, int);
using FnPrepareSearchResults = int(__thiscall*)(void*, int);
using FnSelectSearchResultTab = int(__thiscall*)(void*, int);
using FnActivateWindowObject = int(__thiscall*)(int);
using FnSendMessageA = LRESULT(WINAPI*)(HWND, UINT, WPARAM, LPARAM);
using FnConsumeSearchResultRecord = void(__stdcall*)(void*, int);
using FnFromHandle = void*(__stdcall*)(void*);
using FnEditorGetOuterCount = int(__thiscall*)(void*);
using FnEditorGetInnerCount = int(__thiscall*)(void*, int);
using FnEditorFetchLineTextRaw = int(__thiscall*)(void*, int, int, ExeCStringA*, int, int, int, int*);

constexpr std::uintptr_t kImageBase = 0x400000;

struct NativeSearchAddresses {
    bool initialized = false;
    bool ok = false;
    std::uintptr_t moduleBase = 0;

    std::uintptr_t initContext = 0;
    std::uintptr_t getOuterCount = 0;
    std::uintptr_t getInnerCount = 0;
    std::uintptr_t fetchSearchText = 0;
    std::uintptr_t containerGetAt = 0;
    std::uintptr_t containerGetId = 0;
    std::uintptr_t resolveBucketIndex = 0;
    std::uintptr_t openCodeTarget = 0;
    std::uintptr_t moveToLine = 0;
    std::uintptr_t ensureVisible = 0;
    std::uintptr_t moveCaretToOffset = 0;
    std::uintptr_t activateWindow = 0;
    std::uintptr_t notifyOpenFailure = 0;
    std::uintptr_t cstringDestroy = 0;
    std::uintptr_t emptyCStringData = 0;

    std::uintptr_t type1Container = 0;
    std::uintptr_t type2Data = 0;
    std::uintptr_t type3Data = 0;
    std::uintptr_t type4Data = 0;
    std::uintptr_t type678Data = 0;
    std::uintptr_t mainEditorHost = 0;
    std::uintptr_t ownerObject = 0;
    std::uintptr_t builtinSearchDialogCtor = 0;
    std::uintptr_t builtinSearchDialogDtor = 0;
    std::uintptr_t dialogDoModal = 0;
    std::uintptr_t searchMode = 0;
    std::uintptr_t consumeSearchResultRecord = 0;
    std::uintptr_t fromHandle = 0;
    std::uintptr_t editorGetOuterCount = 0;
    std::uintptr_t editorGetInnerCount = 0;
    std::uintptr_t editorFetchLineText = 0;
    std::uintptr_t appendBytes = 0;
    std::uintptr_t prepareSearchResults = 0;
    std::uintptr_t selectSearchResultTab = 0;
};

constexpr int kSearchTypes[] = {1, 2, 3, 4, 6, 7, 8};
constexpr size_t kBuiltinSearchDialogStorageSize = 0x200;
constexpr size_t kBuiltinResultPreviewLimit = 10;
constexpr ptrdiff_t kOffset_SearchKeywordCtrlHwnd = 300;
constexpr ptrdiff_t kOffset_SearchCaseCtrlHwnd = 180;
constexpr ptrdiff_t kOffset_SearchTypeCtrlHwnd = 360;
constexpr ptrdiff_t kOffset_ResultPageType1 = 0xB88;
constexpr ptrdiff_t kOffset_ResultPageType2 = 0xBD0;
constexpr ptrdiff_t kOffset_ResultRecordContainerType1 = 0xC54;
constexpr ptrdiff_t kOffset_ResultRecordContainerType2 = 0xC68;
constexpr ptrdiff_t kOffset_CWndHwnd = 28;
constexpr ptrdiff_t kOffset_ByteContainerData = 8;
constexpr ptrdiff_t kOffset_ByteContainerUsedBytes = 16;
constexpr UINT kMsg_TcmGetCurSel = 0x130B;
constexpr UINT kMsg_TcmSetCurSel = 0x130C;
constexpr UINT kMsg_TcmGetCurFocus = 0x132F;
constexpr UINT kMsg_TcmSetCurFocus = 0x1330;

struct ExeCStringA {
    const char* data;
};

struct Type1BucketEntry {
    int textPtr;
    int unused1;
    unsigned char flags;
};

struct NativeSearchResultRecordBuffer {
    int unused0;
    int unused1;
    void* data;
    int capacityBytes;
    int usedBytes;
};

using FnBuiltinSearchDialogCtor = void*(__thiscall*)(void*, CWnd*);
using FnBuiltinSearchDialogDtor = void*(__thiscall*)(void*, unsigned char);
using FnDialogDoModal = int(__thiscall*)(void*);
using FnSearchPageDispatch = int(__thiscall*)(void*, int, int, int, int);

struct HiddenBuiltinSearchContext {
    const char* keyword = nullptr;
    std::uintptr_t moduleBase = 0;
    void* dialogObject = nullptr;
    HHOOK cbtHook = nullptr;
    HHOOK callWndHook = nullptr;
    HWND dialogHwnd = nullptr;
    HWND resultListHwnd = nullptr;
    HWND resultPageHwnd = nullptr;
    HWND mainFrameHwnd = nullptr;
    HWND previousFocusHwnd = nullptr;
    void* previousFocusObject = nullptr;
    void* resultContainerObject = nullptr;
    void* resultPageObject = nullptr;
    bool dialogHandled = false;
    bool searchFinished = false;
    size_t fallbackHitCount = 0;
    std::string fallbackFirstResultText;
    std::vector<std::string> fallbackPreviewLines;
    std::vector<e571::DirectGlobalSearch::GlobalSearchHit> rawHits;
};

thread_local HiddenBuiltinSearchContext* g_hiddenBuiltinSearchContext = nullptr;

FnAppendBytes g_originalAppendBytes = nullptr;
FnPrepareSearchResults g_originalPrepareSearchResults = nullptr;
FnSelectSearchResultTab g_originalSelectSearchResultTab = nullptr;
FnActivateWindowObject g_originalActivateWindowObject = nullptr;
FnSendMessageA g_originalSendMessageA = ::SendMessageA;

std::mutex g_nativeSearchAddressMutex;
NativeSearchAddresses g_nativeSearchAddresses;

std::uintptr_t NormalizeRuntimeAddress(std::uintptr_t runtimeAddress, std::uintptr_t moduleBase) {
    if (runtimeAddress == 0 || moduleBase == 0 || runtimeAddress < moduleBase) {
        return 0;
    }
    return runtimeAddress - moduleBase + kImageBase;
}

std::uintptr_t ReadNormalizedImm32(std::uintptr_t instructionAddress, size_t immOffset, std::uintptr_t moduleBase) {
    if (instructionAddress == 0) {
        return 0;
    }
    const auto runtimeValue = static_cast<std::uintptr_t>(
        *reinterpret_cast<const std::uint32_t*>(instructionAddress + immOffset));
    return NormalizeRuntimeAddress(runtimeValue, moduleBase);
}

std::uintptr_t ReadNormalizedAbs32(std::uintptr_t absoluteAddress, std::uintptr_t moduleBase) {
    if (absoluteAddress == 0 || moduleBase == 0 || absoluteAddress < kImageBase) {
        return 0;
    }
    const auto runtimeAddress = absoluteAddress - kImageBase + moduleBase;
    const auto runtimeValue = static_cast<std::uintptr_t>(*reinterpret_cast<const std::uint32_t*>(runtimeAddress));
    return NormalizeRuntimeAddress(runtimeValue, moduleBase);
}

std::uintptr_t ResolveUniqueCodeAddress(
    const char* label,
    const char* pattern,
    std::uintptr_t moduleBase) {
    const auto matches = FindSelfModelMemoryAll(pattern);
    if (matches.size() != 1) {
        OutputStringToELog(std::format(
            "[DirectGlobalSearch] resolve {} failed, matchCount={}",
            label,
            matches.size()));
        return 0;
    }
    return NormalizeRuntimeAddress(static_cast<std::uintptr_t>(matches.front()), moduleBase);
}

std::uintptr_t ResolveUniqueImmAddress(
    const char* label,
    const char* pattern,
    size_t immOffset,
    std::uintptr_t moduleBase) {
    const auto matches = FindSelfModelMemoryAll(pattern);
    if (matches.size() != 1) {
        OutputStringToELog(std::format(
            "[DirectGlobalSearch] resolve {} failed, matchCount={}",
            label,
            matches.size()));
        return 0;
    }
    return ReadNormalizedImm32(static_cast<std::uintptr_t>(matches.front()), immOffset, moduleBase);
}

bool PopulateNativeSearchAddresses(NativeSearchAddresses& addrs, std::uintptr_t moduleBase) {
    addrs = {};
    addrs.moduleBase = moduleBase;

    addrs.initContext = ResolveUniqueCodeAddress(
        "init_context",
        "8B C1 33 C9 89 08 89 48 0C 89 48 04 89 48 08 C3",
        moduleBase);
    addrs.getOuterCount = ResolveUniqueCodeAddress(
        "get_outer_count",
        "83 EC 10 53 55 56 8B F1 33 DB 33 ED 8B 06 57 48 83 F8 07 0F 87 20 01 00 00",
        moduleBase);
    addrs.getInnerCount = ResolveUniqueCodeAddress(
        "get_inner_count",
        "83 EC 0C 8B 01 53 48 55 56 83 F8 07 57 89 4C 24 10 0F 87 0F 01 00 00",
        moduleBase);
    addrs.fetchSearchText = ResolveUniqueCodeAddress(
        "fetch_search_text",
        "83 EC 0C 53 55 56 57 8B F9 8B 4C 24 28 E8 ?? ?? ?? ?? 8B 6C 24 40 85 ED 74 07 C7 45 00 00 00 00 00",
        moduleBase);
    addrs.containerGetAt = ResolveUniqueCodeAddress(
        "container_get_at",
        "8B 41 18 8B 54 24 04 C1 E8 03 3B D0 7D 23 56 8B 71 18 85 F6 5E 75 04 33 C9 EB 03 8B 49 10 03 C2",
        moduleBase);
    addrs.containerGetId = ResolveUniqueCodeAddress(
        "container_get_id",
        "8B 51 18 8B 44 24 04 C1 EA 03 3B C2 7D 18 8B 51 18 85 D2 75 08 33 C9 8B 04 81 C2 04 00 8B 49 10",
        moduleBase);
    addrs.resolveBucketIndex = ResolveUniqueCodeAddress(
        "resolve_bucket_index",
        "53 8B 5C 24 08 56 57 8B F1 53 E8 ?? ?? ?? ?? 83 C4 04 85 C0 75 06 5F 5E 5B C2 0C 00 8B 46 18 85 C0 75 04 33 FF EB 03 8B 7E 10",
        moduleBase);
    addrs.openCodeTarget = ResolveUniqueCodeAddress(
        "open_code_target",
        "83 EC 08 53 56 57 8B F9 8B 87 CC 03 00 00 F7 D0 83 E0 01 3C 01 0F 84 C0 02 00 00 8B 5C 24 18 33 F6 83 FB 01 74 39",
        moduleBase);
    addrs.moveToLine = ResolveUniqueCodeAddress(
        "move_to_line",
        "6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC 90 00 00 00 8B 84 24 A0 00 00 00 55 56 57 8B F1 50",
        moduleBase);
    addrs.ensureVisible = ResolveUniqueCodeAddress(
        "ensure_visible",
        "6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 83 EC 44 8B 44 24 58 53 55 8B 6C 24 5C 56 57 85 C0 8B F1",
        moduleBase);
    addrs.moveCaretToOffset = ResolveUniqueCodeAddress(
        "move_caret_to_offset",
        "64 A1 00 00 00 00 6A FF 68 ?? ?? ?? ?? 50 64 89 25 00 00 00 00 83 EC 28 55 8B 6C 24 3C 56 83 FD FF 8B F1 75 06",
        moduleBase);
    addrs.activateWindow = ResolveUniqueCodeAddress(
        "activate_window",
        "8B 41 38 85 C0 75 10 FF 71 1C FF 15 ?? ?? ?? ?? 50 E8 ?? ?? ?? ?? C3 8B 10 8B C8 FF A2 AC 00 00",
        moduleBase);
    addrs.notifyOpenFailure = ResolveUniqueCodeAddress(
        "notify_open_failure",
        "8B 44 24 04 50 FF 15 ?? ?? ?? ?? C2 04 00 90 90 51 56 8B F1 8B 8E 24 09 00 00 85 C9 74 2B",
        moduleBase);
    addrs.cstringDestroy = ResolveUniqueCodeAddress(
        "cstring_destroy",
        "56 8B F1 8B 06 8D 48 F4 3B 0D ?? ?? ?? ?? 74 18 83 C0 F4 50 FF 15 ?? ?? ?? ?? 85 C0 7F 0A 8B 0E",
        moduleBase);

    addrs.emptyCStringData = ResolveUniqueImmAddress(
        "empty_cstring_data",
        "A1 ?? ?? ?? ?? 53 56 8B F1 33 DB 57 89 74 24 6C 33 FF 89 5C 24 1C",
        1,
        moduleBase);
    addrs.type1Container = ResolveUniqueImmAddress(
        "type1_container",
        "8B C1 41 89 4C 24 48 52 50 B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 85 C0 0F 84 1B 04 00 00 8B 44 24",
        10,
        moduleBase);
    addrs.type2Data = ResolveUniqueImmAddress(
        "type2_data",
        "68 ?? ?? ?? ?? 8D 4C 24 14 C7 44 24 58 ?? ?? ?? ?? E8 ?? ?? ?? ?? EB 54",
        13,
        moduleBase);
    addrs.type3Data = ResolveUniqueImmAddress(
        "type3_data",
        "68 ?? ?? ?? ?? 8D 4C 24 14 C7 44 24 58 ?? ?? ?? ?? E8 ?? ?? ?? ?? EB 3C",
        13,
        moduleBase);
    addrs.type4Data = ResolveUniqueImmAddress(
        "type4_data",
        "68 ?? ?? ?? ?? 8D 4C 24 14 C7 44 24 58 ?? ?? ?? ?? E8 ?? ?? ?? ?? EB 24",
        13,
        moduleBase);
    addrs.type678Data = ResolveUniqueImmAddress(
        "type678_data",
        "E8 ?? ?? ?? ?? C7 44 24 54 ?? ?? ?? ?? A1 ?? ?? ?? ?? 89 44 24 2C 33 F6 33 ED",
        9,
        moduleBase);
    addrs.mainEditorHost = ResolveUniqueImmAddress(
        "main_editor_host",
        "A1 ?? ?? ?? ?? 3B C7 74 0A B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 8D 44 24 14 55 50 B9 ?? ?? ?? ?? E8",
        10,
        moduleBase);
    addrs.ownerObject = ResolveUniqueImmAddress(
        "owner_object",
        "8B 3D ?? ?? ?? ?? 83 C9 FF 33 C0 C7 44 24 5C ?? ?? ?? ?? 89 7C 24 6C F2 AE A1 ?? ?? ?? ?? 33 FF F7 D1 49 33 D2 3B C7 8B 86 68 01 00 00",
        15,
        moduleBase);
    addrs.builtinSearchDialogCtor = ResolveUniqueCodeAddress(
        "builtin_search_dialog_ctor",
        "6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 51 8B 44 24 14 56 57 8B F1 50 68 35 78 00 00",
        moduleBase);
    {
        const auto builtinSearchDialogVftable = ResolveUniqueImmAddress(
            "builtin_search_dialog_vftable",
            "8B 4C 24 0C C7 06 ?? ?? ?? ?? C7 86 8C 01 00 00 00 00 00 00",
            6,
            moduleBase);
        addrs.builtinSearchDialogDtor = ReadNormalizedAbs32(builtinSearchDialogVftable + sizeof(std::uint32_t), moduleBase);
        if (addrs.builtinSearchDialogDtor == 0) {
            OutputStringToELog("[DirectGlobalSearch] resolve builtin_search_dialog_dtor failed");
        }
    }
    addrs.dialogDoModal = ResolveUniqueCodeAddress(
        "dialog_do_modal",
        "B8 ?? ?? ?? ?? E8 ?? ?? ?? ?? 83 EC 18 53 56 8B F1 57 89 65 F0 89 75 E4 8B 46 48 8B 7E 44 89 45 E8",
        moduleBase);
    addrs.searchMode = ResolveUniqueImmAddress(
        "search_mode",
        "FF 15 ?? ?? ?? ?? 48 F7 D8 1B C0 83 C0 02 A3 ?? ?? ?? ?? C3",
        15,
        moduleBase);
    addrs.consumeSearchResultRecord = ResolveUniqueCodeAddress(
        "consume_search_result_record",
        "8B 4C 24 08 56 83 F9 FF 57 0F 84 07 01 00 00 49 0F 88 F4 00 00 00 8B 7C 24 0C B8 CD CC CC CC 8B 77 10",
        moduleBase);
    addrs.fromHandle = ResolveUniqueCodeAddress(
        "from_handle",
        "56 57 6A 01 E8 ?? ?? ?? ?? 8B F0 FF 74 24 0C 8B CE E8 ?? ?? ?? ?? 8B F8 56 8B CF E8 ?? ?? ?? ??",
        moduleBase);
    addrs.editorGetOuterCount = ResolveUniqueCodeAddress(
        "editor_get_outer_count",
        "83 EC 10 51 8D 4C 24 04 E8 33 6F 00 00 8D 4C 24 00 E8 EA E1 FC FF 83 C4 10 C3",
        moduleBase);
    addrs.editorGetInnerCount = ResolveUniqueCodeAddress(
        "editor_get_inner_count",
        "83 EC 10 51 8D 4C 24 04 E8 A3 20 00 00 8B 44 24 14 8D 4C 24 00 50 E8 25 DA FC FF 83 C4 10 C2 04",
        moduleBase);
    addrs.editorFetchLineText = ResolveUniqueCodeAddress(
        "editor_fetch_line_text",
        "83 EC 10 51 8D 4C 24 04 E8 F3 20 00 00 8B 44 24 2C 8B 4C 24 28 8B 54 24 24 50 8B 44 24 24 6A 00",
        moduleBase);
    addrs.appendBytes = ResolveUniqueCodeAddress(
        "append_bytes",
        "56 57 8B 7C 24 10 85 FF 7E 3D 8B 71 10 8D 04 3E 50 E8 ?? ?? ?? ?? 85 C0 75 05 5F 5E C2 08 00",
        moduleBase);
    addrs.prepareSearchResults = ResolveUniqueCodeAddress(
        "prepare_search_results",
        "53 8B 1D ?? ?? ?? ?? 56 57 8B 7C 24 10 8B F1 85 FF 74 05 83 FF 02 74 1D 8B 86 8C 03 00 00 6A 00",
        moduleBase);
    addrs.selectSearchResultTab = ResolveUniqueCodeAddress(
        "select_search_result_tab",
        "8B 54 24 04 33 C0 83 FA 02 53 0F 94 C0 56 57 83 C0 05 8B F1 68 E8 03 00 00 8B F8 E8 ?? ?? ?? ??",
        moduleBase);

    addrs.ok =
        addrs.initContext != 0 &&
        addrs.getOuterCount != 0 &&
        addrs.getInnerCount != 0 &&
        addrs.fetchSearchText != 0 &&
        addrs.containerGetAt != 0 &&
        addrs.containerGetId != 0 &&
        addrs.resolveBucketIndex != 0 &&
        addrs.openCodeTarget != 0 &&
        addrs.moveToLine != 0 &&
        addrs.ensureVisible != 0 &&
        addrs.moveCaretToOffset != 0 &&
        addrs.activateWindow != 0 &&
        addrs.notifyOpenFailure != 0 &&
        addrs.cstringDestroy != 0 &&
        addrs.emptyCStringData != 0 &&
        addrs.type1Container != 0 &&
        addrs.type2Data != 0 &&
        addrs.type3Data != 0 &&
        addrs.type4Data != 0 &&
        addrs.type678Data != 0 &&
        addrs.mainEditorHost != 0 &&
        addrs.ownerObject != 0 &&
        addrs.builtinSearchDialogCtor != 0 &&
        addrs.builtinSearchDialogDtor != 0 &&
        addrs.dialogDoModal != 0 &&
        addrs.searchMode != 0 &&
        addrs.consumeSearchResultRecord != 0 &&
        addrs.fromHandle != 0 &&
        addrs.editorGetOuterCount != 0 &&
        addrs.editorGetInnerCount != 0 &&
        addrs.editorFetchLineText != 0 &&
        addrs.appendBytes != 0 &&
        addrs.prepareSearchResults != 0 &&
        addrs.selectSearchResultTab != 0;
    addrs.initialized = true;
    return addrs.ok;
}

const NativeSearchAddresses& GetNativeSearchAddresses(std::uintptr_t moduleBase) {
    std::lock_guard<std::mutex> lock(g_nativeSearchAddressMutex);
    if (!g_nativeSearchAddresses.initialized || g_nativeSearchAddresses.moduleBase != moduleBase) {
        PopulateNativeSearchAddresses(g_nativeSearchAddresses, moduleBase);
    }
    return g_nativeSearchAddresses;
}

bool HasNativeSearchAddresses(std::uintptr_t moduleBase) {
    return GetNativeSearchAddresses(moduleBase).ok;
}

void* GetMainWindowObject(std::uintptr_t moduleBase) {
    if (g_hwnd == nullptr || !::IsWindow(g_hwnd)) {
        return nullptr;
    }
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (addrs.fromHandle == 0) {
        return nullptr;
    }
    return reinterpret_cast<FnFromHandle>(addrs.fromHandle - kImageBase + moduleBase)(g_hwnd);
}

bool IsBuiltinSearchDecorativeLine(const std::string& text);
bool TryReadBuiltinResultHit(
    std::uintptr_t moduleBase,
    size_t index,
    e571::DirectGlobalSearch::GlobalSearchHit* outHit,
    size_t* outTotalCount);
bool LooksLikeSearchHit(const e571::DirectGlobalSearch::GlobalSearchHit& hit);
bool TryFormatRawSearchHit(
    std::uintptr_t moduleBase,
    const e571::DirectGlobalSearch::GlobalSearchHit& hit,
    std::string* outText);
bool TryResolveEditorObjectForHit(
    std::uintptr_t moduleBase,
    const e571::DirectGlobalSearch::GlobalSearchHit& hit,
    void** outEditorObject,
    std::string* outTrace);
bool TryDumpEditorPageCode(
    std::uintptr_t moduleBase,
    void* editorObject,
    std::string* outCode,
    e571::NativeEditorPageDumpDebugResult* outResult);
bool TryDumpSearchContextCode(
    e571::DirectGlobalSearch::SearchContext* ctx,
    std::uintptr_t moduleBase,
    std::string* outCode,
    int* outOuterCount,
    size_t* outLineCount,
    size_t* outFetchFailures);
bool SafeGetOuterCount(FnGetOuterCount fn, e571::DirectGlobalSearch::SearchContext* ctx, int* outValue);
bool SafeGetInnerCount(FnGetInnerCount fn, e571::DirectGlobalSearch::SearchContext* ctx, int outerIndex, int* outValue);
bool SafeEditorGetOuterCount(FnEditorGetOuterCount fn, void* editorObject, int* outValue);
bool SafeEditorGetInnerCount(FnEditorGetInnerCount fn, void* editorObject, int outerIndex, int* outValue);
bool SafeEditorFetchLineText(
    FnEditorFetchLineTextRaw fn,
    void* editorObject,
    int outerIndex,
    int innerIndex,
    ExeCStringA* outText,
    int* outOk,
    int* optionalOut);
bool InvokeNativeSearchResultRecordConsumer(
    std::uintptr_t moduleBase,
    const e571::DirectGlobalSearch::GlobalSearchHit* hits,
    size_t hitCount,
    size_t hitIndex);
bool SafeFetchSearchText(
    int(__thiscall* fn)(
        e571::DirectGlobalSearch::SearchContext*,
        int,
        int,
        ExeCStringA*,
        int,
        int,
        ExeCStringA*,
        int,
        int*),
    e571::DirectGlobalSearch::SearchContext* ctx,
    int outerIndex,
    int innerIndex,
    ExeCStringA* outText,
    ExeCStringA* outPrefix,
    int* outOk);
bool SafeResolveBucketIndex(
    FnResolveBucketIndex fn,
    void* container,
    int bucketId,
    int* outValue,
    int* outPos);
bool SafeContainerGetAt(
    FnContainerGetAt fn,
    void* container,
    int index,
    int* outPtr,
    int* outOk);
int ResolveTypeDataRaw(int type, std::uintptr_t moduleBase);

class ScopedHiddenBuiltinSearchApiHooks {
public:
    ScopedHiddenBuiltinSearchApiHooks(HiddenBuiltinSearchContext& ctx, std::uintptr_t moduleBase);
    ~ScopedHiddenBuiltinSearchApiHooks();
    bool IsInstalled() const;

private:
    bool installed_;
};

class ScopedHiddenBuiltinSearchUiGuard {
public:
    ScopedHiddenBuiltinSearchUiGuard(HiddenBuiltinSearchContext& ctx, std::uintptr_t moduleBase);
    ~ScopedHiddenBuiltinSearchUiGuard();

private:
    HiddenBuiltinSearchContext* ctx_;
    std::uintptr_t moduleBase_;
    bool redrawDisabled_;
};

template <typename T>
T BindAbsolute(std::uintptr_t moduleBase, std::uintptr_t absoluteAddress) {
    return reinterpret_cast<T>(absoluteAddress - kImageBase + moduleBase);
}

template <typename T>
T* PtrAbsolute(std::uintptr_t moduleBase, std::uintptr_t absoluteAddress) {
    return reinterpret_cast<T*>(absoluteAddress - kImageBase + moduleBase);
}

void InitExeCString(ExeCStringA& value, std::uintptr_t moduleBase) {
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    value.data = addrs.emptyCStringData != 0
        ? *PtrAbsolute<const char*>(moduleBase, addrs.emptyCStringData)
        : "";
}

void DestroyExeCString(ExeCStringA& value, std::uintptr_t moduleBase) {
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (addrs.cstringDestroy == 0) {
        value.data = "";
        return;
    }
    BindAbsolute<FnCStringDestroy>(moduleBase, addrs.cstringDestroy)(&value);
    InitExeCString(value, moduleBase);
}

const char* GetExeCStringData(const ExeCStringA& value) {
    return value.data != nullptr ? value.data : "";
}

bool MatchAtMbcs(const char* haystack, const char* needle, std::size_t count, bool matchCase) {
    if (matchCase) {
        return ::_mbsnbcmp(
                   reinterpret_cast<const unsigned char*>(haystack),
                   reinterpret_cast<const unsigned char*>(needle),
                   static_cast<unsigned int>(count)) == 0;
    }

    return ::_mbsnbicmp(
               reinterpret_cast<const unsigned char*>(haystack),
               reinterpret_cast<const unsigned char*>(needle),
               static_cast<unsigned int>(count)) == 0;
}

std::string BuildDisplayTextRaw(const ExeCStringA& prefix, const ExeCStringA& text) {
    const char* prefixData = GetExeCStringData(prefix);
    const char* textData = GetExeCStringData(text);
    if (prefixData[0] == '\0') {
        return textData;
    }
    return std::string(prefixData) + " " + textData;
}

bool IsListBoxWindow(HWND hwnd) {
    char className[32] = {};
    return ::GetClassNameA(hwnd, className, static_cast<int>(std::size(className))) > 0 &&
           std::strcmp(className, "ListBox") == 0;
}

HWND ReadDialogControlHwnd(void* dialogObject, ptrdiff_t offset) {
    if (dialogObject == nullptr) {
        return nullptr;
    }
    return *reinterpret_cast<HWND*>(reinterpret_cast<unsigned char*>(dialogObject) + offset);
}

LRESULT CALLBACK HiddenBuiltinSearchCbtProc(int code, WPARAM wParam, LPARAM lParam) {
    HiddenBuiltinSearchContext* ctx = g_hiddenBuiltinSearchContext;
    if (ctx != nullptr && code == HCBT_ACTIVATE && !ctx->dialogHandled) {
        HWND hwnd = reinterpret_cast<HWND>(wParam);
        char className[32] = {};
        if (::GetClassNameA(hwnd, className, static_cast<int>(std::size(className))) > 0 &&
            std::strcmp(className, "#32770") == 0) {
            ctx->dialogHandled = true;
            ctx->dialogHwnd = hwnd;

            ::ShowWindow(hwnd, SW_HIDE);

            HWND keywordHwnd = ReadDialogControlHwnd(ctx->dialogObject, kOffset_SearchKeywordCtrlHwnd);
            if (::IsWindow(keywordHwnd) && ctx->keyword != nullptr) {
                ::SetWindowTextA(keywordHwnd, ctx->keyword);
            }

            HWND caseHwnd = ReadDialogControlHwnd(ctx->dialogObject, kOffset_SearchCaseCtrlHwnd);
            if (::IsWindow(caseHwnd)) {
                ::SendMessageA(caseHwnd, BM_SETCHECK, BST_UNCHECKED, 0);
            }

            HWND typeHwnd = ReadDialogControlHwnd(ctx->dialogObject, kOffset_SearchTypeCtrlHwnd);
            if (::IsWindow(typeHwnd)) {
                ::SendMessageA(typeHwnd, CB_SETCURSEL, 0, 0);
            }

            HWND okButton = ::GetDlgItem(hwnd, IDOK);
            ::PostMessageA(
                hwnd,
                WM_COMMAND,
                MAKEWPARAM(IDOK, BN_CLICKED),
                reinterpret_cast<LPARAM>(okButton));
        }
    }

    return ::CallNextHookEx(ctx != nullptr ? ctx->cbtHook : nullptr, code, wParam, lParam);
}

LRESULT CALLBACK HiddenBuiltinSearchCallWndProc(int code, WPARAM wParam, LPARAM lParam) {
    HiddenBuiltinSearchContext* ctx = g_hiddenBuiltinSearchContext;
    if (ctx != nullptr && code >= 0 && lParam != 0) {
        const auto* cwp = reinterpret_cast<const CWPSTRUCT*>(lParam);
        if (IsListBoxWindow(cwp->hwnd)) {
            if (cwp->message == LB_RESETCONTENT) {
                if (ctx->resultListHwnd == nullptr || ctx->resultListHwnd == cwp->hwnd) {
                    ctx->resultListHwnd = cwp->hwnd;
                    ctx->fallbackHitCount = 0;
                    ctx->fallbackFirstResultText.clear();
                    ctx->fallbackPreviewLines.clear();
                    ctx->searchFinished = false;
                }
            } else if (cwp->message == LB_ADDSTRING && cwp->lParam != 0) {
                const char* textPtr = reinterpret_cast<const char*>(cwp->lParam);
                std::string text = textPtr != nullptr ? textPtr : "";
                const bool lineLooksRelevant =
                    (!text.empty() && ctx->keyword != nullptr && std::strstr(text.c_str(), ctx->keyword) != nullptr) ||
                    IsBuiltinSearchDecorativeLine(text);
                if (ctx->resultListHwnd == cwp->hwnd || (ctx->resultListHwnd == nullptr && lineLooksRelevant)) {
                    ctx->resultListHwnd = cwp->hwnd;
                    if (IsBuiltinSearchDecorativeLine(text)) {
                        if ((ctx->keyword == nullptr || std::strstr(text.c_str(), ctx->keyword) == nullptr) &&
                            (ctx->fallbackHitCount > 0 || !ctx->rawHits.empty())) {
                            ctx->searchFinished = true;
                        }
                    } else {
                        ++ctx->fallbackHitCount;
                        if (ctx->fallbackFirstResultText.empty()) {
                            ctx->fallbackFirstResultText = text;
                        }
                        if (ctx->fallbackPreviewLines.size() < kBuiltinResultPreviewLimit) {
                            ctx->fallbackPreviewLines.push_back(std::move(text));
                        }
                    }
                }
            }
        }
    }

    return ::CallNextHookEx(ctx != nullptr ? ctx->callWndHook : nullptr, code, wParam, lParam);
}

class ScopedHiddenBuiltinSearchHook {
public:
    explicit ScopedHiddenBuiltinSearchHook(HiddenBuiltinSearchContext& ctx)
        : ctx_(ctx) {
        g_hiddenBuiltinSearchContext = &ctx_;
        ctx_.cbtHook = ::SetWindowsHookExA(WH_CBT, HiddenBuiltinSearchCbtProc, nullptr, ::GetCurrentThreadId());
        ctx_.callWndHook = ::SetWindowsHookExA(WH_CALLWNDPROC, HiddenBuiltinSearchCallWndProc, nullptr, ::GetCurrentThreadId());
    }

    ~ScopedHiddenBuiltinSearchHook() {
        if (ctx_.callWndHook != nullptr) {
            ::UnhookWindowsHookEx(ctx_.callWndHook);
            ctx_.callWndHook = nullptr;
        }
        if (ctx_.cbtHook != nullptr) {
            ::UnhookWindowsHookEx(ctx_.cbtHook);
            ctx_.cbtHook = nullptr;
        }
        g_hiddenBuiltinSearchContext = nullptr;
    }

    bool IsInstalled() const {
        return ctx_.cbtHook != nullptr && ctx_.callWndHook != nullptr;
    }

private:
    HiddenBuiltinSearchContext& ctx_;
};

HWND GetBuiltinResultListHwnd(std::uintptr_t moduleBase) {
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (addrs.searchMode == 0) {
        return nullptr;
    }
    const int mode = *PtrAbsolute<int>(moduleBase, addrs.searchMode);
    auto* mainWindowObject = reinterpret_cast<unsigned char*>(GetMainWindowObject(moduleBase));
    if (mainWindowObject == nullptr) {
        return nullptr;
    }
    const ptrdiff_t resultPageOffset = mode == 2 ? kOffset_ResultPageType2 : kOffset_ResultPageType1;
    const HWND pageHwnd = *reinterpret_cast<HWND*>(mainWindowObject + resultPageOffset + kOffset_CWndHwnd);
    if (!::IsWindow(pageHwnd)) {
        return nullptr;
    }

    struct FindListBoxContext {
        HWND bestHwnd = nullptr;
        LRESULT bestCount = -1;
    } ctx;

    auto enumProc = [](HWND hwnd, LPARAM lParam) -> BOOL {
        auto& ctxRef = *reinterpret_cast<FindListBoxContext*>(lParam);
        char className[32] = {};
        if (::GetClassNameA(hwnd, className, static_cast<int>(std::size(className))) <= 0) {
            return TRUE;
        }
        if (std::strcmp(className, "ListBox") != 0) {
            return TRUE;
        }

        const LRESULT count = ::SendMessageA(hwnd, LB_GETCOUNT, 0, 0);
        if (count == LB_ERR || count < 0) {
            return TRUE;
        }

        if (count > ctxRef.bestCount) {
            ctxRef.bestCount = count;
            ctxRef.bestHwnd = hwnd;
        }
        return TRUE;
    };

    ::EnumChildWindows(
        pageHwnd,
        enumProc,
        reinterpret_cast<LPARAM>(&ctx));
    return ctx.bestHwnd;
}

bool TryReadBuiltinResultHit(
    std::uintptr_t moduleBase,
    size_t index,
    e571::DirectGlobalSearch::GlobalSearchHit* outHit,
    size_t* outTotalCount) {
    static_assert(sizeof(e571::DirectGlobalSearch::GlobalSearchHit) == 0x14, "Unexpected GlobalSearchHit size.");

    if (outTotalCount != nullptr) {
        *outTotalCount = 0;
    }
    if (outHit != nullptr) {
        std::memset(outHit, 0, sizeof(*outHit));
    }

    __try {
        const auto& addrs = GetNativeSearchAddresses(moduleBase);
        if (addrs.searchMode == 0) {
            return false;
        }
        const int mode = *PtrAbsolute<int>(moduleBase, addrs.searchMode);
        auto* mainWindowObject = reinterpret_cast<unsigned char*>(GetMainWindowObject(moduleBase));
        if (mainWindowObject == nullptr) {
            return false;
        }
        const ptrdiff_t resultRecordOffset =
            mode == 2 ? kOffset_ResultRecordContainerType2 : kOffset_ResultRecordContainerType1;
        auto* containerObject = mainWindowObject + resultRecordOffset;
        const int usedBytes = *reinterpret_cast<int*>(containerObject + kOffset_ByteContainerUsedBytes);
        auto* dataPtr = *reinterpret_cast<unsigned char**>(containerObject + kOffset_ByteContainerData);

        if (usedBytes <= 0 || dataPtr == nullptr) {
            return false;
        }

        const size_t totalCount =
            static_cast<size_t>(usedBytes) / sizeof(e571::DirectGlobalSearch::GlobalSearchHit);
        if (outTotalCount != nullptr) {
            *outTotalCount = totalCount;
        }
        if (totalCount == 0) {
            return false;
        }
        if (outHit == nullptr) {
            return true;
        }
        if (index >= totalCount) {
            return false;
        }

        std::memcpy(
            outHit,
            dataPtr + index * sizeof(e571::DirectGlobalSearch::GlobalSearchHit),
            sizeof(e571::DirectGlobalSearch::GlobalSearchHit));
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        if (outTotalCount != nullptr) {
            *outTotalCount = 0;
        }
        if (outHit != nullptr) {
            std::memset(outHit, 0, sizeof(*outHit));
        }
        return false;
    }
}

bool LooksLikeSearchHit(const e571::DirectGlobalSearch::GlobalSearchHit& hit) {
    switch (hit.type) {
    case 1:
    case 2:
    case 3:
    case 4:
    case 6:
    case 7:
    case 8:
        break;
    default:
        return false;
    }

    if (hit.outerIndex < 0 || hit.innerIndex < 0 || hit.matchOffset < 0) {
        return false;
    }

    if (hit.outerIndex > 1 << 20 || hit.innerIndex > 1 << 20 || hit.matchOffset > 1 << 20) {
        return false;
    }

    if (hit.type != 1 && hit.extra != 0) {
        return false;
    }

    return true;
}

bool TryFormatRawSearchHit(
    std::uintptr_t moduleBase,
    const e571::DirectGlobalSearch::GlobalSearchHit& hit,
    std::string* outText) {
    if (outText == nullptr) {
        return false;
    }
    outText->clear();

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return false;
    }
    const auto initContext = BindAbsolute<FnInitContext>(moduleBase, addrs.initContext);
    const auto resolveBucketIndex = BindAbsolute<FnResolveBucketIndex>(moduleBase, addrs.resolveBucketIndex);

    e571::DirectGlobalSearch::SearchContext ctx{};
    initContext(&ctx);
    ctx.type = hit.type;
    ctx.flag = 0;
    ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.ownerObject)));

    if (hit.type == 1) {
        int bucketData = 0;
        if (!SafeResolveBucketIndex(
                resolveBucketIndex,
                PtrAbsolute<void>(moduleBase, addrs.type1Container),
                hit.extra,
                &bucketData,
                nullptr) ||
            bucketData == 0) {
            return false;
        }
        ctx.data = bucketData;
    } else {
        ctx.data = ResolveTypeDataRaw(hit.type, moduleBase);
        if (ctx.data == 0) {
            return false;
        }
    }

    ExeCStringA text{};
    ExeCStringA prefix{};
    InitExeCString(text, moduleBase);
    InitExeCString(prefix, moduleBase);

    int ok = 0;
    const bool fetchOk = SafeFetchSearchText(
        reinterpret_cast<int(__thiscall*)(
            e571::DirectGlobalSearch::SearchContext*,
            int,
            int,
            ExeCStringA*,
            int,
            int,
            ExeCStringA*,
            int,
            int*)>(BindAbsolute<void*>(moduleBase, addrs.fetchSearchText)),
        &ctx,
        hit.outerIndex,
        hit.innerIndex,
        &text,
        &prefix,
        &ok);
    if (!fetchOk || ok == 0) {
        DestroyExeCString(text, moduleBase);
        DestroyExeCString(prefix, moduleBase);
        return false;
    }

    *outText = BuildDisplayTextRaw(prefix, text);
    DestroyExeCString(text, moduleBase);
    DestroyExeCString(prefix, moduleBase);
    return true;
}

bool InvokeNativeSearchResultRecordConsumer(
    std::uintptr_t moduleBase,
    const e571::DirectGlobalSearch::GlobalSearchHit* hits,
    size_t hitCount,
    size_t hitIndex) {
    if (hits == nullptr || hitCount == 0 || hitIndex >= hitCount) {
        return false;
    }

    static_assert(
        sizeof(e571::DirectGlobalSearch::GlobalSearchHit) == 0x14,
        "GlobalSearchHit must match the native 20-byte search record layout.");
    static_assert(
        offsetof(NativeSearchResultRecordBuffer, data) == kOffset_ByteContainerData,
        "NativeSearchResultRecordBuffer data offset must match the EXE byte container layout.");
    static_assert(
        offsetof(NativeSearchResultRecordBuffer, usedBytes) == kOffset_ByteContainerUsedBytes,
        "NativeSearchResultRecordBuffer usedBytes offset must match the EXE byte container layout.");

    NativeSearchResultRecordBuffer buffer{};
    buffer.data = const_cast<e571::DirectGlobalSearch::GlobalSearchHit*>(hits);
    buffer.capacityBytes = static_cast<int>(hitCount * sizeof(e571::DirectGlobalSearch::GlobalSearchHit));
    buffer.usedBytes = buffer.capacityBytes;

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (addrs.consumeSearchResultRecord == 0) {
        return false;
    }

    __try {
        BindAbsolute<FnConsumeSearchResultRecord>(moduleBase, addrs.consumeSearchResultRecord)(
            &buffer,
            static_cast<int>(hitIndex + 1));
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool TryResolveEditorObjectForHit(
    std::uintptr_t moduleBase,
    const e571::DirectGlobalSearch::GlobalSearchHit& hit,
    void** outEditorObject,
    std::string* outTrace) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return false;
    }
    const auto openCodeTarget = BindAbsolute<FnOpenCodeTarget>(moduleBase, addrs.openCodeTarget);
    void* editorObject = nullptr;

    if (hit.type == 1) {
        int resolvedIndex = -1;
        const int ok = BindAbsolute<FnResolveBucketIndex>(moduleBase, addrs.resolveBucketIndex)(
            PtrAbsolute<void>(moduleBase, addrs.type1Container),
            hit.extra,
            nullptr,
            &resolvedIndex);
        if (ok == 0 || resolvedIndex < 0) {
            if (outTrace != nullptr) {
                *outTrace = "resolve_bucket_failed";
            }
            return false;
        }

        editorObject = reinterpret_cast<void*>(openCodeTarget(
            PtrAbsolute<void>(moduleBase, addrs.mainEditorHost),
            hit.type,
            resolvedIndex,
            0,
            hit.outerIndex,
            hit.innerIndex,
            1,
            -1));
        if (outTrace != nullptr) {
            *outTrace = "open_type1";
        }
    } else {
        editorObject = reinterpret_cast<void*>(openCodeTarget(
            PtrAbsolute<void>(moduleBase, addrs.mainEditorHost),
            hit.type,
            -1,
            -1,
            0,
            0,
            1,
            -1));
        if (outTrace != nullptr) {
            *outTrace = "open_non_type1";
        }
    }

    if (editorObject == nullptr) {
        if (outTrace != nullptr) {
            outTrace->append("_null");
        }
        return false;
    }

    if (outEditorObject != nullptr) {
        *outEditorObject = editorObject;
    }
    return true;
}

bool TryDumpEditorPageCode(
    std::uintptr_t moduleBase,
    void* editorObject,
    std::string* outCode,
    e571::NativeEditorPageDumpDebugResult* outResult) {
    if (editorObject == nullptr || outCode == nullptr) {
        return false;
    }

    outCode->clear();
    if (outResult != nullptr) {
        outResult->ok = false;
        outResult->editorObject = reinterpret_cast<std::uintptr_t>(editorObject);
        outResult->outerCount = 0;
        outResult->lineCount = 0;
        outResult->fetchFailures = 0;
        outResult->trace.clear();
    }

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return false;
    }
    const auto getOuterCount = BindAbsolute<FnEditorGetOuterCount>(moduleBase, addrs.editorGetOuterCount);
    const auto getInnerCount = BindAbsolute<FnEditorGetInnerCount>(moduleBase, addrs.editorGetInnerCount);
    const auto fetchLineText = BindAbsolute<FnEditorFetchLineTextRaw>(moduleBase, addrs.editorFetchLineText);

    int outerCount = 0;
    if (!SafeEditorGetOuterCount(getOuterCount, editorObject, &outerCount)) {
        if (outResult != nullptr) {
            outResult->trace = "outer_count_exception";
        }
        return false;
    }
    if (outResult != nullptr) {
        outResult->outerCount = outerCount;
    }
    if (outerCount <= 0) {
        if (outResult != nullptr) {
            outResult->trace = "outer_count_invalid";
        }
        return false;
    }

    size_t lineCount = 0;
    size_t fetchFailures = 0;
    for (int outerIndex = 0; outerIndex < outerCount; ++outerIndex) {
        int innerCount = 0;
        if (!SafeEditorGetInnerCount(getInnerCount, editorObject, outerIndex, &innerCount)) {
            ++fetchFailures;
            if (outResult != nullptr) {
                outResult->trace = "inner_count_exception outer=" + std::to_string(outerIndex);
            }
            continue;
        }
        if (innerCount <= 0) {
            innerCount = 1;
        }

        for (int innerIndex = 0; innerIndex < innerCount; ++innerIndex) {
            ExeCStringA lineText;
            InitExeCString(lineText, moduleBase);
            int optionalOut = 0;
            int ok = 0;
            if (!SafeEditorFetchLineText(
                    fetchLineText,
                    editorObject,
                    outerIndex,
                    innerIndex,
                    &lineText,
                    &ok,
                    &optionalOut)) {
                DestroyExeCString(lineText, moduleBase);
                ++fetchFailures;
                if (outResult != nullptr && outResult->trace.empty()) {
                    outResult->trace =
                        "fetch_exception outer=" + std::to_string(outerIndex) +
                        " inner=" + std::to_string(innerIndex);
                }
                continue;
            }
            if (ok == 0) {
                DestroyExeCString(lineText, moduleBase);
                ++fetchFailures;
                continue;
            }

            outCode->append(GetExeCStringData(lineText));
            outCode->append("\r\n");
            DestroyExeCString(lineText, moduleBase);
            ++lineCount;
        }
    }

    if (outResult != nullptr) {
        outResult->ok = lineCount != 0;
        outResult->lineCount = lineCount;
        outResult->fetchFailures = fetchFailures;
        outResult->trace = lineCount != 0 ? "enumerate_ok" : "enumerate_empty";
    }
    return lineCount != 0;
}

bool TryDumpSearchContextCode(
    e571::DirectGlobalSearch::SearchContext* ctx,
    std::uintptr_t moduleBase,
    std::string* outCode,
    int* outOuterCount,
    size_t* outLineCount,
    size_t* outFetchFailures) {
    if (ctx == nullptr || outCode == nullptr) {
        return false;
    }

    if (outOuterCount != nullptr) {
        *outOuterCount = 0;
    }
    if (outLineCount != nullptr) {
        *outLineCount = 0;
    }
    if (outFetchFailures != nullptr) {
        *outFetchFailures = 0;
    }
    outCode->clear();

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return false;
    }
    const auto getOuterCount = BindAbsolute<FnGetOuterCount>(moduleBase, addrs.getOuterCount);
    const auto getInnerCount = BindAbsolute<FnGetInnerCount>(moduleBase, addrs.getInnerCount);
    const auto fetchSearchText = BindAbsolute<int(__thiscall*)(
        e571::DirectGlobalSearch::SearchContext*,
        int,
        int,
        ExeCStringA*,
        int,
        int,
        ExeCStringA*,
        int,
        int*)>(moduleBase, addrs.fetchSearchText);

    int outerCount = 0;
    if (!SafeGetOuterCount(getOuterCount, ctx, &outerCount) || outerCount <= 0) {
        return false;
    }
    if (outOuterCount != nullptr) {
        *outOuterCount = outerCount;
    }

    size_t lineCount = 0;
    size_t fetchFailures = 0;
    for (int outerIndex = 0; outerIndex < outerCount; ++outerIndex) {
        int innerCount = 0;
        if (!SafeGetInnerCount(getInnerCount, ctx, outerIndex, &innerCount)) {
            ++fetchFailures;
            continue;
        }
        if (innerCount < 0) {
            innerCount = 1;
        }

        for (int innerIndex = 0; innerIndex < innerCount; ++innerIndex) {
            ExeCStringA text{};
            ExeCStringA prefix{};
            InitExeCString(text, moduleBase);
            InitExeCString(prefix, moduleBase);

            int ok = 0;
            const bool fetchOk = SafeFetchSearchText(fetchSearchText, ctx, outerIndex, innerIndex, &text, &prefix, &ok);
            if (!fetchOk || ok == 0) {
                DestroyExeCString(text, moduleBase);
                DestroyExeCString(prefix, moduleBase);
                ++fetchFailures;
                continue;
            }

            outCode->append(GetExeCStringData(text));
            outCode->append("\r\n");
            DestroyExeCString(text, moduleBase);
            DestroyExeCString(prefix, moduleBase);
            ++lineCount;
        }
    }

    if (outLineCount != nullptr) {
        *outLineCount = lineCount;
    }
    if (outFetchFailures != nullptr) {
        *outFetchFailures = fetchFailures;
    }
    return lineCount != 0;
}

bool TryResolveEditorObjectForProgramTreeItemData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    void** outEditorObject,
    int* outResolvedType,
    int* outResolvedIndex,
    int* outBucketData,
    std::string* outTrace) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (outResolvedType != nullptr) {
        *outResolvedType = 0;
    }
    if (outResolvedIndex != nullptr) {
        *outResolvedIndex = -1;
    }
    if (outBucketData != nullptr) {
        *outBucketData = 0;
    }
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    const unsigned int itemTypeNibble = itemData >> 28;
    if (itemTypeNibble == 0 || itemTypeNibble == 15) {
        if (outTrace != nullptr) {
            *outTrace = "non_page_tree_item";
        }
        return false;
    }

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        if (outTrace != nullptr) {
            *outTrace = "resolve_native_addresses_failed";
        }
        return false;
    }

    const auto openCodeTarget = BindAbsolute<FnOpenCodeTarget>(moduleBase, addrs.openCodeTarget);
    int openType = 0;
    int arg2 = -1;
    int arg3 = -1;
    int resolvedIndex = -1;
    int bucketData = 0;

    switch (itemTypeNibble) {
    case 1: {
        resolvedIndex = static_cast<int>(((itemData >> 16) & 0x0FFFu)) - 1;
        if (resolvedIndex < 0) {
            if (outTrace != nullptr) {
                *outTrace = "type1_resolved_index_invalid";
            }
            return false;
        }

        const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, addrs.containerGetAt);
        int bucketOk = 0;
        if (!SafeContainerGetAt(
                containerGetAt,
                PtrAbsolute<void>(moduleBase, addrs.type1Container),
                resolvedIndex,
                &bucketData,
                &bucketOk) ||
            bucketOk == 0 ||
            bucketData == 0) {
            if (outTrace != nullptr) {
                *outTrace = "type1_bucket_lookup_failed";
            }
            return false;
        }

        openType = 1;
        arg2 = resolvedIndex;
        arg3 = 0;
        break;
    }
    case 2:
        openType = 3;
        break;
    case 3:
        openType = 2;
        break;
    case 4:
        openType = 4;
        break;
    case 5:
        openType = 5;
        break;
    case 6:
        openType = 6;
        break;
    case 7: {
        const unsigned int subType = itemData & 0x0FFFFFFFu;
        openType = (subType == 1u) ? 7 : 8;
        break;
    }
    default:
        if (outTrace != nullptr) {
            *outTrace = "unsupported_tree_item_type";
        }
        return false;
    }

    void* editorObject = reinterpret_cast<void*>(openCodeTarget(
        PtrAbsolute<void>(moduleBase, addrs.mainEditorHost),
        openType,
        arg2,
        arg3,
        0,
        0,
        1,
        -1));
    if (editorObject == nullptr) {
        if (outTrace != nullptr) {
            *outTrace = "open_code_target_null";
        }
        return false;
    }

    if (outEditorObject != nullptr) {
        *outEditorObject = editorObject;
    }
    if (outResolvedType != nullptr) {
        *outResolvedType = openType;
    }
    if (outResolvedIndex != nullptr) {
        *outResolvedIndex = resolvedIndex;
    }
    if (outBucketData != nullptr) {
        *outBucketData = bucketData;
    }
    if (outTrace != nullptr) {
        *outTrace = "open_editor_ok";
    }
    return true;
}

bool SafeEditorGetOuterCount(FnEditorGetOuterCount fn, void* editorObject, int* outValue) {
    if (outValue != nullptr) {
        *outValue = 0;
    }
    __try {
        const int value = fn(editorObject);
        if (outValue != nullptr) {
            *outValue = value;
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeEditorGetInnerCount(FnEditorGetInnerCount fn, void* editorObject, int outerIndex, int* outValue) {
    if (outValue != nullptr) {
        *outValue = 0;
    }
    __try {
        const int value = fn(editorObject, outerIndex);
        if (outValue != nullptr) {
            *outValue = value;
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeEditorFetchLineText(
    FnEditorFetchLineTextRaw fn,
    void* editorObject,
    int outerIndex,
    int innerIndex,
    ExeCStringA* outText,
    int* outOk,
    int* optionalOut) {
    if (outOk != nullptr) {
        *outOk = 0;
    }
    __try {
        const int ok = fn(editorObject, outerIndex, innerIndex, outText, 0, 0, 0, optionalOut);
        if (outOk != nullptr) {
            *outOk = ok;
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

std::string ReadListBoxText(HWND listHwnd, int index) {
    const LRESULT textLen = ::SendMessageA(listHwnd, LB_GETTEXTLEN, static_cast<WPARAM>(index), 0);
    if (textLen == LB_ERR || textLen < 0) {
        return {};
    }

    std::string text(static_cast<size_t>(textLen), '\0');
    if (textLen == 0) {
        return text;
    }

    const LRESULT copied = ::SendMessageA(
        listHwnd,
        LB_GETTEXT,
        static_cast<WPARAM>(index),
        reinterpret_cast<LPARAM>(text.data()));
    if (copied == LB_ERR || copied < 0) {
        return {};
    }

    text.resize(static_cast<size_t>(copied));
    return text;
}

bool IsBuiltinSearchDecorativeLine(const std::string& text) {
    return text.rfind("--- ", 0) == 0;
}

bool IsBuiltinSearchSummaryLine(const std::string& text) {
    return IsBuiltinSearchDecorativeLine(text) && text.find("找到") != std::string::npos;
}

void PumpPendingMessages() {
    MSG msg = {};
    while (::PeekMessageA(&msg, nullptr, 0, 0, PM_REMOVE)) {
        ::TranslateMessage(&msg);
        ::DispatchMessageA(&msg);
    }
}

void WaitForBuiltinSearchResults(std::uintptr_t moduleBase, DWORD timeoutMs) {
    DWORD deadline = ::GetTickCount() + timeoutMs;
    LRESULT lastCount = -1;
    int stableRounds = 0;

    for (;;) {
        PumpPendingMessages();

        HWND listHwnd = GetBuiltinResultListHwnd(moduleBase);
        if (::IsWindow(listHwnd)) {
            const LRESULT rawCount = ::SendMessageA(listHwnd, LB_GETCOUNT, 0, 0);
            if (rawCount != LB_ERR && rawCount >= 2) {
                const std::string lastLine = ReadListBoxText(listHwnd, static_cast<int>(rawCount - 1));
                if (IsBuiltinSearchSummaryLine(lastLine)) {
                    stableRounds = (rawCount == lastCount) ? (stableRounds + 1) : 1;
                    lastCount = rawCount;
                    if (stableRounds >= 3) {
                        break;
                    }
                } else {
                    stableRounds = 0;
                    lastCount = rawCount;
                }
            }
        }

        if (static_cast<DWORD>(::GetTickCount() - deadline) < 0x80000000u) {
            break;
        }

        ::Sleep(20);
    }
}

void WaitForBuiltinSearchResults(HiddenBuiltinSearchContext& ctx, std::uintptr_t moduleBase, DWORD timeoutMs) {
    DWORD deadline = ::GetTickCount() + timeoutMs;
    while (!ctx.searchFinished) {
        PumpPendingMessages();
        if (ctx.searchFinished) {
            break;
        }
        if (ctx.resultListHwnd != nullptr && ::IsWindow(ctx.resultListHwnd)) {
            const LRESULT rawCount = ::SendMessageA(ctx.resultListHwnd, LB_GETCOUNT, 0, 0);
            if (rawCount != LB_ERR && rawCount >= 2) {
                const std::string lastLine = ReadListBoxText(ctx.resultListHwnd, static_cast<int>(rawCount - 1));
                if (IsBuiltinSearchSummaryLine(lastLine)) {
                    ctx.searchFinished = true;
                    break;
                }
            }
        }
        if (static_cast<DWORD>(::GetTickCount() - deadline) < 0x80000000u) {
            break;
        }
        ::Sleep(20);
    }

}

void FillCapturedBuiltinSearchResult(
    const HiddenBuiltinSearchContext& ctx,
    e571::HiddenBuiltinSearchDebugResult& result) {
    if (!ctx.rawHits.empty()) {
        result.hits = ctx.rawHits.size();
        result.rawHitCount = ctx.rawHits.size();
        result.hasRawFirstHit = true;
        result.rawFirstHit.type = ctx.rawHits.front().type;
        result.rawFirstHit.extra = ctx.rawHits.front().extra;
        result.rawFirstHit.outerIndex = ctx.rawHits.front().outerIndex;
        result.rawFirstHit.innerIndex = ctx.rawHits.front().innerIndex;
        result.rawFirstHit.matchOffset = ctx.rawHits.front().matchOffset;

        const size_t formatLimit = (std::min)(ctx.rawHits.size(), static_cast<size_t>(kBuiltinResultPreviewLimit));
        for (size_t i = 0; i < formatLimit; ++i) {
            std::string text;
            if (!TryFormatRawSearchHit(ctx.moduleBase, ctx.rawHits[i], &text) || text.empty()) {
                continue;
            }

            if (result.firstResultText.empty()) {
                result.firstResultText = text;
            }
            result.previewLines.push_back(std::move(text));
        }
        return;
    }

    result.hits = ctx.fallbackHitCount;
    result.firstResultText = ctx.fallbackFirstResultText;
    result.previewLines = ctx.fallbackPreviewLines;
}

int FindFirstBuiltinSearchResultIndex(HWND listHwnd) {
    if (!::IsWindow(listHwnd)) {
        return -1;
    }

    const LRESULT rawCount = ::SendMessageA(listHwnd, LB_GETCOUNT, 0, 0);
    if (rawCount == LB_ERR || rawCount <= 0) {
        return -1;
    }

    for (int i = 0; i < rawCount; ++i) {
        const std::string text = ReadListBoxText(listHwnd, i);
        if (text.empty() || IsBuiltinSearchDecorativeLine(text)) {
            continue;
        }

        if (i == rawCount - 1) {
            continue;
        }

        return i;
    }

    return -1;
}

e571::HiddenBuiltinSearchDebugResult CollectBuiltinSearchResult(std::uintptr_t moduleBase) {
    e571::HiddenBuiltinSearchDebugResult result;
    HWND listHwnd = GetBuiltinResultListHwnd(moduleBase);
    if (!::IsWindow(listHwnd)) {
        return result;
    }

    const LRESULT rawCount = ::SendMessageA(listHwnd, LB_GETCOUNT, 0, 0);
    if (rawCount == LB_ERR || rawCount < 0) {
        return result;
    }

    const int firstResultIndex = FindFirstBuiltinSearchResultIndex(listHwnd);
    if (firstResultIndex >= 0) {
        result.firstResultText = ReadListBoxText(listHwnd, firstResultIndex);

        for (int i = firstResultIndex; i < rawCount - 1; ++i) {
            const std::string text = ReadListBoxText(listHwnd, i);
            if (text.empty() || IsBuiltinSearchDecorativeLine(text)) {
                continue;
            }

            ++result.hits;
            if (result.previewLines.size() < kBuiltinResultPreviewLimit) {
                result.previewLines.push_back(text);
            }
        }
    }

    return result;
}

bool TriggerBuiltinSearchFirstResult(std::uintptr_t moduleBase) {
    HWND listHwnd = GetBuiltinResultListHwnd(moduleBase);
    if (!::IsWindow(listHwnd)) {
        return false;
    }

    const int firstResultIndex = FindFirstBuiltinSearchResultIndex(listHwnd);
    if (firstResultIndex < 0) {
        return false;
    }

    if (::SendMessageA(listHwnd, LB_SETCURSEL, static_cast<WPARAM>(firstResultIndex), 0) == LB_ERR) {
        return false;
    }

    HWND parentHwnd = ::GetParent(listHwnd);
    if (!::IsWindow(parentHwnd)) {
        return false;
    }

    ::SetFocus(listHwnd);
    const int ctrlId = ::GetDlgCtrlID(listHwnd);
    ::SendMessageA(
        parentHwnd,
        WM_COMMAND,
        MAKEWPARAM(ctrlId, LBN_SELCHANGE),
        reinterpret_cast<LPARAM>(listHwnd));
    ::SendMessageA(
        parentHwnd,
        WM_COMMAND,
        MAKEWPARAM(ctrlId, LBN_DBLCLK),
        reinterpret_cast<LPARAM>(listHwnd));
    return true;
}

bool RunBuiltinSearchDialogHidden(
    const char* keyword,
    std::uintptr_t moduleBase,
    bool* outDialogHandled,
    e571::HiddenBuiltinSearchDebugResult* outCapturedResult,
    std::vector<e571::DirectGlobalSearch::GlobalSearchHit>* outRawHits) {
    std::array<unsigned char, kBuiltinSearchDialogStorageSize> dialogStorage = {};
    void* const dialogObject = dialogStorage.data();

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return false;
    }
    auto ctor = BindAbsolute<FnBuiltinSearchDialogCtor>(moduleBase, addrs.builtinSearchDialogCtor);
    auto dtor = BindAbsolute<FnBuiltinSearchDialogDtor>(moduleBase, addrs.builtinSearchDialogDtor);
    auto doModal = BindAbsolute<FnDialogDoModal>(moduleBase, addrs.dialogDoModal);

    HiddenBuiltinSearchContext hookContext;
    hookContext.keyword = keyword;
    hookContext.dialogObject = dialogObject;

    ctor(dialogObject, nullptr);
    {
        ScopedHiddenBuiltinSearchHook hook(hookContext);
        ScopedHiddenBuiltinSearchApiHooks apiHooks(hookContext, moduleBase);
        if (!hook.IsInstalled() || !apiHooks.IsInstalled()) {
            dtor(dialogObject, 0);
            return false;
        }
        ScopedHiddenBuiltinSearchUiGuard uiGuard(hookContext, moduleBase);
        doModal(dialogObject);
        if (!hookContext.rawHits.empty()) {
            DWORD deadline = ::GetTickCount() + 120;
            size_t lastCount = hookContext.rawHits.size();
            int stableRounds = 0;
            while (stableRounds < 2) {
                PumpPendingMessages();
                const size_t currentCount = hookContext.rawHits.size();
                if (currentCount == lastCount) {
                    ++stableRounds;
                } else {
                    lastCount = currentCount;
                    stableRounds = 0;
                }
                if (static_cast<DWORD>(::GetTickCount() - deadline) < 0x80000000u) {
                    break;
                }
                ::Sleep(1);
            }
        } else if (!hookContext.searchFinished) {
            WaitForBuiltinSearchResults(hookContext, moduleBase, 500);
        }
    }
    dtor(dialogObject, 0);

    if (outDialogHandled != nullptr) {
        *outDialogHandled = hookContext.dialogHandled;
    }
    if (outCapturedResult != nullptr) {
        FillCapturedBuiltinSearchResult(hookContext, *outCapturedResult);
    }
    if (outRawHits != nullptr) {
        *outRawHits = hookContext.rawHits;
    }
    return true;
}

bool RunBuiltinSearchDialogHiddenSafe(
    const char* keyword,
    std::uintptr_t moduleBase,
    bool* outDialogHandled,
    e571::HiddenBuiltinSearchDebugResult* outCapturedResult,
    std::vector<e571::DirectGlobalSearch::GlobalSearchHit>* outRawHits) {
    __try {
        return RunBuiltinSearchDialogHidden(keyword, moduleBase, outDialogHandled, outCapturedResult, outRawHits);
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeGetOuterCount(FnGetOuterCount fn, e571::DirectGlobalSearch::SearchContext* ctx, int* outValue) {
    __try {
        *outValue = fn(ctx);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        *outValue = 0;
        return false;
    }
}

bool SafeGetInnerCount(
    FnGetInnerCount fn,
    e571::DirectGlobalSearch::SearchContext* ctx,
    int outerIndex,
    int* outValue) {
    __try {
        *outValue = fn(ctx, outerIndex);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        *outValue = 0;
        return false;
    }
}

bool SafeFetchSearchText(
    int(__thiscall* fn)(
        e571::DirectGlobalSearch::SearchContext*,
        int,
        int,
        ExeCStringA*,
        int,
        int,
        ExeCStringA*,
        int,
        int*),
    e571::DirectGlobalSearch::SearchContext* ctx,
    int outerIndex,
    int innerIndex,
    ExeCStringA* outText,
    ExeCStringA* outPrefix,
    int* outOk) {
    __try {
        *outOk = fn(ctx, outerIndex, innerIndex, outText, 0, 0, outPrefix, 0, nullptr);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        *outOk = 0;
        return false;
    }
}

bool SafeContainerGetAt(FnContainerGetAt fn, void* container, int index, int* outPtr, int* outOk) {
    __try {
        *outOk = fn(container, index, outPtr);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        *outOk = 0;
        return false;
    }
}

bool SafeContainerGetId(FnContainerGetId fn, void* container, int index, int* outId) {
    __try {
        *outId = fn(container, index);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        *outId = 0;
        return false;
    }
}

bool SafeResolveBucketIndex(
    FnResolveBucketIndex fn,
    void* container,
    int bucketId,
    int* outValue,
    int* outPos) {
    __try {
        fn(container, bucketId, outValue, outPos);
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        if (outValue != nullptr) {
            *outValue = 0;
        }
        if (outPos != nullptr) {
            *outPos = -1;
        }
        return false;
    }
}

bool SafeJumpToResult(const e571::DirectGlobalSearch& search, const e571::DirectGlobalSearch::GlobalSearchHit& hit) {
    __try {
        return search.JumpToResult(hit);
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

void InitHiddenBuiltinSearchTargets(HiddenBuiltinSearchContext& ctx, std::uintptr_t moduleBase) {
    ctx.moduleBase = moduleBase;

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return;
    }
    const int mode = *PtrAbsolute<int>(moduleBase, addrs.searchMode);
    auto* mainWindowObject = reinterpret_cast<unsigned char*>(GetMainWindowObject(moduleBase));
    if (mainWindowObject == nullptr) {
        return;
    }
    const ptrdiff_t resultPageOffset = mode == 2 ? kOffset_ResultPageType2 : kOffset_ResultPageType1;
    const ptrdiff_t resultRecordOffset =
        mode == 2 ? kOffset_ResultRecordContainerType2 : kOffset_ResultRecordContainerType1;

    ctx.resultPageObject = mainWindowObject + resultPageOffset;
    ctx.resultContainerObject = mainWindowObject + resultRecordOffset;
    ctx.resultPageHwnd = *reinterpret_cast<HWND*>(reinterpret_cast<unsigned char*>(ctx.resultPageObject) + kOffset_CWndHwnd);
    ctx.mainFrameHwnd = *reinterpret_cast<HWND*>(mainWindowObject + kOffset_CWndHwnd);
}

int __fastcall HiddenBuiltinSearchHook_AppendBytes(void* thisPtr, void* /*edx*/, void* src, int size) {
    HiddenBuiltinSearchContext* ctx = g_hiddenBuiltinSearchContext;
    if (ctx != nullptr &&
        src != nullptr &&
        size == static_cast<int>(sizeof(e571::DirectGlobalSearch::GlobalSearchHit))) {
        e571::DirectGlobalSearch::GlobalSearchHit hit{};
        std::memcpy(&hit, src, sizeof(hit));
        if (LooksLikeSearchHit(hit)) {
            ctx->rawHits.push_back(hit);
            return 1;
        }
    }

    return g_originalAppendBytes(thisPtr, src, size);
}

int __fastcall HiddenBuiltinSearchHook_PrepareSearchResults(void* thisPtr, void* /*edx*/, int mode) {
    HiddenBuiltinSearchContext* ctx = g_hiddenBuiltinSearchContext;
    if (ctx != nullptr) {
        return 0;
    }

    return g_originalPrepareSearchResults(thisPtr, mode);
}

int __fastcall HiddenBuiltinSearchHook_SelectSearchResultTab(void* thisPtr, void* /*edx*/, int mode) {
    return g_originalSelectSearchResultTab(thisPtr, mode);
}

int __fastcall HiddenBuiltinSearchHook_ActivateWindowObject(int thisPtr, void* /*edx*/) {
    HiddenBuiltinSearchContext* ctx = g_hiddenBuiltinSearchContext;
    if (ctx != nullptr && reinterpret_cast<void*>(thisPtr) == ctx->resultPageObject) {
        return thisPtr;
    }

    return g_originalActivateWindowObject(thisPtr);
}

LRESULT WINAPI HiddenBuiltinSearchHook_SendMessageA(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    HiddenBuiltinSearchContext* ctx = g_hiddenBuiltinSearchContext;
    if (ctx != nullptr) {
        char className[32] = {};
        const bool hasClassName =
            ::GetClassNameA(hwnd, className, static_cast<int>(std::size(className))) > 0;
        const bool isTabControl = hasClassName && std::strcmp(className, "SysTabControl32") == 0;

        if (isTabControl) {
            if (msg == kMsg_TcmSetCurSel) {
                return g_originalSendMessageA(hwnd, kMsg_TcmGetCurSel, 0, 0);
            }
            if (msg == kMsg_TcmSetCurFocus) {
                return g_originalSendMessageA(hwnd, kMsg_TcmGetCurFocus, 0, 0);
            }
        }

        if (hwnd == ctx->mainFrameHwnd && msg == WM_COMMAND && LOWORD(wParam) == 0x82) {
            return 0;
        }
        if (hwnd == ctx->resultPageHwnd && (msg == LB_RESETCONTENT || msg == LB_ADDSTRING || msg == LB_SETCOUNT)) {
            if (msg == LB_ADDSTRING) {
                return 0;
            }
            return 0;
        }
    }

    return g_originalSendMessageA(hwnd, msg, wParam, lParam);
}

ScopedHiddenBuiltinSearchApiHooks::ScopedHiddenBuiltinSearchApiHooks(
    HiddenBuiltinSearchContext& ctx,
    std::uintptr_t moduleBase)
    : installed_(false) {
    InitHiddenBuiltinSearchTargets(ctx, moduleBase);

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return;
    }
    g_originalAppendBytes = BindAbsolute<FnAppendBytes>(moduleBase, addrs.appendBytes);
    g_originalPrepareSearchResults = BindAbsolute<FnPrepareSearchResults>(moduleBase, addrs.prepareSearchResults);
    g_originalSelectSearchResultTab = BindAbsolute<FnSelectSearchResultTab>(moduleBase, addrs.selectSearchResultTab);
    g_originalActivateWindowObject = BindAbsolute<FnActivateWindowObject>(moduleBase, addrs.activateWindow);

    DetourTransactionBegin();
    DetourUpdateThread(::GetCurrentThread());
    DetourAttach(&(PVOID&)g_originalAppendBytes, HiddenBuiltinSearchHook_AppendBytes);
    DetourAttach(&(PVOID&)g_originalPrepareSearchResults, HiddenBuiltinSearchHook_PrepareSearchResults);
    DetourAttach(&(PVOID&)g_originalSelectSearchResultTab, HiddenBuiltinSearchHook_SelectSearchResultTab);
    DetourAttach(&(PVOID&)g_originalActivateWindowObject, HiddenBuiltinSearchHook_ActivateWindowObject);
    DetourAttach(&(PVOID&)g_originalSendMessageA, HiddenBuiltinSearchHook_SendMessageA);
    installed_ = (DetourTransactionCommit() == NO_ERROR);
}

ScopedHiddenBuiltinSearchApiHooks::~ScopedHiddenBuiltinSearchApiHooks() {
    if (!installed_) {
        return;
    }

    DetourTransactionBegin();
    DetourUpdateThread(::GetCurrentThread());
    DetourDetach(&(PVOID&)g_originalAppendBytes, HiddenBuiltinSearchHook_AppendBytes);
    DetourDetach(&(PVOID&)g_originalPrepareSearchResults, HiddenBuiltinSearchHook_PrepareSearchResults);
    DetourDetach(&(PVOID&)g_originalSelectSearchResultTab, HiddenBuiltinSearchHook_SelectSearchResultTab);
    DetourDetach(&(PVOID&)g_originalActivateWindowObject, HiddenBuiltinSearchHook_ActivateWindowObject);
    DetourDetach(&(PVOID&)g_originalSendMessageA, HiddenBuiltinSearchHook_SendMessageA);
    DetourTransactionCommit();
}

bool ScopedHiddenBuiltinSearchApiHooks::IsInstalled() const {
    return installed_;
}

ScopedHiddenBuiltinSearchUiGuard::ScopedHiddenBuiltinSearchUiGuard(
    HiddenBuiltinSearchContext& ctx,
    std::uintptr_t moduleBase)
    : ctx_(&ctx),
      moduleBase_(moduleBase),
      redrawDisabled_(false) {
    const auto& addrs = GetNativeSearchAddresses(moduleBase_);
    ctx.previousFocusHwnd = ::GetFocus();
    ctx.previousFocusObject = nullptr;
    if (::IsWindow(ctx.previousFocusHwnd) && addrs.fromHandle != 0) {
        ctx.previousFocusObject = BindAbsolute<FnFromHandle>(moduleBase_, addrs.fromHandle)(ctx.previousFocusHwnd);
    }

    if (::IsWindow(ctx.mainFrameHwnd)) {
        ::SendMessageA(ctx.mainFrameHwnd, WM_SETREDRAW, FALSE, 0);
        redrawDisabled_ = true;
    }
}

ScopedHiddenBuiltinSearchUiGuard::~ScopedHiddenBuiltinSearchUiGuard() {
    if (ctx_ == nullptr) {
        return;
    }

    if (ctx_->previousFocusObject != nullptr) {
        __try {
            BindAbsolute<FnActivateWindowObject>(moduleBase_, GetNativeSearchAddresses(moduleBase_).activateWindow)(
                static_cast<int>(reinterpret_cast<std::uintptr_t>(ctx_->previousFocusObject)));
        } __except (EXCEPTION_EXECUTE_HANDLER) {
        }
    } else if (::IsWindow(ctx_->previousFocusHwnd)) {
        ::SetFocus(ctx_->previousFocusHwnd);
    }

    if (redrawDisabled_ && ::IsWindow(ctx_->mainFrameHwnd)) {
        ::SendMessageA(ctx_->mainFrameHwnd, WM_SETREDRAW, TRUE, 0);
        ::RedrawWindow(
            ctx_->mainFrameHwnd,
            nullptr,
            nullptr,
            RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
    }
}

int ResolveTypeDataRaw(int type, std::uintptr_t moduleBase) {
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    switch (type) {
    case 2:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.type2Data)));
    case 3:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.type3Data)));
    case 4:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.type4Data)));
    case 6:
    case 7:
    case 8:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.type678Data)));
    default:
        return 0;
    }
}

void CollectMatchesRaw(
    int type,
    int extra,
    int outerIndex,
    int innerIndex,
    const char* keyword,
    std::size_t keywordLen,
    bool matchCase,
    const ExeCStringA& prefix,
    const ExeCStringA& text,
    std::vector<e571::DirectGlobalSearchDebugHit>& results) {
    const char* haystack = GetExeCStringData(text);
    if (haystack[0] == '\0') {
        return;
    }

    const std::size_t haystackLen = std::strlen(haystack);
    if (haystackLen < keywordLen) {
        return;
    }

    const std::string displayText = BuildDisplayTextRaw(prefix, text);
    for (std::size_t offset = 0; offset + keywordLen <= haystackLen; ++offset) {
        if (!MatchAtMbcs(haystack + offset, keyword, keywordLen, matchCase)) {
            continue;
        }

        e571::DirectGlobalSearchDebugHit hit;
        hit.type = type;
        hit.extra = extra;
        hit.outerIndex = outerIndex;
        hit.innerIndex = innerIndex;
        hit.matchOffset = static_cast<int>(offset);
        hit.displayText = displayText;
        results.push_back(std::move(hit));
    }
}

void SearchOneContextRaw(
    e571::DirectGlobalSearch::SearchContext& ctx,
    const char* keyword,
    std::size_t keywordLen,
    bool matchCase,
    int extra,
    std::uintptr_t moduleBase,
    std::vector<e571::DirectGlobalSearchDebugHit>& results) {
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        return;
    }
    const auto getOuterCount = BindAbsolute<FnGetOuterCount>(moduleBase, addrs.getOuterCount);
    const auto getInnerCount = BindAbsolute<FnGetInnerCount>(moduleBase, addrs.getInnerCount);
    const auto fetchSearchText = BindAbsolute<int(__thiscall*)(
        e571::DirectGlobalSearch::SearchContext*,
        int,
        int,
        ExeCStringA*,
        int,
        int,
        ExeCStringA*,
        int,
        int*)>(moduleBase, addrs.fetchSearchText);

    int outerCount = 0;
    if (!SafeGetOuterCount(getOuterCount, &ctx, &outerCount)) {
        return;
    }

    for (int outerIndex = 0; outerIndex < outerCount; ++outerIndex) {
        int innerCount = 0;
        if (!SafeGetInnerCount(getInnerCount, &ctx, outerIndex, &innerCount)) {
            continue;
        }
        if (innerCount < 0) {
            innerCount = 1;
        }

        for (int innerIndex = 0; innerIndex < innerCount; ++innerIndex) {
            ExeCStringA text{};
            ExeCStringA prefix{};
            InitExeCString(text, moduleBase);
            InitExeCString(prefix, moduleBase);

            int ok = 0;
            SafeFetchSearchText(fetchSearchText, &ctx, outerIndex, innerIndex, &text, &prefix, &ok);
            if (ok != 0) {
                CollectMatchesRaw(
                    ctx.type,
                    extra,
                    outerIndex,
                    innerIndex,
                    keyword,
                    keywordLen,
                    matchCase,
                    prefix,
                    text,
                    results);
            }

            DestroyExeCString(text, moduleBase);
            DestroyExeCString(prefix, moduleBase);
        }
    }
}

}  // namespace

struct DirectGlobalSearch::DwordContainer {
    int unused0;
    int unused1;
    int unused2;
    int unused3;
};

DirectGlobalSearch::DirectGlobalSearch(std::uintptr_t moduleBase)
    : moduleBase_(moduleBase) {}

std::vector<DirectGlobalSearch::SearchResult> DirectGlobalSearch::Search(
    const char* keyword,
    const SearchOptions& options) const {
    std::vector<SearchResult> results;
    if (keyword == nullptr || keyword[0] == '\0') {
        return results;
    }

    const std::size_t keywordLen = std::strlen(keyword);
    if (keywordLen == 0) {
        return results;
    }
    if (!HasNativeSearchAddresses(moduleBase_)) {
        return results;
    }

    for (int type : kSearchTypes) {
        if (type == 1) {
            EnumerateType1(keyword, keywordLen, options, results);
        } else {
            SearchContext ctx{};
            InitContext(&ctx);

            ctx.type = type;
            ctx.data = ResolveTypeData(type);
            ctx.flag = 0;
            ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(GetNativeSearchAddresses(moduleBase_).ownerObject)));

            SearchOneContext(ctx, keyword, keywordLen, options, 0, results);
        }
    }

    return results;
}

bool DirectGlobalSearch::JumpToResult(const GlobalSearchHit& hit) const {
    if (InvokeNativeSearchResultRecordConsumer(moduleBase_, &hit, 1, 0)) {
        return true;
    }

    if (hit.type == 1) {
        int resolvedIndex = -1;
        const int ok = ResolveBucketIndex(
            Ptr<DwordContainer>(GetNativeSearchAddresses(moduleBase_).type1Container),
            hit.extra,
            nullptr,
            &resolvedIndex);
        if (ok == 0) {
            NotifyOpenFailure(0x30);
            return false;
        }

        HWND hwnd = OpenCodeTarget(
            Ptr<void>(GetNativeSearchAddresses(moduleBase_).mainEditorHost),
            hit.type,
            resolvedIndex,
            0,
            hit.outerIndex,
            hit.innerIndex,
            1,
            -1);
        if (hwnd == nullptr) {
            NotifyOpenFailure(0x30);
            return false;
        }

        return FocusResult(hwnd, hit);
    }

    HWND hwnd = OpenCodeTarget(
        Ptr<void>(GetNativeSearchAddresses(moduleBase_).mainEditorHost),
        hit.type,
        -1,
        -1,
        0,
        0,
        1,
        -1);
    if (hwnd == nullptr) {
        NotifyOpenFailure(0x30);
        return false;
    }

    MoveToLine(hwnd, hit.outerIndex, hit.innerIndex, 0, 1, 0);
    EnsureVisible(hwnd, 0, 0, 1);
    return FocusResult(hwnd, hit);
}

bool DirectGlobalSearch::JumpToResult(const SearchResult& result) const {
    return JumpToResult(result.hit);
}

template <typename T>
T DirectGlobalSearch::Bind(std::uintptr_t absoluteAddress) const {
    return reinterpret_cast<T>(absoluteAddress - kImageBase + moduleBase_);
}

template <typename T>
T* DirectGlobalSearch::Ptr(std::uintptr_t absoluteAddress) const {
    return reinterpret_cast<T*>(absoluteAddress - kImageBase + moduleBase_);
}

void DirectGlobalSearch::InitContext(SearchContext* ctx) const {
    Bind<FnInitContext>(GetNativeSearchAddresses(moduleBase_).initContext)(ctx);
}

int DirectGlobalSearch::GetOuterCount(SearchContext* ctx) const {
    return Bind<FnGetOuterCount>(GetNativeSearchAddresses(moduleBase_).getOuterCount)(ctx);
}

int DirectGlobalSearch::GetInnerCount(SearchContext* ctx, int outerIndex) const {
    return Bind<FnGetInnerCount>(GetNativeSearchAddresses(moduleBase_).getInnerCount)(ctx, outerIndex);
}

int DirectGlobalSearch::FetchSearchText(
    SearchContext* ctx,
    int outerIndex,
    int innerIndex,
    CStringA* outText,
    CStringA* outPrefix,
    int filter,
    int* optionalOut) const {
    return Bind<FnFetchSearchText>(GetNativeSearchAddresses(moduleBase_).fetchSearchText)(
        ctx,
        outerIndex,
        innerIndex,
        outText,
        0,
        0,
        outPrefix,
        filter,
        optionalOut);
}

int DirectGlobalSearch::ContainerGetAt(void* container, int index, int* outPtr) const {
    return Bind<FnContainerGetAt>(GetNativeSearchAddresses(moduleBase_).containerGetAt)(container, index, outPtr);
}

int DirectGlobalSearch::ContainerGetId(void* container, int index) const {
    return Bind<FnContainerGetId>(GetNativeSearchAddresses(moduleBase_).containerGetId)(container, index);
}

int DirectGlobalSearch::ResolveBucketIndex(void* container, int bucketId, int* outValue, int* outPos) const {
    return Bind<FnResolveBucketIndex>(GetNativeSearchAddresses(moduleBase_).resolveBucketIndex)(container, bucketId, outValue, outPos);
}

HWND DirectGlobalSearch::OpenCodeTarget(
    void* mainEditorHost,
    int type,
    int arg2,
    int arg3,
    int outerIndex,
    int innerIndex,
    int activate,
    int arg7) const {
    return Bind<FnOpenCodeTarget>(GetNativeSearchAddresses(moduleBase_).openCodeTarget)(
        mainEditorHost,
        type,
        arg2,
        arg3,
        outerIndex,
        innerIndex,
        activate,
        arg7);
}

void DirectGlobalSearch::MoveToLine(HWND hwnd, int outerIndex, int innerIndex, int arg3, int arg4, int arg5) const {
    Bind<FnMoveToLine>(GetNativeSearchAddresses(moduleBase_).moveToLine)(hwnd, outerIndex, innerIndex, arg3, arg4, arg5);
}

void DirectGlobalSearch::EnsureVisible(HWND hwnd, int arg1, int arg2, int arg3) const {
    Bind<FnEnsureVisible>(GetNativeSearchAddresses(moduleBase_).ensureVisible)(hwnd, arg1, arg2, arg3);
}

void DirectGlobalSearch::MoveCaretToOffset(HWND hwnd, int matchOffset, int force, void* maybeNull, int redraw) const {
    Bind<FnMoveCaretToOffset>(GetNativeSearchAddresses(moduleBase_).moveCaretToOffset)(hwnd, matchOffset, force, maybeNull, redraw);
}

void DirectGlobalSearch::ActivateWindow(HWND hwnd) const {
    Bind<FnActivateWindow>(GetNativeSearchAddresses(moduleBase_).activateWindow)(hwnd);
}

void DirectGlobalSearch::NotifyOpenFailure(int messageId) const {
    const auto& addrs = GetNativeSearchAddresses(moduleBase_);
    Bind<FnNotifyOpenFailure>(addrs.notifyOpenFailure)(Ptr<void>(addrs.mainEditorHost), messageId);
}

int DirectGlobalSearch::ResolveTypeData(int type) const {
    const auto& addrs = GetNativeSearchAddresses(moduleBase_);
    switch (type) {
    case 2:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(addrs.type2Data)));
    case 3:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(addrs.type3Data)));
    case 4:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(addrs.type4Data)));
    case 6:
    case 7:
    case 8:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(addrs.type678Data)));
    default:
        return 0;
    }
}

void DirectGlobalSearch::EnumerateType1(
    const char* keyword,
    std::size_t keywordLen,
    const SearchOptions& options,
    std::vector<SearchResult>& results) const {
    const auto& addrs = GetNativeSearchAddresses(moduleBase_);
    for (int bucketIndex = 0;; ++bucketIndex) {
        SearchContext ctx{};
        InitContext(&ctx);

        int bucketData = 0;
        if (ContainerGetAt(Ptr<DwordContainer>(addrs.type1Container), bucketIndex, &bucketData) == 0) {
            break;
        }

        ctx.type = 1;
        ctx.data = bucketData;
        ctx.flag = 0;
        ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(addrs.ownerObject)));

        const int bucketId = ContainerGetId(Ptr<DwordContainer>(addrs.type1Container), bucketIndex);
        SearchOneContext(ctx, keyword, keywordLen, options, bucketId, results);
    }
}

void DirectGlobalSearch::SearchOneContext(
    SearchContext& ctx,
    const char* keyword,
    std::size_t keywordLen,
    const SearchOptions& options,
    int extra,
    std::vector<SearchResult>& results) const {
    const int outerCount = GetOuterCount(&ctx);
    for (int outerIndex = 0; outerIndex < outerCount; ++outerIndex) {
        int innerCount = GetInnerCount(&ctx, outerIndex);
        if (innerCount < 0) {
            innerCount = 1;
        }

        for (int innerIndex = 0; innerIndex < innerCount; ++innerIndex) {
            CStringA text;
            CStringA prefix;
            const int ok = FetchSearchText(&ctx, outerIndex, innerIndex, &text, &prefix, 0, nullptr);
            if (ok == 0) {
                continue;
            }

            CollectMatches(ctx.type, extra, outerIndex, innerIndex, keyword, keywordLen, options, prefix, text, results);
        }
    }
}

void DirectGlobalSearch::CollectMatches(
    int type,
    int extra,
    int outerIndex,
    int innerIndex,
    const char* keyword,
    std::size_t keywordLen,
    const SearchOptions& options,
    const CStringA& prefix,
    const CStringA& text,
    std::vector<SearchResult>& results) const {
    const char* haystack = text.GetString();
    if (haystack == nullptr || haystack[0] == '\0') {
        return;
    }

    const std::size_t haystackLen = std::strlen(haystack);
    if (haystackLen < keywordLen) {
        return;
    }

    for (std::size_t offset = 0; offset + keywordLen <= haystackLen; ++offset) {
        if (!MatchAt(haystack + offset, keyword, keywordLen, options.matchCase)) {
            continue;
        }

        SearchResult result{};
        result.hit.type = type;
        result.hit.extra = extra;
        result.hit.outerIndex = outerIndex;
        result.hit.innerIndex = innerIndex;
        result.hit.matchOffset = static_cast<int>(offset);
        result.prefix = prefix;
        result.text = text;
        result.displayText = BuildDisplayText(prefix, text);
        results.push_back(result);
    }
}

bool DirectGlobalSearch::MatchAt(const char* haystack, const char* needle, std::size_t count, bool matchCase) const {
    if (matchCase) {
        return ::_mbsnbcmp(
                   reinterpret_cast<const unsigned char*>(haystack),
                   reinterpret_cast<const unsigned char*>(needle),
                   static_cast<unsigned int>(count)) == 0;
    }

    return ::_mbsnbicmp(
               reinterpret_cast<const unsigned char*>(haystack),
               reinterpret_cast<const unsigned char*>(needle),
               static_cast<unsigned int>(count)) == 0;
}

CStringA DirectGlobalSearch::BuildDisplayText(const CStringA& prefix, const CStringA& text) const {
    if (prefix.IsEmpty()) {
        return text;
    }

    CStringA display(prefix);
    display += " ";
    display += text;
    return display;
}

bool DirectGlobalSearch::FocusResult(HWND hwnd, const GlobalSearchHit& hit) const {
    if (hwnd == nullptr) {
        return false;
    }

    auto* state = reinterpret_cast<int*>(hwnd);
    if (state[29] != hit.outerIndex || state[30] != hit.innerIndex) {
        NotifyOpenFailure(0x30);
        return false;
    }

    MoveCaretToOffset(hwnd, hit.matchOffset, 0, nullptr, 1);
    if (state[32] != hit.matchOffset) {
        NotifyOpenFailure(0x30);
        return false;
    }

    ActivateWindow(hwnd);
    return true;
}

}  // namespace e571

std::vector<e571::DirectGlobalSearchDebugHit> e571::DebugSearchDirectGlobalKeyword(
    const char* keyword,
    std::uintptr_t moduleBase) {
    std::vector<DirectGlobalSearchDebugHit> debugHits;
    if (keyword == nullptr || keyword[0] == '\0') {
        return debugHits;
    }

    const std::size_t keywordLen = std::strlen(keyword);
    if (keywordLen == 0) {
        return debugHits;
    }
    if (!HasNativeSearchAddresses(moduleBase)) {
        return debugHits;
    }

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    const auto initContext = BindAbsolute<FnInitContext>(moduleBase, addrs.initContext);
    const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, addrs.containerGetAt);
    const auto containerGetId = BindAbsolute<FnContainerGetId>(moduleBase, addrs.containerGetId);

    for (int type : kSearchTypes) {
        if (type == 1) {
            for (int bucketIndex = 0;; ++bucketIndex) {
                DirectGlobalSearch::SearchContext ctx{};
                initContext(&ctx);

                int bucketData = 0;
                int bucketOk = 0;
                if (!SafeContainerGetAt(
                        containerGetAt,
                        PtrAbsolute<void>(moduleBase, addrs.type1Container),
                        bucketIndex,
                        &bucketData,
                        &bucketOk)) {
                    continue;
                }
                if (bucketOk == 0) {
                    break;
                }
                if ((reinterpret_cast<const Type1BucketEntry*>(bucketData)->flags & 0x02) != 0) {
                    continue;
                }

                ctx.type = 1;
                ctx.data = bucketData;
                ctx.flag = 0;
                ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.ownerObject)));

                int bucketId = 0;
                if (!SafeContainerGetId(
                        containerGetId,
                        PtrAbsolute<void>(moduleBase, addrs.type1Container),
                        bucketIndex,
                        &bucketId)) {
                    continue;
                }
                SearchOneContextRaw(ctx, keyword, keywordLen, false, bucketId, moduleBase, debugHits);
            }
            continue;
        }

        DirectGlobalSearch::SearchContext ctx{};
        initContext(&ctx);
        ctx.type = type;
        ctx.data = ResolveTypeDataRaw(type, moduleBase);
        ctx.flag = 0;
        ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.ownerObject)));
        SearchOneContextRaw(ctx, keyword, keywordLen, false, 0, moduleBase, debugHits);
    }

    return debugHits;
}

bool e571::DebugLocateFirstDirectGlobalKeyword(
    const char* keyword,
    std::uintptr_t moduleBase,
    DirectGlobalSearchDebugHit* outHit,
    size_t* outTotalHits) {
    const auto results = DebugSearchDirectGlobalKeyword(keyword, moduleBase);
    if (outTotalHits != nullptr) {
        *outTotalHits = results.size();
    }

    if (results.empty()) {
        return false;
    }

    if (outHit != nullptr) {
        *outHit = results.front();
    }

    DirectGlobalSearch search(moduleBase);
    DirectGlobalSearch::GlobalSearchHit hit{};
    hit.type = results.front().type;
    hit.extra = results.front().extra;
    hit.outerIndex = results.front().outerIndex;
    hit.innerIndex = results.front().innerIndex;
    hit.matchOffset = results.front().matchOffset;
    return SafeJumpToResult(search, hit);
}

e571::HiddenBuiltinSearchDebugResult e571::DebugSearchDirectGlobalKeywordHidden(
    const char* keyword,
    std::uintptr_t moduleBase) {
    HiddenBuiltinSearchDebugResult result;
    if (keyword == nullptr || keyword[0] == '\0') {
        return result;
    }

    bool dialogHandled = false;
    result = {};
    if (!RunBuiltinSearchDialogHiddenSafe(keyword, moduleBase, &dialogHandled, &result, nullptr)) {
        return result;
    }

    if (result.hits == 0) {
        result = CollectBuiltinSearchResult(moduleBase);
    }

    result.dialogHandled = dialogHandled;
    return result;
}

std::vector<e571::DirectGlobalSearchDebugHit> e571::DebugSearchDirectGlobalKeywordHiddenDetailed(
    const char* keyword,
    std::uintptr_t moduleBase,
    bool* outDialogHandled) {
    std::vector<DirectGlobalSearchDebugHit> hits;
    if (outDialogHandled != nullptr) {
        *outDialogHandled = false;
    }
    if (keyword == nullptr || keyword[0] == '\0') {
        return hits;
    }

    bool dialogHandled = false;
    HiddenBuiltinSearchDebugResult captured;
    std::vector<DirectGlobalSearch::GlobalSearchHit> rawHits;
    if (!RunBuiltinSearchDialogHiddenSafe(keyword, moduleBase, &dialogHandled, &captured, &rawHits)) {
        return hits;
    }
    if (outDialogHandled != nullptr) {
        *outDialogHandled = dialogHandled;
    }

    for (const auto& rawHit : rawHits) {
        DirectGlobalSearchDebugHit hit{};
        hit.type = rawHit.type;
        hit.extra = rawHit.extra;
        hit.outerIndex = rawHit.outerIndex;
        hit.innerIndex = rawHit.innerIndex;
        hit.matchOffset = rawHit.matchOffset;
        TryFormatRawSearchHit(moduleBase, rawHit, &hit.displayText);
        hits.push_back(std::move(hit));
    }
    if (!hits.empty()) {
        return hits;
    }

    HWND listHwnd = GetBuiltinResultListHwnd(moduleBase);
    size_t rawCount = 0;
    for (size_t index = 0;; ++index) {
        DirectGlobalSearch::GlobalSearchHit rawHit{};
        size_t totalCount = 0;
        if (!TryReadBuiltinResultHit(moduleBase, index, &rawHit, &totalCount)) {
            rawCount = totalCount;
            break;
        }
        rawCount = totalCount;

        DirectGlobalSearchDebugHit hit{};
        hit.type = rawHit.type;
        hit.extra = rawHit.extra;
        hit.outerIndex = rawHit.outerIndex;
        hit.innerIndex = rawHit.innerIndex;
        hit.matchOffset = rawHit.matchOffset;
        if (!TryFormatRawSearchHit(moduleBase, rawHit, &hit.displayText) || hit.displayText.empty()) {
            if (::IsWindow(listHwnd)) {
                const int listIndex = static_cast<int>(index);
                const LRESULT count = ::SendMessageA(listHwnd, LB_GETCOUNT, 0, 0);
                if (count != LB_ERR && listIndex >= 0 && listIndex < count) {
                    hit.displayText = ReadListBoxText(listHwnd, listIndex);
                }
            }
        }
        hits.push_back(std::move(hit));
    }

    if (hits.empty() && rawCount == 0) {
        return DebugSearchDirectGlobalKeyword(keyword, moduleBase);
    }
    return hits;
}

bool e571::DebugLocateFirstDirectGlobalKeywordHidden(
    const char* keyword,
    std::uintptr_t moduleBase,
    HiddenBuiltinSearchDebugResult* outResult) {
    HiddenBuiltinSearchDebugResult result = DebugSearchDirectGlobalKeywordHidden(keyword, moduleBase);
    const auto directHits = DebugSearchDirectGlobalKeyword(keyword, moduleBase);
    result.directHitCount = directHits.size();
    if (!directHits.empty()) {
        result.hasDirectFirstHit = true;
        result.directFirstHit = directHits.front();
    }

    DirectGlobalSearch::GlobalSearchHit firstHit{};
    if (result.hasRawFirstHit) {
        firstHit.type = result.rawFirstHit.type;
        firstHit.extra = result.rawFirstHit.extra;
        firstHit.outerIndex = result.rawFirstHit.outerIndex;
        firstHit.innerIndex = result.rawFirstHit.innerIndex;
        firstHit.matchOffset = result.rawFirstHit.matchOffset;
    } else {
        size_t rawHitCount = 0;
        if (TryReadBuiltinResultHit(moduleBase, 0, &firstHit, &rawHitCount)) {
            result.rawHitCount = rawHitCount;
            result.hasRawFirstHit = true;
            result.rawFirstHit.type = firstHit.type;
            result.rawFirstHit.extra = firstHit.extra;
            result.rawFirstHit.outerIndex = firstHit.outerIndex;
            result.rawFirstHit.innerIndex = firstHit.innerIndex;
            result.rawFirstHit.matchOffset = firstHit.matchOffset;
        }
    }

    if (result.hasRawFirstHit && firstHit.type == 1) {
        int bucketData = 0;
        int resolvedIndex = -1;
        const auto& addrs = GetNativeSearchAddresses(moduleBase);
        const int ok = BindAbsolute<FnResolveBucketIndex>(moduleBase, addrs.resolveBucketIndex)(
            PtrAbsolute<void>(moduleBase, addrs.type1Container),
            firstHit.extra,
            &bucketData,
            &resolvedIndex);
        result.rawResolveOk = ok != 0;
        result.rawBucketData = bucketData;
        result.rawResolvedIndex = resolvedIndex;
    }

    if (outResult != nullptr) {
        *outResult = result;
    }

    if (result.hits == 0) {
        return false;
    }

    if (result.hasRawFirstHit) {
        const bool ok = InvokeNativeSearchResultRecordConsumer(moduleBase, &firstHit, 1, 0);
        result.jumpTrace = ok ? "native_consumer_ok" : "native_consumer_fail";
        if (outResult != nullptr) {
            *outResult = result;
        }
        return ok;
    }

    return false;
}

bool e571::DebugDumpCodePageForSearchHit(
    const DirectGlobalSearchDebugHit& hit,
    std::uintptr_t moduleBase,
    std::string* outCode,
    NativeEditorPageDumpDebugResult* outResult) {
    if (outCode != nullptr) {
        outCode->clear();
    }
    if (outResult != nullptr) {
        *outResult = {};
    }

    DirectGlobalSearch::GlobalSearchHit rawHit{};
    rawHit.type = hit.type;
    rawHit.extra = hit.extra;
    rawHit.outerIndex = hit.outerIndex;
    rawHit.innerIndex = hit.innerIndex;
    rawHit.matchOffset = hit.matchOffset;

    void* editorObject = nullptr;
    std::string resolveTrace;
    if (!TryResolveEditorObjectForHit(moduleBase, rawHit, &editorObject, &resolveTrace)) {
        if (outResult != nullptr) {
            outResult->trace = resolveTrace.empty() ? "resolve_editor_failed" : resolveTrace;
        }
        return false;
    }

    NativeEditorPageDumpDebugResult dumpResult{};
    const bool ok = TryDumpEditorPageCode(moduleBase, editorObject, outCode, &dumpResult);
    dumpResult.editorObject = reinterpret_cast<std::uintptr_t>(editorObject);
    if (!resolveTrace.empty()) {
        dumpResult.trace = resolveTrace + "|" + dumpResult.trace;
    }
    if (outResult != nullptr) {
        *outResult = dumpResult;
    }
    return ok;
}

bool e571::DebugDumpCodePageByProgramTreeItemData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::string* outCode,
    RawSearchContextPageDumpDebugResult* outResult) {
    if (outCode != nullptr) {
        outCode->clear();
    }
    if (outResult != nullptr) {
        *outResult = {};
    }

    const unsigned int itemTypeNibble = itemData >> 28;
    if (itemTypeNibble == 0 || itemTypeNibble == 15) {
        if (outResult != nullptr) {
            outResult->trace = "non_page_tree_item";
        }
        return false;
    }
    if (!HasNativeSearchAddresses(moduleBase)) {
        if (outResult != nullptr) {
            outResult->trace = "resolve_native_addresses_failed";
        }
        return false;
    }

    DirectGlobalSearch::SearchContext ctx{};
    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    BindAbsolute<FnInitContext>(moduleBase, addrs.initContext)(&ctx);
    ctx.flag = 0;
    ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, addrs.ownerObject)));

    RawSearchContextPageDumpDebugResult dumpResult{};
    dumpResult.resolvedIndex = -1;
    dumpResult.bucketData = 0;

    switch (itemTypeNibble) {
    case 1: {
        const int resolvedIndex = static_cast<int>(((itemData >> 16) & 0x0FFFu)) - 1;
        if (resolvedIndex < 0) {
            dumpResult.trace = "type1_resolved_index_invalid";
            if (outResult != nullptr) {
                *outResult = dumpResult;
            }
            return false;
        }

        const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, addrs.containerGetAt);
        int bucketData = 0;
        int bucketOk = 0;
        if (!SafeContainerGetAt(
                containerGetAt,
                PtrAbsolute<void>(moduleBase, addrs.type1Container),
                resolvedIndex,
                &bucketData,
                &bucketOk) ||
            bucketOk == 0 ||
            bucketData == 0) {
            dumpResult.trace = "type1_bucket_lookup_failed";
            dumpResult.resolvedIndex = resolvedIndex;
            if (outResult != nullptr) {
                *outResult = dumpResult;
            }
            return false;
        }

        ctx.type = 1;
        ctx.data = bucketData;
        dumpResult.type = 1;
        dumpResult.resolvedIndex = resolvedIndex;
        dumpResult.bucketData = bucketData;
        break;
    }
    case 2:
        ctx.type = 3;
        ctx.data = ResolveTypeDataRaw(3, moduleBase);
        dumpResult.type = 3;
        break;
    case 3:
        ctx.type = 2;
        ctx.data = ResolveTypeDataRaw(2, moduleBase);
        dumpResult.type = 2;
        break;
    case 4:
        ctx.type = 4;
        ctx.data = ResolveTypeDataRaw(4, moduleBase);
        dumpResult.type = 4;
        break;
    case 6:
        ctx.type = 6;
        ctx.data = ResolveTypeDataRaw(6, moduleBase);
        dumpResult.type = 6;
        break;
    case 7: {
        const unsigned int subType = itemData & 0x0FFFFFFFu;
        ctx.type = (subType == 1u) ? 7 : 8;
        ctx.data = ResolveTypeDataRaw(ctx.type, moduleBase);
        dumpResult.type = ctx.type;
        break;
    }
    default:
        dumpResult.trace = "unsupported_tree_item_type";
        if (outResult != nullptr) {
            *outResult = dumpResult;
        }
        return false;
    }

    if (ctx.data == 0) {
        dumpResult.trace = "search_context_data_null";
        if (outResult != nullptr) {
            *outResult = dumpResult;
        }
        return false;
    }

    dumpResult.ok = TryDumpSearchContextCode(
        &ctx,
        moduleBase,
        outCode,
        &dumpResult.outerCount,
        &dumpResult.lineCount,
        &dumpResult.fetchFailures);
    if (dumpResult.ok) {
        dumpResult.trace = "search_context_ok";
    } else {
        void* editorObject = nullptr;
        int resolvedType = dumpResult.type;
        int resolvedIndex = dumpResult.resolvedIndex;
        int bucketData = dumpResult.bucketData;
        std::string fallbackTrace;
        if (TryResolveEditorObjectForProgramTreeItemData(
                itemData,
                moduleBase,
                &editorObject,
                &resolvedType,
                &resolvedIndex,
                &bucketData,
                &fallbackTrace)) {
            NativeEditorPageDumpDebugResult editorDumpResult{};
            if (TryDumpEditorPageCode(moduleBase, editorObject, outCode, &editorDumpResult)) {
                dumpResult.ok = true;
                dumpResult.type = resolvedType;
                dumpResult.resolvedIndex = resolvedIndex;
                dumpResult.bucketData = bucketData;
                dumpResult.outerCount = editorDumpResult.outerCount;
                dumpResult.lineCount = editorDumpResult.lineCount;
                dumpResult.fetchFailures = editorDumpResult.fetchFailures;
                dumpResult.trace = "search_context_dump_failed|" + fallbackTrace + "|" + editorDumpResult.trace;
            } else {
                dumpResult.trace =
                    "search_context_dump_failed|" + fallbackTrace + "|" + editorDumpResult.trace;
            }
        } else {
            dumpResult.trace = "search_context_dump_failed|" + fallbackTrace;
        }
    }
    if (outResult != nullptr) {
        *outResult = dumpResult;
    }
    return dumpResult.ok;
}

bool e571::DebugOpenProgramTreeItemByData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::string* outTrace) {
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    void* editorObject = nullptr;
    int resolvedType = 0;
    int resolvedIndex = -1;
    int bucketData = 0;
    std::string trace;
    const bool ok = TryResolveEditorObjectForProgramTreeItemData(
        itemData,
        moduleBase,
        &editorObject,
        &resolvedType,
        &resolvedIndex,
        &bucketData,
        &trace);
    if (outTrace != nullptr) {
        *outTrace = trace;
    }
    return ok && editorObject != nullptr;
}

bool e571::DebugJumpToSearchHit(
    const DirectGlobalSearchDebugHit& hit,
    std::uintptr_t moduleBase,
    std::string* outTrace) {
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    DirectGlobalSearch search(moduleBase);
    DirectGlobalSearch::GlobalSearchHit rawHit{};
    rawHit.type = hit.type;
    rawHit.extra = hit.extra;
    rawHit.outerIndex = hit.outerIndex;
    rawHit.innerIndex = hit.innerIndex;
    rawHit.matchOffset = hit.matchOffset;

    const bool ok = SafeJumpToResult(search, rawHit);
    if (outTrace != nullptr) {
        *outTrace = ok ? "jump_ok" : "jump_failed";
    }
    return ok;
}
