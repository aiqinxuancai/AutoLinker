#include "direct_global_search.hpp"
#include "direct_global_search_debug.hpp"

#include <detours.h>

#include <array>
#include <mbstring.h>

#include <cstring>

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

constexpr std::uintptr_t kAddr_InitContext = 0x4C74E0;
constexpr std::uintptr_t kAddr_GetOuterCount = 0x48E7B0;
constexpr std::uintptr_t kAddr_GetInnerCount = 0x492E80;
constexpr std::uintptr_t kAddr_FetchSearchText = 0x4C7590;
constexpr std::uintptr_t kAddr_ContainerGetAt = 0x4E7EA0;
constexpr std::uintptr_t kAddr_ContainerGetId = 0x4E7EE0;
constexpr std::uintptr_t kAddr_ResolveBucketIndex = 0x4E7F40;
constexpr std::uintptr_t kAddr_OpenCodeTarget = 0x403A80;
constexpr std::uintptr_t kAddr_MoveToLine = 0x4BA580;
constexpr std::uintptr_t kAddr_EnsureVisible = 0x4BAFB0;
constexpr std::uintptr_t kAddr_MoveCaretToOffset = 0x4AAC10;
constexpr std::uintptr_t kAddr_ActivateWindow = 0x53CF5B;
constexpr std::uintptr_t kAddr_NotifyOpenFailure = 0x47B3C0;
constexpr std::uintptr_t kAddr_CStringDestroy = 0x53BB23;
constexpr std::uintptr_t kAddr_EmptyCStringData = 0x5C72A4;

constexpr std::uintptr_t kAddr_Type1Container = 0x5CB184;
constexpr std::uintptr_t kAddr_Type2Data = 0x5CB12C;
constexpr std::uintptr_t kAddr_Type3Data = 0x5CB148;
constexpr std::uintptr_t kAddr_Type4Data = 0x5CB1A0;
constexpr std::uintptr_t kAddr_Type678Data = 0x5CB1D8;
constexpr std::uintptr_t kAddr_MainEditorHost = 0x5CAE70;
constexpr std::uintptr_t kAddr_OwnerObject = 0x5CAF30;
constexpr std::uintptr_t kAddr_BuiltinSearchDialogCtor = 0x445C00;
constexpr std::uintptr_t kAddr_BuiltinSearchDialogDtor = 0x445CE0;
constexpr std::uintptr_t kAddr_DialogDoModal = 0x53C618;
constexpr std::uintptr_t kAddr_MainWindowObject = 0x5CB790;
constexpr std::uintptr_t kAddr_SearchMode = 0x5CAD6C;
constexpr std::uintptr_t kAddr_ConsumeSearchResultRecord = 0x4A9170;
constexpr std::uintptr_t kAddr_FromHandle = 0x53D57E;
constexpr std::uintptr_t kAddr_EditorGetOuterCount = 0x4C05B0;
constexpr std::uintptr_t kAddr_EditorGetInnerCount = 0x4C5440;
constexpr std::uintptr_t kAddr_EditorFetchLineText = 0x4C53F0;

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
    std::vector<std::string> capturedLines;
    std::vector<e571::DirectGlobalSearch::GlobalSearchHit> rawHits;
};

thread_local HiddenBuiltinSearchContext* g_hiddenBuiltinSearchContext = nullptr;

FnAppendBytes g_originalAppendBytes = nullptr;
FnPrepareSearchResults g_originalPrepareSearchResults = nullptr;
FnSelectSearchResultTab g_originalSelectSearchResultTab = nullptr;
FnActivateWindowObject g_originalActivateWindowObject = nullptr;
FnSendMessageA g_originalSendMessageA = ::SendMessageA;

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
    value.data = *PtrAbsolute<const char*>(moduleBase, kAddr_EmptyCStringData);
}

void DestroyExeCString(ExeCStringA& value, std::uintptr_t moduleBase) {
    BindAbsolute<FnCStringDestroy>(moduleBase, kAddr_CStringDestroy)(&value);
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
                    ctx->capturedLines.clear();
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
                    ctx->capturedLines.push_back(std::move(text));
                    const std::string& last = ctx->capturedLines.back();
                    if (IsBuiltinSearchDecorativeLine(last) &&
                        (ctx->keyword == nullptr || std::strstr(last.c_str(), ctx->keyword) == nullptr) &&
                        ctx->capturedLines.size() >= 2) {
                        ctx->searchFinished = true;
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
    const int mode = *PtrAbsolute<int>(moduleBase, kAddr_SearchMode);
    auto* mainWindowObject = reinterpret_cast<unsigned char*>(PtrAbsolute<void>(moduleBase, kAddr_MainWindowObject));
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
        const int mode = *PtrAbsolute<int>(moduleBase, kAddr_SearchMode);
        auto* mainWindowObject = reinterpret_cast<unsigned char*>(PtrAbsolute<void>(moduleBase, kAddr_MainWindowObject));
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

    const auto initContext = BindAbsolute<FnInitContext>(moduleBase, kAddr_InitContext);
    const auto resolveBucketIndex = BindAbsolute<FnResolveBucketIndex>(moduleBase, kAddr_ResolveBucketIndex);

    e571::DirectGlobalSearch::SearchContext ctx{};
    initContext(&ctx);
    ctx.type = hit.type;
    ctx.flag = 0;
    ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_OwnerObject)));

    if (hit.type == 1) {
        int bucketData = 0;
        if (!SafeResolveBucketIndex(
                resolveBucketIndex,
                PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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
            int*)>(BindAbsolute<void*>(moduleBase, kAddr_FetchSearchText)),
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

    __try {
        BindAbsolute<FnConsumeSearchResultRecord>(moduleBase, kAddr_ConsumeSearchResultRecord)(
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

    const auto openCodeTarget = BindAbsolute<FnOpenCodeTarget>(moduleBase, kAddr_OpenCodeTarget);
    void* editorObject = nullptr;

    if (hit.type == 1) {
        int resolvedIndex = -1;
        const int ok = BindAbsolute<FnResolveBucketIndex>(moduleBase, kAddr_ResolveBucketIndex)(
            PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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
            PtrAbsolute<void>(moduleBase, kAddr_MainEditorHost),
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
            PtrAbsolute<void>(moduleBase, kAddr_MainEditorHost),
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

    const auto getOuterCount = BindAbsolute<FnEditorGetOuterCount>(moduleBase, kAddr_EditorGetOuterCount);
    const auto getInnerCount = BindAbsolute<FnEditorGetInnerCount>(moduleBase, kAddr_EditorGetInnerCount);
    const auto fetchLineText = BindAbsolute<FnEditorFetchLineTextRaw>(moduleBase, kAddr_EditorFetchLineText);

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

    const auto getOuterCount = BindAbsolute<FnGetOuterCount>(moduleBase, kAddr_GetOuterCount);
    const auto getInnerCount = BindAbsolute<FnGetInnerCount>(moduleBase, kAddr_GetInnerCount);
    const auto fetchSearchText = BindAbsolute<int(__thiscall*)(
        e571::DirectGlobalSearch::SearchContext*,
        int,
        int,
        ExeCStringA*,
        int,
        int,
        ExeCStringA*,
        int,
        int*)>(moduleBase, kAddr_FetchSearchText);

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

    const auto openCodeTarget = BindAbsolute<FnOpenCodeTarget>(moduleBase, kAddr_OpenCodeTarget);
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

        const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, kAddr_ContainerGetAt);
        int bucketOk = 0;
        if (!SafeContainerGetAt(
                containerGetAt,
                PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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
        PtrAbsolute<void>(moduleBase, kAddr_MainEditorHost),
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

    WaitForBuiltinSearchResults(moduleBase, 300);
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

        for (size_t i = 0; i < ctx.rawHits.size(); ++i) {
            std::string text;
            if (!TryFormatRawSearchHit(ctx.moduleBase, ctx.rawHits[i], &text) || text.empty()) {
                continue;
            }

            if (result.firstResultText.empty()) {
                result.firstResultText = text;
            }
            if (result.previewLines.size() < kBuiltinResultPreviewLimit) {
                result.previewLines.push_back(std::move(text));
            }
        }
        return;
    }

    for (const std::string& line : ctx.capturedLines) {
        if (line.empty()) {
            continue;
        }
        if (IsBuiltinSearchDecorativeLine(line)) {
            continue;
        }

        if (result.firstResultText.empty()) {
            result.firstResultText = line;
        }
        ++result.hits;
        if (result.previewLines.size() < kBuiltinResultPreviewLimit) {
            result.previewLines.push_back(line);
        }
    }
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
    e571::HiddenBuiltinSearchDebugResult* outCapturedResult) {
    std::array<unsigned char, kBuiltinSearchDialogStorageSize> dialogStorage = {};
    void* const dialogObject = dialogStorage.data();

    auto ctor = BindAbsolute<FnBuiltinSearchDialogCtor>(moduleBase, kAddr_BuiltinSearchDialogCtor);
    auto dtor = BindAbsolute<FnBuiltinSearchDialogDtor>(moduleBase, kAddr_BuiltinSearchDialogDtor);
    auto doModal = BindAbsolute<FnDialogDoModal>(moduleBase, kAddr_DialogDoModal);

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
        WaitForBuiltinSearchResults(hookContext, moduleBase, 5000);
    }
    dtor(dialogObject, 0);

    if (outDialogHandled != nullptr) {
        *outDialogHandled = hookContext.dialogHandled;
    }
    if (outCapturedResult != nullptr) {
        FillCapturedBuiltinSearchResult(hookContext, *outCapturedResult);
    }
    return true;
}

bool RunBuiltinSearchDialogHiddenSafe(
    const char* keyword,
    std::uintptr_t moduleBase,
    bool* outDialogHandled,
    e571::HiddenBuiltinSearchDebugResult* outCapturedResult) {
    __try {
        return RunBuiltinSearchDialogHidden(keyword, moduleBase, outDialogHandled, outCapturedResult);
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

    const int mode = *PtrAbsolute<int>(moduleBase, kAddr_SearchMode);
    auto* mainWindowObject = reinterpret_cast<unsigned char*>(PtrAbsolute<void>(moduleBase, kAddr_MainWindowObject));
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

    g_originalAppendBytes = BindAbsolute<FnAppendBytes>(moduleBase, 0x486920);
    g_originalPrepareSearchResults = BindAbsolute<FnPrepareSearchResults>(moduleBase, 0x4A8C90);
    g_originalSelectSearchResultTab = BindAbsolute<FnSelectSearchResultTab>(moduleBase, 0x4A8C20);
    g_originalActivateWindowObject = BindAbsolute<FnActivateWindowObject>(moduleBase, kAddr_ActivateWindow);

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
    ctx.previousFocusHwnd = ::GetFocus();
    ctx.previousFocusObject = nullptr;
    if (::IsWindow(ctx.previousFocusHwnd)) {
        ctx.previousFocusObject = BindAbsolute<FnFromHandle>(moduleBase_, kAddr_FromHandle)(ctx.previousFocusHwnd);
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
            BindAbsolute<FnActivateWindowObject>(moduleBase_, kAddr_ActivateWindow)(
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
    switch (type) {
    case 2:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_Type2Data)));
    case 3:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_Type3Data)));
    case 4:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_Type4Data)));
    case 6:
    case 7:
    case 8:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_Type678Data)));
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
    const auto getOuterCount = BindAbsolute<FnGetOuterCount>(moduleBase, kAddr_GetOuterCount);
    const auto getInnerCount = BindAbsolute<FnGetInnerCount>(moduleBase, kAddr_GetInnerCount);
    const auto fetchSearchText = BindAbsolute<int(__thiscall*)(
        e571::DirectGlobalSearch::SearchContext*,
        int,
        int,
        ExeCStringA*,
        int,
        int,
        ExeCStringA*,
        int,
        int*)>(moduleBase, kAddr_FetchSearchText);

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

    for (int type : kSearchTypes) {
        if (type == 1) {
            EnumerateType1(keyword, keywordLen, options, results);
        } else {
            SearchContext ctx{};
            InitContext(&ctx);

            ctx.type = type;
            ctx.data = ResolveTypeData(type);
            ctx.flag = 0;
            ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(kAddr_OwnerObject)));

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
            Ptr<DwordContainer>(kAddr_Type1Container),
            hit.extra,
            nullptr,
            &resolvedIndex);
        if (ok == 0) {
            NotifyOpenFailure(0x30);
            return false;
        }

        HWND hwnd = OpenCodeTarget(
            Ptr<void>(kAddr_MainEditorHost),
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
        Ptr<void>(kAddr_MainEditorHost),
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
    Bind<FnInitContext>(kAddr_InitContext)(ctx);
}

int DirectGlobalSearch::GetOuterCount(SearchContext* ctx) const {
    return Bind<FnGetOuterCount>(kAddr_GetOuterCount)(ctx);
}

int DirectGlobalSearch::GetInnerCount(SearchContext* ctx, int outerIndex) const {
    return Bind<FnGetInnerCount>(kAddr_GetInnerCount)(ctx, outerIndex);
}

int DirectGlobalSearch::FetchSearchText(
    SearchContext* ctx,
    int outerIndex,
    int innerIndex,
    CStringA* outText,
    CStringA* outPrefix,
    int filter,
    int* optionalOut) const {
    return Bind<FnFetchSearchText>(kAddr_FetchSearchText)(
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
    return Bind<FnContainerGetAt>(kAddr_ContainerGetAt)(container, index, outPtr);
}

int DirectGlobalSearch::ContainerGetId(void* container, int index) const {
    return Bind<FnContainerGetId>(kAddr_ContainerGetId)(container, index);
}

int DirectGlobalSearch::ResolveBucketIndex(void* container, int bucketId, int* outValue, int* outPos) const {
    return Bind<FnResolveBucketIndex>(kAddr_ResolveBucketIndex)(container, bucketId, outValue, outPos);
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
    return Bind<FnOpenCodeTarget>(kAddr_OpenCodeTarget)(
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
    Bind<FnMoveToLine>(kAddr_MoveToLine)(hwnd, outerIndex, innerIndex, arg3, arg4, arg5);
}

void DirectGlobalSearch::EnsureVisible(HWND hwnd, int arg1, int arg2, int arg3) const {
    Bind<FnEnsureVisible>(kAddr_EnsureVisible)(hwnd, arg1, arg2, arg3);
}

void DirectGlobalSearch::MoveCaretToOffset(HWND hwnd, int matchOffset, int force, void* maybeNull, int redraw) const {
    Bind<FnMoveCaretToOffset>(kAddr_MoveCaretToOffset)(hwnd, matchOffset, force, maybeNull, redraw);
}

void DirectGlobalSearch::ActivateWindow(HWND hwnd) const {
    Bind<FnActivateWindow>(kAddr_ActivateWindow)(hwnd);
}

void DirectGlobalSearch::NotifyOpenFailure(int messageId) const {
    Bind<FnNotifyOpenFailure>(kAddr_NotifyOpenFailure)(Ptr<void>(kAddr_MainEditorHost), messageId);
}

int DirectGlobalSearch::ResolveTypeData(int type) const {
    switch (type) {
    case 2:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(kAddr_Type2Data)));
    case 3:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(kAddr_Type3Data)));
    case 4:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(kAddr_Type4Data)));
    case 6:
    case 7:
    case 8:
        return static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(kAddr_Type678Data)));
    default:
        return 0;
    }
}

void DirectGlobalSearch::EnumerateType1(
    const char* keyword,
    std::size_t keywordLen,
    const SearchOptions& options,
    std::vector<SearchResult>& results) const {
    for (int bucketIndex = 0;; ++bucketIndex) {
        SearchContext ctx{};
        InitContext(&ctx);

        int bucketData = 0;
        if (ContainerGetAt(Ptr<DwordContainer>(kAddr_Type1Container), bucketIndex, &bucketData) == 0) {
            break;
        }

        ctx.type = 1;
        ctx.data = bucketData;
        ctx.flag = 0;
        ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(Ptr<void>(kAddr_OwnerObject)));

        const int bucketId = ContainerGetId(Ptr<DwordContainer>(kAddr_Type1Container), bucketIndex);
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

    const auto initContext = BindAbsolute<FnInitContext>(moduleBase, kAddr_InitContext);
    const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, kAddr_ContainerGetAt);
    const auto containerGetId = BindAbsolute<FnContainerGetId>(moduleBase, kAddr_ContainerGetId);

    for (int type : kSearchTypes) {
        if (type == 1) {
            for (int bucketIndex = 0;; ++bucketIndex) {
                DirectGlobalSearch::SearchContext ctx{};
                initContext(&ctx);

                int bucketData = 0;
                int bucketOk = 0;
                if (!SafeContainerGetAt(
                        containerGetAt,
                        PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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
                ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_OwnerObject)));

                int bucketId = 0;
                if (!SafeContainerGetId(
                        containerGetId,
                        PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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
        ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_OwnerObject)));
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
    if (!RunBuiltinSearchDialogHiddenSafe(keyword, moduleBase, &dialogHandled, &result)) {
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
    if (!RunBuiltinSearchDialogHiddenSafe(keyword, moduleBase, &dialogHandled, &captured)) {
        return hits;
    }
    if (outDialogHandled != nullptr) {
        *outDialogHandled = dialogHandled;
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
        const int ok = BindAbsolute<FnResolveBucketIndex>(moduleBase, kAddr_ResolveBucketIndex)(
            PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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

    DirectGlobalSearch::SearchContext ctx{};
    BindAbsolute<FnInitContext>(moduleBase, kAddr_InitContext)(&ctx);
    ctx.flag = 0;
    ctx.owner = static_cast<int>(reinterpret_cast<std::uintptr_t>(PtrAbsolute<void>(moduleBase, kAddr_OwnerObject)));

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

        const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, kAddr_ContainerGetAt);
        int bucketData = 0;
        int bucketOk = 0;
        if (!SafeContainerGetAt(
                containerGetAt,
                PtrAbsolute<void>(moduleBase, kAddr_Type1Container),
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
