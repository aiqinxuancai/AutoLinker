#include "direct_global_search.hpp"
#include "direct_global_search_debug.hpp"

#include <Windows.h>
#include <detours.h>

#include <cstdint>
#include <format>
#include <initializer_list>

#include <cstring>
#include <mutex>

#include "MemFind.h"

void OutputStringToELog(const std::string& szbuf);

namespace e571 {

namespace {

using FnContainerGetAt = int(__thiscall*)(void*, int, int*);
using FnOpenCodeTarget = HWND(__thiscall*)(void*, int, int, int, int, int, int, int);
using FnCreateCodePage = void*(__thiscall*)(void*, int, int, void*, int);
using FnSetMainEditorActiveEditorObject = int(__thiscall*)(void*, void*, int);
using FnSendMessageA = LRESULT(WINAPI*)(HWND, UINT, WPARAM, LPARAM);
using FnSetFocusApi = HWND(WINAPI*)(HWND);

constexpr std::uintptr_t kImageBase = 0x400000;

struct NativeSearchAddresses {
    bool initialized = false;
    bool ok = false;
    std::uintptr_t moduleBase = 0;

    std::uintptr_t containerGetAt = 0;
    std::uintptr_t openCodeTarget = 0;
    std::uintptr_t createCodePage = 0;

    std::uintptr_t type1Container = 0;
    std::uintptr_t type2Data = 0;
    std::uintptr_t type3Data = 0;
    std::uintptr_t type4Data = 0;
    std::uintptr_t type678Data = 0;
    std::uintptr_t mainEditorHost = 0;
    std::uintptr_t setMainEditorActiveEditorObject = 0;
    ptrdiff_t mainEditorHostActiveEditorObjectOffset = 0;
};

struct OpenCodeTargetAddresses {
    bool initialized = false;
    bool ok = false;
    std::uintptr_t moduleBase = 0;

    std::uintptr_t containerGetAt = 0;
    std::uintptr_t openCodeTarget = 0;
    std::uintptr_t type1Container = 0;
    std::uintptr_t mainEditorHost = 0;
    std::string trace;
};

constexpr ptrdiff_t kOffset_MainEditorHostCreateOwner = 0x0C0;
constexpr ptrdiff_t kOffset_MainEditorHostOpenPageList = 0x924;
constexpr ptrdiff_t kOffset_MainEditorHostFormArray = 0x34C;
constexpr ptrdiff_t kOffset_PageObjectType = 0x50;
constexpr ptrdiff_t kOffset_PageObjectType5Key = 0x60;
constexpr ptrdiff_t kOffset_PageObjectType1Key = 0x68;
constexpr ptrdiff_t kOffset_PageRuntimeEditorObject = 0x40;
constexpr ptrdiff_t kOffset_Type1BucketOwner = 0x50;

struct HiddenOpenCodeTargetContext {
    bool suppressMdiActivate = true;
    bool suppressSetFocus = true;
    HWND blockedFocusHwnd = nullptr;
};

thread_local HiddenOpenCodeTargetContext* g_hiddenOpenCodeTargetContext = nullptr;

FnSendMessageA g_originalHiddenOpenSendMessageA = ::SendMessageA;
FnSetFocusApi g_originalHiddenOpenSetFocus = ::SetFocus;

LRESULT WINAPI HiddenOpenCodeTargetHook_SendMessageA(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
HWND WINAPI HiddenOpenCodeTargetHook_SetFocus(HWND hwnd);

class ScopedHiddenOpenCodeTargetApiHooks {
public:
    explicit ScopedHiddenOpenCodeTargetApiHooks(HiddenOpenCodeTargetContext& ctx);
    ~ScopedHiddenOpenCodeTargetApiHooks();

    bool IsInstalled() const;

private:
    HiddenOpenCodeTargetContext& ctx_;
    bool installed_;
};

std::mutex g_nativeSearchAddressMutex;
NativeSearchAddresses g_nativeSearchAddresses;
std::mutex g_openCodeTargetAddressMutex;
OpenCodeTargetAddresses g_openCodeTargetAddresses;

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

bool HasUnsafeDirectGlobalSearchLayoutSupport(
    std::uintptr_t moduleBase,
    std::string* outTrace = nullptr) {
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    if (moduleBase == 0) {
        if (outTrace != nullptr) {
            *outTrace = "direct_global_search_fixed_layout_unsupported|module_base_invalid";
        }
        return false;
    }

    struct KnownLayoutProfile {
        const char* name;
        std::uintptr_t setActiveEditorObjectRva;
        ptrdiff_t activeEditorObjectOffset;
    };

    static constexpr KnownLayoutProfile kProfiles[] = {
        { "e5.95", 0x47ADF0, 0x464 },
        { "e571", 0x471B30, 0x438 },
    };

    auto tryExtractActiveEditorOffset = [moduleBase](std::uintptr_t functionRva, ptrdiff_t* outOffset) -> bool {
        if (outOffset != nullptr) {
            *outOffset = 0;
        }
        if (functionRva < kImageBase) {
            return false;
        }

        const auto* const code = reinterpret_cast<const std::uint8_t*>(moduleBase + (functionRva - kImageBase));
        __try {
            if (code == nullptr ||
                code[0] != 0x53 ||
                code[1] != 0x8B ||
                code[2] != 0xD9 ||
                code[3] != 0x56 ||
                code[4] != 0x57 ||
                code[5] != 0x8B ||
                code[6] != 0xB3 ||
                code[21] != 0xC7 ||
                code[22] != 0x83 ||
                code[44] != 0x8B ||
                code[45] != 0x83 ||
                code[56] != 0x89 ||
                code[57] != 0xBB) {
                return false;
            }

            const std::uint32_t disp1 = *reinterpret_cast<const std::uint32_t*>(code + 7);
            const std::uint32_t disp2 = *reinterpret_cast<const std::uint32_t*>(code + 23);
            const std::uint32_t disp3 = *reinterpret_cast<const std::uint32_t*>(code + 46);
            const std::uint32_t disp4 = *reinterpret_cast<const std::uint32_t*>(code + 58);
            if (disp1 == 0 || disp1 != disp2 || disp1 != disp3 || disp1 != disp4) {
                return false;
            }

            if (outOffset != nullptr) {
                *outOffset = static_cast<ptrdiff_t>(disp1);
            }
            return true;
        } __except (EXCEPTION_EXECUTE_HANDLER) {
            return false;
        }
    };

    std::string mismatchTrace;
    for (const auto& profile : kProfiles) {
        ptrdiff_t activeOffset = 0;
        if (!tryExtractActiveEditorOffset(profile.setActiveEditorObjectRva, &activeOffset)) {
            mismatchTrace +=
                std::string("|known_profile_try=") +
                profile.name +
                "|set_active_signature_mismatch";
            continue;
        }
        if (activeOffset != profile.activeEditorObjectOffset) {
            mismatchTrace +=
                std::string("|known_profile_try=") +
                profile.name +
                "|active_offset_mismatch=" +
                std::to_string(activeOffset);
            continue;
        }

        if (outTrace != nullptr) {
            *outTrace =
                std::string("direct_global_search_fixed_layout_supported") +
                "|known_profile=" +
                profile.name +
                "|active_offset=" +
                std::to_string(activeOffset);
        }
        return true;
    }

    if (outTrace != nullptr) {
        *outTrace =
            std::string("direct_global_search_fixed_layout_unsupported") +
            "|known_profile=no_match" +
            mismatchTrace;
    }
    return false;
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

std::uintptr_t ResolveUniqueCodeAddressFromPatterns(
    const char* label,
    std::initializer_list<const char*> patterns,
    std::uintptr_t moduleBase) {
    size_t patternIndex = 0;
    for (const char* pattern : patterns) {
        const auto matches = FindSelfModelMemoryAll(pattern);
        if (matches.size() == 1) {
            return NormalizeRuntimeAddress(static_cast<std::uintptr_t>(matches.front()), moduleBase);
        }
        OutputStringToELog(std::format(
            "[DirectGlobalSearch] resolve {} pattern#{} failed, matchCount={}",
            label,
            patternIndex,
            matches.size()));
        ++patternIndex;
    }
    return 0;
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

bool TryExtractActiveEditorOffsetFromSetActiveEditorObject(
    std::uintptr_t moduleBase,
    std::uintptr_t functionRva,
    ptrdiff_t* outOffset) {
    if (outOffset != nullptr) {
        *outOffset = 0;
    }
    if (moduleBase == 0 || functionRva == 0 || functionRva < kImageBase) {
        return false;
    }

    const auto* const code = reinterpret_cast<const std::uint8_t*>(moduleBase + (functionRva - kImageBase));
    __try {
        if (code == nullptr ||
            code[0] != 0x53 ||
            code[1] != 0x8B ||
            code[2] != 0xD9 ||
            code[3] != 0x56 ||
            code[4] != 0x57 ||
            code[5] != 0x8B ||
            code[6] != 0xB3 ||
            code[21] != 0xC7 ||
            code[22] != 0x83 ||
            code[44] != 0x8B ||
            code[45] != 0x83 ||
            code[56] != 0x89 ||
            code[57] != 0xBB) {
            return false;
        }

        const std::uint32_t disp1 = *reinterpret_cast<const std::uint32_t*>(code + 7);
        const std::uint32_t disp2 = *reinterpret_cast<const std::uint32_t*>(code + 23);
        const std::uint32_t disp3 = *reinterpret_cast<const std::uint32_t*>(code + 46);
        const std::uint32_t disp4 = *reinterpret_cast<const std::uint32_t*>(code + 58);
        if (disp1 == 0 || disp1 != disp2 || disp1 != disp3 || disp1 != disp4) {
            return false;
        }

        if (outOffset != nullptr) {
            *outOffset = static_cast<ptrdiff_t>(disp1);
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

std::uintptr_t ResolveSetMainEditorActiveEditorObjectAddress(
    std::uintptr_t moduleBase,
    ptrdiff_t* outActiveEditorOffset) {
    if (outActiveEditorOffset != nullptr) {
        *outActiveEditorOffset = 0;
    }

    const std::uintptr_t functionRva = ResolveUniqueCodeAddress(
        "set_main_editor_active_editor_object",
        "53 8B D9 56 57 8B B3 ?? ?? ?? ?? 8B 7C 24 10 3B FE 74 ?? 85 F6 C7 83 ?? ?? ?? ?? 00 00 00 00 74 ?? 56 57 6A 00 8B CE E8 ?? ?? ?? ?? 8B 83 ?? ?? ?? ?? 85 C0 75 ?? 85 FF 89 BB ?? ?? ?? ?? 74 ?? 83 7C 24 14 01 75 ?? 56 57 6A 01 8B CF E8 ?? ?? ?? ?? 5F 5E 5B C2 08 00",
        moduleBase);
    if (functionRva == 0) {
        return 0;
    }

    if (!TryExtractActiveEditorOffsetFromSetActiveEditorObject(
            moduleBase,
            functionRva,
            outActiveEditorOffset)) {
        OutputStringToELog("[DirectGlobalSearch] resolve main_editor_host_active_editor_offset failed");
    }
    return functionRva;
}

bool TryPopulateKnownOpenCodeTargetMainHost(OpenCodeTargetAddresses& addrs, std::uintptr_t moduleBase) {
    struct KnownProfile {
        const char* name = nullptr;
        std::uintptr_t mainEditorHostRva = 0;
        std::uintptr_t setActiveEditorObjectRva = 0;
        ptrdiff_t activeEditorObjectOffset = 0;
    };

    static constexpr KnownProfile kProfiles[] = {
        { "e5.95", 0x6756E0, 0x47ADF0, 0x464 },
        { "e571", 0x5CAE70, 0x471B30, 0x438 },
    };

    for (const auto& profile : kProfiles) {
        ptrdiff_t activeOffset = 0;
        if (!TryExtractActiveEditorOffsetFromSetActiveEditorObject(
                moduleBase,
                profile.setActiveEditorObjectRva,
                &activeOffset)) {
            addrs.trace +=
                std::string("|known_profile_try=") +
                profile.name +
                "|set_active_signature_mismatch";
            continue;
        }
        if (activeOffset != profile.activeEditorObjectOffset) {
            addrs.trace +=
                std::string("|known_profile_try=") +
                profile.name +
                "|active_offset_mismatch=" +
                std::to_string(activeOffset);
            continue;
        }

        addrs.mainEditorHost = profile.mainEditorHostRva;
        addrs.trace +=
            std::string("|known_profile=") +
            profile.name +
            "|main_editor_host=" +
            std::to_string(addrs.mainEditorHost);
        return true;
    }

    return false;
}

bool PopulateOpenCodeTargetAddresses(OpenCodeTargetAddresses& addrs, std::uintptr_t moduleBase) {
    addrs = {};
    addrs.initialized = true;
    addrs.moduleBase = moduleBase;
    addrs.trace = "open_code_target_probe";
    if (moduleBase == 0) {
        addrs.trace += "|module_base_invalid";
        return false;
    }

    addrs.openCodeTarget = ResolveUniqueCodeAddressFromPatterns(
        "open_code_target",
        {
            "83 EC 08 53 56 57 8B F9 8B 87 ?? ?? ?? ?? F7 D0 83 E0 01 3C 01 0F 84 C0 02 00 00 8B 5C 24 18 33 F6 83 FB 01 74 39",
            "83 EC 08 53 56 57 8B F9 8B 87 ?? ?? ?? ?? F7 D0 83 E0 01 3C 01",
        },
        moduleBase);
    addrs.containerGetAt = ResolveUniqueCodeAddress(
        "container_get_at",
        "8B 41 18 8B 54 24 04 C1 E8 03 3B D0 7D 23 56 8B 71 18 85 F6 5E 75 04 33 C9 EB 03 8B 49 10 03 C2",
        moduleBase);
    addrs.type1Container = ResolveUniqueImmAddress(
        "type1_container",
        "8B C1 41 89 4C 24 48 52 50 B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 85 C0 0F 84 1B 04 00 00 8B 44 24",
        10,
        moduleBase);

    if (!TryPopulateKnownOpenCodeTargetMainHost(addrs, moduleBase)) {
        addrs.trace += "|known_profile=no_match|fallback=pattern_scan";
        addrs.mainEditorHost = ResolveUniqueImmAddress(
            "main_editor_host",
            "A1 ?? ?? ?? ?? 3B C7 74 0A B9 ?? ?? ?? ?? E8 ?? ?? ?? ?? 8D 44 24 14 55 50 B9 ?? ?? ?? ?? E8",
            10,
            moduleBase);
    }

    addrs.ok =
        addrs.openCodeTarget != 0 &&
        addrs.mainEditorHost != 0;
    addrs.trace +=
        std::string("|ok=") + (addrs.ok ? "1" : "0") +
        "|open_code_target=" + std::to_string(addrs.openCodeTarget) +
        "|main_editor_host=" + std::to_string(addrs.mainEditorHost) +
        "|container_get_at=" + std::to_string(addrs.containerGetAt) +
        "|type1_container=" + std::to_string(addrs.type1Container);
    return addrs.ok;
}

const OpenCodeTargetAddresses& GetOpenCodeTargetAddresses(std::uintptr_t moduleBase) {
    std::lock_guard<std::mutex> lock(g_openCodeTargetAddressMutex);
    if (!g_openCodeTargetAddresses.initialized || g_openCodeTargetAddresses.moduleBase != moduleBase) {
        PopulateOpenCodeTargetAddresses(g_openCodeTargetAddresses, moduleBase);
    }
    return g_openCodeTargetAddresses;
}

bool PopulateNativeSearchAddresses(NativeSearchAddresses& addrs, std::uintptr_t moduleBase) {
    addrs = {};
    addrs.moduleBase = moduleBase;

    addrs.containerGetAt = ResolveUniqueCodeAddress(
        "container_get_at",
        "8B 41 18 8B 54 24 04 C1 E8 03 3B D0 7D 23 56 8B 71 18 85 F6 5E 75 04 33 C9 EB 03 8B 49 10 03 C2",
        moduleBase);
    addrs.openCodeTarget = ResolveUniqueCodeAddressFromPatterns(
        "open_code_target",
        {
            "83 EC 08 53 56 57 8B F9 8B 87 ?? ?? ?? ?? F7 D0 83 E0 01 3C 01 0F 84 C0 02 00 00 8B 5C 24 18 33 F6 83 FB 01 74 39",
            "83 EC 08 53 56 57 8B F9 8B 87 ?? ?? ?? ?? F7 D0 83 E0 01 3C 01",
        },
        moduleBase);
    addrs.createCodePage = ResolveUniqueCodeAddress(
        "create_code_page",
        "64 A1 00 00 00 00 6A FF 68 ?? ?? ?? ?? 50 64 89 25 00 00 00 00 53 8B D9 55 56 57 8B 8B 24 09 00 00 8B 01 FF 50 6C",
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
    addrs.setMainEditorActiveEditorObject = ResolveSetMainEditorActiveEditorObjectAddress(
        moduleBase,
        &addrs.mainEditorHostActiveEditorObjectOffset);

    addrs.ok =
        addrs.containerGetAt != 0 &&
        addrs.openCodeTarget != 0 &&
        addrs.createCodePage != 0 &&
        addrs.type1Container != 0 &&
        addrs.type2Data != 0 &&
        addrs.type3Data != 0 &&
        addrs.type4Data != 0 &&
        addrs.type678Data != 0 &&
        addrs.mainEditorHost != 0 &&
        addrs.setMainEditorActiveEditorObject != 0 &&
        addrs.mainEditorHostActiveEditorObjectOffset > 0;
    addrs.initialized = true;
    return addrs.ok;
}
const NativeSearchAddresses& GetNativeSearchAddresses(std::uintptr_t moduleBase) {
    std::lock_guard<std::mutex> lock(g_nativeSearchAddressMutex);
    if (!g_nativeSearchAddresses.initialized || g_nativeSearchAddresses.moduleBase != moduleBase) {
        g_nativeSearchAddresses = {};
        g_nativeSearchAddresses.moduleBase = moduleBase;

        std::string supportTrace;
        if (!HasUnsafeDirectGlobalSearchLayoutSupport(moduleBase, &supportTrace)) {
            g_nativeSearchAddresses.initialized = true;
            g_nativeSearchAddresses.ok = false;
            return g_nativeSearchAddresses;
        }

        PopulateNativeSearchAddresses(g_nativeSearchAddresses, moduleBase);
    }
    return g_nativeSearchAddresses;
}

bool HasNativeSearchAddresses(std::uintptr_t moduleBase) {
    return GetNativeSearchAddresses(moduleBase).ok;
}

bool DebugIsDirectGlobalSearchSupportedImpl(
    std::uintptr_t moduleBase,
    std::string* outTrace) {
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    std::string supportTrace;
    if (!HasUnsafeDirectGlobalSearchLayoutSupport(moduleBase, &supportTrace)) {
        if (outTrace != nullptr) {
            *outTrace = supportTrace.empty()
                ? "direct_global_search_fixed_layout_unsupported_without_probe"
                : supportTrace;
        }
        return false;
    }

    if (!HasNativeSearchAddresses(moduleBase)) {
        if (outTrace != nullptr) {
            *outTrace = "resolve_native_addresses_failed";
        }
        return false;
    }

    if (outTrace != nullptr) {
        *outTrace = "direct_global_search_supported";
    }
    return true;
}

bool SafeContainerGetAt(
    FnContainerGetAt fn,
    void* container,
    int index,
    int* outPtr,
    int* outOk);
bool SafeCallOpenCodeTarget(
    FnOpenCodeTarget fn,
    void* mainEditorHost,
    int openType,
    int arg2,
    int arg3,
    void** outEditorObject);
int ResolveTypeDataRaw(int type, std::uintptr_t moduleBase);

template <typename T>
T BindAbsolute(std::uintptr_t moduleBase, std::uintptr_t absoluteAddress) {
    return reinterpret_cast<T>(absoluteAddress - kImageBase + moduleBase);
}

template <typename T>
T* PtrAbsolute(std::uintptr_t moduleBase, std::uintptr_t absoluteAddress) {
    return reinterpret_cast<T*>(absoluteAddress - kImageBase + moduleBase);
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

    const auto& addrs = GetOpenCodeTargetAddresses(moduleBase);
    if (!addrs.ok) {
        if (outTrace != nullptr) {
            *outTrace = "resolve_open_code_target_addresses_failed|" + addrs.trace;
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

        if (addrs.containerGetAt != 0 && addrs.type1Container != 0) {
            const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, addrs.containerGetAt);
            int bucketOk = 0;
            if (!SafeContainerGetAt(
                    containerGetAt,
                    PtrAbsolute<void>(moduleBase, addrs.type1Container),
                    resolvedIndex,
                    &bucketData,
                    &bucketOk) ||
                bucketOk == 0) {
                bucketData = 0;
            }
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
        resolvedIndex = static_cast<int>(itemData & 0x0FFFFFFFu);
        openType = 5;
        arg2 = resolvedIndex;
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

    void* editorObject = nullptr;
    const DWORD openStartTick = ::GetTickCount();
    if (!SafeCallOpenCodeTarget(
            openCodeTarget,
            PtrAbsolute<void>(moduleBase, addrs.mainEditorHost),
            openType,
            arg2,
            arg3,
            &editorObject)) {
        if (outTrace != nullptr) {
            *outTrace =
                "open_code_target_exception"
                "|open_type=" + std::to_string(openType) +
                "|arg2=" + std::to_string(arg2) +
                "|arg3=" + std::to_string(arg3) +
                "|resolved_index=" + std::to_string(resolvedIndex) +
                "|bucket_data=" + std::to_string(bucketData) +
                "|" +
                addrs.trace;
        }
        return false;
    }
    const DWORD openElapsed = ::GetTickCount() - openStartTick;
    if (editorObject == nullptr) {
        if (outTrace != nullptr) {
            *outTrace =
                "open_code_target_null"
                "|open_ms=" + std::to_string(openElapsed) +
                "|open_type=" + std::to_string(openType) +
                "|arg2=" + std::to_string(arg2) +
                "|arg3=" + std::to_string(arg3) +
                "|resolved_index=" + std::to_string(resolvedIndex) +
                "|bucket_data=" + std::to_string(bucketData) +
                "|" +
                addrs.trace;
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
        *outTrace =
            "open_editor_ok"
            "|open_ms=" + std::to_string(openElapsed) +
            "|open_type=" + std::to_string(openType) +
            "|arg2=" + std::to_string(arg2) +
            "|arg3=" + std::to_string(arg3) +
            "|resolved_index=" + std::to_string(resolvedIndex) +
            "|bucket_data=" + std::to_string(bucketData) +
            "|" +
            addrs.trace;
    }
    return true;
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

bool SafeFindExistingOpenEditorObject(
    void* mainEditorHost,
    int openType,
    int matchKey,
    void** outEditorObject) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (mainEditorHost == nullptr) {
        return false;
    }

    __try {
        auto* const hostBytes = reinterpret_cast<std::uint8_t*>(mainEditorHost);
        void* const openPageList =
            *reinterpret_cast<void**>(hostBytes + kOffset_MainEditorHostOpenPageList);
        if (openPageList == nullptr) {
            return false;
        }

        const auto vtable = *reinterpret_cast<std::uintptr_t*>(openPageList);
        if (vtable == 0) {
            return false;
        }

        const auto getFirst = reinterpret_cast<int(__thiscall*)(void*)>(
            *reinterpret_cast<const std::uintptr_t*>(vtable + 0x54));
        const auto getNext = reinterpret_cast<void*(__thiscall*)(void*, int*)>(
            *reinterpret_cast<const std::uintptr_t*>(vtable + 0x58));
        if (getFirst == nullptr || getNext == nullptr) {
            return false;
        }

        int cursor = getFirst(openPageList);
        while (cursor != 0) {
            void* const candidate = getNext(openPageList, &cursor);
            if (candidate == nullptr) {
                continue;
            }

            auto* const candidateBytes = reinterpret_cast<const std::uint8_t*>(candidate);
            const int candidateType =
                *reinterpret_cast<const int*>(candidateBytes + kOffset_PageObjectType);
            if (candidateType != openType) {
                continue;
            }

            if (openType == 1) {
                const int bucketKey =
                    *reinterpret_cast<const int*>(candidateBytes + kOffset_PageObjectType1Key);
                if (bucketKey != matchKey) {
                    continue;
                }
            } else if (openType == 5) {
                const int formKey =
                    *reinterpret_cast<const int*>(candidateBytes + kOffset_PageObjectType5Key);
                if (formKey != matchKey) {
                    continue;
                }
            }

            if (outEditorObject != nullptr) {
                *outEditorObject = candidate;
            }
            return true;
        }
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }

    return false;
}

bool SafeResolveRuntimeEditorObjectFromPageObject(
    void* mainEditorHost,
    ptrdiff_t activeEditorObjectOffset,
    void* pageObject,
    void** outEditorObject) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (mainEditorHost == nullptr || pageObject == nullptr) {
        return false;
    }

    __try {
        std::uintptr_t preferredEditorObject = 0;
        if (activeEditorObjectOffset > 0) {
            preferredEditorObject = *reinterpret_cast<const std::uintptr_t*>(
                reinterpret_cast<const std::uint8_t*>(mainEditorHost) + activeEditorObjectOffset);
        }
        const auto pageVtable = *reinterpret_cast<const std::uintptr_t*>(pageObject);
        if (pageVtable == 0) {
            return false;
        }

        const auto getFirst = reinterpret_cast<int(__thiscall*)(void*)>(
            *reinterpret_cast<const std::uintptr_t*>(pageVtable + 96));
        const auto getNext = reinterpret_cast<void*(__thiscall*)(void*, int*)>(
            *reinterpret_cast<const std::uintptr_t*>(pageVtable + 100));
        if (getFirst == nullptr || getNext == nullptr) {
            return false;
        }

        void* fallbackEditorObject = nullptr;
        int cursor = getFirst(pageObject);
        if (cursor == 0) {
            return false;
        }

        do {
            void* const pageNode = getNext(pageObject, &cursor);
            if (pageNode == nullptr) {
                continue;
            }

            void* const editorObject = *reinterpret_cast<void**>(
                reinterpret_cast<std::uint8_t*>(pageNode) + kOffset_PageRuntimeEditorObject);
            if (editorObject == nullptr) {
                continue;
            }

            if (fallbackEditorObject == nullptr) {
                fallbackEditorObject = editorObject;
            }
            if (preferredEditorObject != 0 &&
                reinterpret_cast<std::uintptr_t>(editorObject) == preferredEditorObject) {
                if (outEditorObject != nullptr) {
                    *outEditorObject = editorObject;
                }
                return true;
            }
        } while (cursor != 0);

        if (outEditorObject != nullptr) {
            *outEditorObject = fallbackEditorObject;
        }
        return fallbackEditorObject != nullptr;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeCreateEditorObjectNoActivate(
    FnCreateCodePage fn,
    void* mainEditorHost,
    int ownerArg,
    int openType,
    int dataArg,
    int extraArg,
    void** outEditorObject) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (fn == nullptr || mainEditorHost == nullptr) {
        return false;
    }

    __try {
        void* const editorObject = fn(
            mainEditorHost,
            ownerArg,
            openType,
            reinterpret_cast<void*>(dataArg),
            extraArg);
        if (outEditorObject != nullptr) {
            *outEditorObject = editorObject;
        }
        return editorObject != nullptr;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeSetMainEditorActiveEditorObject(
    FnSetMainEditorActiveEditorObject fn,
    void* mainEditorHost,
    ptrdiff_t activeEditorObjectOffset,
    void* editorObject,
    int notifyMode,
    std::uintptr_t* outPreviousEditorObject,
    std::uintptr_t* outCurrentEditorObject) {
    if (outPreviousEditorObject != nullptr) {
        *outPreviousEditorObject = 0;
    }
    if (outCurrentEditorObject != nullptr) {
        *outCurrentEditorObject = 0;
    }
    if (fn == nullptr || mainEditorHost == nullptr) {
        return false;
    }

    __try {
        auto* const hostBytes = reinterpret_cast<std::uint8_t*>(mainEditorHost);
        std::uintptr_t previousEditorObject = 0;
        std::uintptr_t currentEditorObject = 0;
        if (activeEditorObjectOffset > 0) {
            previousEditorObject =
                *reinterpret_cast<const std::uintptr_t*>(hostBytes + activeEditorObjectOffset);
        }
        fn(mainEditorHost, editorObject, notifyMode);
        if (activeEditorObjectOffset > 0) {
            currentEditorObject =
                *reinterpret_cast<const std::uintptr_t*>(hostBytes + activeEditorObjectOffset);
        }
        if (outPreviousEditorObject != nullptr) {
            *outPreviousEditorObject = previousEditorObject;
        }
        if (outCurrentEditorObject != nullptr) {
            *outCurrentEditorObject = currentEditorObject;
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeCallOpenCodeTarget(
    FnOpenCodeTarget fn,
    void* mainEditorHost,
    int openType,
    int arg2,
    int arg3,
    void** outEditorObject) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (fn == nullptr || mainEditorHost == nullptr) {
        return false;
    }

    __try {
        void* const editorObject = reinterpret_cast<void*>(fn(
            mainEditorHost,
            openType,
            arg2,
            arg3,
            0,
            0,
            1,
            -1));
        if (outEditorObject != nullptr) {
            *outEditorObject = editorObject;
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool SafeReadMainEditorActiveEditorObject(
    void* mainEditorHost,
    ptrdiff_t activeEditorObjectOffset,
    std::uintptr_t* outEditorObject) {
    if (outEditorObject != nullptr) {
        *outEditorObject = 0;
    }
    if (mainEditorHost == nullptr || activeEditorObjectOffset <= 0) {
        return false;
    }

    __try {
        const auto* const hostBytes = reinterpret_cast<const std::uint8_t*>(mainEditorHost);
        const std::uintptr_t editorObject = *reinterpret_cast<const std::uintptr_t*>(
            hostBytes + activeEditorObjectOffset);
        if (outEditorObject != nullptr) {
            *outEditorObject = editorObject;
        }
        return true;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

bool TryOpenCodeTargetNoActivate(
    FnOpenCodeTarget fn,
    void* mainEditorHost,
    int openType,
    int arg2,
    int arg3,
    void** outEditorObject,
    std::string* outTrace) {
    if (outEditorObject != nullptr) {
        *outEditorObject = nullptr;
    }
    if (outTrace != nullptr) {
        outTrace->clear();
    }
    if (fn == nullptr || mainEditorHost == nullptr) {
        if (outTrace != nullptr) {
            *outTrace = "hidden_open_invalid_argument";
        }
        return false;
    }

    HiddenOpenCodeTargetContext hookContext;
    ScopedHiddenOpenCodeTargetApiHooks hooks(hookContext);
    if (!hooks.IsInstalled()) {
        if (outTrace != nullptr) {
            *outTrace = "hidden_open_hook_install_failed";
        }
        return false;
    }

    void* editorObject = nullptr;
    if (!SafeCallOpenCodeTarget(fn, mainEditorHost, openType, arg2, arg3, &editorObject)) {
        if (outTrace != nullptr) {
            *outTrace = "hidden_open_exception";
        }
        return false;
    }

    if (outEditorObject != nullptr) {
        *outEditorObject = editorObject;
    }
    if (outTrace != nullptr) {
        *outTrace =
            std::string(editorObject != nullptr ? "hidden_open_ok" : "hidden_open_null") +
            "|blocked_focus_hwnd=" +
            std::to_string(reinterpret_cast<std::uintptr_t>(hookContext.blockedFocusHwnd));
    }
    return editorObject != nullptr;
}

bool TryResolveEditorObjectForProgramTreeItemDataNoActivate(
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

    std::string supportTrace;
    if (!HasUnsafeDirectGlobalSearchLayoutSupport(moduleBase, &supportTrace)) {
        if (outTrace != nullptr) {
            *outTrace = supportTrace;
        }
        return false;
    }

    const unsigned int itemTypeNibble = itemData >> 28;
    if (itemTypeNibble == 0 || itemTypeNibble == 15) {
        if (outTrace != nullptr) {
            *outTrace = "non_page_tree_item";
        }
        return false;
    }

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    const auto resolveLightweightByHiddenOpen = [&](const char* reason) -> bool {
        const auto& openAddrs = GetOpenCodeTargetAddresses(moduleBase);
        if (!openAddrs.ok) {
            if (outTrace != nullptr) {
                *outTrace =
                    std::string(reason == nullptr ? "native_addresses_unavailable" : reason) +
                    "|lightweight_open_code_target_unavailable|" +
                    openAddrs.trace;
            }
            return false;
        }

        int openType = 0;
        int openArg2 = -1;
        int openArg3 = -1;
        int resolvedIndex = -1;
        switch (itemTypeNibble) {
        case 1:
            resolvedIndex = static_cast<int>(((itemData >> 16) & 0x0FFFu)) - 1;
            if (resolvedIndex < 0) {
                if (outTrace != nullptr) {
                    *outTrace = "type1_resolved_index_invalid";
                }
                return false;
            }
            openType = 1;
            openArg2 = resolvedIndex;
            openArg3 = 0;
            break;
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
            resolvedIndex = static_cast<int>(itemData & 0x0FFFFFFFu);
            openType = 5;
            openArg2 = resolvedIndex;
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

        void* const mainEditorHost = PtrAbsolute<void>(moduleBase, openAddrs.mainEditorHost);
        if (mainEditorHost == nullptr) {
            if (outTrace != nullptr) {
                *outTrace =
                    std::string(reason == nullptr ? "native_addresses_unavailable" : reason) +
                    "|lightweight_main_editor_host_null|" +
                    openAddrs.trace;
            }
            return false;
        }

        const auto openCodeTarget = BindAbsolute<FnOpenCodeTarget>(moduleBase, openAddrs.openCodeTarget);
        void* editorObject = nullptr;
        std::string hiddenOpenTrace;
        if (!TryOpenCodeTargetNoActivate(
                openCodeTarget,
                mainEditorHost,
                openType,
                openArg2,
                openArg3,
                &editorObject,
                &hiddenOpenTrace) ||
            editorObject == nullptr) {
            if (outTrace != nullptr) {
                *outTrace =
                    std::string(reason == nullptr ? "native_addresses_unavailable" : reason) +
                    "|lightweight_hidden_open_failed|" +
                    (hiddenOpenTrace.empty() ? std::string("hidden_open_failed") : hiddenOpenTrace) +
                    "|" +
                    openAddrs.trace;
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
            *outBucketData = 0;
        }
        if (outTrace != nullptr) {
            *outTrace =
                std::string("open_editor_no_activate_lightweight_hidden_open") +
                "|" +
                (reason == nullptr ? "native_addresses_unavailable" : reason) +
                "|open_type=" + std::to_string(openType) +
                "|arg2=" + std::to_string(openArg2) +
                "|arg3=" + std::to_string(openArg3) +
                "|resolved_index=" + std::to_string(resolvedIndex) +
                "|bucket_data=0|" +
                hiddenOpenTrace +
                "|" +
                openAddrs.trace;
        }
        return true;
    };

    if (!addrs.ok) {
        return resolveLightweightByHiddenOpen("resolve_native_addresses_failed");
    }
    if (addrs.createCodePage == 0) {
        return resolveLightweightByHiddenOpen("create_code_page_resolve_failed");
    }

    void* const mainEditorHost = PtrAbsolute<void>(moduleBase, addrs.mainEditorHost);
    if (mainEditorHost == nullptr) {
        if (outTrace != nullptr) {
            *outTrace = "main_editor_host_null";
        }
        return false;
    }

    const auto containerGetAt = BindAbsolute<FnContainerGetAt>(moduleBase, addrs.containerGetAt);
    const auto openCodeTarget = BindAbsolute<FnOpenCodeTarget>(moduleBase, addrs.openCodeTarget);
    const auto createCodePage = BindAbsolute<FnCreateCodePage>(moduleBase, addrs.createCodePage);

    int openType = 0;
    int resolvedIndex = -1;
    int bucketData = 0;
    int openArg2 = -1;
    int openArg3 = -1;
    int matchKey = 0;
    int createOwnerArg = 0;
    int createDataArg = 0;
    int createExtraArg = 0;

    switch (itemTypeNibble) {
    case 1: {
        resolvedIndex = static_cast<int>(((itemData >> 16) & 0x0FFFu)) - 1;
        if (resolvedIndex < 0) {
            if (outTrace != nullptr) {
                *outTrace = "type1_resolved_index_invalid";
            }
            return false;
        }

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
        openArg2 = resolvedIndex;
        openArg3 = 0;
        matchKey = bucketData;
        createOwnerArg = *reinterpret_cast<int*>(bucketData + kOffset_Type1BucketOwner);
        createDataArg = bucketData;
        break;
    }
    case 2:
        openType = 3;
        openArg2 = -1;
        openArg3 = -1;
        createOwnerArg = static_cast<int>(
            reinterpret_cast<std::uintptr_t>(
                reinterpret_cast<std::uint8_t*>(mainEditorHost) + kOffset_MainEditorHostCreateOwner));
        createDataArg = ResolveTypeDataRaw(3, moduleBase);
        break;
    case 3:
        openType = 2;
        openArg2 = -1;
        openArg3 = -1;
        createOwnerArg = static_cast<int>(
            reinterpret_cast<std::uintptr_t>(
                reinterpret_cast<std::uint8_t*>(mainEditorHost) + kOffset_MainEditorHostCreateOwner));
        createDataArg = ResolveTypeDataRaw(2, moduleBase);
        break;
    case 4:
        openType = 4;
        openArg2 = -1;
        openArg3 = -1;
        createOwnerArg = static_cast<int>(
            reinterpret_cast<std::uintptr_t>(
                reinterpret_cast<std::uint8_t*>(mainEditorHost) + kOffset_MainEditorHostCreateOwner));
        createDataArg = ResolveTypeDataRaw(4, moduleBase);
        break;
    case 5: {
        openType = 5;
        resolvedIndex = static_cast<int>(itemData & 0x0FFFFFFFu);
        if (resolvedIndex < 0) {
            if (outTrace != nullptr) {
                *outTrace = "type5_resolved_index_invalid";
            }
            return false;
        }

        int formData = 0;
        int formOk = 0;
        if (!SafeContainerGetAt(
                containerGetAt,
                reinterpret_cast<std::uint8_t*>(mainEditorHost) + kOffset_MainEditorHostFormArray,
                resolvedIndex,
                &formData,
                &formOk) ||
            formOk == 0 ||
            formData == 0) {
            if (outTrace != nullptr) {
                *outTrace = "type5_form_lookup_failed";
            }
            return false;
        }

        matchKey = formData;
        openArg2 = resolvedIndex;
        openArg3 = -1;
        createOwnerArg = *reinterpret_cast<int*>(formData);
        createDataArg = formData;
        break;
    }
    case 6:
        openType = 6;
        openArg2 = -1;
        openArg3 = -1;
        createOwnerArg = static_cast<int>(
            reinterpret_cast<std::uintptr_t>(
                reinterpret_cast<std::uint8_t*>(mainEditorHost) + kOffset_MainEditorHostCreateOwner));
        createDataArg = ResolveTypeDataRaw(6, moduleBase);
        break;
    case 7: {
        const unsigned int subType = itemData & 0x0FFFFFFFu;
        openType = (subType == 1u) ? 7 : 8;
        openArg2 = -1;
        openArg3 = -1;
        createOwnerArg = static_cast<int>(
            reinterpret_cast<std::uintptr_t>(
                reinterpret_cast<std::uint8_t*>(mainEditorHost) + kOffset_MainEditorHostCreateOwner));
        createDataArg = ResolveTypeDataRaw(openType, moduleBase);
        break;
    }
    default:
        if (outTrace != nullptr) {
            *outTrace = "unsupported_tree_item_type";
        }
        return false;
    }

    if (createOwnerArg == 0 || createDataArg == 0) {
        if (outTrace != nullptr) {
            *outTrace = "create_args_invalid";
        }
        return false;
    }

    void* editorObject = nullptr;
    void* pageObject = nullptr;
    std::string directTrace;
    std::string hiddenOpenTrace;
    if (SafeFindExistingOpenEditorObject(mainEditorHost, openType, matchKey, &pageObject) &&
        pageObject != nullptr &&
        SafeResolveRuntimeEditorObjectFromPageObject(
            mainEditorHost,
            addrs.mainEditorHostActiveEditorObjectOffset,
            pageObject,
            &editorObject) &&
        editorObject != nullptr) {
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
            *outTrace = "open_editor_no_activate_existing";
        }
        return true;
    }

    pageObject = nullptr;
    if (SafeCreateEditorObjectNoActivate(
            createCodePage,
            mainEditorHost,
            createOwnerArg,
            openType,
            createDataArg,
            createExtraArg,
            &pageObject) &&
        pageObject != nullptr) {
        editorObject = nullptr;
        if (SafeResolveRuntimeEditorObjectFromPageObject(
                mainEditorHost,
                addrs.mainEditorHostActiveEditorObjectOffset,
                pageObject,
                &editorObject) &&
            editorObject != nullptr) {
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
                *outTrace = "open_editor_no_activate_created";
            }
            return true;
        }

        directTrace = "open_editor_no_activate_created_runtime_unavailable";
    } else {
        directTrace = "create_editor_no_activate_failed";
    }

    editorObject = nullptr;
    if (TryOpenCodeTargetNoActivate(
            openCodeTarget,
            mainEditorHost,
            openType,
            openArg2,
            openArg3,
            &editorObject,
            &hiddenOpenTrace) &&
        editorObject != nullptr) {
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
            *outTrace = directTrace.empty()
                ? ("open_editor_no_activate_hidden_open|" + hiddenOpenTrace)
                : (directTrace + "|fallback_hidden_open|" + hiddenOpenTrace);
        }
        return true;
    }

    if (outTrace != nullptr) {
        *outTrace = directTrace.empty()
            ? (hiddenOpenTrace.empty() ? "open_editor_no_activate_failed" : ("hidden_open_failed|" + hiddenOpenTrace))
            : (hiddenOpenTrace.empty() ? directTrace : (directTrace + "|hidden_open_failed|" + hiddenOpenTrace));
    }
    return false;
}

bool DebugSetMainEditorActiveEditorObjectImpl(
    std::uintptr_t moduleBase,
    std::uintptr_t editorObject,
    int notifyMode,
    std::uintptr_t* outPreviousEditorObject,
    std::string* outTrace) {
    if (outPreviousEditorObject != nullptr) {
        *outPreviousEditorObject = 0;
    }
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    std::string supportTrace;
    if (!HasUnsafeDirectGlobalSearchLayoutSupport(moduleBase, &supportTrace)) {
        if (outTrace != nullptr) {
            *outTrace = supportTrace;
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

    void* const mainEditorHost = PtrAbsolute<void>(moduleBase, addrs.mainEditorHost);
    if (mainEditorHost == nullptr) {
        if (outTrace != nullptr) {
            *outTrace = "main_editor_host_null";
        }
        return false;
    }

    if (addrs.setMainEditorActiveEditorObject == 0) {
        if (outTrace != nullptr) {
            *outTrace = "set_active_editor_object_resolve_failed";
        }
        return false;
    }
    if (addrs.mainEditorHostActiveEditorObjectOffset <= 0) {
        if (outTrace != nullptr) {
            *outTrace = "main_editor_host_active_editor_offset_unresolved";
        }
        return false;
    }

    const auto setActiveEditorObject = BindAbsolute<FnSetMainEditorActiveEditorObject>(
        moduleBase,
        addrs.setMainEditorActiveEditorObject);
    std::uintptr_t previousEditorObject = 0;
    std::uintptr_t currentEditorObject = 0;
    if (!SafeSetMainEditorActiveEditorObject(
            setActiveEditorObject,
            mainEditorHost,
            addrs.mainEditorHostActiveEditorObjectOffset,
            reinterpret_cast<void*>(editorObject),
            notifyMode,
            &previousEditorObject,
            &currentEditorObject)) {
        if (outTrace != nullptr) {
            *outTrace = "set_active_editor_object_exception";
        }
        return false;
    }

    if (outPreviousEditorObject != nullptr) {
        *outPreviousEditorObject = previousEditorObject;
    }
    if (outTrace != nullptr) {
        *outTrace =
            "set_active_editor_object_ok"
            "|notify=" + std::to_string(notifyMode) +
            "|previous=" + std::to_string(previousEditorObject) +
            "|current=" + std::to_string(currentEditorObject) +
            "|requested=" + std::to_string(editorObject);
    }
    return true;
}

bool DebugGetMainEditorActiveEditorObjectImpl(
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    std::string* outTrace) {
    if (outEditorObject != nullptr) {
        *outEditorObject = 0;
    }
    if (outTrace != nullptr) {
        outTrace->clear();
    }

    const auto& addrs = GetNativeSearchAddresses(moduleBase);
    if (!addrs.ok) {
        if (outTrace != nullptr) {
            *outTrace = "resolve_native_addresses_failed";
        }
        return false;
    }

    void* const mainEditorHost = PtrAbsolute<void>(moduleBase, addrs.mainEditorHost);
    if (mainEditorHost == nullptr) {
        if (outTrace != nullptr) {
            *outTrace = "main_editor_host_null";
        }
        return false;
    }

    if (addrs.mainEditorHostActiveEditorObjectOffset <= 0) {
        if (outTrace != nullptr) {
            *outTrace = "main_editor_host_active_editor_offset_unresolved";
        }
        return false;
    }

    std::uintptr_t editorObject = 0;
    if (!SafeReadMainEditorActiveEditorObject(
            mainEditorHost,
            addrs.mainEditorHostActiveEditorObjectOffset,
            &editorObject)) {
        if (outTrace != nullptr) {
            *outTrace = "read_active_editor_object_exception";
        }
        return false;
    }

    if (outEditorObject != nullptr) {
        *outEditorObject = editorObject;
    }
    if (outTrace != nullptr) {
        *outTrace = "get_active_editor_object_ok|current=" + std::to_string(editorObject);
    }
    return editorObject != 0;
}

LRESULT WINAPI HiddenOpenCodeTargetHook_SendMessageA(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    HiddenOpenCodeTargetContext* ctx = g_hiddenOpenCodeTargetContext;
    if (ctx != nullptr && ctx->suppressMdiActivate && msg == WM_MDIACTIVATE) {
        return 0;
    }

    return g_originalHiddenOpenSendMessageA(hwnd, msg, wParam, lParam);
}

HWND WINAPI HiddenOpenCodeTargetHook_SetFocus(HWND hwnd) {
    HiddenOpenCodeTargetContext* ctx = g_hiddenOpenCodeTargetContext;
    if (ctx != nullptr && ctx->suppressSetFocus) {
        ctx->blockedFocusHwnd = hwnd;
        return ::GetFocus();
    }

    return g_originalHiddenOpenSetFocus(hwnd);
}

ScopedHiddenOpenCodeTargetApiHooks::ScopedHiddenOpenCodeTargetApiHooks(HiddenOpenCodeTargetContext& ctx)
    : ctx_(ctx), installed_(false) {
    g_originalHiddenOpenSendMessageA = ::SendMessageA;
    g_originalHiddenOpenSetFocus = ::SetFocus;
    g_hiddenOpenCodeTargetContext = &ctx_;

    DetourTransactionBegin();
    DetourUpdateThread(::GetCurrentThread());
    DetourAttach(&(PVOID&)g_originalHiddenOpenSendMessageA, HiddenOpenCodeTargetHook_SendMessageA);
    DetourAttach(&(PVOID&)g_originalHiddenOpenSetFocus, HiddenOpenCodeTargetHook_SetFocus);
    installed_ = (DetourTransactionCommit() == NO_ERROR);
    if (!installed_) {
        g_hiddenOpenCodeTargetContext = nullptr;
    }
}

ScopedHiddenOpenCodeTargetApiHooks::~ScopedHiddenOpenCodeTargetApiHooks() {
    if (installed_) {
        DetourTransactionBegin();
        DetourUpdateThread(::GetCurrentThread());
        DetourDetach(&(PVOID&)g_originalHiddenOpenSendMessageA, HiddenOpenCodeTargetHook_SendMessageA);
        DetourDetach(&(PVOID&)g_originalHiddenOpenSetFocus, HiddenOpenCodeTargetHook_SetFocus);
        DetourTransactionCommit();
    }
    g_hiddenOpenCodeTargetContext = nullptr;
}

bool ScopedHiddenOpenCodeTargetApiHooks::IsInstalled() const {
    return installed_;
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

}  // namespace

bool e571::DebugSetMainEditorActiveEditorObject(
    std::uintptr_t moduleBase,
    std::uintptr_t editorObject,
    int notifyMode,
    std::uintptr_t* outPreviousEditorObject,
    std::string* outTrace) {
    return DebugSetMainEditorActiveEditorObjectImpl(
        moduleBase,
        editorObject,
        notifyMode,
        outPreviousEditorObject,
        outTrace);
}

bool e571::DebugGetMainEditorActiveEditorObject(
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    std::string* outTrace) {
    return DebugGetMainEditorActiveEditorObjectImpl(moduleBase, outEditorObject, outTrace);
}

}  // namespace e571

bool e571::DebugIsDirectGlobalSearchSupported(
    std::uintptr_t moduleBase,
    std::string* outTrace) {
    return DebugIsDirectGlobalSearchSupportedImpl(moduleBase, outTrace);
}

bool e571::DebugResolveEditorObjectByProgramTreeItemData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    int* outResolvedType,
    int* outResolvedIndex,
    int* outBucketData,
    std::string* outTrace) {
    if (outEditorObject != nullptr) {
        *outEditorObject = 0;
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
    if (outEditorObject != nullptr) {
        *outEditorObject = reinterpret_cast<std::uintptr_t>(editorObject);
    }
    if (outResolvedType != nullptr) {
        *outResolvedType = resolvedType;
    }
    if (outResolvedIndex != nullptr) {
        *outResolvedIndex = resolvedIndex;
    }
    if (outBucketData != nullptr) {
        *outBucketData = bucketData;
    }
    if (outTrace != nullptr) {
        *outTrace = trace;
    }
    return ok && editorObject != nullptr;
}

bool e571::DebugResolveEditorObjectByProgramTreeItemDataNoActivate(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    int* outResolvedType,
    int* outResolvedIndex,
    int* outBucketData,
    std::string* outTrace) {
    if (outEditorObject != nullptr) {
        *outEditorObject = 0;
    }

    void* editorObject = nullptr;
    int resolvedType = 0;
    int resolvedIndex = -1;
    int bucketData = 0;
    std::string trace;
    const bool ok = TryResolveEditorObjectForProgramTreeItemDataNoActivate(
        itemData,
        moduleBase,
        &editorObject,
        &resolvedType,
        &resolvedIndex,
        &bucketData,
        &trace);
    if (outEditorObject != nullptr) {
        *outEditorObject = reinterpret_cast<std::uintptr_t>(editorObject);
    }
    if (outResolvedType != nullptr) {
        *outResolvedType = resolvedType;
    }
    if (outResolvedIndex != nullptr) {
        *outResolvedIndex = resolvedIndex;
    }
    if (outBucketData != nullptr) {
        *outBucketData = bucketData;
    }
    if (outTrace != nullptr) {
        *outTrace = trace;
    }
    return ok && editorObject != nullptr;
}
