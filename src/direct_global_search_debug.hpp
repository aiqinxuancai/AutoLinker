#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

namespace e571 {

struct DirectGlobalSearchDebugHit {
    int type = 0;
    int extra = 0;
    int outerIndex = 0;
    int innerIndex = 0;
    int matchOffset = 0;
    std::string displayText;
};

struct HiddenBuiltinSearchDebugResult {
    bool dialogHandled = false;
    size_t hits = 0;
    std::string firstResultText;
    std::vector<std::string> previewLines;
    bool hasRawFirstHit = false;
    DirectGlobalSearchDebugHit rawFirstHit;
    size_t rawHitCount = 0;
    bool hasDirectFirstHit = false;
    DirectGlobalSearchDebugHit directFirstHit;
    size_t directHitCount = 0;
    bool rawResolveOk = false;
    int rawResolvedIndex = -1;
    int rawBucketData = 0;
    std::string jumpTrace;
};

struct NativeEditorPageDumpDebugResult {
    bool ok = false;
    std::uintptr_t editorObject = 0;
    int outerCount = 0;
    size_t lineCount = 0;
    size_t fetchFailures = 0;
    std::string trace;
};

struct RawSearchContextPageDumpDebugResult {
    bool ok = false;
    int type = 0;
    int resolvedIndex = -1;
    int bucketData = 0;
    int outerCount = 0;
    size_t lineCount = 0;
    size_t fetchFailures = 0;
    std::string trace;
};

bool DebugIsDirectGlobalSearchSupported(
    std::uintptr_t moduleBase = 0x400000,
    std::string* outTrace = nullptr);

std::vector<DirectGlobalSearchDebugHit> DebugSearchDirectGlobalKeyword(
    const char* keyword,
    std::uintptr_t moduleBase = 0x400000);

bool DebugLocateFirstDirectGlobalKeyword(
    const char* keyword,
    std::uintptr_t moduleBase,
    DirectGlobalSearchDebugHit* outHit = nullptr,
    size_t* outTotalHits = nullptr);

HiddenBuiltinSearchDebugResult DebugSearchDirectGlobalKeywordHidden(
    const char* keyword,
    std::uintptr_t moduleBase = 0x400000);

std::vector<DirectGlobalSearchDebugHit> DebugSearchDirectGlobalKeywordHiddenDetailed(
    const char* keyword,
    std::uintptr_t moduleBase = 0x400000,
    bool* outDialogHandled = nullptr);

bool DebugLocateFirstDirectGlobalKeywordHidden(
    const char* keyword,
    std::uintptr_t moduleBase,
    HiddenBuiltinSearchDebugResult* outResult = nullptr);

bool DebugDumpCodePageForSearchHit(
    const DirectGlobalSearchDebugHit& hit,
    std::uintptr_t moduleBase,
    std::string* outCode,
    NativeEditorPageDumpDebugResult* outResult = nullptr);

bool DebugDumpCodePageByProgramTreeItemData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::string* outCode,
    RawSearchContextPageDumpDebugResult* outResult = nullptr);

bool DebugOpenProgramTreeItemByData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::string* outTrace = nullptr);

bool DebugResolveEditorObjectByProgramTreeItemData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    int* outResolvedType = nullptr,
    int* outResolvedIndex = nullptr,
    int* outBucketData = nullptr,
    std::string* outTrace = nullptr);

bool DebugResolveEditorObjectByProgramTreeItemDataNoActivate(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    int* outResolvedType = nullptr,
    int* outResolvedIndex = nullptr,
    int* outBucketData = nullptr,
    std::string* outTrace = nullptr);

bool DebugSetMainEditorActiveEditorObject(
    std::uintptr_t moduleBase,
    std::uintptr_t editorObject,
    int notifyMode,
    std::uintptr_t* outPreviousEditorObject = nullptr,
    std::string* outTrace = nullptr);

bool DebugGetMainEditorActiveEditorObject(
    std::uintptr_t moduleBase,
    std::uintptr_t* outEditorObject,
    std::string* outTrace = nullptr);

bool DebugJumpToSearchHit(
    const DirectGlobalSearchDebugHit& hit,
    std::uintptr_t moduleBase,
    std::string* outTrace = nullptr);

}  // namespace e571
