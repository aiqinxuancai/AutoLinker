#pragma once

#include <cstdint>
#include <string>

namespace e571 {

// 易语言IDE全局搜索内部桥接：仅暴露当前真实页读写链路需要的编辑器对象接口。
bool DebugIsDirectGlobalSearchSupported(
    std::uintptr_t moduleBase = 0x400000,
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

}  // namespace e571
