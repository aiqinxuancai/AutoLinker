#pragma once

#include <cstdint>
#include <string>

namespace e571 {

// 判断当前易语言 IDE 版本是否支持内部编辑器对象解析。
bool IsEditorObjectResolverSupported(
	std::uintptr_t moduleBase = 0x400000,
	std::string* outTrace = nullptr);

// 解析程序树项目对应的编辑器对象，必要时允许 IDE 激活目标页面。
bool ResolveEditorObjectByProgramTreeItemData(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	int* outResolvedType = nullptr,
	int* outResolvedIndex = nullptr,
	int* outBucketData = nullptr,
	std::string* outTrace = nullptr);

// 解析程序树项目对应的编辑器对象，不主动激活目标页面。
bool ResolveEditorObjectByProgramTreeItemDataNoActivate(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	int* outResolvedType = nullptr,
	int* outResolvedIndex = nullptr,
	int* outBucketData = nullptr,
	std::string* outTrace = nullptr);

// 设置主编辑器当前活动对象，并可返回切换前的对象。
bool SetMainEditorActiveEditorObject(
	std::uintptr_t moduleBase,
	std::uintptr_t editorObject,
	int notifyMode,
	std::uintptr_t* outPreviousEditorObject = nullptr,
	std::string* outTrace = nullptr);

// 获取主编辑器当前活动对象。
bool GetMainEditorActiveEditorObject(
	std::uintptr_t moduleBase,
	std::uintptr_t* outEditorObject,
	std::string* outTrace = nullptr);

} // namespace e571
