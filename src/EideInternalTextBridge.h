#pragma once

#include <cstddef>
#include <cstdint>
#include <string>

namespace e571 {

// 真实源码整页读写结果。
struct NativeRealPageAccessResult {
	bool ok = false;
	std::uintptr_t editorObject = 0;
	bool usedClipboardEmulation = false;
	bool capturedCustomFormat = false;
	bool rollbackAttempted = false;
	bool rollbackSucceeded = false;
	size_t textBytes = 0;
	std::string trace;
};

// 按编辑器对象读取真实整页源码。
bool GetRealPageCodeByEditorObject(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	std::string* outCode,
	NativeRealPageAccessResult* outResult = nullptr);

// 按程序树页数据读取真实整页源码。
bool GetRealPageCodeByProgramTreeItemData(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	std::string* outCode,
	NativeRealPageAccessResult* outResult = nullptr);

// 按程序树页数据切换到对应页面。
bool OpenProgramTreeItemPageByData(
	unsigned int itemData,
	std::string* outTrace = nullptr);

// 按编辑器对象整页覆盖真实源码。
bool ReplaceRealPageCodeByEditorObject(
	std::uintptr_t editorObject,
	std::uintptr_t moduleBase,
	const std::string& newPageCode,
	const std::string* rollbackPageCode,
	NativeRealPageAccessResult* outResult = nullptr);

// 按程序树页数据整页覆盖真实源码。
bool ReplaceRealPageCodeByProgramTreeItemData(
	unsigned int itemData,
	std::uintptr_t moduleBase,
	const std::string& newPageCode,
	const std::string* rollbackPageCode,
	NativeRealPageAccessResult* outResult = nullptr);

}  // namespace e571
