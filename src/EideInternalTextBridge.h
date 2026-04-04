#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

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

// 当前活动页编辑器对象解析结果。
struct ActiveEditorObjectInfo {
	bool ok = false;
	std::uintptr_t rawEditorObject = 0;
	std::uintptr_t innerEditorObject = 0;
	unsigned int pageType = 0;
	std::string trace;
};

// 内部 thiscall 命令抓取规格。
struct InternalThiscallCommandSpec {
	std::uintptr_t targetObject = 0;
	std::uintptr_t functionAddress = 0;
	std::uintptr_t arg1 = 0;
	std::uintptr_t arg2 = 0;
	std::uintptr_t arg3 = 0;
	std::uintptr_t arg4 = 0;
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

// 解析当前活动代码页的编辑器对象。
bool ResolveCurrentActiveEditorObject(
	std::uintptr_t moduleBase,
	ActiveEditorObjectInfo* outInfo);

// 执行内部 thiscall 命令并抓取写入自定义剪贴板格式的二进制。
bool CaptureCustomClipboardPayloadByThiscall(
	const InternalThiscallCommandSpec& spec,
	std::vector<unsigned char>& outBytes,
	std::string* outTrace = nullptr);

}  // namespace e571
