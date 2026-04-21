// ResourceTextLoader.cpp - 资源文本加载工具实现
#include "ResourceTextLoader.h"

#include <Windows.h>

#include <string>

namespace {

constexpr WORD kHtmlResourceTypeId = 23;

HMODULE GetCurrentModuleHandle()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandle),
		&module);
	return module;
}

std::string LoadUtf8TextResourceInternal(int resourceId, const char* resourceType)
{
	const HMODULE module = GetCurrentModuleHandle();
	if (module == nullptr || resourceType == nullptr) {
		return std::string();
	}

	// HTML 在 RC 中会被编译为预定义资源类型 RT_HTML(#23)，
	// 但为兼容不同写法，这里同时兼容数值类型和字符串类型。
	HRSRC resourceInfo = FindResourceA(
		module,
		MAKEINTRESOURCEA(resourceId),
		MAKEINTRESOURCEA(kHtmlResourceTypeId));
	if (resourceInfo == nullptr) {
		resourceInfo = FindResourceA(module, MAKEINTRESOURCEA(resourceId), resourceType);
	}
	if (resourceInfo == nullptr) {
		return std::string();
	}

	const DWORD resourceSize = SizeofResource(module, resourceInfo);
	if (resourceSize == 0) {
		return std::string();
	}

	const HGLOBAL resourceData = LoadResource(module, resourceInfo);
	if (resourceData == nullptr) {
		return std::string();
	}

	const void* const lockedData = LockResource(resourceData);
	if (lockedData == nullptr) {
		return std::string();
	}

	std::string text(
		reinterpret_cast<const char*>(lockedData),
		reinterpret_cast<const char*>(lockedData) + resourceSize);
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		text.erase(0, 3);
	}
	return text;
}

} // namespace

std::string LoadUtf8HtmlResourceText(int resourceId)
{
	return LoadUtf8TextResourceInternal(resourceId, "HTML");
}
