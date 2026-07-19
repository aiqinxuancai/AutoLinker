#include "IdeOutputTabController.h"

#include <CommCtrl.h>

#include <algorithm>
#include <array>
#include <format>
#include <string>

#include "Logger.h"

namespace IdeOutputTabController {
namespace {

constexpr wchar_t kTabControlClassName[] = L"SysTabControl32";
constexpr int kMaximumReasonableItemCount = 128;

bool IsTabControl(HWND window) noexcept
{
	if (window == nullptr || !IsWindow(window)) {
		return false;
	}
	wchar_t className[64] = {};
	return GetClassNameW(
		window,
		className,
		static_cast<int>(std::size(className))) > 0 &&
		wcscmp(className, kTabControlClassName) == 0;
}

std::wstring LocalToWide(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (length <= 0) {
		return {};
	}
	std::wstring wide(static_cast<std::size_t>(length), L'\0');
	if (MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		length) <= 0) {
		return {};
	}
	return wide;
}

bool ItemCaptionEquals(HWND tabWindow, int index, const std::string& localCaption)
{
	const bool unicode = SendMessageW(tabWindow, CCM_GETUNICODEFORMAT, 0, 0) != FALSE;
	if (unicode) {
		std::array<wchar_t, 512> text = {};
		TCITEMW item = {};
		item.mask = TCIF_TEXT;
		item.pszText = text.data();
		item.cchTextMax = static_cast<int>(text.size());
		return SendMessageW(
			tabWindow,
			TCM_GETITEMW,
			static_cast<WPARAM>(index),
			reinterpret_cast<LPARAM>(&item)) != FALSE &&
			LocalToWide(localCaption) == text.data();
	}

	std::array<char, 512> text = {};
	TCITEMA item = {};
	item.mask = TCIF_TEXT;
	item.pszText = text.data();
	item.cchTextMax = static_cast<int>(text.size());
	return SendMessageA(
		tabWindow,
		TCM_GETITEMA,
		static_cast<WPARAM>(index),
		reinterpret_cast<LPARAM>(&item)) != FALSE &&
		localCaption == text.data();
}

int FindItemByCaption(HWND tabWindow, const std::string& caption)
{
	const int itemCount = static_cast<int>(
		SendMessageW(tabWindow, TCM_GETITEMCOUNT, 0, 0));
	if (itemCount <= 0 || itemCount > kMaximumReasonableItemCount) {
		return -1;
	}
	for (int index = 0; index < itemCount; ++index) {
		if (ItemCaptionEquals(tabWindow, index, caption)) {
			return index;
		}
	}
	return -1;
}

HWND FindOwningTab(HWND mainWindow, HWND pageWindow, const std::string& caption)
{
	const HWND parent = GetParent(pageWindow);
	if (IsTabControl(parent) && FindItemByCaption(parent, caption) >= 0) {
		return parent;
	}

	// 一些旧版宿主把新增页面与 Tab 作为同级子窗口；优先检查页面父窗口，
	// 再对主窗口做一次递归枚举，标题匹配用于避免触碰无关 Tab 控件。
	if (parent != nullptr) {
		for (HWND tab = FindWindowExW(parent, nullptr, kTabControlClassName, nullptr);
			tab != nullptr;
			tab = FindWindowExW(parent, tab, kTabControlClassName, nullptr)) {
			if (FindItemByCaption(tab, caption) >= 0) {
				return tab;
			}
		}
	}

	struct SearchContext {
		const std::string* caption = nullptr;
		HWND result = nullptr;
	} search{&caption, nullptr};
	EnumChildWindows(
		mainWindow,
		[](HWND window, LPARAM parameter) -> BOOL {
			auto* context = reinterpret_cast<SearchContext*>(parameter);
			if (context != nullptr && IsTabControl(window) &&
				FindItemByCaption(window, *context->caption) >= 0) {
				context->result = window;
				return FALSE;
			}
			return TRUE;
		},
		reinterpret_cast<LPARAM>(&search));
	return search.result;
}

bool ClickTabItem(HWND tabWindow, int index) noexcept
{
	RECT itemRect = {};
	if (index < 0 || SendMessageW(
		tabWindow,
		TCM_GETITEMRECT,
		static_cast<WPARAM>(index),
		reinterpret_cast<LPARAM>(&itemRect)) == FALSE ||
		itemRect.right <= itemRect.left || itemRect.bottom <= itemRect.top) {
		return false;
	}
	const int x = itemRect.left + (itemRect.right - itemRect.left) / 2;
	const int y = itemRect.top + (itemRect.bottom - itemRect.top) / 2;
	SendMessageW(tabWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(x, y));
	SendMessageW(tabWindow, WM_LBUTTONUP, 0, MAKELPARAM(x, y));
	return true;
}

int InsertTabItem(HWND tabWindow, int index, const std::string& localCaption)
{
	const bool unicode = SendMessageW(tabWindow, CCM_GETUNICODEFORMAT, 0, 0) != FALSE;
	if (unicode) {
		std::wstring text = LocalToWide(localCaption);
		if (text.empty()) {
			return -1;
		}
		TCITEMW item = {};
		item.mask = TCIF_TEXT;
		item.pszText = text.data();
		return static_cast<int>(SendMessageW(
			tabWindow,
			TCM_INSERTITEMW,
			static_cast<WPARAM>(index),
			reinterpret_cast<LPARAM>(&item)));
	}

	TCITEMA item = {};
	item.mask = TCIF_TEXT;
	item.pszText = const_cast<char*>(localCaption.c_str());
	return static_cast<int>(SendMessageA(
		tabWindow,
		TCM_INSERTITEMA,
		static_cast<WPARAM>(index),
		reinterpret_cast<LPARAM>(&item)));
}

} // namespace

bool HideTabByCaption(
	HWND mainWindow,
	HWND pageWindow,
	const std::string& caption,
	HiddenTabState& state)
{
	if (mainWindow == nullptr || !IsWindow(mainWindow) ||
		pageWindow == nullptr || !IsWindow(pageWindow) || caption.empty()) {
		return false;
	}

	const HWND tabWindow = FindOwningTab(mainWindow, pageWindow, caption);
	const int itemIndex = tabWindow != nullptr ? FindItemByCaption(tabWindow, caption) : -1;
	const int itemCount = tabWindow != nullptr
		? static_cast<int>(SendMessageW(tabWindow, TCM_GETITEMCOUNT, 0, 0))
		: 0;
	if (tabWindow == nullptr || itemIndex < 0 || itemCount <= 0) {
		Logger::Instance().Write("IdeOutputTabController", "output tab was not found");
		return false;
	}

	const int selectedIndex = static_cast<int>(
		SendMessageW(tabWindow, TCM_GETCURSEL, 0, 0));
	if (selectedIndex == itemIndex && itemCount > 1) {
		const int fallbackIndex = itemIndex > 0 ? itemIndex - 1 : 1;
		ClickTabItem(tabWindow, fallbackIndex);
	}

	if (SendMessageW(
		tabWindow,
		TCM_DELETEITEM,
		static_cast<WPARAM>(itemIndex),
		0) == FALSE) {
		Logger::Instance().Write("IdeOutputTabController", "standard output tab removal failed");
		return false;
	}

	state.tabWindow = tabWindow;
	state.itemIndex = itemIndex;
	state.hidden = true;
	Logger::Instance().Write(
		"IdeOutputTabController",
		std::format(
			"hid standard output tab index={} count_before={}",
			itemIndex,
			itemCount));
	return true;
}

bool RestoreHiddenTab(
	HWND pageWindow,
	const std::string& caption,
	HiddenTabState& state)
{
	if (pageWindow == nullptr || !IsWindow(pageWindow) || caption.empty() ||
		!state.hidden || !IsTabControl(state.tabWindow)) {
		return false;
	}

	int itemIndex = FindItemByCaption(state.tabWindow, caption);
	if (itemIndex < 0) {
		const int itemCount = static_cast<int>(
			SendMessageW(state.tabWindow, TCM_GETITEMCOUNT, 0, 0));
		if (itemCount < 0 || itemCount > kMaximumReasonableItemCount) {
			return false;
		}
		const int requestedIndex = (std::clamp)(state.itemIndex, 0, itemCount);
		itemIndex = InsertTabItem(state.tabWindow, requestedIndex, caption);
		if (itemIndex < 0) {
			Logger::Instance().Write("IdeOutputTabController", "restore output tab insertion failed");
			return false;
		}
	}

	ShowWindow(pageWindow, SW_SHOW);
	BringWindowToTop(pageWindow);
	if (!ClickTabItem(state.tabWindow, itemIndex)) {
		SendMessageW(
			state.tabWindow,
			TCM_SETCURSEL,
			static_cast<WPARAM>(itemIndex),
			0);
	}
	InvalidateRect(state.tabWindow, nullptr, TRUE);
	state.itemIndex = itemIndex;
	state.hidden = false;
	Logger::Instance().Write(
		"IdeOutputTabController",
		std::format("restored standard output tab index={}", itemIndex));
	return true;
}

} // namespace IdeOutputTabController
