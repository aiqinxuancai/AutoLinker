#include "AIConfigDialog.h"

#include <array>
#include <CommCtrl.h>

#pragma comment(lib, "comctl32.lib")
#if defined _M_IX86
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

namespace {
constexpr int IDC_CFG_BASE_URL = 1001;
constexpr int IDC_CFG_API_KEY = 1002;
constexpr int IDC_CFG_MODEL = 1003;
constexpr int IDC_CFG_EXTRA_PROMPT = 1004;
constexpr int IDC_CFG_SAVE = 1;
constexpr int IDC_CFG_CANCEL = 2;

constexpr int IDC_PREVIEW_EDIT = 1101;
constexpr int IDC_PREVIEW_OK = 1;
constexpr int IDC_PREVIEW_CANCEL = 2;

constexpr int IDC_INPUT_EDIT = 1201;
constexpr int IDC_INPUT_OK = 1;
constexpr int IDC_INPUT_CANCEL = 2;

HFONT GetDialogFont()
{
	static HFONT s_dialogFont = nullptr;
	if (s_dialogFont != nullptr) {
		return s_dialogFont;
	}

	NONCLIENTMETRICSA ncm = {};
	ncm.cbSize = sizeof(ncm);
	if (SystemParametersInfoA(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0) != FALSE) {
		s_dialogFont = CreateFontIndirectA(&ncm.lfMessageFont);
	}
	if (s_dialogFont == nullptr) {
		s_dialogFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
	}
	return s_dialogFont;
}

std::string GetEditTextA(HWND hEdit)
{
	if (hEdit == nullptr) {
		return std::string();
	}

	int len = GetWindowTextLengthA(hEdit);
	if (len <= 0) {
		return std::string();
	}

	std::string text(static_cast<size_t>(len + 1), '\0');
	GetWindowTextA(hEdit, &text[0], len + 1);
	text.resize(static_cast<size_t>(len));
	return text;
}

std::string NormalizeMultilineForEdit(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}

	std::string expanded = text;
	const bool hasRealLineBreak = expanded.find('\n') != std::string::npos || expanded.find('\r') != std::string::npos;
	if (!hasRealLineBreak) {
		std::string unescaped;
		unescaped.reserve(expanded.size());
		for (size_t i = 0; i < expanded.size(); ++i) {
			if (expanded[i] == '\\' && i + 1 < expanded.size()) {
				if (expanded[i + 1] == 'n') {
					unescaped.push_back('\n');
					++i;
					continue;
				}
				if (expanded[i + 1] == 'r') {
					if (i + 3 < expanded.size() && expanded[i + 2] == '\\' && expanded[i + 3] == 'n') {
						unescaped.push_back('\n');
						i += 3;
						continue;
					}
					unescaped.push_back('\n');
					++i;
					continue;
				}
			}
			unescaped.push_back(expanded[i]);
		}
		expanded.swap(unescaped);
	}

	std::string normalized;
	normalized.reserve(expanded.size() + 16);
	for (size_t i = 0; i < expanded.size(); ++i) {
		const char ch = expanded[i];
		if (ch == '\r') {
			if (i + 1 < expanded.size() && expanded[i + 1] == '\n') {
				++i;
			}
			normalized += "\r\n";
			continue;
		}
		if (ch == '\n') {
			normalized += "\r\n";
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

void SetDefaultFont(HWND hWnd)
{
	if (hWnd == nullptr) {
		return;
	}
	SendMessageA(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetDialogFont()), TRUE);
}

struct AIConfigDialogContext {
	AISettings* settings = nullptr;
	bool accepted = false;
	HWND hBaseUrl = nullptr;
	HWND hApiKey = nullptr;
	HWND hModel = nullptr;
	HWND hExtraPrompt = nullptr;
};

LRESULT CALLBACK AIConfigDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIConfigDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTA*>(lParam);
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr || ctx->settings == nullptr) {
			return -1;
		}

		HWND hBaseUrlLabel = CreateWindowA("STATIC", "Base URL:", WS_CHILD | WS_VISIBLE,
			16, 16, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hBaseUrl = CreateWindowExA(0, "EDIT", ctx->settings->baseUrl.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 14, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_BASE_URL), nullptr, nullptr);

		HWND hApiKeyLabel = CreateWindowA("STATIC", "API Key:", WS_CHILD | WS_VISIBLE,
			16, 50, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hApiKey = CreateWindowExA(0, "EDIT", ctx->settings->apiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 48, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_API_KEY), nullptr, nullptr);

		HWND hModelLabel = CreateWindowA("STATIC", "Model:", WS_CHILD | WS_VISIBLE,
			16, 84, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hModel = CreateWindowExA(0, "EDIT", ctx->settings->model.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 82, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_MODEL), nullptr, nullptr);

		HWND hExtraPromptLabel = CreateWindowA("STATIC", "System Prompt Extra:", WS_CHILD | WS_VISIBLE,
			16, 118, 140, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hExtraPrompt = CreateWindowExA(0, "EDIT", ctx->settings->extraSystemPrompt.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			120, 118, 500, 180, hWnd, reinterpret_cast<HMENU>(IDC_CFG_EXTRA_PROMPT), nullptr, nullptr);

		HWND hSave = CreateWindowA("BUTTON", "Save and Continue", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			420, 314, 100, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_SAVE), nullptr, nullptr);
		HWND hCancel = CreateWindowA("BUTTON", "Cancel", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			530, 314, 90, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_CANCEL), nullptr, nullptr);

		std::array<HWND, 10> controls = {
			hBaseUrlLabel,
			hApiKeyLabel,
			hModelLabel,
			hExtraPromptLabel,
			ctx->hBaseUrl,
			ctx->hApiKey,
			ctx->hModel,
			ctx->hExtraPrompt,
			hSave,
			hCancel
		};
		for (HWND hControl : controls) {
			SetDefaultFont(hControl);
		}
		SetFocus(ctx->hBaseUrl);
		return 0;
	}

	case WM_COMMAND: {
		const int id = LOWORD(wParam);
		if (ctx == nullptr || ctx->settings == nullptr) {
			return 0;
		}

		if (id == IDC_CFG_SAVE) {
			AISettings next = *ctx->settings;
			next.baseUrl = GetEditTextA(ctx->hBaseUrl);
			next.apiKey = GetEditTextA(ctx->hApiKey);
			next.model = GetEditTextA(ctx->hModel);
			next.extraSystemPrompt = GetEditTextA(ctx->hExtraPrompt);

			if (next.baseUrl.empty() || next.apiKey.empty() || next.model.empty()) {
				MessageBoxA(hWnd, "baseUrl / apiKey / model cannot be empty.", "AI Config", MB_ICONWARNING | MB_OK);
				return 0;
			}

			*ctx->settings = next;
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}

		if (id == IDC_CFG_CANCEL) {
			ctx->accepted = false;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;
	default:
		break;
	}

	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

struct AIPreviewDialogContext {
	std::string title;
	std::string content;
	std::string confirmText;
	bool accepted = false;
	HWND hEdit = nullptr;
};

struct AIInputDialogContext {
	std::string title;
	std::string hint;
	std::string text;
	bool accepted = false;
	HWND hEdit = nullptr;
};

LRESULT CALLBACK AIPreviewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIPreviewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTA*>(lParam);
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr) {
			return -1;
		}

		ctx->hEdit = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", "",
			WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
			14, 14, 752, 450, hWnd, reinterpret_cast<HMENU>(IDC_PREVIEW_EDIT), nullptr, nullptr);
		const std::string displayText = NormalizeMultilineForEdit(ctx->content);
		SetWindowTextA(ctx->hEdit, displayText.c_str());

		HWND hOk = CreateWindowA("BUTTON", ctx->confirmText.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 590, 474, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_PREVIEW_OK), nullptr, nullptr);

		HWND hCancel = CreateWindowA("BUTTON", "Cancel",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 682, 474, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_PREVIEW_CANCEL), nullptr, nullptr);

		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(hOk);
		SetDefaultFont(hCancel);
		return 0;
	}

	case WM_COMMAND: {
		const int id = LOWORD(wParam);
		if (ctx == nullptr) {
			return 0;
		}
		if (id == IDC_PREVIEW_OK) {
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_PREVIEW_CANCEL) {
			ctx->accepted = false;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;
	default:
		break;
	}
	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK AIInputDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIInputDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTA*>(lParam);
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr) {
			return -1;
		}

		HWND hHint = CreateWindowA("STATIC", ctx->hint.c_str(),
			WS_CHILD | WS_VISIBLE, 14, 14, 612, 24, hWnd, nullptr, nullptr, nullptr);

		ctx->hEdit = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", ctx->text.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			14, 44, 612, 200, hWnd, reinterpret_cast<HMENU>(IDC_INPUT_EDIT), nullptr, nullptr);

		HWND hOk = CreateWindowA("BUTTON", "OK",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 448, 254, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_INPUT_OK), nullptr, nullptr);
		HWND hCancel = CreateWindowA("BUTTON", "Cancel",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 542, 254, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_INPUT_CANCEL), nullptr, nullptr);

		SetDefaultFont(hHint);
		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(hOk);
		SetDefaultFont(hCancel);
		SetFocus(ctx->hEdit);
		return 0;
	}

	case WM_COMMAND: {
		if (ctx == nullptr) {
			return 0;
		}
		const int id = LOWORD(wParam);
		if (id == IDC_INPUT_OK) {
			ctx->text = GetEditTextA(ctx->hEdit);
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_INPUT_CANCEL) {
			ctx->accepted = false;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;
	default:
		break;
	}
	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

bool RunModalWindow(HWND owner, HWND hDialog)
{
	if (hDialog == nullptr) {
		return false;
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, FALSE);
	}

	ShowWindow(hDialog, SW_SHOW);
	UpdateWindow(hDialog);

	MSG msg = {};
	while (IsWindow(hDialog)) {
		const BOOL ok = GetMessageA(&msg, nullptr, 0, 0);
		if (ok <= 0) {
			break;
		}
		if (!IsDialogMessageA(hDialog, &msg)) {
			TranslateMessage(&msg);
			DispatchMessageA(&msg);
		}
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, TRUE);
		SetForegroundWindow(owner);
	}
	return true;
}
} // namespace

bool ShowAIConfigDialog(HWND owner, AISettings& ioSettings)
{
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIConfigDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIConfigDialogWindow";
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIConfigDialogContext ctx = {};
	ctx.settings = &ioSettings;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 650, 390,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		return false;
	}

	RunModalWindow(owner, hDialog);
	return ctx.accepted;
}

bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText)
{
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIPreviewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIPreviewDialogWindow";
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIPreviewDialogContext ctx = {};
	ctx.title = title;
	ctx.content = content;
	ctx.confirmText = confirmText.empty() ? "OK" : confirmText;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		ctx.title.c_str(),
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_SIZEBOX,
		CW_USEDEFAULT, CW_USEDEFAULT, 800, 560,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		return false;
	}

	RunModalWindow(owner, hDialog);
	return ctx.accepted;
}

bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText)
{
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIInputDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIInputDialogWindow";
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIInputDialogContext ctx = {};
	ctx.title = title;
	ctx.hint = hint;
	ctx.text = ioText;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		ctx.title.c_str(),
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 650, 330,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		return false;
	}

	RunModalWindow(owner, hDialog);
	if (ctx.accepted) {
		ioText = ctx.text;
	}
	return ctx.accepted;
}
