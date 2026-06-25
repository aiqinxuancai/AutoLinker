#include "AIConfigDialog.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <CommCtrl.h>
#include <filesystem>
#include <format>
#include <fstream>
#include <new>
#include <Shellapi.h>
#include <set>
#include <system_error>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "AIService.h"
#include "Global.h"
#include "PathHelper.h"
#include "ResourceTextLoader.h"
#include "WinINetUtil.h"

#pragma comment(lib, "comctl32.lib")
#if defined _M_IX86
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

namespace {
constexpr int IDC_CFG_PROTOCOL = 1000;
constexpr int IDC_CFG_BASE_URL = 1001;
constexpr int IDC_CFG_API_KEY = 1002;
constexpr int IDC_CFG_MODEL = 1003;
constexpr int IDC_CFG_EXTRA_PROMPT = 1004;
constexpr int IDC_CFG_GET_KEY_LINK = 1005;
constexpr int IDC_CFG_TAVILY_API_KEY = 1007;
constexpr int IDC_CFG_TEST_CONNECTION = 1009;
constexpr int IDC_CFG_THINKING_LEVEL = 1010;
constexpr int IDC_CFG_CUSTOM_HEADERS = 1011;
constexpr int IDC_CFG_PROFILE_COMBO = 1012;
constexpr int IDC_CFG_PROFILE_ADD = 1013;
constexpr int IDC_CFG_PROFILE_RENAME = 1014;
constexpr int IDC_CFG_PROFILE_DELETE = 1015;
constexpr int IDC_CFG_PROFILE_ADD_PRESET = 1016;
constexpr int IDC_CFG_SAVE = 1;
constexpr int IDC_CFG_CANCEL = 2;
constexpr UINT kAIConfigNativePresetMenuBase = 30000;

constexpr int IDC_PREVIEW_EDIT = 1101;
constexpr int IDC_PREVIEW_OK = 1;
constexpr int IDC_PREVIEW_SECONDARY = 1102;
constexpr int IDC_PREVIEW_CANCEL = 2;

constexpr int IDC_INPUT_EDIT = 1201;
constexpr int IDC_INPUT_OK = 1;
constexpr int IDC_INPUT_CANCEL = 2;
constexpr UINT_PTR kAIConfigWebViewInitTimerId = 0xAC01;
constexpr UINT kAIConfigWebViewInitTimeoutMs = 12000;
constexpr UINT WM_AUTOLINKER_AI_CONFIG_TEST_DONE = WM_APP + 301;
constexpr UINT WM_AUTOLINKER_AI_CONFIG_MODEL_LIST_DONE = WM_APP + 302;
constexpr UINT_PTR kAIPreviewWebViewInitTimerId = 0xAC02;
constexpr UINT kAIPreviewWebViewInitTimeoutMs = 12000;
constexpr UINT_PTR kLinkerConfigWebViewInitTimerId = 0xAC03;
constexpr UINT kLinkerConfigWebViewInitTimeoutMs = 12000;

HMODULE GetCurrentModuleHandle()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandle),
		&module);
	return module;
}
HICON LoadAppIconHandle(int cx, int cy)
{
	const HMODULE module = GetCurrentModuleHandle();
	if (module == nullptr) {
		return nullptr;
	}
	return reinterpret_cast<HICON>(LoadImageA(
		module,
		MAKEINTRESOURCEA(IDI_APP_ICON),
		IMAGE_ICON,
		cx,
		cy,
		LR_DEFAULTCOLOR));
}

HICON GetAppIconLarge()
{
	static HICON s_largeIcon = LoadAppIconHandle(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
	return s_largeIcon;
}

HICON GetAppIconSmall()
{
	static HICON s_smallIcon = LoadAppIconHandle(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
	return s_smallIcon;
}

void ApplyWindowIcon(HWND hWnd)
{
	if (hWnd == nullptr) {
		return;
	}
	SendMessageA(hWnd, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(GetAppIconLarge()));
	SendMessageA(hWnd, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(GetAppIconSmall()));
}

void EnsureWindowTitle(HWND hWnd, const std::string& fallbackTitle)
{
	if (hWnd == nullptr) {
		return;
	}
	if (GetWindowTextLengthA(hWnd) > 0) {
		return;
	}
	if (fallbackTitle.empty()) {
		SetWindowTextA(hWnd, "AutoLinker");
		return;
	}
	SetWindowTextA(hWnd, fallbackTitle.c_str());
}

class ComCtl6ActivationScope {
public:
	ComCtl6ActivationScope()
	{
		const HMODULE module = GetCurrentModuleHandle();
		if (module == nullptr) {
			return;
		}

		ACTCTXW act = {};
		act.cbSize = sizeof(act);
		act.dwFlags = ACTCTX_FLAG_HMODULE_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
		act.hModule = module;
		act.lpResourceName = MAKEINTRESOURCEW(2);

		m_actCtx = CreateActCtxW(&act);
		if (m_actCtx == INVALID_HANDLE_VALUE) {
			m_actCtx = nullptr;
			return;
		}

		ULONG_PTR localCookie = 0;
		if (ActivateActCtx(m_actCtx, &localCookie) != FALSE) {
			m_cookie = localCookie;
			m_active = true;
		}
	}

	~ComCtl6ActivationScope()
	{
		if (m_active) {
			DeactivateActCtx(0, m_cookie);
		}
		if (m_actCtx != nullptr && m_actCtx != INVALID_HANDLE_VALUE) {
			ReleaseActCtx(m_actCtx);
		}
	}

	ComCtl6ActivationScope(const ComCtl6ActivationScope&) = delete;
	ComCtl6ActivationScope& operator=(const ComCtl6ActivationScope&) = delete;

private:
	HANDLE m_actCtx = nullptr;
	ULONG_PTR m_cookie = 0;
	bool m_active = false;
};

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

void SetEditTextA(HWND hEdit, const std::string& text)
{
	if (hEdit == nullptr) {
		return;
	}
	SetWindowTextA(hEdit, text.c_str());
}

bool IsValidUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return text;
	}

	const int outLen = WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		out.data(),
		outLen,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return out;
}

std::string LocalToUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocalText(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	const int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::wstring WideFromLocal(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	const int wideLen = MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int utf8Len = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (utf8Len <= 0) {
		return std::string();
	}

	std::string utf8(static_cast<size_t>(utf8Len), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		utf8.data(),
		utf8Len,
		nullptr,
		nullptr) <= 0) {
		return std::string();
	}
	return utf8;
}

std::string JsonStringLiteral(const std::string& utf8Text)
{
	return nlohmann::json(utf8Text).dump();
}

bool IsWebView2RuntimeAvailableWithTag(const char* logTag)
{
	(void)logTag;
	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	const bool available = SUCCEEDED(hr);
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return available;
}

bool IsWebView2RuntimeAvailable()
{
	return IsWebView2RuntimeAvailableWithTag("AI Config");
}

struct AIConfigDialogRunResult {
	bool accepted = false;
	bool fallbackRequested = false;
};

struct AIConfigProfileEntry {
	std::string id;
	std::string name;
	AISettings settings;
};

struct AIPreviewWebViewRunResult {
	AIPreviewAction action = AIPreviewAction::Cancel;
	bool fallbackRequested = false;
};

struct AIConfigDialogContext {
	AIJsonConfig* jsonConfig = nullptr;
	AISettings* settings = nullptr;
	std::vector<AIConfigProfileEntry> profiles;
	std::string activeProfileId;
	bool accepted = false;
	bool useNativeLink = false;
	bool testInFlight = false;
	HWND hProfileCombo = nullptr;
	HWND hProtocol = nullptr;
	HWND hBaseUrl = nullptr;
	HWND hApiKey = nullptr;
	HWND hModel = nullptr;
	HWND hThinkingLevel = nullptr;
	HWND hTavilyApiKey = nullptr;
	HWND hExtraPrompt = nullptr;
	HWND hCustomHeaders = nullptr;
	HWND hGetKeyLink = nullptr;
	HWND hTestConnection = nullptr;
};

struct AIConfigWebViewDialogContext {
	AIJsonConfig* jsonConfig = nullptr;
	AISettings* settings = nullptr;
	std::vector<AIConfigProfileEntry> profiles;
	std::string activeProfileId;
	bool accepted = false;
	bool fallbackRequested = false;
	bool webViewReady = false;
	bool testInFlight = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct AIConfigConnectionTestRequest {
	HWND dialogHwnd = nullptr;
	AISettings settings;
	bool forWebView = false;
};

struct AIConfigConnectionTestResult {
	bool forWebView = false;
	AIResult result;
};

struct AIConfigModelListRequest {
	HWND dialogHwnd = nullptr;
	AISettings settings;
};

struct AIConfigModelListResult {
	bool ok = false;
	std::vector<std::string> models;
	std::string error;
	int httpStatus = 0;
};

struct AIPreviewWebViewDialogContext {
	std::string title;
	std::string content;
	std::string primaryText;
	std::string secondaryText;
	AIPreviewAction action = AIPreviewAction::Cancel;
	bool fallbackRequested = false;
	bool webViewReady = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

std::wstring EscapeJsSingleQuotedWide(const std::wstring& text)
{
	std::wstring out;
	out.reserve(text.size() + 32);
	for (wchar_t ch : text) {
		switch (ch)
		{
		case L'\\': out += L"\\\\"; break;
		case L'\'': out += L"\\'"; break;
		case L'\r': out += L"\\r"; break;
		case L'\n': out += L"\\n"; break;
		case 0x2028: out += L"\\u2028"; break;
		case 0x2029: out += L"\\u2029"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

void PopulateProtocolCombo(HWND hCombo, AIProtocolType selected)
{
	if (hCombo == nullptr) {
		return;
	}

	SendMessageA(hCombo, CB_RESETCONTENT, 0, 0);
	const int idxOpenAI = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("OpenAI Chat")));
	const int idxOpenAIResponses = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("OpenAI Responses")));
	const int idxGemini = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("Gemini")));
	const int idxClaude = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("Claude")));
	SendMessageA(hCombo, CB_SETITEMDATA, idxOpenAI, static_cast<LPARAM>(AIProtocolType::OpenAI));
	SendMessageA(hCombo, CB_SETITEMDATA, idxOpenAIResponses, static_cast<LPARAM>(AIProtocolType::OpenAIResponses));
	SendMessageA(hCombo, CB_SETITEMDATA, idxGemini, static_cast<LPARAM>(AIProtocolType::Gemini));
	SendMessageA(hCombo, CB_SETITEMDATA, idxClaude, static_cast<LPARAM>(AIProtocolType::Claude));

	int selectedIndex = idxOpenAI;
	if (selected == AIProtocolType::OpenAIResponses) {
		selectedIndex = idxOpenAIResponses;
	}
	else if (selected == AIProtocolType::Gemini) {
		selectedIndex = idxGemini;
	}
	else if (selected == AIProtocolType::Claude) {
		selectedIndex = idxClaude;
	}
	SendMessageA(hCombo, CB_SETCURSEL, selectedIndex, 0);
}

AIProtocolType GetSelectedProtocol(HWND hCombo)
{
	if (hCombo == nullptr) {
		return AIProtocolType::OpenAI;
	}

	const int selected = static_cast<int>(SendMessageA(hCombo, CB_GETCURSEL, 0, 0));
	if (selected == CB_ERR) {
		return AIProtocolType::OpenAI;
	}

	const LRESULT data = SendMessageA(hCombo, CB_GETITEMDATA, selected, 0);
	switch (static_cast<AIProtocolType>(data)) {
	case AIProtocolType::OpenAIResponses:
		return AIProtocolType::OpenAIResponses;
	case AIProtocolType::Gemini:
		return AIProtocolType::Gemini;
	case AIProtocolType::Claude:
		return AIProtocolType::Claude;
	case AIProtocolType::OpenAI:
	default:
		return AIProtocolType::OpenAI;
	}
}

void ApplyProfileValuesToSettings(const std::map<std::string, std::string>& values, AISettings& settings)
{
	settings.protocolType = AIService::ParseProtocolType([&]() {
		const auto it = values.find("protocol_type");
		return it != values.end() ? it->second : AIService::ProtocolTypeToString(settings.protocolType);
	}());
	settings.thinkingLevel = AIService::ParseThinkingLevel([&]() {
		const auto it = values.find("thinking_level");
		return it != values.end() ? it->second : AIService::ThinkingLevelToString(settings.thinkingLevel);
	}());
	if (const auto it = values.find("base_url"); it != values.end()) settings.baseUrl = it->second;
	if (const auto it = values.find("api_key"); it != values.end()) settings.apiKey = it->second;
	if (const auto it = values.find("model"); it != values.end()) settings.model = it->second;
	if (const auto it = values.find("system_prompt_extra"); it != values.end()) settings.extraSystemPrompt = it->second;
	if (const auto it = values.find("custom_headers"); it != values.end()) settings.customHeadersText = it->second;
	if (const auto it = values.find("tavily_api_key"); it != values.end()) settings.tavilyApiKey = it->second;
	if (const auto it = values.find("timeout_ms"); it != values.end()) {
		try { settings.timeoutMs = (std::max)(1000, std::stoi(it->second)); } catch (...) {}
	}
	if (const auto it = values.find("temperature"); it != values.end()) {
		try { settings.temperature = std::stod(it->second); } catch (...) {}
	}
	if (const auto it = values.find("context_window"); it != values.end()) {
		try { settings.contextWindowTokens = (std::max)(0, std::stoi(it->second)); } catch (...) {}
	}
}

std::map<std::string, std::string> BuildProfileValuesFromSettings(const AISettings& settings)
{
	return {
		{ "protocol_type", AIService::ProtocolTypeToString(settings.protocolType) },
		{ "thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel) },
		{ "base_url", settings.baseUrl },
		{ "api_key", settings.apiKey },
		{ "model", settings.model },
		{ "system_prompt_extra", settings.extraSystemPrompt },
		{ "custom_headers", settings.customHeadersText },
		{ "tavily_api_key", settings.tavilyApiKey },
		{ "timeout_ms", std::to_string(settings.timeoutMs) },
		{ "temperature", std::format("{:.2f}", settings.temperature) },
		{ "context_window", std::to_string(settings.contextWindowTokens) }
	};
}

std::vector<AIConfigProfileEntry> LoadProfileEntriesFromJsonConfig(AIJsonConfig& jsonConfig, AISettings& ioSettings, std::string& outActiveProfileId)
{
	std::vector<AIConfigProfileEntry> entries;
	const std::vector<AIJsonConfigProfileSnapshot> snapshots = jsonConfig.getProfilesLocal();
	outActiveProfileId = jsonConfig.getActiveProfileId();

	for (const auto& snapshot : snapshots) {
		AIConfigProfileEntry entry;
		entry.id = snapshot.id;
		entry.name = snapshot.name;
		entry.settings = {};
		ApplyProfileValuesToSettings(snapshot.values, entry.settings);
		entries.push_back(std::move(entry));
	}

	if (entries.empty()) {
		AIConfigProfileEntry fallback;
		fallback.id = "default";
		fallback.name = "默认";
		fallback.settings = ioSettings;
		entries.push_back(std::move(fallback));
		outActiveProfileId = "default";
	}

	bool foundActive = false;
	for (const auto& entry : entries) {
		if (entry.id == outActiveProfileId) {
			ioSettings = entry.settings;
			foundActive = true;
			break;
		}
	}
	if (!foundActive) {
		outActiveProfileId = entries.front().id;
		ioSettings = entries.front().settings;
	}
	return entries;
}

std::vector<AIJsonConfigProfileSnapshot> BuildProfileSnapshotsForSave(const std::vector<AIConfigProfileEntry>& entries)
{
	std::vector<AIJsonConfigProfileSnapshot> snapshots;
	snapshots.reserve(entries.size());
	for (const auto& entry : entries) {
		AIJsonConfigProfileSnapshot snapshot;
		snapshot.id = entry.id;
		snapshot.name = entry.name;
		snapshot.values = BuildProfileValuesFromSettings(entry.settings);
		snapshots.push_back(std::move(snapshot));
	}
	return snapshots;
}

int FindProfileIndexById(const std::vector<AIConfigProfileEntry>& profiles, const std::string& id)
{
	for (size_t i = 0; i < profiles.size(); ++i) {
		if (profiles[i].id == id) {
			return static_cast<int>(i);
		}
	}
	return -1;
}

void PopulateThinkingLevelCombo(HWND hCombo, AIThinkingLevel selected)
{
	if (hCombo == nullptr) {
		return;
	}

	SendMessageA(hCombo, CB_RESETCONTENT, 0, 0);
	const int idxOff = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("关闭")));
	const int idxLow = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("低")));
	const int idxMedium = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("中")));
	const int idxHigh = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("高")));
	SendMessageA(hCombo, CB_SETITEMDATA, idxOff, static_cast<LPARAM>(AIThinkingLevel::Off));
	SendMessageA(hCombo, CB_SETITEMDATA, idxLow, static_cast<LPARAM>(AIThinkingLevel::Low));
	SendMessageA(hCombo, CB_SETITEMDATA, idxMedium, static_cast<LPARAM>(AIThinkingLevel::Medium));
	SendMessageA(hCombo, CB_SETITEMDATA, idxHigh, static_cast<LPARAM>(AIThinkingLevel::High));

	int selectedIndex = idxOff;
	switch (selected) {
	case AIThinkingLevel::Low:
		selectedIndex = idxLow;
		break;
	case AIThinkingLevel::Medium:
		selectedIndex = idxMedium;
		break;
	case AIThinkingLevel::High:
		selectedIndex = idxHigh;
		break;
	case AIThinkingLevel::Off:
	default:
		selectedIndex = idxOff;
		break;
	}
	SendMessageA(hCombo, CB_SETCURSEL, selectedIndex, 0);
}

AIThinkingLevel GetSelectedThinkingLevel(HWND hCombo)
{
	if (hCombo == nullptr) {
		return AIThinkingLevel::Off;
	}

	const int selected = static_cast<int>(SendMessageA(hCombo, CB_GETCURSEL, 0, 0));
	if (selected == CB_ERR) {
		return AIThinkingLevel::Off;
	}

	const LRESULT data = SendMessageA(hCombo, CB_GETITEMDATA, selected, 0);
	switch (static_cast<AIThinkingLevel>(data)) {
	case AIThinkingLevel::Low:
		return AIThinkingLevel::Low;
	case AIThinkingLevel::Medium:
		return AIThinkingLevel::Medium;
	case AIThinkingLevel::High:
		return AIThinkingLevel::High;
	case AIThinkingLevel::Off:
	default:
		return AIThinkingLevel::Off;
	}
}
HFONT GetLinkFont()
{
	static HFONT s_linkFont = nullptr;
	if (s_linkFont != nullptr) {
		return s_linkFont;
	}

	LOGFONTA lf = {};
	HFONT base = GetDialogFont();
	if (base != nullptr && GetObjectA(base, sizeof(lf), &lf) == sizeof(lf)) {
		lf.lfUnderline = TRUE;
		s_linkFont = CreateFontIndirectA(&lf);
	}
	if (s_linkFont == nullptr) {
		s_linkFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
	}
	return s_linkFont;
}

AISettings ReadAISettingsFromNativeDialog(AIConfigDialogContext* ctx)
{
	AISettings next = {};
	if (ctx == nullptr || ctx->settings == nullptr) {
		return next;
	}

	next = *ctx->settings;
	next.protocolType = GetSelectedProtocol(ctx->hProtocol);
	next.thinkingLevel = GetSelectedThinkingLevel(ctx->hThinkingLevel);
	next.baseUrl = GetEditTextA(ctx->hBaseUrl);
	next.apiKey = GetEditTextA(ctx->hApiKey);
	next.model = GetEditTextA(ctx->hModel);
	next.tavilyApiKey = GetEditTextA(ctx->hTavilyApiKey);
	next.extraSystemPrompt = GetEditTextA(ctx->hExtraPrompt);
	next.customHeadersText = GetEditTextA(ctx->hCustomHeaders);
	return next;
}

void ApplyAISettingsToNativeDialog(AIConfigDialogContext* ctx, const AISettings& settings)
{
	if (ctx == nullptr) {
		return;
	}
	PopulateProtocolCombo(ctx->hProtocol, settings.protocolType);
	PopulateThinkingLevelCombo(ctx->hThinkingLevel, settings.thinkingLevel);
	SetEditTextA(ctx->hBaseUrl, settings.baseUrl);
	SetEditTextA(ctx->hApiKey, settings.apiKey);
	SetEditTextA(ctx->hModel, settings.model);
	SetEditTextA(ctx->hTavilyApiKey, settings.tavilyApiKey);
	SetEditTextA(ctx->hExtraPrompt, NormalizeMultilineForEdit(settings.extraSystemPrompt));
	SetEditTextA(ctx->hCustomHeaders, NormalizeMultilineForEdit(settings.customHeadersText));
}

void PopulateProfileCombo(HWND hCombo, const std::vector<AIConfigProfileEntry>& profiles, const std::string& activeProfileId)
{
	if (hCombo == nullptr) {
		return;
	}
	SendMessageA(hCombo, CB_RESETCONTENT, 0, 0);
	int selectedIndex = 0;
	for (size_t i = 0; i < profiles.size(); ++i) {
		const int index = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(profiles[i].name.c_str())));
		SendMessageA(hCombo, CB_SETITEMDATA, index, static_cast<LPARAM>(i));
		if (profiles[i].id == activeProfileId) {
			selectedIndex = index;
		}
	}
	SendMessageA(hCombo, CB_SETCURSEL, selectedIndex, 0);
}

struct AIConfigPresetSite {
	const wchar_t* label;
	const char* baseUrl;
	const char* const* models;
	size_t modelCount;
	AIProtocolType protocol;
};

constexpr const char* kRightPresetModels[] = { "gpt-5.5", "gpt-5.4", "gpt-5.4-mini" };
constexpr const char* kDeepseekPresetModels[] = { "deepseek-v4-flash", "deepseek-v4-pro" };
constexpr const char* kZhipuPresetModels[] = { "glm-5.2", "glm-5-turbo", "glm-4.7", "glm-4.5-air" };
constexpr const char* kQwenPresetModels[] = { "qwen3.7-plus", "qwen3.7-max", "qwen3.6-flash", "qwen3-coder-next", "qwen3-coder-plus" };
constexpr const char* kKimiPresetModels[] = { "kimi-k2.7-code", "kimi-k2.7-code-highspeed", "kimi-k2.6", "kimi-k2.5" };
constexpr const char* kDoubaoPresetModels[] = { "doubao-seed-2.0-pro", "doubao-seed-2.0-code", "doubao-seed-2.0-lite", "doubao-seed-1.8" };
constexpr const char* kMiniMaxPresetModels[] = { "MiniMax-M3", "MiniMax-M2.7", "MiniMax-M2.7-highspeed", "MiniMax-M2.5" };
constexpr const char* kAihubmixPresetModels[] = { "gpt-5.5", "claude-opus-4-8", "claude-sonnet-4-6", "deepseek-v4-pro", "deepseek-v4-flash", "gemini-3.1-pro-preview" };
constexpr const char* kSiliconFlowPresetModels[] = { "deepseek-ai/DeepSeek-V4-Flash", "deepseek-ai/DeepSeek-V4-Pro", "Pro/zai-org/GLM-5", "zai-org/GLM-5.1", "Qwen/Qwen3.5-397B-A17B" };
constexpr const char* kOpenAIPresetModels[] = { "gpt-5.5", "gpt-5.4", "gpt-5.4-mini", "gpt-5.3-codex" };
constexpr const char* kClaudePresetModels[] = { "claude-opus-4-8", "claude-sonnet-4-6", "claude-haiku-4-5" };
constexpr const char* kGeminiPresetModels[] = { "gemini-3.1-pro-preview", "gemini-3.1-pro-preview-customtools", "gemini-3.5-flash", "gemini-3.1-flash-lite", "gemini-2.5-pro" };

#define AI_PRESET_MODELS(name) name, std::size(name)

constexpr AIConfigPresetSite kAIConfigPresetSites[] = {
	{ L"Right",            "https://right.codes/codex",                         AI_PRESET_MODELS(kRightPresetModels),        AIProtocolType::OpenAI },
	{ L"Deepseek",         "https://api.deepseek.com",                          AI_PRESET_MODELS(kDeepseekPresetModels),     AIProtocolType::OpenAI },
	{ L"\u667A\u8C31",     "https://open.bigmodel.cn/api/paas/v4",              AI_PRESET_MODELS(kZhipuPresetModels),        AIProtocolType::OpenAI },
	{ L"\u5343\u95EE",     "https://dashscope.aliyuncs.com/compatible-mode/v1", AI_PRESET_MODELS(kQwenPresetModels),         AIProtocolType::OpenAI },
	{ L"Kimi",             "https://api.moonshot.cn/v1",                        AI_PRESET_MODELS(kKimiPresetModels),         AIProtocolType::OpenAI },
	{ L"\u8C46\u5305",     "https://ark.cn-beijing.volces.com/api/v3",          AI_PRESET_MODELS(kDoubaoPresetModels),       AIProtocolType::OpenAI },
	{ L"MiniMax",          "https://api.minimax.chat/v1",                       AI_PRESET_MODELS(kMiniMaxPresetModels),      AIProtocolType::OpenAI },
	{ L"aihubmix",         "https://aihubmix.com/v1",                           AI_PRESET_MODELS(kAihubmixPresetModels),     AIProtocolType::OpenAI },
	{ L"\u7845\u57FA\u6D41\u52A8", "https://api.siliconflow.cn/v1",             AI_PRESET_MODELS(kSiliconFlowPresetModels),  AIProtocolType::OpenAI },
	{ L"OpenAI",           "https://api.openai.com/v1",                         AI_PRESET_MODELS(kOpenAIPresetModels),       AIProtocolType::OpenAIResponses },
	{ L"Claude",           "https://api.anthropic.com",                         AI_PRESET_MODELS(kClaudePresetModels),       AIProtocolType::Claude },
	{ L"Gemini",           "https://generativelanguage.googleapis.com",         AI_PRESET_MODELS(kGeminiPresetModels),       AIProtocolType::Gemini },
};

#undef AI_PRESET_MODELS

const char* GetPresetSitePrimaryModel(const AIConfigPresetSite& site)
{
	return site.modelCount > 0 && site.models != nullptr ? site.models[0] : "";
}

std::string BuildPresetProfileNameLocal(const AIConfigPresetSite& site, const char* model)
{
	const std::string label = Utf8ToLocalText(WideToUtf8(site.label == nullptr ? L"" : site.label));
	return label + "(" + (model == nullptr ? "" : model) + ")";
}

void AddNativePresetProfile(HWND hWnd, AIConfigDialogContext* ctx, const AIConfigPresetSite& site, const char* model)
{
	if (ctx == nullptr || ctx->settings == nullptr) {
		return;
	}
	const int oldIndex = FindProfileIndexById(ctx->profiles, ctx->activeProfileId);
	if (oldIndex >= 0) {
		ctx->profiles[static_cast<size_t>(oldIndex)].settings = ReadAISettingsFromNativeDialog(ctx);
	}

	AIConfigProfileEntry entry;
	entry.id = std::format("profile_{}", GetTickCount64());
	entry.name = BuildPresetProfileNameLocal(site, model);
	entry.settings = {};
	entry.settings.protocolType = site.protocol;
	entry.settings.thinkingLevel = AIThinkingLevel::Off;
	entry.settings.baseUrl = site.baseUrl == nullptr ? std::string() : site.baseUrl;
	entry.settings.model = model == nullptr ? std::string() : model;

	ctx->profiles.push_back(std::move(entry));
	ctx->activeProfileId = ctx->profiles.back().id;
	*ctx->settings = ctx->profiles.back().settings;
	PopulateProfileCombo(ctx->hProfileCombo, ctx->profiles, ctx->activeProfileId);
	ApplyAISettingsToNativeDialog(ctx, *ctx->settings);
	SetFocus(ctx->hApiKey);
}

void ShowNativePresetProfileMenu(HWND hWnd, AIConfigDialogContext* ctx, HWND hButton)
{
	if (hWnd == nullptr || ctx == nullptr || hButton == nullptr) {
		return;
	}
	HMENU menu = CreatePopupMenu();
	if (menu == nullptr) {
		return;
	}
	std::vector<std::pair<size_t, size_t>> menuItems;
	for (size_t i = 0; i < std::size(kAIConfigPresetSites); ++i) {
		const AIConfigPresetSite& site = kAIConfigPresetSites[i];
		for (size_t modelIndex = 0; modelIndex < site.modelCount; ++modelIndex) {
			const char* model = site.models == nullptr ? nullptr : site.models[modelIndex];
			if (model == nullptr || model[0] == '\0') {
				continue;
			}
			const UINT commandId = kAIConfigNativePresetMenuBase + static_cast<UINT>(menuItems.size());
			const std::wstring text = std::wstring(site.label) + L"(" + Utf8ToWide(model) + L")";
			AppendMenuW(menu, MF_STRING, commandId, text.c_str());
			menuItems.emplace_back(i, modelIndex);
		}
	}
	if (menuItems.empty()) {
		DestroyMenu(menu);
		return;
	}

	RECT rc = {};
	GetWindowRect(hButton, &rc);
	const int command = TrackPopupMenu(
		menu,
		TPM_RETURNCMD | TPM_LEFTALIGN | TPM_TOPALIGN,
		rc.left,
		rc.bottom,
		0,
		hWnd,
		nullptr);
	DestroyMenu(menu);

	if (command >= static_cast<int>(kAIConfigNativePresetMenuBase)) {
		const size_t index = static_cast<size_t>(command - kAIConfigNativePresetMenuBase);
		if (index < menuItems.size()) {
			const auto [siteIndex, modelIndex] = menuItems[index];
			const AIConfigPresetSite& site = kAIConfigPresetSites[siteIndex];
			const char* model = site.models == nullptr || modelIndex >= site.modelCount ? GetPresetSitePrimaryModel(site) : site.models[modelIndex];
			AddNativePresetProfile(hWnd, ctx, site, model);
		}
	}
}

bool ValidateAISettingsForConnection(HWND hWnd, const AISettings& settings)
{
	if (AIService::Trim(settings.baseUrl).empty() ||
		AIService::Trim(settings.apiKey).empty() ||
		AIService::Trim(settings.model).empty()) {
		MessageBoxA(hWnd, "baseUrl / apiKey / model cannot be empty.", "AI Config", MB_ICONWARNING | MB_OK);
		return false;
	}
	std::string headerError;
	if (!AIService::ValidateCustomHeadersText(settings.customHeadersText, headerError)) {
		MessageBoxA(hWnd, headerError.c_str(), "AI Config", MB_ICONWARNING | MB_OK);
		return false;
	}
	return true;
}

std::string BuildAIConnectionTestMessage(const AIResult& result)
{
	if (result.ok) {
		std::string message = "连通性测试成功。";
		if (result.httpStatus > 0) {
			message += "\nHTTP: " + std::to_string(result.httpStatus);
		}
		const std::string trimmed = AIService::Trim(result.content);
		if (!trimmed.empty()) {
			message += "\n模型返回：";
			message += trimmed;
		}
		return message;
	}

	std::string message = "连通性测试失败。";
	if (result.httpStatus > 0) {
		message += "\nHTTP: " + std::to_string(result.httpStatus);
	}
	if (!result.error.empty()) {
		message += "\n错误：";
		message += result.error;
	}
	return message;
}

void SetAIConfigNativeTestBusy(AIConfigDialogContext* ctx, bool busy)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->testInFlight = busy;
	if (ctx->hTestConnection != nullptr) {
		SetWindowTextA(ctx->hTestConnection, busy ? "测试中..." : "测试连通性");
		EnableWindow(ctx->hTestConnection, busy ? FALSE : TRUE);
	}
}

DWORD WINAPI AIConfigConnectionTestWorkerProc(LPVOID param)
{
	AIConfigConnectionTestRequest* request = reinterpret_cast<AIConfigConnectionTestRequest*>(param);
	if (request == nullptr) {
		return 0;
	}

	AIConfigConnectionTestResult* result = new (std::nothrow) AIConfigConnectionTestResult();
	if (result == nullptr) {
		delete request;
		return 0;
	}

	result->forWebView = request->forWebView;
	try {
		result->result = AIService::TestConnection(request->settings);
	}
	catch (const std::exception& ex) {
		result->result.ok = false;
		result->result.error = std::string("connection test exception: ") + ex.what();
	}
	catch (...) {
		result->result.ok = false;
		result->result.error = "connection test unknown exception";
	}

	const HWND hWnd = request->dialogHwnd;
	delete request;
	if (hWnd == nullptr || !PostMessageA(hWnd, WM_AUTOLINKER_AI_CONFIG_TEST_DONE, 0, reinterpret_cast<LPARAM>(result))) {
		delete result;
	}
	return 0;
}

bool StartAIConfigConnectionTest(HWND hWnd, const AISettings& settings, bool forWebView)
{
	AIConfigConnectionTestRequest* request = new (std::nothrow) AIConfigConnectionTestRequest();
	if (request == nullptr) {
		MessageBoxA(hWnd, "无法启动连通性测试线程。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}

	request->dialogHwnd = hWnd;
	request->settings = settings;
	request->forWebView = forWebView;

	const HANDLE workerHandle = CreateThread(nullptr, 0, AIConfigConnectionTestWorkerProc, request, 0, nullptr);
	if (workerHandle == nullptr) {
		delete request;
		MessageBoxA(hWnd, "无法启动连通性测试线程。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}

	CloseHandle(workerHandle);
	return true;
}

std::string TrimTrailingSlashesCopy(std::string text)
{
	text = AIService::Trim(text);
	while (!text.empty() && text.back() == '/') {
		text.pop_back();
	}
	return text;
}

bool EndsWithAsciiInsensitiveLocal(const std::string& text, const std::string& suffix)
{
	if (suffix.size() > text.size()) {
		return false;
	}
	return _stricmp(text.c_str() + text.size() - suffix.size(), suffix.c_str()) == 0;
}

std::string BuildModelListUrl(const AISettings& settings)
{
	std::string url = TrimTrailingSlashesCopy(settings.baseUrl);
	if (url.empty()) {
		return std::string();
	}
	if (settings.protocolType == AIProtocolType::Gemini) {
		if (EndsWithAsciiInsensitiveLocal(url, "/models")) {
			return url;
		}
		if (EndsWithAsciiInsensitiveLocal(url, "/v1beta") || EndsWithAsciiInsensitiveLocal(url, "/v1")) {
			return url + "/models";
		}
		return url + "/v1beta/models";
	}
	if (EndsWithAsciiInsensitiveLocal(url, "/models")) {
		return url;
	}
	if (EndsWithAsciiInsensitiveLocal(url, "/chat/completions")) {
		url = url.substr(0, url.size() - std::string("/chat/completions").size());
	}
	else if (EndsWithAsciiInsensitiveLocal(url, "/responses")) {
		url = url.substr(0, url.size() - std::string("/responses").size());
	}
	else if (EndsWithAsciiInsensitiveLocal(url, "/messages")) {
		url = url.substr(0, url.size() - std::string("/messages").size());
	}
	if (EndsWithAsciiInsensitiveLocal(url, "/v1")) {
		return url + "/models";
	}
	return url + "/v1/models";
}

struct AIConfigHeaderEntry {
	std::string name;
	std::string value;
};

void UpsertAIConfigHeader(
	std::vector<AIConfigHeaderEntry>& headers,
	const std::string& name,
	const std::string& value)
{
	for (auto& header : headers) {
		if (_stricmp(header.name.c_str(), name.c_str()) == 0) {
			header.name = name;
			header.value = value;
			return;
		}
	}
	headers.push_back({ name, value });
}

void MergeAIConfigCustomHeaders(
	std::vector<AIConfigHeaderEntry>& headers,
	const std::string& customText)
{
	size_t lineStart = 0;
	while (lineStart <= customText.size()) {
		const size_t lineEnd = customText.find_first_of("\r\n", lineStart);
		const std::string line = lineEnd == std::string::npos
			? customText.substr(lineStart)
			: customText.substr(lineStart, lineEnd - lineStart);
		if (lineEnd == std::string::npos) {
			lineStart = customText.size() + 1;
		}
		else if (customText[lineEnd] == '\r' && lineEnd + 1 < customText.size() && customText[lineEnd + 1] == '\n') {
			lineStart = lineEnd + 2;
		}
		else {
			lineStart = lineEnd + 1;
		}

		const std::string trimmed = AIService::Trim(line);
		if (trimmed.empty()) {
			continue;
		}
		const size_t colon = trimmed.find(':');
		if (colon == std::string::npos || colon == 0) {
			continue;
		}
		UpsertAIConfigHeader(
			headers,
			AIService::Trim(trimmed.substr(0, colon)),
			AIService::Trim(trimmed.substr(colon + 1)));
	}
}

std::string BuildModelListHeaders(const AISettings& settings)
{
	std::vector<AIConfigHeaderEntry> headerEntries;
	if (settings.protocolType == AIProtocolType::Claude) {
		headerEntries.push_back({ "x-api-key", settings.apiKey });
		headerEntries.push_back({ "anthropic-version", "2023-06-01" });
		headerEntries.push_back({ "Accept", "application/json" });
	}
	else if (settings.protocolType == AIProtocolType::Gemini) {
		headerEntries.push_back({ "x-goog-api-key", settings.apiKey });
		headerEntries.push_back({ "Accept", "application/json" });
	}
	else {
		headerEntries.push_back({ "Authorization", "Bearer " + settings.apiKey });
		headerEntries.push_back({ "Accept", "application/json" });
	}

	MergeAIConfigCustomHeaders(headerEntries, settings.customHeadersText);

	std::string headers;
	for (const auto& entry : headerEntries) {
		headers += entry.name;
		headers += ": ";
		headers += entry.value;
		headers += "\r\n";
	}
	return headers;
}

void AddModelNameFromJsonValue(
	const nlohmann::json& value,
	std::set<std::string>& outModels,
	bool stripModelPrefix)
{
	if (!value.is_string()) {
		return;
	}
	std::string model = AIService::Trim(value.get<std::string>());
	if (stripModelPrefix && model.rfind("models/", 0) == 0) {
		model = model.substr(std::string("models/").size());
	}
	if (!model.empty()) {
		outModels.insert(model);
	}
}

bool SupportsGeminiGenerateContent(const nlohmann::json& item)
{
	if (!item.is_object() || !item.contains("supportedGenerationMethods") || !item["supportedGenerationMethods"].is_array()) {
		return true;
	}
	for (const auto& method : item["supportedGenerationMethods"]) {
		if (method.is_string() && method.get<std::string>() == "generateContent") {
			return true;
		}
	}
	return false;
}

void CollectModelNamesFromArray(
	const nlohmann::json& arrayValue,
	std::set<std::string>& outModels,
	bool stripModelPrefix,
	bool requireGeminiGenerateContent)
{
	if (!arrayValue.is_array()) {
		return;
	}
	for (const auto& item : arrayValue) {
		if (item.is_string()) {
			AddModelNameFromJsonValue(item, outModels, stripModelPrefix);
			continue;
		}
		if (!item.is_object()) {
			continue;
		}
		if (requireGeminiGenerateContent && !SupportsGeminiGenerateContent(item)) {
			continue;
		}
		if (item.contains("id")) {
			AddModelNameFromJsonValue(item["id"], outModels, stripModelPrefix);
		}
		else if (item.contains("name")) {
			AddModelNameFromJsonValue(item["name"], outModels, stripModelPrefix);
		}
	}
}

bool TryParseModelListResponse(
	const std::string& responseBody,
	AIProtocolType protocolType,
	std::vector<std::string>& outModels,
	std::string& outError)
{
	outModels.clear();
	outError.clear();
	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		std::set<std::string> models;
		const bool gemini = protocolType == AIProtocolType::Gemini;
		if (parsed.is_array()) {
			CollectModelNamesFromArray(parsed, models, gemini, gemini);
		}
		else if (parsed.is_object()) {
			if (parsed.contains("data")) {
				CollectModelNamesFromArray(parsed["data"], models, gemini, gemini);
			}
			if (parsed.contains("models")) {
				CollectModelNamesFromArray(parsed["models"], models, gemini, gemini);
			}
		}
		outModels.assign(models.begin(), models.end());
		if (outModels.empty()) {
			outError = "模型列表为空。";
			return false;
		}
		return true;
	}
	catch (const std::exception& ex) {
		outError = std::string("解析模型列表失败：") + ex.what();
		return false;
	}
}

DWORD WINAPI AIConfigModelListWorkerProc(LPVOID param)
{
	AIConfigModelListRequest* request = reinterpret_cast<AIConfigModelListRequest*>(param);
	if (request == nullptr) {
		return 0;
	}

	AIConfigModelListResult* result = new (std::nothrow) AIConfigModelListResult();
	if (result == nullptr) {
		delete request;
		return 0;
	}

	const HWND hWnd = request->dialogHwnd;
	try {
		std::string headerError;
		if (AIService::Trim(request->settings.baseUrl).empty()) {
			result->error = "请输入 Base URL 后再获取模型列表。";
		}
		else if (AIService::Trim(request->settings.apiKey).empty()) {
			result->error = "请输入 API Key 后再获取模型列表。";
		}
		else if (!AIService::ValidateCustomHeadersText(request->settings.customHeadersText, headerError)) {
			result->error = headerError;
		}
		else {
			const std::string url = BuildModelListUrl(request->settings);
			const std::string headers = BuildModelListHeaders(request->settings);
			const int timeout = (std::clamp)(request->settings.timeoutMs <= 0 ? 30000 : request->settings.timeoutMs, 1000, 300000);
			const auto [body, statusCode] = PerformGetRequest(url, headers, timeout, false, false);
			result->httpStatus = statusCode;
			if (statusCode < 200 || statusCode >= 300) {
				result->error = std::format("HTTP {}: {}", statusCode, body.substr(0, (std::min)(body.size(), size_t(500))));
			}
			else {
				result->ok = TryParseModelListResponse(body, request->settings.protocolType, result->models, result->error);
			}
		}
	}
	catch (const std::exception& ex) {
		result->error = std::string("获取模型列表异常：") + ex.what();
	}
	catch (...) {
		result->error = "获取模型列表异常：unknown";
	}

	delete request;
	if (hWnd == nullptr || !PostMessageA(hWnd, WM_AUTOLINKER_AI_CONFIG_MODEL_LIST_DONE, 0, reinterpret_cast<LPARAM>(result))) {
		delete result;
	}
	return 0;
}

bool StartAIConfigModelListFetch(HWND hWnd, const AISettings& settings)
{
	AIConfigModelListRequest* request = new (std::nothrow) AIConfigModelListRequest();
	if (request == nullptr) {
		MessageBoxA(hWnd, "无法启动模型列表获取线程。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}
	request->dialogHwnd = hWnd;
	request->settings = settings;

	const HANDLE workerHandle = CreateThread(nullptr, 0, AIConfigModelListWorkerProc, request, 0, nullptr);
	if (workerHandle == nullptr) {
		delete request;
		MessageBoxA(hWnd, "无法启动模型列表获取线程。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}

	CloseHandle(workerHandle);
	return true;
}

std::string BuildAIConfigWebViewSettingsJson(const std::vector<AIConfigProfileEntry>& profiles, const std::string& activeProfileId)
{
	nlohmann::json initialSettings;
	initialSettings["activeProfileId"] = LocalToUtf8Text(activeProfileId);
	initialSettings["profiles"] = nlohmann::json::array();
	for (const auto& profile : profiles) {
		nlohmann::json item;
		item["id"] = LocalToUtf8Text(profile.id);
		item["name"] = LocalToUtf8Text(profile.name);
		item["protocolType"] = AIService::ProtocolTypeToString(profile.settings.protocolType);
		item["thinkingLevel"] = AIService::ThinkingLevelToString(profile.settings.thinkingLevel);
		item["baseUrl"] = LocalToUtf8Text(profile.settings.baseUrl);
		item["apiKey"] = LocalToUtf8Text(profile.settings.apiKey);
		item["model"] = LocalToUtf8Text(profile.settings.model);
		item["extraPrompt"] = LocalToUtf8Text(profile.settings.extraSystemPrompt);
		item["customHeaders"] = LocalToUtf8Text(profile.settings.customHeadersText);
		item["tavilyApiKey"] = LocalToUtf8Text(profile.settings.tavilyApiKey);
		item["contextWindow"] = profile.settings.contextWindowTokens;
		initialSettings["profiles"].push_back(std::move(item));
	}
	return initialSettings.dump();
}

std::string BuildAIConfigWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_CONFIG_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>Config shell resource missing.</body></html>";
}

void LayoutAIConfigWebViewDialog(HWND hWnd, AIConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom));
	const int loadingLeft = 16;
	const int loadingTop = 14;
	const int loadingWidth = static_cast<int>((std::max)(0L, rc.right - loadingLeft * 2L));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, 0, 0, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, loadingLeft, loadingTop, loadingWidth, 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = {};
		bounds.left = 0;
		bounds.top = 0;
		bounds.right = static_cast<LONG>(hostWidth);
		bounds.bottom = static_cast<LONG>(hostHeight);
		ctx->webViewController->put_Bounds(bounds);
	}
}

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

		HWND hProfileLabel = CreateWindowA("STATIC", "Profile:", WS_CHILD | WS_VISIBLE,
			16, 16, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hProfileCombo = CreateWindowExA(0, "COMBOBOX", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
			120, 14, 230, 260, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROFILE_COMBO), nullptr, nullptr);
		HWND hAddProfile = CreateWindowW(L"BUTTON", L"\u65B0\u5EFA", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			360, 14, 70, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROFILE_ADD), nullptr, nullptr);
		HWND hAddPresetProfile = CreateWindowW(L"BUTTON", L"\u4F7F\u7528\u9884\u8BBE\u7AD9\u70B9\u65B0\u5EFA", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			438, 14, 150, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROFILE_ADD_PRESET), nullptr, nullptr);
		HWND hRenameProfile = CreateWindowW(L"BUTTON", L"\u91CD\u547D\u540D", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			596, 14, 70, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROFILE_RENAME), nullptr, nullptr);
		HWND hDeleteProfile = CreateWindowW(L"BUTTON", L"\u5220\u9664", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			674, 14, 70, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROFILE_DELETE), nullptr, nullptr);
		PopulateProfileCombo(ctx->hProfileCombo, ctx->profiles, ctx->activeProfileId);

		HWND hProtocolLabel = CreateWindowA("STATIC", "Protocol:", WS_CHILD | WS_VISIBLE,
			16, 50, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hProtocol = CreateWindowExA(0, "COMBOBOX", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
			120, 48, 180, 220, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROTOCOL), nullptr, nullptr);
		PopulateProtocolCombo(ctx->hProtocol, ctx->settings->protocolType);

		HWND hBaseUrlLabel = CreateWindowA("STATIC", "Base URL:", WS_CHILD | WS_VISIBLE,
			16, 84, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hBaseUrl = CreateWindowExA(0, "EDIT", ctx->settings->baseUrl.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 82, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_BASE_URL), nullptr, nullptr);

		HWND hApiKeyLabel = CreateWindowA("STATIC", "API Key:", WS_CHILD | WS_VISIBLE,
			16, 118, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hApiKey = CreateWindowExA(0, "EDIT", ctx->settings->apiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 116, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_API_KEY), nullptr, nullptr);

		HWND hModelLabel = CreateWindowA("STATIC", "Model:", WS_CHILD | WS_VISIBLE,
			16, 152, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hModel = CreateWindowExA(0, "EDIT", ctx->settings->model.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 150, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_MODEL), nullptr, nullptr);

		HWND hThinkingLabel = CreateWindowA("STATIC", "Thinking:", WS_CHILD | WS_VISIBLE,
			16, 186, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hThinkingLevel = CreateWindowExA(0, "COMBOBOX", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
			120, 184, 180, 220, hWnd, reinterpret_cast<HMENU>(IDC_CFG_THINKING_LEVEL), nullptr, nullptr);
		PopulateThinkingLevelCombo(ctx->hThinkingLevel, ctx->settings->thinkingLevel);

		const std::wstring getKeyLinkText =
			L"<a href=\"https://right.codes/register?aff=3dc87885\">从转发平台获取Key</a>";
		HWND hGetKeyLink = CreateWindowExW(0, L"SysLink",
			getKeyLinkText.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			330, 184, 240, 22, hWnd, reinterpret_cast<HMENU>(IDC_CFG_GET_KEY_LINK), nullptr, nullptr);
		if (hGetKeyLink != nullptr) {
			ctx->useNativeLink = true;
			ctx->hGetKeyLink = hGetKeyLink;
		}
		else {
			ctx->useNativeLink = false;
			ctx->hGetKeyLink = CreateWindowW(
				L"STATIC",
				L"\u4ECE\u8F6C\u53D1\u5E73\u53F0\u83B7\u53D6Key",
				WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY,
				330, 184, 240, 22, hWnd, reinterpret_cast<HMENU>(IDC_CFG_GET_KEY_LINK), nullptr, nullptr);
			hGetKeyLink = ctx->hGetKeyLink;
		}

		HWND hTavilyApiKeyLabel = CreateWindowA("STATIC", "Tavily API Key:", WS_CHILD | WS_VISIBLE,
			16, 248, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hTavilyApiKey = CreateWindowExA(0, "EDIT", ctx->settings->tavilyApiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 246, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_TAVILY_API_KEY), nullptr, nullptr);

		HWND hExtraPromptLabel = CreateWindowA("STATIC", "System Prompt:", WS_CHILD | WS_VISIBLE,
			16, 282, 140, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hExtraPrompt = CreateWindowExA(0, "EDIT", NormalizeMultilineForEdit(ctx->settings->extraSystemPrompt).c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			120, 282, 500, 96, hWnd, reinterpret_cast<HMENU>(IDC_CFG_EXTRA_PROMPT), nullptr, nullptr);

		HWND hCustomHeadersLabel = CreateWindowA("STATIC", "Custom Headers:", WS_CHILD | WS_VISIBLE,
			16, 388, 140, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hCustomHeaders = CreateWindowExA(0, "EDIT", NormalizeMultilineForEdit(ctx->settings->customHeadersText).c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			120, 388, 500, 100, hWnd, reinterpret_cast<HMENU>(IDC_CFG_CUSTOM_HEADERS), nullptr, nullptr);

		ctx->hTestConnection = CreateWindowW(L"BUTTON", L"\u6D4B\u8BD5\u8FDE\u901A\u6027", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			300, 486, 110, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_TEST_CONNECTION), nullptr, nullptr);
		HWND hSave = CreateWindowW(L"BUTTON", L"\u4FDD\u5B58\u5E76\u7EE7\u7EED", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			420, 486, 100, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_SAVE), nullptr, nullptr);
		HWND hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			530, 486, 90, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_CANCEL), nullptr, nullptr);

		std::array<HWND, 26> controls = {
			hProfileLabel,
			ctx->hProfileCombo,
			hAddProfile,
			hAddPresetProfile,
			hRenameProfile,
			hDeleteProfile,
			hProtocolLabel,
			ctx->hProtocol,
			hBaseUrlLabel,
			hApiKeyLabel,
			hModelLabel,
			hThinkingLabel,
			hTavilyApiKeyLabel,
			ctx->hGetKeyLink,
			hExtraPromptLabel,
			hCustomHeadersLabel,
			ctx->hBaseUrl,
			ctx->hApiKey,
			ctx->hModel,
			ctx->hThinkingLevel,
			ctx->hTavilyApiKey,
			ctx->hExtraPrompt,
			ctx->hCustomHeaders,
			ctx->hTestConnection,
			hSave,
			hCancel
		};
		for (HWND hControl : controls) {
			SetDefaultFont(hControl);
		}
		if (ctx->hGetKeyLink != nullptr && !ctx->useNativeLink) {
			SendMessageA(ctx->hGetKeyLink, WM_SETFONT, reinterpret_cast<WPARAM>(GetLinkFont()), TRUE);
		}
		SetFocus(ctx->hBaseUrl);
		return 0;
	}

	case WM_COMMAND: {
		const int id = LOWORD(wParam);
		if (ctx == nullptr || ctx->settings == nullptr) {
			return 0;
		}

		if (id == IDC_CFG_PROFILE_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
			const int oldIndex = FindProfileIndexById(ctx->profiles, ctx->activeProfileId);
			if (oldIndex >= 0) {
				ctx->profiles[oldIndex].settings = ReadAISettingsFromNativeDialog(ctx);
			}
			const int selected = static_cast<int>(SendMessageA(ctx->hProfileCombo, CB_GETCURSEL, 0, 0));
			if (selected != CB_ERR) {
				const LRESULT itemData = SendMessageA(ctx->hProfileCombo, CB_GETITEMDATA, selected, 0);
				if (itemData >= 0 && itemData < static_cast<LRESULT>(ctx->profiles.size())) {
					ctx->activeProfileId = ctx->profiles[static_cast<size_t>(itemData)].id;
					*ctx->settings = ctx->profiles[static_cast<size_t>(itemData)].settings;
					ApplyAISettingsToNativeDialog(ctx, *ctx->settings);
				}
			}
			return 0;
		}

		if (id == IDC_CFG_PROFILE_ADD) {
			std::string name = "新配置";
			if (!ShowAITextInputDialog(hWnd, "新建AI配置组", "请输入新的配置组名称。", name)) {
				return 0;
			}
			if (AIService::Trim(name).empty()) {
				MessageBoxA(hWnd, "配置组名称不能为空。", "AI Config", MB_ICONWARNING | MB_OK);
				return 0;
			}
			const int oldIndex = FindProfileIndexById(ctx->profiles, ctx->activeProfileId);
			if (oldIndex >= 0) {
				ctx->profiles[oldIndex].settings = ReadAISettingsFromNativeDialog(ctx);
			}
			AIConfigProfileEntry entry;
			entry.id = std::format("profile_{}", GetTickCount64());
			entry.name = name;
			entry.settings = {};
			ctx->profiles.push_back(std::move(entry));
			ctx->activeProfileId = ctx->profiles.back().id;
			*ctx->settings = ctx->profiles.back().settings;
			PopulateProfileCombo(ctx->hProfileCombo, ctx->profiles, ctx->activeProfileId);
			ApplyAISettingsToNativeDialog(ctx, *ctx->settings);
			return 0;
		}

		if (id == IDC_CFG_PROFILE_ADD_PRESET) {
			ShowNativePresetProfileMenu(hWnd, ctx, reinterpret_cast<HWND>(lParam));
			return 0;
		}

		if (id == IDC_CFG_PROFILE_RENAME) {
			const int index = FindProfileIndexById(ctx->profiles, ctx->activeProfileId);
			if (index < 0) {
				return 0;
			}
			std::string name = ctx->profiles[static_cast<size_t>(index)].name;
			if (!ShowAITextInputDialog(hWnd, "重命名AI配置组", "请输入新的配置组名称。", name)) {
				return 0;
			}
			if (AIService::Trim(name).empty()) {
				MessageBoxA(hWnd, "配置组名称不能为空。", "AI Config", MB_ICONWARNING | MB_OK);
				return 0;
			}
			ctx->profiles[static_cast<size_t>(index)].name = name;
			PopulateProfileCombo(ctx->hProfileCombo, ctx->profiles, ctx->activeProfileId);
			return 0;
		}

		if (id == IDC_CFG_PROFILE_DELETE) {
			if (ctx->profiles.size() <= 1) {
				MessageBoxA(hWnd, "至少保留一个配置组。", "AI Config", MB_ICONWARNING | MB_OK);
				return 0;
			}
			const int index = FindProfileIndexById(ctx->profiles, ctx->activeProfileId);
			if (index < 0) {
				return 0;
			}
			const std::string confirmText = "确定删除配置组：“" + ctx->profiles[static_cast<size_t>(index)].name + "”？";
			if (MessageBoxA(hWnd, confirmText.c_str(), "AI Config", MB_ICONQUESTION | MB_YESNO) != IDYES) {
				return 0;
			}
			ctx->profiles.erase(ctx->profiles.begin() + index);
			ctx->activeProfileId = ctx->profiles.front().id;
			*ctx->settings = ctx->profiles.front().settings;
			PopulateProfileCombo(ctx->hProfileCombo, ctx->profiles, ctx->activeProfileId);
			ApplyAISettingsToNativeDialog(ctx, *ctx->settings);
			return 0;
		}

		if (id == IDC_CFG_SAVE) {
			AISettings next = ReadAISettingsFromNativeDialog(ctx);
			if (!ValidateAISettingsForConnection(hWnd, next)) {
				return 0;
			}

			const int index = FindProfileIndexById(ctx->profiles, ctx->activeProfileId);
			if (index >= 0) {
				ctx->profiles[static_cast<size_t>(index)].settings = next;
			}
			if (ctx->jsonConfig == nullptr ||
				!ctx->jsonConfig->replaceProfiles(BuildProfileSnapshotsForSave(ctx->profiles), ctx->activeProfileId)) {
				MessageBoxA(hWnd, "保存 AI 配置组失败。", "AI Config", MB_ICONERROR | MB_OK);
				return 0;
			}
			*ctx->settings = next;
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}

		if (id == IDC_CFG_TEST_CONNECTION) {
			if (ctx->testInFlight) {
				return 0;
			}

			const AISettings next = ReadAISettingsFromNativeDialog(ctx);
			if (!ValidateAISettingsForConnection(hWnd, next)) {
				return 0;
			}
			if (!StartAIConfigConnectionTest(hWnd, next, false)) {
				return 0;
			}
			SetAIConfigNativeTestBusy(ctx, true);
			return 0;
		}

		if (id == IDC_CFG_GET_KEY_LINK && HIWORD(wParam) == STN_CLICKED) {
			ShellExecuteA(hWnd, "open", "https://right.codes/register?aff=3dc87885", nullptr, nullptr, SW_SHOWNORMAL);
			return 0;
		}

		if (id == IDC_CFG_CANCEL) {
			ctx->accepted = false;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_NOTIFY: {
		const NMHDR* hdr = reinterpret_cast<const NMHDR*>(lParam);
		if (hdr != nullptr && ctx != nullptr && ctx->useNativeLink && hdr->idFrom == IDC_CFG_GET_KEY_LINK &&
			(hdr->code == NM_CLICK || hdr->code == NM_RETURN)) {
			ShellExecuteA(hWnd, "open", "https://right.codes/register?aff=3dc87885", nullptr, nullptr, SW_SHOWNORMAL);
			return 0;
		}
		break;
	}

	case WM_CTLCOLORSTATIC: {
		if (ctx != nullptr && !ctx->useNativeLink && reinterpret_cast<HWND>(lParam) == ctx->hGetKeyLink) {
			HDC hdc = reinterpret_cast<HDC>(wParam);
			SetTextColor(hdc, RGB(0, 102, 204));
			SetBkMode(hdc, TRANSPARENT);
			return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
		}
		break;
	}

	case WM_AUTOLINKER_AI_CONFIG_TEST_DONE: {
		AIConfigConnectionTestResult* result = reinterpret_cast<AIConfigConnectionTestResult*>(lParam);
		if (ctx == nullptr || result == nullptr) {
			delete result;
			return 0;
		}

		SetAIConfigNativeTestBusy(ctx, false);
		const std::string message = BuildAIConnectionTestMessage(result->result);
		MessageBoxA(hWnd, message.c_str(), "AI Config", (result->result.ok ? MB_ICONINFORMATION : MB_ICONERROR) | MB_OK);
		delete result;
		return 0;
	}

	case WM_AUTOLINKER_AI_CONFIG_MODEL_LIST_DONE: {
		AIConfigModelListResult* result = reinterpret_cast<AIConfigModelListResult*>(lParam);
		delete result;
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

AISettings ReadAISettingsFromWebProfilePayload(const AISettings& current, const nlohmann::json& data)
{
	AISettings next = current;
	next.protocolType = AIService::ParseProtocolType(data.value("protocolType", AIService::ProtocolTypeToString(next.protocolType)));
	next.thinkingLevel = AIService::ParseThinkingLevel(data.value("thinkingLevel", AIService::ThinkingLevelToString(next.thinkingLevel)));
	next.baseUrl = Utf8ToLocalText(data.value("baseUrl", ""));
	next.apiKey = Utf8ToLocalText(data.value("apiKey", ""));
	next.model = Utf8ToLocalText(data.value("model", ""));
	next.extraSystemPrompt = Utf8ToLocalText(data.value("extraPrompt", ""));
	next.customHeadersText = Utf8ToLocalText(data.value("customHeaders", ""));
	next.tavilyApiKey = Utf8ToLocalText(data.value("tavilyApiKey", ""));
	next.contextWindowTokens = (std::max)(0, data.value("contextWindow", 0));
	return next;
}

bool TryApplyAISettingsFromWebPayload(HWND hWnd, AIConfigWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr || ctx->settings == nullptr || ctx->jsonConfig == nullptr) {
		return false;
	}

	if (!data.contains("profiles") || !data["profiles"].is_array()) {
		return false;
	}

	const std::string activeProfileId = Utf8ToLocalText(data.value("activeProfileId", ""));
	std::vector<AIConfigProfileEntry> nextProfiles;
	nextProfiles.reserve(data["profiles"].size());
	for (const auto& item : data["profiles"]) {
		if (!item.is_object()) {
			return false;
		}
		AIConfigProfileEntry entry;
		entry.id = Utf8ToLocalText(item.value("id", ""));
		entry.name = Utf8ToLocalText(item.value("name", ""));
		entry.settings = ReadAISettingsFromWebProfilePayload({}, item);
		if (entry.id.empty() || entry.name.empty()) {
			return false;
		}
		if (!ValidateAISettingsForConnection(hWnd, entry.settings)) {
			return false;
		}
		nextProfiles.push_back(std::move(entry));
	}

	if (nextProfiles.empty()) {
		return false;
	}

	const std::vector<AIJsonConfigProfileSnapshot> snapshots = BuildProfileSnapshotsForSave(nextProfiles);
	if (!ctx->jsonConfig->replaceProfiles(snapshots, activeProfileId)) {
		MessageBoxA(hWnd, "保存 AI 配置组失败。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}

	ctx->profiles = nextProfiles;
	ctx->activeProfileId = activeProfileId;
	for (const auto& entry : ctx->profiles) {
		if (entry.id == ctx->activeProfileId) {
			*ctx->settings = entry.settings;
			break;
		}
	}
	ctx->accepted = true;
	DestroyWindow(hWnd);
	return true;
}

void ExecuteAIConfigWebViewScript(AIConfigWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void SetAIConfigWebViewTestBusy(AIConfigWebViewDialogContext* ctx, bool busy)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->testInFlight = busy;
	ExecuteAIConfigWebViewScript(ctx, busy
		? L"window.autolinkerSetTestBusy(true);"
		: L"window.autolinkerSetTestBusy(false);");
}

void ShowAIConfigWebViewTestResult(AIConfigWebViewDialogContext* ctx, const AIResult& result)
{
	if (ctx == nullptr) {
		return;
	}

	nlohmann::json payload;
	payload["ok"] = result.ok;
	payload["message"] = LocalToUtf8Text(BuildAIConnectionTestMessage(result));

	const std::wstring payloadJsonWide = Utf8ToWide(payload.dump());
	if (payloadJsonWide.empty()) {
		return;
	}

	std::wstring script = L"window.autolinkerShowTestResult(JSON.parse('";
	script += EscapeJsSingleQuotedWide(payloadJsonWide);
	script += L"'));";
	ExecuteAIConfigWebViewScript(ctx, script);
}

void ShowAIConfigWebViewModelListResult(AIConfigWebViewDialogContext* ctx, const AIConfigModelListResult& result)
{
	if (ctx == nullptr) {
		return;
	}

	nlohmann::json payload;
	payload["ok"] = result.ok;
	payload["error"] = LocalToUtf8Text(result.error);
	payload["httpStatus"] = result.httpStatus;
	payload["models"] = nlohmann::json::array();
	for (const std::string& model : result.models) {
		payload["models"].push_back(LocalToUtf8Text(model));
	}

	const std::wstring payloadJsonWide = Utf8ToWide(payload.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace));
	if (payloadJsonWide.empty()) {
		return;
	}

	std::wstring script = L"window.autolinkerShowModelListResult(JSON.parse('";
	script += EscapeJsSingleQuotedWide(payloadJsonWide);
	script += L"'));";
	ExecuteAIConfigWebViewScript(ctx, script);
}

void ApplyAIConfigWebViewSettings(AIConfigWebViewDialogContext* ctx)
{
	if (ctx == nullptr || ctx->settings == nullptr || !ctx->webViewReady) {
		return;
	}

	const std::string settingsJsonUtf8 = BuildAIConfigWebViewSettingsJson(ctx->profiles, ctx->activeProfileId);
	const std::wstring settingsJsonWide = Utf8ToWide(settingsJsonUtf8);
	if (settingsJsonWide.empty()) {
		OutputStringToELog("[AI Config][WebView2] settings json conversion failed");
		return;
	}

	std::wstring script = L"window.autolinkerApplySettings(JSON.parse('";
	script += EscapeJsSingleQuotedWide(settingsJsonWide);
	script += L"'));window.autolinkerFocusPrimary();";
	ExecuteAIConfigWebViewScript(ctx, script);
}

void StartAIConfigWebView(HWND hWnd, AIConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHost == nullptr) {
		return;
	}

	const std::wstring webViewUserDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		webViewUserDataFolder.empty() ? nullptr : webViewUserDataFolder.c_str(),
		nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT envResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* innerCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[AI Config][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					innerCtx->fallbackRequested = true;
					DestroyWindow(hWnd);
					return S_OK;
				}

				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* readyCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
							if (readyCtx == nullptr || !IsWindow(hWnd)) {
								return S_OK;
							}
							if (FAILED(controllerResult) || controller == nullptr) {
								OutputStringToELog(std::format("[AI Config][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							readyCtx->webViewController = controller;
							readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
							if (readyCtx->webView == nullptr) {
								OutputStringToELog("[AI Config][WebView2] get_CoreWebView2 returned null");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							Microsoft::WRL::ComPtr<ICoreWebView2Settings> webSettings;
							if (SUCCEEDED(readyCtx->webView->get_Settings(&webSettings)) && webSettings != nullptr) {
								webSettings->put_AreDevToolsEnabled(FALSE);
								webSettings->put_IsStatusBarEnabled(FALSE);
								webSettings->put_IsZoomControlEnabled(FALSE);
							}

							readyCtx->webView->add_WebMessageReceived(
								Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
										auto* messageCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (messageCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										LPWSTR rawMessage = nullptr;
										if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
											return S_OK;
										}

										const std::string utf8Message = WideToUtf8(rawMessage);
										CoTaskMemFree(rawMessage);
										try {
											const nlohmann::json payload = nlohmann::json::parse(utf8Message);
											const std::string action = payload.value("action", "");
											if (action == "save" && payload.contains("data") && payload["data"].is_object()) {
												TryApplyAISettingsFromWebPayload(hWnd, messageCtx, payload["data"]);
											}
											else if (action == "test_connection" && payload.contains("data") && payload["data"].is_object()) {
												if (!messageCtx->testInFlight) {
													const auto& data = payload["data"];
													const std::string activeProfileId = Utf8ToLocalText(data.value("activeProfileId", ""));
													AISettings next = *messageCtx->settings;
													if (data.contains("profiles") && data["profiles"].is_array()) {
														for (const auto& item : data["profiles"]) {
															if (!item.is_object()) {
																continue;
															}
															if (Utf8ToLocalText(item.value("id", "")) != activeProfileId) {
																continue;
															}
															next = ReadAISettingsFromWebProfilePayload(*messageCtx->settings, item);
															break;
														}
													}
													if (ValidateAISettingsForConnection(hWnd, next) &&
														StartAIConfigConnectionTest(hWnd, next, true)) {
														SetAIConfigWebViewTestBusy(messageCtx, true);
													}
												}
											}
											else if (action == "fetch_models" && payload.contains("data") && payload["data"].is_object()) {
												const AISettings next = ReadAISettingsFromWebProfilePayload(
													*messageCtx->settings,
													payload["data"]);
												if (!StartAIConfigModelListFetch(hWnd, next)) {
													ExecuteAIConfigWebViewScript(messageCtx, L"window.autolinkerSetModelFetchBusy(false);");
												}
											}
											else if (action == "cancel") {
												DestroyWindow(hWnd);
											}
											else if (action == "open_url") {
												const std::string openUrl = payload.value("url", "");
												if (!openUrl.empty() &&
													(openUrl.rfind("https://", 0) == 0 || openUrl.rfind("http://", 0) == 0)) {
													ShellExecuteA(hWnd, "open", openUrl.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
												}
											}
										}
										catch (...) {
										}
										return S_OK;
									}).Get(),
								nullptr);

							readyCtx->webView->add_NavigationCompleted(
								Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										auto* navCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										BOOL isSuccess = FALSE;
										args->get_IsSuccess(&isSuccess);
										if (isSuccess == TRUE) {
											navCtx->webViewReady = true;
											KillTimer(hWnd, kAIConfigWebViewInitTimerId);
											if (navCtx->hLoading != nullptr) {
												ShowWindow(navCtx->hLoading, SW_HIDE);
											}
											LayoutAIConfigWebViewDialog(hWnd, navCtx);
											ApplyAIConfigWebViewSettings(navCtx);
											return S_OK;
										}

										COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
										args->get_WebErrorStatus(&webErrorStatus);
										if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
											OutputStringToELog(std::format(
												"[AI Config][WebView2] navigation superseded errorStatus={}",
												static_cast<int>(webErrorStatus)));
											return S_OK;
										}

										OutputStringToELog(std::format(
											"[AI Config][WebView2] navigation failed errorStatus={}",
											static_cast<int>(webErrorStatus)));
										navCtx->fallbackRequested = true;
										DestroyWindow(hWnd);
										return S_OK;
									}).Get(),
								nullptr);

							LayoutAIConfigWebViewDialog(hWnd, readyCtx);
							const std::wstring shellHtml = Utf8ToWide(BuildAIConfigWebViewShellHtml());
							if (shellHtml.empty()) {
								OutputStringToELog("[AI Config][WebView2] shell html conversion failed");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}
							readyCtx->webView->NavigateToString(shellHtml.c_str());
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[AI Config][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		ctx->fallbackRequested = true;
		DestroyWindow(hWnd);
	}
}

LRESULT CALLBACK AIConfigWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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

		ctx->hHost = CreateWindowExA(
			0,
			"STATIC",
			"",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		ctx->hLoading = CreateWindowExW(
			0,
			L"STATIC",
			L"正在初始化 WebView2 设置页...",
			WS_CHILD | WS_VISIBLE,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		SetDefaultFont(ctx->hLoading);
		LayoutAIConfigWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kAIConfigWebViewInitTimerId, kAIConfigWebViewInitTimeoutMs, nullptr);
		StartAIConfigWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 820L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 620L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutAIConfigWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kAIConfigWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[AI Config][WebView2] initialization timed out, fallback to native dialog");
				ctx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kAIConfigWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_AUTOLINKER_AI_CONFIG_TEST_DONE: {
		AIConfigConnectionTestResult* result = reinterpret_cast<AIConfigConnectionTestResult*>(lParam);
		if (ctx == nullptr || result == nullptr) {
			delete result;
			return 0;
		}

		SetAIConfigWebViewTestBusy(ctx, false);
		ShowAIConfigWebViewTestResult(ctx, result->result);
		delete result;
		return 0;
	}

	case WM_AUTOLINKER_AI_CONFIG_MODEL_LIST_DONE: {
		AIConfigModelListResult* result = reinterpret_cast<AIConfigModelListResult*>(lParam);
		if (ctx == nullptr || result == nullptr) {
			delete result;
			return 0;
		}

		ShowAIConfigWebViewModelListResult(ctx, *result);
		delete result;
		return 0;
	}

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kAIConfigWebViewInitTimerId);
		if (ctx != nullptr) {
			ctx->webView = nullptr;
			ctx->webViewController = nullptr;
			ctx->webViewEnvironment = nullptr;
		}
		return 0;

	default:
		break;
	}

	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

std::string BuildAIPreviewWebViewPayloadJson(const AIPreviewWebViewDialogContext& ctx)
{
	nlohmann::json payload;
	payload["title"] = LocalToUtf8Text(ctx.title);
	payload["content"] = LocalToUtf8Text(ctx.content);
	payload["primaryText"] = LocalToUtf8Text(ctx.primaryText);
	payload["secondaryText"] = LocalToUtf8Text(ctx.secondaryText);
	return payload.dump();
}

std::string BuildAIPreviewWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_PREVIEW_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>Preview shell resource missing.</body></html>";
}

void LayoutAIPreviewWebViewDialog(HWND hWnd, AIPreviewWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom));
	const int loadingLeft = 16;
	const int loadingTop = 14;
	const int loadingWidth = static_cast<int>((std::max)(0L, rc.right - loadingLeft * 2L));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, 0, 0, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, loadingLeft, loadingTop, loadingWidth, 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = {};
		bounds.left = 0;
		bounds.top = 0;
		bounds.right = static_cast<LONG>(hostWidth);
		bounds.bottom = static_cast<LONG>(hostHeight);
		ctx->webViewController->put_Bounds(bounds);
	}
}

void ExecuteAIPreviewWebViewScript(AIPreviewWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplyAIPreviewWebViewPayload(AIPreviewWebViewDialogContext* ctx)
{
	if (ctx == nullptr || !ctx->webViewReady) {
		return;
	}

	const std::string payloadJsonUtf8 = BuildAIPreviewWebViewPayloadJson(*ctx);
	const std::wstring payloadJsonWide = Utf8ToWide(payloadJsonUtf8);
	if (payloadJsonWide.empty()) {
		OutputStringToELog("[AI Preview][WebView2] payload json conversion failed");
		return;
	}

	std::wstring script = L"window.autolinkerApplyPreview(JSON.parse('";
	script += EscapeJsSingleQuotedWide(payloadJsonWide);
	script += L"'));window.autolinkerFocusPrimary();";
	OutputStringToELog(std::format("[AI Preview][WebView2] apply payload jsonChars={}", payloadJsonWide.size()));
	ExecuteAIPreviewWebViewScript(ctx, script);
}

void StartAIPreviewWebView(HWND hWnd, AIPreviewWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHost == nullptr) {
		return;
	}

	const std::wstring webViewUserDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		webViewUserDataFolder.empty() ? nullptr : webViewUserDataFolder.c_str(),
		nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT envResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* innerCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[AI Preview][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					innerCtx->fallbackRequested = true;
					DestroyWindow(hWnd);
					return S_OK;
				}

				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* readyCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
							if (readyCtx == nullptr || !IsWindow(hWnd)) {
								return S_OK;
							}
							if (FAILED(controllerResult) || controller == nullptr) {
								OutputStringToELog(std::format("[AI Preview][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							readyCtx->webViewController = controller;
							readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
							if (readyCtx->webView == nullptr) {
								OutputStringToELog("[AI Preview][WebView2] get_CoreWebView2 returned null");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							Microsoft::WRL::ComPtr<ICoreWebView2Settings> webSettings;
							if (SUCCEEDED(readyCtx->webView->get_Settings(&webSettings)) && webSettings != nullptr) {
								webSettings->put_AreDevToolsEnabled(FALSE);
								webSettings->put_AreDefaultContextMenusEnabled(FALSE);
								webSettings->put_IsStatusBarEnabled(FALSE);
								webSettings->put_IsZoomControlEnabled(FALSE);
							}

							readyCtx->webView->add_WebMessageReceived(
								Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
										auto* messageCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (messageCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										LPWSTR rawMessage = nullptr;
										if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
											return S_OK;
										}

										const std::string utf8Message = WideToUtf8(rawMessage);
										CoTaskMemFree(rawMessage);
										try {
											const nlohmann::json payload = nlohmann::json::parse(utf8Message);
											const std::string action = payload.value("action", "");
											if (action == "primary") {
												messageCtx->action = AIPreviewAction::PrimaryConfirm;
												DestroyWindow(hWnd);
											}
											else if (action == "secondary") {
												messageCtx->action = AIPreviewAction::SecondaryConfirm;
												DestroyWindow(hWnd);
											}
											else if (action == "cancel") {
												messageCtx->action = AIPreviewAction::Cancel;
												DestroyWindow(hWnd);
											}
										}
										catch (...) {
										}
										return S_OK;
									}).Get(),
								nullptr);

							readyCtx->webView->add_NavigationCompleted(
								Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										auto* navCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										BOOL isSuccess = FALSE;
										args->get_IsSuccess(&isSuccess);
										if (isSuccess == TRUE) {
											navCtx->webViewReady = true;
											KillTimer(hWnd, kAIPreviewWebViewInitTimerId);
											if (navCtx->hLoading != nullptr) {
												ShowWindow(navCtx->hLoading, SW_HIDE);
											}
											OutputStringToELog("[AI Preview][WebView2] navigation completed successfully");
											LayoutAIPreviewWebViewDialog(hWnd, navCtx);
											ApplyAIPreviewWebViewPayload(navCtx);
											return S_OK;
										}

										COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
										args->get_WebErrorStatus(&webErrorStatus);
										if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
											OutputStringToELog(std::format(
												"[AI Preview][WebView2] navigation superseded errorStatus={}",
												static_cast<int>(webErrorStatus)));
											return S_OK;
										}

										OutputStringToELog(std::format(
											"[AI Preview][WebView2] navigation failed errorStatus={}",
											static_cast<int>(webErrorStatus)));
										navCtx->fallbackRequested = true;
										DestroyWindow(hWnd);
										return S_OK;
									}).Get(),
								nullptr);

							LayoutAIPreviewWebViewDialog(hWnd, readyCtx);
							const std::wstring shellHtml = Utf8ToWide(BuildAIPreviewWebViewShellHtml());
							if (shellHtml.empty()) {
								OutputStringToELog("[AI Preview][WebView2] shell html conversion failed");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}
							OutputStringToELog(std::format("[AI Preview][WebView2] controller ready=1 shellHtmlChars={}", shellHtml.size()));
							readyCtx->webView->NavigateToString(shellHtml.c_str());
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[AI Preview][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		ctx->fallbackRequested = true;
		DestroyWindow(hWnd);
	}
}

LRESULT CALLBACK AIPreviewWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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

		ctx->hHost = CreateWindowExA(
			0,
			"STATIC",
			"",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		ctx->hLoading = CreateWindowExW(
			0,
			L"STATIC",
			L"正在初始化 WebView2 确认窗口...",
			WS_CHILD | WS_VISIBLE,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		SetDefaultFont(ctx->hLoading);
		LayoutAIPreviewWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kAIPreviewWebViewInitTimerId, kAIPreviewWebViewInitTimeoutMs, nullptr);
		StartAIPreviewWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 760L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 520L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutAIPreviewWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kAIPreviewWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[AI Preview][WebView2] initialization timed out, fallback to native dialog");
				ctx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kAIPreviewWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_CLOSE:
		if (ctx != nullptr) {
			ctx->action = AIPreviewAction::Cancel;
		}
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kAIPreviewWebViewInitTimerId);
		if (ctx != nullptr) {
			ctx->webView = nullptr;
			ctx->webViewController = nullptr;
			ctx->webViewEnvironment = nullptr;
		}
		return 0;

	default:
		break;
	}

	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

struct AIPreviewDialogContext {
	std::string title;
	std::string content;
	std::string primaryText;
	std::string secondaryText;
	AIPreviewAction action = AIPreviewAction::Cancel;
	HWND hEdit = nullptr;
	HWND hPrimary = nullptr;
	HWND hSecondary = nullptr;
	HWND hCancel = nullptr;
};

struct AIInputDialogContext {
	std::string title;
	std::string hint;
	std::string text;
	bool accepted = false;
	HWND hEdit = nullptr;
};

void LayoutAIPreviewDialog(HWND hWnd, AIPreviewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int margin = 14;
	const int gap = 8;
	const int buttonW = 120;
	const int buttonH = 30;

	int right = static_cast<int>(rc.right) - margin;
	const int buttonTop = (std::max)(margin, static_cast<int>(rc.bottom) - margin - buttonH);

	if (ctx->hCancel != nullptr) {
		MoveWindow(ctx->hCancel, right - buttonW, buttonTop, buttonW, buttonH, TRUE);
		right -= (buttonW + gap);
	}
	if (ctx->hSecondary != nullptr) {
		MoveWindow(ctx->hSecondary, right - buttonW, buttonTop, buttonW, buttonH, TRUE);
		right -= (buttonW + gap);
	}
	if (ctx->hPrimary != nullptr) {
		MoveWindow(ctx->hPrimary, right - buttonW, buttonTop, buttonW, buttonH, TRUE);
	}

	const int editLeft = margin;
	const int editTop = margin;
	const int editWidth = (std::max)(120, static_cast<int>(rc.right) - margin * 2);
	const int editHeight = (std::max)(120, buttonTop - margin - editTop);
	if (ctx->hEdit != nullptr) {
		MoveWindow(ctx->hEdit, editLeft, editTop, editWidth, editHeight, TRUE);
	}
}

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

		ctx->hPrimary = CreateWindowA("BUTTON", ctx->primaryText.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			590, 474, 120, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_PREVIEW_OK), nullptr, nullptr);

		if (!ctx->secondaryText.empty()) {
			ctx->hSecondary = CreateWindowA("BUTTON", ctx->secondaryText.c_str(),
				WS_CHILD | WS_VISIBLE | WS_TABSTOP,
				460, 474, 120, 30, hWnd,
				reinterpret_cast<HMENU>(IDC_PREVIEW_SECONDARY), nullptr, nullptr);
		}

		ctx->hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			682, 474, 120, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_PREVIEW_CANCEL), nullptr, nullptr);

		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(ctx->hPrimary);
		SetDefaultFont(ctx->hSecondary);
		SetDefaultFont(ctx->hCancel);
		LayoutAIPreviewDialog(hWnd, ctx);
		SetFocus(ctx->hPrimary != nullptr ? ctx->hPrimary : ctx->hEdit);
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutAIPreviewDialog(hWnd, ctx);
		}
		return 0;

	case WM_GETMINMAXINFO: {
		auto* mm = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mm != nullptr) {
			mm->ptMinTrackSize.x = 680;
			mm->ptMinTrackSize.y = 420;
		}
		return 0;
	}

	case WM_COMMAND: {
		const int id = LOWORD(wParam);
		if (ctx == nullptr) {
			return 0;
		}
		if (id == IDC_PREVIEW_OK) {
			ctx->action = AIPreviewAction::PrimaryConfirm;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_PREVIEW_SECONDARY) {
			ctx->action = AIPreviewAction::SecondaryConfirm;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_PREVIEW_CANCEL) {
			ctx->action = AIPreviewAction::Cancel;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_CLOSE:
		if (ctx != nullptr) {
			ctx->action = AIPreviewAction::Cancel;
		}
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

		HWND hOk = CreateWindowW(L"BUTTON", L"\u786E\u5B9A",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 448, 254, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_INPUT_OK), nullptr, nullptr);
		HWND hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88",
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

void CenterWindowOnOwnerOrScreen(HWND hDialog, HWND owner)
{
	if (hDialog == nullptr || !IsWindow(hDialog)) {
		return;
	}

	RECT dlgRc = {};
	if (!GetWindowRect(hDialog, &dlgRc)) {
		return;
	}
	const int dlgW = dlgRc.right - dlgRc.left;
	const int dlgH = dlgRc.bottom - dlgRc.top;

	// Reference rect: owner window if available, otherwise the work area of
	// the monitor nearest the dialog.
	RECT refRc = {};
	bool haveRef = false;
	if (owner != nullptr && IsWindow(owner) && !IsIconic(owner)) {
		haveRef = GetWindowRect(owner, &refRc) != FALSE;
	}
	if (!haveRef) {
		HMONITOR hMon = MonitorFromWindow(hDialog, MONITOR_DEFAULTTOPRIMARY);
		MONITORINFO mi = {};
		mi.cbSize = sizeof(mi);
		if (GetMonitorInfoA(hMon, &mi)) {
			refRc = mi.rcWork;
			haveRef = true;
		}
	}
	if (!haveRef) {
		return;
	}

	int x = refRc.left + ((refRc.right - refRc.left) - dlgW) / 2;
	int y = refRc.top + ((refRc.bottom - refRc.top) - dlgH) / 2;

	// Clamp to the work area of the monitor the centered position lands on.
	HMONITOR hMon = MonitorFromPoint(POINT{ x + dlgW / 2, y + dlgH / 2 }, MONITOR_DEFAULTTONEAREST);
	MONITORINFO mi = {};
	mi.cbSize = sizeof(mi);
	if (GetMonitorInfoA(hMon, &mi)) {
		const RECT& work = mi.rcWork;
		if (x + dlgW > work.right) {
			x = work.right - dlgW;
		}
		if (y + dlgH > work.bottom) {
			y = work.bottom - dlgH;
		}
		if (x < work.left) {
			x = work.left;
		}
		if (y < work.top) {
			y = work.top;
		}
	}

	SetWindowPos(hDialog, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

bool RunModalWindow(HWND owner, HWND hDialog)
{
	if (hDialog == nullptr) {
		return false;
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, FALSE);
	}

	CenterWindowOnOwnerOrScreen(hDialog, owner);
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

bool ShowAIConfigDialogNative(HWND owner, AIJsonConfig& jsonConfig, AISettings& ioSettings)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIConfigDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIConfigDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIConfigDialogContext ctx = {};
	ctx.jsonConfig = &jsonConfig;
	ctx.settings = &ioSettings;
	ctx.profiles = LoadProfileEntriesFromJsonConfig(jsonConfig, ioSettings, ctx.activeProfileId);

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 780, 560,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		return false;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, "AutoLinker AI Config");
	RunModalWindow(owner, hDialog);
	return ctx.accepted;
}

AIConfigDialogRunResult ShowAIConfigDialogWebView(HWND owner, AIJsonConfig& jsonConfig, AISettings& ioSettings)
{
	AIConfigDialogRunResult result = {};
	if (!IsWebView2RuntimeAvailable()) {
		OutputStringToELog("[AI Config][WebView2] runtime unavailable, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIConfigWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIConfigWebViewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIConfigWebViewDialogContext ctx = {};
	ctx.jsonConfig = &jsonConfig;
	ctx.settings = &ioSettings;
	ctx.profiles = LoadProfileEntriesFromJsonConfig(jsonConfig, ioSettings, ctx.activeProfileId);

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 700, 860,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[AI Config][WebView2] CreateWindowExA failed, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, "AutoLinker AI Config");
	RunModalWindow(owner, hDialog);
	result.accepted = ctx.accepted;
	result.fallbackRequested = ctx.fallbackRequested;
	return result;
}

bool ShowAIConfigDialog(HWND owner, AIJsonConfig& jsonConfig, AISettings& ioSettings)
{
	const AIConfigDialogRunResult webViewResult = ShowAIConfigDialogWebView(owner, jsonConfig, ioSettings);
	if (webViewResult.fallbackRequested) {
		OutputStringToELog("[AI Config][WebView2] fallback to native dialog");
		return ShowAIConfigDialogNative(owner, jsonConfig, ioSettings);
	}
	return webViewResult.accepted;
}

AIPreviewAction ShowAIPreviewDialogNative(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIPreviewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIPreviewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIPreviewDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Preview" : title;
	ctx.content = content;
	ctx.primaryText = primaryText.empty() ? "\u786E\u5B9A" : primaryText;
	ctx.secondaryText = secondaryText;
	ctx.action = AIPreviewAction::Cancel;

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
		return AIPreviewAction::Cancel;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, ctx.title);
	RunModalWindow(owner, hDialog);
	return ctx.action;
}

AIPreviewWebViewRunResult ShowAIPreviewDialogWebView(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText)
{
	AIPreviewWebViewRunResult result = {};
	if (!IsWebView2RuntimeAvailableWithTag("AI Preview")) {
		OutputStringToELog("[AI Preview][WebView2] runtime unavailable, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIPreviewWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIPreviewWebViewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIPreviewWebViewDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Preview" : title;
	ctx.content = content;
	ctx.primaryText = primaryText.empty() ? "\u786E\u5B9A" : primaryText;
	ctx.secondaryText = secondaryText;
	ctx.action = AIPreviewAction::Cancel;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		ctx.title.c_str(),
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 960, 700,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[AI Preview][WebView2] CreateWindowExA failed, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, ctx.title);
	RunModalWindow(owner, hDialog);
	result.action = ctx.action;
	result.fallbackRequested = ctx.fallbackRequested;
	return result;
}

AIPreviewAction ShowAIPreviewDialogEx(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText)
{
	const AIPreviewWebViewRunResult webViewResult = ShowAIPreviewDialogWebView(
		owner,
		title,
		content,
		primaryText,
		secondaryText);
	if (webViewResult.fallbackRequested) {
		OutputStringToELog("[AI Preview][WebView2] fallback to native dialog");
		return ShowAIPreviewDialogNative(owner, title, content, primaryText, secondaryText);
	}
	return webViewResult.action;
}

bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText)
{
	return ShowAIPreviewDialogEx(owner, title, content, confirmText, "") == AIPreviewAction::PrimaryConfirm;
}

bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIInputDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIInputDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIInputDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Input" : title;
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

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, ctx.title);
	RunModalWindow(owner, hDialog);
	if (ctx.accepted) {
		ioText = ctx.text;
	}
	return ctx.accepted;
}

namespace {

constexpr UINT WM_AUTOLINKER_LINKER_SAVE_DONE = WM_APP + 303;

struct LinkerEntry {
	std::string name;     // UTF-8 显示名（即 .ini 文件名，不含后缀）
	std::string content;  // UTF-8 文件内容
};

struct LinkerConfigWebViewDialogContext {
	bool webViewReady = false;
	bool fallbackRequested = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	std::vector<LinkerEntry> linkers;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

// link.ini 为易语言使用的 GBK/ANSI 文本：读取时若非 UTF-8 则按 936 转 UTF-8，
// 写入时再由 UTF-8 转回 936，避免中文乱码。
std::string LinkerFileBytesToUtf8(const std::string& bytes)
{
	if (bytes.empty()) {
		return std::string();
	}
	if (bytes.size() >= 3 &&
		static_cast<unsigned char>(bytes[0]) == 0xEF &&
		static_cast<unsigned char>(bytes[1]) == 0xBB &&
		static_cast<unsigned char>(bytes[2]) == 0xBF) {
		return bytes.substr(3);
	}
	if (IsValidUtf8Text(bytes)) {
		return bytes;
	}
	return ConvertCodePage(bytes, 936, CP_UTF8, 0);
}

std::string Utf8ToLinkerFileBytes(const std::string& utf8)
{
	if (utf8.empty()) {
		return std::string();
	}
	if (!IsValidUtf8Text(utf8)) {
		return utf8;
	}
	return ConvertCodePage(utf8, CP_UTF8, 936, 0);
}

std::string NormalizeToCrlf(const std::string& text)
{
	std::string out;
	out.reserve(text.size() + 16);
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			out += "\r\n";
		}
		else if (ch == '\n') {
			out += "\r\n";
		}
		else {
			out.push_back(ch);
		}
	}
	return out;
}

std::filesystem::path GetLinkerConfigDirectoryPath()
{
	std::filesystem::path dir = GetAutoLinkerDirectoryPath() / "Config";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir;
}

std::string SanitizeLinkerName(const std::string& name)
{
	std::string out;
	out.reserve(name.size());
	for (char ch : name) {
		switch (ch) {
		case '\\': case '/': case ':': case '*':
		case '?': case '"': case '<': case '>': case '|':
			out.push_back('_');
			break;
		default:
			out.push_back(ch);
			break;
		}
	}
	// 去除首尾空白与点号（Windows 文件名不允许结尾点/空格）。
	const auto notTrim = [](char c) { return c != ' ' && c != '.' && c != '\t'; };
	auto begin = std::find_if(out.begin(), out.end(), notTrim);
	auto end = std::find_if(out.rbegin(), out.rend(), notTrim).base();
	if (begin >= end) {
		return std::string();
	}
	return std::string(begin, end);
}

std::vector<LinkerEntry> LoadLinkerEntriesFromDisk()
{
	std::vector<LinkerEntry> entries;
	const std::filesystem::path dir = GetLinkerConfigDirectoryPath();
	std::error_code ec;
	if (!std::filesystem::exists(dir, ec)) {
		return entries;
	}
	for (const auto& item : std::filesystem::directory_iterator(dir, ec)) {
		if (ec) {
			break;
		}
		if (!item.is_regular_file() || item.path().extension() != ".ini") {
			continue;
		}
		LinkerEntry entry;
		entry.name = WideToUtf8(item.path().stem().wstring());
		std::ifstream in(item.path(), std::ios::binary);
		if (in.is_open()) {
			std::string bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
			entry.content = LinkerFileBytesToUtf8(bytes);
		}
		entries.push_back(std::move(entry));
	}
	std::sort(entries.begin(), entries.end(), [](const LinkerEntry& a, const LinkerEntry& b) {
		return _stricmp(a.name.c_str(), b.name.c_str()) < 0;
	});
	return entries;
}

// 将 webview 提交的链接器列表写回 Config 目录：写入/覆盖所有提交项，
// 删除目录中未出现的 .ini（涵盖重命名与删除）。返回 false 并填充 errorMessage 表示失败。
bool SaveLinkerEntriesToDisk(const std::vector<LinkerEntry>& entries, std::string& errorMessage)
{
	const std::filesystem::path dir = GetLinkerConfigDirectoryPath();
	std::set<std::wstring> keepStems;

	for (const auto& entry : entries) {
		const std::string safeName = SanitizeLinkerName(entry.name);
		if (safeName.empty()) {
			errorMessage = "存在名称为空的链接器，已取消保存。";
			return false;
		}
		const std::wstring wideStem = Utf8ToWide(safeName);
		keepStems.insert(wideStem);
		std::filesystem::path file = dir / (wideStem + L".ini");
		std::ofstream out(file, std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			errorMessage = "无法写入链接器文件：" + safeName + ".ini";
			return false;
		}
		const std::string bytes = Utf8ToLinkerFileBytes(NormalizeToCrlf(entry.content));
		out.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
	}

	std::error_code ec;
	for (const auto& item : std::filesystem::directory_iterator(dir, ec)) {
		if (ec) {
			break;
		}
		if (!item.is_regular_file() || item.path().extension() != ".ini") {
			continue;
		}
		if (keepStems.find(item.path().stem().wstring()) == keepStems.end()) {
			std::error_code removeEc;
			std::filesystem::remove(item.path(), removeEc);
		}
	}
	return true;
}

std::string BuildLinkerWebViewPayloadJson(const std::vector<LinkerEntry>& linkers)
{
	nlohmann::json payload;
	payload["linkers"] = nlohmann::json::array();
	for (const auto& entry : linkers) {
		nlohmann::json item;
		item["name"] = entry.name;
		item["content"] = entry.content;
		payload["linkers"].push_back(std::move(item));
	}
	return payload.dump();
}

std::string BuildLinkerWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_LINKER_CONFIG_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>Linker config shell resource missing.</body></html>";
}

void ExecuteLinkerWebViewScript(LinkerConfigWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplyLinkerWebViewData(LinkerConfigWebViewDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	const std::wstring payloadWide = Utf8ToWide(BuildLinkerWebViewPayloadJson(ctx->linkers));
	std::wstring script = L"window.autolinkerApplyLinkers(JSON.parse('";
	script += EscapeJsSingleQuotedWide(payloadWide);
	script += L"'));";
	ExecuteLinkerWebViewScript(ctx, script);
}

void NotifyLinkerSaveResult(LinkerConfigWebViewDialogContext* ctx, bool ok, const std::string& message)
{
	nlohmann::json result;
	result["ok"] = ok;
	result["message"] = message;
	std::wstring script = L"window.autolinkerLinkerSaveResult(JSON.parse('";
	script += EscapeJsSingleQuotedWide(Utf8ToWide(result.dump()));
	script += L"'));";
	ExecuteLinkerWebViewScript(ctx, script);
}

void LayoutLinkerConfigWebViewDialog(HWND hWnd, LinkerConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}
	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, 0, 0, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, 16, 14, (std::max)(0, hostWidth - 32), 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = { 0, 0, static_cast<LONG>(hostWidth), static_cast<LONG>(hostHeight) };
		ctx->webViewController->put_Bounds(bounds);
	}
}

void HandleLinkerWebViewSave(HWND hWnd, LinkerConfigWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr) {
		return;
	}
	std::vector<LinkerEntry> next;
	std::set<std::string> seenLower;
	if (data.contains("linkers") && data["linkers"].is_array()) {
		for (const auto& item : data["linkers"]) {
			if (!item.is_object()) {
				continue;
			}
			LinkerEntry entry;
			entry.name = SanitizeLinkerName(item.value("name", std::string()));
			entry.content = item.value("content", std::string());
			if (entry.name.empty()) {
				NotifyLinkerSaveResult(ctx, false, "链接器名称不能为空。");
				return;
			}
			std::string lower = entry.name;
			std::transform(lower.begin(), lower.end(), lower.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
			if (!seenLower.insert(lower).second) {
				NotifyLinkerSaveResult(ctx, false, "存在重名链接器：" + entry.name);
				return;
			}
			next.push_back(std::move(entry));
		}
	}

	std::string errorMessage;
	if (!SaveLinkerEntriesToDisk(next, errorMessage)) {
		NotifyLinkerSaveResult(ctx, false, errorMessage.empty() ? "保存失败。" : errorMessage);
		return;
	}
	ctx->linkers = std::move(next);
	NotifyLinkerSaveResult(ctx, true, std::format("已保存 {} 个链接器配置。", ctx->linkers.size()));
}

HRESULT OnLinkerConfigControllerCreated(HWND hWnd, HRESULT controllerResult, ICoreWebView2Controller* controller)
{
	auto* readyCtx = reinterpret_cast<LinkerConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	if (readyCtx == nullptr || !IsWindow(hWnd)) {
		return S_OK;
	}
	if (FAILED(controllerResult) || controller == nullptr) {
		OutputStringToELog(std::format("[Linker Config][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
		readyCtx->fallbackRequested = true;
		DestroyWindow(hWnd);
		return S_OK;
	}

	readyCtx->webViewController = controller;
	readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
	if (readyCtx->webView == nullptr) {
		readyCtx->fallbackRequested = true;
		DestroyWindow(hWnd);
		return S_OK;
	}

	Microsoft::WRL::ComPtr<ICoreWebView2Settings> webSettings;
	if (SUCCEEDED(readyCtx->webView->get_Settings(&webSettings)) && webSettings != nullptr) {
		webSettings->put_AreDevToolsEnabled(FALSE);
		webSettings->put_IsStatusBarEnabled(FALSE);
		webSettings->put_IsZoomControlEnabled(FALSE);
	}

	readyCtx->webView->add_WebMessageReceived(
		Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
			[hWnd](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
				auto* messageCtx = reinterpret_cast<LinkerConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (messageCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				LPWSTR rawMessage = nullptr;
				if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
					return S_OK;
				}
				const std::string utf8Message = WideToUtf8(rawMessage);
				CoTaskMemFree(rawMessage);
				try {
					const nlohmann::json payload = nlohmann::json::parse(utf8Message);
					const std::string action = payload.value("action", "");
					if (action == "save" && payload.contains("data") && payload["data"].is_object()) {
						HandleLinkerWebViewSave(hWnd, messageCtx, payload["data"]);
					}
					else if (action == "cancel") {
						DestroyWindow(hWnd);
					}
				}
				catch (...) {
				}
				return S_OK;
			}).Get(),
		nullptr);

	readyCtx->webView->add_NavigationCompleted(
		Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
			[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
				auto* navCtx = reinterpret_cast<LinkerConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				BOOL isSuccess = FALSE;
				args->get_IsSuccess(&isSuccess);
				if (isSuccess == TRUE) {
					navCtx->webViewReady = true;
					KillTimer(hWnd, kLinkerConfigWebViewInitTimerId);
					if (navCtx->hLoading != nullptr) {
						ShowWindow(navCtx->hLoading, SW_HIDE);
					}
					LayoutLinkerConfigWebViewDialog(hWnd, navCtx);
					ApplyLinkerWebViewData(navCtx);
					return S_OK;
				}
				COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
				args->get_WebErrorStatus(&webErrorStatus);
				if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
					webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
					webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
					return S_OK;
				}
				navCtx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return S_OK;
			}).Get(),
		nullptr);

	LayoutLinkerConfigWebViewDialog(hWnd, readyCtx);
	const std::wstring shellHtml = Utf8ToWide(BuildLinkerWebViewShellHtml());
	if (shellHtml.empty()) {
		readyCtx->fallbackRequested = true;
		DestroyWindow(hWnd);
		return S_OK;
	}
	readyCtx->webView->NavigateToString(shellHtml.c_str());
	return S_OK;
}

void StartLinkerConfigWebView(HWND hWnd, LinkerConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHost == nullptr) {
		return;
	}

	const std::wstring webViewUserDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		webViewUserDataFolder.empty() ? nullptr : webViewUserDataFolder.c_str(),
		nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT envResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* innerCtx = reinterpret_cast<LinkerConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[Linker Config][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					innerCtx->fallbackRequested = true;
					DestroyWindow(hWnd);
					return S_OK;
				}
				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							return OnLinkerConfigControllerCreated(hWnd, controllerResult, controller);
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[Linker Config][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		ctx->fallbackRequested = true;
		DestroyWindow(hWnd);
	}
}

LRESULT CALLBACK LinkerConfigWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<LinkerConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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
		ctx->hHost = CreateWindowExA(
			0, "STATIC", "",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0, 0, 0, 0, hWnd, nullptr, nullptr, nullptr);
		ctx->hLoading = CreateWindowExW(
			0, L"STATIC", L"正在初始化 WebView2 链接器设置页...",
			WS_CHILD | WS_VISIBLE,
			0, 0, 0, 0, hWnd, nullptr, nullptr, nullptr);
		SetDefaultFont(ctx->hLoading);
		LayoutLinkerConfigWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kLinkerConfigWebViewInitTimerId, kLinkerConfigWebViewInitTimeoutMs, nullptr);
		StartLinkerConfigWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 640L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 520L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutLinkerConfigWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kLinkerConfigWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[Linker Config][WebView2] initialization timed out");
				ctx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kLinkerConfigWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kLinkerConfigWebViewInitTimerId);
		if (ctx != nullptr) {
			ctx->webView = nullptr;
			ctx->webViewController = nullptr;
			ctx->webViewEnvironment = nullptr;
		}
		return 0;

	default:
		break;
	}
	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

} // namespace

void ShowLinkerConfigDialog(HWND owner)
{
	if (!IsWebView2RuntimeAvailable()) {
		MessageBoxW(owner,
			L"未检测到 WebView2 运行时，无法打开链接器设置页。\n请安装 Microsoft Edge WebView2 Runtime 后重试。",
			L"AutoLinker 链接器设置",
			MB_OK | MB_ICONWARNING);
		return;
	}

	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = LinkerConfigWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerLinkerConfigWebViewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	LinkerConfigWebViewDialogContext ctx = {};
	ctx.linkers = LoadLinkerEntriesFromDisk();

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker Linker Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 760, 720,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[Linker Config][WebView2] CreateWindowExA failed");
		return;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, "AutoLinker 链接器设置");
	RunModalWindow(owner, hDialog);
}





