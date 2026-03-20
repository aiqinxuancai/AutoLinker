#include "AIConfigDialog.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <CommCtrl.h>
#include <format>
#include <Shellapi.h>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "Global.h"

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
constexpr int IDC_CFG_FILL_RIGHT_CODES = 1006;
constexpr int IDC_CFG_TAVILY_API_KEY = 1007;
constexpr int IDC_CFG_SAVE = 1;
constexpr int IDC_CFG_CANCEL = 2;

constexpr int IDC_PREVIEW_EDIT = 1101;
constexpr int IDC_PREVIEW_OK = 1;
constexpr int IDC_PREVIEW_SECONDARY = 1102;
constexpr int IDC_PREVIEW_CANCEL = 2;

constexpr int IDC_INPUT_EDIT = 1201;
constexpr int IDC_INPUT_OK = 1;
constexpr int IDC_INPUT_CANCEL = 2;

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

bool IsWebView2RuntimeAvailable()
{
	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	const bool available = SUCCEEDED(hr);
	OutputStringToELog(std::format("[AI Config][WebView2] runtime available={}", available ? 1 : 0));
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return available;
}

struct AIConfigDialogRunResult {
	bool accepted = false;
	bool fallbackRequested = false;
};

struct AIConfigDialogContext {
	AISettings* settings = nullptr;
	bool accepted = false;
	bool useNativeLink = false;
	HWND hProtocol = nullptr;
	HWND hBaseUrl = nullptr;
	HWND hApiKey = nullptr;
	HWND hModel = nullptr;
	HWND hTavilyApiKey = nullptr;
	HWND hExtraPrompt = nullptr;
	HWND hGetKeyLink = nullptr;
};

struct AIConfigWebViewDialogContext {
	AISettings* settings = nullptr;
	bool accepted = false;
	bool fallbackRequested = false;
	bool webViewReady = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

void PopulateProtocolCombo(HWND hCombo, AIProtocolType selected)
{
	if (hCombo == nullptr) {
		return;
	}

	SendMessageA(hCombo, CB_RESETCONTENT, 0, 0);
	const int idxOpenAI = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("OpenAI")));
	const int idxGemini = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("Gemini")));
	const int idxClaude = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("Claude")));
	SendMessageA(hCombo, CB_SETITEMDATA, idxOpenAI, static_cast<LPARAM>(AIProtocolType::OpenAI));
	SendMessageA(hCombo, CB_SETITEMDATA, idxGemini, static_cast<LPARAM>(AIProtocolType::Gemini));
	SendMessageA(hCombo, CB_SETITEMDATA, idxClaude, static_cast<LPARAM>(AIProtocolType::Claude));

	int selectedIndex = idxOpenAI;
	if (selected == AIProtocolType::Gemini) {
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
	case AIProtocolType::Gemini:
		return AIProtocolType::Gemini;
	case AIProtocolType::Claude:
		return AIProtocolType::Claude;
	case AIProtocolType::OpenAI:
	default:
		return AIProtocolType::OpenAI;
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

std::string BuildAIConfigWebViewHtml(const AISettings& settings)
{
	nlohmann::json initialSettings;
	initialSettings["protocolType"] = AIService::ProtocolTypeToString(settings.protocolType);
	initialSettings["baseUrl"] = LocalToUtf8Text(settings.baseUrl);
	initialSettings["apiKey"] = LocalToUtf8Text(settings.apiKey);
	initialSettings["model"] = LocalToUtf8Text(settings.model);
	initialSettings["extraPrompt"] = LocalToUtf8Text(settings.extraSystemPrompt);
	initialSettings["tavilyApiKey"] = LocalToUtf8Text(settings.tavilyApiKey);

	std::string html;
	html.reserve(7000);
	html += "<!doctype html><html><head><meta charset=\"utf-8\"><meta name=\"color-scheme\" content=\"light only\">";
	html += "<style>";
	html += "html,body{margin:0;padding:0;background:#f4f6f8;color:#1f2937;font:14px/1.5 'Microsoft YaHei UI','Segoe UI',sans-serif;}";
	html += "*{box-sizing:border-box;} body{padding:16px;} h1{margin:0 0 8px 0;font-size:22px;} h2{margin:0 0 10px 0;font-size:16px;} p{margin:0 0 10px 0;color:#4b5563;} .wrap{max-width:1100px;margin:0 auto;} .panel{background:#fff;border:1px solid #d7dce2;border-radius:10px;padding:16px 18px;margin-bottom:14px;} .row{margin-bottom:12px;} .row label{display:block;margin-bottom:6px;font-weight:700;color:#374151;} .input,.select,.textarea{width:100%;border:1px solid #c7d0da;border-radius:8px;padding:9px 10px;background:#fff;font:14px/1.5 'Microsoft YaHei UI','Segoe UI',sans-serif;color:#111827;} .textarea{min-height:140px;resize:vertical;} .inline{display:flex;gap:8px;align-items:center;} .inline .input{flex:1 1 auto;} .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;} .btns{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px;} .btn{border:1px solid #bfc8d4;background:#fff;border-radius:8px;padding:8px 14px;cursor:pointer;font:600 13px/1 'Microsoft YaHei UI','Segoe UI',sans-serif;} .btn.primary{background:#1565c0;border-color:#1565c0;color:#fff;} .btn.secondary{background:#eb6c2d;border-color:#eb6c2d;color:#fff;} .tip{font-size:12px;color:#6b7280;} .code{font-family:Consolas,'Courier New',monospace;background:#eef2f7;border-radius:4px;padding:1px 5px;} @media (max-width:900px){.grid{grid-template-columns:1fr;}}";
	html += "</style></head><body><div class='wrap'><div class='panel'><h1>AutoLinker AI 设置</h1><p>统一配置主模型参数和 Tavily 联网搜索密钥。保存时由宿主程序完成最终校验。</p></div><div class='grid'><div class='panel'><h2>主模型</h2><div class='row'><label for='protocol'>Protocol</label><select id='protocol' class='select'><option value='OpenAI'>OpenAI</option><option value='Gemini'>Gemini</option><option value='Claude'>Claude</option></select></div><div class='row'><label for='baseUrl'>Base URL</label><input id='baseUrl' class='input' placeholder='https://right.codes/codex' /></div><div class='row'><label for='apiKey'>API Key</label><div class='inline'><input id='apiKey' class='input' type='password' placeholder='主模型 API Key' /><button id='toggleApiKey' class='btn' type='button'>显示</button></div></div><div class='row'><label for='model'>Model</label><input id='model' class='input' placeholder='例如 gpt-5.2-medium' /></div><div class='row'><label for='extraPrompt'>System Prompt</label><textarea id='extraPrompt' class='textarea' placeholder='附加系统提示词'></textarea></div><div class='btns'><button id='fillRightCodes' class='btn secondary' type='button'>一键填入 right.codes</button><button id='openRightCodes' class='btn' type='button'>打开 right.codes</button></div><p class='tip'>主模型配置为空时无法开始 AI 对话。</p></div>";
	html += "<div class='panel'><h2>Tavily 联网搜索</h2><div class='row'><label for='tavilyApiKey'>Tavily API Key</label><div class='inline'><input id='tavilyApiKey' class='input' type='password' placeholder='用于 search_web_tavily' /><button id='toggleTavilyKey' class='btn' type='button'>显示</button></div></div><p>这个 Key 仅供 <span class='code'>search_web_tavily</span> 工具使用，不影响主模型配置。</p><p class='tip'>PowerShell 命令执行工具仍然会逐次弹窗确认。</p><div class='btns'><button id='cancelBtn' class='btn' type='button'>取消</button><button id='saveBtn' class='btn primary' type='button'>保存并继续</button></div></div></div><script>";
	html += "const initialSettings=" + initialSettings.dump() + ";";
	html += "function $(id){return document.getElementById(id);} function post(payload){if(window.chrome&&window.chrome.webview){window.chrome.webview.postMessage(JSON.stringify(payload));}}";
	html += "function setField(id,value){const el=$(id); if(el){el.value=value||'';}}";
	html += "function toggleSecret(id,btnId){const input=$(id),btn=$(btnId); if(!input||!btn){return;} const next=input.type==='password'?'text':'password'; input.type=next; btn.textContent=next==='password'?'显示':'隐藏';}";
	html += "function collect(){return {protocol_type:$('protocol').value||'OpenAI',base_url:$('baseUrl').value||'',api_key:$('apiKey').value||'',model:$('model').value||'',extra_system_prompt:$('extraPrompt').value||'',tavily_api_key:$('tavilyApiKey').value||''};}";
	html += "function validate(data){if(!data.base_url.trim()||!data.api_key.trim()||!data.model.trim()){alert('baseUrl / apiKey / model 不能为空。'); return false;} return true;}";
	html += "function applyInitial(){var protocol=$('protocol'); if(protocol){protocol.value='OpenAI'; var desired=initialSettings.protocolType||'OpenAI'; protocol.value=desired; if(!protocol.value){protocol.selectedIndex=0;}} setField('baseUrl',initialSettings.baseUrl);setField('apiKey',initialSettings.apiKey);setField('model',initialSettings.model);setField('extraPrompt',initialSettings.extraPrompt);setField('tavilyApiKey',initialSettings.tavilyApiKey);}";
	html += "$('toggleApiKey').addEventListener('click',function(){toggleSecret('apiKey','toggleApiKey');}); $('toggleTavilyKey').addEventListener('click',function(){toggleSecret('tavilyApiKey','toggleTavilyKey');});";
	html += "$('fillRightCodes').addEventListener('click',function(){$('protocol').value='OpenAI';$('baseUrl').value='https://right.codes/codex';$('model').value='gpt-5.2-medium';$('apiKey').focus();});";
	html += "$('openRightCodes').addEventListener('click',function(){post({action:'open_right_codes'});});";
	html += "$('cancelBtn').addEventListener('click',function(){post({action:'cancel'});});";
	html += "$('saveBtn').addEventListener('click',function(){const data=collect(); if(!validate(data)){return;} post({action:'save',data:data});});";
	html += "applyInitial(); setTimeout(function(){if($('baseUrl')){$('baseUrl').focus();}},0);";
	html += "</script></body></html>";
	return html;
}

void LayoutAIConfigWebViewDialog(HWND hWnd, AIConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int margin = 12;
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right - margin * 2L));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom - margin * 2L));
	const int loadingWidth = static_cast<int>((std::max)(0L, rc.right - margin * 2L - 24L));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, margin, margin, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, margin + 12, margin + 12, loadingWidth, 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = {};
		bounds.left = margin;
		bounds.top = margin;
		bounds.right = static_cast<LONG>(margin + hostWidth);
		bounds.bottom = static_cast<LONG>(margin + hostHeight);
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

		HWND hProtocolLabel = CreateWindowA("STATIC", "Protocol:", WS_CHILD | WS_VISIBLE,
			16, 16, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hProtocol = CreateWindowExA(0, "COMBOBOX", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
			120, 14, 180, 220, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROTOCOL), nullptr, nullptr);
		PopulateProtocolCombo(ctx->hProtocol, ctx->settings->protocolType);

		HWND hBaseUrlLabel = CreateWindowA("STATIC", "Base URL:", WS_CHILD | WS_VISIBLE,
			16, 50, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hBaseUrl = CreateWindowExA(0, "EDIT", ctx->settings->baseUrl.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 48, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_BASE_URL), nullptr, nullptr);

		HWND hApiKeyLabel = CreateWindowA("STATIC", "API Key:", WS_CHILD | WS_VISIBLE,
			16, 84, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hApiKey = CreateWindowExA(0, "EDIT", ctx->settings->apiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 82, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_API_KEY), nullptr, nullptr);

		HWND hModelLabel = CreateWindowA("STATIC", "Model:", WS_CHILD | WS_VISIBLE,
			16, 118, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hModel = CreateWindowExA(0, "EDIT", ctx->settings->model.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 116, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_MODEL), nullptr, nullptr);

		const std::wstring getKeyLinkText =
			L"<a href=\"https://right.codes/register?aff=3dc87885\">从转发平台获取Key</a>";
		HWND hGetKeyLink = CreateWindowExW(0, L"SysLink",
			getKeyLinkText.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			120, 146, 240, 22, hWnd, reinterpret_cast<HMENU>(IDC_CFG_GET_KEY_LINK), nullptr, nullptr);
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
				120, 146, 240, 22, hWnd, reinterpret_cast<HMENU>(IDC_CFG_GET_KEY_LINK), nullptr, nullptr);
			hGetKeyLink = ctx->hGetKeyLink;
		}

		HWND hFillRightCodes = CreateWindowW(L"BUTTON", L"\u4E00\u952E\u586B\u5165 right.codes",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			430, 142, 190, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_FILL_RIGHT_CODES), nullptr, nullptr);

		HWND hTavilyApiKeyLabel = CreateWindowA("STATIC", "Tavily API Key:", WS_CHILD | WS_VISIBLE,
			16, 182, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hTavilyApiKey = CreateWindowExA(0, "EDIT", ctx->settings->tavilyApiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 180, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_TAVILY_API_KEY), nullptr, nullptr);

		HWND hExtraPromptLabel = CreateWindowA("STATIC", "System Prompt:", WS_CHILD | WS_VISIBLE,
			16, 216, 140, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hExtraPrompt = CreateWindowExA(0, "EDIT", ctx->settings->extraSystemPrompt.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			120, 216, 500, 150, hWnd, reinterpret_cast<HMENU>(IDC_CFG_EXTRA_PROMPT), nullptr, nullptr);

		HWND hSave = CreateWindowW(L"BUTTON", L"\u4FDD\u5B58\u5E76\u7EE7\u7EED", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			420, 382, 100, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_SAVE), nullptr, nullptr);
		HWND hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			530, 382, 90, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_CANCEL), nullptr, nullptr);

		std::array<HWND, 16> controls = {
			hProtocolLabel,
			ctx->hProtocol,
			hBaseUrlLabel,
			hApiKeyLabel,
			hModelLabel,
			hTavilyApiKeyLabel,
			ctx->hGetKeyLink,
			hFillRightCodes,
			hExtraPromptLabel,
			ctx->hBaseUrl,
			ctx->hApiKey,
			ctx->hModel,
			ctx->hTavilyApiKey,
			ctx->hExtraPrompt,
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

		if (id == IDC_CFG_SAVE) {
			AISettings next = *ctx->settings;
			next.protocolType = GetSelectedProtocol(ctx->hProtocol);
			next.baseUrl = GetEditTextA(ctx->hBaseUrl);
			next.apiKey = GetEditTextA(ctx->hApiKey);
			next.model = GetEditTextA(ctx->hModel);
			next.tavilyApiKey = GetEditTextA(ctx->hTavilyApiKey);
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

		if (id == IDC_CFG_FILL_RIGHT_CODES) {
			PopulateProtocolCombo(ctx->hProtocol, AIProtocolType::OpenAI);
			SetWindowTextA(ctx->hBaseUrl, "https://right.codes/codex");
			SetWindowTextA(ctx->hModel, "gpt-5.2-medium");
			SetFocus(ctx->hApiKey);
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

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;
	default:
		break;
	}

	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

bool TryApplyAISettingsFromWebPayload(HWND hWnd, AIConfigWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr || ctx->settings == nullptr) {
		return false;
	}

	AISettings next = *ctx->settings;
	next.protocolType = AIService::ParseProtocolType(data.value("protocol_type", AIService::ProtocolTypeToString(next.protocolType)));
	next.baseUrl = Utf8ToLocalText(data.value("base_url", ""));
	next.apiKey = Utf8ToLocalText(data.value("api_key", ""));
	next.model = Utf8ToLocalText(data.value("model", ""));
	next.extraSystemPrompt = Utf8ToLocalText(data.value("extra_system_prompt", ""));
	next.tavilyApiKey = Utf8ToLocalText(data.value("tavily_api_key", ""));

	if (AIService::Trim(next.baseUrl).empty() || AIService::Trim(next.apiKey).empty() || AIService::Trim(next.model).empty()) {
		MessageBoxA(hWnd, "baseUrl / apiKey / model cannot be empty.", "AI Config", MB_ICONWARNING | MB_OK);
		return false;
	}

	*ctx->settings = next;
	ctx->accepted = true;
	DestroyWindow(hWnd);
	return true;
}

void StartAIConfigWebView(HWND hWnd, AIConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHost == nullptr) {
		return;
	}

	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		nullptr,
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
					hWnd,
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
											else if (action == "cancel") {
												DestroyWindow(hWnd);
											}
											else if (action == "open_right_codes") {
												ShellExecuteA(hWnd, "open", "https://right.codes/register?aff=3dc87885", nullptr, nullptr, SW_SHOWNORMAL);
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
											if (navCtx->hLoading != nullptr) {
												ShowWindow(navCtx->hLoading, SW_HIDE);
											}
											OutputStringToELog("[AI Config][WebView2] navigation completed successfully");
											LayoutAIConfigWebViewDialog(hWnd, navCtx);
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
							const std::wstring html = WideFromLocal(BuildAIConfigWebViewHtml(*readyCtx->settings));
							if (html.empty()) {
								OutputStringToELog("[AI Config][WebView2] html conversion failed");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}
							OutputStringToELog(std::format("[AI Config][WebView2] controller ready=1 htmlChars={}", html.size()));
							readyCtx->webView->NavigateToString(html.c_str());
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
			WS_EX_CLIENTEDGE,
			"STATIC",
			"",
			WS_CHILD,
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
		ShowWindow(ctx->hHost, SW_HIDE);
		LayoutAIConfigWebViewDialog(hWnd, ctx);
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

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
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

bool ShowAIConfigDialogNative(HWND owner, AISettings& ioSettings)
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
	ctx.settings = &ioSettings;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 650, 464,
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

AIConfigDialogRunResult ShowAIConfigDialogWebView(HWND owner, AISettings& ioSettings)
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
	ctx.settings = &ioSettings;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 1060, 760,
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

bool ShowAIConfigDialog(HWND owner, AISettings& ioSettings)
{
	const AIConfigDialogRunResult webViewResult = ShowAIConfigDialogWebView(owner, ioSettings);
	if (webViewResult.fallbackRequested) {
		OutputStringToELog("[AI Config][WebView2] fallback to native dialog");
		return ShowAIConfigDialogNative(owner, ioSettings);
	}
	return webViewResult.accepted;
}

AIPreviewAction ShowAIPreviewDialogEx(
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
