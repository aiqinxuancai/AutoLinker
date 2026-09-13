"""使用真实 Win32 消息循环验证生产代码的网页刷新调度，无需启动 IDE。"""
import os
from pathlib import Path
import subprocess
import sys
import tempfile


root = Path(__file__).resolve().parents[1]
source = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "src/AIChatFeature.cpp"
text = source.read_text(encoding="utf-8-sig")
start = text.index("void UpdateHistoryWebViewHtml(")
end = text.index("void TryInitializeHistoryWebView(", start)
functions = text[start:end]
harness = r'''
#include <windows.h>
#include <string>
#include <iostream>
#include <stdexcept>
struct ChatDialogContext {
    bool webViewReady = true, webViewContentReady = true, webViewFlushScheduled = false;
    void* webView = this;
    std::string pendingHistoryHtml, pendingHistorySessionId;
};
HWND g_chatDialog;
constexpr UINT_PTR kHistoryWebViewFlushTimerId = 0xA17;
ChatDialogContext context;
int renders = 0;
std::wstring lastScript;
std::wstring WideFromLocal(const std::string& s) { return {s.begin(), s.end()}; }
std::wstring WideFromUtf8Text(const std::string& s) { return WideFromLocal(s); }
std::wstring EscapeJsSingleQuotedWide(const std::wstring& s) { return s; }
void ExecuteWebViewScript(ChatDialogContext*, const std::wstring& script, bool = false) {
    ++renders;
    lastScript = script;
}
void FlushHistoryWebViewHtml(ChatDialogContext*);
/* PRODUCTION_FUNCTIONS */
LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_TIMER && wp == kHistoryWebViewFlushTimerId) {
        KillTimer(hwnd, wp);
        if (context.webViewFlushScheduled) FlushHistoryWebViewHtml(&context);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}
void pump(DWORD duration, bool streaming = false) {
    const auto start = GetTickCount64();
    ULONGLONG nextUpdate = 0;
    while (GetTickCount64() - start < duration) {
        auto elapsed = GetTickCount64() - start;
        if (streaming && elapsed >= nextUpdate) {
            UpdateHistoryWebViewHtml(&context, std::to_string(elapsed), "session-a");
            nextUpdate = elapsed + 10;
        }
        MSG msg;
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) DispatchMessageW(&msg);
        Sleep(1);
    }
}
void check(bool ok, const char* message) { if (!ok) throw std::runtime_error(message); }
int main() {
    WNDCLASSW wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"ChatRefreshRegression";
    RegisterClassW(&wc);
    g_chatDialog = CreateWindowW(wc.lpszClassName, L"", 0, 0, 0, 0, 0,
        HWND_MESSAGE, nullptr, wc.hInstance, nullptr);
    try {
        check(g_chatDialog != nullptr, "Cannot create test window");
        pump(700, true);
        check(renders >= 3, "Continuous updates starved history rendering");
        UpdateHistoryWebViewHtml(&context, "final", "session-b");
        pump(200);
        check(lastScript.find(L"final', 'session-b") != std::wstring::npos,
            "Latest content/session was lost");
        const auto settled = renders;
        pump(150);
        check(renders == settled && !context.webViewFlushScheduled, "Timer did not settle");
        UpdateHistoryWebViewHtml(&context, "before-navigation", "session-b");
        context.webViewContentReady = false;
        pump(200);
        check(!context.webViewFlushScheduled, "Unready page left scheduling latched");
        UpdateHistoryWebViewHtml(&context, "after-navigation", "session-b");
        context.webViewContentReady = true;
        UpdateHistoryWebViewHtml(&context, context.pendingHistoryHtml, context.pendingHistorySessionId);
        pump(200);
        check(lastScript.find(L"after-navigation") != std::wstring::npos, "Navigation recovery failed");
        UpdateHistoryWebViewHtml(&context, "", "session-c");
        pump(200);
        check(lastScript.find(L"('', 'session-c')") != std::wstring::npos, "Empty history not rendered");
        std::cout << "PASS: continuous output, latest snapshot, timer cleanup, navigation, empty history\n";
    } catch (const std::exception& error) {
        std::cerr << "FAIL: " << error.what() << '\n';
        DestroyWindow(g_chatDialog);
        return 1;
    }
    DestroyWindow(g_chatDialog);
}
'''.replace("/* PRODUCTION_FUNCTIONS */", functions)

vswhere = Path(os.environ["ProgramFiles(x86)"]) / "Microsoft Visual Studio/Installer/vswhere.exe"
installation = subprocess.check_output([
    str(vswhere), "-latest", "-requires", "Microsoft.VisualStudio.Component.VC.Tools.x86.x64",
    "-property", "installationPath"
], text=True).strip()
with tempfile.TemporaryDirectory(prefix="chat-refresh-") as directory:
    work = Path(directory)
    (work / "test.cpp").write_text(harness, encoding="utf-8-sig", newline="\r\n")
    devcmd = Path(installation) / "Common7/Tools/VsDevCmd.bat"
    command = f'call "{devcmd}" -arch=x86 -host_arch=x64 >nul && cl /nologo /EHsc /std:c++20 /utf-8 test.cpp user32.lib /Fe:test.exe'
    subprocess.run('cmd.exe /d /s /c "' + command + '"', cwd=work, check=True)
    subprocess.run([str(work / "test.exe")], cwd=work, check=True)
