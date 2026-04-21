#include "AutoLinker.h"
#include "AutoLinkerInternal.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <process.h>

#include <array>
#include <filesystem>
#include <fstream>
#include <format>
#include <mutex>
#include <string>

#include "AIChatTooling.h"
#include "AIChatFeature.h"
#include "AIConfigDialog.h"
#include "ECOMEx.h"
#include "Global.h"
#include "IDEFacade.h"
#include "LocalMcpServer.h"
#include "Logger.h"
#include "MemFind.h"
#include "MouseBack.h"
#include "PathHelper.h"
#include "ProjectSourceCacheManager.h"
#include "Version.h"
#include "WinINetUtil.h"
#include "WindowHelper.h"

#pragma comment(lib, "comctl32.lib")

bool FneInit();

namespace {

bool g_mainWindowSubclassInstalled = false;
bool g_initTraceSessionStarted = false;
std::mutex g_initTraceMutex;

constexpr const char* kDebugStartPattern =
	"55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC ?? ?? 00 00 56 89 8D ?? FB FF FF 8B 8D ?? FB FF FF 81 C1 C0 00 00 00 E8 ?? ?? ?? ?? 85 C0 74 05 E9 ?? 08 00 00";
constexpr const char* kCompileStartPattern =
	"55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 81 EC C4 00";

std::filesystem::path GetInitTraceLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "startup_init_last.log";
}

void TraceInitStep(const std::string& step)
{
	std::lock_guard<std::mutex> lock(g_initTraceMutex);
	const auto path = GetInitTraceLogPath();
	const auto mode = g_initTraceSessionStarted ? (std::ios::app | std::ios::binary) : (std::ios::trunc | std::ios::binary);
	std::ofstream out(path, mode);
	if (!out.is_open()) {
		return;
	}
	g_initTraceSessionStarted = true;
	out << "[" << Logger::BuildTimestamp() << "] " << step << "\r\n";
}

void ResolveCompileDebugStartAddressesForInit()
{
	if (g_debugStartAddress > 0 && g_compileStartAddress > 0) {
		TraceInitStep(std::format(
			"编译/调试入口地址已就绪 debug=0x{:X} compile=0x{:X}",
			g_debugStartAddress,
			g_compileStartAddress));
		return;
	}

	TraceInitStep("开始解析编译/调试入口地址");
	if (g_debugStartAddress <= 0) {
		g_debugStartAddress = FindSelfModelMemory(kDebugStartPattern);
	}
	if (g_compileStartAddress <= 0) {
		g_compileStartAddress = FindSelfModelMemory(kCompileStartPattern);
	}

	if (g_debugStartAddress > 0 && g_compileStartAddress > 0) {
		TraceInitStep(std::format(
			"编译/调试入口地址解析完成 debug=0x{:X} compile=0x{:X}",
			g_debugStartAddress,
			g_compileStartAddress));
		return;
	}

	TraceInitStep(std::format(
		"编译/调试入口地址未完整解析 debug=0x{:X} compile=0x{:X}，将跳过对应 Hook",
		g_debugStartAddress,
		g_compileStartAddress));
}

struct AddInMenuEntry {
	const char* title;
	const char* description;
	void (*handler)();
};

void OpenProjectDirectoryAddIn()
{
	std::string cmd = std::format("/select,{}", g_nowOpenSourceFilePath);
	ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
}

void OpenAutoLinkerConfigDirectoryAddIn()
{
	std::string cmd = std::format("{}\\AutoLinker", GetBasePath());
	ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
}

void OpenELanguageDirectoryAddIn()
{
	std::string cmd = std::format("{}", GetBasePath());
	ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
}

void ShowAISettingsAddIn()
{
	AISettings settings = {};
	AIService::LoadSettings(g_aiJsonConfig, &g_configManager, settings);
	if (!ShowAIConfigDialog(g_hwnd, settings)) {
		OutputStringToELog("AI配置已取消");
		return;
	}
	AIService::SaveSettings(g_aiJsonConfig, settings);
	OutputStringToELog("AI配置已保存");
}

const auto& GetAddInMenuEntries()
{
	static const std::array<AddInMenuEntry, 4> kEntries = { {
		{ "打开项目目录", "这是个用作测试的辅助工具功能。", &OpenProjectDirectoryAddIn },
		{ "打开AutoLinker配置目录", "这是个用作测试的辅助工具功能。", &OpenAutoLinkerConfigDirectoryAddIn },
		{ "打开E语言目录", "这是个用作测试的辅助工具功能。", &OpenELanguageDirectoryAddIn },
		{ "AutoLinker AI接口设置", "编辑AI接口地址、API Key、模型和提示词等配置。", &ShowAISettingsAddIn },
	} };
	return kEntries;
}

const std::string& GetAddInMenuInfoText()
{
	static const std::string kMenuInfo = [] {
		std::string text;
		for (const auto& entry : GetAddInMenuEntries()) {
			text.append(entry.title);
			text.push_back('\0');
			text.append(entry.description);
			text.push_back('\0');
		}
		text.push_back('\0');
		return text;
	}();
	return kMenuInfo;
}

void RefreshSourcePathAfterWindowStateChange(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return;
	}

	if (uMsg == WM_SETTEXT || uMsg == WM_MDIACTIVATE) {
		(void)wParam;
		(void)lParam;
		UpdateCurrentOpenSourceFile();
	}
}

} // namespace

void ChangeVMProtectModel(bool isLib)
{
	if (isLib) {
		int sdk = FindECOMNameIndex("VMPSDK");
		if (sdk != -1) {
			RemoveECOM(sdk);
		}
		int sdkLib = FindECOMNameIndex("VMPSDK_LIB");
		if (sdkLib == -1) {
			OutputStringToELog("切换到静态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK_LIB.ec", buffer);
			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	else {
		int sdk = FindECOMNameIndex("VMPSDK_LIB");
		if (sdk != -1) {
			RemoveECOM(sdk);
		}
		int sdkLib = FindECOMNameIndex("VMPSDK");
		if (sdkLib == -1) {
			OutputStringToELog("切换到动态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK.ec", buffer);
			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
}

LRESULT CALLBACK MainWindowSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	if (uMsg == WM_AUTOLINKER_AI_TASK_DONE) {
		HandleAiTaskCompletionMessage(lParam);
		return 0;
	}
	if (uMsg == WM_AUTOLINKER_AI_APPLY_RESULT) {
		HandleAiApplyMessage(lParam);
		return 0;
	}
	if (uMsg == WM_NCDESTROY) {
		LocalMcpServer::Shutdown();
		AIChatFeature::Shutdown();
		g_mainWindowSubclassInstalled = false;
		RemoveWindowSubclass(hWnd, MainWindowSubclassProc, uIdSubclass);
		return DefSubclassProc(hWnd, uMsg, wParam, lParam);
	}
	if (AIChatFeature::HandleMainWindowMessage(uMsg, wParam, lParam)) {
		return 0;
	}

	if (uMsg == WM_COMMAND) {
		UINT cmd = LOWORD(wParam);
		if (HandleTopLinkerMenuCommand(cmd)) {
			return 0;
		}
		if (IDEFacade::Instance().HandleMainWindowCommand(wParam)) {
			return 0;
		}
	}

	if (uMsg == 20708) {
		BOOL result = SetWindowSubclass((HWND)wParam, EditViewSubclassProc, 0, 0);
		return result ? 1 : 0;
	}

	if (uMsg == WM_INITMENUPOPUP) {
		LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		FinalizeAutoLinkerPopupMenu(reinterpret_cast<HMENU>(wParam));
		return result;
	}

	if (uMsg == WM_SETTEXT || uMsg == WM_MDIACTIVATE) {
		LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		RefreshSourcePathAfterWindowStateChange(hWnd, uMsg, wParam, lParam);
		return result;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2)
{
	return IDEFacade::Instance().RunFunctionRaw(code, p1, p2);
}

INT WINAPI fnAddInFunc(INT nAddInFnIndex)
{
	const auto& entries = GetAddInMenuEntries();
	if (nAddInFnIndex < 0 || static_cast<size_t>(nAddInFnIndex) >= entries.size()) {
		return 0;
	}

	const auto handler = entries[static_cast<size_t>(nAddInFnIndex)].handler;
	if (handler != nullptr) {
		handler();
	}

	return 0;
}

void FneCheckNewVersion(void* pParams)
{
	Sleep(1000);

	OutputStringToELog("AutoLinker开源下载地址：https://github.com/aiqinxuancai/AutoLinker");
	std::string url = "https://api.github.com/repos/aiqinxuancai/AutoLinker/releases";
	auto response = PerformGetRequest(url);

	std::string currentVersion = AUTOLINKER_VERSION;
	if (response.second == 200) {
		std::string nowGithubVersion = "0.0.0";
		if (strcmp(AUTOLINKER_VERSION, "0.0.0") == 0) {
			OutputStringToELog(std::format("自编译版本，不检查更新，当前版本：{}", currentVersion));
		}
		else if (!response.first.empty()) {
			try {
				auto releases = json::parse(response.first);
				for (const auto& release : releases) {
					if (!release["prerelease"].get<bool>()) {
						nowGithubVersion = release["tag_name"];
						break;
					}
				}

				Version nowGithubVersionObj(nowGithubVersion);
				Version currentVersionObj(AUTOLINKER_VERSION);
				if (nowGithubVersionObj > currentVersionObj) {
					OutputStringToELog(std::format("有新版本：{}", nowGithubVersion));
				}
			}
			catch (const std::exception& e) {
				OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误：{}", currentVersion, e.what()));
			}
		}
	}
}

bool FneInit()
{
	if (g_uiInitialized) {
		TraceInitStep("FneInit 已初始化，跳过重复执行");
		return true;
	}

	TraceInitStep("进入 FneInit");
	OutputStringToELog("开始初始化");

	if (!g_notifySysReady) {
		TraceInitStep("系统通知接口尚未就绪，初始化终止");
		OutputStringToELog("系统通知接口尚未就绪");
		return false;
	}

	if (g_hwnd == NULL || !IsWindow(g_hwnd)) {
		TraceInitStep("开始获取主窗口句柄");
		g_hwnd = reinterpret_cast<HWND>(NotifySys(NES_GET_MAIN_HWND, 0, 0));
		if (g_hwnd == nullptr || !IsWindow(g_hwnd)) {
			TraceInitStep("NotifySys 获取主窗口句柄失败，改用进程枚举");
			g_hwnd = GetMainWindowByProcessId();
		}
	}

	if (g_hwnd == NULL || !IsWindow(g_hwnd)) {
		TraceInitStep("主窗口句柄无效，初始化终止");
		OutputStringToELog("主窗口句柄无效");
		return false;
	}
	TraceInitStep(std::format(
		"主窗口句柄有效 hwnd={}",
		static_cast<unsigned long long>(reinterpret_cast<ULONG_PTR>(g_hwnd))));

	if (!g_mainWindowSubclassInstalled) {
		TraceInitStep("开始安装主窗口子类");
		if (SetWindowSubclass(g_hwnd, MainWindowSubclassProc, 0, 0) == FALSE) {
			TraceInitStep("主窗口子类化失败，初始化终止");
			OutputStringToELog("主窗口子类化失败");
			return false;
		}
		g_mainWindowSubclassInstalled = true;
		TraceInitStep("主窗口子类化完成");
	}

	DWORD processID = GetCurrentProcessId();
	TraceInitStep(std::format(
		"记录进程信息 pid={} hwnd={}",
		processID,
		static_cast<unsigned long long>(reinterpret_cast<ULONG_PTR>(g_hwnd))));
	OutputStringToELog(std::format("E进程ID{} 主句柄{}", processID, (int)g_hwnd));

	TraceInitStep("开始注册 IDE 右键菜单");
	RegisterIDEContextMenu();
	TraceInitStep("IDE 右键菜单注册完成");
	TraceInitStep("开始启动编辑框子类化巡检线程");
	StartEditViewSubclassTask();
	TraceInitStep("编辑框子类化巡检线程已启动");
	TraceInitStep("开始输出当前源码链接器信息");
	OutputCurrentSourceLinker();
	TraceInitStep("当前源码链接器信息输出完成");
	TraceInitStep("开始初始化 AI Chat");
	AIChatFeature::Initialize(g_hwnd, &g_configManager, &g_aiJsonConfig);
	TraceInitStep("AI Chat 初始化完成");
	TraceInitStep("开始初始化 Local MCP");
	LocalMcpServer::Initialize();
	TraceInitStep("Local MCP 初始化完成");
	ResolveCompileDebugStartAddressesForInit();
	TraceInitStep("开始安装文件与编译相关 Hook");
	StartHookCreateFileA();
	TraceInitStep("文件与编译相关 Hook 安装完成");
	TraceInitStep("开始预热当前工程源码解析缓存");
	{
		std::string warmupError;
		std::string warmupTrace;
		if (!project_source_cache::ProjectSourceCacheManager::Instance().WarmupCurrentSource(
				&warmupError,
				&warmupTrace)) {
			OutputStringToELog(std::format(
				"[ProjectSourceCacheWarmup] 跳过或失败 error={} trace={}",
				warmupError.empty() ? "warmup_failed" : warmupError,
				warmupTrace));
		}
	}
	TraceInitStep("当前工程源码解析缓存预热结束");

	TraceInitStep("开始启动版本检查线程");
	_beginthread(FneCheckNewVersion, 0, NULL);
	TraceInitStep("版本检查线程已启动");
	g_uiInitialized = true;
	TraceInitStep("FneInit 完成");
	OutputStringToELog("初始化完成");
	return true;
}

EXTERN_C INT WINAPI AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	std::string s = std::format("AutoLinker_MessageNotify {0} {1} {2}", (int)nMsg, dwParam1, dwParam2);
	OutputStringToELog(s);

#ifndef __E_STATIC_LIB
	if (nMsg == NL_GET_CMD_FUNC_NAMES) {
		return NULL;
	}
	else if (nMsg == NL_GET_NOTIFY_LIB_FUNC_NAME) {
		return (INT)LIBARAYNAME;
	}
	else if (nMsg == NL_GET_DEPENDENT_LIBS) {
		return (INT)NULL;
	}
	else if (nMsg == NL_SYS_NOTIFY_FUNCTION) {
		if (dwParam1 && !g_notifySysReady) {
			g_notifySysReady = true;
			TraceInitStep("收到 NL_SYS_NOTIFY_FUNCTION，系统通知接口已就绪");
			OutputStringToELog("系统通知接口已就绪，等待 IDE 就绪通知");
		}
	}
	else if (nMsg == NL_IDE_READY) {
		TraceInitStep("收到 NL_IDE_READY");
		if (!g_notifySysReady) {
			TraceInitStep("系统通知接口尚未就绪，无法执行 IDE 就绪初始化");
			OutputStringToELog("收到 NL_IDE_READY，但系统通知接口尚未就绪");
		}
		else if (FneInit()) {
			TraceInitStep("NL_IDE_READY 初始化成功");
			OutputStringToELog("收到 NL_IDE_READY，界面初始化成功");
		}
		else {
			TraceInitStep("NL_IDE_READY 初始化失败");
			OutputStringToELog("收到 NL_IDE_READY，但界面初始化失败");
		}
	}
#endif
	return ProcessNotifyLib(nMsg, dwParam1, dwParam2);
}

#ifndef __E_STATIC_LIB
static LIB_INFOX LibInfo =
{
	LIB_FORMAT_VER,
	_T(LIB_GUID_STR),
	LIB_MajorVersion,
	LIB_MinorVersion,
	LIB_BuildNumber,
	LIB_SysMajorVer,
	LIB_SysMinorVer,
	LIB_KrnlLibMajorVer,
	LIB_KrnlLibMinorVer,
	_T(LIB_NAME_STR),
	__GBK_LANG_VER,
	_WT(LIB_DESCRIPTION_STR),
	LBS_IDE_PLUGIN | LBS_LIB_INFO2,
	_WT(LIB_Author),
	_WT(LIB_ZipCode),
	_WT(LIB_Address),
	_WT(LIB_Phone),
	_WT(LIB_Fax),
	_WT(LIB_Email),
	_WT(LIB_HomePage),
	_WT(LIB_Other),
	0,
	NULL,
	LIB_TYPE_COUNT,
	_WT(LIB_TYPE_STR),
	0,
	NULL,
	NULL,
	fnAddInFunc,
	NULL,
	AutoLinker_MessageNotify,
	NULL,
	NULL,
	0,
	NULL,
	NULL,
	NULL,
	0,
	NULL,
	"You",
};

PLIB_INFOX WINAPI GetNewInf()
{
	LibInfo.m_szzAddInFnInfo = GetAddInMenuInfoText().c_str();
	return (&LibInfo);
}
#endif
