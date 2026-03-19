#include "AutoLinker.h"
#include "AutoLinkerInternal.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <process.h>

#include <filesystem>
#include <format>
#include <string>

#include "AIChatFeature.h"
#include "AIConfigDialog.h"
#include "ECOMEx.h"
#include "Global.h"
#include "IDEFacade.h"
#include "LocalMcpServer.h"
#include "MouseBack.h"
#include "PathHelper.h"
#include "Version.h"
#include "WinINetUtil.h"
#include "WindowHelper.h"

#pragma comment(lib, "comctl32.lib")

bool FneInit();

LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	if (uMsg == WM_INITMENUPOPUP) {
		LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		FinalizeAutoLinkerPopupMenu(reinterpret_cast<HMENU>(wParam));
		return result;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

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
	if (uMsg == WM_AUTOLINKER_RUN_MODULE_INFO_TEST) {
		RunFirstImportedModulePublicInfoTest();
		return 0;
	}
	if (uMsg == WM_NCDESTROY) {
		LocalMcpServer::Shutdown();
		AIChatFeature::Shutdown();
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

	if (uMsg == WM_AUTOLINKER_INIT) {
		OutputStringToELog("收到初始化消息，尝试初始化");
		if (FneInit()) {
			OutputStringToELog("初始化成功");
			return 1;
		}
		return 0;
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

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2)
{
	return IDEFacade::Instance().RunFunctionRaw(code, p1, p2);
}

INT WINAPI fnAddInFunc(INT nAddInFnIndex)
{
	switch (nAddInFnIndex) {
	case 0: {
		std::string cmd = std::format("/select,{}", g_nowOpenSourceFilePath);
		ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
		break;
	}
	case 1: {
		std::string cmd = std::format("{}\\AutoLinker", GetBasePath());
		ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
		break;
	}
	case 2: {
		std::string cmd = std::format("{}", GetBasePath());
		ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
		break;
	}
	case 3:
		TryCopyCurrentFunctionCode();
		break;
	case 4: {
		AISettings settings = {};
		AIService::LoadSettings(g_configManager, settings);
		if (!ShowAIConfigDialog(g_hwnd, settings)) {
			OutputStringToELog("AI配置已取消");
			break;
		}
		AIService::SaveSettings(g_configManager, settings);
		OutputStringToELog("AI配置已保存");
		break;
	}
	case 5:
		RunFnAddTabStructPassThroughTest();
		break;
	case 6:
		RunDirectGlobalSearchKeywordTest();
		break;
	case 7:
		RunDirectGlobalSearchLocateKeywordTest();
		break;
	case 8:
		RunDirectGlobalSearchLocateAndDumpCurrentPageTest();
		break;
	case 9:
		RunTreeViewProbeTest();
		break;
	case 10:
		RunProgramTreeDirectPageDumpTest();
		break;
	case 11:
		RunProgramTreeListTest();
		break;
	case 12:
		RunCurrentPageWindowProbeTest();
		break;
	case 13:
		RunCurrentPageNameTest();
		break;
	case 14:
		RunImportedModuleListTest();
		break;
	case 15:
		RunFirstImportedModulePublicInfoTest();
		break;
	case 16:
		RunFirstSupportLibraryInfoTest();
		break;
	case 17:
		RunStaticCompileWindowsExeTest();
		break;
	default:
		break;
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
	OutputStringToELog("开始初始化");

	if (g_hwnd == NULL) {
		OutputStringToELog("g_hwnd 为空");
		return false;
	}

	g_toolBarHwnd = FindMenuBar(g_hwnd);

	DWORD processID = GetCurrentProcessId();
	std::string s = std::format("E进程ID{} 主句柄{} 菜单栏句柄{}", processID, (int)g_hwnd, (int)g_toolBarHwnd);
	OutputStringToELog(s);

	if (g_toolBarHwnd != NULL) {
		StartEditViewSubclassTask();
		RebuildTopLinkerSubMenu();
		OutputCurrentSourceLinker();
		AIChatFeature::EnsureTabCreated();

		OutputStringToELog("找到工具条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		StartHookCreateFileA();
		PostAppMessageA(g_toolBarHwnd, WM_PRINT, 0, 0);
		OutputStringToELog("初始化完成");

		const auto autoRunMarker = GetAutoRunModulePublicInfoTestMarkerPath();
		if (std::filesystem::exists(autoRunMarker)) {
			std::error_code removeEc;
			std::filesystem::remove(autoRunMarker, removeEc);
			OutputStringToELog("[ModulePublicInfoTest] 检测到自动测试标记，稍后执行首个导入模块公开信息测试");
			_beginthread(AutoRunModulePublicInfoTestThread, 0, nullptr);
		}

		_beginthread(FneCheckNewVersion, 0, NULL);
		return true;
	}

	OutputStringToELog(std::format("初始化失败，未找到工具条窗口 {}", (int)g_toolBarHwnd));
	return false;
}

void InitRetryThread(void* pParams)
{
	const int MAX_RETRY_COUNT = 5;
	int retryCount = 0;

	while (retryCount < MAX_RETRY_COUNT) {
		Sleep(1000);
		++retryCount;

		OutputStringToELog(std::format("初始化重试第 {}/{} 次", retryCount, MAX_RETRY_COUNT));
		LRESULT result = SendMessage(g_hwnd, WM_AUTOLINKER_INIT, 0, 0);
		if (result == 1) {
			OutputStringToELog("初始化线程完成");
			return;
		}
	}

	OutputStringToELog(std::format("初始化失败，已重试 {} 次", MAX_RETRY_COUNT));
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
		if (dwParam1 && !g_initStarted) {
			g_hwnd = GetMainWindowByProcessId();
			if (g_hwnd) {
				g_initStarted = true;
				SetWindowSubclass(g_hwnd, MainWindowSubclassProc, 0, 0);
				AIChatFeature::Initialize(g_hwnd, &g_configManager);
				LocalMcpServer::Initialize();
				RegisterIDEContextMenu();
				OutputStringToELog("主窗口子类化完成，启动初始化线程");
				_beginthread(InitRetryThread, 0, NULL);
			}
			else {
				OutputStringToELog("无法获取主窗口句柄");
			}
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
	_T("打开项目目录\0这是个用作测试的辅助工具功能。\0打开AutoLinker配置目录\0这是个用作测试的辅助工具功能。\0打开E语言目录\0这是个用作测试的辅助工具功能。\0复制当前函数代码\0复制当前光标所在子程序完整代码到剪贴板。\0AutoLinker AI接口设置\0编辑AI接口地址、API Key、模型和提示词等配置。\0FN_ADD_TAB结构传递测试\0构造ADD_TAB_INF调用FN_ADD_TAB，并打印调用前后结构体字段。\0测试整体搜索subWinHwnd\0调用direct_global_search固定搜索subWinHwnd，并输出命中结果到E输出窗口。\0测试定位subWinHwnd首个结果\0调用direct_global_search固定搜索subWinHwnd，并跳转到首个命中位置。\0测试定位后抓取当前页代码\0先定位到subWinHwnd首个命中，再抓取当前代码页完整代码并写入AutoLinker目录。\0测试枚举左侧TreeView\0枚举主窗口下所有SysTreeView32，并输出前几层节点文本与item data特征。\0测试程序树按名称抓代码\0在程序树中固定查找Class_HWND，并根据tree item data直接抓取整页代码。\0测试枚举程序树页面\0枚举程序树中所有页面节点，输出名称、类型和item data，并写入文件。\0测试当前页窗口与页签\0探测MDIClient当前活动子页与CCustomTabCtrl当前选中项文本，用于定位当前页名称来源。\0测试获取当前页名称\0调用IDEFacade当前页名称接口，输出当前页名称、类型和来源链路。\0测试枚举导入模块\0枚举当前项目导入的易模块路径。\0测试首个模块公开信息\0对当前项目首个导入模块执行原生公开信息抓取，并输出摘要与日志文件路径。\0测试首个支持库公开信息\0枚举当前已选支持库，并对首个可定位文件的支持库执行GetNewInf公开信息抓取，输出摘要与日志文件路径。\0测试编译静态EXE\0调用静态编译窗口程序EXE测试，并将输出文件写入AutoLinker\\StaticCompileTest目录。\0\0"),
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
	return (&LibInfo);
}
#endif
