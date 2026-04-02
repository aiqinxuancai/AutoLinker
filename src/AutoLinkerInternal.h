#pragma once



#include <Windows.h>



#include <chrono>

#include <filesystem>

#include <string>



#include "ConfigManager.h"

#include "LinkerManager.h"

#include "ModelManager.h"

#include "AIService.h"



extern std::string g_nowOpenSourceFilePath;

extern ConfigManager g_configManager;

extern LinkerManager g_linkerManager;

extern ModelManager g_modelManager;

extern HWND g_hwnd;

extern bool g_preDebugging;

extern bool g_preCompiling;

// 已完成基础挂接。
extern bool g_initStarted;

// 已完成 IDE 就绪后的界面初始化。
extern bool g_uiInitialized;



inline constexpr UINT IDM_AUTOLINKER_CTX_COPY_FUNC = 31001;

inline constexpr UINT IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC = 31101;

inline constexpr UINT IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC = 31102;

inline constexpr UINT IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC = 31103;

inline constexpr UINT IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT = 31104;

inline constexpr UINT IDM_AUTOLINKER_CTX_AI_ADD_BY_PAGE = 31106;

inline constexpr UINT IDM_AUTOLINKER_LINKER_BASE = 34000;

inline constexpr UINT IDM_AUTOLINKER_LINKER_MAX = 34999;

inline constexpr UINT WM_AUTOLINKER_AI_TASK_DONE = WM_USER + 1001;

inline constexpr UINT WM_AUTOLINKER_AI_APPLY_RESULT = WM_USER + 1002;

inline constexpr UINT WM_AUTOLINKER_RUN_MODULE_INFO_TEST = WM_USER + 1003;



using PerfClock = std::chrono::steady_clock;



std::string TrimAsciiCopy(const std::string& text);

std::string ToLowerAsciiCopy(const std::string& text);

std::string TruncateForPerfLog(const std::string& text, size_t maxLen = 120);

long long ElapsedMs(const PerfClock::time_point& start);
std::string EscapeOneLineForLog(std::string text);



void UpdateCurrentOpenSourceFile();
void OutputCurrentSourceLinker();

void RebuildTopLinkerSubMenu();

bool HandleTopLinkerMenuCommand(UINT cmd);

void PrepareAutoLinkerPopupMenu(HMENU hMenu);

void FinalizeAutoLinkerPopupMenu(HMENU hMenu);

void RegisterIDEContextMenu();



void BeginAIRoundtripLogSession(uint64_t traceId, const std::string& scene, const std::string& taskName);

void AppendAIRoundtripLogLine(const std::string& line);

void AppendAIRoundtripLogBlock(const std::string& title, const std::string& text);



void HandleAiTaskCompletionMessage(LPARAM lParam);

void HandleAiApplyMessage(LPARAM lParam);

void RunAiFunctionReplaceTask(AITaskKind kind);

void RunAiTranslateSelectedTextTask();

void RunAiAddByCurrentPageTypeTask();

void TryCopyCurrentFunctionCode();



void RunFnAddTabStructPassThroughTest();

void RunDirectGlobalSearchKeywordTest();

void RunDirectGlobalSearchLocateKeywordTest();

void RunDirectGlobalSearchLocateAndDumpCurrentPageTest();

void RunTreeViewProbeTest();

void RunProgramTreeDirectPageDumpTest();

// 测试切换到常量表页面。
void RunProgramTreeSwitchToConstantTableTest();

// 测试抓取常量表页面真实代码。
void RunProgramTreeReadConstantTableCodeTest();

void RunProgramTreeListTest();

void RunCurrentPageWindowProbeTest();

void RunCurrentPageNameTest();

void RunImportedModuleListTest();

void RunFirstImportedModulePublicInfoTest();

void RunFirstSupportLibraryInfoTest();

void RunStaticCompileWindowsExeTest();



std::filesystem::path GetAutoRunModulePublicInfoTestMarkerPath();

void AutoRunModulePublicInfoTestThread(void* pParams);

