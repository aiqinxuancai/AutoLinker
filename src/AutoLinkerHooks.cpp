#include "AutoLinkerInternal.h"

#include <Windows.h>
#include <CommDlg.h>

#include <algorithm>
#include <chrono>
#include <cstring>
#include <filesystem>
#include <format>
#include <initializer_list>
#include <mutex>
#include <string>

#include <detours.h>

#include "ECOMEx.h"
#include "EideProjectBinarySerializer.h"
#include "Global.h"
#include "HeadlessCompileRunner.h"
#include "IDEFacade.h"
#include "MemFind.h"
#include "PathHelper.h"
#include "StringHelper.h"

namespace {

struct SilentCompileOutputPathState {
	bool active = false;
	bool consumed = false;
	DWORD ownerThreadId = 0;
	std::string outputPath;
	std::chrono::steady_clock::time_point createdAt = {};
};

std::mutex g_silentCompileOutputPathMutex;
SilentCompileOutputPathState g_silentCompileOutputPathState;
std::mutex g_detourHookMutex;
bool g_messageBoxHookInstalled = false;
bool g_fullHookInstalled = false;

constexpr auto kSilentCompileOutputPathTimeout = std::chrono::seconds(30);

std::string NormalizeSilentCompileOutputPath(const std::string& outputPath)
{
	if (outputPath.empty()) {
		return std::string();
	}

	try {
		std::filesystem::path path(outputPath);
		path = path.lexically_normal();
		if (path.is_relative()) {
			path = std::filesystem::absolute(path);
		}
		path = path.lexically_normal();
		return path.string();
	}
	catch (...) {
		return outputPath;
	}
}

bool FillOpenFileNamePathA(LPOPENFILENAMEA item, const std::string& outputPath, std::string* diagnostics)
{
	if (item == nullptr) {
		if (diagnostics != nullptr) {
			*diagnostics = "item_null";
		}
		return false;
	}
	if (item->lpstrFile == nullptr || item->nMaxFile <= 1) {
		if (diagnostics != nullptr) {
			*diagnostics = "file_buffer_invalid";
		}
		return false;
	}
	if (outputPath.empty()) {
		if (diagnostics != nullptr) {
			*diagnostics = "output_path_empty";
		}
		return false;
	}
	if (outputPath.size() + 1 > static_cast<size_t>(item->nMaxFile)) {
		if (diagnostics != nullptr) {
			*diagnostics = std::format(
				"file_buffer_too_small pathLen={} maxFile={}",
				outputPath.size(),
				item->nMaxFile);
		}
		return false;
	}

	std::memset(item->lpstrFile, 0, static_cast<size_t>(item->nMaxFile));
	std::memcpy(item->lpstrFile, outputPath.c_str(), outputPath.size());

	try {
		const std::filesystem::path path(outputPath);
		const std::string fileName = path.filename().string();
		if (item->lpstrFileTitle != nullptr && item->nMaxFileTitle > 0) {
			if (fileName.size() + 1 > static_cast<size_t>(item->nMaxFileTitle)) {
				if (diagnostics != nullptr) {
					*diagnostics = std::format(
						"file_title_buffer_too_small titleLen={} maxFileTitle={}",
						fileName.size(),
						item->nMaxFileTitle);
				}
				return false;
			}
			std::memset(item->lpstrFileTitle, 0, static_cast<size_t>(item->nMaxFileTitle));
			std::memcpy(item->lpstrFileTitle, fileName.c_str(), fileName.size());
		}

		const std::string nativePath = path.string();
		const size_t fileOffset = nativePath.size() - fileName.size();
		item->nFileOffset = static_cast<WORD>(fileOffset);

		const std::string extension = path.extension().string();
		if (!extension.empty()) {
			item->nFileExtension = static_cast<WORD>(nativePath.size() - extension.size() + 1);
		}
		else {
			item->nFileExtension = 0;
		}
	}
	catch (...) {
		const char* lastSlash = (std::max)(strrchr(outputPath.c_str(), '\\'), strrchr(outputPath.c_str(), '/'));
		const char* fileName = lastSlash != nullptr ? (lastSlash + 1) : outputPath.c_str();
		if (item->lpstrFileTitle != nullptr && item->nMaxFileTitle > 0) {
			const size_t titleLen = std::strlen(fileName);
			if (titleLen + 1 <= static_cast<size_t>(item->nMaxFileTitle)) {
				std::memset(item->lpstrFileTitle, 0, static_cast<size_t>(item->nMaxFileTitle));
				std::memcpy(item->lpstrFileTitle, fileName, titleLen);
			}
		}
		item->nFileOffset = static_cast<WORD>(fileName - outputPath.c_str());
		const char* ext = strrchr(fileName, '.');
		item->nFileExtension = ext != nullptr ? static_cast<WORD>(ext - outputPath.c_str() + 1) : 0;
	}

	if (item->nFilterIndex == 0) {
		item->nFilterIndex = 1;
	}
	if (diagnostics != nullptr) {
		*diagnostics = "ok";
	}
	return true;
}

bool TryHandleSilentCompileOutputPathRequest(LPOPENFILENAMEA item, std::string* diagnostics)
{
	std::lock_guard<std::mutex> lock(g_silentCompileOutputPathMutex);
	if (!g_silentCompileOutputPathState.active) {
		if (diagnostics != nullptr) {
			*diagnostics = "request_inactive";
		}
		return false;
	}

	const auto now = std::chrono::steady_clock::now();
	if (now - g_silentCompileOutputPathState.createdAt > kSilentCompileOutputPathTimeout) {
		g_silentCompileOutputPathState = {};
		if (diagnostics != nullptr) {
			*diagnostics = "request_expired";
		}
		return false;
	}

	const DWORD currentThreadId = GetCurrentThreadId();
	if (g_silentCompileOutputPathState.ownerThreadId != 0 &&
		g_silentCompileOutputPathState.ownerThreadId != currentThreadId) {
		if (diagnostics != nullptr) {
			*diagnostics = std::format(
				"thread_mismatch owner={} current={}",
				g_silentCompileOutputPathState.ownerThreadId,
				currentThreadId);
		}
		return false;
	}

	std::string fillDiagnostics;
	if (!FillOpenFileNamePathA(item, g_silentCompileOutputPathState.outputPath, &fillDiagnostics)) {
		g_silentCompileOutputPathState = {};
		if (diagnostics != nullptr) {
			*diagnostics = "fill_failed:" + fillDiagnostics;
		}
		return false;
	}

	g_silentCompileOutputPathState.active = false;
	g_silentCompileOutputPathState.consumed = true;
	if (diagnostics != nullptr) {
		*diagnostics = "handled";
	}
	return true;
}

bool IsSilentCompileOutputPathRequestOngoingForCurrentThread()
{
	std::lock_guard<std::mutex> lock(g_silentCompileOutputPathMutex);
	if (!g_silentCompileOutputPathState.active && !g_silentCompileOutputPathState.consumed) {
		return false;
	}

	const DWORD currentThreadId = GetCurrentThreadId();
	if (g_silentCompileOutputPathState.ownerThreadId != 0 &&
		g_silentCompileOutputPathState.ownerThreadId != currentThreadId) {
		return false;
	}

	return true;
}

bool IsSilentCompileModulePublishPrompt(const std::string& text)
{
	return text.find("编译当前易模块程序到其发布版本成功") != std::string::npos &&
		text.find("系统模块库目录") != std::string::npos &&
		text.find("ecom") != std::string::npos;
}

} // namespace

bool BeginSilentCompileOutputPathRequest(
	const std::string& outputPath,
	DWORD ownerThreadId,
	std::string* diagnostics)
{
	const std::string normalized = NormalizeSilentCompileOutputPath(outputPath);
	if (normalized.empty()) {
		if (diagnostics != nullptr) {
			*diagnostics = "normalized_output_path_empty";
		}
		return false;
	}

	std::lock_guard<std::mutex> lock(g_silentCompileOutputPathMutex);
	g_silentCompileOutputPathState = {};
	g_silentCompileOutputPathState.active = true;
	g_silentCompileOutputPathState.consumed = false;
	g_silentCompileOutputPathState.ownerThreadId =
		ownerThreadId != 0 ? ownerThreadId : GetCurrentThreadId();
	g_silentCompileOutputPathState.outputPath = normalized;
	g_silentCompileOutputPathState.createdAt = std::chrono::steady_clock::now();
	if (diagnostics != nullptr) {
		*diagnostics = normalized;
	}
	return true;
}

void CancelSilentCompileOutputPathRequest()
{
	std::lock_guard<std::mutex> lock(g_silentCompileOutputPathMutex);
	g_silentCompileOutputPathState = {};
}

bool WasSilentCompileOutputPathRequestConsumed()
{
	std::lock_guard<std::mutex> lock(g_silentCompileOutputPathMutex);
	return g_silentCompileOutputPathState.consumed;
}

static auto originalCreateFileA = CreateFileA;
static auto originalGetSaveFileNameA = GetSaveFileNameA;
static auto originalCreateProcessA = CreateProcessA;
static auto originalMessageBoxA = MessageBoxA;
static auto originalMessageBoxW = MessageBoxW;
static auto originalTrackPopupMenu = TrackPopupMenu;
static auto originalTrackPopupMenuEx = TrackPopupMenuEx;

#if defined(_M_IX86)
constexpr std::uintptr_t kImageBase = 0x400000;
constexpr const char* kProjectSerializeToFilePattern =
	"55 8B EC 6A FF 68 ?? ?? ?? ?? 64 A1 00 00 00 00 50 64 89 25 00 00 00 00 "
	"81 EC ?? ?? ?? ?? 53 56 8B D9 57 B9 ?? ?? ?? ?? 89 65 ?? E8 ?? ?? ?? ?? "
	"8B F8 8B 83 ?? ?? ?? ?? 89 7D ?? 8B 48 F8 85 C9 75 ?? 6A 14 C7 45 ?? 00 00 00 00 "
	"E8 ?? ?? ?? ?? 8B F0";
typedef int(__thiscall* OriginalProjectSerializeToFileFuncType)(void* thisPtr, LPCSTR lpFileName, int a3);
OriginalProjectSerializeToFileFuncType originalProjectSerializeToFile = nullptr;

template <typename T>
T ResolveInternalAddress(std::uintptr_t moduleBase, std::uintptr_t rva)
{
	if (moduleBase == 0 || rva < kImageBase) {
		return nullptr;
	}
	return reinterpret_cast<T>(moduleBase + (rva - kImageBase));
}

std::uintptr_t NormalizeRuntimeAddress(std::uintptr_t runtimeAddress, std::uintptr_t moduleBase)
{
	if (runtimeAddress == 0 || moduleBase == 0 || runtimeAddress < moduleBase) {
		return 0;
	}
	return runtimeAddress - moduleBase + kImageBase;
}

std::uintptr_t ResolveUniqueCodeAddress(const char* pattern, std::uintptr_t moduleBase)
{
	if (pattern == nullptr || *pattern == '\0') {
		return 0;
	}

	const auto matches = FindSelfModelMemoryAll(pattern);
	if (matches.size() != 1) {
		return 0;
	}

	return NormalizeRuntimeAddress(static_cast<std::uintptr_t>(matches.front()), moduleBase);
}

std::uintptr_t ResolveUniqueCodeAddressFromPatterns(
	const std::initializer_list<const char*> patterns,
	std::uintptr_t moduleBase)
{
	for (const char* pattern : patterns) {
		const std::uintptr_t resolvedAddress = ResolveUniqueCodeAddress(pattern, moduleBase);
		if (resolvedAddress != 0) {
			return resolvedAddress;
		}
	}
	return 0;
}
#endif

typedef int(__thiscall* OriginalEStartDebugFuncType)(DWORD* thisPtr, int a2, int a3);
OriginalEStartDebugFuncType originalEStartDebugFunc = nullptr;

typedef int(__thiscall* OriginalEStartCompileFuncType)(DWORD* thisPtr, int a2);
OriginalEStartCompileFuncType originalEStartCompileFunc = nullptr;

int __fastcall MyEStartCompileFunc(DWORD* thisPtr, int dummy, int a2)
{
	OutputStringToELog("compile start");
	RunChangeECOM(true);
	return originalEStartCompileFunc(thisPtr, a2);
}

int __fastcall MyEStartDebugFunc(DWORD* thisPtr, int dummy, int a2, int a3)
{
	OutputStringToELog("debug start");
	RunChangeECOM(false);
	return originalEStartDebugFunc(thisPtr, a2, a3);
}

#if defined(_M_IX86)
int __fastcall MyProjectSerializeToFile(void* thisPtr, int /*dummy*/, LPCSTR lpFileName, int a3)
{
	if (thisPtr != nullptr) {
		const std::string sourcePath = lpFileName != nullptr ? std::string(lpFileName) : std::string();
		e571::ProjectBinarySerializer::Instance().RecordVerifiedSerializerContext(
			thisPtr,
			sourcePath);
	}
	return originalProjectSerializeToFile(thisPtr, lpFileName, a3);
}
#endif

HANDLE WINAPI MyCreateFileA(
	LPCSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDisposition,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile)
{
	if (lpFileName != nullptr && std::string(lpFileName).find("\\Temp\\e_debug\\") != std::string::npos) {
		OutputStringToELog("debug precompile finished");
		g_preDebugging = false;
	}

	std::filesystem::path currentPath = GetBasePath();
	std::filesystem::path autoLinkerPath = currentPath / "tools" / "link.ini";
	if (lpFileName != nullptr && autoLinkerPath.string() == std::string(lpFileName)) {
		g_preCompiling = false;

		auto linkName = g_configManager.getValue(g_nowOpenSourceFilePath);
		if (!linkName.empty()) {
			auto linkConfig = g_linkerManager.getConfig(linkName);
			if (std::filesystem::exists(linkConfig.path)) {
				OutputStringToELog("switch linker: " + linkConfig.name + " " + linkConfig.path);
				return originalCreateFileA(
					linkConfig.path.c_str(),
					dwDesiredAccess,
					dwShareMode,
					lpSecurityAttributes,
					dwCreationDisposition,
					dwFlagsAndAttributes,
					hTemplateFile);
			}

			OutputStringToELog("failed to switch linker: file missing");
		}
		else {
			OutputStringToELog("linker not configured for current source file");
		}
	}

	return originalCreateFileA(
		lpFileName,
		dwDesiredAccess,
		dwShareMode,
		lpSecurityAttributes,
		dwCreationDisposition,
		dwFlagsAndAttributes,
		hTemplateFile);
}

BOOL APIENTRY MyGetSaveFileNameA(LPOPENFILENAMEA item)
{
	std::string silentDiagnostics;
	if (TryHandleSilentCompileOutputPathRequest(item, &silentDiagnostics)) {
		OutputStringToELog(std::format(
			"[SilentCompile] suppressed GetSaveFileNameA path={}",
			item != nullptr && item->lpstrFile != nullptr ? item->lpstrFile : ""));
		return TRUE;
	}

	if (g_preCompiling) {
		OutputStringToELog("precompile finished");
		g_preCompiling = false;
	}

	return originalGetSaveFileNameA(item);
}

BOOL WINAPI MyCreateProcessA(
	LPCSTR lpApplicationName,
	LPSTR lpCommandLine,
	LPSECURITY_ATTRIBUTES lpProcessAttributes,
	LPSECURITY_ATTRIBUTES lpThreadAttributes,
	BOOL bInheritHandles,
	DWORD dwCreationFlags,
	LPVOID lpEnvironment,
	LPCSTR lpCurrentDirectory,
	LPSTARTUPINFOA lpStartupInfo,
	LPPROCESS_INFORMATION lpProcessInformation)
{
	std::string commandLine = lpCommandLine != nullptr ? lpCommandLine : "";

	const std::string krnlnPath = GetLinkerCommandKrnlnFileName(lpCommandLine);
	OutputStringToELog(krnlnPath);

	if (!krnlnPath.empty()) {
		const std::string libFilePath = std::format("{}\\AutoLinker\\ForceLinkLib.txt", GetBasePath());
		const auto libList = ReadFileAndSplitLines(libFilePath);

		if (!libList.empty()) {
			const std::string currentLinkerName = g_configManager.getValue(g_nowOpenSourceFilePath);
			OutputStringToELog(std::format("current linker: {}", currentLinkerName));

			std::string libCmd;
			for (const auto& line : libList) {
				const auto parts = SplitStringTwo(line, '=');
				std::string libPath = line;
				std::string linkerName;
				if (parts.size() == 2) {
					linkerName = parts[0];
					libPath = parts[1];
				}

				if (!linkerName.empty()) {
					if (currentLinkerName.find(linkerName) != std::string::npos) {
						if (std::filesystem::exists(libPath)) {
							OutputStringToELog(std::format("force link lib: {} -> {}", linkerName, libPath));
							if (!libCmd.empty()) {
								libCmd += " ";
							}
							libCmd += "\"" + libPath + "\"";
						}
					}
					else {
						OutputStringToELog(std::format(
							"skip lib for linker mismatch: rule={} current={} path={}",
							linkerName,
							currentLinkerName,
							libPath));
					}
				}
				else if (std::filesystem::exists(libPath)) {
					OutputStringToELog(std::format("force link lib: {}", libPath));
					if (!libCmd.empty()) {
						libCmd += " ";
					}
					libCmd += "\"" + libPath + "\"";
				}
			}

			const std::string newLibs = libCmd + " \"" + krnlnPath + "\" /FORCE";
			commandLine = ReplaceSubstring(commandLine, "\"" + krnlnPath + "\"", newLibs);
		}
	}

	const std::string outFileName = GetLinkerCommandOutFileName(lpCommandLine);
	if (!outFileName.empty()) {
		if (commandLine.find("/pdb:\"build.pdb\"") != std::string::npos) {
			const std::string newPdbCommand = std::format("/pdb:\"{}.pdb\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/pdb:\"build.pdb\"", newPdbCommand);
		}
		if (commandLine.find("/map:\"build.map\"") != std::string::npos) {
			const std::string newMapCommand = std::format("/map:\"{}.map\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/map:\"build.map\"", newMapCommand);
		}
	}

	OutputStringToELog("launch command: " + commandLine);
	return originalCreateProcessA(
		lpApplicationName,
		commandLine.data(),
		lpProcessAttributes,
		lpThreadAttributes,
		bInheritHandles,
		dwCreationFlags,
		lpEnvironment,
		lpCurrentDirectory,
		lpStartupInfo,
		lpProcessInformation);
}

int WINAPI MyMessageBoxA(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType)
{
	const std::string caption = lpCaption != nullptr ? lpCaption : "";
	const std::string text = lpText != nullptr ? lpText : "";
	if (IsSilentCompileOutputPathRequestOngoingForCurrentThread() &&
		IsSilentCompileModulePublishPrompt(text)) {
		OutputStringToELog("auto-confirm ecom publish prompt during silent compile");
		return IDYES;
	}
	if (caption.find("linker output contains too many errors or warnings") != std::string::npos) {
		HeadlessCompileRunner::ReportIdeMessageBoxA(caption, text, uType);
		return IDNO;
	}
	if (text.find("linker output contains too many errors or warnings") != std::string::npos) {
		HeadlessCompileRunner::ReportIdeMessageBoxA(caption, text, uType);
		return IDNO;
	}
	if (HeadlessCompileRunner::HasHeadlessCompileRequest()) {
		HeadlessCompileRunner::ReportIdeMessageBoxA(caption, text, uType);
		return HeadlessCompileRunner::GetMessageBoxAutoResponse(uType);
	}
	return originalMessageBoxA(hWnd, lpText, lpCaption, uType);
}

int WINAPI MyMessageBoxW(HWND hWnd, LPCWSTR lpText, LPCWSTR lpCaption, UINT uType)
{
	const std::wstring caption = lpCaption != nullptr ? lpCaption : L"";
	const std::wstring text = lpText != nullptr ? lpText : L"";
	if (HeadlessCompileRunner::HasHeadlessCompileRequest()) {
		HeadlessCompileRunner::ReportIdeMessageBoxW(caption, text, uType);
		return HeadlessCompileRunner::GetMessageBoxAutoResponse(uType);
	}
	return originalMessageBoxW(hWnd, lpText, lpCaption, uType);
}

BOOL WINAPI MyTrackPopupMenu(
	HMENU hMenu,
	UINT uFlags,
	int x,
	int y,
	int nReserved,
	HWND hWnd,
	const RECT* prcRect)
{
	PrepareAutoLinkerPopupMenu(hMenu);
	return originalTrackPopupMenu(hMenu, uFlags, x, y, nReserved, hWnd, prcRect);
}

BOOL WINAPI MyTrackPopupMenuEx(
	HMENU hMenu,
	UINT uFlags,
	int x,
	int y,
	HWND hWnd,
	LPTPMPARAMS lptpm)
{
	PrepareAutoLinkerPopupMenu(hMenu);
	return originalTrackPopupMenuEx(hMenu, uFlags, x, y, hWnd, lptpm);
}

void StartHookCreateFileA()
{
	std::lock_guard<std::mutex> lock(g_detourHookMutex);
	if (g_fullHookInstalled) {
		return;
	}

	DetourRestoreAfterWith();
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)originalCreateFileA, MyCreateFileA);
	DetourAttach(&(PVOID&)originalGetSaveFileNameA, MyGetSaveFileNameA);
	DetourAttach(&(PVOID&)originalCreateProcessA, MyCreateProcessA);
	const bool attachMessageBoxHooks = !g_messageBoxHookInstalled;
	if (attachMessageBoxHooks) {
		DetourAttach(&(PVOID&)originalMessageBoxA, MyMessageBoxA);
		DetourAttach(&(PVOID&)originalMessageBoxW, MyMessageBoxW);
	}
	DetourAttach(&(PVOID&)originalTrackPopupMenu, MyTrackPopupMenu);
	DetourAttach(&(PVOID&)originalTrackPopupMenuEx, MyTrackPopupMenuEx);

	if (g_debugStartAddress > 0 && g_compileStartAddress > 0) {
		originalEStartCompileFunc = (OriginalEStartCompileFuncType)g_compileStartAddress;
		originalEStartDebugFunc = (OriginalEStartDebugFuncType)g_debugStartAddress;
		DetourAttach(&(PVOID&)originalEStartCompileFunc, MyEStartCompileFunc);
		DetourAttach(&(PVOID&)originalEStartDebugFunc, MyEStartDebugFunc);
	}
	else {
		OutputStringToELog(std::format(
			"missing compile/debug start address, skip hook debug=0x{:X} compile=0x{:X}",
			g_debugStartAddress,
			g_compileStartAddress));
	}

#if defined(_M_IX86)
	const std::uintptr_t moduleBase = reinterpret_cast<std::uintptr_t>(GetModuleHandleA(nullptr));
	if (moduleBase != 0) {
		const std::uintptr_t projectSerializeToFileRva = ResolveUniqueCodeAddressFromPatterns(
			{
				kProjectSerializeToFilePattern,
			},
			moduleBase);
		if (projectSerializeToFileRva != 0) {
			originalProjectSerializeToFile = reinterpret_cast<OriginalProjectSerializeToFileFuncType>(
				ResolveInternalAddress<void*>(
					moduleBase,
					projectSerializeToFileRva));
		}
		if (originalProjectSerializeToFile != nullptr && projectSerializeToFileRva != 0) {
			OutputStringToELog(std::format(
				"[ProjectBinarySerializer] hook serialize_to_file rva=0x{:X} via={}",
				projectSerializeToFileRva,
				"pattern"));
			DetourAttach(&(PVOID&)originalProjectSerializeToFile, MyProjectSerializeToFile);
		}
		else {
			OutputStringToELog("[ProjectBinarySerializer] hook serialize_to_file address unavailable");
		}
	}
#endif

	const LONG error = DetourTransactionCommit();
	if (error == NO_ERROR) {
		g_fullHookInstalled = true;
		if (attachMessageBoxHooks) {
			g_messageBoxHookInstalled = true;
		}
	}
	else {
		OutputStringToELog(std::format("hook transaction failed, error={}", error));
	}
}

void StartHeadlessMessageBoxHook()
{
	if (!HeadlessCompileRunner::HasHeadlessCompileRequest()) {
		return;
	}

	std::lock_guard<std::mutex> lock(g_detourHookMutex);
	if (g_messageBoxHookInstalled) {
		return;
	}

	DetourRestoreAfterWith();
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)originalMessageBoxA, MyMessageBoxA);
	DetourAttach(&(PVOID&)originalMessageBoxW, MyMessageBoxW);
	const LONG error = DetourTransactionCommit();
	if (error == NO_ERROR) {
		g_messageBoxHookInstalled = true;
	}
}
