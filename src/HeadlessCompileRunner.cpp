#include "HeadlessCompileRunner.h"

#include <Windows.h>
#include <Shellapi.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <format>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AIChatFeature.h"
#include "Global.h"
#include "PathHelper.h"

#pragma comment(lib, "Shell32.lib")

namespace {

constexpr int kDefaultStartupTimeoutSeconds = 60;
constexpr int kMaxStartupTimeoutSeconds = 600;
constexpr int kExitDelayMilliseconds = 300;

struct HeadlessCompileRequest {
	bool enabled = false;
	std::string target = "auto";
	std::string outputPath;
	bool staticCompile = false;
	bool hideWindow = true;
	bool exitAfterCompile = true;
	int startupTimeoutSeconds = kDefaultStartupTimeoutSeconds;
	std::string resultPath;
};

std::once_flag g_parseOnce;
HeadlessCompileRequest g_request;
std::atomic_bool g_started = false;

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int size = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (size <= 0) {
		return std::string();
	}

	std::string out(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		out.data(),
		size,
		nullptr,
		nullptr) <= 0) {
		return std::string();
	}
	return out;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	int size = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	DWORD flags = MB_ERR_INVALID_CHARS;
	UINT codePage = CP_UTF8;
	if (size <= 0) {
		size = MultiByteToWideChar(
			CP_ACP,
			0,
			text.data(),
			static_cast<int>(text.size()),
			nullptr,
			0);
		flags = 0;
		codePage = CP_ACP;
	}
	if (size <= 0) {
		return std::wstring();
	}

	std::wstring out(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(
		codePage,
		flags,
		text.data(),
		static_cast<int>(text.size()),
		out.data(),
		size) <= 0) {
		return std::wstring();
	}
	return out;
}

std::string TrimAsciiCopyLocal(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string ToLowerAsciiCopyLocal(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string NormalizeOptionName(std::string name)
{
	name = ToLowerAsciiCopyLocal(TrimAsciiCopyLocal(name));
	while (!name.empty() && (name.front() == '-' || name.front() == '/')) {
		name.erase(name.begin());
	}
	for (char& ch : name) {
		if (ch == '_') {
			ch = '-';
		}
	}
	return name;
}

bool TryParseBoolText(const std::string& raw, bool& outValue)
{
	const std::string value = ToLowerAsciiCopyLocal(TrimAsciiCopyLocal(raw));
	if (value == "1" || value == "true" || value == "yes" || value == "on") {
		outValue = true;
		return true;
	}
	if (value == "0" || value == "false" || value == "no" || value == "off") {
		outValue = false;
		return true;
	}
	return false;
}

bool GetJsonBool(const nlohmann::json& json, const char* key, bool fallback)
{
	if (!json.is_object() || key == nullptr || !json.contains(key)) {
		return fallback;
	}
	const nlohmann::json& value = json[key];
	if (value.is_boolean()) {
		return value.get<bool>();
	}
	if (value.is_number_integer()) {
		return value.get<int>() != 0;
	}
	if (value.is_string()) {
		bool parsed = false;
		if (TryParseBoolText(value.get<std::string>(), parsed)) {
			return parsed;
		}
	}
	return fallback;
}

std::string GetJsonString(const nlohmann::json& json, const char* key, const std::string& fallback)
{
	if (!json.is_object() || key == nullptr || !json.contains(key) || !json[key].is_string()) {
		return fallback;
	}
	return json[key].get<std::string>();
}

int GetJsonBoundedInt(const nlohmann::json& json, const char* key, int fallback, int minValue, int maxValue)
{
	if (!json.is_object() || key == nullptr || !json.contains(key)) {
		return fallback;
	}
	try {
		int value = fallback;
		if (json[key].is_number_integer()) {
			value = json[key].get<int>();
		}
		else if (json[key].is_string()) {
			value = std::stoi(json[key].get<std::string>());
		}
		return (std::clamp)(value, minValue, maxValue);
	}
	catch (...) {
		return fallback;
	}
}

std::wstring GetEnvironmentVariableWide(const wchar_t* name)
{
	if (name == nullptr || *name == L'\0') {
		return std::wstring();
	}

	DWORD size = GetEnvironmentVariableW(name, nullptr, 0);
	if (size == 0) {
		return std::wstring();
	}

	std::wstring value(static_cast<size_t>(size), L'\0');
	const DWORD copied = GetEnvironmentVariableW(name, value.data(), size);
	if (copied == 0 || copied >= size) {
		return std::wstring();
	}
	value.resize(copied);
	return value;
}

void ApplyJsonRequest(const nlohmann::json& json, HeadlessCompileRequest& request)
{
	request.enabled = GetJsonBool(json, "enabled", request.enabled);
	request.target = GetJsonString(json, "target", request.target);
	request.outputPath = GetJsonString(json, "output_path", request.outputPath);
	request.staticCompile = GetJsonBool(json, "static_compile", request.staticCompile);
	request.hideWindow = GetJsonBool(json, "hide_window", request.hideWindow);
	request.exitAfterCompile = GetJsonBool(json, "exit_after_compile", request.exitAfterCompile);
	request.startupTimeoutSeconds = GetJsonBoundedInt(
		json,
		"startup_timeout_seconds",
		request.startupTimeoutSeconds,
		1,
		kMaxStartupTimeoutSeconds);
	request.resultPath = GetJsonString(json, "result_path", request.resultPath);
}

void ApplyEnvironmentRequest(HeadlessCompileRequest& request)
{
	const std::wstring raw = GetEnvironmentVariableWide(L"AUTOLINKER_HEADLESS_COMPILE");
	if (!raw.empty()) {
		const std::string rawUtf8 = WideToUtf8(raw);
		const std::string trimmed = TrimAsciiCopyLocal(rawUtf8);
		if (!trimmed.empty() && trimmed.front() == '{') {
			try {
				ApplyJsonRequest(nlohmann::json::parse(trimmed), request);
			}
			catch (...) {
				request.enabled = true;
			}
		}
		else {
			bool enabled = true;
			TryParseBoolText(trimmed, enabled);
			request.enabled = enabled;
		}
	}

	const auto applyStringEnv = [](const wchar_t* name, std::string& target) {
		const std::wstring value = GetEnvironmentVariableWide(name);
		if (!value.empty()) {
			target = WideToUtf8(value);
		}
	};
	const auto applyBoolEnv = [](const wchar_t* name, bool& target) {
		const std::wstring value = GetEnvironmentVariableWide(name);
		if (!value.empty()) {
			bool parsed = target;
			if (TryParseBoolText(WideToUtf8(value), parsed)) {
				target = parsed;
			}
		}
	};
	const auto applyIntEnv = [](const wchar_t* name, int& target) {
		const std::wstring value = GetEnvironmentVariableWide(name);
		if (!value.empty()) {
			try {
				target = (std::clamp)(std::stoi(WideToUtf8(value)), 1, kMaxStartupTimeoutSeconds);
			}
			catch (...) {
			}
		}
	};

	applyStringEnv(L"AUTOLINKER_HEADLESS_TARGET", request.target);
	applyStringEnv(L"AUTOLINKER_HEADLESS_OUTPUT", request.outputPath);
	applyStringEnv(L"AUTOLINKER_HEADLESS_OUTPUT_PATH", request.outputPath);
	applyStringEnv(L"AUTOLINKER_HEADLESS_RESULT", request.resultPath);
	applyStringEnv(L"AUTOLINKER_HEADLESS_RESULT_PATH", request.resultPath);
	applyBoolEnv(L"AUTOLINKER_HEADLESS_STATIC", request.staticCompile);
	applyBoolEnv(L"AUTOLINKER_HEADLESS_STATIC_COMPILE", request.staticCompile);
	applyBoolEnv(L"AUTOLINKER_HEADLESS_HIDE_WINDOW", request.hideWindow);
	applyBoolEnv(L"AUTOLINKER_HEADLESS_EXIT", request.exitAfterCompile);
	applyIntEnv(L"AUTOLINKER_HEADLESS_STARTUP_TIMEOUT", request.startupTimeoutSeconds);
}

std::vector<std::wstring> GetCommandLineArguments()
{
	std::vector<std::wstring> args;
	int argc = 0;
	LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
	if (argv == nullptr) {
		return args;
	}

	args.reserve(static_cast<size_t>(argc));
	for (int i = 0; i < argc; ++i) {
		args.emplace_back(argv[i] == nullptr ? L"" : argv[i]);
	}
	LocalFree(argv);
	return args;
}

bool TrySplitInlineOption(
	const std::wstring& arg,
	std::string& outName,
	std::string& outValue,
	bool& outHasValue)
{
	outName.clear();
	outValue.clear();
	outHasValue = false;

	if (arg.rfind(L"--", 0) != 0 && arg.rfind(L"/autolinker-", 0) != 0 && arg.rfind(L"/AutoLinker-", 0) != 0) {
		return false;
	}

	const size_t prefixLen = arg.rfind(L"--", 0) == 0 ? 2 : 1;
	size_t splitPos = arg.find(L'=', prefixLen);
	if (splitPos == std::wstring::npos) {
		splitPos = arg.find(L':', prefixLen);
	}

	if (splitPos == std::wstring::npos) {
		outName = NormalizeOptionName(WideToUtf8(arg.substr(prefixLen)));
		return true;
	}

	outName = NormalizeOptionName(WideToUtf8(arg.substr(prefixLen, splitPos - prefixLen)));
	outValue = WideToUtf8(arg.substr(splitPos + 1));
	outHasValue = true;
	return true;
}

std::string ReadNextArgumentValue(const std::vector<std::wstring>& args, size_t& index)
{
	if (index + 1 >= args.size()) {
		return std::string();
	}
	++index;
	return WideToUtf8(args[index]);
}

void ApplyCommandLineRequest(HeadlessCompileRequest& request)
{
	const std::vector<std::wstring> args = GetCommandLineArguments();
	for (size_t i = 1; i < args.size(); ++i) {
		std::string name;
		std::string value;
		bool hasValue = false;
		if (!TrySplitInlineOption(args[i], name, value, hasValue)) {
			continue;
		}

		if (name == "autolinker-headless-compile" || name == "autolinker-compile") {
			request.enabled = true;
			if (hasValue) {
				bool parsed = true;
				if (TryParseBoolText(value, parsed)) {
					request.enabled = parsed;
				}
			}
			continue;
		}
		if (name == "autolinker-output" || name == "autolinker-output-path") {
			request.outputPath = hasValue ? value : ReadNextArgumentValue(args, i);
			continue;
		}
		if (name == "autolinker-target") {
			request.target = hasValue ? value : ReadNextArgumentValue(args, i);
			continue;
		}
		if (name == "autolinker-static" || name == "autolinker-static-compile") {
			request.staticCompile = true;
			if (hasValue) {
				bool parsed = true;
				if (TryParseBoolText(value, parsed)) {
					request.staticCompile = parsed;
				}
			}
			continue;
		}
		if (name == "autolinker-no-static") {
			request.staticCompile = false;
			continue;
		}
		if (name == "autolinker-result" || name == "autolinker-result-path") {
			request.resultPath = hasValue ? value : ReadNextArgumentValue(args, i);
			continue;
		}
		if (name == "autolinker-startup-timeout" || name == "autolinker-startup-timeout-seconds") {
			const std::string raw = hasValue ? value : ReadNextArgumentValue(args, i);
			try {
				request.startupTimeoutSeconds = (std::clamp)(std::stoi(raw), 1, kMaxStartupTimeoutSeconds);
			}
			catch (...) {
			}
			continue;
		}
		if (name == "autolinker-show-window") {
			request.hideWindow = false;
			continue;
		}
		if (name == "autolinker-hide-window") {
			request.hideWindow = true;
			if (hasValue) {
				bool parsed = true;
				if (TryParseBoolText(value, parsed)) {
					request.hideWindow = parsed;
				}
			}
			continue;
		}
		if (name == "autolinker-no-exit" || name == "autolinker-keep-open") {
			request.exitAfterCompile = false;
			continue;
		}
		if (name == "autolinker-exit") {
			request.exitAfterCompile = true;
			if (hasValue) {
				bool parsed = true;
				if (TryParseBoolText(value, parsed)) {
					request.exitAfterCompile = parsed;
				}
			}
			continue;
		}
	}
}

void EnsureRequestParsed()
{
	std::call_once(g_parseOnce, []() {
		HeadlessCompileRequest request;
		ApplyEnvironmentRequest(request);
		ApplyCommandLineRequest(request);
		request.target = ToLowerAsciiCopyLocal(TrimAsciiCopyLocal(request.target.empty() ? "auto" : request.target));
		request.outputPath = TrimAsciiCopyLocal(request.outputPath);
		request.resultPath = TrimAsciiCopyLocal(request.resultPath);
		g_request = request;
	});
}

HeadlessCompileRequest GetRequestCopy()
{
	EnsureRequestParsed();
	return g_request;
}

std::string NormalizeCompileTarget(std::string target)
{
	target = ToLowerAsciiCopyLocal(TrimAsciiCopyLocal(target));
	for (char& ch : target) {
		if (ch == '-') {
			ch = '_';
		}
	}

	if (target.empty() || target == "auto") {
		return "auto";
	}
	if (target == "exe" || target == "win" || target == "windows" || target == "window" || target == "winexe") {
		return "win_exe";
	}
	if (target == "console" || target == "conole" || target == "win_console" || target == "winconsole") {
		return "win_console_exe";
	}
	if (target == "dll" || target == "win_dll") {
		return "win_dll";
	}
	if (target == "ec" || target == "ecom" || target == "module") {
		return "ecom";
	}
	return target;
}

bool IsSupportedCompileTarget(const std::string& target)
{
	return target == "win_exe" ||
		target == "win_console_exe" ||
		target == "win_dll" ||
		target == "ecom";
}

std::filesystem::path MakeUtf8Path(const std::string& utf8Path)
{
	const std::wstring wide = Utf8ToWide(utf8Path);
	if (!wide.empty()) {
		return std::filesystem::path(wide);
	}
	return std::filesystem::path(utf8Path);
}

std::filesystem::path GetDefaultResultPath()
{
	return GetAutoLinkerLogFilePath("headless_compile_last.json");
}

bool WriteTextFileUtf8(const std::filesystem::path& path, const std::string& text, std::string* outError)
{
	try {
		const std::filesystem::path parent = path.parent_path();
		if (!parent.empty()) {
			std::filesystem::create_directories(parent);
		}

		std::ofstream out(path, std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			if (outError != nullptr) {
				*outError = "open_result_file_failed";
			}
			return false;
		}
		out.write(text.data(), static_cast<std::streamsize>(text.size()));
		return true;
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) {
			*outError = ex.what();
		}
		return false;
	}
	catch (...) {
		if (outError != nullptr) {
			*outError = "unknown";
		}
		return false;
	}
}

void WriteResultFiles(const HeadlessCompileRequest& request, const nlohmann::json& result)
{
	const std::string text = result.dump(2);
	std::string error;
	if (!WriteTextFileUtf8(GetDefaultResultPath(), text, &error)) {
		OutputStringToELog("[HeadlessCompile] 写入默认结果文件失败: " + error);
	}

	if (!request.resultPath.empty() &&
		!WriteTextFileUtf8(MakeUtf8Path(request.resultPath), text, &error)) {
		OutputStringToELog("[HeadlessCompile] 写入指定结果文件失败: " + error);
	}
}

bool CallPublicToolUtf8(
	const std::string& toolName,
	const nlohmann::json& arguments,
	nlohmann::json& outJson,
	bool& outToolOk,
	std::string& outRawText)
{
	outJson = nlohmann::json::object();
	outToolOk = false;
	outRawText.clear();

	std::string resultText;
	bool toolOk = false;
	if (!AIChatFeature::ExecutePublicTool(toolName, arguments.dump(), resultText, toolOk)) {
		outRawText = "ExecutePublicTool transport failed";
		return false;
	}

	outRawText = resultText;
	outToolOk = toolOk;
	try {
		outJson = resultText.empty() ? nlohmann::json::object() : nlohmann::json::parse(resultText);
		return true;
	}
	catch (const std::exception& ex) {
		outJson = {
			{"ok", false},
			{"error", std::string("parse tool result failed: ") + ex.what()},
			{"raw_text", resultText}
		};
		return false;
	}
}

bool WaitForEideInfo(
	const HeadlessCompileRequest& request,
	nlohmann::json& outInfo,
	std::string& outError)
{
	const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(request.startupTimeoutSeconds);
	std::string lastError;

	while (std::chrono::steady_clock::now() < deadline) {
		bool toolOk = false;
		std::string rawText;
		nlohmann::json info;
		if (CallPublicToolUtf8("get_current_eide_info", nlohmann::json::object(), info, toolOk, rawText) &&
			toolOk &&
			info.value("ok", false)) {
			const std::string projectType = info.value("project_type", std::string("unknown"));
			const bool projectTypeKnown = projectType != "unknown";
			const bool autoTargetResolved = NormalizeCompileTarget(request.target) != "auto" || projectTypeKnown;
			const bool sourcePathReady =
				info.value("source_file_exists", false) ||
				!TrimAsciiCopyLocal(info.value("source_file_path", std::string())).empty();
			const std::string mainTitle = info.value("main_window_title", std::string());
			const bool projectReady = projectTypeKnown || sourcePathReady;
			if (autoTargetResolved && projectReady && !mainTitle.empty()) {
				outInfo = std::move(info);
				return true;
			}
			lastError = std::format("waiting_project_ready target={} project_type={} source_ready={} title_empty={}",
				request.target,
				projectType,
				sourcePathReady ? 1 : 0,
				mainTitle.empty() ? 1 : 0);
		}
		else {
			lastError = rawText.empty() ? "get_current_eide_info_failed" : rawText;
		}
		Sleep(500);
	}

	outError = lastError.empty() ? "wait_eide_info_timeout" : lastError;
	return false;
}

std::string ResolveTargetFromInfo(
	const HeadlessCompileRequest& request,
	const nlohmann::json& info,
	std::string& outError)
{
	std::string target = NormalizeCompileTarget(request.target);
	if (target == "auto") {
		target = NormalizeCompileTarget(info.value("project_type", std::string("unknown")));
	}

	if (!IsSupportedCompileTarget(target)) {
		outError = "unsupported_or_unknown_target: " + target;
		return std::string();
	}
	if (target == "ecom" && request.staticCompile) {
		outError = "ecom_does_not_support_static_compile";
		return std::string();
	}
	return target;
}

bool IsCompileArtifactUpdated(const nlohmann::json& compileResult)
{
	return compileResult.value("ok", false) &&
		compileResult.value("output_file_exists", false) &&
		compileResult.value("output_file_modified_after_compile", false);
}

nlohmann::json BuildRequestJson(const HeadlessCompileRequest& request)
{
	return {
		{"target", request.target},
		{"output_path", request.outputPath},
		{"static_compile", request.staticCompile},
		{"hide_window", request.hideWindow},
		{"exit_after_compile", request.exitAfterCompile},
		{"startup_timeout_seconds", request.startupTimeoutSeconds},
		{"result_path", request.resultPath}
	};
}

void FinishHeadlessRun(const HeadlessCompileRequest& request, nlohmann::json result, int exitCode)
{
	result["exit_code"] = exitCode;
	WriteResultFiles(request, result);

	if (result.value("ok", false)) {
		OutputStringToELog("[HeadlessCompile] 编译完成");
	}
	else {
		OutputStringToELog("[HeadlessCompile] 编译失败: " + result.value("error", std::string("unknown")));
	}

	if (request.exitAfterCompile) {
		Sleep(kExitDelayMilliseconds);
		TerminateProcess(GetCurrentProcess(), static_cast<UINT>(exitCode));
		ExitProcess(static_cast<UINT>(exitCode));
	}
}

void HeadlessWorkerMain()
{
	const HeadlessCompileRequest request = GetRequestCopy();
	nlohmann::json result;
	result["ok"] = false;
	result["mode"] = "autolinker_headless_compile";
	result["request"] = BuildRequestJson(request);
	result["process_id"] = GetCurrentProcessId();
	result["result_file"] = GetDefaultResultPath().string();

	OutputStringToELog("[HeadlessCompile] 已进入无头编译模式");

	if (request.hideWindow) {
		HeadlessCompileRunner::ApplyInitialWindowState(g_hwnd);
	}

	if (TrimAsciiCopyLocal(request.outputPath).empty()) {
		result["error"] = "output_path is required";
		FinishHeadlessRun(request, std::move(result), 2);
		return;
	}

	nlohmann::json eideInfo;
	std::string waitError;
	if (!WaitForEideInfo(request, eideInfo, waitError)) {
		result["error"] = "wait_eide_info_timeout";
		result["trace"] = waitError;
		FinishHeadlessRun(request, std::move(result), 3);
		return;
	}
	result["eide_info"] = eideInfo;

	std::string targetError;
	const std::string resolvedTarget = ResolveTargetFromInfo(request, eideInfo, targetError);
	if (resolvedTarget.empty()) {
		result["error"] = targetError.empty() ? "resolve_target_failed" : targetError;
		FinishHeadlessRun(request, std::move(result), 2);
		return;
	}
	result["resolved_target"] = resolvedTarget;

	nlohmann::json compileArgs = {
		{"target", resolvedTarget},
		{"output_path", request.outputPath},
		{"static_compile", request.staticCompile}
	};

	OutputStringToELog(std::format(
		"[HeadlessCompile] 开始编译 target={} static={} output={}",
		resolvedTarget,
		request.staticCompile ? 1 : 0,
		request.outputPath));

	nlohmann::json compileResult;
	bool toolOk = false;
	std::string rawText;
	const bool callOk = CallPublicToolUtf8("compile_with_output_path", compileArgs, compileResult, toolOk, rawText);
	result["compile_tool_ok"] = toolOk;
	result["compile_result"] = compileResult;

	if (!callOk) {
		result["error"] = "compile_tool_transport_or_parse_failed";
		result["raw_compile_result"] = rawText;
		FinishHeadlessRun(request, std::move(result), 4);
		return;
	}

	if (!toolOk || !compileResult.value("ok", false)) {
		result["error"] = compileResult.value("error", std::string("compile_with_output_path_failed"));
		FinishHeadlessRun(request, std::move(result), 1);
		return;
	}

	if (!IsCompileArtifactUpdated(compileResult)) {
		result["error"] = "compile_invoked_but_output_file_not_updated";
		FinishHeadlessRun(request, std::move(result), 1);
		return;
	}

	result["ok"] = true;
	FinishHeadlessRun(request, std::move(result), 0);
}

} // namespace

namespace HeadlessCompileRunner {

bool HasHeadlessCompileRequest()
{
	return GetRequestCopy().enabled;
}

void ApplyInitialWindowState(HWND mainWindow)
{
	const HeadlessCompileRequest request = GetRequestCopy();
	if (!request.enabled || !request.hideWindow || mainWindow == nullptr || !IsWindow(mainWindow)) {
		return;
	}

	ShowWindow(mainWindow, SW_HIDE);
	SetWindowPos(
		mainWindow,
		nullptr,
		0,
		0,
		0,
		0,
		SWP_HIDEWINDOW | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

void StartIfRequested()
{
	const HeadlessCompileRequest request = GetRequestCopy();
	if (!request.enabled) {
		return;
	}

	bool expected = false;
	if (!g_started.compare_exchange_strong(expected, true)) {
		return;
	}

	std::thread(HeadlessWorkerMain).detach();
}

} // namespace HeadlessCompileRunner
