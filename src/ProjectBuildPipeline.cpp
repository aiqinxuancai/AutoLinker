#include "ProjectBuildPipeline.h"

#include <Windows.h>

#include <atomic>
#include <chrono>
#include <exception>
#include <filesystem>
#include <fstream>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include "..\\thirdparty\\json.hpp"
#include "EideProjectBinarySerializer.h"
#include "Global.h"
#include "HeadlessCompileRunner.h"
#include "Logger.h"
#include "PowerShellToolRunner.h"
#include "UnicodeTextCodec.h"

namespace {

constexpr DWORD kHeadlessCompileTimeoutMilliseconds = 630000;
constexpr ULONGLONG kHeadlessStartupWindowSuppressionMilliseconds = 10000;
constexpr DWORD kHeadlessStartupWindowPollMilliseconds = 10;
std::atomic<unsigned long long> g_projectBuildJobCounter = 0;

std::string ProjectBuildUtf8(std::u8string_view text)
{
	return std::string(reinterpret_cast<const char*>(text.data()), text.size());
}

void OutputProjectBuildUtf8(const std::string& text)
{
	OutputStringToELog(UnicodeTextCodec::Utf8ToLocalPreservingUnicode(text));
}

struct ProjectBuildJob {
	std::string id;
	ProjectBuildConfig config;
	ProjectBuildVariableContext variables;
	std::filesystem::path jobDirectory;
	std::filesystem::path snapshotPath;
	std::filesystem::path resultPath;
	std::filesystem::path eidePath;
};

struct ProjectBuildJobCleanup {
	std::filesystem::path snapshotPath;
	std::filesystem::path jobDirectory;
	~ProjectBuildJobCleanup()
	{
		std::error_code error;
		std::filesystem::remove(snapshotPath, error);
		error.clear();
		std::filesystem::remove_all(jobDirectory, error);
	}
};

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) return {};
	const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) return {};
	std::string result(static_cast<size_t>(size), '\0');
	WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr);
	return result;
}

// IDE 的 EProjectSerializer 接收本地代码页路径，而配置与进程参数使用 UTF-8/Unicode。
std::string WideToLocal(const std::wstring& text)
{
	if (text.empty()) return {};
	const int size = WideCharToMultiByte(
		CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) return {};
	std::string result(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(
		CP_ACP, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr) <= 0) {
		return {};
	}
	return result;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) return {};
	const int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) return {};
	std::wstring result(static_cast<size_t>(size), L'\0');
	MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), size);
	return result;
}

std::filesystem::path GetCurrentExecutablePath()
{
	std::vector<wchar_t> buffer(MAX_PATH);
	for (;;) {
		const DWORD length = GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
		if (length == 0) return {};
		if (length < buffer.size() - 1) return std::filesystem::path(std::wstring(buffer.data(), length));
		buffer.resize(buffer.size() * 2);
	}
}

std::wstring QuoteCommandLineArgument(const std::wstring& argument)
{
	std::wstring quoted = L"\"";
	size_t backslashes = 0;
	for (const wchar_t ch : argument) {
		if (ch == L'\\') {
			++backslashes;
			continue;
		}
		if (ch == L'\"') {
			quoted.append(backslashes * 2 + 1, L'\\');
			quoted.push_back(ch);
			backslashes = 0;
			continue;
		}
		quoted.append(backslashes, L'\\');
		backslashes = 0;
		quoted.push_back(ch);
	}
	quoted.append(backslashes * 2, L'\\');
	quoted.push_back(L'\"');
	return quoted;
}

void AppendCommandLineArgument(std::wstring& commandLine, const std::wstring& value)
{
	if (!commandLine.empty()) commandLine.push_back(L' ');
	commandLine += QuoteCommandLineArgument(value);
}

std::string JoinUnknownVariables(const std::vector<std::string>& variables)
{
	std::string result;
	for (const auto& variable : variables) {
		if (!result.empty()) result += ", ";
		result += "{" + variable + "}";
	}
	return result;
}

bool IsPowerShellProgressClixml(const std::string& text)
{
	std::string_view view(text);
	if (view.starts_with("\xEF\xBB\xBF")) view.remove_prefix(3);
	while (!view.empty() && (view.front() == ' ' || view.front() == '\t' || view.front() == '\r' || view.front() == '\n')) {
		view.remove_prefix(1);
	}
	if (!view.starts_with("#< CLIXML")) return false;

	bool foundProgressStream = false;
	size_t position = 0;
	while ((position = view.find(" S=\"", position)) != std::string_view::npos) {
		const size_t valueStart = position + 4;
		const size_t valueEnd = view.find('"', valueStart);
		if (valueEnd == std::string_view::npos) return false;
		const std::string_view stream = view.substr(valueStart, valueEnd - valueStart);
		if (stream != "progress" && stream != "Progress") return false;
		foundProgressStream = true;
		position = valueEnd + 1;
	}
	return foundProgressStream;
}

void OutputProjectBuildCommandStream(const std::string& text)
{
	if (!text.empty() && !IsPowerShellProgressClixml(text)) OutputProjectBuildUtf8(text);
}

bool WriteProjectSnapshot(
	const std::filesystem::path& path,
	size_t& bytesWritten,
	std::string& error,
	std::string& trace)
{
	return e571::ProjectBinarySerializer::Instance().WriteCurrentProjectFileSnapshot(
		WideToLocal(path.wstring()),
		&bytesWritten,
		&error,
		&trace);
}

bool RunCommands(
	const std::vector<std::string>& commands,
	const ProjectBuildVariableContext& context,
	const char* stage,
	std::string& error)
{
	for (const auto& command : commands) {
		if (command.empty()) continue;
		std::vector<std::string> unknown;
		const std::string expanded = ExpandProjectBuildVariables(command, context, &unknown);
		if (!unknown.empty()) {
			error = std::string(stage) + " action contains unknown variables: " + JoinUnknownVariables(unknown);
			return false;
		}
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] ") + stage + ProjectBuildUtf8(u8"动作: ") + expanded);
		const auto actionResult = PowerShellToolRunner::Run(expanded, WideToUtf8(context.projectDir.wstring()), 600);
		OutputProjectBuildCommandStream(actionResult.stdOut);
		OutputProjectBuildCommandStream(actionResult.stdErr);
		if (!actionResult.ok || actionResult.exitCode != 0) {
			error = std::string(stage) + " action failed (exit=" + std::to_string(actionResult.exitCode) + ")";
			if (!actionResult.error.empty()) error += ": " + actionResult.error;
			return false;
		}
	}
	return true;
}

bool ReadHeadlessResult(
	const std::filesystem::path& resultPath,
	DWORD processExitCode,
	nlohmann::json& result,
	std::string& error)
{
	std::ifstream input(resultPath, std::ios::binary);
	if (!input.is_open()) {
		error = "headless result file is missing (process exit=" + std::to_string(processExitCode) + ")";
		return false;
	}
	const std::string text((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
	try {
		result = nlohmann::json::parse(text);
	}
	catch (const std::exception& ex) {
		error = std::string("invalid headless result: ") + ex.what();
		return false;
	}
	if (!result.value("ok", false)) {
		error = result.value("error", std::string("headless compile failed"));
		return false;
	}
	return true;
}

bool RunHeadlessCompile(ProjectBuildJob& job, std::string& outputPath, std::string& error)
{
	std::wstring commandLine;
	AppendCommandLineArgument(commandLine, job.eidePath.wstring());
	AppendCommandLineArgument(commandLine, job.snapshotPath.wstring());
	AppendCommandLineArgument(commandLine, L"--autolinker-headless-compile");
	AppendCommandLineArgument(commandLine, L"--autolinker-output");
	AppendCommandLineArgument(commandLine, job.variables.outputPath.wstring());
	AppendCommandLineArgument(commandLine, L"--autolinker-target");
	AppendCommandLineArgument(commandLine, Utf8ToWide(ProjectBuildTargetToString(job.config.target)));
	AppendCommandLineArgument(commandLine, L"--autolinker-result");
	AppendCommandLineArgument(commandLine, job.resultPath.wstring());
	AppendCommandLineArgument(commandLine, L"--autolinker-startup-timeout");
	AppendCommandLineArgument(commandLine, L"600");
	AppendCommandLineArgument(commandLine, L"--autolinker-invocation-id");
	AppendCommandLineArgument(commandLine, Utf8ToWide(job.id));
	AppendCommandLineArgument(commandLine, L"--autolinker-project-source");
	AppendCommandLineArgument(commandLine, job.variables.sourcePath.wstring());
	AppendCommandLineArgument(commandLine, job.config.staticCompile ? L"--autolinker-static" : L"--autolinker-no-static");
	AppendCommandLineArgument(commandLine, L"--autolinker-hide-window");
	AppendCommandLineArgument(commandLine, L"--autolinker-exit");

	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');
	STARTUPINFOW startup = {};
	startup.cb = sizeof(startup);
	startup.dwFlags = STARTF_USESHOWWINDOW;
	startup.wShowWindow = SW_HIDE;
	PROCESS_INFORMATION process = {};
	const BOOL created = CreateProcessW(
		job.eidePath.c_str(),
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		FALSE,
		CREATE_UNICODE_ENVIRONMENT | CREATE_SUSPENDED,
		nullptr,
		job.variables.projectDir.c_str(),
		&startup,
		&process);
	if (!created) {
		error = "CreateProcessW failed, win32=" + std::to_string(GetLastError());
		return false;
	}

	if (ResumeThread(process.hThread) == static_cast<DWORD>(-1)) {
		const DWORD resumeError = GetLastError();
		TerminateProcess(process.hProcess, 5);
		WaitForSingleObject(process.hProcess, 5000);
		CloseHandle(process.hThread);
		CloseHandle(process.hProcess);
		error = "ResumeThread failed, win32=" + std::to_string(resumeError);
		return false;
	}
	CloseHandle(process.hThread);
	const ULONGLONG waitStartedAt = GetTickCount64();
	std::size_t hiddenWindowEvents = 0;
	DWORD waitResult = WAIT_TIMEOUT;
	for (;;) {
		const ULONGLONG elapsed = GetTickCount64() - waitStartedAt;
		if (elapsed < kHeadlessStartupWindowSuppressionMilliseconds) {
			hiddenWindowEvents += HeadlessCompileRunner::HideIdeWindowsForHeadlessProcess(process.dwProcessId);
		}
		if (elapsed >= kHeadlessCompileTimeoutMilliseconds) {
			break;
		}

		const ULONGLONG remaining = kHeadlessCompileTimeoutMilliseconds - elapsed;
		DWORD waitSlice = elapsed < kHeadlessStartupWindowSuppressionMilliseconds
			? kHeadlessStartupWindowPollMilliseconds
			: 1000;
		if (remaining < waitSlice) {
			waitSlice = static_cast<DWORD>(remaining);
		}
		waitResult = WaitForSingleObject(process.hProcess, waitSlice);
		if (waitResult != WAIT_TIMEOUT) {
			break;
		}
	}
	if (hiddenWindowEvents > 0) {
		Logger::Instance().Write(
			"ProjectBuild",
			"job=" + job.id + " hidden_headless_window_events=" + std::to_string(hiddenWindowEvents));
	}
	if (waitResult == WAIT_TIMEOUT) {
		TerminateProcess(process.hProcess, 5);
		WaitForSingleObject(process.hProcess, 5000);
		CloseHandle(process.hProcess);
		error = "headless compile timed out";
		return false;
	}
	if (waitResult != WAIT_OBJECT_0) {
		const DWORD waitError = GetLastError();
		CloseHandle(process.hProcess);
		error = "wait for headless compile failed, win32=" + std::to_string(waitError);
		return false;
	}

	DWORD processExitCode = 0;
	GetExitCodeProcess(process.hProcess, &processExitCode);
	CloseHandle(process.hProcess);
	nlohmann::json headlessResult;
	if (!ReadHeadlessResult(job.resultPath, processExitCode, headlessResult, error)) return false;
	if (headlessResult.contains("compile_result") && headlessResult["compile_result"].is_object()) {
		outputPath = headlessResult["compile_result"].value("output_path", WideToUtf8(job.variables.outputPath.wstring()));
	}
	if (outputPath.empty()) outputPath = WideToUtf8(job.variables.outputPath.wstring());
	return true;
}

void RunProjectBuildWorker(ProjectBuildJob job)
{
	ProjectBuildJobCleanup cleanup{job.snapshotPath, job.jobDirectory};
	try {
		std::string error;
		if (!RunCommands(job.config.preBuildCommands, job.variables, "pre-build", error)) {
			OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 异步任务失败：") + job.config.name + " - " + error);
			return;
		}

		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 无头编译进程已启动：") + job.config.name + ProjectBuildUtf8(u8"，任务=") + job.id);
		std::string outputPath;
		if (!RunHeadlessCompile(job, outputPath, error)) {
			OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 异步编译失败：") + job.config.name + " - " + error);
			Logger::Instance().Write("ProjectBuild", "job=" + job.id + " failed: " + error);
			return;
		}

		job.variables.outputPath = std::filesystem::path(Utf8ToWide(outputPath));
		if (!RunCommands(job.config.postBuildCommands, job.variables, "post-build", error)) {
			OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 编译后动作失败：") + job.config.name + " - " + error);
			return;
		}
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 异步编译完成：") + outputPath);
		Logger::Instance().Write("ProjectBuild", "job=" + job.id + " complete output=" + outputPath);
	}
	catch (const std::exception& ex) {
		const std::string error = std::string("async project build exception: ") + ex.what();
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 异步任务异常：") + error);
		Logger::Instance().Write("ProjectBuild", "job=" + job.id + " " + error);
	}
	catch (...) {
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 异步任务发生未知异常。"));
		Logger::Instance().Write("ProjectBuild", "job=" + job.id + " unknown exception");
	}
}

} // namespace

ProjectBuildPipelineResult StartProjectBuildPipelineAsync(
	const std::filesystem::path& sourcePath,
	const ProjectBuildConfig& config)
{
	ProjectBuildPipelineResult result;
	try {
		result.stage = "validate";
		if (sourcePath.empty() || !std::filesystem::exists(sourcePath)) {
			result.message = "source file does not exist";
			return result;
		}

		ProjectBuildJob job;
		job.config = config;
		job.variables.projectDir = sourcePath.parent_path();
		job.variables.projectName = WideToUtf8(sourcePath.stem().wstring());
		job.variables.sourcePath = sourcePath;
		job.eidePath = GetCurrentExecutablePath();
		if (job.eidePath.empty()) {
			result.message = "cannot resolve current IDE executable path";
			return result;
		}
		job.variables.programDir = job.eidePath.parent_path();

		std::vector<std::string> unknown;
		const std::string expandedOutput = ExpandProjectBuildVariables(config.outputPath, job.variables, &unknown);
		if (!unknown.empty()) {
			result.message = "output path contains unknown variables: " + JoinUnknownVariables(unknown);
			return result;
		}
		const std::wstring outputWide = Utf8ToWide(expandedOutput);
		if (outputWide.empty()) {
			result.message = "output path is empty or not valid UTF-8";
			return result;
		}
		job.variables.outputPath = std::filesystem::path(outputWide);
		if (job.variables.outputPath.is_relative()) {
			job.variables.outputPath = job.variables.projectDir / job.variables.outputPath;
		}
		job.variables.outputPath = job.variables.outputPath.lexically_normal();
		result.outputPath = WideToUtf8(job.variables.outputPath.wstring());

		const unsigned long long counter = g_projectBuildJobCounter.fetch_add(1) + 1;
		job.id = "project-build-" + std::to_string(GetCurrentProcessId()) + "-" +
			std::to_string(GetTickCount64()) + "-" + std::to_string(counter);
		std::error_code pathError;
		job.jobDirectory = std::filesystem::temp_directory_path(pathError) /
			L"AutoLinker" / L"project-build" / Utf8ToWide(job.id);
		if (pathError) {
			result.stage = "snapshot";
			result.message = "cannot resolve system temp directory";
			return result;
		}
		std::filesystem::create_directories(job.jobDirectory, pathError);
		if (pathError) {
			result.stage = "snapshot";
			result.message = "cannot create project build task directory";
			return result;
		}
		job.snapshotPath = job.variables.projectDir /
			std::filesystem::path(L"~autolinker-" + Utf8ToWide(job.id) + L".e");
		job.resultPath = job.jobDirectory / L"result.json";

		result.stage = "snapshot";
		size_t bytesWritten = 0;
		std::string snapshotError;
		std::string snapshotTrace;
		if (!WriteProjectSnapshot(job.snapshotPath, bytesWritten, snapshotError, snapshotTrace)) {
			std::filesystem::remove(job.snapshotPath, pathError);
			pathError.clear();
			std::filesystem::remove_all(job.jobDirectory, pathError);
			result.message = snapshotError.empty() ? "serialize current project failed" : snapshotError;
			Logger::Instance().Write("ProjectBuild", "snapshot failed: " + snapshotTrace);
			return result;
		}
		Logger::Instance().Write(
			"ProjectBuild",
			"job=" + job.id +
			" source=" + WideToUtf8(job.variables.sourcePath.wstring()) +
			" snapshot=" + WideToUtf8(job.snapshotPath.wstring()) +
			" snapshot_bytes=" + std::to_string(bytesWritten) +
			" trace=" + snapshotTrace);

		result.stage = "queued";
		result.message = "project build queued";
		result.ok = true;
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 已生成当前工程快照并提交异步编译：") + config.name + ProjectBuildUtf8(u8"，任务=") + job.id);
		std::thread(RunProjectBuildWorker, std::move(job)).detach();
		return result;
	}
	catch (const std::exception& ex) {
		result.ok = false;
		if (result.stage.empty()) result.stage = "exception";
		result.message = std::string("start async project build exception: ") + ex.what();
		Logger::Instance().Write("ProjectBuild", result.message);
		return result;
	}
	catch (...) {
		result.ok = false;
		if (result.stage.empty()) result.stage = "exception";
		result.message = "start async project build exception: unknown exception";
		Logger::Instance().Write("ProjectBuild", result.message);
		return result;
	}
}
