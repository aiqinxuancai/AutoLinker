#include <cstdlib>
#include <algorithm>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <format>
#include <iostream>
#include <string>
#include <thread>
#include <vector>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"

#include "..\\src\\AutoLinkerTestApi.h"

namespace {

void PrintUsage();

struct HeadlessLauncherOptions {
	std::string eExePath;
	std::string projectPath;
	std::string outputPath;
	std::string resultPath;
	std::string target = "auto";
	bool staticCompile = false;
	bool hideWindow = true;
	int timeoutSeconds = 120;
};

struct CapturedDialog {
	std::string kind = "info";
	std::string caption;
	std::string text;
	std::vector<std::string> listItems;
	std::vector<std::string> supportLibraries;
};

struct DialogEnumContext {
	DWORD processId = 0;
	std::vector<CapturedDialog>* dialogs = nullptr;
};

std::wstring AnsiToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::wstring();
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}
	const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return std::string();
	}
	std::string utf8(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), utf8.data(), size, nullptr, nullptr) <= 0) {
		return std::string();
	}
	return utf8;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return AnsiToWide(text);
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return AnsiToWide(text);
	}
	return wide;
}

std::wstring QuoteCommandLineArgWide(const std::wstring& arg)
{
	std::wstring quoted = L"\"";
	for (wchar_t ch : arg) {
		if (ch == L'"') {
			quoted += L"\\\"";
		}
		else {
			quoted.push_back(ch);
		}
	}
	quoted += L"\"";
	return quoted;
}

std::filesystem::path MakePathFromText(const std::string& text)
{
	const std::wstring wide = Utf8ToWide(text);
	if (!wide.empty()) {
		return std::filesystem::path(wide);
	}
	return std::filesystem::path(text);
}

std::string DefaultHeadlessResultPath(const std::string& outputPath)
{
	const std::filesystem::path path = MakePathFromText(outputPath);
	if (path.empty()) {
		return "autolinker-headless-result.json";
	}
	return WideToUtf8(path.wstring() + L".headless.json");
}

std::filesystem::path GetHeadlessRequestPath(const std::string& eExePath)
{
	const std::filesystem::path exePath = MakePathFromText(eExePath);
	const std::filesystem::path basePath = exePath.parent_path().empty()
		? std::filesystem::current_path()
		: exePath.parent_path();
	return basePath / "AutoLinker" / "Log" / "headless_compile_request.json";
}

std::wstring GetWindowTextWideLocal(HWND hWnd)
{
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return std::wstring();
	}
	const int len = GetWindowTextLengthW(hWnd);
	if (len <= 0) {
		return std::wstring();
	}
	std::wstring text(static_cast<size_t>(len) + 1, L'\0');
	const int copied = GetWindowTextW(hWnd, text.data(), len + 1);
	if (copied <= 0) {
		return std::wstring();
	}
	text.resize(static_cast<size_t>(copied));
	return text;
}

std::wstring GetWindowClassWideLocal(HWND hWnd)
{
	wchar_t className[128] = {};
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return std::wstring();
	}
	if (GetClassNameW(hWnd, className, static_cast<int>(sizeof(className) / sizeof(className[0]))) <= 0) {
		return std::wstring();
	}
	return className;
}

struct ChildTextContext {
	std::vector<std::wstring> texts;
	std::vector<std::wstring> listItems;
};

std::vector<std::wstring> ReadListBoxItems(HWND hWnd)
{
	std::vector<std::wstring> items;
	const LRESULT count = SendMessageW(hWnd, LB_GETCOUNT, 0, 0);
	if (count <= 0 || count == LB_ERR) {
		return items;
	}

	for (int i = 0; i < count; ++i) {
		const LRESULT len = SendMessageW(hWnd, LB_GETTEXTLEN, static_cast<WPARAM>(i), 0);
		if (len == LB_ERR) {
			continue;
		}
		std::wstring item(static_cast<size_t>(len) + 1, L'\0');
		const LRESULT copied = SendMessageW(hWnd, LB_GETTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(item.data()));
		if (copied == LB_ERR || copied <= 0) {
			continue;
		}
		item.resize(static_cast<size_t>(copied));
		if (std::find(items.begin(), items.end(), item) == items.end()) {
			items.push_back(std::move(item));
		}
	}
	return items;
}

BOOL CALLBACK EnumChildTextProc(HWND hWnd, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<ChildTextContext*>(lParam);
	if (ctx == nullptr || hWnd == nullptr || !IsWindow(hWnd)) {
		return TRUE;
	}
	const std::wstring className = GetWindowClassWideLocal(hWnd);
	if (_wcsicmp(className.c_str(), L"ListBox") == 0) {
		for (auto& item : ReadListBoxItems(hWnd)) {
			if (!item.empty() && std::find(ctx->listItems.begin(), ctx->listItems.end(), item) == ctx->listItems.end()) {
				ctx->listItems.push_back(std::move(item));
			}
		}
		return TRUE;
	}
	if (_wcsicmp(className.c_str(), L"Static") != 0 &&
		_wcsicmp(className.c_str(), L"Edit") != 0) {
		return TRUE;
	}
	std::wstring text = GetWindowTextWideLocal(hWnd);
	if (!text.empty() && std::find(ctx->texts.begin(), ctx->texts.end(), text) == ctx->texts.end()) {
		ctx->texts.push_back(std::move(text));
	}
	return TRUE;
}

std::wstring CollectDialogText(HWND hWnd, std::vector<std::wstring>& listItems)
{
	ChildTextContext ctx;
	EnumChildWindows(hWnd, EnumChildTextProc, reinterpret_cast<LPARAM>(&ctx));
	listItems = ctx.listItems;
	std::wstring merged;
	for (const auto& text : ctx.texts) {
		if (text.empty()) {
			continue;
		}
		if (!merged.empty()) {
			merged += L"\r\n";
		}
		merged += text;
	}
	for (const auto& item : ctx.listItems) {
		if (item.empty() || std::find(ctx.texts.begin(), ctx.texts.end(), item) != ctx.texts.end()) {
			continue;
		}
		if (!merged.empty()) {
			merged += L"\r\n";
		}
		merged += item;
	}
	return merged;
}

bool HasEcomModuleListItem(const std::vector<std::wstring>& listItems)
{
	for (const auto& item : listItems) {
		if (item.find(L".ec") != std::wstring::npos || item.find(L".EC") != std::wstring::npos) {
			return true;
		}
	}
	return false;
}

bool TextLooksLikeMissingEcomModule(const std::wstring& text)
{
	return text.find(L"易模块文件已经无法找到") != std::wstring::npos ||
		text.find(L"模块文件已经无法找到") != std::wstring::npos ||
		(text.find(L"易模块文件") != std::wstring::npos && text.find(L"不存在") != std::wstring::npos) ||
		text.find(L".ec") != std::wstring::npos ||
		text.find(L".EC") != std::wstring::npos;
}

bool TextLooksLikeMissingSupportLibrary(const std::wstring& text)
{
	return text.find(L"不能载入支持库") != std::wstring::npos ||
		text.find(L"支持库不能被载入") != std::wstring::npos;
}

std::vector<std::wstring> ExtractMissingSupportLibraries(const std::wstring& text)
{
	std::vector<std::wstring> rows;
	size_t begin = 0;
	while (begin <= text.size()) {
		size_t end = text.find_first_of(L"\r\n", begin);
		if (end == std::wstring::npos) {
			end = text.size();
		}
		std::wstring line = text.substr(begin, end - begin);
		if (line.find(L"不能载入支持库") != std::wstring::npos &&
			std::find(rows.begin(), rows.end(), line) == rows.end()) {
			rows.push_back(std::move(line));
		}
		begin = end + 1;
		while (begin < text.size() && (text[begin] == L'\r' || text[begin] == L'\n')) {
			++begin;
		}
		if (end == text.size()) {
			break;
		}
	}
	if (rows.empty() && TextLooksLikeMissingSupportLibrary(text)) {
		rows.push_back(text);
	}
	return rows;
}

std::string ClassifyDialog(const std::wstring& text, const std::vector<std::wstring>& listItems)
{
	if (TextLooksLikeMissingSupportLibrary(text)) {
		return "missing_support_library";
	}
	if (TextLooksLikeMissingEcomModule(text) || HasEcomModuleListItem(listItems)) {
		return "missing_ecom_module";
	}
	return "info";
}

bool ContainsDialog(const std::vector<CapturedDialog>& dialogs, const CapturedDialog& dialog)
{
	return std::find_if(dialogs.begin(), dialogs.end(), [&dialog](const CapturedDialog& item) {
		return item.kind == dialog.kind && item.caption == dialog.caption && item.text == dialog.text;
	}) != dialogs.end();
}

bool HasBlockingDialog(const std::vector<CapturedDialog>& dialogs)
{
	return std::find_if(dialogs.begin(), dialogs.end(), [](const CapturedDialog& dialog) {
		return dialog.kind != "info";
	}) != dialogs.end();
}

BOOL CALLBACK EnumLauncherDialogProc(HWND hWnd, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<DialogEnumContext*>(lParam);
	if (ctx == nullptr || ctx->dialogs == nullptr || hWnd == nullptr || !IsWindow(hWnd) || !IsWindowVisible(hWnd)) {
		return TRUE;
	}
	DWORD pid = 0;
	GetWindowThreadProcessId(hWnd, &pid);
	if (pid != ctx->processId || _wcsicmp(GetWindowClassWideLocal(hWnd).c_str(), L"#32770") != 0) {
		return TRUE;
	}
	CapturedDialog dialog;
	std::vector<std::wstring> listItems;
	const std::wstring text = CollectDialogText(hWnd, listItems);
	dialog.kind = ClassifyDialog(text, listItems);
	dialog.caption = WideToUtf8(GetWindowTextWideLocal(hWnd));
	dialog.text = WideToUtf8(text);
	for (const auto& item : listItems) {
		dialog.listItems.push_back(WideToUtf8(item));
	}
	for (const auto& item : ExtractMissingSupportLibraries(text)) {
		dialog.supportLibraries.push_back(WideToUtf8(item));
	}
	if (!ContainsDialog(*ctx->dialogs, dialog)) {
		ctx->dialogs->push_back(dialog);
	}
	PostMessageW(hWnd, WM_COMMAND, static_cast<WPARAM>(IDCANCEL), 0);
	PostMessageW(hWnd, WM_CLOSE, 0, 0);
	return TRUE;
}

void CaptureAndDismissProcessDialogs(DWORD processId, std::vector<CapturedDialog>& dialogs)
{
	DialogEnumContext ctx;
	ctx.processId = processId;
	ctx.dialogs = &dialogs;
	EnumWindows(EnumLauncherDialogProc, reinterpret_cast<LPARAM>(&ctx));
}

bool WriteTextFile(const std::filesystem::path& path, const std::string& text)
{
	std::error_code ec;
	const std::filesystem::path parent = path.parent_path();
	if (!parent.empty()) {
		std::filesystem::create_directories(parent, ec);
	}
	std::ofstream out(path, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		return false;
	}
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	return true;
}

std::string ReadTextFile(const std::filesystem::path& path)
{
	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		return std::string();
	}
	return std::string(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

void PrintHeadlessSummary(const nlohmann::json& result)
{
	const bool ok = result.value("ok", false);
	std::ostream& out = ok ? std::cout : std::cerr;
	out << (ok ? "[AutoLinkerLauncher] compile success" : "[AutoLinkerLauncher] compile failed") << std::endl;
	const std::string error = result.value("error", std::string());
	if (!ok && !error.empty()) {
		out << "error: " << error << std::endl;
	}
	if (!ok && result.contains("launcher_process_exit_code")) {
		const DWORD exitCode = result.value("launcher_process_exit_code", 0u);
		out << std::format("process_exit_code: {} (0x{:08X})", exitCode, exitCode) << std::endl;
	}
	if (result.contains("resolved_target")) {
		out << "target: " << result.value("resolved_target", std::string()) << std::endl;
	}
	if (result.contains("compile_result") && result["compile_result"].is_object()) {
		const auto& compileResult = result["compile_result"];
		out << "output: " << compileResult.value("output_path", std::string()) << std::endl;
		if (!ok) {
			out << "error_location: page=" << compileResult.value("caret_page_name", std::string())
				<< " type=" << compileResult.value("caret_page_type", std::string())
				<< " row=" << compileResult.value("caret_row", -1) << std::endl;
			const std::string line = compileResult.value("caret_line_text", std::string());
			if (!line.empty()) {
				out << "error_line: " << line << std::endl;
			}
		}
		const std::string outputText = compileResult.value("output_window_text", std::string());
		if (!outputText.empty()) {
			out << "[IDE output]" << std::endl << outputText << std::endl;
		}
	}
	const auto printBoxes = [](const nlohmann::json& boxes, const char* label) {
		if (!boxes.is_array()) {
			return;
		}
		for (const auto& box : boxes) {
			const std::string kind = box.value("kind", std::string("info"));
			std::cerr << label;
			if (kind == "missing_ecom_module") {
				std::cerr << "[MissingModule] ";
			}
			else if (kind == "missing_support_library") {
				std::cerr << "[MissingSupportLibrary] ";
			}
			else if (kind == "info") {
				std::cerr << "[Info] ";
			}
			std::cerr << box.value("caption", std::string()) << std::endl;
			const std::string text = box.value("text", std::string());
			if (!text.empty()) {
				std::cerr << text << std::endl;
			}
			if (box.contains("support_libraries") && box["support_libraries"].is_array() && !box["support_libraries"].empty()) {
				std::cerr << "support_libraries:" << std::endl;
				for (const auto& item : box["support_libraries"]) {
					if (item.is_string()) {
						std::cerr << "  " << item.get<std::string>() << std::endl;
					}
				}
			}
			if (box.contains("list_items") && box["list_items"].is_array() && !box["list_items"].empty()) {
				std::cerr << "list_items:" << std::endl;
				for (const auto& item : box["list_items"]) {
					if (item.is_string()) {
						std::cerr << "  " << item.get<std::string>() << std::endl;
					}
				}
			}
		}
	};
	if (result.contains("launcher_message_boxes")) {
		printBoxes(result["launcher_message_boxes"], "[AutoLinkerLauncher][MessageBox] ");
	}
	if (result.contains("ide_message_boxes")) {
		printBoxes(result["ide_message_boxes"], "[AutoLinker][MessageBox] ");
	}
}

int RunHeadlessCompile(int argc, char* argv[])
{
	if (argc < 5) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	HeadlessLauncherOptions options;
	options.eExePath = argv[2];
	options.projectPath = argv[3];
	options.outputPath = argv[4];
	options.resultPath = DefaultHeadlessResultPath(options.outputPath);

	for (int i = 5; i < argc; ++i) {
		const std::string arg = argv[i];
		if (arg == "--static") {
			options.staticCompile = true;
		}
		else if (arg == "--show-window") {
			options.hideWindow = false;
		}
		else if (arg == "--target" && i + 1 < argc) {
			options.target = argv[++i];
		}
		else if (arg == "--result" && i + 1 < argc) {
			options.resultPath = argv[++i];
		}
		else if (arg == "--timeout" && i + 1 < argc) {
			options.timeoutSeconds = (std::max)(1, std::atoi(argv[++i]));
		}
		else {
			std::cerr << "unknown headless-compile option: " << arg << std::endl;
			return EXIT_FAILURE;
		}
	}

	nlohmann::json request = {
		{"enabled", true},
		{"target", options.target},
		{"static_compile", options.staticCompile},
		{"output_path", options.outputPath},
		{"result_path", options.resultPath},
		{"startup_timeout_seconds", options.timeoutSeconds},
		{"hide_window", options.hideWindow},
		{"exit_after_compile", true}
	};

	const std::filesystem::path requestPath = GetHeadlessRequestPath(options.eExePath);
	if (!WriteTextFile(requestPath, request.dump(2))) {
		std::cerr << "write headless request failed: " << WideToUtf8(requestPath.wstring()) << std::endl;
		return EXIT_FAILURE;
	}

	STARTUPINFOW si = {};
	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = options.hideWindow ? SW_HIDE : SW_SHOW;
	PROCESS_INFORMATION pi = {};

	const std::wstring exePath = Utf8ToWide(options.eExePath);
	const std::wstring projectPath = Utf8ToWide(options.projectPath);
	std::wstring commandLine = QuoteCommandLineArgWide(exePath) + L" " + QuoteCommandLineArgWide(projectPath);
	const std::filesystem::path projectFsPath = MakePathFromText(options.projectPath);
	const std::wstring workingDir = projectFsPath.parent_path().empty()
		? std::filesystem::current_path().wstring()
		: projectFsPath.parent_path().wstring();

	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');
	const BOOL created = CreateProcessW(
		exePath.c_str(),
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		FALSE,
		0,
		nullptr,
		workingDir.c_str(),
		&si,
		&pi);

	if (created == FALSE) {
		std::error_code ec;
		std::filesystem::remove(requestPath, ec);
		std::cerr << "CreateProcessW failed, error=" << GetLastError() << std::endl;
		return EXIT_FAILURE;
	}

	std::vector<CapturedDialog> dialogs;
	const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(options.timeoutSeconds + 30);
	bool timedOut = false;
	for (;;) {
		const DWORD waitResult = WaitForSingleObject(pi.hProcess, 200);
		CaptureAndDismissProcessDialogs(pi.dwProcessId, dialogs);
		if (waitResult == WAIT_OBJECT_0) {
			break;
		}
		if (std::chrono::steady_clock::now() >= deadline) {
			timedOut = true;
			TerminateProcess(pi.hProcess, 5);
			break;
		}
	}

	DWORD processExitCode = 0;
	GetExitCodeProcess(pi.hProcess, &processExitCode);
	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);
	{
		std::error_code ec;
		std::filesystem::remove(requestPath, ec);
	}

	const std::filesystem::path resultPath = MakePathFromText(options.resultPath);
	nlohmann::json result;
	const std::string resultText = ReadTextFile(resultPath);
	const bool resultLoadedFromFile = !resultText.empty();
	if (!resultText.empty()) {
		try {
			result = nlohmann::json::parse(resultText);
		}
		catch (const std::exception& ex) {
			result = {
				{"ok", false},
				{"error", std::string("parse_headless_result_failed: ") + ex.what()},
				{"raw_result", resultText}
			};
		}
	}
	else {
		result = {
			{"ok", false},
			{"mode", "autolinker_headless_launcher"},
			{"error", timedOut ? "headless_process_timeout" : "headless_result_missing"},
			{"process_exit_code", processExitCode},
			{"request", request},
			{"request_file", WideToUtf8(requestPath.wstring())}
		};
	}

	if (!dialogs.empty()) {
		nlohmann::json rows = nlohmann::json::array();
		for (const auto& dialog : dialogs) {
			nlohmann::json row = {
				{"kind", dialog.kind},
				{"caption", dialog.caption},
				{"text", dialog.text}
			};
			if (!dialog.listItems.empty()) {
				row["list_items"] = dialog.listItems;
			}
			if (!dialog.supportLibraries.empty()) {
				row["support_libraries"] = dialog.supportLibraries;
			}
			rows.push_back(std::move(row));
		}
		result["launcher_message_boxes"] = std::move(rows);
		if (HasBlockingDialog(dialogs) && !result.value("ok", false) && !timedOut && (!resultLoadedFromFile || !result.contains("error"))) {
			result["error"] = "ide_startup_message_box_captured";
		}
	}
	result["launcher_process_exit_code"] = processExitCode;
	result["launcher_timed_out"] = timedOut;
	result["launcher_request_file"] = WideToUtf8(requestPath.wstring());
	WriteTextFile(resultPath, result.dump(2));
	PrintHeadlessSummary(result);
	std::cout << "result: " << options.resultPath << std::endl;

	if (result.value("ok", false)) {
		return EXIT_SUCCESS;
	}
	return result.value("exit_code", timedOut ? 5 : 1);
}

std::string JoinArguments(int startIndex, int argc, char* argv[])
{
	std::string text;
	for (int i = startIndex; i < argc; ++i) {
		if (!text.empty()) {
			text.push_back(' ');
		}
		text.append(argv[i]);
	}
	return text;
}

int PrintStringResult(const char* label, int result, const char* text)
{
	if (result >= 0) {
		std::cout << label << ": " << text << std::endl;
		return EXIT_SUCCESS;
	}

	if (result == AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL) {
		std::cerr << label << " failed: buffer too small" << std::endl;
	}
	else {
		std::cerr << label << " failed: invalid argument" << std::endl;
	}
	return EXIT_FAILURE;
}

int RunSmokeTest()
{
	char buffer[512] = {};
	int compareResult = 0;

	if (!AutoLinkerTest_CompareVersion("1.2.3", "1.2.0", &compareResult)) {
		std::cerr << "version-compare failed" << std::endl;
		return EXIT_FAILURE;
	}

	std::cout << "version-compare: " << compareResult << std::endl;

	int result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("version-text", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	const char* linkerCommand = "link /out:\"D:\\demo\\AutoLinkerTest.exe\" \"D:\\demo\\main.obj\" \"D:\\deps\\static_lib\\krnln_static.lib\"";

	result = AutoLinkerTest_GetLinkerOutFileName(linkerCommand, buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("linker-out", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	result = AutoLinkerTest_GetLinkerKrnlnFileName(linkerCommand, buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("linker-krnln", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	result = AutoLinkerTest_ExtractBetweenDashes("before - middle - after", buffer, static_cast<int>(sizeof(buffer)));
	return PrintStringResult("extract-between-dashes", result, buffer);
}

int RunVersionCompare(const std::string& left, const std::string& right)
{
	int compareResult = 0;
	if (!AutoLinkerTest_CompareVersion(left.c_str(), right.c_str(), &compareResult)) {
		std::cerr << "version-compare failed" << std::endl;
		return EXIT_FAILURE;
	}

	std::cout << compareResult << std::endl;
	return EXIT_SUCCESS;
}

int RunStringCommand(const std::string& commandName, const std::string& input)
{
	char buffer[524288] = {};
	int result = AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;

	if (commandName == "linker-out") {
		result = AutoLinkerTest_GetLinkerOutFileName(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "linker-krnln") {
		result = AutoLinkerTest_GetLinkerKrnlnFileName(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "between-dashes") {
		result = AutoLinkerTest_ExtractBetweenDashes(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "version-text") {
		result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "module-local-dump") {
		result = AutoLinkerTest_DumpLocalModulePublicInfo(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}

	if (result < 0) {
		return PrintStringResult(commandName.c_str(), result, buffer);
	}

	std::cout << buffer << std::endl;
	return EXIT_SUCCESS;
}

void PrintUsage()
{
	std::cout << "AutoLinkerTest commands:" << std::endl;
	std::cout << "  AutoLinkerTest" << std::endl;
	std::cout << "  AutoLinkerTest version-compare <left> <right>" << std::endl;
	std::cout << "  AutoLinkerTest linker-out <link-command>" << std::endl;
	std::cout << "  AutoLinkerTest linker-krnln <link-command>" << std::endl;
	std::cout << "  AutoLinkerTest between-dashes <text>" << std::endl;
	std::cout << "  AutoLinkerTest version-text" << std::endl;
	std::cout << "  AutoLinkerTest module-local-dump <module-path>" << std::endl;
	std::cout << "  AutoLinkerTest e2txt-generate <input-path> <output-path>" << std::endl;
	std::cout << "  AutoLinkerTest e2txt-restore <input-path> <output-path>" << std::endl;
	std::cout << "  AutoLinkerTest bundle-digest-compare <input.e> <input-dir>" << std::endl;
	std::cout << "  AutoLinkerTest headless-compile <e.exe> <input.e> <output> [--target auto|win_exe|win_console_exe|win_dll|ecom] [--static] [--result path] [--timeout seconds]" << std::endl;
}

}

int main(int argc, char* argv[])
{
	if (argc == 1) {
		return RunSmokeTest();
	}

	const std::string commandName = argv[1];
	if (commandName == "headless-compile") {
		return RunHeadlessCompile(argc, argv);
	}

	if (commandName == "version-compare") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunVersionCompare(argv[2], argv[3]);
	}

	if (commandName == "linker-out" || commandName == "linker-krnln" || commandName == "between-dashes" || commandName == "module-local-dump") {
		if (argc < 3) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunStringCommand(commandName, JoinArguments(2, argc, argv));
	}

	if (commandName == "version-text") {
		if (argc != 2) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunStringCommand(commandName, "");
	}

	if (commandName == "e2txt-generate") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_GenerateE2Txt(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "e2txt-restore") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_RestoreE2Txt(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "bundle-digest-compare") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_CompareBundleDigest(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	PrintUsage();
	return EXIT_FAILURE;
}
