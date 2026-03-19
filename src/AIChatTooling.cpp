#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"

#include <Windows.h>

#include <chrono>
#include <cctype>
#include <condition_variable>
#include <mutex>
#include <string>

#include "..\\thirdparty\\json.hpp"

namespace {std::string TrimAsciiCopy(const std::string& text)
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

std::string LocalFromWide(const wchar_t* text)
{
	if (text == nullptr || *text == L'\0') {
		return std::string();
	}
	const int size = WideCharToMultiByte(CP_ACP, 0, text, -1, nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return std::string();
	}
	std::string out(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(CP_ACP, 0, text, -1, out.data(), size, nullptr, nullptr) <= 0) {
		return std::string();
	}
	if (!out.empty() && out.back() == '\0') {
		out.pop_back();
	}
	return out;
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
bool RequestToolExecutionFromMainThread(
	const std::string& toolName,
	const std::string& argumentsJson,
	std::string& outResultJson,
	bool& outOk)
{
	outResultJson.clear();
	outOk = false;
	const HWND mainWindow = GetAIChatMainWindowForTooling();
	const UINT toolExecMessage = GetAIChatToolExecMessageForTooling();
	if (mainWindow == nullptr || !IsWindow(mainWindow) || toolExecMessage == 0) {
		return false;
	}

	ToolExecutionRequest request;
	request.toolName = toolName;
	request.argumentsJson = argumentsJson;
	if (PostMessage(mainWindow, toolExecMessage, 0, reinterpret_cast<LPARAM>(&request)) == FALSE) {
		return false;
	}

	std::unique_lock<std::mutex> lock(request.mutex);
	if (!request.cv.wait_for(lock, std::chrono::minutes(20), [&request]() { return request.done; })) {
		return false;
	}

	outResultJson = request.resultJson;
	outOk = request.ok;
	return true;
}

} // namespace
std::string ExecuteToolCall(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;

	if (toolName == "request_code_edit") {
		std::string title = LocalFromWide(L"AI\u4ee3\u7801\u7f16\u8f91");
		std::string hint;
		std::string initialCode;
		try {
			const nlohmann::json args = nlohmann::json::parse(argumentsJson);
			if (args.contains("title") && args["title"].is_string()) {
				title = Utf8ToLocalText(args["title"].get<std::string>());
			}
			if (args.contains("hint") && args["hint"].is_string()) {
				hint = Utf8ToLocalText(args["hint"].get<std::string>());
			}
			if (args.contains("initial_code") && args["initial_code"].is_string()) {
				initialCode = Utf8ToLocalText(args["initial_code"].get<std::string>());
			}
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		std::string editedCode;
		if (!RequestCodeEditForTooling(title, hint, initialCode, editedCode)) {
			return R"({"ok":false,"cancelled":true})";
		}

		nlohmann::json r;
		r["ok"] = true;
		r["code"] = LocalToUtf8Text(editedCode);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_current_page_code" ||
		toolName == "get_current_page_info" ||
		toolName == "list_imported_modules" ||
		toolName == "list_support_libraries" ||
		toolName == "get_support_library_info" ||
		toolName == "search_support_library_info" ||
		toolName == "get_module_public_info" ||
		toolName == "search_module_public_info" ||
		toolName == "list_program_items" ||
		toolName == "get_program_item_code" ||
		toolName == "switch_to_program_item_page" ||
		toolName == "search_project_keyword" ||
		toolName == "jump_to_search_result" ||
		toolName == "compile_with_output_path") {
		std::string resultJson;
		if (!RequestToolExecutionFromMainThread(toolName, argumentsJson, resultJson, outOk)) {
			return R"({"ok":false,"error":"main thread tool execution failed"})";
		}
		return resultJson;
	}

	nlohmann::json r;
	r["ok"] = false;
	r["error"] = "unknown tool: " + toolName;
	return Utf8ToLocalText(r.dump());
}

