#pragma once

#include <functional>
#include <string>

// 基于 WorkspaceMirror 的统一文件读取、搜索和列出工具。
namespace WorkspaceFileTools {

// 执行 read_file/read_files/read_code_item/list_files/search_code 工具，返回本地编码 JSON 字符串。
bool CanHandleTool(const std::string& toolName);
std::string ExecuteTool(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk,
	const std::function<bool()>& cancelCallback = {});

// 构建分页与精确省略范围元数据的内部自测报告。
std::string BuildPaginationSelfTestJson();

} // namespace WorkspaceFileTools
