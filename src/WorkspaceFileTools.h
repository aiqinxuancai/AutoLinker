#pragma once

#include <string>

// 基于 WorkspaceMirror 的统一文件读取、搜索和列出工具。
namespace WorkspaceFileTools {

// 执行 read_file/list_files/search_code 工具，返回本地编码 JSON 字符串。
bool CanHandleTool(const std::string& toolName);
std::string ExecuteTool(const std::string& toolName, const std::string& argumentsJson, bool& outOk);

} // namespace WorkspaceFileTools
