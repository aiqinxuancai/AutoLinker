#pragma once

#include <string>

namespace LocalMcpServer {

// 初始化本地 MCP 服务
void Initialize();

// 关闭本地 MCP 服务
void Shutdown();

// 判断本地 MCP 是否正在运行
bool IsRunning();

// 获取当前绑定端口
int GetBoundPort();

// 获取当前实例唯一标识
std::string GetInstanceId();

// 获取当前 MCP 完整地址
std::string GetEndpoint();

// 更新当前实例的易语言上下文提示
void UpdateInstanceHints(
	const std::string& sourceFilePath,
	const std::string& pageName,
	const std::string& pageType);

} // namespace LocalMcpServer