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

// 获取固定 MCP 网关地址
std::string GetGatewayEndpoint();

// 判断当前实例是否持有固定 MCP 网关
bool IsGatewayOwner();

// 构建外部 MCP 首次刷新门禁的内部自测报告。
std::string BuildWorkspaceRefreshGateSelfTestJson();

// 构建多实例会话路由与网关接管的内部自测报告
std::string BuildMultiInstanceRoutingSelfTestJson();

// 更新当前实例的易语言上下文提示
void UpdateInstanceHints(
	const std::string& sourceFilePath,
	const std::string& pageName,
	const std::string& pageType);

} // namespace LocalMcpServer
