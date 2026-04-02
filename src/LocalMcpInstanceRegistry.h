#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace LocalMcpInstanceRegistry {

// 本机 MCP 实例登记信息
struct InstanceRecord {
	std::string instanceId; // 实例唯一标识
	unsigned long processId = 0; // 进程 ID
	std::string processPath; // 易语言主程序路径
	std::string processName; // 易语言主程序名
	int port = 0; // MCP 监听端口
	std::string endpoint; // MCP 完整地址
	std::string sourceFilePathHint; // 当前源码路径提示
	std::string pageNameHint; // 当前页面名提示
	std::string pageTypeHint; // 当前页面类型提示
	std::uint64_t lastSeenUnixMs = 0; // 最后心跳时间
};

// 获取本机实例注册文件路径
std::string GetRegistryFilePath();

// 写入或刷新当前实例登记信息
bool UpsertCurrentInstance(const InstanceRecord& record, std::string* outError = nullptr);

// 删除指定实例的登记信息
bool RemoveCurrentInstance(const std::string& instanceId, std::string* outError = nullptr);

// 读取当前有效的实例列表
bool LoadInstances(std::vector<InstanceRecord>& outRecords, std::string* outError = nullptr);

} // namespace LocalMcpInstanceRegistry