#pragma once

#include <map>
#include <optional>
#include <string>
#include <vector>

class ConfigManager;

namespace GameAnalyticsClient {

// Remote Config 当前快照，仅保存在进程内供业务代码读取。
struct RemoteConfigSnapshot {
	bool ready = false;
	bool requestInFlight = false;
	int httpStatus = 0;
	long long serverTs = 0;
	std::string error;
	std::string configsHash;
	std::string abId;
	std::string abVariantId;
	std::map<std::string, std::string> values;
	std::vector<std::string> valueTypes;
};

// 初始化 GameAnalytics 后台客户端。
void Initialize(ConfigManager* configManager);

// 关闭后台客户端，最多短暂等待当前异步任务结束。
void Shutdown();

// 判断后台客户端是否正在运行。
bool IsRunning();

// 异步刷新 Remote Config。
void RefreshRemoteConfigsAsync();

// 读取 Remote Config 当前快照。
RemoteConfigSnapshot GetRemoteConfigs();

// 读取指定 Remote Config 值。
std::optional<std::string> GetRemoteConfigValue(const std::string& key);

// 构建不依赖网络的自检报告。
std::string BuildSelfTestReportJson();

} // namespace GameAnalyticsClient
