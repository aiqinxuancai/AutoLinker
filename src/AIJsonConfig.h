// AIJsonConfig.h
// 管理 AIConfig.json 中的端点、端点分组和全局 AI 设置。
#pragma once

#ifndef AI_JSON_CONFIG_H
#define AI_JSON_CONFIG_H

#include <filesystem>
#include <map>
#include <string>
#include <vector>

// AI 端点快照，字符串均使用调用方本地编码。
struct AIJsonConfigEndpointSnapshot {
	std::string id;
	std::string name;
	std::map<std::string, std::string> values;
};

// AI 端点分组快照，endpointIds 的顺序就是故障转移顺序。
struct AIJsonConfigEndpointGroupSnapshot {
	std::string id;
	std::string name;
	std::vector<std::string> endpointIds;
};

// 唯一启用的 AI 目标。
struct AIJsonConfigActiveTargetSnapshot {
	std::string type; // endpoint | group
	std::string id;
};

// 旧“配置组”接口的兼容别名；v2 中每个旧配置对应一个端点。
using AIJsonConfigProfileSnapshot = AIJsonConfigEndpointSnapshot;

class AIJsonConfig {
public:
	AIJsonConfig();
	// 仅供无 IDE 测试使用，直接读写指定配置文件。
	explicit AIJsonConfig(const std::filesystem::path& filePath);

	// 获取当前活动目标首个端点的值（UTF-8 编码）。
	std::string getValue(const std::string& key) const;
	// 获取端点值并转换为本地编码（ANSI/GBK）。
	std::string getValueLocal(const std::string& key) const;
	// 获取全局配置值（UTF-8 编码）。
	std::string getGlobalValue(const std::string& key) const;
	// 获取全局配置值并转换为本地编码（ANSI/GBK）。
	std::string getGlobalValueLocal(const std::string& key) const;

	// 设置当前活动目标首个端点的值（本地编码）。
	void setValue(const std::string& key, const std::string& localValue);
	void setValues(const std::map<std::string, std::string>& localPairs);
	void removeValues(const std::vector<std::string>& keys);
	void setGlobalValues(const std::map<std::string, std::string>& localPairs);

	bool hasAnyData() const;
	bool hasKey(const std::string& key) const;

	// 获取 v2 端点、分组和活动目标（本地编码）。
	std::vector<AIJsonConfigEndpointSnapshot> getEndpointsLocal() const;
	std::vector<AIJsonConfigEndpointGroupSnapshot> getEndpointGroupsLocal() const;
	AIJsonConfigActiveTargetSnapshot getActiveTargetLocal() const;
	// 切换活动端点或分组；目标不存在、分组为空时保持原状态。
	bool setActiveTargetLocal(const AIJsonConfigActiveTargetSnapshot& activeTarget);

	// 原子替换端点拓扑；失败时保持原配置不变。
	bool replaceEndpointConfiguration(
		const std::vector<AIJsonConfigEndpointSnapshot>& endpoints,
		const std::vector<AIJsonConfigEndpointGroupSnapshot>& groups,
		const AIJsonConfigActiveTargetSnapshot& activeTarget);

	// 旧配置组 API，供原生回退窗口兼容使用。
	std::vector<AIJsonConfigProfileSnapshot> getProfilesLocal() const;
	std::string getActiveProfileId() const;
	bool setActiveProfileId(const std::string& activeProfileId);
	bool replaceProfiles(const std::vector<AIJsonConfigProfileSnapshot>& profiles, const std::string& activeProfileId);

private:
	struct StoredEndpoint {
		std::string id;
		std::string name;
		std::map<std::string, std::string> values;
	};

	struct StoredEndpointGroup {
		std::string id;
		std::string name;
		std::vector<std::string> endpointIds;
	};

	void load();
	void save() const;
	StoredEndpoint* findEndpoint(const std::string& id);
	const StoredEndpoint* findEndpoint(const std::string& id) const;
	StoredEndpointGroup* findGroup(const std::string& id);
	const StoredEndpointGroup* findGroup(const std::string& id) const;
	StoredEndpoint* findActiveEndpoint();
	const StoredEndpoint* findActiveEndpoint() const;
	void ensureWritableEndpoint();

	std::filesystem::path m_filePath;
	std::string m_activeTargetType;
	std::string m_activeTargetId;
	std::map<std::string, std::string> m_globalValues;
	std::vector<StoredEndpoint> m_endpoints;
	std::vector<StoredEndpointGroup> m_groups;
};

#endif // AI_JSON_CONFIG_H
