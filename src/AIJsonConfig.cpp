// AIJsonConfig.cpp
#include "AIJsonConfig.h"

#include <algorithm>
#include <fstream>
#include <set>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"
#include "PathHelper.h"

namespace {

constexpr int kAIConfigSchemaVersion = 2;

bool IsReservedRootKey(const std::string& key)
{
	return key == "schema_version" || key == "active_target" || key == "endpoints" ||
		key == "endpoint_groups" || key == "active_profile_id" || key == "profiles";
}

bool IsLegacyGlobalRootKey(const std::string& key)
{
	return key == "source_edit_mode" || key == "tavily_api_key";
}

bool IsValidUtf8(const std::string& text)
{
	const auto* cursor = reinterpret_cast<const unsigned char*>(text.data());
	const auto* end = cursor + text.size();
	while (cursor < end) {
		const unsigned char ch = *cursor++;
		int extra = 0;
		if (ch < 0x80) extra = 0;
		else if (ch < 0xC0) return false;
		else if (ch < 0xE0) extra = 1;
		else if (ch < 0xF0) extra = 2;
		else if (ch < 0xF8) extra = 3;
		else return false;
		for (int i = 0; i < extra; ++i) {
			if (cursor >= end || (*cursor & 0xC0) != 0x80) return false;
			++cursor;
		}
	}
	return true;
}

std::string LocalToUtf8(const std::string& text)
{
	if (text.empty() || IsValidUtf8(text)) return text;
	const int wideLength = MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, nullptr, 0);
	if (wideLength <= 0) return text;
	std::wstring wide(static_cast<size_t>(wideLength), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.c_str(), -1, wide.data(), wideLength) <= 0) return text;
	const int utf8Length = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (utf8Length <= 0) return text;
	std::string utf8(static_cast<size_t>(utf8Length), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, utf8.data(), utf8Length, nullptr, nullptr) <= 0) return text;
	utf8.pop_back();
	return utf8;
}

std::string Utf8ToLocal(const std::string& text)
{
	if (text.empty() || !IsValidUtf8(text)) return text;
	const int wideLength = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
	if (wideLength <= 0) return text;
	std::wstring wide(static_cast<size_t>(wideLength), L'\0');
	if (MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, wide.data(), wideLength) <= 0) return text;
	const int localLength = WideCharToMultiByte(CP_ACP, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (localLength <= 0) return text;
	std::string local(static_cast<size_t>(localLength), '\0');
	if (WideCharToMultiByte(CP_ACP, 0, wide.c_str(), -1, local.data(), localLength, nullptr, nullptr) <= 0) return text;
	local.pop_back();
	return local;
}

std::string JsonScalarToString(const nlohmann::json& value)
{
	return value.is_string() ? value.get<std::string>() : value.dump();
}

} // namespace

AIJsonConfig::AIJsonConfig()
{
	const std::filesystem::path directory = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code error;
	std::filesystem::create_directories(directory, error);
	m_filePath = directory / "AIConfig.json";
	load();
}

AIJsonConfig::AIJsonConfig(const std::filesystem::path& filePath)
	: m_filePath(filePath)
{
	load();
}

AIJsonConfig::StoredEndpoint* AIJsonConfig::findEndpoint(const std::string& id)
{
	const auto it = std::find_if(m_endpoints.begin(), m_endpoints.end(), [&id](const StoredEndpoint& endpoint) {
		return endpoint.id == id;
	});
	return it == m_endpoints.end() ? nullptr : &*it;
}

const AIJsonConfig::StoredEndpoint* AIJsonConfig::findEndpoint(const std::string& id) const
{
	const auto it = std::find_if(m_endpoints.begin(), m_endpoints.end(), [&id](const StoredEndpoint& endpoint) {
		return endpoint.id == id;
	});
	return it == m_endpoints.end() ? nullptr : &*it;
}

AIJsonConfig::StoredEndpointGroup* AIJsonConfig::findGroup(const std::string& id)
{
	const auto it = std::find_if(m_groups.begin(), m_groups.end(), [&id](const StoredEndpointGroup& group) {
		return group.id == id;
	});
	return it == m_groups.end() ? nullptr : &*it;
}

const AIJsonConfig::StoredEndpointGroup* AIJsonConfig::findGroup(const std::string& id) const
{
	const auto it = std::find_if(m_groups.begin(), m_groups.end(), [&id](const StoredEndpointGroup& group) {
		return group.id == id;
	});
	return it == m_groups.end() ? nullptr : &*it;
}

AIJsonConfig::StoredEndpoint* AIJsonConfig::findActiveEndpoint()
{
	if (m_activeTargetType == "endpoint") return findEndpoint(m_activeTargetId);
	if (m_activeTargetType == "group") {
		const StoredEndpointGroup* group = findGroup(m_activeTargetId);
		if (group != nullptr && !group->endpointIds.empty()) return findEndpoint(group->endpointIds.front());
	}
	return nullptr;
}

const AIJsonConfig::StoredEndpoint* AIJsonConfig::findActiveEndpoint() const
{
	if (m_activeTargetType == "endpoint") return findEndpoint(m_activeTargetId);
	if (m_activeTargetType == "group") {
		const StoredEndpointGroup* group = findGroup(m_activeTargetId);
		if (group != nullptr && !group->endpointIds.empty()) return findEndpoint(group->endpointIds.front());
	}
	return nullptr;
}

void AIJsonConfig::ensureWritableEndpoint()
{
	if (findActiveEndpoint() != nullptr) return;
	StoredEndpoint endpoint;
	endpoint.id = "default";
	endpoint.name = "默认";
	m_endpoints.push_back(endpoint);
	m_activeTargetType = "endpoint";
	m_activeTargetId = endpoint.id;
}

std::string AIJsonConfig::getValue(const std::string& key) const
{
	const StoredEndpoint* active = findActiveEndpoint();
	if (active == nullptr) return {};
	const auto it = active->values.find(key);
	return it == active->values.end() ? std::string() : it->second;
}

std::string AIJsonConfig::getValueLocal(const std::string& key) const
{
	return Utf8ToLocal(getValue(key));
}

std::string AIJsonConfig::getGlobalValue(const std::string& key) const
{
	const auto it = m_globalValues.find(key);
	return it == m_globalValues.end() ? std::string() : it->second;
}

std::string AIJsonConfig::getGlobalValueLocal(const std::string& key) const
{
	return Utf8ToLocal(getGlobalValue(key));
}

void AIJsonConfig::setValue(const std::string& key, const std::string& localValue)
{
	setValues({ { key, localValue } });
}

void AIJsonConfig::setValues(const std::map<std::string, std::string>& localPairs)
{
	ensureWritableEndpoint();
	StoredEndpoint* active = findActiveEndpoint();
	if (active == nullptr) return;
	for (const auto& [key, value] : localPairs) active->values[key] = LocalToUtf8(value);
	save();
}

void AIJsonConfig::removeValues(const std::vector<std::string>& keys)
{
	StoredEndpoint* active = findActiveEndpoint();
	if (active == nullptr) return;
	for (const std::string& key : keys) active->values.erase(key);
	save();
}

void AIJsonConfig::setGlobalValues(const std::map<std::string, std::string>& localPairs)
{
	for (const auto& [key, value] : localPairs) {
		if (!IsReservedRootKey(key)) m_globalValues[key] = LocalToUtf8(value);
	}
	save();
}

bool AIJsonConfig::hasAnyData() const
{
	const StoredEndpoint* active = findActiveEndpoint();
	return !m_globalValues.empty() || (active != nullptr && !active->values.empty());
}

bool AIJsonConfig::hasKey(const std::string& key) const
{
	const StoredEndpoint* active = findActiveEndpoint();
	return active != nullptr && active->values.contains(key);
}

std::vector<AIJsonConfigEndpointSnapshot> AIJsonConfig::getEndpointsLocal() const
{
	std::vector<AIJsonConfigEndpointSnapshot> result;
	result.reserve(m_endpoints.size());
	for (const StoredEndpoint& endpoint : m_endpoints) {
		AIJsonConfigEndpointSnapshot snapshot;
		snapshot.id = Utf8ToLocal(endpoint.id);
		snapshot.name = Utf8ToLocal(endpoint.name);
		for (const auto& [key, value] : endpoint.values) snapshot.values[key] = Utf8ToLocal(value);
		result.push_back(std::move(snapshot));
	}
	return result;
}

std::vector<AIJsonConfigEndpointGroupSnapshot> AIJsonConfig::getEndpointGroupsLocal() const
{
	std::vector<AIJsonConfigEndpointGroupSnapshot> result;
	result.reserve(m_groups.size());
	for (const StoredEndpointGroup& group : m_groups) {
		AIJsonConfigEndpointGroupSnapshot snapshot;
		snapshot.id = Utf8ToLocal(group.id);
		snapshot.name = Utf8ToLocal(group.name);
		for (const std::string& endpointId : group.endpointIds) snapshot.endpointIds.push_back(Utf8ToLocal(endpointId));
		result.push_back(std::move(snapshot));
	}
	return result;
}

AIJsonConfigActiveTargetSnapshot AIJsonConfig::getActiveTargetLocal() const
{
	return { m_activeTargetType, Utf8ToLocal(m_activeTargetId) };
}

bool AIJsonConfig::setActiveTargetLocal(const AIJsonConfigActiveTargetSnapshot& activeTarget)
{
	const std::string id = LocalToUtf8(activeTarget.id);
	if (activeTarget.type == "endpoint") {
		if (findEndpoint(id) == nullptr) return false;
	}
	else if (activeTarget.type == "group") {
		const StoredEndpointGroup* group = findGroup(id);
		if (group == nullptr || group->endpointIds.empty()) return false;
	}
	else {
		return false;
	}
	m_activeTargetType = activeTarget.type;
	m_activeTargetId = id;
	save();
	return true;
}

bool AIJsonConfig::replaceEndpointConfiguration(
	const std::vector<AIJsonConfigEndpointSnapshot>& endpoints,
	const std::vector<AIJsonConfigEndpointGroupSnapshot>& groups,
	const AIJsonConfigActiveTargetSnapshot& activeTarget)
{
	std::vector<StoredEndpoint> nextEndpoints;
	std::vector<StoredEndpointGroup> nextGroups;
	std::set<std::string> endpointIds;
	std::set<std::string> groupIds;
	nextEndpoints.reserve(endpoints.size());
	nextGroups.reserve(groups.size());

	for (const AIJsonConfigEndpointSnapshot& endpoint : endpoints) {
		StoredEndpoint stored;
		stored.id = LocalToUtf8(endpoint.id);
		stored.name = LocalToUtf8(endpoint.name);
		if (stored.id.empty() || stored.name.empty() || !endpointIds.insert(stored.id).second) return false;
		for (const auto& [key, value] : endpoint.values) stored.values[key] = LocalToUtf8(value);
		nextEndpoints.push_back(std::move(stored));
	}
	if (nextEndpoints.empty()) return false;

	for (const AIJsonConfigEndpointGroupSnapshot& group : groups) {
		StoredEndpointGroup stored;
		stored.id = LocalToUtf8(group.id);
		stored.name = LocalToUtf8(group.name);
		if (stored.id.empty() || stored.name.empty() || !groupIds.insert(stored.id).second) return false;
		std::set<std::string> memberIds;
		for (const std::string& localId : group.endpointIds) {
			const std::string id = LocalToUtf8(localId);
			if (!endpointIds.contains(id) || !memberIds.insert(id).second) return false;
			stored.endpointIds.push_back(id);
		}
		nextGroups.push_back(std::move(stored));
	}

	const std::string targetType = activeTarget.type;
	const std::string targetId = LocalToUtf8(activeTarget.id);
	if (targetType == "endpoint") {
		if (!endpointIds.contains(targetId)) return false;
	}
	else if (targetType == "group") {
		const auto it = std::find_if(nextGroups.begin(), nextGroups.end(), [&targetId](const StoredEndpointGroup& group) {
			return group.id == targetId;
		});
		if (it == nextGroups.end() || it->endpointIds.empty()) return false;
	}
	else {
		return false;
	}

	m_endpoints = std::move(nextEndpoints);
	m_groups = std::move(nextGroups);
	m_activeTargetType = targetType;
	m_activeTargetId = targetId;
	save();
	return true;
}

std::vector<AIJsonConfigProfileSnapshot> AIJsonConfig::getProfilesLocal() const
{
	return getEndpointsLocal();
}

std::string AIJsonConfig::getActiveProfileId() const
{
	const StoredEndpoint* active = findActiveEndpoint();
	return active == nullptr ? std::string() : Utf8ToLocal(active->id);
}

bool AIJsonConfig::setActiveProfileId(const std::string& activeProfileId)
{
	return setActiveTargetLocal({ "endpoint", activeProfileId });
}

bool AIJsonConfig::replaceProfiles(
	const std::vector<AIJsonConfigProfileSnapshot>& profiles,
	const std::string& activeProfileId)
{
	auto groups = getEndpointGroupsLocal();
	std::set<std::string> retainedIds;
	for (const auto& profile : profiles) retainedIds.insert(profile.id);
	for (auto& group : groups) {
		group.endpointIds.erase(
			std::remove_if(group.endpointIds.begin(), group.endpointIds.end(), [&retainedIds](const std::string& id) {
				return !retainedIds.contains(id);
			}),
			group.endpointIds.end());
	}
	AIJsonConfigActiveTargetSnapshot target = getActiveTargetLocal();
	if (target.type != "group") target = { "endpoint", activeProfileId };
	return replaceEndpointConfiguration(profiles, groups, target);
}

void AIJsonConfig::load()
{
	m_endpoints.clear();
	m_groups.clear();
	m_activeTargetType.clear();
	m_activeTargetId.clear();
	m_globalValues.clear();
	if (!std::filesystem::exists(m_filePath)) return;

	try {
		std::ifstream file(m_filePath, std::ios::binary);
		if (!file.is_open()) return;
		const nlohmann::json root = nlohmann::json::parse(file, nullptr, false);
		if (root.is_discarded() || !root.is_object()) return;

		const auto loadStructuredGlobals = [this, &root]() {
			for (const auto& [key, value] : root.items()) {
				if (!IsReservedRootKey(key)) m_globalValues[key] = JsonScalarToString(value);
			}
		};

		if (root.contains("endpoints") && root["endpoints"].is_array()) {
			loadStructuredGlobals();
			for (const auto& item : root["endpoints"]) {
				if (!item.is_object()) continue;
				StoredEndpoint endpoint;
				endpoint.id = item.value("id", "");
				endpoint.name = item.value("name", "");
				if (item.contains("values") && item["values"].is_object()) {
					for (const auto& [key, value] : item["values"].items()) endpoint.values[key] = JsonScalarToString(value);
				}
				if (!endpoint.id.empty() && !endpoint.name.empty() && findEndpoint(endpoint.id) == nullptr) {
					m_endpoints.push_back(std::move(endpoint));
				}
			}
			if (root.contains("endpoint_groups") && root["endpoint_groups"].is_array()) {
				for (const auto& item : root["endpoint_groups"]) {
					if (!item.is_object()) continue;
					StoredEndpointGroup group;
					group.id = item.value("id", "");
					group.name = item.value("name", "");
					std::set<std::string> memberIds;
					if (item.contains("endpoint_ids") && item["endpoint_ids"].is_array()) {
						for (const auto& idValue : item["endpoint_ids"]) {
							if (!idValue.is_string()) continue;
							const std::string id = idValue.get<std::string>();
							if (findEndpoint(id) != nullptr && memberIds.insert(id).second) group.endpointIds.push_back(id);
						}
					}
					if (!group.id.empty() && !group.name.empty() && findGroup(group.id) == nullptr) {
						m_groups.push_back(std::move(group));
					}
				}
			}
			if (root.contains("active_target") && root["active_target"].is_object()) {
				m_activeTargetType = root["active_target"].value("type", "");
				m_activeTargetId = root["active_target"].value("id", "");
			}
		}
		else if (root.contains("profiles") && root["profiles"].is_array()) {
			loadStructuredGlobals();
			// v1 配置在内存中迁移为端点，等用户保存后再写成 v2。
			for (const auto& item : root["profiles"]) {
				if (!item.is_object()) continue;
				StoredEndpoint endpoint;
				endpoint.id = item.value("id", "");
				endpoint.name = item.value("name", "");
				if (item.contains("values") && item["values"].is_object()) {
					for (const auto& [key, value] : item["values"].items()) endpoint.values[key] = JsonScalarToString(value);
				}
				if (!endpoint.id.empty() && !endpoint.name.empty() && findEndpoint(endpoint.id) == nullptr) {
					m_endpoints.push_back(std::move(endpoint));
				}
			}
			m_activeTargetType = "endpoint";
			m_activeTargetId = root.value("active_profile_id", "");
		}
		else {
			StoredEndpoint endpoint;
			endpoint.id = "default";
			endpoint.name = "默认";
			for (const auto& [key, value] : root.items()) {
				if (IsLegacyGlobalRootKey(key)) m_globalValues[key] = JsonScalarToString(value);
				else endpoint.values[key] = JsonScalarToString(value);
			}
			m_endpoints.push_back(std::move(endpoint));
			m_activeTargetType = "endpoint";
			m_activeTargetId = "default";
		}

		const bool targetValid =
			(m_activeTargetType == "endpoint" && findEndpoint(m_activeTargetId) != nullptr) ||
			(m_activeTargetType == "group" && findGroup(m_activeTargetId) != nullptr && !findGroup(m_activeTargetId)->endpointIds.empty());
		if (!targetValid && !m_endpoints.empty()) {
			m_activeTargetType = "endpoint";
			m_activeTargetId = m_endpoints.front().id;
		}
	}
	catch (...) {
		m_endpoints.clear();
		m_groups.clear();
		m_activeTargetType.clear();
		m_activeTargetId.clear();
		m_globalValues.clear();
	}
}

void AIJsonConfig::save() const
{
	try {
		nlohmann::json root = nlohmann::json::object();
		for (const auto& [key, value] : m_globalValues) {
			if (!IsReservedRootKey(key)) root[key] = value;
		}
		root["schema_version"] = kAIConfigSchemaVersion;
		root["active_target"] = { { "type", m_activeTargetType }, { "id", m_activeTargetId } };
		root["endpoints"] = nlohmann::json::array();
		for (const StoredEndpoint& endpoint : m_endpoints) {
			nlohmann::json item = {
				{ "id", endpoint.id },
				{ "name", endpoint.name },
				{ "values", nlohmann::json::object() }
			};
			for (const auto& [key, value] : endpoint.values) item["values"][key] = value;
			root["endpoints"].push_back(std::move(item));
		}
		root["endpoint_groups"] = nlohmann::json::array();
		for (const StoredEndpointGroup& group : m_groups) {
			root["endpoint_groups"].push_back({
				{ "id", group.id },
				{ "name", group.name },
				{ "endpoint_ids", group.endpointIds }
			});
		}

		const std::string dumped = root.dump(4);
		std::error_code error;
		if (!m_filePath.parent_path().empty()) std::filesystem::create_directories(m_filePath.parent_path(), error);
		std::ofstream file(m_filePath, std::ios::binary | std::ios::trunc);
		if (file.is_open()) file << dumped;
	}
	catch (...) {
	}
}
