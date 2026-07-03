// AIJsonConfig.cpp
#include "AIJsonConfig.h"

#include <fstream>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"
#include "PathHelper.h"

namespace {

bool IsReservedRootKey(const std::string& key)
{
    return key == "active_profile_id" || key == "profiles";
}

// 检查字符串是否为合法 UTF-8 编码
bool IsValidUtf8(const std::string& s)
{
    const auto* p = reinterpret_cast<const unsigned char*>(s.data());
    const auto* end = p + s.size();
    while (p < end) {
        unsigned char c = *p++;
        int extra = 0;
        if      (c < 0x80)  extra = 0;
        else if (c < 0xC0)  return false;
        else if (c < 0xE0)  extra = 1;
        else if (c < 0xF0)  extra = 2;
        else if (c < 0xF8)  extra = 3;
        else                return false;
        for (int i = 0; i < extra; ++i) {
            if (p >= end || (*p & 0xC0) != 0x80) return false;
            ++p;
        }
    }
    return true;
}

// 将本地编码（ANSI/GBK）字符串转为 UTF-8
std::string LocalToUtf8(const std::string& s)
{
    if (s.empty() || IsValidUtf8(s)) return s;
    const int wlen = MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return s;
    std::wstring ws(wlen - 1, L'\0');
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, ws.data(), wlen);
    const int ulen = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (ulen <= 0) return s;
    std::string u(ulen - 1, '\0');
    WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), -1, u.data(), ulen, nullptr, nullptr);
    return u;
}

// 将 UTF-8 字符串转为本地编码（ANSI/GBK）
std::string Utf8ToLocal(const std::string& s)
{
    if (s.empty() || !IsValidUtf8(s)) return s;
    const int wlen = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return s;
    std::wstring ws(wlen - 1, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, ws.data(), wlen);
    const int alen = WideCharToMultiByte(CP_ACP, 0, ws.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (alen <= 0) return s;
    std::string a(alen - 1, '\0');
    WideCharToMultiByte(CP_ACP, 0, ws.c_str(), -1, a.data(), alen, nullptr, nullptr);
    return a;
}

} // namespace

AIJsonConfig::AIJsonConfig()
{
    std::filesystem::path basePath = GetBasePath();
    std::filesystem::path autoLinkerPath = basePath / "AutoLinker";
    if (!std::filesystem::exists(autoLinkerPath)) {
        std::filesystem::create_directory(autoLinkerPath);
    }
    m_filePath = autoLinkerPath / "AIConfig.json";
    load();
}

AIJsonConfig::StoredProfile* AIJsonConfig::findActiveProfile()
{
    for (auto& profile : m_profiles) {
        if (profile.id == m_activeProfileId) {
            return &profile;
        }
    }
    return nullptr;
}

const AIJsonConfig::StoredProfile* AIJsonConfig::findActiveProfile() const
{
    for (const auto& profile : m_profiles) {
        if (profile.id == m_activeProfileId) {
            return &profile;
        }
    }
    return nullptr;
}

void AIJsonConfig::ensureWritableProfile()
{
    if (findActiveProfile() != nullptr) {
        return;
    }

    StoredProfile profile;
    profile.id = "default";
    profile.name = "默认";
    m_profiles.push_back(profile);
    m_activeProfileId = profile.id;
}

std::string AIJsonConfig::getValue(const std::string& key) const
{
    const StoredProfile* active = findActiveProfile();
    if (active == nullptr) {
        return std::string();
    }
    const auto it = active->values.find(key);
    return it != active->values.end() ? it->second : std::string();
}

std::string AIJsonConfig::getValueLocal(const std::string& key) const
{
    return Utf8ToLocal(getValue(key));
}

std::string AIJsonConfig::getGlobalValue(const std::string& key) const
{
    const auto it = m_globalValues.find(key);
    return it != m_globalValues.end() ? it->second : std::string();
}

std::string AIJsonConfig::getGlobalValueLocal(const std::string& key) const
{
    return Utf8ToLocal(getGlobalValue(key));
}

void AIJsonConfig::setValue(const std::string& key, const std::string& localValue)
{
    ensureWritableProfile();
    StoredProfile* active = findActiveProfile();
    if (active == nullptr) {
        return;
    }
    active->values[key] = LocalToUtf8(localValue);
    save();
}

void AIJsonConfig::setValues(const std::map<std::string, std::string>& localPairs)
{
    ensureWritableProfile();
    StoredProfile* active = findActiveProfile();
    if (active == nullptr) {
        return;
    }
    for (const auto& [k, v] : localPairs) {
        active->values[k] = LocalToUtf8(v);
    }
    save();
}

void AIJsonConfig::removeValues(const std::vector<std::string>& keys)
{
    StoredProfile* active = findActiveProfile();
    if (active == nullptr) {
        return;
    }
    for (const auto& key : keys) {
        active->values.erase(key);
    }
    save();
}

void AIJsonConfig::setGlobalValues(const std::map<std::string, std::string>& localPairs)
{
    for (const auto& [k, v] : localPairs) {
        if (IsReservedRootKey(k)) {
            continue;
        }
        m_globalValues[k] = LocalToUtf8(v);
    }
    save();
}

bool AIJsonConfig::hasAnyData() const
{
    const StoredProfile* active = findActiveProfile();
    return !m_globalValues.empty() || (active != nullptr && !active->values.empty());
}

bool AIJsonConfig::hasKey(const std::string& key) const
{
    const StoredProfile* active = findActiveProfile();
    return active != nullptr && active->values.count(key) > 0;
}

std::vector<AIJsonConfigProfileSnapshot> AIJsonConfig::getProfilesLocal() const
{
    std::vector<AIJsonConfigProfileSnapshot> snapshots;
    snapshots.reserve(m_profiles.size());
    for (const auto& profile : m_profiles) {
        AIJsonConfigProfileSnapshot snapshot;
        snapshot.id = Utf8ToLocal(profile.id);
        snapshot.name = Utf8ToLocal(profile.name);
        for (const auto& [key, value] : profile.values) {
            snapshot.values[key] = Utf8ToLocal(value);
        }
        snapshots.push_back(std::move(snapshot));
    }
    return snapshots;
}

std::string AIJsonConfig::getActiveProfileId() const
{
    return Utf8ToLocal(m_activeProfileId);
}

bool AIJsonConfig::replaceProfiles(const std::vector<AIJsonConfigProfileSnapshot>& profiles, const std::string& activeProfileId)
{
    std::vector<StoredProfile> nextProfiles;
    nextProfiles.reserve(profiles.size());
    bool foundActive = false;
    for (const auto& profile : profiles) {
        const std::string idUtf8 = LocalToUtf8(profile.id);
        const std::string nameUtf8 = LocalToUtf8(profile.name);
        if (idUtf8.empty() || nameUtf8.empty()) {
            return false;
        }

        StoredProfile stored;
        stored.id = idUtf8;
        stored.name = nameUtf8;
        for (const auto& [key, value] : profile.values) {
            stored.values[key] = LocalToUtf8(value);
        }
        if (idUtf8 == LocalToUtf8(activeProfileId)) {
            foundActive = true;
        }
        nextProfiles.push_back(std::move(stored));
    }

    if (nextProfiles.empty() || !foundActive) {
        return false;
    }

    m_profiles = std::move(nextProfiles);
    m_activeProfileId = LocalToUtf8(activeProfileId);
    save();
    return true;
}

void AIJsonConfig::load()
{
    m_profiles.clear();
    m_activeProfileId.clear();
    m_globalValues.clear();
    if (!std::filesystem::exists(m_filePath)) {
        return;
    }

    try {
        std::ifstream f(m_filePath);
        if (!f.is_open()) {
            return;
        }
        const nlohmann::json j = nlohmann::json::parse(f, nullptr, /*allow_exceptions=*/false);
        if (j.is_discarded() || !j.is_object()) {
            return;
        }

        if (j.contains("profiles") && j["profiles"].is_array()) {
            for (const auto& [key, val] : j.items()) {
                if (IsReservedRootKey(key)) {
                    continue;
                }
                m_globalValues[key] = val.is_string() ? val.get<std::string>() : val.dump();
            }

            for (const auto& item : j["profiles"]) {
                if (!item.is_object()) {
                    continue;
                }

                StoredProfile profile;
                if (item.contains("id") && item["id"].is_string()) {
                    profile.id = item["id"].get<std::string>();
                }
                if (item.contains("name") && item["name"].is_string()) {
                    profile.name = item["name"].get<std::string>();
                }
                if (item.contains("values") && item["values"].is_object()) {
                    for (const auto& [key, val] : item["values"].items()) {
                        profile.values[key] = val.is_string() ? val.get<std::string>() : val.dump();
                    }
                }
                if (!profile.id.empty() && !profile.name.empty()) {
                    m_profiles.push_back(std::move(profile));
                }
            }
            if (j.contains("active_profile_id") && j["active_profile_id"].is_string()) {
                m_activeProfileId = j["active_profile_id"].get<std::string>();
            }
            if (findActiveProfile() == nullptr && !m_profiles.empty()) {
                m_activeProfileId = m_profiles.front().id;
            }
            return;
        }

        StoredProfile legacyProfile;
        legacyProfile.id = "default";
        legacyProfile.name = "默认";
        for (const auto& [key, val] : j.items()) {
            if (val.is_string()) {
                legacyProfile.values[key] = val.get<std::string>();
            } else {
                legacyProfile.values[key] = val.dump();
            }
        }
        m_profiles.push_back(std::move(legacyProfile));
        m_activeProfileId = "default";
    } catch (...) {
        // 忽略解析错误，保持空数据
    }
}

void AIJsonConfig::save() const
{
    try {
        nlohmann::json j = nlohmann::json::object();
        for (const auto& [key, value] : m_globalValues) {
            if (!IsReservedRootKey(key)) {
                j[key] = value;
            }
        }
        j["active_profile_id"] = m_activeProfileId;
        j["profiles"] = nlohmann::json::array();
        for (const auto& profile : m_profiles) {
            nlohmann::json item = nlohmann::json::object();
            item["id"] = profile.id;
            item["name"] = profile.name;
            item["values"] = nlohmann::json::object();
            for (const auto& [key, value] : profile.values) {
                item["values"][key] = value;
            }
            j["profiles"].push_back(std::move(item));
        }
        // 先生成序列化内容，再打开文件，避免 dump 失败时留下空文件
        const std::string dumped = j.dump(4);
        std::ofstream f(m_filePath);
        if (f.is_open()) {
            f << dumped;
        }
    } catch (...) {
        // 忽略写入错误
    }
}

