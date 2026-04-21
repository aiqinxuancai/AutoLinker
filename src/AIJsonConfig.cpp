// AIJsonConfig.cpp
#include "AIJsonConfig.h"

#include <fstream>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"
#include "PathHelper.h"

namespace {

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

std::string AIJsonConfig::getValue(const std::string& key) const
{
    const auto it = m_data.find(key);
    return it != m_data.end() ? it->second : std::string();
}

std::string AIJsonConfig::getValueLocal(const std::string& key) const
{
    return Utf8ToLocal(getValue(key));
}

void AIJsonConfig::setValue(const std::string& key, const std::string& localValue)
{
    m_data[key] = LocalToUtf8(localValue);
    save();
}

void AIJsonConfig::setValues(const std::map<std::string, std::string>& localPairs)
{
    for (const auto& [k, v] : localPairs) {
        m_data[k] = LocalToUtf8(v);
    }
    save();
}

bool AIJsonConfig::hasAnyData() const
{
    return !m_data.empty();
}

bool AIJsonConfig::hasKey(const std::string& key) const
{
    return m_data.count(key) > 0;
}

void AIJsonConfig::load()
{
    m_data.clear();
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
        for (const auto& [key, val] : j.items()) {
            if (val.is_string()) {
                // JSON 文件内部始终存储 UTF-8
                m_data[key] = val.get<std::string>();
            } else {
                m_data[key] = val.dump();
            }
        }
    } catch (...) {
        // 忽略解析错误，保持空数据
    }
}

void AIJsonConfig::save() const
{
    try {
        nlohmann::json j = nlohmann::json::object();
        for (const auto& [key, value] : m_data) {
            // m_data 内部值均为 UTF-8，可安全写入 JSON
            j[key] = value;
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

