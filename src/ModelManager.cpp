// ModelManager.cpp - EC 模块动静态切换配置管理
#include "ModelManager.h"
#include "PathHelper.h"

#include <Windows.h>

#include <algorithm>
#include <fstream>
#include <iterator>
#include <system_error>

namespace {

bool IsValidUtf8(const std::string& text)
{
    if (text.empty()) {
        return true;
    }
    return MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        text.data(),
        static_cast<int>(text.size()),
        nullptr,
        0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
    if (text.empty()) {
        return std::string();
    }

    const int wideLen = MultiByteToWideChar(
        fromCodePage,
        fromFlags,
        text.data(),
        static_cast<int>(text.size()),
        nullptr,
        0);
    if (wideLen <= 0) {
        return text;
    }

    std::wstring wide(static_cast<size_t>(wideLen), L'\0');
    if (MultiByteToWideChar(
        fromCodePage,
        fromFlags,
        text.data(),
        static_cast<int>(text.size()),
        wide.data(),
        wideLen) <= 0) {
        return text;
    }

    const int outLen = WideCharToMultiByte(
        toCodePage,
        0,
        wide.data(),
        wideLen,
        nullptr,
        0,
        nullptr,
        nullptr);
    if (outLen <= 0) {
        return text;
    }

    std::string out(static_cast<size_t>(outLen), '\0');
    if (WideCharToMultiByte(
        toCodePage,
        0,
        wide.data(),
        wideLen,
        out.data(),
        outLen,
        nullptr,
        nullptr) <= 0) {
        return text;
    }
    return out;
}

std::string Utf8ToLocalText(const std::string& text)
{
    if (text.empty()) {
        return std::string();
    }
    return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string LocalToUtf8Text(const std::string& text)
{
    if (text.empty()) {
        return std::string();
    }
    return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

std::string ModelConfigBytesToLocalText(const std::string& bytes)
{
    if (bytes.empty()) {
        return std::string();
    }
    if (bytes.size() >= 3 &&
        static_cast<unsigned char>(bytes[0]) == 0xEF &&
        static_cast<unsigned char>(bytes[1]) == 0xBB &&
        static_cast<unsigned char>(bytes[2]) == 0xBF) {
        return Utf8ToLocalText(bytes.substr(3));
    }
    if (IsValidUtf8(bytes)) {
        return Utf8ToLocalText(bytes);
    }
    return bytes;
}

std::string TrimAscii(const std::string& text)
{
    size_t begin = 0;
    size_t end = text.size();
    while (begin < end) {
        const unsigned char ch = static_cast<unsigned char>(text[begin]);
        if (ch != ' ' && ch != '\t') {
            break;
        }
        ++begin;
    }
    while (end > begin) {
        const unsigned char ch = static_cast<unsigned char>(text[end - 1]);
        if (ch != ' ' && ch != '\t' && ch != '\r') {
            break;
        }
        --end;
    }
    return text.substr(begin, end - begin);
}

std::string NormalizeLineBreaksToCrlf(const std::string& text)
{
    std::string out;
    out.reserve(text.size() + 16);
    for (size_t i = 0; i < text.size(); ++i) {
        const char ch = text[i];
        if (ch == '\r') {
            if (i + 1 < text.size() && text[i + 1] == '\n') {
                ++i;
            }
            out += "\r\n";
        }
        else if (ch == '\n') {
            out += "\r\n";
        }
        else {
            out.push_back(ch);
        }
    }
    return out;
}

} // namespace

ModelManager::ModelManager()
{
    std::filesystem::path autoLinkerPath = GetAutoLinkerDirectoryPath();
    std::error_code ec;
    std::filesystem::create_directories(autoLinkerPath, ec);
    configFilePath = autoLinkerPath / "ModelManager.ini";
    loadConfig();
}

std::string ModelManager::getValue(const std::string& key)
{
    const auto it = configData.find(key);
    return it != configData.end() ? it->second : std::string();
}

void ModelManager::setValue(const std::string& key, const std::string& value)
{
    configData[key] = value;
    saveConfig();
}

std::string ModelManager::getKeyFromValue(const std::string& value)
{
    for (const auto& kv : configData) {
        if (kv.second == value) {
            return kv.first;
        }
    }
    return std::string();
}

std::map<std::string, std::string> ModelManager::getAllValues() const
{
    return configData;
}

bool ModelManager::replaceAllValues(const std::map<std::string, std::string>& values, std::string* errorMessage)
{
    configData = values;
    return saveConfig(errorMessage);
}

const std::filesystem::path& ModelManager::getConfigFilePath() const
{
    return configFilePath;
}

void ModelManager::loadConfig()
{
    configData.clear();

    std::ifstream configFile(configFilePath, std::ios::binary);
    if (!configFile.is_open()) {
        return;
    }

    const std::string bytes((std::istreambuf_iterator<char>(configFile)), std::istreambuf_iterator<char>());
    const std::string text = ModelConfigBytesToLocalText(bytes);

    size_t lineBegin = 0;
    while (lineBegin <= text.size()) {
        const size_t lineEnd = text.find('\n', lineBegin);
        std::string line = lineEnd == std::string::npos
            ? text.substr(lineBegin)
            : text.substr(lineBegin, lineEnd - lineBegin);
        if (!line.empty() && line.back() == '\r') {
            line.pop_back();
        }

        const size_t pos = line.find('=');
        if (pos != std::string::npos) {
            const std::string key = TrimAscii(line.substr(0, pos));
            const std::string value = TrimAscii(line.substr(pos + 1));
            if (!key.empty() && !value.empty()) {
                configData[key] = value;
            }
        }

        if (lineEnd == std::string::npos) {
            break;
        }
        lineBegin = lineEnd + 1;
    }
}

bool ModelManager::saveConfig(std::string* errorMessage) const
{
    std::error_code ec;
    std::filesystem::create_directories(configFilePath.parent_path(), ec);
    if (ec) {
        if (errorMessage != nullptr) {
            *errorMessage = "无法创建配置目录：" + ec.message();
        }
        return false;
    }

    std::string text;
    text.reserve(configData.size() * 32);
    for (const auto& [key, value] : configData) {
        if (key.empty() || value.empty()) {
            continue;
        }
        text += key;
        text.push_back('=');
        text += value;
        text += "\r\n";
    }

    std::ofstream configFile(configFilePath, std::ios::binary | std::ios::trunc);
    if (!configFile.is_open()) {
        if (errorMessage != nullptr) {
            *errorMessage = "无法写入配置文件。";
        }
        return false;
    }

    static constexpr unsigned char kUtf8Bom[] = { 0xEF, 0xBB, 0xBF };
    configFile.write(reinterpret_cast<const char*>(kUtf8Bom), sizeof(kUtf8Bom));
    const std::string normalized = NormalizeLineBreaksToCrlf(text);
    const std::string utf8 = LocalToUtf8Text(normalized);
    configFile.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
    if (!configFile.good()) {
        if (errorMessage != nullptr) {
            *errorMessage = "写入配置文件失败。";
        }
        return false;
    }
    return true;
}
