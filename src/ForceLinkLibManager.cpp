#include "ForceLinkLibManager.h"

#include "PathHelper.h"

#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <fstream>
#include <iterator>
#include <system_error>
#include <utility>

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

std::string ConfigBytesToLocalText(const std::string& bytes)
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
		if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n') {
			break;
		}
		++begin;
	}
	while (end > begin) {
		const unsigned char ch = static_cast<unsigned char>(text[end - 1]);
		if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n') {
			break;
		}
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

bool ParseBoolValue(const std::string& text, bool fallback)
{
	const std::string value = ToLowerAscii(TrimAscii(text));
	if (value == "1" || value == "true" || value == "yes" || value == "on") {
		return true;
	}
	if (value == "0" || value == "false" || value == "no" || value == "off") {
		return false;
	}
	return fallback;
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

bool IsSectionLine(const std::string& line)
{
	return line.size() >= 2 && line.front() == '[' && line.back() == ']';
}

void PushRuleIfValid(std::vector<ForceLinkLibRule>& rules, ForceLinkLibRule& rule)
{
	rule.linkerName = TrimAscii(rule.linkerName);
	rule.libPath = TrimAscii(rule.libPath);
	if (!rule.libPath.empty()) {
		rules.push_back(rule);
	}
	rule = ForceLinkLibRule{};
}

} // namespace

ForceLinkLibManager::ForceLinkLibManager()
{
	std::filesystem::path autoLinkerPath = GetAutoLinkerDirectoryPath();
	std::error_code ec;
	std::filesystem::create_directories(autoLinkerPath, ec);
	configFilePath = autoLinkerPath / "ForceLinkLib.ini";
	loadConfig();
}

std::vector<ForceLinkLibRule> ForceLinkLibManager::getRules() const
{
	return rules_;
}

bool ForceLinkLibManager::replaceRules(const std::vector<ForceLinkLibRule>& rules, std::string* errorMessage)
{
	rules_.clear();
	for (ForceLinkLibRule rule : rules) {
		rule.linkerName = TrimAscii(rule.linkerName);
		rule.libPath = TrimAscii(rule.libPath);
		if (!rule.libPath.empty()) {
			rules_.push_back(std::move(rule));
		}
	}
	return saveConfig(errorMessage);
}

const std::filesystem::path& ForceLinkLibManager::getConfigFilePath() const
{
	return configFilePath;
}

void ForceLinkLibManager::loadConfig()
{
	rules_.clear();

	std::ifstream configFile(configFilePath, std::ios::binary);
	if (!configFile.is_open()) {
		return;
	}

	const std::string bytes((std::istreambuf_iterator<char>(configFile)), std::istreambuf_iterator<char>());
	const std::string text = ConfigBytesToLocalText(bytes);

	ForceLinkLibRule current;
	bool inRuleSection = false;
	size_t lineBegin = 0;
	while (lineBegin <= text.size()) {
		const size_t lineEnd = text.find('\n', lineBegin);
		std::string line = lineEnd == std::string::npos
			? text.substr(lineBegin)
			: text.substr(lineBegin, lineEnd - lineBegin);
		line = TrimAscii(line);

		if (line.empty() || line[0] == ';' || line[0] == '#') {
			// skip
		}
		else if (IsSectionLine(line)) {
			if (inRuleSection) {
				PushRuleIfValid(rules_, current);
			}
			const std::string sectionName = ToLowerAscii(TrimAscii(line.substr(1, line.size() - 2)));
			inRuleSection = sectionName.rfind("rule", 0) == 0;
			current = ForceLinkLibRule{};
		}
		else if (inRuleSection) {
			const size_t pos = line.find('=');
			if (pos != std::string::npos) {
				const std::string key = ToLowerAscii(TrimAscii(line.substr(0, pos)));
				const std::string value = TrimAscii(line.substr(pos + 1));
				if (key == "enabled") {
					current.enabled = ParseBoolValue(value, true);
				}
				else if (key == "linker") {
					current.linkerName = value;
				}
				else if (key == "path") {
					current.libPath = value;
				}
			}
		}

		if (lineEnd == std::string::npos) {
			break;
		}
		lineBegin = lineEnd + 1;
	}

	if (inRuleSection) {
		PushRuleIfValid(rules_, current);
	}
}

bool ForceLinkLibManager::saveConfig(std::string* errorMessage) const
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
	text += "; AutoLinker 核心库函数重写强制链接配置\r\n";
	text += "; linker 为空表示所有链接器生效；path 为要追加到 krnln_static.lib 前面的 .lib 文件路径。\r\n";
	for (size_t i = 0; i < rules_.size(); ++i) {
		const ForceLinkLibRule& rule = rules_[i];
		if (TrimAscii(rule.libPath).empty()) {
			continue;
		}
		text += "\r\n[rule.";
		text += std::to_string(i + 1);
		text += "]\r\n";
		text += "enabled=";
		text += rule.enabled ? "1" : "0";
		text += "\r\n";
		text += "linker=";
		text += TrimAscii(rule.linkerName);
		text += "\r\n";
		text += "path=";
		text += TrimAscii(rule.libPath);
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
