#include "AIChatThemeManager.h"

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <set>
#include <string>

#include "PathHelper.h"
#include "StringHelper.h"

namespace AIChatThemeManager {
namespace {

constexpr const char* kDefaultThemeId = "default";
constexpr const char* kDefaultThemeName = "默认配色";
constexpr const char* kCurrentThemeFileName = "current_theme.json";
constexpr const char* kThemesDirectoryName = "AIChatThemes";

std::filesystem::path GetThemeRoot()
{
	std::filesystem::path dir = GetAutoLinkerDirectoryPath() / "Config" / kThemesDirectoryName;
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir;
}

std::filesystem::path GetCurrentThemePath()
{
	return GetThemeRoot() / kCurrentThemeFileName;
}

std::string TrimAscii(std::string text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
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

std::string SanitizeThemeId(const std::string& id, const std::string& fallbackName)
{
	std::string candidate = TrimAscii(id);
	if (candidate.empty()) {
		candidate = TrimAscii(fallbackName);
	}
	if (candidate.empty()) {
		candidate = "theme";
	}
	std::string out;
	out.reserve(candidate.size());
	for (char ch : candidate) {
		if ((ch >= 'a' && ch <= 'z') ||
			(ch >= 'A' && ch <= 'Z') ||
			(ch >= '0' && ch <= '9') ||
			ch == '_' || ch == '-') {
			out.push_back(ch);
		}
		else {
			out.push_back('_');
		}
	}
	while (!out.empty() && out.front() == '_') {
		out.erase(out.begin());
	}
	while (!out.empty() && out.back() == '_') {
		out.pop_back();
	}
	if (out.empty()) {
		out = "theme";
	}
	if (ToLowerAscii(out) == kDefaultThemeId) {
		out = "theme_default";
	}
	return out;
}

std::filesystem::path ThemeFilePathForId(const std::string& id)
{
	return GetThemeRoot() / (SanitizeThemeId(id, "theme") + ".json");
}

std::string ReadFileUtf8(const std::filesystem::path& path)
{
	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		return {};
	}
	std::string text((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		text.erase(0, 3);
	}
	return text;
}

bool WriteFileUtf8Bom(const std::filesystem::path& path, const std::string& text)
{
	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);
	std::ofstream out(path, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		return false;
	}
	static constexpr unsigned char kBom[] = { 0xEF, 0xBB, 0xBF };
	out.write(reinterpret_cast<const char*>(kBom), sizeof(kBom));
	std::string normalized = text;
	std::string crlf;
	crlf.reserve(normalized.size() + 16);
	for (size_t i = 0; i < normalized.size(); ++i) {
		const char ch = normalized[i];
		if (ch == '\r') {
			if (i + 1 < normalized.size() && normalized[i + 1] == '\n') {
				++i;
			}
			crlf += "\r\n";
		}
		else if (ch == '\n') {
			crlf += "\r\n";
		}
		else {
			crlf.push_back(ch);
		}
	}
	out.write(crlf.data(), static_cast<std::streamsize>(crlf.size()));
	return out.good();
}

bool IsHexColor(const std::string& value)
{
	if (value.size() != 7 || value[0] != '#') {
		return false;
	}
	for (size_t i = 1; i < value.size(); ++i) {
		if (std::isxdigit(static_cast<unsigned char>(value[i])) == 0) {
			return false;
		}
	}
	return true;
}

std::string NormalizeColorValue(const nlohmann::json& value, const std::string& fallback)
{
	if (!value.is_string()) {
		return fallback;
	}
	std::string text = TrimAscii(value.get<std::string>());
	if (!IsHexColor(text)) {
		return fallback;
	}
	std::transform(text.begin() + 1, text.end(), text.begin() + 1, [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

nlohmann::json MergeWithDefaults(const nlohmann::json& colors)
{
	nlohmann::json merged = GetDefaultColors();
	if (!colors.is_object()) {
		return merged;
	}
	for (auto& item : merged.items()) {
		if (colors.contains(item.key())) {
			item.value() = NormalizeColorValue(colors[item.key()], item.value().get<std::string>());
		}
	}
	return merged;
}

ThemeEntry BuildDefaultTheme()
{
	ThemeEntry entry;
	entry.id = kDefaultThemeId;
	entry.name = kDefaultThemeName;
	entry.colors = GetDefaultColors();
	entry.isDefault = true;
	return entry;
}

bool TryParseThemeFile(const std::filesystem::path& path, ThemeEntry& outEntry)
{
	try {
		const std::string text = ReadFileUtf8(path);
		if (text.empty()) {
			return false;
		}
		const nlohmann::json root = nlohmann::json::parse(text);
		if (!root.is_object()) {
			return false;
		}
		outEntry.id = SanitizeThemeId(root.value("id", path.stem().string()), path.stem().string());
		outEntry.name = TrimAscii(root.value("name", outEntry.id));
		if (outEntry.name.empty()) {
			outEntry.name = outEntry.id;
		}
		outEntry.colors = MergeWithDefaults(root.value("colors", nlohmann::json::object()));
		outEntry.isDefault = false;
		return true;
	}
	catch (...) {
		return false;
	}
}

std::vector<ThemeEntry> LoadUserThemes()
{
	std::vector<ThemeEntry> entries;
	const std::filesystem::path dir = GetThemeRoot();
	std::error_code ec;
	if (!std::filesystem::exists(dir, ec)) {
		return entries;
	}
	std::set<std::string> seen;
	for (const auto& item : std::filesystem::directory_iterator(dir, ec)) {
		if (ec) {
			break;
		}
		if (!item.is_regular_file() ||
			item.path().extension() != ".json" ||
			item.path().filename() == kCurrentThemeFileName) {
			continue;
		}
		ThemeEntry entry;
		if (!TryParseThemeFile(item.path(), entry)) {
			continue;
		}
		const std::string lowerId = ToLowerAscii(entry.id);
		if (!seen.insert(lowerId).second) {
			continue;
		}
		entries.push_back(std::move(entry));
	}
	std::sort(entries.begin(), entries.end(), [](const ThemeEntry& a, const ThemeEntry& b) {
		return _stricmp(a.name.c_str(), b.name.c_str()) < 0;
	});
	return entries;
}

std::string LoadCurrentThemeId()
{
	try {
		const std::string text = ReadFileUtf8(GetCurrentThemePath());
		if (text.empty()) {
			return kDefaultThemeId;
		}
		const nlohmann::json root = nlohmann::json::parse(text);
		if (!root.is_object()) {
			return kDefaultThemeId;
		}
		return SanitizeThemeId(root.value("currentThemeId", std::string(kDefaultThemeId)), kDefaultThemeId);
	}
	catch (...) {
		return kDefaultThemeId;
	}
}

bool SaveCurrentThemeId(const std::string& id)
{
	nlohmann::json root;
	root["currentThemeId"] = id.empty() ? kDefaultThemeId : id;
	return WriteFileUtf8Bom(GetCurrentThemePath(), root.dump(2, ' ', false, nlohmann::json::error_handler_t::replace));
}

nlohmann::json ThemeEntryToJson(const ThemeEntry& entry)
{
	nlohmann::json item;
	item["id"] = entry.id;
	item["name"] = entry.name;
	item["isDefault"] = entry.isDefault;
	item["colors"] = entry.colors;
	return item;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int len = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (len <= 0) {
		return {};
	}
	std::wstring wide(static_cast<size_t>(len), L'\0');
	MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), wide.data(), len);
	return wide;
}

std::wstring EscapeJsSingleQuotedWide(const std::wstring& text)
{
	std::wstring out;
	out.reserve(text.size() + 16);
	for (wchar_t ch : text) {
		switch (ch) {
		case L'\\': out += L"\\\\"; break;
		case L'\'': out += L"\\'"; break;
		case L'\r': out += L"\\r"; break;
		case L'\n': out += L"\\n"; break;
		case L'\t': out += L"\\t"; break;
		case 0x2028: out += L"\\u2028"; break;
		case 0x2029: out += L"\\u2029"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

} // namespace

nlohmann::json GetDefaultColors()
{
	return {
		{"pageBg", "#e6e9ef"},
		{"text", "#4c4f69"},
		{"mutedText", "#5d6475"},
		{"subtleText", "#9aa0b0"},
		{"brand", "#8839ef"},
		{"primary", "#1e66f5"},
		{"primaryHover", "#1554d1"},
		{"accent", "#04a5e5"},
		{"success", "#2e7d32"},
		{"warning", "#8a5a00"},
		{"danger", "#b42318"},
		{"border", "#d7dbe3"},
		{"softBorder", "#ccd3dd"},
		{"panelBg", "#ffffff"},
		{"toolbarBg", "#e6e9ef"},
		{"composerBg", "#e6e9ef"},
		{"inputBg", "#ffffff"},
		{"inputBorder", "#d4d9e3"},
		{"buttonBg", "#ffffff"},
		{"buttonHoverBg", "#eef1f6"},
		{"userMsgBg", "#eff1f5"},
		{"assistantMsgBg", "#e6e9ef"},
		{"systemMsgBg", "#e6e9ef"},
		{"toolMsgBg", "#e6e9ef"},
		{"planMsgBg", "#eef3ff"},
		{"roleText", "#4c4f69"},
		{"bodyText", "#4c4f69"},
		{"linkText", "#0b63c9"},
		{"inlineCodeBg", "#cfd1d7"},
		{"codeBlockBg", "#ccd0da"},
		{"codeText", "#1f2937"},
		{"planBg", "#f7faff"},
		{"planBorder", "#b9cffb"},
		{"planText", "#374151"},
		{"chipBg", "#dce9ff"},
		{"approvalBg", "#ffffff"},
		{"diffBg", "#f7f9fc"},
		{"diffAdd", "#1b7f37"},
		{"diffDel", "#b42318"},
		{"placeholder", "#9ca3b4"}
	};
}

ThemeEntry LoadCurrentTheme()
{
	const std::string currentId = ToLowerAscii(LoadCurrentThemeId());
	if (currentId == kDefaultThemeId) {
		return BuildDefaultTheme();
	}
	for (const auto& entry : LoadUserThemes()) {
		if (ToLowerAscii(entry.id) == currentId) {
			return entry;
		}
	}
	return BuildDefaultTheme();
}

std::string BuildConfigPayloadJson()
{
	nlohmann::json payload;
	payload["currentThemeId"] = LoadCurrentTheme().id;
	payload["colorKeys"] = nlohmann::json::array();
	const nlohmann::json defaults = GetDefaultColors();
	for (const auto& item : defaults.items()) {
		payload["colorKeys"].push_back(item.key());
	}
	payload["themes"] = nlohmann::json::array();
	payload["themes"].push_back(ThemeEntryToJson(BuildDefaultTheme()));
	for (const auto& entry : LoadUserThemes()) {
		payload["themes"].push_back(ThemeEntryToJson(entry));
	}
	return payload.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

bool SaveConfigPayload(const nlohmann::json& data, std::string& outMessage)
{
	if (!data.is_object()) {
		outMessage = "提交数据格式错误。";
		return false;
	}
	const std::string currentThemeId = SanitizeThemeId(data.value("currentThemeId", std::string(kDefaultThemeId)), kDefaultThemeId);
	if (!data.contains("themes") || !data["themes"].is_array()) {
		outMessage = "缺少配色列表。";
		return false;
	}

	std::vector<ThemeEntry> nextThemes;
	std::set<std::string> seenIds;
	for (const auto& item : data["themes"]) {
		if (!item.is_object()) {
			continue;
		}
		if (item.value("isDefault", false)) {
			continue;
		}
		ThemeEntry entry;
		entry.id = SanitizeThemeId(item.value("id", std::string()), item.value("name", std::string()));
		entry.name = TrimAscii(item.value("name", entry.id));
		if (entry.name.empty()) {
			outMessage = "配色名称不能为空。";
			return false;
		}
		const std::string lowerId = ToLowerAscii(entry.id);
		if (!seenIds.insert(lowerId).second) {
			outMessage = "存在重复配色标识：" + entry.id;
			return false;
		}
		entry.colors = MergeWithDefaults(item.value("colors", nlohmann::json::object()));
		entry.isDefault = false;
		nextThemes.push_back(std::move(entry));
	}

	const std::filesystem::path dir = GetThemeRoot();
	std::set<std::wstring> keepFiles;
	for (const auto& entry : nextThemes) {
		const std::filesystem::path path = ThemeFilePathForId(entry.id);
		keepFiles.insert(path.filename().wstring());
		nlohmann::json root = ThemeEntryToJson(entry);
		root.erase("isDefault");
		if (!WriteFileUtf8Bom(path, root.dump(2, ' ', false, nlohmann::json::error_handler_t::replace))) {
			outMessage = "无法写入配色文件：" + entry.name;
			return false;
		}
	}

	std::error_code ec;
	for (const auto& item : std::filesystem::directory_iterator(dir, ec)) {
		if (ec) {
			break;
		}
		if (!item.is_regular_file() ||
			item.path().extension() != ".json" ||
			item.path().filename() == kCurrentThemeFileName) {
			continue;
		}
		if (keepFiles.find(item.path().filename().wstring()) == keepFiles.end()) {
			std::error_code removeEc;
			std::filesystem::remove(item.path(), removeEc);
		}
	}

	std::string selectedId = currentThemeId;
	if (ToLowerAscii(selectedId) != kDefaultThemeId &&
		seenIds.find(ToLowerAscii(selectedId)) == seenIds.end()) {
		selectedId = kDefaultThemeId;
	}
	if (!SaveCurrentThemeId(selectedId)) {
		outMessage = "无法保存当前配色选择。";
		return false;
	}

	outMessage = "AI 对话配色已保存。";
	return true;
}

std::wstring BuildApplyCurrentThemeScript()
{
	nlohmann::json payload = ThemeEntryToJson(LoadCurrentTheme());
	std::wstring script = L"if(window.autolinkerApplyTheme){window.autolinkerApplyTheme(JSON.parse('";
	script += EscapeJsSingleQuotedWide(Utf8ToWide(payload.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace)));
	script += L"'));}";
	return script;
}

} // namespace AIChatThemeManager
