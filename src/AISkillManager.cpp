#include "AISkillManager.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cctype>
#include <cwctype>
#include <fstream>
#include <format>
#include <mutex>
#include <set>
#include <system_error>
#include <string_view>
#include <unordered_set>

#include "Global.h"
#include "AISkillLocalPackage.h"
#include "PathHelper.h"
#include "PowerShellToolRunner.h"
#include "WinINetUtil.h"

namespace {

using nlohmann::json;

constexpr int kConfigVersion = 2;
constexpr int kMaxScanDepth = 6;
constexpr size_t kMaxScannedDirectories = 2000;
constexpr size_t kMaxSkillFileBytes = 4 * 1024 * 1024;
constexpr size_t kMaxSkillNameBytes = 64;
constexpr size_t kMaxSkillDescriptionBytes = 1024;
constexpr size_t kPromptCatalogBudgetBytes = 8000;
constexpr size_t kMaxReadResourceBytes = 256 * 1024;
constexpr size_t kMaxArchiveBytes = 100 * 1024 * 1024;
constexpr char kSourceMetadataFileName[] = ".autolinker-skill.json";
constexpr char kGitHubHeaders[] =
	"Accept: application/vnd.github+json\r\n"
	"User-Agent: AutoLinker-SkillManager\r\n"
	"X-GitHub-Api-Version: 2022-11-28\r\n";

std::recursive_mutex g_skillMutex;
std::atomic_ullong g_tempCounter = 1;

struct SkillConfig {
	std::unordered_set<std::string> disabledPaths;
	struct ExternalReference {
		std::filesystem::path skillFile;
		AISkillScope scope = AISkillScope::User;
		std::filesystem::path projectDirectory;
		long long addedAtUnixMs = 0;
	};
	std::vector<ExternalReference> externalReferences;
};

struct ParsedGithubSource {
	std::string owner;
	std::string repository;
	std::string reference;
	std::string subpath;
};

struct PreparedRepository {
	ParsedGithubSource source;
	std::string commit;
	std::filesystem::path stagingRoot;
	std::filesystem::path repositoryRoot;
	std::vector<AISkillInfo> candidates;
};

struct PreparedLocalSkills {
	AISkillPreparedLocalSource source;
	std::vector<AISkillInfo> candidates;
};

std::string TrimAscii(std::string text)
{
	while (!text.empty() && std::isspace(static_cast<unsigned char>(text.back()))) {
		text.pop_back();
	}
	size_t begin = 0;
	while (begin < text.size() && std::isspace(static_cast<unsigned char>(text[begin]))) {
		++begin;
	}
	return begin == 0 ? text : text.substr(begin);
}

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int size = WideCharToMultiByte(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return {};
	}
	std::string result(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(
			CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr) <= 0) {
		return {};
	}
	return result;
}

std::filesystem::path Utf8ToPath(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int size = MultiByteToWideChar(
		CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::filesystem::path(text);
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(
			CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return std::filesystem::path(text);
	}
	return std::filesystem::path(wide);
}

std::filesystem::path LocalToPath(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int size = MultiByteToWideChar(
		CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::filesystem::path(text);
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(
			CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return std::filesystem::path(text);
	}
	return std::filesystem::path(wide);
}

std::string PathToUtf8(const std::filesystem::path& path)
{
	return WideToUtf8(path.wstring());
}

std::string NormalizePathKey(const std::filesystem::path& path)
{
	std::error_code ec;
	std::filesystem::path normalized = std::filesystem::weakly_canonical(path, ec);
	if (ec) {
		normalized = std::filesystem::absolute(path, ec);
	}
	std::wstring wide = normalized.lexically_normal().wstring();
	std::transform(wide.begin(), wide.end(), wide.begin(), [](wchar_t ch) {
		return static_cast<wchar_t>(std::towlower(ch));
	});
	return WideToUtf8(wide);
}

bool IsPathInside(const std::filesystem::path& candidate, const std::filesystem::path& root)
{
	std::error_code ec;
	const auto normalizedCandidate = std::filesystem::weakly_canonical(candidate, ec);
	if (ec) {
		return false;
	}
	const auto normalizedRoot = std::filesystem::weakly_canonical(root, ec);
	if (ec) {
		return false;
	}
	auto candidateIt = normalizedCandidate.begin();
	for (auto rootIt = normalizedRoot.begin(); rootIt != normalizedRoot.end(); ++rootIt, ++candidateIt) {
		if (candidateIt == normalizedCandidate.end() || _wcsicmp(rootIt->c_str(), candidateIt->c_str()) != 0) {
			return false;
		}
	}
	return true;
}

bool IsPathStrictlyInside(const std::filesystem::path& candidate, const std::filesystem::path& root)
{
	return NormalizePathKey(candidate) != NormalizePathKey(root) && IsPathInside(candidate, root);
}

std::filesystem::path GetConfigPath()
{
	return GetAutoLinkerDirectoryPath() / "AISkillsConfig.json";
}

bool ReadFileBytes(const std::filesystem::path& path, size_t maxBytes, std::string& out, std::string& outError)
{
	out.clear();
	outError.clear();
	std::error_code ec;
	const auto size = std::filesystem::file_size(path, ec);
	if (ec) {
		outError = "无法读取文件大小：" + ec.message();
		return false;
	}
	if (size > maxBytes) {
		outError = std::format("文件超过大小限制（{} 字节）", maxBytes);
		return false;
	}
	std::ifstream input(path, std::ios::binary);
	if (!input.is_open()) {
		outError = "无法打开文件";
		return false;
	}
	out.assign(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
	if (out.size() >= 3 && static_cast<unsigned char>(out[0]) == 0xEF &&
		static_cast<unsigned char>(out[1]) == 0xBB && static_cast<unsigned char>(out[2]) == 0xBF) {
		out.erase(0, 3);
	}
	return true;
}

bool WriteFileBytesAtomic(const std::filesystem::path& path, const std::string& bytes, std::string& outError)
{
	outError.clear();
	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);
	if (ec) {
		outError = "创建目录失败：" + ec.message();
		return false;
	}
	std::filesystem::path temp = path;
	temp += std::format(L".tmp.{}.{}", GetCurrentProcessId(), g_tempCounter.fetch_add(1));
	{
		std::ofstream output(temp, std::ios::binary | std::ios::trunc);
		if (!output.is_open()) {
			outError = "无法写入临时文件";
			return false;
		}
		output.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
		if (!output.good()) {
			outError = "写入临时文件失败";
			output.close();
			std::filesystem::remove(temp, ec);
			return false;
		}
	}
	if (!MoveFileExW(temp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
		outError = "替换配置文件失败：" + std::system_category().message(GetLastError());
		std::filesystem::remove(temp, ec);
		return false;
	}
	return true;
}

SkillConfig LoadConfigFromPath(const std::filesystem::path& configPath)
{
	SkillConfig config;
	std::string text;
	std::string error;
	if (!ReadFileBytes(configPath, 2 * 1024 * 1024, text, error)) {
		return config;
	}
	const json root = json::parse(text, nullptr, false);
	if (!root.is_object()) {
		return config;
	}
	if (root.contains("disabled_paths") && root["disabled_paths"].is_array()) {
		for (const auto& item : root["disabled_paths"]) {
			if (item.is_string()) {
				config.disabledPaths.insert(ToLowerAscii(item.get<std::string>()));
			}
		}
	}
	if (root.contains("external_references") && root["external_references"].is_array()) {
		for (const auto& item : root["external_references"]) {
			if (!item.is_object()) {
				continue;
			}
			const std::string skillFile = TrimAscii(item.value("skill_file", std::string()));
			const std::string scope = ToLowerAscii(item.value("scope", std::string("global")));
			if (skillFile.empty() || (scope != "global" && scope != "project")) {
				continue;
			}
			SkillConfig::ExternalReference reference;
			reference.skillFile = Utf8ToPath(skillFile);
			reference.scope = scope == "project" ? AISkillScope::Repo : AISkillScope::User;
			reference.projectDirectory = Utf8ToPath(item.value("project_directory", std::string()));
			reference.addedAtUnixMs = item.value("added_at_unix_ms", 0LL);
			if (!reference.skillFile.is_absolute() || (reference.scope == AISkillScope::Repo &&
				(reference.projectDirectory.empty() || !reference.projectDirectory.is_absolute()))) {
				continue;
			}
			config.externalReferences.push_back(std::move(reference));
		}
	}
	return config;
}

SkillConfig LoadConfig()
{
	return LoadConfigFromPath(GetConfigPath());
}

bool SaveConfigToPath(const SkillConfig& config, const std::filesystem::path& configPath, std::string& outError)
{
	json disabled = json::array();
	std::vector<std::string> sorted(config.disabledPaths.begin(), config.disabledPaths.end());
	std::sort(sorted.begin(), sorted.end());
	for (const auto& path : sorted) {
		disabled.push_back(path);
	}
	std::vector<SkillConfig::ExternalReference> references = config.externalReferences;
	std::stable_sort(references.begin(), references.end(), [](const auto& left, const auto& right) {
		const std::string leftKey = NormalizePathKey(left.skillFile);
		const std::string rightKey = NormalizePathKey(right.skillFile);
		if (leftKey != rightKey) {
			return leftKey < rightKey;
		}
		if (left.scope != right.scope) {
			return left.scope == AISkillScope::User;
		}
		return NormalizePathKey(left.projectDirectory) < NormalizePathKey(right.projectDirectory);
	});
	json external = json::array();
	for (const auto& reference : references) {
		external.push_back({
			{"skill_file", PathToUtf8(reference.skillFile)},
			{"scope", reference.scope == AISkillScope::Repo ? "project" : "global"},
			{"project_directory", reference.scope == AISkillScope::Repo
				? PathToUtf8(reference.projectDirectory) : std::string()},
			{"added_at_unix_ms", reference.addedAtUnixMs}
		});
	}
	const json root = {
		{"version", kConfigVersion},
		{"disabled_paths", std::move(disabled)},
		{"external_references", std::move(external)}
	};
	return WriteFileBytesAtomic(configPath, root.dump(2), outError);
}

bool SaveConfig(const SkillConfig& config, std::string& outError)
{
	return SaveConfigToPath(config, GetConfigPath(), outError);
}

std::vector<std::string> SplitLines(const std::string& text)
{
	std::vector<std::string> lines;
	size_t begin = 0;
	while (begin <= text.size()) {
		const size_t end = text.find('\n', begin);
		std::string line = end == std::string::npos ? text.substr(begin) : text.substr(begin, end - begin);
		if (!line.empty() && line.back() == '\r') {
			line.pop_back();
		}
		lines.push_back(std::move(line));
		if (end == std::string::npos) {
			break;
		}
		begin = end + 1;
	}
	return lines;
}

std::string UnquoteYamlScalar(std::string value)
{
	value = TrimAscii(std::move(value));
	if (value.size() >= 2 && ((value.front() == '"' && value.back() == '"') ||
		(value.front() == '\'' && value.back() == '\''))) {
		const char quote = value.front();
		value = value.substr(1, value.size() - 2);
		if (quote == '"') {
			std::string decoded;
			decoded.reserve(value.size());
			for (size_t i = 0; i < value.size(); ++i) {
				if (value[i] == '\\' && i + 1 < value.size()) {
					const char next = value[++i];
					decoded.push_back(next == 'n' ? '\n' : next == 't' ? '\t' : next);
				}
				else {
					decoded.push_back(value[i]);
				}
			}
			value = std::move(decoded);
		}
	}
	return TrimAscii(std::move(value));
}

std::string SingleLine(std::string value)
{
	for (char& ch : value) {
		if (ch == '\r' || ch == '\n' || ch == '\t') {
			ch = ' ';
		}
	}
	std::string result;
	result.reserve(value.size());
	bool previousSpace = false;
	for (char ch : value) {
		const bool space = ch == ' ';
		if (!space || !previousSpace) {
			result.push_back(ch);
		}
		previousSpace = space;
	}
	return TrimAscii(std::move(result));
}

bool ParseSkillFile(const std::filesystem::path& path, AISkillInfo& outInfo)
{
	outInfo.skillFile = path;
	outInfo.directory = path.parent_path();
	outInfo.name = PathToUtf8(path.parent_path().filename());
	std::string text;
	if (!ReadFileBytes(path, kMaxSkillFileBytes, text, outInfo.error)) {
		return false;
	}
	const auto lines = SplitLines(text);
	if (lines.empty() || TrimAscii(lines.front()) != "---") {
		outInfo.error = "SKILL.md 缺少 YAML frontmatter";
		return false;
	}
	size_t closing = 0;
	for (size_t i = 1; i < lines.size(); ++i) {
		if (TrimAscii(lines[i]) == "---") {
			closing = i;
			break;
		}
	}
	if (closing == 0) {
		outInfo.error = "SKILL.md frontmatter 未闭合";
		return false;
	}
	bool inMetadata = false;
	for (size_t i = 1; i < closing; ++i) {
		const std::string raw = lines[i];
		const size_t indent = raw.find_first_not_of(" \t");
		if (indent == std::string::npos || TrimAscii(raw).starts_with('#')) {
			continue;
		}
		const size_t colon = raw.find(':', indent);
		if (colon == std::string::npos) {
			continue;
		}
		const std::string key = ToLowerAscii(TrimAscii(raw.substr(indent, colon - indent)));
		std::string value = TrimAscii(raw.substr(colon + 1));
		if (indent == 0) {
			inMetadata = key == "metadata";
		}
		if ((value == ">" || value == "|" || value == ">-" || value == "|-") && i + 1 < closing) {
			std::string folded;
			while (i + 1 < closing) {
				const std::string& next = lines[i + 1];
				const size_t nextIndent = next.find_first_not_of(" \t");
				if (nextIndent == std::string::npos) {
					++i;
					continue;
				}
				if (nextIndent <= indent) {
					break;
				}
				if (!folded.empty()) {
					folded.push_back(value.front() == '|' ? '\n' : ' ');
				}
				folded += TrimAscii(next);
				++i;
			}
			value = std::move(folded);
		}
		value = SingleLine(UnquoteYamlScalar(std::move(value)));
		if (indent == 0 && key == "name" && !value.empty()) {
			outInfo.name = value;
		}
		else if (indent == 0 && key == "description") {
			outInfo.description = value;
		}
		else if (indent > 0 && inMetadata && (key == "short-description" || key == "short_description")) {
			outInfo.shortDescription = value;
		}
	}
	if (outInfo.name.empty() || outInfo.name.size() > kMaxSkillNameBytes) {
		outInfo.error = "技能名称为空或超过 64 字节";
		return false;
	}
	if (outInfo.description.empty()) {
		outInfo.error = "SKILL.md frontmatter 缺少 description";
		return false;
	}
	if (outInfo.description.size() > kMaxSkillDescriptionBytes) {
		outInfo.description.resize(kMaxSkillDescriptionBytes);
	}
	outInfo.valid = true;
	return true;
}

void LoadSourceMetadata(AISkillInfo& skill)
{
	std::string text;
	std::string error;
	if (!ReadFileBytes(skill.directory / kSourceMetadataFileName, 64 * 1024, text, error)) {
		return;
	}
	const json root = json::parse(text, nullptr, false);
	if (!root.is_object()) {
		return;
	}
	skill.sourceRepository = root.value("repository", std::string());
	skill.sourceRef = root.value("ref", std::string());
	skill.sourcePath = root.value("path", std::string());
	skill.sourceCommit = root.value("commit", std::string());
	skill.installedAtUnixMs = root.value("installed_at_unix_ms", 0LL);
}

void DiscoverUnderRoot(
	const std::filesystem::path& root,
	AISkillScope scope,
	const SkillConfig& config,
	std::vector<AISkillInfo>& out)
{
	std::error_code ec;
	if (!std::filesystem::is_directory(root, ec)) {
		return;
	}
	size_t scannedDirectories = 0;
	std::filesystem::recursive_directory_iterator it(
		root, std::filesystem::directory_options::skip_permission_denied, ec), end;
	for (; it != end && !ec; it.increment(ec)) {
		if (it.depth() >= kMaxScanDepth && it->is_directory(ec)) {
			it.disable_recursion_pending();
		}
		if (it->is_symlink(ec) && it->is_directory(ec)) {
			it.disable_recursion_pending();
			continue;
		}
		if (it->is_directory(ec)) {
			if (++scannedDirectories >= kMaxScannedDirectories) {
				break;
			}
			continue;
		}
		if (!it->is_regular_file(ec) || _wcsicmp(it->path().filename().c_str(), L"SKILL.md") != 0) {
			continue;
		}
		AISkillInfo skill;
		skill.scope = scope;
		ParseSkillFile(it->path(), skill);
		skill.enabled = !config.disabledPaths.contains(NormalizePathKey(it->path()));
		LoadSourceMetadata(skill);
		out.push_back(std::move(skill));
		it.disable_recursion_pending();
	}
}

std::optional<std::filesystem::path> CurrentProjectDirectory()
{
	const auto projectRoot = AISkillManager::GetProjectSkillsRoot();
	if (!projectRoot) {
		return std::nullopt;
	}
	return projectRoot->parent_path().parent_path();
}

bool ReferenceAppliesToProject(
	const SkillConfig::ExternalReference& reference,
	const std::optional<std::filesystem::path>& projectDirectory)
{
	if (reference.scope == AISkillScope::User) {
		return true;
	}
	return projectDirectory &&
		NormalizePathKey(reference.projectDirectory) == NormalizePathKey(*projectDirectory);
}

void DiscoverExternalReferences(
	const SkillConfig& config,
	const std::optional<std::filesystem::path>& projectDirectory,
	std::vector<AISkillInfo>& out)
{
	std::unordered_set<std::string> discovered;
	for (const auto& skill : out) {
		discovered.insert(NormalizePathKey(skill.skillFile));
	}
	for (const auto& reference : config.externalReferences) {
		if (!ReferenceAppliesToProject(reference, projectDirectory)) {
			continue;
		}
		const std::string key = NormalizePathKey(reference.skillFile);
		if (!discovered.insert(key).second) {
			continue;
		}
		AISkillInfo skill;
		skill.scope = reference.scope;
		skill.externalReference = true;
		skill.installedAtUnixMs = reference.addedAtUnixMs;
		ParseSkillFile(reference.skillFile, skill);
		skill.enabled = !config.disabledPaths.contains(key);
		out.push_back(std::move(skill));
	}
}

std::vector<AISkillInfo> DiscoverSkillsUnlocked()
{
	const SkillConfig config = LoadConfig();
	std::vector<AISkillInfo> result;
	DiscoverUnderRoot(AISkillManager::GetGlobalSkillsRoot(), AISkillScope::User, config, result);
	if (const auto projectRoot = AISkillManager::GetProjectSkillsRoot()) {
		DiscoverUnderRoot(*projectRoot, AISkillScope::Repo, config, result);
	}
	DiscoverExternalReferences(config, CurrentProjectDirectory(), result);
	std::stable_sort(result.begin(), result.end(), [](const AISkillInfo& left, const AISkillInfo& right) {
		if (left.scope != right.scope) {
			return left.scope == AISkillScope::Repo;
		}
		return ToLowerAscii(left.name) < ToLowerAscii(right.name);
	});
	return result;
}

json SkillToJson(const AISkillInfo& skill)
{
	return {
		{"name", skill.name},
		{"description", skill.description},
		{"short_description", skill.shortDescription},
		{"scope", skill.scope == AISkillScope::Repo ? "project" : "global"},
		{"directory", PathToUtf8(skill.directory)},
		{"skill_file", PathToUtf8(skill.skillFile)},
		{"enabled", skill.enabled},
		{"valid", skill.valid},
		{"external_reference", skill.externalReference},
		{"error", skill.error},
		{"source_repository", skill.sourceRepository},
		{"source_ref", skill.sourceRef},
		{"source_path", skill.sourcePath},
		{"source_commit", skill.sourceCommit},
		{"installed_at_unix_ms", skill.installedAtUnixMs},
		{"managed", !skill.sourceRepository.empty()}
	};
}

std::vector<AISkillInfo> EffectiveSkills(const std::vector<AISkillInfo>& skills)
{
	std::vector<AISkillInfo> result;
	std::unordered_set<std::string> names;
	for (const auto& skill : skills) {
		if (!skill.valid || !skill.enabled) {
			continue;
		}
		const std::string key = ToLowerAscii(skill.name);
		if (names.insert(key).second) {
			result.push_back(skill);
		}
	}
	return result;
}

std::unordered_set<std::string> ExtractSkillMentions(const std::string& text)
{
	std::unordered_set<std::string> result;
	for (size_t i = 0; i < text.size(); ++i) {
		if (text[i] != '$') {
			continue;
		}
		size_t end = i + 1;
		while (end < text.size()) {
			const unsigned char ch = static_cast<unsigned char>(text[end]);
			if (!std::isalnum(ch) && ch != '-' && ch != '_' && ch != '.' && ch != ':') {
				break;
			}
			++end;
		}
		if (end > i + 1) {
			result.insert(ToLowerAscii(text.substr(i + 1, end - i - 1)));
			i = end - 1;
		}
	}
	return result;
}

std::string SanitizeDirectoryName(const std::string& text)
{
	std::string result;
	result.reserve(text.size());
	for (unsigned char ch : text) {
		if (std::isalnum(ch) || ch == '-' || ch == '_' || ch == '.') {
			result.push_back(static_cast<char>(ch));
		}
		else if (!result.empty() && result.back() != '-') {
			result.push_back('-');
		}
	}
	while (!result.empty() && (result.back() == '.' || result.back() == '-')) {
		result.pop_back();
	}
	return result.empty() ? "skill" : result;
}

std::string UrlEncode(const std::string& value)
{
	constexpr char hex[] = "0123456789ABCDEF";
	std::string result;
	for (unsigned char ch : value) {
		if (std::isalnum(ch) || ch == '-' || ch == '_' || ch == '.' || ch == '~') {
			result.push_back(static_cast<char>(ch));
		}
		else {
			result.push_back('%');
			result.push_back(hex[(ch >> 4) & 0x0F]);
			result.push_back(hex[ch & 0x0F]);
		}
	}
	return result;
}

std::vector<std::string> SplitPath(std::string text)
{
	std::replace(text.begin(), text.end(), '\\', '/');
	std::vector<std::string> parts;
	size_t begin = 0;
	while (begin <= text.size()) {
		const size_t end = text.find('/', begin);
		std::string part = end == std::string::npos ? text.substr(begin) : text.substr(begin, end - begin);
		if (!part.empty()) {
			parts.push_back(std::move(part));
		}
		if (end == std::string::npos) {
			break;
		}
		begin = end + 1;
	}
	return parts;
}

bool IsGithubIdentifier(const std::string& value)
{
	return !value.empty() && std::all_of(value.begin(), value.end(), [](unsigned char ch) {
		return std::isalnum(ch) || ch == '-' || ch == '_' || ch == '.';
	});
}

bool ParseGithubSource(std::string source, ParsedGithubSource& out, std::string& outError)
{
	out = {};
	outError.clear();
	source = TrimAscii(std::move(source));
	for (const std::string prefix : {"https://github.com/", "http://github.com/", "github.com/"}) {
		if (ToLowerAscii(source).starts_with(prefix)) {
			source.erase(0, prefix.size());
			break;
		}
	}
	if (source.ends_with(".git")) {
		source.resize(source.size() - 4);
	}
	const size_t query = source.find_first_of("?#");
	if (query != std::string::npos) {
		source.resize(query);
	}
	const auto parts = SplitPath(source);
	if (parts.size() < 2 || !IsGithubIdentifier(parts[0]) || !IsGithubIdentifier(parts[1])) {
		outError = "请输入公开 GitHub 仓库地址或 owner/repo";
		return false;
	}
	out.owner = parts[0];
	out.repository = parts[1];
	if (parts.size() >= 4 && (parts[2] == "tree" || parts[2] == "blob")) {
		out.reference = parts[3];
		for (size_t i = 4; i < parts.size(); ++i) {
			if (!out.subpath.empty()) {
				out.subpath += '/';
			}
			out.subpath += parts[i];
		}
		if (parts[2] == "blob" && !out.subpath.empty()) {
			const size_t slash = out.subpath.find_last_of('/');
			out.subpath = slash == std::string::npos ? std::string() : out.subpath.substr(0, slash);
		}
	}
	return true;
}

std::string RepositoryName(const ParsedGithubSource& source)
{
	return source.owner + "/" + source.repository;
}

bool FetchJson(const std::string& url, json& out, std::string& outError)
{
	const auto response = PerformGetRequest(url, kGitHubHeaders, 30000, false, false);
	if (response.second < 200 || response.second >= 300) {
		outError = std::format("HTTP {}", response.second);
		if (!response.first.empty()) {
			const json errorJson = json::parse(response.first, nullptr, false);
			if (errorJson.is_object() && errorJson.contains("message") && errorJson["message"].is_string()) {
				outError += ": " + errorJson["message"].get<std::string>();
			}
		}
		return false;
	}
	out = json::parse(response.first, nullptr, false);
	if (out.is_discarded()) {
		outError = "服务器返回了无效 JSON";
		return false;
	}
	return true;
}

bool WriteBinaryFile(const std::filesystem::path& path, const std::string& bytes, std::string& outError)
{
	std::ofstream output(path, std::ios::binary | std::ios::trunc);
	if (!output.is_open()) {
		outError = "无法写入下载文件";
		return false;
	}
	output.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
	if (!output.good()) {
		outError = "写入下载文件失败";
		return false;
	}
	return true;
}

std::string PowerShellLiteral(const std::filesystem::path& path)
{
	std::string value = PathToUtf8(path);
	std::string escaped;
	escaped.reserve(value.size() + 8);
	for (char ch : value) {
		escaped.push_back(ch);
		if (ch == '\'') {
			escaped.push_back('\'');
		}
	}
	return "'" + escaped + "'";
}

bool ExtractArchive(
	const std::filesystem::path& archive,
	const std::filesystem::path& destination,
	std::string& outError)
{
	const std::string command = "Expand-Archive -LiteralPath " + PowerShellLiteral(archive) +
		" -DestinationPath " + PowerShellLiteral(destination) + " -Force";
	const PowerShellRunResult result = PowerShellToolRunner::Run(command, PathToUtf8(destination.parent_path()), 120);
	if (!result.ok) {
		outError = !result.error.empty() ? result.error : !result.stdErr.empty() ? result.stdErr : "解压失败";
		return false;
	}
	return true;
}

void CleanupStaging(const std::filesystem::path& root)
{
	if (root.empty()) {
		return;
	}
	std::error_code ec;
	const auto cache = std::filesystem::weakly_canonical(GetAutoLinkerCacheDirectoryPath(), ec);
	const auto target = std::filesystem::weakly_canonical(root, ec);
	if (!ec && IsPathInside(target, cache)) {
		std::filesystem::remove_all(target, ec);
	}
}

bool PrepareGithubRepository(const std::string& sourceText, PreparedRepository& out, std::string& outError)
{
	out = {};
	if (!ParseGithubSource(sourceText, out.source, outError)) {
		return false;
	}
	json repositoryJson;
	const std::string repoApi = "https://api.github.com/repos/" + out.source.owner + "/" + out.source.repository;
	if (!FetchJson(repoApi, repositoryJson, outError) || !repositoryJson.is_object()) {
		outError = "读取 GitHub 仓库信息失败：" + outError;
		return false;
	}
	if (repositoryJson.value("private", false)) {
		outError = "首版仅支持公开 GitHub 仓库";
		return false;
	}
	if (out.source.reference.empty()) {
		out.source.reference = repositoryJson.value("default_branch", std::string());
	}
	if (out.source.reference.empty()) {
		outError = "无法确定 GitHub 默认分支";
		return false;
	}
	json commitJson;
	std::string commitError;
	if (FetchJson(repoApi + "/commits/" + UrlEncode(out.source.reference), commitJson, commitError) &&
		commitJson.is_object()) {
		out.commit = commitJson.value("sha", std::string());
	}
	out.stagingRoot = GetAutoLinkerCacheDirectoryPath() /
		std::format(L"SkillInstall.{}.{}.{}", GetCurrentProcessId(), GetTickCount64(), g_tempCounter.fetch_add(1));
	const auto archive = out.stagingRoot / "repository.zip";
	const auto extract = out.stagingRoot / "extract";
	std::error_code ec;
	std::filesystem::create_directories(extract, ec);
	if (ec) {
		outError = "创建技能安装临时目录失败：" + ec.message();
		return false;
	}
	const std::string archiveUrl = "https://codeload.github.com/" + out.source.owner + "/" +
		out.source.repository + "/zip/" + UrlEncode(out.source.reference);
	auto response = PerformGetRequest(archiveUrl, "User-Agent: AutoLinker-SkillManager\r\n", 120000, false, false);
	if (response.second != 200 || response.first.size() < 4 || response.first[0] != 'P' || response.first[1] != 'K') {
		outError = std::format("下载 GitHub 仓库失败：HTTP {}", response.second);
		CleanupStaging(out.stagingRoot);
		return false;
	}
	if (response.first.size() > kMaxArchiveBytes) {
		outError = "GitHub 仓库压缩包超过 100 MiB 限制";
		CleanupStaging(out.stagingRoot);
		return false;
	}
	if (!WriteBinaryFile(archive, response.first, outError) || !ExtractArchive(archive, extract, outError)) {
		CleanupStaging(out.stagingRoot);
		return false;
	}
	for (std::filesystem::directory_iterator it(extract, ec), end; it != end && !ec; it.increment(ec)) {
		if (it->is_directory(ec)) {
			out.repositoryRoot = it->path();
			break;
		}
	}
	if (out.repositoryRoot.empty()) {
		outError = "GitHub 压缩包中没有仓库目录";
		CleanupStaging(out.stagingRoot);
		return false;
	}
	SkillConfig emptyConfig;
	DiscoverUnderRoot(out.repositoryRoot, AISkillScope::User, emptyConfig, out.candidates);
	for (auto& candidate : out.candidates) {
		candidate.sourceRepository = RepositoryName(out.source);
		candidate.sourceRef = out.source.reference;
		candidate.sourceCommit = out.commit;
		candidate.sourcePath = PathToUtf8(candidate.directory.lexically_relative(out.repositoryRoot));
		std::replace(candidate.sourcePath.begin(), candidate.sourcePath.end(), '\\', '/');
	}
	if (!out.source.subpath.empty()) {
		std::string prefix = out.source.subpath;
		std::replace(prefix.begin(), prefix.end(), '\\', '/');
		while (!prefix.empty() && prefix.back() == '/') {
			prefix.pop_back();
		}
		std::erase_if(out.candidates, [&prefix](const AISkillInfo& candidate) {
			return candidate.sourcePath != prefix && !candidate.sourcePath.starts_with(prefix + "/");
		});
	}
	if (out.candidates.empty()) {
		outError = "仓库指定范围内未找到有效的 SKILL.md";
		CleanupStaging(out.stagingRoot);
		return false;
	}
	return true;
}

json CandidatesJson(const std::vector<AISkillInfo>& candidates)
{
	json result = json::array();
	for (const auto& candidate : candidates) {
		result.push_back({
			{"name", candidate.name},
			{"description", candidate.description},
			{"repository_path", candidate.sourcePath}
		});
	}
	return result;
}

const AISkillInfo* SelectInstallCandidate(
	const PreparedRepository& repository,
	const std::string& selectedPath,
	const std::string& expectedSkillId)
{
	if (!selectedPath.empty()) {
		std::string normalized = selectedPath;
		std::replace(normalized.begin(), normalized.end(), '\\', '/');
		for (const auto& candidate : repository.candidates) {
			if (ToLowerAscii(candidate.sourcePath) == ToLowerAscii(normalized)) {
				return &candidate;
			}
		}
	}
	if (!expectedSkillId.empty()) {
		const std::string expected = ToLowerAscii(expectedSkillId);
		const AISkillInfo* match = nullptr;
		for (const auto& candidate : repository.candidates) {
			const std::string directoryName = ToLowerAscii(PathToUtf8(candidate.directory.filename()));
			if (ToLowerAscii(candidate.name) == expected || directoryName == expected) {
				if (match != nullptr) {
					return nullptr;
				}
				match = &candidate;
			}
		}
		if (match != nullptr) {
			return match;
		}
	}
	return repository.candidates.size() == 1 ? &repository.candidates.front() : nullptr;
}

long long CurrentUnixMs()
{
	return std::chrono::duration_cast<std::chrono::milliseconds>(
		std::chrono::system_clock::now().time_since_epoch()).count();
}

bool WriteSourceMetadata(const std::filesystem::path& directory, const AISkillInfo& source, std::string& outError)
{
	const json metadata = {
		{"version", 1},
		{"repository", source.sourceRepository},
		{"ref", source.sourceRef},
		{"path", source.sourcePath},
		{"commit", source.sourceCommit},
		{"installed_at_unix_ms", CurrentUnixMs()}
	};
	return WriteFileBytesAtomic(directory / kSourceMetadataFileName, metadata.dump(2), outError);
}

bool ValidateInstallTree(const std::filesystem::path& source, std::string& outError)
{
	const DWORD rootAttributes = GetFileAttributesW(source.c_str());
	if (rootAttributes == INVALID_FILE_ATTRIBUTES || (rootAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
		outError = "技能来源不是可读取的目录：" + PathToUtf8(source);
		return false;
	}
	if ((rootAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
		outError = "技能来源目录是符号链接或目录联接：" + PathToUtf8(source);
		return false;
	}
	std::error_code ec;
	std::filesystem::recursive_directory_iterator it(
		source, std::filesystem::directory_options::skip_permission_denied, ec), end;
	if (ec) {
		outError = "无法检查技能目录：" + ec.message();
		return false;
	}
	for (; it != end; it.increment(ec)) {
		if (ec) {
			outError = "检查技能目录失败：" + ec.message();
			return false;
		}
		const DWORD attributes = GetFileAttributesW(it->path().c_str());
		if (attributes == INVALID_FILE_ATTRIBUTES) {
			outError = "无法读取技能文件属性：" + PathToUtf8(it->path());
			return false;
		}
		if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
			outError = "技能目录包含不受支持的符号链接或目录联接：" + PathToUtf8(it->path());
			return false;
		}
	}
	return true;
}

bool ReplaceDirectoryAtomically(
	const std::filesystem::path& source,
	const std::filesystem::path& target,
	bool allowReplace,
	std::string& outError,
	bool stripSourceMetadata = false)
{
	if (!ValidateInstallTree(source, outError)) {
		return false;
	}
	std::error_code ec;
	std::filesystem::create_directories(target.parent_path(), ec);
	if (ec) {
		outError = "创建技能目录失败：" + ec.message();
		return false;
	}
	const std::filesystem::path staged = target.parent_path() /
		(target.filename().wstring() + std::format(L".install.{}.{}", GetCurrentProcessId(), g_tempCounter.fetch_add(1)));
	std::filesystem::copy(
		source, staged,
		std::filesystem::copy_options::recursive,
		ec);
	if (ec) {
		outError = "复制技能目录失败：" + ec.message();
		std::filesystem::remove_all(staged, ec);
		return false;
	}
	if (stripSourceMetadata) {
		std::filesystem::remove_all(staged / kSourceMetadataFileName, ec);
		if (ec) {
			outError = "清理本地技能来源标记失败：" + ec.message();
			std::filesystem::remove_all(staged, ec);
			return false;
		}
	}
	const bool targetExists = std::filesystem::exists(target, ec);
	if (targetExists && !allowReplace) {
		outError = "目标技能目录已存在，请确认替换或先卸载";
		std::filesystem::remove_all(staged, ec);
		return false;
	}
	const std::filesystem::path backup = target.parent_path() /
		(target.filename().wstring() + std::format(L".backup.{}.{}", GetCurrentProcessId(), g_tempCounter.fetch_add(1)));
	if (targetExists) {
		std::filesystem::rename(target, backup, ec);
		if (ec) {
			outError = "暂存旧技能失败：" + ec.message();
			std::filesystem::remove_all(staged, ec);
			return false;
		}
	}
	std::filesystem::rename(staged, target, ec);
	if (ec) {
		outError = "启用新技能目录失败：" + ec.message();
		if (targetExists) {
			std::error_code rollbackEc;
			std::filesystem::rename(backup, target, rollbackEc);
		}
		std::filesystem::remove_all(staged, ec);
		return false;
	}
	if (targetExists) {
		std::filesystem::remove_all(backup, ec);
	}
	return true;
}

void CleanupPreparedLocalSkills(PreparedLocalSkills& prepared)
{
	AISkillLocalPackage::Cleanup(prepared.source, GetAutoLinkerCacheDirectoryPath());
	prepared.candidates.clear();
}

bool PrepareLocalSkills(const std::string& sourceText, PreparedLocalSkills& out, std::string& outError)
{
	out = {};
	const std::string trimmed = TrimAscii(sourceText);
	if (!AISkillLocalPackage::Prepare(
		Utf8ToPath(trimmed), GetAutoLinkerCacheDirectoryPath(), out.source, outError)) {
		return false;
	}
	if (out.source.kind == AISkillLocalSourceKind::ZipArchive &&
		!ValidateInstallTree(out.source.scanRoot, outError)) {
		CleanupPreparedLocalSkills(out);
		return false;
	}
	if (!out.source.directSkillFile.empty()) {
		AISkillInfo candidate;
		if (ParseSkillFile(out.source.directSkillFile, candidate)) {
			out.candidates.push_back(std::move(candidate));
		}
		else {
			outError = "SKILL.md 无效：" + candidate.error;
			CleanupPreparedLocalSkills(out);
			return false;
		}
	}
	else {
		SkillConfig emptyConfig;
		DiscoverUnderRoot(out.source.scanRoot, AISkillScope::User, emptyConfig, out.candidates);
		std::erase_if(out.candidates, [](const AISkillInfo& candidate) { return !candidate.valid; });
	}
	for (auto& candidate : out.candidates) {
		candidate.sourcePath = PathToUtf8(candidate.directory.lexically_relative(out.source.scanRoot));
		std::replace(candidate.sourcePath.begin(), candidate.sourcePath.end(), '\\', '/');
		if (candidate.sourcePath.empty() || candidate.sourcePath == ".") {
			candidate.sourcePath = ".";
		}
	}
	std::stable_sort(out.candidates.begin(), out.candidates.end(), [](const auto& left, const auto& right) {
		return ToLowerAscii(left.sourcePath) < ToLowerAscii(right.sourcePath);
	});
	if (out.candidates.empty()) {
		outError = "本地来源中未找到有效的 SKILL.md";
		CleanupPreparedLocalSkills(out);
		return false;
	}
	return true;
}

json LocalCandidatesJson(const PreparedLocalSkills& prepared)
{
	json result = json::array();
	for (const auto& candidate : prepared.candidates) {
		result.push_back({
			{"name", candidate.name},
			{"description", candidate.description},
			{"candidate_path", candidate.sourcePath},
			{"skill_file", prepared.source.kind == AISkillLocalSourceKind::ZipArchive
				? (candidate.sourcePath == "." ? "SKILL.md" : candidate.sourcePath + "/SKILL.md")
				: PathToUtf8(candidate.skillFile)}
		});
	}
	return result;
}

json LocalInspectionJson(const PreparedLocalSkills& prepared)
{
	return {
		{"ok", true},
		{"source_path", PathToUtf8(prepared.source.sourcePath)},
		{"source_kind", AISkillLocalPackage::KindName(prepared.source.kind)},
		{"reference_allowed", prepared.source.kind != AISkillLocalSourceKind::ZipArchive},
		{"candidates", LocalCandidatesJson(prepared)}
	};
}

const AISkillInfo* SelectLocalCandidate(
	const PreparedLocalSkills& prepared,
	std::string selectedPath)
{
	selectedPath = TrimAscii(std::move(selectedPath));
	std::replace(selectedPath.begin(), selectedPath.end(), '\\', '/');
	while (selectedPath.starts_with("./")) {
		selectedPath.erase(0, 2);
	}
	if (selectedPath.empty()) {
		return prepared.candidates.size() == 1 ? &prepared.candidates.front() : nullptr;
	}
	for (const auto& candidate : prepared.candidates) {
		if (ToLowerAscii(candidate.sourcePath) == ToLowerAscii(selectedPath)) {
			return &candidate;
		}
	}
	return nullptr;
}

bool SameReferenceRegistration(
	const SkillConfig::ExternalReference& left,
	const SkillConfig::ExternalReference& right)
{
	return left.scope == right.scope &&
		(left.scope == AISkillScope::User ||
			NormalizePathKey(left.projectDirectory) == NormalizePathKey(right.projectDirectory));
}

bool AddExternalReference(
	SkillConfig& config,
	const SkillConfig::ExternalReference& requested,
	bool& outAlreadyRegistered,
	std::string& outError)
{
	outAlreadyRegistered = false;
	const std::string requestedKey = NormalizePathKey(requested.skillFile);
	for (const auto& existing : config.externalReferences) {
		if (NormalizePathKey(existing.skillFile) != requestedKey) {
			continue;
		}
		if (SameReferenceRegistration(existing, requested)) {
			outAlreadyRegistered = true;
			return true;
		}
		outError = "同一个 SKILL.md 已在其他作用域或项目中引用，不能重复登记";
		return false;
	}
	config.externalReferences.push_back(requested);
	return true;
}

bool RemoveExternalReference(
	SkillConfig& config,
	const std::filesystem::path& skillFile,
	AISkillScope scope,
	const std::optional<std::filesystem::path>& projectDirectory)
{
	const std::string key = NormalizePathKey(skillFile);
	const size_t before = config.externalReferences.size();
	std::erase_if(config.externalReferences, [&](const SkillConfig::ExternalReference& reference) {
		if (NormalizePathKey(reference.skillFile) != key || reference.scope != scope) {
			return false;
		}
		return scope == AISkillScope::User || (projectDirectory &&
			NormalizePathKey(reference.projectDirectory) == NormalizePathKey(*projectDirectory));
	});
	return config.externalReferences.size() != before;
}

bool PathsOverlap(const std::filesystem::path& left, const std::filesystem::path& right)
{
	return IsPathInside(left, right) || IsPathInside(right, left);
}

const AISkillInfo* FindSkillByPath(const std::vector<AISkillInfo>& skills, const std::filesystem::path& skillFile)
{
	const std::string key = NormalizePathKey(skillFile);
	for (const auto& skill : skills) {
		if (NormalizePathKey(skill.skillFile) == key) {
			return &skill;
		}
	}
	return nullptr;
}

std::string ScopeText(AISkillScope scope)
{
	return scope == AISkillScope::Repo ? "project" : "global";
}

bool ParseSkillsShDetailDocument(const std::string& html, json& outDetails)
{
	constexpr std::string_view marker = "<script type=\"application/ld+json\">";
	constexpr std::string_view closing = "</script>";
	size_t offset = 0;
	while ((offset = html.find(marker, offset)) != std::string::npos) {
		const size_t jsonBegin = offset + marker.size();
		const size_t jsonEnd = html.find(closing, jsonBegin);
		if (jsonEnd == std::string::npos) {
			break;
		}
		json document = json::parse(html.substr(jsonBegin, jsonEnd - jsonBegin), nullptr, false);
		if (document.is_object() && document.value("@type", std::string()) == "SoftwareApplication" &&
			document.contains("description") && document["description"].is_string()) {
			outDetails = std::move(document);
			return true;
		}
		offset = jsonEnd + closing.size();
	}
	return false;
}

bool CreateZipFromDirectory(
	const std::filesystem::path& source,
	const std::filesystem::path& archive,
	std::string& outError)
{
	const std::string command =
		"Add-Type -AssemblyName System.IO.Compression.FileSystem; "
		"[System.IO.Compression.ZipFile]::CreateFromDirectory(" + PowerShellLiteral(source) + "," +
		PowerShellLiteral(archive) + ",[System.IO.Compression.CompressionLevel]::Optimal,$false)";
	const PowerShellRunResult result = PowerShellToolRunner::Run(command, PathToUtf8(source.parent_path()), 60);
	if (!result.ok) {
		outError = !result.error.empty() ? result.error : result.stdErr;
		return false;
	}
	return true;
}

bool CreateTraversalZip(const std::filesystem::path& archive, std::string& outError)
{
	const std::string command =
		"Add-Type -AssemblyName System.IO.Compression; "
		"$stream=[System.IO.File]::Open(" + PowerShellLiteral(archive) + ",[System.IO.FileMode]::Create); "
		"try { $zip=[System.IO.Compression.ZipArchive]::new($stream,[System.IO.Compression.ZipArchiveMode]::Create,$true); "
		"try { $entry=$zip.CreateEntry('../escape.txt'); $writer=[System.IO.StreamWriter]::new($entry.Open()); "
		"try { $writer.Write('escape') } finally { $writer.Dispose() } } finally { $zip.Dispose() } } "
		"finally { $stream.Dispose() }";
	const PowerShellRunResult result = PowerShellToolRunner::Run(command, PathToUtf8(archive.parent_path()), 60);
	if (!result.ok) {
		outError = !result.error.empty() ? result.error : result.stdErr;
		return false;
	}
	return true;
}

void AppendUInt16(std::string& bytes, const unsigned int value)
{
	bytes.push_back(static_cast<char>(value & 0xFF));
	bytes.push_back(static_cast<char>((value >> 8) & 0xFF));
}

void AppendUInt32(std::string& bytes, const unsigned long value)
{
	for (int shift = 0; shift < 32; shift += 8) {
		bytes.push_back(static_cast<char>((value >> shift) & 0xFF));
	}
}

std::string BuildDeclaredOversizeZip()
{
	constexpr unsigned long declaredSize = 512UL * 1024 * 1024 + 1;
	constexpr std::string_view name = "huge.bin";
	std::string bytes;
	AppendUInt32(bytes, 0x04034B50);
	AppendUInt16(bytes, 20);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt32(bytes, 0);
	AppendUInt32(bytes, 0);
	AppendUInt32(bytes, declaredSize);
	AppendUInt16(bytes, static_cast<unsigned int>(name.size()));
	AppendUInt16(bytes, 0);
	bytes.append(name);
	const unsigned long centralOffset = static_cast<unsigned long>(bytes.size());
	AppendUInt32(bytes, 0x02014B50);
	AppendUInt16(bytes, 20);
	AppendUInt16(bytes, 20);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt32(bytes, 0);
	AppendUInt32(bytes, 0);
	AppendUInt32(bytes, declaredSize);
	AppendUInt16(bytes, static_cast<unsigned int>(name.size()));
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt32(bytes, 0);
	AppendUInt32(bytes, 0);
	bytes.append(name);
	const unsigned long centralSize = static_cast<unsigned long>(bytes.size()) - centralOffset;
	AppendUInt32(bytes, 0x06054B50);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 0);
	AppendUInt16(bytes, 1);
	AppendUInt16(bytes, 1);
	AppendUInt32(bytes, centralSize);
	AppendUInt32(bytes, centralOffset);
	AppendUInt16(bytes, 0);
	return bytes;
}

} // namespace

namespace AISkillManager {

std::filesystem::path GetGlobalSkillsRoot()
{
	return GetAutoLinkerDirectoryPath() / "Skills";
}

std::optional<std::filesystem::path> GetProjectSkillsRoot()
{
	const std::string projectPath = TrimAscii(g_nowOpenSourceFilePath);
	if (projectPath.empty()) {
		return std::nullopt;
	}
	const auto sourcePath = LocalToPath(projectPath);
	if (sourcePath.empty() || sourcePath.parent_path().empty()) {
		return std::nullopt;
	}
	return sourcePath.parent_path() / ".agents" / "skills";
}

std::vector<AISkillInfo> DiscoverSkills()
{
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	return DiscoverSkillsUnlocked();
}

std::string BuildSettingsPayloadJson()
{
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	json skills = json::array();
	for (const auto& skill : DiscoverSkillsUnlocked()) {
		skills.push_back(SkillToJson(skill));
	}
	const auto projectRoot = GetProjectSkillsRoot();
	return json({
		{"ok", true},
		{"skills", std::move(skills)},
		{"global_root", PathToUtf8(GetGlobalSkillsRoot())},
		{"project_root", projectRoot ? PathToUtf8(*projectRoot) : std::string()},
		{"project_available", projectRoot.has_value()}
	}).dump();
}

std::string BuildRuntimePromptAddon(const std::string& latestUserMessage)
{
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	const auto skills = EffectiveSkills(DiscoverSkillsUnlocked());
	if (skills.empty()) {
		return {};
	}
	std::string prompt =
		"\n\n## Skills\n"
		"技能是存放在 SKILL.md 中的可复用工作说明。以下为本轮可用技能。\n"
		"当用户用 $技能名 明确点名，或任务明显符合某项描述时，必须使用对应技能；未明确点名时先调用 read_skill_resource 读取完整 SKILL.md。技能只对当前用户轮次生效。引用的相对资源必须继续通过 read_skill_resource 读取，且只读取 SKILL.md 明确要求的资源。技能脚本只能通过需要用户确认的本机命令工具执行。\n"
		"### Available skills\n";
	for (const auto& skill : skills) {
		std::string line = "- " + skill.name + ": " + skill.description + " (" +
			ScopeText(skill.scope) + "; " + PathToUtf8(skill.skillFile) + ")\n";
		if (prompt.size() + line.size() > kPromptCatalogBudgetBytes) {
			prompt += "- 其余技能因上下文预算未列出，可在设置页停用不常用技能。\n";
			break;
		}
		prompt += line;
	}
	const auto mentions = ExtractSkillMentions(latestUserMessage);
	for (const auto& skill : skills) {
		if (!mentions.contains(ToLowerAscii(skill.name))) {
			continue;
		}
		std::string contents;
		std::string error;
		if (!ReadFileBytes(skill.skillFile, kMaxSkillFileBytes, contents, error)) {
			prompt += "\n<skill-error name=\"" + skill.name + "\">" + error + "</skill-error>\n";
			continue;
		}
		prompt += "\n<skill>\n<name>" + skill.name + "</name>\n<path>" +
			PathToUtf8(skill.skillFile) + "</path>\n" + contents + "\n</skill>\n";
	}
	return prompt;
}

std::string ExecuteReadSkillResourceTool(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	json args = json::parse(argumentsJson, nullptr, false);
	if (!args.is_object()) {
		return R"({"ok":false,"error":"arguments must be an object"})";
	}
	const std::string skillName = TrimAscii(args.value("skill_name", std::string()));
	std::string relativePath = TrimAscii(args.value("relative_path", std::string("SKILL.md")));
	const long long byteOffset = (std::max)(0LL, args.value("byte_offset", 0LL));
	const size_t maxBytes = static_cast<size_t>((std::clamp)(
		args.value("max_bytes", static_cast<long long>(64 * 1024)), 4096LL,
		static_cast<long long>(kMaxReadResourceBytes)));
	if (skillName.empty()) {
		return R"({"ok":false,"error":"skill_name is required"})";
	}
	if (relativePath.empty()) {
		relativePath = "SKILL.md";
	}
	const auto skills = EffectiveSkills(DiscoverSkills());
	const AISkillInfo* selected = nullptr;
	for (const auto& skill : skills) {
		if (ToLowerAscii(skill.name) == ToLowerAscii(skillName)) {
			selected = &skill;
			break;
		}
	}
	if (selected == nullptr) {
		return json({{"ok", false}, {"error", "skill not found or disabled"}, {"skill_name", skillName}}).dump();
	}
	const auto requested = selected->directory / Utf8ToPath(relativePath);
	std::error_code ec;
	const auto canonical = std::filesystem::weakly_canonical(requested, ec);
	if (ec || !IsPathInside(canonical, selected->directory) || !std::filesystem::is_regular_file(canonical, ec)) {
		return json({{"ok", false}, {"error", "resource path is outside the skill or is not a file"}}).dump();
	}
	const auto size = std::filesystem::file_size(canonical, ec);
	if (ec || static_cast<unsigned long long>(byteOffset) > size) {
		return json({{"ok", false}, {"error", "invalid byte_offset"}}).dump();
	}
	std::ifstream input(canonical, std::ios::binary);
	if (!input.is_open()) {
		return json({{"ok", false}, {"error", "failed to open skill resource"}}).dump();
	}
	input.seekg(byteOffset);
	const size_t remaining = static_cast<size_t>((std::min<unsigned long long>)(size - byteOffset, maxBytes));
	std::string content(remaining, '\0');
	input.read(content.data(), static_cast<std::streamsize>(content.size()));
	content.resize(static_cast<size_t>(input.gcount()));
	if (content.find('\0') != std::string::npos) {
		return json({{"ok", false}, {"error", "binary skill resources cannot be returned as text"}}).dump();
	}
	if (byteOffset == 0 && content.size() >= 3 && static_cast<unsigned char>(content[0]) == 0xEF &&
		static_cast<unsigned char>(content[1]) == 0xBB && static_cast<unsigned char>(content[2]) == 0xBF) {
		content.erase(0, 3);
	}
	const long long nextOffset = byteOffset + static_cast<long long>(input.gcount());
	outOk = true;
	return json({
		{"ok", true},
		{"skill_name", selected->name},
		{"relative_path", PathToUtf8(canonical.lexically_relative(selected->directory))},
		{"content", content},
		{"byte_offset", byteOffset},
		{"next_byte_offset", nextOffset < static_cast<long long>(size) ? json(nextOffset) : json(nullptr)},
		{"total_bytes", size}
	}).dump(-1, ' ', false, json::error_handler_t::replace);
}

std::string SearchSkillsSh(const std::string& query)
{
	const std::string trimmed = TrimAscii(query);
	if (trimmed.size() < 2) {
		return json({{"ok", false}, {"error", "请输入至少两个字符"}}).dump();
	}
	const auto response = PerformGetRequest(
		"https://skills.sh/api/search?q=" + UrlEncode(trimmed) + "&limit=100",
		"Accept: application/json\r\nUser-Agent: AutoLinker-SkillManager\r\n",
		30000, false, false);
	if (response.second != 200) {
		return json({{"ok", false}, {"error", std::format("skills.sh 搜索失败：HTTP {}", response.second)}}).dump();
	}
	json root = json::parse(response.first, nullptr, false);
	if (!root.is_object() || !root.contains("skills") || !root["skills"].is_array()) {
		return json({{"ok", false}, {"error", "skills.sh 返回格式无效"}}).dump();
	}
	json results = json::array();
	for (const auto& item : root["skills"]) {
		if (!item.is_object()) {
			continue;
		}
		results.push_back({
			{"name", item.value("name", std::string())},
			{"skill_id", item.value("skillId", std::string())},
			{"source", item.value("source", std::string())},
			{"installs", item.value("installs", 0LL)}
		});
	}
	return json({{"ok", true}, {"query", trimmed}, {"results", std::move(results)}}).dump();
}

std::string GetSkillsShSkillDetails(const std::string& source, const std::string& skillId)
{
	const auto sourceParts = SplitPath(TrimAscii(source));
	const auto skillParts = SplitPath(TrimAscii(skillId));
	if (sourceParts.size() != 2 || skillParts.empty() ||
		!IsGithubIdentifier(sourceParts[0]) || !IsGithubIdentifier(sourceParts[1]) ||
		!std::all_of(skillParts.begin(), skillParts.end(), IsGithubIdentifier)) {
		return json({{"ok", false}, {"error", "skills.sh 技能来源或标识无效"}}).dump();
	}
	std::string detailUrl = "https://skills.sh/" + UrlEncode(sourceParts[0]) + "/" + UrlEncode(sourceParts[1]);
	for (const auto& part : skillParts) {
		detailUrl += "/" + UrlEncode(part);
	}
	const auto response = PerformGetRequest(
		detailUrl,
		"Accept: text/html,application/xhtml+xml\r\nUser-Agent: AutoLinker-SkillManager\r\n",
		30000, false, false);
	if (response.second != 200) {
		return json({{"ok", false}, {"error", std::format("读取 skills.sh 技能详情失败：HTTP {}", response.second)}}).dump();
	}
	if (response.first.size() > 8 * 1024 * 1024) {
		return json({{"ok", false}, {"error", "skills.sh 技能详情页超过大小限制"}}).dump();
	}
	json details;
	if (!ParseSkillsShDetailDocument(response.first, details)) {
		return json({{"ok", false}, {"error", "skills.sh 技能详情中没有可用的结构化介绍"}}).dump();
	}
	std::string description = details.value("description", std::string());
	if (description.size() > 16 * 1024) {
		description.resize(16 * 1024);
	}
	long long installs = 0;
	if (details.contains("interactionStatistic") && details["interactionStatistic"].is_object() &&
		details["interactionStatistic"].contains("userInteractionCount")) {
		const auto& count = details["interactionStatistic"]["userInteractionCount"];
		if (count.is_number_integer() || count.is_number_unsigned()) {
			installs = count.get<long long>();
		}
	}
	return json({
		{"ok", true},
		{"name", details.value("name", skillParts.back())},
		{"description", std::move(description)},
		{"source", sourceParts[0] + "/" + sourceParts[1]},
		{"skill_id", skillId},
		{"installs", installs},
		{"detail_url", detailUrl}
	}).dump(-1, ' ', false, json::error_handler_t::replace);
}

std::string InspectGitHubRepository(const std::string& source)
{
	PreparedRepository repository;
	std::string error;
	if (!PrepareGithubRepository(source, repository, error)) {
		return json({{"ok", false}, {"error", error}}).dump();
	}
	const json result = {
		{"ok", true},
		{"repository", RepositoryName(repository.source)},
		{"ref", repository.source.reference},
		{"commit", repository.commit},
		{"candidates", CandidatesJson(repository.candidates)}
	};
	CleanupStaging(repository.stagingRoot);
	return result.dump();
}

std::string InspectLocalSource(const std::string& sourcePath)
{
	PreparedLocalSkills prepared;
	std::string error;
	if (!PrepareLocalSkills(sourcePath, prepared, error)) {
		return json({{"ok", false}, {"error", error}}).dump();
	}
	const json result = LocalInspectionJson(prepared);
	CleanupPreparedLocalSkills(prepared);
	return result.dump();
}

std::string InstallFromGitHub(
	const std::string& source,
	AISkillScope scope,
	const std::string& selectedRepositoryPath,
	const std::string& expectedSkillId,
	bool allowReplace)
{
	const auto requestedProjectRoot = scope == AISkillScope::Repo ? GetProjectSkillsRoot() : std::nullopt;
	if (scope == AISkillScope::Repo && !requestedProjectRoot) {
		return json({{"ok", false}, {"error", "当前没有打开可定位的易语言项目"}}).dump();
	}
	PreparedRepository repository;
	std::string error;
	if (!PrepareGithubRepository(source, repository, error)) {
		return json({{"ok", false}, {"error", error}}).dump();
	}
	const AISkillInfo* selected = SelectInstallCandidate(repository, selectedRepositoryPath, expectedSkillId);
	if (selected == nullptr) {
		const json result = {
			{"ok", false},
			{"selection_required", true},
			{"error", "仓库中包含多个技能，请选择要安装的技能"},
			{"repository", RepositoryName(repository.source)},
			{"candidates", CandidatesJson(repository.candidates)}
		};
		CleanupStaging(repository.stagingRoot);
		return result.dump();
	}
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	const auto root = scope == AISkillScope::Repo ? *requestedProjectRoot : GetGlobalSkillsRoot();
	const auto target = root / Utf8ToPath(SanitizeDirectoryName(selected->name));
	if (std::filesystem::exists(target)) {
		AISkillInfo existing;
		if (ParseSkillFile(target / "SKILL.md", existing)) {
			LoadSourceMetadata(existing);
			if (ToLowerAscii(existing.sourceRepository) == ToLowerAscii(selected->sourceRepository)) {
				allowReplace = true;
			}
		}
	}
	if (!ReplaceDirectoryAtomically(selected->directory, target, allowReplace, error)) {
		CleanupStaging(repository.stagingRoot);
		return json({
			{"ok", false}, {"error", error}, {"conflict", std::filesystem::exists(target)},
			{"target", PathToUtf8(target)}
		}).dump();
	}
	AISkillInfo installedSource = *selected;
	if (!WriteSourceMetadata(target, installedSource, error)) {
		CleanupStaging(repository.stagingRoot);
		return json({{"ok", false}, {"error", "技能已复制，但写入来源信息失败：" + error}}).dump();
	}
	CleanupStaging(repository.stagingRoot);
	return json({
		{"ok", true},
		{"message", "技能安装完成"},
		{"name", selected->name},
		{"scope", ScopeText(scope)},
		{"skill_file", PathToUtf8(target / "SKILL.md")},
		{"commit", selected->sourceCommit}
	}).dump();
}

std::string InstallFromLocal(
	const std::string& sourcePath,
	AISkillScope scope,
	const std::string& selectedCandidatePath,
	AISkillLocalInstallMode mode,
	bool allowReplace)
{
	const auto requestedProjectRoot = scope == AISkillScope::Repo ? GetProjectSkillsRoot() : std::nullopt;
	const auto projectDirectory = scope == AISkillScope::Repo ? CurrentProjectDirectory() : std::nullopt;
	if (scope == AISkillScope::Repo && (!requestedProjectRoot || !projectDirectory)) {
		return json({{"ok", false}, {"error", "当前没有打开可定位的易语言项目"}}).dump();
	}
	PreparedLocalSkills prepared;
	std::string error;
	if (!PrepareLocalSkills(sourcePath, prepared, error)) {
		return json({{"ok", false}, {"error", error}}).dump();
	}
	const AISkillInfo* selected = SelectLocalCandidate(prepared, selectedCandidatePath);
	if (selected == nullptr) {
		json result = LocalInspectionJson(prepared);
		result["ok"] = false;
		result["selection_required"] = true;
		result["error"] = "本地来源包含多个技能，请选择要添加的技能";
		CleanupPreparedLocalSkills(prepared);
		return result.dump();
	}
	const std::string selectedName = selected->name;
	const std::filesystem::path selectedDirectory = selected->directory;
	const std::filesystem::path selectedSkillFile = selected->skillFile;
	if (mode == AISkillLocalInstallMode::Reference) {
		if (prepared.source.kind == AISkillLocalSourceKind::ZipArchive) {
			CleanupPreparedLocalSkills(prepared);
			return json({{"ok", false}, {"error", "ZIP 来源只能复制安装，不能引用临时解压路径"}}).dump();
		}
		std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
		const auto globalRoot = GetGlobalSkillsRoot();
		const auto currentProjectRoot = GetProjectSkillsRoot();
		if (IsPathInside(selectedSkillFile, globalRoot) ||
			(currentProjectRoot && IsPathInside(selectedSkillFile, *currentProjectRoot))) {
			CleanupPreparedLocalSkills(prepared);
			return json({{"ok", false}, {"error", "该技能已经位于 AutoLinker 管理目录中，无需登记外部引用"}}).dump();
		}
		SkillConfig config = LoadConfig();
		SkillConfig::ExternalReference reference;
		reference.skillFile = selectedSkillFile;
		reference.scope = scope;
		reference.projectDirectory = projectDirectory.value_or(std::filesystem::path());
		reference.addedAtUnixMs = CurrentUnixMs();
		bool alreadyRegistered = false;
		if (!AddExternalReference(config, reference, alreadyRegistered, error)) {
			CleanupPreparedLocalSkills(prepared);
			return json({{"ok", false}, {"error", error}}).dump();
		}
		if (!alreadyRegistered && !SaveConfig(config, error)) {
			CleanupPreparedLocalSkills(prepared);
			return json({{"ok", false}, {"error", "保存技能引用失败：" + error}}).dump();
		}
		CleanupPreparedLocalSkills(prepared);
		return json({
			{"ok", true},
			{"message", alreadyRegistered ? "该原路径引用已经存在" : "已添加原路径引用"},
			{"name", selectedName},
			{"scope", ScopeText(scope)},
			{"skill_file", PathToUtf8(selectedSkillFile)},
			{"external_reference", true},
			{"already_registered", alreadyRegistered}
		}).dump();
	}
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	const auto root = scope == AISkillScope::Repo ? *requestedProjectRoot : GetGlobalSkillsRoot();
	const auto target = root / Utf8ToPath(SanitizeDirectoryName(selectedName));
	if (PathsOverlap(selectedDirectory, target)) {
		CleanupPreparedLocalSkills(prepared);
		return json({{"ok", false}, {"error", "技能来源目录与安装目标重叠，拒绝自复制"}}).dump();
	}
	if (!ReplaceDirectoryAtomically(selectedDirectory, target, allowReplace, error, true)) {
		const bool conflict = std::filesystem::exists(target);
		CleanupPreparedLocalSkills(prepared);
		return json({
			{"ok", false}, {"error", error}, {"conflict", conflict}, {"target", PathToUtf8(target)}
		}).dump();
	}
	CleanupPreparedLocalSkills(prepared);
	return json({
		{"ok", true},
		{"message", "本地技能已复制安装"},
		{"name", selectedName},
		{"scope", ScopeText(scope)},
		{"skill_file", PathToUtf8(target / "SKILL.md")},
		{"external_reference", false}
	}).dump();
}

std::string UpdateInstalledSkill(const std::string& skillFilePath)
{
	AISkillInfo skill;
	{
		std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
		const auto skills = DiscoverSkillsUnlocked();
		const AISkillInfo* found = FindSkillByPath(skills, Utf8ToPath(skillFilePath));
		if (found == nullptr || found->sourceRepository.empty()) {
			return json({{"ok", false}, {"error", "该技能不是由 AutoLinker 从 GitHub 安装，无法自动更新"}}).dump();
		}
		skill = *found;
	}
	const std::string source = "https://github.com/" + skill.sourceRepository + "/tree/" +
		skill.sourceRef + (skill.sourcePath.empty() ? std::string() : "/" + skill.sourcePath);
	return InstallFromGitHub(source, skill.scope, skill.sourcePath, skill.name, true);
}

std::string RemoveInstalledSkill(const std::string& skillFilePath)
{
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	const auto skills = DiscoverSkillsUnlocked();
	const AISkillInfo* skill = FindSkillByPath(skills, Utf8ToPath(skillFilePath));
	if (skill == nullptr) {
		return json({{"ok", false}, {"error", "未找到技能"}}).dump();
	}
	if (skill->externalReference) {
		SkillConfig config = LoadConfig();
		if (!RemoveExternalReference(config, skill->skillFile, skill->scope, CurrentProjectDirectory())) {
			return json({{"ok", false}, {"error", "未找到对应的原路径引用配置"}}).dump();
		}
		config.disabledPaths.erase(NormalizePathKey(skill->skillFile));
		std::string saveError;
		if (!SaveConfig(config, saveError)) {
			return json({{"ok", false}, {"error", "移除技能引用失败：" + saveError}}).dump();
		}
		return json({{"ok", true}, {"message", "技能引用已移除，原路径文件未被删除"}}).dump();
	}
	const auto globalRoot = GetGlobalSkillsRoot();
	const auto projectRoot = GetProjectSkillsRoot();
	if (!IsPathStrictlyInside(skill->directory, globalRoot) &&
		(!projectRoot || !IsPathStrictlyInside(skill->directory, *projectRoot))) {
		return json({{"ok", false}, {"error", "拒绝删除技能根目录本身或根目录外的路径"}}).dump();
	}
	std::error_code ec;
	std::filesystem::remove_all(skill->directory, ec);
	if (ec) {
		return json({{"ok", false}, {"error", "删除技能失败：" + ec.message()}}).dump();
	}
	SkillConfig config = LoadConfig();
	config.disabledPaths.erase(NormalizePathKey(skill->skillFile));
	std::string saveError;
	SaveConfig(config, saveError);
	return json({{"ok", true}, {"message", "技能已卸载"}}).dump();
}

bool SetSkillEnabled(const std::string& skillFilePath, bool enabled, std::string& outError)
{
	std::lock_guard<std::recursive_mutex> guard(g_skillMutex);
	const auto skills = DiscoverSkillsUnlocked();
	const AISkillInfo* skill = FindSkillByPath(skills, Utf8ToPath(skillFilePath));
	if (skill == nullptr) {
		outError = "未找到技能";
		return false;
	}
	SkillConfig config = LoadConfig();
	const std::string key = NormalizePathKey(skill->skillFile);
	if (enabled) {
		config.disabledPaths.erase(key);
	}
	else {
		config.disabledPaths.insert(key);
	}
	return SaveConfig(config, outError);
}

std::string BuildSelfTestJson()
{
	const auto tempRoot = GetAutoLinkerCacheDirectoryPath() /
		std::format(L"SkillSelfTest.{}.{}", GetCurrentProcessId(), g_tempCounter.fetch_add(1));
	std::error_code ec;
	std::filesystem::create_directories(tempRoot, ec);
	std::string error;
	auto writeSkill = [&](const std::filesystem::path& directory, const std::string& name) {
		const std::string contents = "\xEF\xBB\xBF---\r\nname: " + name +
			"\r\ndescription: >\r\n  Test skill for parser\r\nmetadata:\r\n  short-description: Demo\r\n---\r\nBody\r\n";
		return WriteFileBytesAtomic(directory / "SKILL.md", contents, error);
	};
	const auto sourceRoot = tempRoot / "source";
	const auto oneDirectory = sourceRoot / "one";
	const auto twoDirectory = sourceRoot / "nested" / "two";
	const auto threeDirectory = sourceRoot / "three";
	const bool skillsWritten = writeSkill(oneDirectory, "one-skill") &&
		writeSkill(twoDirectory, "two-skill") && writeSkill(threeDirectory, "three-skill");
	AISkillInfo valid;
	const bool parseOk = skillsWritten && ParseSkillFile(oneDirectory / "SKILL.md", valid) &&
		valid.name == "one-skill" && valid.description == "Test skill for parser" &&
		valid.shortDescription == "Demo";
	const bool containmentOk = IsPathInside(tempRoot / "demo" / "SKILL.md", tempRoot) &&
		!IsPathInside(tempRoot.parent_path() / "outside.txt", tempRoot) &&
		IsPathStrictlyInside(tempRoot / "demo", tempRoot) &&
		!IsPathStrictlyInside(tempRoot, tempRoot);

	const auto configV1Path = tempRoot / "config-v1.json";
	const bool configV1Written = WriteFileBytesAtomic(
		configV1Path, R"({"version":1,"disabled_paths":["Legacy/Skill"]})", error);
	const SkillConfig configV1 = LoadConfigFromPath(configV1Path);
	const bool configV1Ok = configV1Written && configV1.disabledPaths.contains("legacy/skill") &&
		configV1.externalReferences.empty();
	SkillConfig configV2;
	configV2.disabledPaths.insert("disabled/example");
	configV2.externalReferences = {
		{oneDirectory / "SKILL.md", AISkillScope::User, {}, 100},
		{twoDirectory / "SKILL.md", AISkillScope::Repo, tempRoot / "project-a", 200}
	};
	const auto configV2Path = tempRoot / "config-v2.json";
	const bool configV2Saved = SaveConfigToPath(configV2, configV2Path, error);
	const SkillConfig configV2Loaded = LoadConfigFromPath(configV2Path);
	std::string configV2Text;
	std::string readError;
	const bool configV2Read = ReadFileBytes(configV2Path, 2 * 1024 * 1024, configV2Text, readError);
	const json configV2Json = json::parse(configV2Text, nullptr, false);
	const bool configV2Ok = configV2Saved && configV2Read && configV2Json.is_object() &&
		configV2Json.value("version", 0) == 2 && configV2Loaded.disabledPaths.contains("disabled/example") &&
		configV2Loaded.externalReferences.size() == 2 &&
		std::ranges::any_of(configV2Loaded.externalReferences, [](const auto& reference) {
			return reference.scope == AISkillScope::Repo;
		});

	SkillConfig filterConfig = configV2;
	filterConfig.externalReferences.push_back(
		{threeDirectory / "SKILL.md", AISkillScope::Repo, tempRoot / "project-b", 300});
	std::vector<AISkillInfo> filteredReferences;
	DiscoverExternalReferences(filterConfig, tempRoot / "project-a", filteredReferences);
	const bool projectFilterOk = filteredReferences.size() == 2 &&
		std::ranges::any_of(filteredReferences, [](const AISkillInfo& skill) {
			return skill.scope == AISkillScope::User && skill.name == "one-skill";
		}) && std::ranges::any_of(filteredReferences, [](const AISkillInfo& skill) {
			return skill.scope == AISkillScope::Repo && skill.name == "two-skill";
		});
	bool alreadyRegistered = false;
	std::string duplicateError;
	const bool idempotentReferenceOk = AddExternalReference(
		filterConfig, filterConfig.externalReferences.front(), alreadyRegistered, duplicateError) && alreadyRegistered;
	SkillConfig::ExternalReference conflictingReference = filterConfig.externalReferences.front();
	conflictingReference.scope = AISkillScope::Repo;
	conflictingReference.projectDirectory = tempRoot / "project-a";
	const bool duplicateScopeRejected = !AddExternalReference(
		filterConfig, conflictingReference, alreadyRegistered, duplicateError);
	filterConfig.externalReferences.push_back(
		{tempRoot / "missing" / "SKILL.md", AISkillScope::User, {}, 400});
	std::vector<AISkillInfo> withMissing;
	DiscoverExternalReferences(filterConfig, tempRoot / "project-a", withMissing);
	const bool missingReferenceVisible = std::ranges::any_of(withMissing, [](const AISkillInfo& skill) {
		return skill.externalReference && !skill.valid && skill.directory.filename() == L"missing";
	});
	SkillConfig missingRemovableConfig = filterConfig;
	const bool missingReferenceRemovable = RemoveExternalReference(
		missingRemovableConfig, tempRoot / "missing" / "SKILL.md", AISkillScope::User, std::nullopt);
	SkillConfig removableConfig = filterConfig;
	const bool sourceExistedBeforeRemove = std::filesystem::exists(oneDirectory / "SKILL.md");
	const bool unregisterOk = RemoveExternalReference(
		removableConfig, oneDirectory / "SKILL.md", AISkillScope::User, std::nullopt) &&
		sourceExistedBeforeRemove && std::filesystem::exists(oneDirectory / "SKILL.md");

	PreparedLocalSkills directoryPrepared;
	const bool directoryDiscoveryOk = PrepareLocalSkills(PathToUtf8(sourceRoot), directoryPrepared, error) &&
		directoryPrepared.candidates.size() == 3;
	CleanupPreparedLocalSkills(directoryPrepared);
	PreparedLocalSkills directPrepared;
	const bool directFileDiscoveryOk = PrepareLocalSkills(
		PathToUtf8(oneDirectory / "SKILL.md"), directPrepared, error) &&
		directPrepared.candidates.size() == 1 && directPrepared.candidates.front().name == "one-skill";
	CleanupPreparedLocalSkills(directPrepared);

	const bool sourceMetadataWritten = WriteFileBytesAtomic(
		oneDirectory / kSourceMetadataFileName, R"({"repository":"fake/source"})", error);
	const auto copiedTarget = tempRoot / "installed" / "one-skill";
	const bool localCopyOk = sourceMetadataWritten &&
		ReplaceDirectoryAtomically(oneDirectory, copiedTarget, false, error, true) &&
		std::filesystem::exists(copiedTarget / "SKILL.md") &&
		!std::filesystem::exists(copiedTarget / kSourceMetadataFileName);
	const bool overlapDetectionOk = PathsOverlap(oneDirectory, oneDirectory / "child") &&
		!PathsOverlap(oneDirectory, tempRoot / "separate");

	const auto singleZipSource = tempRoot / "zip-single-source";
	const bool singleZipSkillWritten = writeSkill(singleZipSource / "only", "zip-one");
	const auto singleZip = tempRoot / "single.zip";
	const bool singleZipCreated = singleZipSkillWritten && CreateZipFromDirectory(singleZipSource, singleZip, error);
	PreparedLocalSkills singleZipPrepared;
	const bool singleZipOk = singleZipCreated && PrepareLocalSkills(PathToUtf8(singleZip), singleZipPrepared, error) &&
		singleZipPrepared.source.kind == AISkillLocalSourceKind::ZipArchive &&
		singleZipPrepared.candidates.size() == 1;
	const auto singleZipStaging = singleZipPrepared.source.stagingRoot;
	CleanupPreparedLocalSkills(singleZipPrepared);
	const bool zipCleanupOk = singleZipStaging.empty() || !std::filesystem::exists(singleZipStaging);
	const auto multiZip = tempRoot / "multi.ZIP";
	const bool multiZipCreated = CreateZipFromDirectory(sourceRoot, multiZip, error);
	PreparedLocalSkills multiZipPrepared;
	const bool multiZipOk = multiZipCreated && PrepareLocalSkills(PathToUtf8(multiZip), multiZipPrepared, error) &&
		multiZipPrepared.candidates.size() == 3 &&
		std::ranges::any_of(multiZipPrepared.candidates, [](const AISkillInfo& skill) {
			return skill.sourcePath == "nested/two";
		});
	CleanupPreparedLocalSkills(multiZipPrepared);

	const auto invalidSignatureZip = tempRoot / "invalid.zip";
	const bool invalidSignatureWritten = WriteBinaryFile(invalidSignatureZip, "not-a-zip", error);
	AISkillPreparedLocalSource invalidPrepared;
	std::string invalidSignatureError;
	const bool invalidSignatureRejected = invalidSignatureWritten && !AISkillLocalPackage::Prepare(
		invalidSignatureZip, GetAutoLinkerCacheDirectoryPath(), invalidPrepared, invalidSignatureError) &&
		invalidSignatureError.find("PK") != std::string::npos;
	const auto compressedOversizeZip = tempRoot / "compressed-oversize.zip";
	bool compressedOversizeWritten = false;
	{
		std::ofstream output(compressedOversizeZip, std::ios::binary | std::ios::trunc);
		if (output.is_open()) {
			output.write("PK", 2);
			output.seekp(static_cast<std::streamoff>(100ULL * 1024 * 1024));
			output.put('\0');
			compressedOversizeWritten = output.good();
		}
	}
	AISkillPreparedLocalSource compressedOversizePrepared;
	std::string compressedOversizeError;
	const bool compressedOversizeRejected = compressedOversizeWritten && !AISkillLocalPackage::Prepare(
		compressedOversizeZip, GetAutoLinkerCacheDirectoryPath(), compressedOversizePrepared, compressedOversizeError) &&
		compressedOversizeError.find("100 MiB") != std::string::npos;
	const auto declaredOversizeZip = tempRoot / "declared-oversize.zip";
	const bool declaredOversizeWritten = WriteBinaryFile(declaredOversizeZip, BuildDeclaredOversizeZip(), error);
	AISkillPreparedLocalSource declaredOversizePrepared;
	std::string declaredOversizeError;
	const bool declaredOversizeRejected = declaredOversizeWritten && !AISkillLocalPackage::Prepare(
		declaredOversizeZip, GetAutoLinkerCacheDirectoryPath(), declaredOversizePrepared, declaredOversizeError) &&
		declaredOversizeError.find("512 MiB") != std::string::npos;
	const auto traversalZip = tempRoot / "traversal.zip";
	const bool traversalZipCreated = CreateTraversalZip(traversalZip, error);
	AISkillPreparedLocalSource traversalPrepared;
	std::string traversalError;
	const bool traversalRejected = traversalZipCreated && !AISkillLocalPackage::Prepare(
		traversalZip, GetAutoLinkerCacheDirectoryPath(), traversalPrepared, traversalError) &&
		traversalError.find("穿越") != std::string::npos &&
		traversalPrepared.stagingRoot.empty();

	json parsedDetails;
	const bool detailsParserOk = ParseSkillsShDetailDocument(
		R"(<script type="application/ld+json">{"@type":"SoftwareApplication","name":"demo","description":"Demo details"}</script>)",
		parsedDetails) && parsedDetails.value("description", std::string()) == "Demo details";
	std::filesystem::remove_all(tempRoot, ec);
	const bool configOk = configV1Ok && configV2Ok && projectFilterOk && idempotentReferenceOk &&
		duplicateScopeRejected && missingReferenceVisible && missingReferenceRemovable && unregisterOk;
	const bool localSourceOk = directoryDiscoveryOk && directFileDiscoveryOk && localCopyOk && overlapDetectionOk;
	const bool zipOk = singleZipOk && multiZipOk && invalidSignatureRejected && compressedOversizeRejected &&
		declaredOversizeRejected && traversalRejected && zipCleanupOk;
	return json({
		{"name", "ai-skill-manager-self-test"},
		{"ok", parseOk && containmentOk && configOk && localSourceOk && zipOk && detailsParserOk},
		{"frontmatter_parser", parseOk},
		{"path_containment", containmentOk},
		{"config_v1_compatibility", configV1Ok},
		{"config_v2_round_trip", configV2Ok},
		{"external_project_filter", projectFilterOk},
		{"external_duplicate_registration", idempotentReferenceOk && duplicateScopeRejected},
		{"external_missing_removable", missingReferenceVisible && missingReferenceRemovable},
		{"external_unregister_preserves_source", unregisterOk},
		{"local_directory_discovery", directoryDiscoveryOk},
		{"local_skill_file_discovery", directFileDiscoveryOk},
		{"local_copy_metadata_cleanup", localCopyOk},
		{"local_copy_overlap_guard", overlapDetectionOk},
		{"zip_single_and_multi", singleZipOk && multiZipOk},
		{"zip_signature_and_size_limits", invalidSignatureRejected && compressedOversizeRejected && declaredOversizeRejected},
		{"zip_invalid_signature", invalidSignatureRejected},
		{"zip_compressed_size_limit", compressedOversizeRejected},
		{"zip_declared_size_limit", declaredOversizeRejected},
		{"zip_declared_size_error", declaredOversizeError},
		{"zip_traversal_and_cleanup", traversalRejected && zipCleanupOk},
		{"zip_traversal_created", traversalZipCreated},
		{"zip_traversal_rejected", traversalRejected},
		{"zip_traversal_error", traversalError},
		{"zip_cleanup", zipCleanupOk},
		{"skills_sh_details_parser", detailsParserOk},
		{"skills_sh_endpoint", "https://skills.sh/api/search"},
		{"global_root", PathToUtf8(GetGlobalSkillsRoot())}
	}).dump();
}

} // namespace AISkillManager
