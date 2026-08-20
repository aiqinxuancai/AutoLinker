#include "ProjectBuildConfigManager.h"

#include <Windows.h>

#include <algorithm>
#include <fstream>
#include <iterator>
#include <format>
#include <system_error>

#include "..\\thirdparty\\json.hpp"
#include "PathHelper.h"

namespace {

using json = nlohmann::json;

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) return {};
	const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) return {};
	std::string result(static_cast<size_t>(size), '\0');
	WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr);
	return result;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) return {};
	const int size = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) return {};
	std::wstring result(static_cast<size_t>(size), L'\0');
	MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size);
	return result;
}

json ToJson(const ProjectBuildConfig& config)
{
	json pre = json::array();
	for (const auto& command : config.preBuildCommands) pre.push_back(command);
	json post = json::array();
	for (const auto& command : config.postBuildCommands) post.push_back(command);
	return {
		{"name", config.name},
		{"target", ProjectBuildTargetToString(config.target)},
		{"static_compile", config.staticCompile},
		{"output_path", config.outputPath},
		{"pre_build", pre},
		{"post_build", post}
	};
}

ProjectBuildConfig FromJson(const json& value)
{
	ProjectBuildConfig config;
	config.name = value.value("name", std::string());
	config.target = ProjectBuildTargetFromString(value.value("target", std::string("auto")));
	config.staticCompile = value.value("static_compile", false);
	config.outputPath = value.value("output_path", std::string());
	if (value.contains("pre_build") && value["pre_build"].is_array()) {
		for (const auto& item : value["pre_build"]) if (item.is_string()) config.preBuildCommands.push_back(item.get<std::string>());
	}
	if (value.contains("post_build") && value["post_build"].is_array()) {
		for (const auto& item : value["post_build"]) if (item.is_string()) config.postBuildCommands.push_back(item.get<std::string>());
	}
	return config;
}

std::string PathUtf8(const std::filesystem::path& path)
{
	return WideToUtf8(path.wstring());
}

std::string ReplaceVariables(const std::string& text, const ProjectBuildVariableContext& context, std::vector<std::string>* unknown)
{
	std::string result;
	result.reserve(text.size() + 32);
	for (size_t i = 0; i < text.size();) {
		if (text[i] != '{') {
			result.push_back(text[i++]);
			continue;
		}
		const size_t end = text.find('}', i + 1);
		if (end == std::string::npos) {
			result.append(text, i, std::string::npos);
			break;
		}
		const std::string name = text.substr(i + 1, end - i - 1);
		std::string replacement;
		if (name == "ProgramDir") replacement = PathUtf8(context.programDir);
		else if (name == "ProjectDir") replacement = PathUtf8(context.projectDir);
		else if (name == "ProjectName") replacement = context.projectName;
		else if (name == "SourcePath") replacement = PathUtf8(context.sourcePath);
		else if (name == "OutputPath") replacement = PathUtf8(context.outputPath);
		else {
			if (unknown != nullptr && std::find(unknown->begin(), unknown->end(), name) == unknown->end()) unknown->push_back(name);
			result.append(text, i, end - i + 1);
			i = end + 1;
			continue;
		}
		result += replacement;
		i = end + 1;
	}
	return result;
}

} // namespace

std::filesystem::path GetProjectBuildConfigPath(const std::filesystem::path& sourcePath)
{
	return sourcePath.parent_path() / (sourcePath.stem().wstring() + L".autolinker.json");
}

const char* ProjectBuildTargetToString(ProjectBuildTarget target)
{
	switch (target) {
	case ProjectBuildTarget::WinExe: return "win_exe";
	case ProjectBuildTarget::WinConsoleExe: return "win_console_exe";
	case ProjectBuildTarget::WinDll: return "win_dll";
	case ProjectBuildTarget::Ecom: return "ecom";
	default: return "auto";
	}
}

ProjectBuildTarget ProjectBuildTargetFromString(const std::string& value)
{
	if (value == "win_exe") return ProjectBuildTarget::WinExe;
	if (value == "win_console_exe") return ProjectBuildTarget::WinConsoleExe;
	if (value == "win_dll") return ProjectBuildTarget::WinDll;
	if (value == "ecom") return ProjectBuildTarget::Ecom;
	return ProjectBuildTarget::Auto;
}

ProjectBuildConfigFile LoadProjectBuildConfigFile(const std::filesystem::path& sourcePath, std::string* outError)
{
	if (outError != nullptr) outError->clear();
	const auto path = GetProjectBuildConfigPath(sourcePath);
	std::error_code existsError;
	const bool exists = std::filesystem::exists(path, existsError);
	if (existsError) {
		if (outError != nullptr) *outError = "check config file failed: " + existsError.message();
		return {};
	}
	if (!exists) return {};
	std::ifstream input(path, std::ios::binary);
	if (!input.is_open()) {
		if (outError != nullptr) *outError = "open config file failed";
		return {};
	}
	std::string text((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
	if (text.size() >= 3 && static_cast<unsigned char>(text[0]) == 0xEF && static_cast<unsigned char>(text[1]) == 0xBB && static_cast<unsigned char>(text[2]) == 0xBF) text.erase(0, 3);
	try {
		const json root = json::parse(text);
		if (!root.contains("configurations") || !root["configurations"].is_array()) {
			throw std::runtime_error("configurations is missing or not an array");
		}
		ProjectBuildConfigFile file;
		file.version = root.value("version", 1);
		file.activeConfiguration = root.value("active_configuration", std::string());
		for (const auto& item : root["configurations"]) {
			if (item.is_object()) file.configurations.push_back(FromJson(item));
		}
		if (file.configurations.empty()) {
			file.activeConfiguration.clear();
		}
		else if (file.activeConfiguration.empty()) {
			file.activeConfiguration = file.configurations.front().name;
		}
		return file;
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) *outError = ex.what();
		return {};
	}
}

bool SaveProjectBuildConfigFile(const std::filesystem::path& sourcePath, const ProjectBuildConfigFile& file, std::string* outError)
{
	if (outError != nullptr) outError->clear();
	try {
		const auto path = GetProjectBuildConfigPath(sourcePath);
		if (file.configurations.empty()) {
			std::error_code removeError;
			std::filesystem::remove(path, removeError);
			if (removeError) throw std::runtime_error("remove empty config file failed: " + removeError.message());
			return true;
		}
		json root = { {"version", file.version}, {"active_configuration", file.activeConfiguration}, {"configurations", json::array()} };
		for (const auto& config : file.configurations) root["configurations"].push_back(ToJson(config));
		std::filesystem::create_directories(path.parent_path());
		const auto temporary = path.wstring() + L".tmp";
		std::ofstream output(temporary, std::ios::binary | std::ios::trunc);
		if (!output.is_open()) throw std::runtime_error("open config temp file failed");
		const std::string serialized = root.dump(2, ' ', false, json::error_handler_t::replace);
		output.write("\xEF\xBB\xBF", 3);
		output.write(serialized.data(), static_cast<std::streamsize>(serialized.size()));
		output.close();
		if (!MoveFileExW(temporary.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) throw std::runtime_error("replace config file failed");
		return true;
	}
	catch (const std::exception& ex) {
		if (outError != nullptr) *outError = ex.what();
		return false;
	}
}

std::string ExpandProjectBuildVariables(const std::string& text, const ProjectBuildVariableContext& context, std::vector<std::string>* outUnknownVariables)
{
	if (outUnknownVariables != nullptr) outUnknownVariables->clear();
	return ReplaceVariables(text, context, outUnknownVariables);
}

std::string BuildProjectBuildConfigSelfTestJson()
{
	const auto root = std::filesystem::temp_directory_path() /
		std::format(L"AutoLinkerProjectBuildSelfTest-{}", GetCurrentProcessId());
	const auto sourcePath = root / L"中文工程.e";
	std::error_code cleanupError;
	std::filesystem::remove_all(root, cleanupError);
	std::filesystem::create_directories(root, cleanupError);

	const ProjectBuildConfigFile missing = LoadProjectBuildConfigFile(sourcePath);
	ProjectBuildConfigFile file;
	ProjectBuildConfig debug;
	debug.name = "Debug";
	debug.outputPath = "{ProjectDir}\\build\\Debug\\{ProjectName}";
	debug.preBuildCommands = { "Write-Output '{ProjectName}'" };
	ProjectBuildConfig release = debug;
	release.name = "Release";
	release.staticCompile = true;
	file.activeConfiguration = debug.name;
	file.configurations = { debug, release };
	std::string saveError;
	const bool saveOk = SaveProjectBuildConfigFile(sourcePath, file, &saveError);
	std::string loadError;
	const ProjectBuildConfigFile loaded = LoadProjectBuildConfigFile(sourcePath, &loadError);

	ProjectBuildVariableContext context;
	context.programDir = L"C:\\eide";
	context.projectDir = root;
	context.projectName = WideToUtf8(sourcePath.stem().wstring());
	context.sourcePath = sourcePath;
	context.outputPath = root / L"build\\中文工程.exe";
	std::vector<std::string> unknown;
	const std::string expanded = ExpandProjectBuildVariables(
		"{ProjectDir}\\{ProjectName}|{Unknown}|{OutputPath}", context, &unknown);

	const bool missingOk = missing.configurations.empty() && missing.activeConfiguration.empty();
	const bool sampleOk = file.configurations.size() == 2 &&
		file.configurations[0].name == "Debug" &&
		file.configurations[1].name == "Release" &&
		!file.configurations[0].staticCompile &&
		file.configurations[1].staticCompile;
	const bool roundTripOk = saveOk && loadError.empty() && loaded.configurations.size() == 2 &&
		loaded.configurations[0].preBuildCommands == file.configurations[0].preBuildCommands;
	std::string clearError;
	const bool clearSaved = SaveProjectBuildConfigFile(sourcePath, ProjectBuildConfigFile{}, &clearError);
	std::string clearedLoadError;
	const ProjectBuildConfigFile cleared = LoadProjectBuildConfigFile(sourcePath, &clearedLoadError);
	const bool clearOk = clearSaved && clearError.empty() && clearedLoadError.empty() &&
		cleared.configurations.empty() && cleared.activeConfiguration.empty() &&
		!std::filesystem::exists(GetProjectBuildConfigPath(sourcePath));
	const bool variablesOk = expanded.find("{ProjectDir}") == std::string::npos &&
		expanded.find("{ProjectName}") == std::string::npos &&
		expanded.find("{OutputPath}") == std::string::npos &&
		expanded.find("{Unknown}") != std::string::npos && unknown == std::vector<std::string>{ "Unknown" };
	const bool ok = missingOk && sampleOk && roundTripOk && clearOk && variablesOk;

	nlohmann::json report = {
		{"name", "project-build-config-self-test"},
		{"ok", ok},
		{"missing_file_has_no_defaults", missingOk},
		{"sample_config", sampleOk},
		{"round_trip", roundTripOk},
		{"clear_last_config", clearOk},
		{"variables", variablesOk},
		{"save_error", saveError},
		{"load_error", loadError},
		{"clear_error", clearError},
		{"cleared_load_error", clearedLoadError}
	};
	std::filesystem::remove_all(root, cleanupError);
	return report.dump();
}

