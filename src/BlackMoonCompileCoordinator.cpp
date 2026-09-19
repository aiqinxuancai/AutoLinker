#include "BlackMoonCompileCoordinator.h"
#include "StaticJson.h"
#include "Global.h"
#include "ECOMEx.h"
#include "Logger.h"
#include <algorithm>
#include <cctype>
#include <cstdint>
#include <format>
#include <fstream>
#include <thread>

namespace {
using nlohmann::json;
constexpr const char* kProfiles[] = {"汇编模式", "C/C++模式", "VC++模式"};
constexpr const char* kModes[] = {"asm", "cxx", "vcxx"};

HMODULE Plugin()
{
	const auto module = GetModuleHandleW(L"BlackMoon.fne");
	return module ? module : GetModuleHandleW(L"BlackMoon.dll");
}

std::filesystem::path ConfigPath()
{
	wchar_t buffer[32768] = {};
	const DWORD length = GetModuleFileNameW(nullptr, buffer, _countof(buffer));
	if (length && length < _countof(buffer)) {
		const auto dir = std::filesystem::path(buffer).parent_path() / L"BlackMoon";
		if (GetPrivateProfileIntW(L"BlackMoon", L"ConfigOnInstallPath", 0,
			(dir / L"BlackMoon.ini").c_str()) == 1) return dir / L"Config.json";
	}
	const DWORD count = GetEnvironmentVariableW(L"ProgramData", buffer, _countof(buffer));
	return count && count < _countof(buffer)
		? std::filesystem::path(buffer) / L"BlackMoon" / L"Config.json" : std::filesystem::path();
}

std::string Convert(const std::string& text, UINT from, UINT to)
{
	const int count = MultiByteToWideChar(from, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (count <= 0) throw std::runtime_error("blackmoon_config_decode_failed");
	std::wstring wide(count, L'\0');
	MultiByteToWideChar(from, 0, text.data(), static_cast<int>(text.size()), wide.data(), count);
	BOOL substituted = FALSE;
	const auto flag = to == CP_UTF8 ? 0 : WC_NO_BEST_FIT_CHARS;
	BOOL* usedDefault = to == CP_UTF8 ? nullptr : &substituted;
	const int size = WideCharToMultiByte(to, flag, wide.data(), count, nullptr, 0, nullptr, usedDefault);
	if (size <= 0 || substituted) throw std::runtime_error("blackmoon_config_encode_failed");
	std::string result(size, '\0');
	WideCharToMultiByte(to, flag, wide.data(), count, result.data(), size, nullptr, usedDefault);
	return result;
}

std::string Read(const std::filesystem::path& path)
{
	std::ifstream input(path, std::ios::binary);
	if (!input) throw std::runtime_error("blackmoon_config_read_failed");
	std::string bytes((std::istreambuf_iterator<char>(input)), {});
	if (input.bad()) throw std::runtime_error("blackmoon_config_read_failed");
	return bytes;
}

// 同目录临时文件经刷盘后原子替换，失败时原文件保持完整。
bool WriteAtomic(const std::filesystem::path& path, const std::string& bytes)
{
	const auto temp = std::filesystem::path(path.wstring() + std::format(L".autolinker.{}.{}.tmp",
		GetCurrentProcessId(), GetTickCount64()));
	HANDLE file = CreateFileW(temp.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
	if (file == INVALID_HANDLE_VALUE) return false;
	DWORD written = 0;
	bool ok = WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) &&
		written == bytes.size() && FlushFileBuffers(file);
	CloseHandle(file);
	if (ok) ok = MoveFileExW(temp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) != FALSE;
	if (!ok) DeleteFileW(temp.c_str());
	return ok;
}

std::filesystem::path Journal(const std::filesystem::path& path)
{
	return std::filesystem::path(path.wstring() + L".autolinker-recovery.json");
}

bool Recover(const std::filesystem::path& path, std::string& error)
{
	if (!std::filesystem::exists(Journal(path))) return true;
	const auto journal = json::parse(Read(Journal(path)));
	const auto original = journal.at("original").get<std::vector<unsigned char>>();
	const auto temporary = journal.at("temporary").get<std::vector<unsigned char>>();
	const std::string before(original.begin(), original.end());
	const std::string during(temporary.begin(), temporary.end());
	const std::string current = Read(path);
	if (current != before && current != during) {
		error = "blackmoon_config_recovery_conflict: Config.json was changed externally; preserve and reconcile Config.json.autolinker-recovery.json before retrying";
		return false;
	}
	if (current != before && !WriteAtomic(path, before)) {
		error = "blackmoon_config_restore_failed";
		return false;
	}
	if (!DeleteFileW(Journal(path).c_str())) {
		error = "blackmoon_config_recovery_cleanup_failed";
		return false;
	}
	return true;
}

bool HasMode(const json& config, int mode)
{
	if (!config.is_object() || !config.contains("Config") || !config["Config"].is_object()) return false;
	const auto profile = config["Config"].find(kProfiles[mode]);
	if (profile == config["Config"].end() || !profile->is_object()) return false;
	const auto normal = profile->find("Normal");
	if (normal == profile->end() || !normal->is_object() || !normal->contains("BMType")) return false;
	const auto& type = (*normal)["BMType"];
	return type == mode || type == std::to_string(mode);
}

std::string Lower(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
	return text;
}
}

bool BlackMoonCompileCoordinator::ValidateArguments(json& args, json& error)
{
	const auto fail = [&](const char* code, const char* hint) {
		error = {{"ok", false}, {"error", code}, {"hint", hint}};
		return false;
	};
	if (!args.is_object()) return fail("invalid_compile_arguments", "Pass an object containing output_path, compiler and optional backend-specific parameters.");
	for (const char* field : {"compiler", "blackmoon_mode", "compile_mode", "target", "output_path"}) {
		if (args.contains(field) && !args[field].is_string())
			return fail("invalid_compile_parameter_type", "compiler, blackmoon_mode, target and output_path must be strings; omit unused optional fields.");
	}
	if (args.contains("static_compile") && !args["static_compile"].is_boolean())
		return fail("invalid_static_compile_type", "static_compile must be true or false for compiler=eide; omit it for compiler=blackmoon.");
	const auto compiler = Lower(args.value("compiler", std::string("eide")));
	if (compiler != "eide" && compiler != "blackmoon")
		return fail("unsupported_compiler", "Set compiler to eide or blackmoon. Omitting compiler selects eide.");
	const auto mode = Lower(args.value("blackmoon_mode", args.value("compile_mode", std::string())));
	if (args.contains("compile_mode") && args.contains("blackmoon_mode") &&
		Lower(args["compile_mode"].get<std::string>()) != mode)
		return fail("conflicting_blackmoon_modes", "Remove legacy compile_mode and use only blackmoon_mode=asm, cxx or vcxx.");
	args["compiler"] = compiler;
	if (compiler == "eide") {
		if (!mode.empty()) return fail("blackmoon_mode_requires_blackmoon", "For ordinary compilation remove blackmoon_mode/compile_mode. To use BlackMoon explicitly set compiler=blackmoon and omit static_compile.");
		return true;
	}
	if (args.contains("static_compile"))
		return fail("blackmoon_static_compile_not_supported", "Remove static_compile entirely (including false). BlackMoon controls its compilation pipeline; use compiler=eide if you need to choose static or dynamic compilation.");
	int index = -1;
	if (mode == "asm" || mode == "assembly" || mode == "0" || mode == "汇编模式") index = 0;
	if (mode == "cxx" || mode == "c" || mode == "c/c++" || mode == "1" || mode == "c/c++模式") index = 1;
	if (mode.empty() || mode == "vcxx" || mode == "vc" || mode == "vc++" || mode == "2" || mode == "vc++模式") index = 2;
	if (index < 0) return fail("unsupported_blackmoon_mode", "Use blackmoon_mode=asm, cxx or vcxx; omit it to select vcxx.");
	args["blackmoon_mode"] = kModes[index];
	return true;
}

bool BlackMoonCompileCoordinator::Begin(const std::filesystem::path& path, int mode, DWORD timeout, std::string& error)
{
	if (mutex_ || mode < 0 || mode > 2 || path.empty()) { error = "blackmoon_invalid_session"; return false; }
	path_ = std::filesystem::absolute(path).lexically_normal();
	auto key = path_.wstring();
	CharLowerBuffW(key.data(), static_cast<DWORD>(key.size()));
	std::uint64_t hash = 14695981039346656037ull;
	for (wchar_t ch : key) { hash ^= static_cast<std::uint16_t>(ch); hash *= 1099511628211ull; }
	const auto name = std::format(L"Global\\AutoLinker.BlackMoon.Config.{:016X}", hash);
	mutex_ = CreateMutexW(nullptr, FALSE, name.c_str());
	if (!mutex_) { error = "blackmoon_config_lock_failed"; return false; }
	const DWORD wait = WaitForSingleObject(mutex_, timeout);
	if (wait != WAIT_OBJECT_0 && wait != WAIT_ABANDONED) {
		error = wait == WAIT_TIMEOUT ? "blackmoon_config_lock_timeout" : "blackmoon_config_lock_failed";
		return false;
	}
	locked_ = true;
	if (!Recover(path_, error)) return false;
	original_ = Read(path_);
	auto config = json::parse(Convert(original_, CP_ACP, CP_UTF8));
	if (!HasMode(config, mode)) {
		error = std::string("blackmoon_mode_unavailable: restore profile ") + kModes[mode] + " with matching Normal.BMType in BlackMoon Config.json";
		return false;
	}
	config["Use"] = kProfiles[mode];
	temporary_ = Convert(config.dump(4), CP_UTF8, CP_ACP);
	const json journal = {
		{"original", std::vector<unsigned char>(original_.begin(), original_.end())},
		{"temporary", std::vector<unsigned char>(temporary_.begin(), temporary_.end())}
	};
	if (!WriteAtomic(Journal(path_), journal.dump())) { error = "blackmoon_config_journal_write_failed"; return false; }
	changed_ = true;
	if (!WriteAtomic(path_, temporary_)) { error = "blackmoon_config_write_failed"; return false; }
	return true;
}

bool BlackMoonCompileCoordinator::Finish(std::string& error)
{
	bool ok = true;
	try {
		if (changed_) ok = Recover(path_, error);
	} catch (const std::exception& ex) { error = std::string("blackmoon_config_restore_failed: ") + ex.what(); ok = false; }
	changed_ = false; // 恢复失败保留恢复日志，由下一次会话检查，不能盲目覆盖外部更改。
	if (locked_) ReleaseMutex(mutex_);
	if (mutex_) CloseHandle(mutex_);
	locked_ = false;
	mutex_ = nullptr;
	return ok;
}

BlackMoonCompileCoordinator::~BlackMoonCompileCoordinator()
{
	std::string error;
	if (!Finish(error)) Logger::Instance().Write("BlackMoon", error);
}

bool BlackMoonCompileCoordinator::Invoke(IDEFacade::CompileOutputKind kind, const std::string& outputPath,
	IDEFacade::BlackMoonCompileMode mode, std::string* normalizedPath, std::string* diagnostics)
{
	std::string error;
	const auto fail = [&]() { if (diagnostics) *diagnostics = error; return false; };
	if (kind == IDEFacade::CompileOutputKind::Ecom) { error = "blackmoon_ecom_not_supported: use compiler=eide to publish .ec"; return fail(); }
	HMODULE module = Plugin();
	if (!module) { error = "blackmoon_plugin_not_loaded: enable BlackMoon in this IDE and restart it"; return fail(); }
	using CompileProc = int(WINAPI*)(LPCSTR, LPCSTR);
	const auto compile = reinterpret_cast<CompileProc>(GetProcAddress(module, "BMCompile"));
	if (!compile || !GetProcAddress(module, "GetBMVersion")) { error = "blackmoon_plugin_api_unavailable"; return fail(); }
	std::string path;
	if (!IDEFacade::Instance().NormalizeCompileOutputPath(kind, outputPath, path, &error)) return fail();
	if (normalizedPath) *normalizedPath = path;
	try {
		if (!Begin(ConfigPath(), static_cast<int>(mode), 30000, error)) return fail();
	} catch (const std::exception& ex) { error = std::string("blackmoon_config_failed: ") + ex.what(); return fail(); }
	RunChangeECOM(true);
	if (!BeginSilentCompileOutputPathRequest(path, GetCurrentThreadId(), &error)) return fail();
	if (!compile(nullptr, nullptr)) { error = "blackmoon_invoke_failed"; return fail(); }
	if (diagnostics) *diagnostics = WasSilentCompileOutputPathRequestConsumed()
		? "blackmoon_compile_invoked_dialog_suppressed" : "blackmoon_compile_invoked_dialog_pending";
	return true;
}

json BlackMoonCompileCoordinator::Inspect()
{
	HMODULE module = Plugin();
	json result = {{"blackmoon_plugin_loaded", module && GetProcAddress(module, "BMCompile") && GetProcAddress(module, "GetBMVersion")},
		{"blackmoon_config_available", false}, {"blackmoon_supported_compile_modes", json::array()},
		{"blackmoon_default_compile_mode", "vcxx"}};
	try {
		const auto config = json::parse(Convert(Read(ConfigPath()), CP_ACP, CP_UTF8));
		for (int i = 0; i < 3; ++i) if (HasMode(config, i)) result["blackmoon_supported_compile_modes"].push_back(kModes[i]);
		result["blackmoon_config_available"] = !result["blackmoon_supported_compile_modes"].empty();
	} catch (const std::exception& ex) { result["blackmoon_config_error"] = ex.what(); }
	return result;
}

json BlackMoonCompileCoordinator::SelfTest()
{
	json checks = json::object();
	json error;
	json args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"eide","blackmoon_mode":"asm"})json");
	checks["ordinary_mode_rejected_with_hint"] = !ValidateArguments(args, error) && error.value("error", "") == "blackmoon_mode_requires_blackmoon" && error.contains("hint");
	args = AUTOLINKER_STATIC_JSON(R"json({"blackmoon_mode":"vcxx"})json");
	checks["implicit_backend_rejected"] = !ValidateArguments(args, error);
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"blackmoon","static_compile":false})json");
	checks["false_static_rejected"] = !ValidateArguments(args, error);
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"blackmoon"})json");
	checks["default_vcxx"] = ValidateArguments(args, error) && args["blackmoon_mode"] == "vcxx";
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"typo"})json");
	checks["unknown_backend_rejected"] = !ValidateArguments(args, error);
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"blackmoon","blackmoon_mode":2})json");
	checks["wrong_type_rejected"] = !ValidateArguments(args, error);
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"eide","static_compile":false})json");
	checks["ordinary_dynamic_allowed"] = ValidateArguments(args, error);
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"blackmoon","blackmoon_mode":"invalid"})json");
	checks["unknown_mode_rejected"] = !ValidateArguments(args, error) && error.contains("hint");
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"blackmoon","blackmoon_mode":"VC++"})json");
	checks["alias_normalized"] = ValidateArguments(args, error) && args["blackmoon_mode"] == "vcxx";
	args = AUTOLINKER_STATIC_JSON(R"json({"compiler":"eide","static_compile":"false"})json");
	checks["string_boolean_rejected"] = !ValidateArguments(args, error);
	const auto dir = std::filesystem::temp_directory_path() / std::format("AutoLinker-BM-test-{}-{}", GetCurrentProcessId(), GetTickCount64());
	std::filesystem::create_directory(dir);
	const auto path = dir / "Config.json";
	try {
		json config = AUTOLINKER_STATIC_JSON(R"json({"Use":"VC++\u6a21\u5f0f","Config":{}})json");
		for (int i = 0; i < 3; ++i) config["Config"][kProfiles[i]] = {{"Normal", {{"BMType", std::to_string(i)}}}};
		const auto original = Convert(config.dump(4), CP_UTF8, CP_ACP);
		if (!WriteAtomic(path, original)) throw std::runtime_error("fixture_write_failed");
		for (int i = 0; i < 3; ++i) {
			BlackMoonCompileCoordinator session;
			std::string detail;
			bool ok = session.Begin(path, i, 1000, detail);
			ok = ok && json::parse(Convert(Read(path), CP_ACP, CP_UTF8))["Use"] == kProfiles[i];
			bool blocked = false;
			std::thread contender([&] { BlackMoonCompileCoordinator second; std::string err;
				blocked = !second.Begin(path, 0, 30, err) && err == "blackmoon_config_lock_timeout"; });
			contender.join();
			ok = session.Finish(detail) && ok && blocked && Read(path) == original;
			checks[std::string("switch_restore_lock_") + kModes[i]] = ok;
		}
		{ BlackMoonCompileCoordinator session; std::string detail;
			checks["prepare_recovery"] = session.Begin(path, 0, 1000, detail);
			session.changed_ = false; // 模拟写入后退出，仅释放锁，保留持久恢复记录。
		}
		{ BlackMoonCompileCoordinator session; std::string detail;
			checks["crash_recovery"] = session.Begin(path, 1, 1000, detail) && session.Finish(detail) && Read(path) == original;
		}
		config["Config"][kProfiles[0]]["Normal"]["BMType"] = "2";
		const auto invalid = Convert(config.dump(4), CP_UTF8, CP_ACP);
		if (!WriteAtomic(path, invalid)) throw std::runtime_error("fixture_write_failed");
		{ BlackMoonCompileCoordinator session; std::string detail;
			checks["mismatched_profile_rejected_without_write"] = !session.Begin(path, 0, 1000, detail) &&
				Read(path) == invalid && !std::filesystem::exists(Journal(path));
		}
		if (!WriteAtomic(path, original)) throw std::runtime_error("fixture_write_failed");
		{ BlackMoonCompileCoordinator session; std::string detail;
			bool ok = session.Begin(path, 0, 1000, detail);
			ok = WriteAtomic(path, original + "\r\n") && ok;
			checks["external_edit_preserved"] = ok && !session.Finish(detail) && Read(path) == original + "\r\n";
		}
	} catch (const std::exception& ex) { checks["exception"] = ex.what(); }
	std::error_code ec;
	std::filesystem::remove_all(dir, ec); // 仅本次唯一测试目录。
	bool ok = true;
	for (const auto& check : checks) ok = ok && check == true;
	return {{"name", "blackmoon-compile"}, {"ok", ok}, {"checks", checks}};
}
