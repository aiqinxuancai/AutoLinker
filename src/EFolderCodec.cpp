#include "EFolderCodec.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <unordered_map>
#include <unordered_set>

#include "..\thirdparty\json.hpp"

namespace e2txt {

namespace {

constexpr std::int32_t kBundleDirectoryFormatVersion = 2;

using json = nlohmann::json;

constexpr unsigned char kUtf8Bom[3] = { 0xEF, 0xBB, 0xBF };

std::string EncodeBase64(const std::vector<std::uint8_t>& bytes);
std::vector<std::uint8_t> DecodeBase64(const std::string& text);

std::string ConvertCodePage(const std::string& text, const UINT fromCodePage, const UINT toCodePage, const DWORD fromFlags)
{
	if (text.empty() || fromCodePage == toCodePage) {
		return text;
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

std::string LocalToUtf8Text(const std::string& text)
{
	return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocalText(const std::string& text)
{
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string NormalizeCrLf(const std::string& text)
{
	std::string out;
	out.reserve(text.size() + 16);
	for (size_t index = 0; index < text.size(); ++index) {
		const char ch = text[index];
		if (ch == '\r') {
			out += "\r\n";
			if (index + 1 < text.size() && text[index + 1] == '\n') {
				++index;
			}
			continue;
		}
		if (ch == '\n') {
			out += "\r\n";
			continue;
		}
		out.push_back(ch);
	}
	return out;
}

bool ReadFileBytes(const std::filesystem::path& path, std::vector<std::uint8_t>& outBytes)
{
	outBytes.clear();
	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		return false;
	}

	in.seekg(0, std::ios::end);
	const std::streamoff size = in.tellg();
	if (size < 0) {
		return false;
	}
	in.seekg(0, std::ios::beg);

	outBytes.resize(static_cast<size_t>(size));
	in.read(reinterpret_cast<char*>(outBytes.data()), size);
	return in.good() || static_cast<size_t>(in.gcount()) == outBytes.size();
}

bool WriteBinaryFile(const std::filesystem::path& path, const std::vector<std::uint8_t>& data)
{
	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);
	std::ofstream out(path, std::ios::binary);
	if (!out.is_open()) {
		return false;
	}
	if (!data.empty()) {
		out.write(reinterpret_cast<const char*>(data.data()), static_cast<std::streamsize>(data.size()));
	}
	return out.good();
}

bool WriteUtf8TextFileBom(const std::filesystem::path& path, const std::string& utf8Text)
{
	std::error_code ec;
	std::filesystem::create_directories(path.parent_path(), ec);
	std::ofstream out(path, std::ios::binary);
	if (!out.is_open()) {
		return false;
	}
	out.write(reinterpret_cast<const char*>(kUtf8Bom), sizeof(kUtf8Bom));
	if (!utf8Text.empty()) {
		out.write(utf8Text.data(), static_cast<std::streamsize>(utf8Text.size()));
	}
	return out.good();
}

bool WriteLocalTextFileUtf8Bom(const std::filesystem::path& path, const std::string& text)
{
	return WriteUtf8TextFileBom(path, LocalToUtf8Text(NormalizeCrLf(text)));
}

bool ReadUtf8TextFile(const std::filesystem::path& path, std::string& outText)
{
	std::vector<std::uint8_t> bytes;
	if (!ReadFileBytes(path, bytes)) {
		return false;
	}

	size_t offset = 0;
	if (bytes.size() >= 3 &&
		bytes[0] == kUtf8Bom[0] &&
		bytes[1] == kUtf8Bom[1] &&
		bytes[2] == kUtf8Bom[2]) {
		offset = 3;
	}
	outText.assign(reinterpret_cast<const char*>(bytes.data() + offset), bytes.size() - offset);
	return true;
}

bool ReadTextFileUtf8ToLocal(const std::filesystem::path& path, std::string& outText)
{
	std::string utf8Text;
	if (!ReadUtf8TextFile(path, utf8Text)) {
		return false;
	}
	outText = Utf8ToLocalText(utf8Text);
	return true;
}

std::string NormalizeSourceRelativePathForWrite(const e2txt::BundleSourceFile& file)
{
	std::filesystem::path path = file.relativePath.empty()
		? std::filesystem::path("src") / (file.logicalName + ".txt")
		: std::filesystem::path(file.relativePath);
	if (path.extension().empty() || path.extension() == std::filesystem::path(L".efile")) {
		path.replace_extension(L".txt");
	}
	return path.generic_string();
}

bool ReadJsonFile(const std::filesystem::path& path, json& outJson)
{
	std::string utf8Text;
	if (!ReadUtf8TextFile(path, utf8Text)) {
		return false;
	}

	try {
		outJson = json::parse(utf8Text);
		return true;
	}
	catch (...) {
		return false;
	}
}

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string TrimAsciiCopy(std::string text)
{
	size_t begin = 0;
	while (begin < text.size() && static_cast<unsigned char>(text[begin]) <= 0x20) {
		++begin;
	}

	size_t end = text.size();
	while (end > begin && static_cast<unsigned char>(text[end - 1]) <= 0x20) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string TrimWindowsFileName(std::string text)
{
	while (!text.empty() && (text.back() == ' ' || text.back() == '.')) {
		text.pop_back();
	}
	while (!text.empty() && (text.front() == ' ')) {
		text.erase(text.begin());
	}
	return text;
}

bool IsReservedWindowsName(const std::string& name)
{
	static const std::unordered_set<std::string> kReserved = {
		"con", "prn", "aux", "nul",
		"com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8", "com9",
		"lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
	};
	return kReserved.contains(ToLowerAscii(name));
}

std::string SanitizePathSegment(std::string segment)
{
	if (segment.empty() || segment == "." || segment == "..") {
		return "_";
	}

	for (char& ch : segment) {
		const unsigned char uch = static_cast<unsigned char>(ch);
		if (uch < 0x20 || ch == '<' || ch == '>' || ch == ':' || ch == '"' ||
			ch == '/' || ch == '\\' || ch == '|' || ch == '?' || ch == '*') {
			ch = '_';
		}
	}
	segment = TrimWindowsFileName(std::move(segment));
	if (segment.empty()) {
		segment = "_";
	}

	std::filesystem::path path(segment);
	const std::string stemLower = ToLowerAscii(path.stem().string());
	if (IsReservedWindowsName(stemLower)) {
		segment.insert(segment.begin(), '_');
	}
	return segment;
}

std::string SanitizeRelativePath(const std::string& rawRelativePath)
{
	std::filesystem::path rawPath(rawRelativePath);
	std::filesystem::path sanitized;
	for (const auto& part : rawPath) {
		const std::string segment = part.generic_string();
		if (segment.empty() || segment == "/" || segment == "\\") {
			continue;
		}
		sanitized /= SanitizePathSegment(segment);
	}
	return sanitized.generic_string();
}

std::string MakeUniqueRelativePath(
	const std::string& rawRelativePath,
	std::unordered_set<std::string>& usedRelativePaths)
{
	std::filesystem::path path = SanitizeRelativePath(rawRelativePath);
	if (path.empty()) {
		path = "src/_";
	}

	const std::string extension = path.extension().generic_string();
	const std::string stem = path.stem().generic_string();
	const std::filesystem::path parent = path.parent_path();

	std::filesystem::path candidate = path;
	int suffix = 2;
	while (true) {
		const std::string key = ToLowerAscii(candidate.generic_string());
		if (!usedRelativePaths.contains(key)) {
			usedRelativePaths.insert(key);
			return candidate.generic_string();
		}

		candidate = parent / (stem + "_" + std::to_string(suffix) + extension);
		++suffix;
	}
}

void ReserveRelativePath(
	const std::string& rawRelativePath,
	std::unordered_set<std::string>& usedRelativePaths)
{
	const std::string sanitized = SanitizeRelativePath(rawRelativePath);
	if (!sanitized.empty()) {
		usedRelativePaths.insert(ToLowerAscii(sanitized));
	}
}

json DependencyToJson(const Dependency& dependency)
{
	json item;
	item["kind"] = dependency.kind == DependencyKind::ELib ? "elib" : "ecom";
	item["name"] = LocalToUtf8Text(dependency.name);
	item["fileName"] = LocalToUtf8Text(dependency.fileName);
	item["guid"] = LocalToUtf8Text(dependency.guid);
	item["versionText"] = LocalToUtf8Text(dependency.versionText);
	item["path"] = LocalToUtf8Text(dependency.path);
	item["reExport"] = dependency.reExport;
	return item;
}

Dependency DependencyFromJson(const json& item)
{
	Dependency dependency;
	const std::string kind = item.value("kind", "elib");
	dependency.kind = kind == "ecom" ? DependencyKind::ECom : DependencyKind::ELib;
	dependency.name = Utf8ToLocalText(item.value("name", ""));
	dependency.fileName = Utf8ToLocalText(item.value("fileName", ""));
	dependency.guid = Utf8ToLocalText(item.value("guid", ""));
	dependency.versionText = Utf8ToLocalText(item.value("versionText", ""));
	dependency.path = Utf8ToLocalText(item.value("path", ""));
	dependency.reExport = item.value("reExport", false);
	return dependency;
}

json FileMetaToJson(const BundleSourceFile& file)
{
	return json{
		{ "key", LocalToUtf8Text(file.key) },
		{ "logicalName", LocalToUtf8Text(file.logicalName) },
		{ "relativePath", LocalToUtf8Text(file.relativePath) },
	};
}

json FormMetaToJson(const BundleFormFile& file)
{
	return json{
		{ "key", LocalToUtf8Text(file.key) },
		{ "logicalName", LocalToUtf8Text(file.logicalName) },
		{ "relativePath", LocalToUtf8Text(file.relativePath) },
	};
}

json FolderMetaToJson(const BundleFolder& folder)
{
	json childKeys = json::array();
	for (const auto& childKey : folder.childKeys) {
		childKeys.push_back(LocalToUtf8Text(childKey));
	}
	return json{
		{ "key", folder.key },
		{ "parentKey", folder.parentKey },
		{ "expand", folder.expand },
		{ "name", LocalToUtf8Text(folder.name) },
		{ "childKeys", std::move(childKeys) },
	};
}

json WindowBindingToJson(const WindowBinding& binding)
{
	return json{
		{ "formName", LocalToUtf8Text(binding.formName) },
		{ "className", LocalToUtf8Text(binding.className) },
	};
}

json ResourceMetaToJson(const BundleBinaryResource& resource)
{
	return json{
		{ "key", LocalToUtf8Text(resource.key) },
		{ "logicalName", LocalToUtf8Text(resource.logicalName) },
		{ "relativePath", LocalToUtf8Text(resource.relativePath) },
		{ "comment", LocalToUtf8Text(resource.comment) },
		{ "isPublic", resource.isPublic },
	};
}

json NativeMethodSnapshotToJson(const BundleNativeMethodSnapshot& snapshot)
{
	return json{
		{ "name", LocalToUtf8Text(snapshot.name) },
		{ "textDigest", snapshot.textDigest },
		{ "id", snapshot.id },
		{ "memoryAddress", snapshot.memoryAddress },
		{ "attr", snapshot.attr },
		{ "paramIds", snapshot.paramIds },
		{ "localIds", snapshot.localIds },
		{ "lineOffset", EncodeBase64(snapshot.lineOffset) },
		{ "blockOffset", EncodeBase64(snapshot.blockOffset) },
		{ "methodReference", EncodeBase64(snapshot.methodReference) },
		{ "variableReference", EncodeBase64(snapshot.variableReference) },
		{ "constantReference", EncodeBase64(snapshot.constantReference) },
		{ "expressionData", EncodeBase64(snapshot.expressionData) },
	};
}

json NativeProgramHeaderSnapshotToJson(const BundleNativeProgramHeaderSnapshot& snapshot)
{
	json supportLibraryInfo = json::array();
	for (const auto& item : snapshot.supportLibraryInfo) {
		supportLibraryInfo.push_back(LocalToUtf8Text(item));
	}
	return json{
		{ "versionFlag1", snapshot.versionFlag1 },
		{ "unk1", snapshot.unk1 },
		{ "unk2_1", EncodeBase64(snapshot.unk2_1) },
		{ "unk2_2", EncodeBase64(snapshot.unk2_2) },
		{ "unk2_3", EncodeBase64(snapshot.unk2_3) },
		{ "supportLibraryInfo", std::move(supportLibraryInfo) },
		{ "flag1", snapshot.flag1 },
		{ "flag2", snapshot.flag2 },
		{ "unk3Op", EncodeBase64(snapshot.unk3Op) },
		{ "icon", EncodeBase64(snapshot.icon) },
		{ "debugCommandLine", LocalToUtf8Text(snapshot.debugCommandLine) },
	};
}

json NativeGlobalSnapshotToJson(const BundleNativeGlobalSnapshot& snapshot)
{
	return json{
		{ "name", LocalToUtf8Text(snapshot.name) },
		{ "textDigest", snapshot.textDigest },
		{ "id", snapshot.id },
	};
}

json NativeDllSnapshotToJson(const BundleNativeDllSnapshot& snapshot)
{
	return json{
		{ "name", LocalToUtf8Text(snapshot.name) },
		{ "textDigest", snapshot.textDigest },
		{ "id", snapshot.id },
		{ "memoryAddress", snapshot.memoryAddress },
		{ "paramIds", snapshot.paramIds },
	};
}

json NativeStructSnapshotToJson(const BundleNativeStructSnapshot& snapshot)
{
	return json{
		{ "name", LocalToUtf8Text(snapshot.name) },
		{ "textDigest", snapshot.textDigest },
		{ "id", snapshot.id },
		{ "memoryAddress", snapshot.memoryAddress },
		{ "memberIds", snapshot.memberIds },
	};
}

json NativeConstantSnapshotToJson(const BundleNativeConstantSnapshot& snapshot)
{
	return json{
		{ "name", LocalToUtf8Text(snapshot.name) },
		{ "key", LocalToUtf8Text(snapshot.key) },
		{ "textDigest", snapshot.textDigest },
		{ "id", snapshot.id },
		{ "pageType", snapshot.pageType },
	};
}

json NativeSourceSnapshotToJson(const BundleNativeSourceFileSnapshot& snapshot)
{
	json methods = json::array();
	for (const auto& method : snapshot.methods) {
		methods.push_back(NativeMethodSnapshotToJson(method));
	}
	return json{
		{ "contentDigest", snapshot.contentDigest },
		{ "classShapeDigest", snapshot.classShapeDigest },
		{ "classId", snapshot.classId },
		{ "classMemoryAddress", snapshot.classMemoryAddress },
		{ "formId", snapshot.formId },
		{ "baseClass", snapshot.baseClass },
		{ "classVarIds", snapshot.classVarIds },
		{ "methods", std::move(methods) },
	};
}

BundleSourceFile SourceFileMetaFromJson(const json& item)
{
	BundleSourceFile file;
	file.key = Utf8ToLocalText(item.value("key", ""));
	file.logicalName = Utf8ToLocalText(item.value("logicalName", ""));
	file.relativePath = Utf8ToLocalText(item.value("relativePath", ""));
	return file;
}

BundleFormFile FormFileMetaFromJson(const json& item)
{
	BundleFormFile file;
	file.key = Utf8ToLocalText(item.value("key", ""));
	file.logicalName = Utf8ToLocalText(item.value("logicalName", ""));
	file.relativePath = Utf8ToLocalText(item.value("relativePath", ""));
	return file;
}

BundleFolder FolderMetaFromJson(const json& item)
{
	BundleFolder folder;
	folder.key = item.value("key", 0);
	folder.parentKey = item.value("parentKey", 0);
	folder.expand = item.value("expand", true);
	folder.name = Utf8ToLocalText(item.value("name", ""));
	if (const auto it = item.find("childKeys"); it != item.end() && it->is_array()) {
		for (const auto& child : *it) {
			if (child.is_string()) {
				folder.childKeys.push_back(Utf8ToLocalText(child.get<std::string>()));
			}
		}
	}
	return folder;
}

WindowBinding WindowBindingFromJson(const json& item)
{
	WindowBinding binding;
	binding.formName = Utf8ToLocalText(item.value("formName", ""));
	binding.className = Utf8ToLocalText(item.value("className", ""));
	return binding;
}

BundleBinaryResource ResourceMetaFromJson(const json& item, BundleResourceKind kind)
{
	BundleBinaryResource resource;
	resource.kind = kind;
	resource.key = Utf8ToLocalText(item.value("key", ""));
	resource.logicalName = Utf8ToLocalText(item.value("logicalName", ""));
	resource.relativePath = Utf8ToLocalText(item.value("relativePath", ""));
	resource.comment = Utf8ToLocalText(item.value("comment", ""));
	resource.isPublic = item.value("isPublic", false);
	return resource;
}

BundleNativeMethodSnapshot NativeMethodSnapshotFromJson(const json& item)
{
	BundleNativeMethodSnapshot snapshot;
	snapshot.name = Utf8ToLocalText(item.value("name", ""));
	snapshot.textDigest = item.value("textDigest", "");
	snapshot.id = item.value("id", 0);
	snapshot.memoryAddress = item.value("memoryAddress", 0);
	snapshot.attr = item.value("attr", 0);
	if (const auto it = item.find("paramIds"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			if (value.is_number_integer()) {
				snapshot.paramIds.push_back(value.get<std::int32_t>());
			}
		}
	}
	if (const auto it = item.find("localIds"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			if (value.is_number_integer()) {
				snapshot.localIds.push_back(value.get<std::int32_t>());
			}
		}
	}
	snapshot.lineOffset = DecodeBase64(item.value("lineOffset", ""));
	snapshot.blockOffset = DecodeBase64(item.value("blockOffset", ""));
	snapshot.methodReference = DecodeBase64(item.value("methodReference", ""));
	snapshot.variableReference = DecodeBase64(item.value("variableReference", ""));
	snapshot.constantReference = DecodeBase64(item.value("constantReference", ""));
	snapshot.expressionData = DecodeBase64(item.value("expressionData", ""));
	return snapshot;
}

BundleNativeProgramHeaderSnapshot NativeProgramHeaderSnapshotFromJson(const json& item)
{
	BundleNativeProgramHeaderSnapshot snapshot;
	snapshot.versionFlag1 = item.value("versionFlag1", 0);
	snapshot.unk1 = item.value("unk1", 0);
	snapshot.unk2_1 = DecodeBase64(item.value("unk2_1", ""));
	snapshot.unk2_2 = DecodeBase64(item.value("unk2_2", ""));
	snapshot.unk2_3 = DecodeBase64(item.value("unk2_3", ""));
	if (const auto it = item.find("supportLibraryInfo"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			if (value.is_string()) {
				snapshot.supportLibraryInfo.push_back(Utf8ToLocalText(value.get<std::string>()));
			}
		}
	}
	snapshot.flag1 = item.value("flag1", 0);
	snapshot.flag2 = item.value("flag2", 0);
	snapshot.unk3Op = DecodeBase64(item.value("unk3Op", ""));
	snapshot.icon = DecodeBase64(item.value("icon", ""));
	snapshot.debugCommandLine = Utf8ToLocalText(item.value("debugCommandLine", ""));
	return snapshot;
}

BundleNativeGlobalSnapshot NativeGlobalSnapshotFromJson(const json& item)
{
	BundleNativeGlobalSnapshot snapshot;
	snapshot.name = Utf8ToLocalText(item.value("name", ""));
	snapshot.textDigest = item.value("textDigest", "");
	snapshot.id = item.value("id", 0);
	return snapshot;
}

BundleNativeDllSnapshot NativeDllSnapshotFromJson(const json& item)
{
	BundleNativeDllSnapshot snapshot;
	snapshot.name = Utf8ToLocalText(item.value("name", ""));
	snapshot.textDigest = item.value("textDigest", "");
	snapshot.id = item.value("id", 0);
	snapshot.memoryAddress = item.value("memoryAddress", 0);
	if (const auto it = item.find("paramIds"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			if (value.is_number_integer()) {
				snapshot.paramIds.push_back(value.get<std::int32_t>());
			}
		}
	}
	return snapshot;
}

BundleNativeStructSnapshot NativeStructSnapshotFromJson(const json& item)
{
	BundleNativeStructSnapshot snapshot;
	snapshot.name = Utf8ToLocalText(item.value("name", ""));
	snapshot.textDigest = item.value("textDigest", "");
	snapshot.id = item.value("id", 0);
	snapshot.memoryAddress = item.value("memoryAddress", 0);
	if (const auto it = item.find("memberIds"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			if (value.is_number_integer()) {
				snapshot.memberIds.push_back(value.get<std::int32_t>());
			}
		}
	}
	return snapshot;
}

BundleNativeConstantSnapshot NativeConstantSnapshotFromJson(const json& item)
{
	BundleNativeConstantSnapshot snapshot;
	snapshot.name = Utf8ToLocalText(item.value("name", ""));
	snapshot.key = Utf8ToLocalText(item.value("key", ""));
	snapshot.textDigest = item.value("textDigest", "");
	snapshot.id = item.value("id", 0);
	snapshot.pageType = item.value("pageType", 0);
	return snapshot;
}

BundleNativeSourceFileSnapshot NativeSourceSnapshotFromJson(const json& item)
{
	BundleNativeSourceFileSnapshot snapshot;
	snapshot.contentDigest = item.value("contentDigest", "");
	snapshot.classShapeDigest = item.value("classShapeDigest", "");
	snapshot.classId = item.value("classId", 0);
	snapshot.classMemoryAddress = item.value("classMemoryAddress", 0);
	snapshot.formId = item.value("formId", 0);
	snapshot.baseClass = item.value("baseClass", 0);
	if (const auto it = item.find("classVarIds"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			if (value.is_number_integer()) {
				snapshot.classVarIds.push_back(value.get<std::int32_t>());
			}
		}
	}
	if (const auto it = item.find("methods"); it != item.end() && it->is_array()) {
		for (const auto& value : *it) {
			snapshot.methods.push_back(NativeMethodSnapshotFromJson(value));
		}
	}
	return snapshot;
}

std::string DumpJson(const json& value)
{
	return value.dump(2, ' ', false, json::error_handler_t::replace);
}

std::filesystem::path GetProjectMetadataDirPath(const std::filesystem::path& root)
{
	return root / "project";
}

std::filesystem::path GetModuleJsonPath(const std::filesystem::path& root)
{
	return GetProjectMetadataDirPath(root) / "模块.json";
}

std::filesystem::path GetMetaJsonPath(const std::filesystem::path& root)
{
	return GetProjectMetadataDirPath(root) / "_meta.json";
}

std::filesystem::path GetNativeSourceSnapshotPath(const std::filesystem::path& root)
{
	return GetProjectMetadataDirPath(root) / ".native_source.bin";
}

std::filesystem::path GetNativeSourceMapPath(const std::filesystem::path& root)
{
	return GetProjectMetadataDirPath(root) / ".native_source_map.json";
}

std::filesystem::path GetNativeSymbolMapPath(const std::filesystem::path& root)
{
	return GetProjectMetadataDirPath(root) / ".native_symbol_map.json";
}

std::filesystem::path GetLegacyModuleJsonPath(const std::filesystem::path& root)
{
	return root / "src" / "模块.json";
}

std::filesystem::path GetLegacyMetaJsonPath(const std::filesystem::path& root)
{
	return root / "src" / "_meta.json";
}

bool ReadProjectJsonWithLegacyFallback(
	const std::filesystem::path& root,
	const std::filesystem::path& primaryPath,
	const std::filesystem::path& legacyPath,
	json& outJson)
{
	if (ReadJsonFile(primaryPath, outJson)) {
		return true;
	}
	return ReadJsonFile(legacyPath, outJson);
}

std::string EncodeBase64(const std::vector<std::uint8_t>& bytes)
{
	static constexpr char kBase64Table[] =
		"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	std::string out;
	out.reserve(((bytes.size() + 2) / 3) * 4);
	size_t index = 0;
	while (index + 3 <= bytes.size()) {
		const std::uint32_t value =
			(static_cast<std::uint32_t>(bytes[index]) << 16) |
			(static_cast<std::uint32_t>(bytes[index + 1]) << 8) |
			static_cast<std::uint32_t>(bytes[index + 2]);
		out.push_back(kBase64Table[(value >> 18) & 0x3F]);
		out.push_back(kBase64Table[(value >> 12) & 0x3F]);
		out.push_back(kBase64Table[(value >> 6) & 0x3F]);
		out.push_back(kBase64Table[value & 0x3F]);
		index += 3;
	}
	if (index + 1 == bytes.size()) {
		const std::uint32_t value = static_cast<std::uint32_t>(bytes[index]) << 16;
		out.push_back(kBase64Table[(value >> 18) & 0x3F]);
		out.push_back(kBase64Table[(value >> 12) & 0x3F]);
		out.push_back('=');
		out.push_back('=');
	}
	else if (index + 2 == bytes.size()) {
		const std::uint32_t value =
			(static_cast<std::uint32_t>(bytes[index]) << 16) |
			(static_cast<std::uint32_t>(bytes[index + 1]) << 8);
		out.push_back(kBase64Table[(value >> 18) & 0x3F]);
		out.push_back(kBase64Table[(value >> 12) & 0x3F]);
		out.push_back(kBase64Table[(value >> 6) & 0x3F]);
		out.push_back('=');
	}
	return out;
}

std::vector<std::uint8_t> DecodeBase64(const std::string& text)
{
	static constexpr std::array<std::uint8_t, 256> kMap = [] {
		std::array<std::uint8_t, 256> map = {};
		map.fill(0xFF);
		for (char ch = 'A'; ch <= 'Z'; ++ch) {
			map[static_cast<unsigned char>(ch)] = static_cast<std::uint8_t>(ch - 'A');
		}
		for (char ch = 'a'; ch <= 'z'; ++ch) {
			map[static_cast<unsigned char>(ch)] = static_cast<std::uint8_t>(ch - 'a' + 26);
		}
		for (char ch = '0'; ch <= '9'; ++ch) {
			map[static_cast<unsigned char>(ch)] = static_cast<std::uint8_t>(ch - '0' + 52);
		}
		map[static_cast<unsigned char>('+')] = 62;
		map[static_cast<unsigned char>('/')] = 63;
		return map;
	}();

	std::vector<std::uint8_t> out;
	std::uint32_t value = 0;
	int bits = 0;
	for (const unsigned char ch : text) {
		if (std::isspace(ch) != 0) {
			continue;
		}
		if (ch == '=') {
			break;
		}
		const std::uint8_t mapped = kMap[ch];
		if (mapped == 0xFF) {
			return {};
		}
		value = (value << 6) | mapped;
		bits += 6;
		if (bits >= 8) {
			bits -= 8;
			out.push_back(static_cast<std::uint8_t>((value >> bits) & 0xFF));
		}
	}
	return out;
}

std::string NormalizeRelativePathKey(const std::filesystem::path& path)
{
	return ToLowerAscii(path.lexically_normal().generic_string());
}

std::string NormalizeRelativePathKey(const std::string& pathText)
{
	return NormalizeRelativePathKey(std::filesystem::path(pathText));
}

bool AppendUniqueKey(std::vector<std::string>& keys, const std::string& key)
{
	if (key.empty()) {
		return false;
	}
	if (std::find(keys.begin(), keys.end(), key) != keys.end()) {
		return false;
	}
	keys.push_back(key);
	return true;
}

std::string BuildFolderKey(const std::int32_t folderKey)
{
	return "folder:" + std::to_string(folderKey);
}

std::string BuildBundleItemKey(
	const std::string& prefix,
	const std::string& rawName,
	std::unordered_map<std::string, int>& counters)
{
	std::string logicalName = TrimAsciiCopy(rawName);
	if (logicalName.empty()) {
		logicalName = prefix;
	}

	const std::string baseKey = prefix + ":" + logicalName;
	int& counter = counters[baseKey];
	++counter;
	if (counter == 1) {
		return baseKey;
	}
	return baseKey + "#" + std::to_string(counter);
}

void ObserveBundleItemKeyBase(
	const std::string& prefix,
	const std::string& rawName,
	std::unordered_map<std::string, int>& counters)
{
	std::string logicalName = TrimAsciiCopy(rawName);
	if (logicalName.empty()) {
		logicalName = prefix;
	}
	++counters[prefix + ":" + logicalName];
}

bool IsReservedBundleSourceFile(const std::filesystem::path& relativePath)
{
	const std::string normalized = NormalizeRelativePathKey(relativePath);
	static const std::unordered_set<std::string> kReserved = {
		"src/.数据类型.txt",
		"src/.dll声明.txt",
		"src/.常量.txt",
		"src/.全局变量.txt",
	};
	return kReserved.contains(normalized);
}

bool IsAutoDiscoverableSourceTextPath(const std::filesystem::path& relativePath)
{
	const std::string normalized = NormalizeRelativePathKey(relativePath);
	if (!normalized.starts_with("src/")) {
		return false;
	}
	if (ToLowerAscii(relativePath.extension().generic_string()) != ".txt") {
		return false;
	}
	if (IsReservedBundleSourceFile(relativePath)) {
		return false;
	}

	const std::string fileName = relativePath.filename().generic_string();
	return !fileName.empty() && fileName.front() != '.';
}

BundleFolder* FindFolderByKey(std::vector<BundleFolder>& folders, const std::int32_t key)
{
	for (auto& folder : folders) {
		if (folder.key == key) {
			return &folder;
		}
	}
	return nullptr;
}

BundleFolder* FindFolderByParentAndName(
	std::vector<BundleFolder>& folders,
	const std::int32_t parentKey,
	const std::string& name)
{
	for (auto& folder : folders) {
		if (folder.parentKey == parentKey && folder.name == name) {
			return &folder;
		}
	}
	return nullptr;
}

bool ContainsAssignedTreeKey(
	const std::vector<BundleFolder>& folders,
	const std::vector<std::string>& rootChildKeys,
	const std::string& key)
{
	if (std::find(rootChildKeys.begin(), rootChildKeys.end(), key) != rootChildKeys.end()) {
		return true;
	}

	for (const auto& folder : folders) {
		if (std::find(folder.childKeys.begin(), folder.childKeys.end(), key) != folder.childKeys.end()) {
			return true;
		}
	}
	return false;
}

std::int32_t GetMaxExistingFolderKey(const std::vector<BundleFolder>& folders)
{
	std::int32_t maxKey = 0;
	for (const auto& folder : folders) {
		maxKey = (std::max)(maxKey, folder.key);
		maxKey = (std::max)(maxKey, folder.parentKey);
	}
	return maxKey;
}

std::int32_t EnsureFolderPath(
	const std::vector<std::string>& segments,
	std::vector<BundleFolder>& folders,
	std::vector<std::string>& rootChildKeys,
	std::int32_t& folderAllocatedKey)
{
	std::int32_t nextFolderKey = (std::max)(folderAllocatedKey, GetMaxExistingFolderKey(folders));
	std::int32_t parentKey = 0;
	for (const auto& rawSegment : segments) {
		const std::string segment = TrimAsciiCopy(rawSegment);
		if (segment.empty()) {
			continue;
		}

		BundleFolder* existing = FindFolderByParentAndName(folders, parentKey, segment);
		if (existing != nullptr) {
			parentKey = existing->key;
			continue;
		}

		BundleFolder created;
		created.key = nextFolderKey + 1;
		created.parentKey = parentKey;
		created.expand = true;
		created.name = segment;
		nextFolderKey = created.key;
		folderAllocatedKey = created.key;
		folders.push_back(created);

		if (parentKey == 0) {
			AppendUniqueKey(rootChildKeys, BuildFolderKey(created.key));
		}
		else if (BundleFolder* parent = FindFolderByKey(folders, parentKey); parent != nullptr) {
			AppendUniqueKey(parent->childKeys, BuildFolderKey(created.key));
		}

		parentKey = created.key;
	}

	return parentKey;
}

std::vector<std::string> GetFolderSegmentsFromRelativePath(const std::string& relativePath)
{
	std::vector<std::string> segments;
	std::filesystem::path path(relativePath);
	bool skippedSrc = false;
	for (const auto& part : path.parent_path()) {
		const std::string segment = part.generic_string();
		if (segment.empty() || segment == "/" || segment == "\\") {
			continue;
		}
		if (!skippedSrc) {
			skippedSrc = true;
			if (ToLowerAscii(segment) == "src") {
				continue;
			}
		}
		segments.push_back(segment);
	}
	return segments;
}

void AttachSourceFileToTree(
	const BundleSourceFile& file,
	std::vector<BundleFolder>& folders,
	std::vector<std::string>& rootChildKeys,
	std::int32_t& folderAllocatedKey)
{
	if (file.key.empty() || ContainsAssignedTreeKey(folders, rootChildKeys, file.key)) {
		return;
	}

	const std::vector<std::string> folderSegments = GetFolderSegmentsFromRelativePath(
		file.relativePath.empty() ? NormalizeSourceRelativePathForWrite(file) : file.relativePath);
	if (folderSegments.empty()) {
		AppendUniqueKey(rootChildKeys, file.key);
		return;
	}

	const std::int32_t folderKey = EnsureFolderPath(folderSegments, folders, rootChildKeys, folderAllocatedKey);
	if (folderKey == 0) {
		AppendUniqueKey(rootChildKeys, file.key);
		return;
	}

	if (BundleFolder* folder = FindFolderByKey(folders, folderKey); folder != nullptr) {
		AppendUniqueKey(folder->childKeys, file.key);
	}
	else {
		AppendUniqueKey(rootChildKeys, file.key);
	}
}

bool ReadAutoDiscoveredSourceFiles(
	const std::filesystem::path& root,
	ProjectBundle& bundle,
	std::string* outError)
{
	const std::filesystem::path sourceRoot = root / "src";
	std::error_code ec;
	if (!std::filesystem::exists(sourceRoot, ec) || ec) {
		return !ec;
	}

	std::unordered_set<std::string> knownRelativePaths;
	std::unordered_map<std::string, int> sourceKeyCounters;
	for (const auto& file : bundle.sourceFiles) {
		knownRelativePaths.insert(NormalizeRelativePathKey(
			file.relativePath.empty() ? NormalizeSourceRelativePathForWrite(file) : file.relativePath));
		ObserveBundleItemKeyBase("class", file.logicalName, sourceKeyCounters);
	}

	for (const auto& file : bundle.sourceFiles) {
		AttachSourceFileToTree(file, bundle.folders, bundle.rootChildKeys, bundle.folderAllocatedKey);
	}

	std::vector<std::filesystem::path> pendingRelativePaths;
	for (std::filesystem::recursive_directory_iterator it(sourceRoot, ec), end; it != end; it.increment(ec)) {
		if (ec) {
			if (outError != nullptr) {
				*outError = "enumerate_source_files_failed";
			}
			return false;
		}
		if (!it->is_regular_file()) {
			continue;
		}

		const std::filesystem::path relativePath = std::filesystem::relative(it->path(), root, ec);
		if (ec) {
			if (outError != nullptr) {
				*outError = "relative_source_path_failed";
			}
			return false;
		}
		if (!IsAutoDiscoverableSourceTextPath(relativePath)) {
			continue;
		}

		const std::string relativeKey = NormalizeRelativePathKey(relativePath);
		if (knownRelativePaths.contains(relativeKey)) {
			continue;
		}
		pendingRelativePaths.push_back(relativePath);
	}

	std::sort(
		pendingRelativePaths.begin(),
		pendingRelativePaths.end(),
		[](const std::filesystem::path& left, const std::filesystem::path& right) {
			return NormalizeRelativePathKey(left) < NormalizeRelativePathKey(right);
		});

	for (const auto& relativePath : pendingRelativePaths) {
		BundleSourceFile file;
		file.logicalName = relativePath.stem().string();
		file.relativePath = relativePath.generic_string();
		file.key = BuildBundleItemKey("class", file.logicalName, sourceKeyCounters);
		if (!ReadTextFileUtf8ToLocal(root / relativePath, file.content)) {
			if (outError != nullptr) {
				*outError = "read_auto_discovered_source_file_failed";
			}
			return false;
		}

		knownRelativePaths.insert(NormalizeRelativePathKey(relativePath));
		bundle.sourceFiles.push_back(std::move(file));
		AttachSourceFileToTree(bundle.sourceFiles.back(), bundle.folders, bundle.rootChildKeys, bundle.folderAllocatedKey);
	}

	return true;
}

}  // namespace

bool BundleDirectoryCodec::WriteBundle(const ProjectBundle& bundle, const std::string& outputDir, std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}

	const std::filesystem::path root = std::filesystem::path(outputDir);
	std::error_code ec;
	std::filesystem::create_directories(root, ec);
	if (ec) {
		if (outError != nullptr) {
			*outError = "create_output_dir_failed";
		}
		return false;
	}

	std::unordered_set<std::string> usedRelativePaths;
	ReserveRelativePath("project/模块.json", usedRelativePaths);
	ReserveRelativePath("project/_meta.json", usedRelativePaths);
	ReserveRelativePath("project/.native_source.bin", usedRelativePaths);
	ReserveRelativePath("project/.native_source_map.json", usedRelativePaths);
	ReserveRelativePath("project/.native_symbol_map.json", usedRelativePaths);
	ReserveRelativePath("src/.数据类型.txt", usedRelativePaths);
	ReserveRelativePath("src/.DLL声明.txt", usedRelativePaths);
	ReserveRelativePath("src/.常量.txt", usedRelativePaths);
	ReserveRelativePath("src/.全局变量.txt", usedRelativePaths);
	ReserveRelativePath("image/list.json", usedRelativePaths);
	ReserveRelativePath("audio/list.json", usedRelativePaths);

	json moduleJson;
	moduleJson["projectName"] = LocalToUtf8Text(bundle.projectName);
	moduleJson["versionText"] = LocalToUtf8Text(bundle.versionText);
	moduleJson["sourcePath"] = LocalToUtf8Text(bundle.sourcePath);
	moduleJson["dependencies"] = json::array();
	for (const auto& dependency : bundle.dependencies) {
		moduleJson["dependencies"].push_back(DependencyToJson(dependency));
	}
	if (!WriteUtf8TextFileBom(GetModuleJsonPath(root), NormalizeCrLf(DumpJson(moduleJson)))) {
		if (outError != nullptr) {
			*outError = "write_module_json_failed";
		}
		return false;
	}

	if (!WriteLocalTextFileUtf8Bom(root / "src" / ".数据类型.txt", bundle.dataTypeText) ||
		!WriteLocalTextFileUtf8Bom(root / "src" / ".DLL声明.txt", bundle.dllDeclareText) ||
		!WriteLocalTextFileUtf8Bom(root / "src" / ".常量.txt", bundle.constantText) ||
		!WriteLocalTextFileUtf8Bom(root / "src" / ".全局变量.txt", bundle.globalText)) {
		if (outError != nullptr) {
			*outError = "write_fixed_text_failed";
		}
		return false;
	}

	json metaJson;
	metaJson["formatVersion"] = kBundleDirectoryFormatVersion;
	metaJson["projectNameStored"] = bundle.projectNameStored;
	metaJson["sourceFiles"] = json::array();
	metaJson["formFiles"] = json::array();
	metaJson["folders"] = json::array();
	metaJson["folderAllocatedKey"] = bundle.folderAllocatedKey;
	metaJson["rootChildKeys"] = json::array();
	for (const auto& childKey : bundle.rootChildKeys) {
		metaJson["rootChildKeys"].push_back(LocalToUtf8Text(childKey));
	}
	metaJson["windowBindings"] = json::array();

	ProjectBundle persistedBundle;
	persistedBundle.sourcePath = bundle.sourcePath;
	persistedBundle.projectName = bundle.projectName;
	persistedBundle.projectNameStored = bundle.projectNameStored;
	persistedBundle.versionText = bundle.versionText;
	persistedBundle.bundleFormatVersion = kBundleDirectoryFormatVersion;
	persistedBundle.dependencies = bundle.dependencies;
	persistedBundle.dataTypeText = NormalizeCrLf(bundle.dataTypeText);
	persistedBundle.dllDeclareText = NormalizeCrLf(bundle.dllDeclareText);
	persistedBundle.constantText = NormalizeCrLf(bundle.constantText);
	persistedBundle.globalText = NormalizeCrLf(bundle.globalText);
	persistedBundle.folderAllocatedKey = bundle.folderAllocatedKey;
	persistedBundle.folders = bundle.folders;
	persistedBundle.rootChildKeys = bundle.rootChildKeys;
	persistedBundle.windowBindings = bundle.windowBindings;
	persistedBundle.nativeProgramHeader = bundle.nativeProgramHeader;
	persistedBundle.nativeGlobalSnapshots = bundle.nativeGlobalSnapshots;
	persistedBundle.nativeStructSnapshots = bundle.nativeStructSnapshots;
	persistedBundle.nativeDllSnapshots = bundle.nativeDllSnapshots;
	persistedBundle.nativeConstantSnapshots = bundle.nativeConstantSnapshots;

	for (auto file : bundle.sourceFiles) {
		const std::string desiredRelativePath = NormalizeSourceRelativePathForWrite(file);
		file.relativePath = MakeUniqueRelativePath(desiredRelativePath, usedRelativePaths);
		if (!WriteLocalTextFileUtf8Bom(root / file.relativePath, file.content)) {
			if (outError != nullptr) {
				*outError = "write_source_file_failed";
			}
			return false;
		}
		file.content = NormalizeCrLf(file.content);
		metaJson["sourceFiles"].push_back(FileMetaToJson(file));
		persistedBundle.sourceFiles.push_back(file);
	}

	for (auto file : bundle.formFiles) {
		const std::string desiredRelativePath = file.relativePath.empty()
			? std::string("src/") + file.logicalName + ".xml"
			: file.relativePath;
		file.relativePath = MakeUniqueRelativePath(desiredRelativePath, usedRelativePaths);
		if (!WriteLocalTextFileUtf8Bom(root / file.relativePath, file.xmlText)) {
			if (outError != nullptr) {
				*outError = "write_form_file_failed";
			}
			return false;
		}
		file.xmlText = NormalizeCrLf(file.xmlText);
		metaJson["formFiles"].push_back(FormMetaToJson(file));
		persistedBundle.formFiles.push_back(file);
	}

	for (const auto& folder : bundle.folders) {
		metaJson["folders"].push_back(FolderMetaToJson(folder));
	}
	for (const auto& binding : bundle.windowBindings) {
		metaJson["windowBindings"].push_back(WindowBindingToJson(binding));
	}

	if (!bundle.nativeSourceBytes.empty()) {
		if (!WriteBinaryFile(GetNativeSourceSnapshotPath(root), bundle.nativeSourceBytes)) {
			if (outError != nullptr) {
				*outError = "write_native_source_snapshot_failed";
			}
			return false;
		}
	}
	else {
		std::error_code removeEc;
		std::filesystem::remove(GetNativeSourceSnapshotPath(root), removeEc);
	}

	if (!bundle.nativeSourceSnapshots.empty()) {
		json nativeSourceMap = json::array();
		for (const auto& snapshot : bundle.nativeSourceSnapshots) {
			nativeSourceMap.push_back(NativeSourceSnapshotToJson(snapshot));
		}
		if (!WriteUtf8TextFileBom(GetNativeSourceMapPath(root), NormalizeCrLf(DumpJson(nativeSourceMap)))) {
			if (outError != nullptr) {
				*outError = "write_native_source_map_failed";
			}
			return false;
		}
	}
	else {
		std::error_code removeEc;
		std::filesystem::remove(GetNativeSourceMapPath(root), removeEc);
	}

	if (!bundle.nativeGlobalSnapshots.empty() ||
		!bundle.nativeStructSnapshots.empty() ||
		!bundle.nativeDllSnapshots.empty() ||
		!bundle.nativeConstantSnapshots.empty()) {
		json nativeSymbolMap;
		if (bundle.nativeProgramHeader.has_value()) {
			nativeSymbolMap["programHeader"] = NativeProgramHeaderSnapshotToJson(*bundle.nativeProgramHeader);
		}
		nativeSymbolMap["globals"] = json::array();
		nativeSymbolMap["structs"] = json::array();
		nativeSymbolMap["dlls"] = json::array();
		nativeSymbolMap["constants"] = json::array();
		for (const auto& snapshot : bundle.nativeGlobalSnapshots) {
			nativeSymbolMap["globals"].push_back(NativeGlobalSnapshotToJson(snapshot));
		}
		for (const auto& snapshot : bundle.nativeStructSnapshots) {
			nativeSymbolMap["structs"].push_back(NativeStructSnapshotToJson(snapshot));
		}
		for (const auto& snapshot : bundle.nativeDllSnapshots) {
			nativeSymbolMap["dlls"].push_back(NativeDllSnapshotToJson(snapshot));
		}
		for (const auto& snapshot : bundle.nativeConstantSnapshots) {
			nativeSymbolMap["constants"].push_back(NativeConstantSnapshotToJson(snapshot));
		}
		if (!WriteUtf8TextFileBom(GetNativeSymbolMapPath(root), NormalizeCrLf(DumpJson(nativeSymbolMap)))) {
			if (outError != nullptr) {
				*outError = "write_native_symbol_map_failed";
			}
			return false;
		}
	}
	else {
		std::error_code removeEc;
		std::filesystem::remove(GetNativeSymbolMapPath(root), removeEc);
	}

	json imageList;
	imageList["items"] = json::array();
	json audioList;
	audioList["items"] = json::array();
	for (auto resource : bundle.resources) {
		const std::string defaultRelativePath =
			std::string(resource.kind == BundleResourceKind::Image ? "image/" : "audio/") +
			resource.logicalName + ".bin";
		resource.relativePath = MakeUniqueRelativePath(
			resource.relativePath.empty() ? defaultRelativePath : resource.relativePath,
			usedRelativePaths);
		if (!WriteBinaryFile(root / resource.relativePath, resource.data)) {
			if (outError != nullptr) {
				*outError = "write_resource_file_failed";
			}
			return false;
		}

		if (resource.kind == BundleResourceKind::Image) {
			imageList["items"].push_back(ResourceMetaToJson(resource));
		}
		else {
			audioList["items"].push_back(ResourceMetaToJson(resource));
		}
		persistedBundle.resources.push_back(resource);
	}

	metaJson["nativeBundleDigest"] = LocalToUtf8Text(ComputeBundleDigest(persistedBundle));

	if (!WriteUtf8TextFileBom(GetMetaJsonPath(root), NormalizeCrLf(DumpJson(metaJson)))) {
		if (outError != nullptr) {
			*outError = "write_meta_json_failed";
		}
		return false;
	}

	if (!WriteUtf8TextFileBom(root / "image" / "list.json", NormalizeCrLf(DumpJson(imageList))) ||
		!WriteUtf8TextFileBom(root / "audio" / "list.json", NormalizeCrLf(DumpJson(audioList)))) {
		if (outError != nullptr) {
			*outError = "write_resource_list_failed";
		}
		return false;
	}

	return true;
}

bool BundleDirectoryCodec::ReadBundle(const std::string& inputDir, ProjectBundle& outBundle, std::string* outError) const
{
	if (outError != nullptr) {
		outError->clear();
	}

	const std::filesystem::path root = std::filesystem::path(inputDir);
	json moduleJson;
	if (!ReadProjectJsonWithLegacyFallback(
			root,
			GetModuleJsonPath(root),
			GetLegacyModuleJsonPath(root),
			moduleJson)) {
		if (outError != nullptr) {
			*outError = "read_module_json_failed";
		}
		return false;
	}

	json metaJson;
	if (!ReadProjectJsonWithLegacyFallback(
			root,
			GetMetaJsonPath(root),
			GetLegacyMetaJsonPath(root),
			metaJson)) {
		if (outError != nullptr) {
			*outError = "read_meta_json_failed";
		}
		return false;
	}

	ProjectBundle bundle;
	bundle.projectName = Utf8ToLocalText(moduleJson.value("projectName", ""));
	bundle.versionText = Utf8ToLocalText(moduleJson.value("versionText", ""));
	bundle.sourcePath = Utf8ToLocalText(moduleJson.value("sourcePath", ""));
	bundle.bundleFormatVersion = metaJson.value("formatVersion", 0);
	bundle.projectNameStored = metaJson.value("projectNameStored", true);
	if (const auto it = moduleJson.find("dependencies"); it != moduleJson.end() && it->is_array()) {
		for (const auto& dependencyItem : *it) {
			bundle.dependencies.push_back(DependencyFromJson(dependencyItem));
		}
	}

	ReadTextFileUtf8ToLocal(root / "src" / ".数据类型.txt", bundle.dataTypeText);
	ReadTextFileUtf8ToLocal(root / "src" / ".DLL声明.txt", bundle.dllDeclareText);
	ReadTextFileUtf8ToLocal(root / "src" / ".常量.txt", bundle.constantText);
	ReadTextFileUtf8ToLocal(root / "src" / ".全局变量.txt", bundle.globalText);

	if (const auto it = metaJson.find("sourceFiles"); it != metaJson.end() && it->is_array()) {
		for (const auto& item : *it) {
			BundleSourceFile file = SourceFileMetaFromJson(item);
			if (!ReadTextFileUtf8ToLocal(root / file.relativePath, file.content)) {
				if (outError != nullptr) {
					*outError = "read_source_file_failed";
				}
				return false;
			}
			bundle.sourceFiles.push_back(std::move(file));
		}
	}

	if (const auto it = metaJson.find("formFiles"); it != metaJson.end() && it->is_array()) {
		for (const auto& item : *it) {
			BundleFormFile file = FormFileMetaFromJson(item);
			if (!ReadTextFileUtf8ToLocal(root / file.relativePath, file.xmlText)) {
				if (outError != nullptr) {
					*outError = "read_form_file_failed";
				}
				return false;
			}
			bundle.formFiles.push_back(std::move(file));
		}
	}

	if (const auto it = metaJson.find("folders"); it != metaJson.end() && it->is_array()) {
		for (const auto& item : *it) {
			bundle.folders.push_back(FolderMetaFromJson(item));
		}
	}
	if (const auto it = metaJson.find("folderAllocatedKey"); it != metaJson.end() && it->is_number_integer()) {
		bundle.folderAllocatedKey = it->get<std::int32_t>();
	}
	if (const auto it = metaJson.find("rootChildKeys"); it != metaJson.end() && it->is_array()) {
		for (const auto& child : *it) {
			if (child.is_string()) {
				bundle.rootChildKeys.push_back(Utf8ToLocalText(child.get<std::string>()));
			}
		}
	}
	if (const auto it = metaJson.find("windowBindings"); it != metaJson.end() && it->is_array()) {
		for (const auto& item : *it) {
			bundle.windowBindings.push_back(WindowBindingFromJson(item));
		}
	}
	if (const auto it = metaJson.find("nativeBundleDigest"); it != metaJson.end() && it->is_string()) {
		bundle.nativeBundleDigest = Utf8ToLocalText(it->get<std::string>());
	}
	(void)ReadFileBytes(GetNativeSourceSnapshotPath(root), bundle.nativeSourceBytes);
	json nativeSourceMapJson;
	if (ReadJsonFile(GetNativeSourceMapPath(root), nativeSourceMapJson) && nativeSourceMapJson.is_array()) {
		for (const auto& item : nativeSourceMapJson) {
			bundle.nativeSourceSnapshots.push_back(NativeSourceSnapshotFromJson(item));
		}
	}
	json nativeSymbolMapJson;
	if (ReadJsonFile(GetNativeSymbolMapPath(root), nativeSymbolMapJson) && nativeSymbolMapJson.is_object()) {
		if (const auto it = nativeSymbolMapJson.find("programHeader");
			it != nativeSymbolMapJson.end() && it->is_object()) {
			bundle.nativeProgramHeader = NativeProgramHeaderSnapshotFromJson(*it);
		}
		if (const auto it = nativeSymbolMapJson.find("globals"); it != nativeSymbolMapJson.end() && it->is_array()) {
			for (const auto& item : *it) {
				bundle.nativeGlobalSnapshots.push_back(NativeGlobalSnapshotFromJson(item));
			}
		}
		if (const auto it = nativeSymbolMapJson.find("structs"); it != nativeSymbolMapJson.end() && it->is_array()) {
			for (const auto& item : *it) {
				bundle.nativeStructSnapshots.push_back(NativeStructSnapshotFromJson(item));
			}
		}
		if (const auto it = nativeSymbolMapJson.find("dlls"); it != nativeSymbolMapJson.end() && it->is_array()) {
			for (const auto& item : *it) {
				bundle.nativeDllSnapshots.push_back(NativeDllSnapshotFromJson(item));
			}
		}
		if (const auto it = nativeSymbolMapJson.find("constants"); it != nativeSymbolMapJson.end() && it->is_array()) {
			for (const auto& item : *it) {
				bundle.nativeConstantSnapshots.push_back(NativeConstantSnapshotFromJson(item));
			}
		}
	}

	const auto readResourceList = [&](const std::filesystem::path& path, BundleResourceKind kind) -> bool {
		json listJson;
		if (!ReadJsonFile(path, listJson)) {
			return false;
		}
		const auto it = listJson.find("items");
		if (it == listJson.end() || !it->is_array()) {
			return false;
		}
		for (const auto& item : *it) {
			BundleBinaryResource resource = ResourceMetaFromJson(item, kind);
			if (!ReadFileBytes(root / resource.relativePath, resource.data)) {
				return false;
			}
			bundle.resources.push_back(std::move(resource));
		}
		return true;
	};

	if (!readResourceList(root / "image" / "list.json", BundleResourceKind::Image)) {
		if (outError != nullptr) {
			*outError = "read_image_list_failed";
		}
		return false;
	}
	if (!readResourceList(root / "audio" / "list.json", BundleResourceKind::Sound)) {
		if (outError != nullptr) {
			*outError = "read_audio_list_failed";
		}
		return false;
	}

	if (!ReadAutoDiscoveredSourceFiles(root, bundle, outError)) {
		return false;
	}

	outBundle = std::move(bundle);
	return true;
}

}  // namespace e2txt
