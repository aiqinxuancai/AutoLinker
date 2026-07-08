#include "AIChatMcpClient.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <functional>
#include <format>
#include <mutex>
#include <optional>
#include <thread>
#include <sstream>
#include <unordered_map>

#include "AutoLinkerVersion.h"
#include "Logger.h"
#include "WinINetUtil.h"

namespace {

constexpr const char* kMcpProtocolVersion = "2025-11-25";
constexpr const char* kMcpCompatProtocolVersion = "2025-03-26";
constexpr const char* kMcpToolPrefix = "mcp_";
constexpr size_t kMaxMcpDescriptionBytes = 900;
constexpr auto kToolCacheTtl = std::chrono::seconds(30);

struct McpToolMapping {
	AIChatMcpToolInfo tool;
	AIChatMcpServerConfig server;
};

struct McpToolCache {
	std::mutex mutex;
	std::chrono::steady_clock::time_point refreshedAt = {};
	std::string configFingerprint;
	std::vector<AIChatMcpToolInfo> tools;
	std::unordered_map<std::string, McpToolMapping> mappings;
};

struct JsonRpcResult {
	bool ok = false;
	bool cancelled = false;
	int httpStatus = 0;
	std::string sessionId;
	nlohmann::json payload = nlohmann::json::object();
	std::string error;
};

McpToolCache g_toolCache;

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string TrimAscii(std::string text)
{
	while (!text.empty() && (text.back() == ' ' || text.back() == '\t' || text.back() == '\r' || text.back() == '\n')) {
		text.pop_back();
	}
	size_t begin = 0;
	while (begin < text.size() && (text[begin] == ' ' || text[begin] == '\t' || text[begin] == '\r' || text[begin] == '\n')) {
		++begin;
	}
	if (begin > 0) {
		text.erase(0, begin);
	}
	return text;
}

bool IsStdioTransport(const AIChatMcpServerConfig& server)
{
	return ToLowerAscii(TrimAscii(server.transport)) == "stdio";
}

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0) > 0;
}

std::string LocalToUtf8Text(const std::string& text)
{
	if (text.empty() || IsValidUtf8(text)) {
		return text;
	}
	const int wideLen = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLen <= 0) {
		return text;
	}
	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), wideLen) <= 0) {
		return text;
	}
	const int utf8Len = WideCharToMultiByte(CP_UTF8, 0, wide.data(), wideLen, nullptr, 0, nullptr, nullptr);
	if (utf8Len <= 0) {
		return text;
	}
	std::string utf8(static_cast<size_t>(utf8Len), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, wide.data(), wideLen, utf8.data(), utf8Len, nullptr, nullptr) <= 0) {
		return text;
	}
	return utf8;
}

std::string Utf8ToLocalText(const std::string& text)
{
	if (text.empty()) {
		return text;
	}
	const int wideLen = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLen <= 0) {
		return text;
	}
	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), wide.data(), wideLen) <= 0) {
		return text;
	}
	const int localLen = WideCharToMultiByte(CP_ACP, 0, wide.data(), wideLen, nullptr, 0, nullptr, nullptr);
	if (localLen <= 0) {
		return text;
	}
	std::string local(static_cast<size_t>(localLen), '\0');
	if (WideCharToMultiByte(CP_ACP, 0, wide.data(), wideLen, local.data(), localLen, nullptr, nullptr) <= 0) {
		return text;
	}
	return local;
}

std::wstring Utf8ToWideText(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int wideLen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLen > 0) {
		std::wstring wide(static_cast<size_t>(wideLen), L'\0');
		if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), wide.data(), wideLen) > 0) {
			return wide;
		}
	}

	const int localWideLen = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (localWideLen <= 0) {
		return std::wstring();
	}
	std::wstring wide(static_cast<size_t>(localWideLen), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), localWideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::wstring QuoteCommandLineArgument(const std::wstring& arg)
{
	if (arg.empty()) {
		return L"\"\"";
	}
	bool needsQuote = false;
	for (wchar_t ch : arg) {
		if (ch == L' ' || ch == L'\t' || ch == L'\n' || ch == L'\r' || ch == L'"') {
			needsQuote = true;
			break;
		}
	}
	if (!needsQuote) {
		return arg;
	}

	std::wstring quoted = L"\"";
	size_t backslashes = 0;
	for (wchar_t ch : arg) {
		if (ch == L'\\') {
			++backslashes;
			continue;
		}
		if (ch == L'"') {
			quoted.append(backslashes * 2 + 1, L'\\');
			quoted.push_back(L'"');
			backslashes = 0;
			continue;
		}
		quoted.append(backslashes, L'\\');
		backslashes = 0;
		quoted.push_back(ch);
	}
	quoted.append(backslashes * 2, L'\\');
	quoted.push_back(L'"');
	return quoted;
}

std::wstring BuildStdioCommandLine(const AIChatMcpServerConfig& server)
{
	std::wstring commandLine = QuoteCommandLineArgument(Utf8ToWideText(server.command));
	for (const auto& arg : server.arguments) {
		commandLine.push_back(L' ');
		commandLine += QuoteCommandLineArgument(Utf8ToWideText(arg));
	}
	return commandLine;
}

std::wstring BuildStdioEnvironmentBlock(const AIChatMcpServerConfig& server)
{
	if (server.env.empty()) {
		return std::wstring();
	}

	LPWCH currentEnv = GetEnvironmentStringsW();
	std::unordered_map<std::wstring, std::wstring> values;
	if (currentEnv != nullptr) {
		for (LPWCH p = currentEnv; *p != L'\0';) {
			std::wstring entry = p;
			const size_t eq = entry.find(L'=');
			if (eq != std::wstring::npos && eq > 0) {
				values[entry.substr(0, eq)] = entry.substr(eq + 1);
			}
			p += entry.size() + 1;
		}
		FreeEnvironmentStringsW(currentEnv);
	}

	for (const auto& env : server.env) {
		const std::wstring name = Utf8ToWideText(env.name);
		if (!name.empty()) {
			values[name] = Utf8ToWideText(env.value);
		}
	}

	std::vector<std::wstring> keys;
	keys.reserve(values.size());
	for (const auto& item : values) {
		keys.push_back(item.first);
	}
	std::sort(keys.begin(), keys.end(), [](const std::wstring& left, const std::wstring& right) {
		return _wcsicmp(left.c_str(), right.c_str()) < 0;
	});

	std::wstring block;
	for (const auto& key : keys) {
		block += key;
		block.push_back(L'=');
		block += values[key];
		block.push_back(L'\0');
	}
	block.push_back(L'\0');
	return block;
}

uint64_t Fnv1a64(const std::string& text)
{
	uint64_t hash = 1469598103934665603ull;
	for (unsigned char ch : text) {
		hash ^= static_cast<uint64_t>(ch);
		hash *= 1099511628211ull;
	}
	return hash;
}

std::string HexHash(const std::string& text, size_t chars)
{
	const std::string hex = std::format("{:016x}", Fnv1a64(text));
	return hex.substr(0, (std::min)(chars, hex.size()));
}

std::string SlugifyAscii(std::string text, size_t maxBytes)
{
	text = ToLowerAscii(TrimAscii(text));
	std::string out;
	out.reserve((std::min)(text.size(), maxBytes));
	for (char ch : text) {
		const unsigned char uch = static_cast<unsigned char>(ch);
		if ((uch >= 'a' && uch <= 'z') || (uch >= '0' && uch <= '9')) {
			out.push_back(ch);
		}
		else if (ch == '_' || ch == '-' || ch == '.') {
			out.push_back('_');
		}
		else if (!out.empty() && out.back() != '_') {
			out.push_back('_');
		}
		if (out.size() >= maxBytes) {
			break;
		}
	}
	while (!out.empty() && out.back() == '_') {
		out.pop_back();
	}
	if (out.empty()) {
		return "tool";
	}
	return out;
}

std::string MakeModelToolName(
	const AIChatMcpServerConfig& server,
	const std::string& originalToolName,
	const std::string& schemaHash)
{
	const std::string serverKey = SlugifyAscii(server.id.empty() ? server.name : server.id, 16);
	const std::string toolSlug = SlugifyAscii(originalToolName, 32);
	std::string modelName = std::string(kMcpToolPrefix) + serverKey + "_" + toolSlug + "_" + HexHash(server.id + "\n" + originalToolName + "\n" + schemaHash, 8);
	if (modelName.size() > 64) {
		modelName.resize(64);
	}
	return modelName;
}

constexpr int kMaxSchemaNormalizeDepth = 24;

bool CopySchemaStringKeyword(const nlohmann::json& schema, nlohmann::json& out, const char* key)
{
	if (key != nullptr && schema.contains(key) && schema[key].is_string()) {
		out[key] = schema[key];
		return true;
	}
	return false;
}

bool CopySchemaNumberKeyword(const nlohmann::json& schema, nlohmann::json& out, const char* key)
{
	if (key != nullptr && schema.contains(key) && schema[key].is_number()) {
		out[key] = schema[key];
		return true;
	}
	return false;
}

bool CopySchemaBooleanKeyword(const nlohmann::json& schema, nlohmann::json& out, const char* key)
{
	if (key != nullptr && schema.contains(key) && schema[key].is_boolean()) {
		out[key] = schema[key];
		return true;
	}
	return false;
}

void AppendUniqueRequired(nlohmann::json& target, const nlohmann::json& source)
{
	if (!source.is_array()) {
		return;
	}
	if (!target.is_array()) {
		target = nlohmann::json::array();
	}
	for (const auto& item : source) {
		if (!item.is_string()) {
			continue;
		}
		const std::string value = item.get<std::string>();
		bool exists = false;
		for (const auto& current : target) {
			if (current.is_string() && current.get<std::string>() == value) {
				exists = true;
				break;
			}
		}
		if (!exists) {
			target.push_back(value);
		}
	}
}

bool ResolveLocalSchemaRef(
	const nlohmann::json& root,
	const std::string& ref,
	const nlohmann::json*& outTarget,
	std::string& outError)
{
	outTarget = nullptr;
	if (ref.empty() || ref[0] != '#') {
		outError = "external $ref is not supported: " + ref;
		return false;
	}
	if (ref == "#") {
		outTarget = &root;
		return true;
	}
	try {
		const nlohmann::json::json_pointer pointer(ref.substr(1));
		if (!root.contains(pointer)) {
			outError = "unresolved local $ref: " + ref;
			return false;
		}
		outTarget = &root.at(pointer);
		return true;
	}
	catch (const std::exception& ex) {
		outError = std::string("invalid local $ref '") + ref + "': " + ex.what();
		return false;
	}
}

bool NormalizeSchemaNode(
	const nlohmann::json& schema,
	const nlohmann::json& root,
	int depth,
	std::vector<std::string>& refStack,
	nlohmann::json& out,
	std::string& outError);

nlohmann::json MergeSchemaNodes(nlohmann::json base, const nlohmann::json& extra)
{
	if (!base.is_object()) {
		base = nlohmann::json::object();
	}
	if (!extra.is_object()) {
		return base;
	}

	for (const auto& item : extra.items()) {
		const std::string& key = item.key();
		const nlohmann::json& value = item.value();
		if (key == "required") {
			AppendUniqueRequired(base["required"], value);
		}
		else if ((key == "properties" || key == "patternProperties") && value.is_object()) {
			if (!base.contains(key) || !base[key].is_object()) {
				base[key] = nlohmann::json::object();
			}
			for (const auto& prop : value.items()) {
				if (base[key].contains(prop.key()) && base[key][prop.key()].is_object() && prop.value().is_object()) {
					base[key][prop.key()] = MergeSchemaNodes(base[key][prop.key()], prop.value());
				}
				else {
					base[key][prop.key()] = prop.value();
				}
			}
		}
		else if (key == "additionalProperties" && base.contains(key)) {
			if (base[key].is_boolean() && base[key].get<bool>() == false) {
				continue;
			}
			if (value.is_boolean() && value.get<bool>() == false) {
				base[key] = false;
			}
			else if (base[key].is_object() && value.is_object()) {
				base[key] = MergeSchemaNodes(base[key], value);
			}
			else {
				base[key] = value;
			}
		}
		else if (!base.contains(key)) {
			base[key] = value;
		}
		else if (base[key].is_object() && value.is_object() && key != "items") {
			base[key] = MergeSchemaNodes(base[key], value);
		}
		else {
			base[key] = value;
		}
	}
	return base;
}

bool NormalizeSchemaArray(
	const nlohmann::json& schemas,
	const nlohmann::json& root,
	int depth,
	std::vector<std::string>& refStack,
	nlohmann::json& out,
	std::string& outError)
{
	out = nlohmann::json::array();
	if (!schemas.is_array()) {
		outError = "schema combinator must be an array";
		return false;
	}
	for (const auto& item : schemas) {
		nlohmann::json normalized;
		if (!NormalizeSchemaNode(item, root, depth + 1, refStack, normalized, outError)) {
			return false;
		}
		out.push_back(std::move(normalized));
	}
	return true;
}

void CopySchemaTypeKeyword(const nlohmann::json& schema, nlohmann::json& out)
{
	if (!schema.contains("type")) {
		return;
	}
	if (schema["type"].is_string()) {
		out["type"] = schema["type"];
		return;
	}
	if (!schema["type"].is_array()) {
		return;
	}

	nlohmann::json anyOf = nlohmann::json::array();
	bool nullable = false;
	for (const auto& item : schema["type"]) {
		if (!item.is_string()) {
			continue;
		}
		const std::string type = item.get<std::string>();
		if (type == "null") {
			nullable = true;
			continue;
		}
		anyOf.push_back({ {"type", type} });
	}
	if (nullable) {
		out["nullable"] = true;
	}
	if (anyOf.size() == 1) {
		out["type"] = anyOf.front()["type"];
	}
	else if (!anyOf.empty()) {
		out["anyOf"] = std::move(anyOf);
	}
}

bool NormalizeSchemaNode(
	const nlohmann::json& schema,
	const nlohmann::json& root,
	int depth,
	std::vector<std::string>& refStack,
	nlohmann::json& out,
	std::string& outError)
{
	out = nlohmann::json::object();
	if (!schema.is_object()) {
		return true;
	}
	if (depth > kMaxSchemaNormalizeDepth) {
		outError = "inputSchema nesting is too deep";
		return false;
	}

	if (schema.contains("$ref")) {
		if (!schema["$ref"].is_string()) {
			outError = "$ref must be a string";
			return false;
		}
		const std::string ref = schema["$ref"].get<std::string>();
		if (std::find(refStack.begin(), refStack.end(), ref) != refStack.end()) {
			outError = "circular $ref detected: " + ref;
			return false;
		}

		const nlohmann::json* target = nullptr;
		if (!ResolveLocalSchemaRef(root, ref, target, outError)) {
			return false;
		}
		refStack.push_back(ref);
		nlohmann::json resolved;
		const bool resolvedOk = NormalizeSchemaNode(*target, root, depth + 1, refStack, resolved, outError);
		refStack.pop_back();
		if (!resolvedOk) {
			return false;
		}

		nlohmann::json siblings = schema;
		siblings.erase("$ref");
		if (siblings.empty()) {
			out = std::move(resolved);
			return true;
		}

		nlohmann::json normalizedSiblings;
		if (!NormalizeSchemaNode(siblings, root, depth + 1, refStack, normalizedSiblings, outError)) {
			return false;
		}
		out = MergeSchemaNodes(std::move(resolved), normalizedSiblings);
		return true;
	}

	CopySchemaTypeKeyword(schema, out);
	CopySchemaStringKeyword(schema, out, "title");
	CopySchemaStringKeyword(schema, out, "description");
	CopySchemaStringKeyword(schema, out, "format");
	CopySchemaStringKeyword(schema, out, "pattern");
	CopySchemaStringKeyword(schema, out, "contentEncoding");
	CopySchemaStringKeyword(schema, out, "contentMediaType");
	CopySchemaNumberKeyword(schema, out, "minimum");
	CopySchemaNumberKeyword(schema, out, "maximum");
	CopySchemaNumberKeyword(schema, out, "exclusiveMinimum");
	CopySchemaNumberKeyword(schema, out, "exclusiveMaximum");
	CopySchemaNumberKeyword(schema, out, "multipleOf");
	CopySchemaNumberKeyword(schema, out, "minLength");
	CopySchemaNumberKeyword(schema, out, "maxLength");
	CopySchemaNumberKeyword(schema, out, "minItems");
	CopySchemaNumberKeyword(schema, out, "maxItems");
	CopySchemaNumberKeyword(schema, out, "minProperties");
	CopySchemaNumberKeyword(schema, out, "maxProperties");
	CopySchemaBooleanKeyword(schema, out, "nullable");
	CopySchemaBooleanKeyword(schema, out, "deprecated");

	if (schema.contains("enum") && schema["enum"].is_array()) {
		out["enum"] = schema["enum"];
	}
	if (schema.contains("const")) {
		out["const"] = schema["const"];
	}
	if (schema.contains("default")) {
		out["default"] = schema["default"];
	}
	if (schema.contains("examples") && schema["examples"].is_array()) {
		out["examples"] = schema["examples"];
	}
	if (schema.contains("required") && schema["required"].is_array()) {
		AppendUniqueRequired(out["required"], schema["required"]);
	}
	if (schema.contains("additionalProperties")) {
		if (schema["additionalProperties"].is_boolean()) {
			out["additionalProperties"] = schema["additionalProperties"];
		}
		else if (schema["additionalProperties"].is_object()) {
			nlohmann::json normalized;
			if (!NormalizeSchemaNode(schema["additionalProperties"], root, depth + 1, refStack, normalized, outError)) {
				return false;
			}
			out["additionalProperties"] = std::move(normalized);
		}
	}
	if (schema.contains("items") && schema["items"].is_object()) {
		nlohmann::json normalized;
		if (!NormalizeSchemaNode(schema["items"], root, depth + 1, refStack, normalized, outError)) {
			return false;
		}
		out["items"] = std::move(normalized);
	}
	if (schema.contains("prefixItems") && schema["prefixItems"].is_array()) {
		nlohmann::json normalizedItems;
		if (!NormalizeSchemaArray(schema["prefixItems"], root, depth, refStack, normalizedItems, outError)) {
			return false;
		}
		out["prefixItems"] = std::move(normalizedItems);
	}
	if (schema.contains("properties") && schema["properties"].is_object()) {
		nlohmann::json properties = nlohmann::json::object();
		for (const auto& prop : schema["properties"].items()) {
			nlohmann::json normalized;
			if (!NormalizeSchemaNode(prop.value(), root, depth + 1, refStack, normalized, outError)) {
				return false;
			}
			properties[prop.key()] = std::move(normalized);
		}
		out["properties"] = std::move(properties);
	}
	if (schema.contains("patternProperties") && schema["patternProperties"].is_object()) {
		nlohmann::json properties = nlohmann::json::object();
		for (const auto& prop : schema["patternProperties"].items()) {
			nlohmann::json normalized;
			if (!NormalizeSchemaNode(prop.value(), root, depth + 1, refStack, normalized, outError)) {
				return false;
			}
			properties[prop.key()] = std::move(normalized);
		}
		out["patternProperties"] = std::move(properties);
	}
	if (schema.contains("propertyNames") && schema["propertyNames"].is_object()) {
		nlohmann::json normalized;
		if (!NormalizeSchemaNode(schema["propertyNames"], root, depth + 1, refStack, normalized, outError)) {
			return false;
		}
		out["propertyNames"] = std::move(normalized);
	}
	if (schema.contains("allOf")) {
		nlohmann::json allOf;
		if (!NormalizeSchemaArray(schema["allOf"], root, depth, refStack, allOf, outError)) {
			return false;
		}
		for (const auto& item : allOf) {
			out = MergeSchemaNodes(std::move(out), item);
		}
	}
	if (schema.contains("anyOf")) {
		nlohmann::json anyOf;
		if (!NormalizeSchemaArray(schema["anyOf"], root, depth, refStack, anyOf, outError)) {
			return false;
		}
		if (out.contains("anyOf") && out["anyOf"].is_array()) {
			for (auto& item : anyOf) {
				out["anyOf"].push_back(std::move(item));
			}
		}
		else {
			out["anyOf"] = std::move(anyOf);
		}
	}
	if (schema.contains("oneOf")) {
		nlohmann::json oneOf;
		if (!NormalizeSchemaArray(schema["oneOf"], root, depth, refStack, oneOf, outError)) {
			return false;
		}
		out["oneOf"] = std::move(oneOf);
	}
	if (schema.contains("not") && schema["not"].is_object()) {
		nlohmann::json normalized;
		if (!NormalizeSchemaNode(schema["not"], root, depth + 1, refStack, normalized, outError)) {
			return false;
		}
		out["not"] = std::move(normalized);
	}
	return true;
}

bool SchemaAllowsObjectRoot(const nlohmann::json& schema)
{
	if (!schema.is_object()) {
		return false;
	}
	if (schema.contains("type")) {
		if (schema["type"].is_string()) {
			return schema["type"].get<std::string>() == "object";
		}
		if (schema["type"].is_array()) {
			for (const auto& item : schema["type"]) {
				if (item.is_string() && item.get<std::string>() == "object") {
					return true;
				}
			}
			return false;
		}
	}
	if (schema.contains("properties") || schema.contains("additionalProperties")) {
		return true;
	}
	const auto combinatorAllowsObject = [](const nlohmann::json& values) {
		if (!values.is_array()) {
			return false;
		}
		for (const auto& item : values) {
			if (SchemaAllowsObjectRoot(item)) {
				return true;
			}
		}
		return false;
	};
	return combinatorAllowsObject(schema.value("anyOf", nlohmann::json::array())) ||
		combinatorAllowsObject(schema.value("oneOf", nlohmann::json::array())) ||
		combinatorAllowsObject(schema.value("allOf", nlohmann::json::array()));
}

bool NormalizeInputSchema(const nlohmann::json& inputSchema, nlohmann::json& outSchema, std::string& outError)
{
	outError.clear();
	outSchema = nlohmann::json::object();
	if (!inputSchema.is_object()) {
		outError = "inputSchema must be object";
		return false;
	}

	std::vector<std::string> refStack;
	if (!NormalizeSchemaNode(inputSchema, inputSchema, 0, refStack, outSchema, outError)) {
		return false;
	}
	const std::string type = outSchema.value("type", std::string());
	if (type.empty()) {
		outSchema["type"] = "object";
	}
	else if (!SchemaAllowsObjectRoot(outSchema)) {
		outError = "inputSchema root type must be object";
		return false;
	}
	if (!outSchema.contains("properties") || !outSchema["properties"].is_object()) {
		outSchema["properties"] = nlohmann::json::object();
	}
	if (!outSchema.contains("additionalProperties")) {
		outSchema["additionalProperties"] = false;
	}
	return true;
}

std::string BuildSchemaHash(const nlohmann::json& schema)
{
	return HexHash(schema.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace), 16);
}

std::string TruncateUtf8(std::string text, size_t maxBytes)
{
	if (text.size() <= maxBytes) {
		return text;
	}
	text.resize(maxBytes);
	while (!text.empty() && (static_cast<unsigned char>(text.back()) & 0xC0) == 0x80) {
		text.pop_back();
	}
	text += "...";
	return text;
}

std::string BuildCustomHeaders(const AIChatMcpServerConfig& server, const std::string& sessionId)
{
	std::string headers;
	headers += "Content-Type: application/json\r\n";
	headers += "Accept: application/json, text/event-stream\r\n";
	headers += std::string("MCP-Protocol-Version: ") + kMcpProtocolVersion + "\r\n";
	if (!sessionId.empty()) {
		headers += "Mcp-Session-Id: " + sessionId + "\r\n";
	}
	for (const auto& header : server.headers) {
		if (!header.name.empty()) {
			headers += header.name;
			headers += ": ";
			headers += header.value;
			headers += "\r\n";
		}
	}
	return headers;
}

std::string RedactHeadersForLog(const AIChatMcpServerConfig& server)
{
	std::string text;
	for (const auto& header : server.headers) {
		if (!text.empty()) {
			text += ",";
		}
		text += header.name;
		text += "=<redacted>";
	}
	return text;
}

void LogMcpLine(const std::string& message)
{
	Logger::Instance().WriteAndIde("MCP", message);
}

void LogMcpJsonRpcRequest(const AIChatMcpServerConfig& server, const std::string& method, const nlohmann::json& params)
{
	nlohmann::json logged = {
		{"server_id", server.id},
		{"server_name", server.name},
		{"transport", IsStdioTransport(server) ? "stdio" : "streamable_http"},
		{"method", method},
		{"params", params}
	};
	if (IsStdioTransport(server)) {
		logged["command"] = server.command;
		logged["arguments"] = server.arguments;
	}
	else {
		logged["url"] = server.url;
	}
	const std::string headers = RedactHeadersForLog(server);
	if (!headers.empty()) {
		logged["headers"] = headers;
	}
	Logger::Instance().WriteSplit(
		"MCP",
		">> " + logged.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace),
		std::format(">> {} {}", server.id, method));
}

void LogMcpJsonRpcResponse(const AIChatMcpServerConfig& server, const std::string& method, const JsonRpcResult& result, double elapsedMs)
{
	nlohmann::json logged = {
		{"server_id", server.id},
		{"method", method},
		{"ok", result.ok},
		{"http_status", result.httpStatus},
		{"elapsed_ms", elapsedMs}
	};
	if (!result.error.empty()) {
		logged["error"] = result.error;
	}
	Logger::Instance().WriteSplit(
		"MCP",
		"<< " + logged.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace),
		std::format("<< {} {} ok={} http={} elapsed_ms={:.1f}",
			server.id,
			method,
			result.ok ? 1 : 0,
			result.httpStatus,
			elapsedMs));
}

nlohmann::json BuildRpcRequest(int id, const std::string& method, const nlohmann::json& params, bool notification)
{
	nlohmann::json request = {
		{"jsonrpc", "2.0"},
		{"method", method}
	};
	if (!notification) {
		request["id"] = id;
	}
	if (!params.is_null()) {
		request["params"] = params;
	}
	return request;
}

class StdioMcpSession {
public:
	StdioMcpSession() = default;
	~StdioMcpSession()
	{
		Close();
	}

	StdioMcpSession(const StdioMcpSession&) = delete;
	StdioMcpSession& operator=(const StdioMcpSession&) = delete;

	bool Start(const AIChatMcpServerConfig& server, std::string& outError)
	{
		outError.clear();
		if (server.command.empty()) {
			outError = "stdio command is empty";
			return false;
		}

		SECURITY_ATTRIBUTES sa = {};
		sa.nLength = sizeof(sa);
		sa.bInheritHandle = TRUE;
		sa.lpSecurityDescriptor = nullptr;

		HANDLE childStdInRead = nullptr;
		HANDLE childStdInWrite = nullptr;
		HANDLE childStdOutRead = nullptr;
		HANDLE childStdOutWrite = nullptr;
		HANDLE childStdErrWrite = nullptr;
		if (!CreatePipe(&childStdInRead, &childStdInWrite, &sa, 0) ||
			!CreatePipe(&childStdOutRead, &childStdOutWrite, &sa, 0)) {
			outError = "CreatePipe failed";
			CloseHandleIfValid(childStdInRead);
			CloseHandleIfValid(childStdInWrite);
			CloseHandleIfValid(childStdOutRead);
			CloseHandleIfValid(childStdOutWrite);
			return false;
		}
		childStdErrWrite = CreateFileW(
			L"NUL",
			GENERIC_WRITE,
			FILE_SHARE_READ | FILE_SHARE_WRITE,
			&sa,
			OPEN_EXISTING,
			FILE_ATTRIBUTE_NORMAL,
			nullptr);
		if (childStdErrWrite == INVALID_HANDLE_VALUE) {
			childStdErrWrite = nullptr;
		}
		SetHandleInformation(childStdInWrite, HANDLE_FLAG_INHERIT, 0);
		SetHandleInformation(childStdOutRead, HANDLE_FLAG_INHERIT, 0);

		STARTUPINFOW si = {};
		si.cb = sizeof(si);
		si.dwFlags = STARTF_USESTDHANDLES;
		si.hStdInput = childStdInRead;
		si.hStdOutput = childStdOutWrite;
		si.hStdError = childStdErrWrite != nullptr ? childStdErrWrite : GetStdHandle(STD_ERROR_HANDLE);

		std::wstring commandLine = BuildStdioCommandLine(server);
		std::wstring workingDirectory = Utf8ToWideText(server.workingDirectory);
		std::wstring environmentBlock = BuildStdioEnvironmentBlock(server);
		PROCESS_INFORMATION pi = {};
		const BOOL created = CreateProcessW(
			nullptr,
			commandLine.empty() ? nullptr : commandLine.data(),
			nullptr,
			nullptr,
			TRUE,
			CREATE_NO_WINDOW | (environmentBlock.empty() ? 0 : CREATE_UNICODE_ENVIRONMENT),
			environmentBlock.empty() ? nullptr : environmentBlock.data(),
			workingDirectory.empty() ? nullptr : workingDirectory.c_str(),
			&si,
			&pi);

		CloseHandleIfValid(childStdInRead);
		CloseHandleIfValid(childStdOutWrite);
		CloseHandleIfValid(childStdErrWrite);
		if (!created) {
			const DWORD error = GetLastError();
			CloseHandleIfValid(childStdInWrite);
			CloseHandleIfValid(childStdOutRead);
			outError = std::format("CreateProcessW failed, error={}", error);
			return false;
		}

		processHandle_ = pi.hProcess;
		threadHandle_ = pi.hThread;
		stdinWrite_ = childStdInWrite;
		stdoutRead_ = childStdOutRead;
		return true;
	}

	JsonRpcResult Post(
		const AIChatMcpServerConfig& server,
		const std::string& method,
		const nlohmann::json& params,
		int id,
		bool notification,
		HttpRequestCancellation* cancellation)
	{
		const nlohmann::json request = BuildRpcRequest(id, method, params, notification);
		const std::string body = request.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
		const auto start = std::chrono::steady_clock::now();
		JsonRpcResult result;
		if (!WriteFrame(body)) {
			result.error = "stdio write failed";
			LogElapsed(server, method, result, start);
			return result;
		}
		if (notification) {
			result.ok = true;
			result.httpStatus = 202;
			LogElapsed(server, method, result, start);
			return result;
		}

		nlohmann::json responsePayload;
		std::string error;
		if (!ReadJsonRpcResponse(server.timeoutMs, cancellation, id, responsePayload, error)) {
			result.cancelled = cancellation != nullptr && cancellation->IsCancelled();
			result.httpStatus = result.cancelled ? 499 : 0;
			result.error = error.empty() ? "stdio read failed" : error;
			LogElapsed(server, method, result, start);
			return result;
		}

		result.payload = std::move(responsePayload);
		if (result.payload.contains("error")) {
			result.error = result.payload["error"].is_object()
				? result.payload["error"].dump(-1, ' ', false, nlohmann::json::error_handler_t::replace)
				: result.payload["error"].dump();
			LogElapsed(server, method, result, start);
			return result;
		}
		result.ok = true;
		result.httpStatus = 200;
		LogElapsed(server, method, result, start);
		return result;
	}

	void Close()
	{
		CloseHandleIfValid(stdinWrite_);
		CloseHandleIfValid(stdoutRead_);
		if (processHandle_ != nullptr) {
			DWORD exitCode = 0;
			if (GetExitCodeProcess(processHandle_, &exitCode) && exitCode == STILL_ACTIVE) {
				TerminateProcess(processHandle_, 0);
			}
		}
		CloseHandleIfValid(threadHandle_);
		CloseHandleIfValid(processHandle_);
	}

private:
	static void CloseHandleIfValid(HANDLE& handle)
	{
		if (handle != nullptr && handle != INVALID_HANDLE_VALUE) {
			CloseHandle(handle);
		}
		handle = nullptr;
	}

	void LogElapsed(
		const AIChatMcpServerConfig& server,
		const std::string& method,
		const JsonRpcResult& result,
		const std::chrono::steady_clock::time_point& start)
	{
		const double elapsedMs = std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - start).count();
		LogMcpJsonRpcResponse(server, method, result, elapsedMs);
	}

	bool WriteAll(const std::string& text)
	{
		size_t writtenTotal = 0;
		while (writtenTotal < text.size()) {
			DWORD written = 0;
			const DWORD chunk = static_cast<DWORD>((std::min)(text.size() - writtenTotal, static_cast<size_t>(64 * 1024)));
			if (!WriteFile(stdinWrite_, text.data() + writtenTotal, chunk, &written, nullptr) || written == 0) {
				return false;
			}
			writtenTotal += written;
		}
		return true;
	}

	bool WriteFrame(const std::string& body)
	{
		if (stdinWrite_ == nullptr) {
			return false;
		}
		const std::string frame = "Content-Length: " + std::to_string(body.size()) + "\r\n\r\n" + body;
		return WriteAll(frame);
	}

	bool ReadAvailable(std::string& buffer)
	{
		DWORD available = 0;
		if (!PeekNamedPipe(stdoutRead_, nullptr, 0, nullptr, &available, nullptr)) {
			return false;
		}
		if (available == 0) {
			return true;
		}
		std::string chunk(static_cast<size_t>(available), '\0');
		DWORD read = 0;
		if (!ReadFile(stdoutRead_, chunk.data(), available, &read, nullptr)) {
			return false;
		}
		chunk.resize(static_cast<size_t>(read));
		buffer += chunk;
		return true;
	}

	static std::optional<size_t> TryParseContentLength(const std::string& headerText)
	{
		size_t begin = 0;
		while (begin <= headerText.size()) {
			size_t end = headerText.find_first_of("\r\n", begin);
			std::string line = end == std::string::npos ? headerText.substr(begin) : headerText.substr(begin, end - begin);
			if (end == std::string::npos) {
				begin = headerText.size() + 1;
			}
			else if (headerText[end] == '\r' && end + 1 < headerText.size() && headerText[end + 1] == '\n') {
				begin = end + 2;
			}
			else {
				begin = end + 1;
			}
			const std::string lower = ToLowerAscii(TrimAscii(line));
			const std::string marker = "content-length:";
			if (lower.rfind(marker, 0) != 0) {
				continue;
			}
			try {
				return static_cast<size_t>(std::stoull(TrimAscii(line.substr(marker.size()))));
			}
			catch (...) {
				return std::nullopt;
			}
		}
		return std::nullopt;
	}

	bool TryExtractFrame(std::string& buffer, std::string& outBody, std::string& outError)
	{
		const size_t headerEnd = buffer.find("\r\n\r\n");
		const size_t delimiterSize = 4;
		size_t actualHeaderEnd = headerEnd;
		size_t actualDelimiterSize = delimiterSize;
		if (actualHeaderEnd == std::string::npos) {
			actualHeaderEnd = buffer.find("\n\n");
			actualDelimiterSize = 2;
		}
		if (actualHeaderEnd == std::string::npos) {
			return false;
		}
		const std::string headerText = buffer.substr(0, actualHeaderEnd);
		const std::optional<size_t> contentLength = TryParseContentLength(headerText);
		if (!contentLength) {
			outError = "stdio frame missing Content-Length";
			return false;
		}
		const size_t bodyStart = actualHeaderEnd + actualDelimiterSize;
		if (buffer.size() < bodyStart + *contentLength) {
			outError.clear();
			return false;
		}
		outBody = buffer.substr(bodyStart, *contentLength);
		buffer.erase(0, bodyStart + *contentLength);
		return true;
	}

	static bool JsonRpcIdsEqual(const nlohmann::json& value, int expectedId)
	{
		if (value.is_number_integer()) {
			return value.get<int>() == expectedId;
		}
		if (value.is_number_unsigned()) {
			return value.get<unsigned int>() == static_cast<unsigned int>(expectedId);
		}
		if (value.is_string()) {
			return value.get<std::string>() == std::to_string(expectedId);
		}
		return false;
	}

	bool ReadJsonRpcResponse(
		int timeoutMs,
		HttpRequestCancellation* cancellation,
		int expectedId,
		nlohmann::json& outPayload,
		std::string& outError)
	{
		outPayload = nlohmann::json::object();
		outError.clear();
		for (;;) {
			std::string frameBody;
			if (!ReadFrame(timeoutMs, cancellation, frameBody, outError)) {
				return false;
			}

			nlohmann::json parsed = nlohmann::json::parse(frameBody, nullptr, false);
			if (parsed.is_discarded() || !parsed.is_object()) {
				outError = "failed to parse JSON-RPC response";
				return false;
			}
			if (!parsed.contains("id")) {
				const std::string method = parsed.value("method", std::string());
				if (!method.empty()) {
					LogMcpLine("skip stdio MCP notification: " + method);
				}
				continue;
			}
			if (!JsonRpcIdsEqual(parsed["id"], expectedId)) {
				LogMcpLine(std::format(
					"skip stdio MCP response with unexpected id, expected={} actual={}",
					expectedId,
					parsed["id"].dump(-1, ' ', false, nlohmann::json::error_handler_t::replace)));
				continue;
			}

			outPayload = std::move(parsed);
			return true;
		}
	}

	bool ReadFrame(int timeoutMs, HttpRequestCancellation* cancellation, std::string& outBody, std::string& outError)
	{
		outBody.clear();
		outError.clear();
		const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds((std::max)(1000, timeoutMs));
		for (;;) {
			if (cancellation != nullptr && cancellation->IsCancelled()) {
				outError = "request cancelled";
				return false;
			}
			if (TryExtractFrame(readBuffer_, outBody, outError)) {
				return true;
			}
			if (!outError.empty()) {
				return false;
			}
			if (!ReadAvailable(readBuffer_)) {
				outError = "stdio read failed or process exited";
				return false;
			}
			if (TryExtractFrame(readBuffer_, outBody, outError)) {
				return true;
			}
			if (!outError.empty()) {
				return false;
			}
			DWORD exitCode = 0;
			if (processHandle_ != nullptr && GetExitCodeProcess(processHandle_, &exitCode) && exitCode != STILL_ACTIVE) {
				outError = std::format("stdio process exited with code {}", exitCode);
				return false;
			}
			if (std::chrono::steady_clock::now() >= deadline) {
				outError = "stdio read timed out";
				return false;
			}
			std::this_thread::sleep_for(std::chrono::milliseconds(10));
		}
	}

	HANDLE processHandle_ = nullptr;
	HANDLE threadHandle_ = nullptr;
	HANDLE stdinWrite_ = nullptr;
	HANDLE stdoutRead_ = nullptr;
	std::string readBuffer_;
};

std::optional<nlohmann::json> ParseJsonObject(const std::string& text)
{
	nlohmann::json parsed = nlohmann::json::parse(text, nullptr, false);
	if (!parsed.is_discarded() && parsed.is_object()) {
		return parsed;
	}
	return std::nullopt;
}

std::vector<nlohmann::json> ParseSseJsonObjects(const std::string& body)
{
	std::vector<nlohmann::json> objects;
	std::string data;
	size_t begin = 0;
	const auto flush = [&]() {
		const std::string trimmed = TrimAscii(data);
		if (!trimmed.empty()) {
			if (auto parsed = ParseJsonObject(trimmed)) {
				objects.push_back(*parsed);
			}
		}
		data.clear();
	};

	while (begin <= body.size()) {
		size_t end = body.find_first_of("\r\n", begin);
		std::string line = end == std::string::npos ? body.substr(begin) : body.substr(begin, end - begin);
		if (end == std::string::npos) {
			begin = body.size() + 1;
		}
		else if (body[end] == '\r' && end + 1 < body.size() && body[end + 1] == '\n') {
			begin = end + 2;
		}
		else {
			begin = end + 1;
		}

		if (line.empty()) {
			flush();
			continue;
		}
		if (line.rfind("data:", 0) == 0) {
			if (!data.empty()) {
				data.push_back('\n');
			}
			data += TrimAscii(line.substr(5));
		}
	}
	flush();
	return objects;
}

JsonRpcResult ParseJsonRpcResponseBody(const HttpResponseDetails& response)
{
	JsonRpcResult result;
	result.httpStatus = response.statusCode;
	result.sessionId = response.GetHeaderValue("Mcp-Session-Id");
	if (response.statusCode == 499) {
		result.cancelled = true;
		result.error = "request cancelled";
		return result;
	}
	if (response.statusCode < 200 || response.statusCode >= 300) {
		result.error = std::format("http_status={} body={}", response.statusCode, TruncateUtf8(response.body, 800));
		return result;
	}
	if (TrimAscii(response.body).empty()) {
		result.ok = true;
		result.payload = nlohmann::json::object();
		return result;
	}

	std::vector<nlohmann::json> candidates;
	if (auto parsed = ParseJsonObject(response.body)) {
		candidates.push_back(*parsed);
	}
	else {
		candidates = ParseSseJsonObjects(response.body);
	}

	for (const auto& candidate : candidates) {
		if (!candidate.is_object()) {
			continue;
		}
		result.payload = candidate;
		if (candidate.contains("error")) {
			result.error = candidate["error"].is_object()
				? candidate["error"].dump(-1, ' ', false, nlohmann::json::error_handler_t::replace)
				: candidate["error"].dump();
			return result;
		}
		result.ok = true;
		return result;
	}

	result.error = "failed to parse JSON-RPC response";
	return result;
}

JsonRpcResult PostHttpJsonRpc(
	const AIChatMcpServerConfig& server,
	const std::string& method,
	const nlohmann::json& params,
	int id,
	bool notification,
	std::string& sessionId,
	HttpRequestCancellation* cancellation)
{
	LogMcpJsonRpcRequest(server, method, params);
	const nlohmann::json request = BuildRpcRequest(id, method, params, notification);
	const std::string body = request.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
	const auto start = std::chrono::steady_clock::now();
	const HttpResponseDetails http = PerformPostRequestDetailed(
		server.url,
		body,
		BuildCustomHeaders(server, sessionId),
		server.timeoutMs,
		false,
		true,
		cancellation);
	JsonRpcResult result = ParseJsonRpcResponseBody(http);
	if (!result.sessionId.empty()) {
		sessionId = result.sessionId;
	}
	const double elapsedMs = std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - start).count();
	LogMcpJsonRpcResponse(server, method, result, elapsedMs);
	return result;
}

bool InitializeHttpSession(
	const AIChatMcpServerConfig& server,
	std::string& sessionId,
	std::string& outError,
	HttpRequestCancellation* cancellation)
{
	outError.clear();
	sessionId.clear();

	nlohmann::json params = {
		{"protocolVersion", kMcpProtocolVersion},
		{"capabilities", nlohmann::json::object()},
		{"clientInfo", {
			{"name", "AutoLinker"},
			{"version", AUTOLINKER_VERSION}
		}}
	};
	JsonRpcResult init = PostHttpJsonRpc(server, "initialize", params, 1, false, sessionId, cancellation);
	if (!init.ok) {
		outError = init.error.empty() ? "initialize failed" : init.error;
		return false;
	}

	const nlohmann::json result = init.payload.value("result", nlohmann::json::object());
	const std::string protocolVersion = result.value("protocolVersion", std::string());
	if (!protocolVersion.empty() &&
		protocolVersion != kMcpProtocolVersion &&
		protocolVersion != kMcpCompatProtocolVersion) {
		LogMcpLine(std::format("server {} returned MCP protocolVersion {}", server.id, protocolVersion));
	}

	nlohmann::json initializedParams = nlohmann::json::object();
	JsonRpcResult initialized = PostHttpJsonRpc(server, "notifications/initialized", initializedParams, 0, true, sessionId, cancellation);
	if (!initialized.ok && initialized.httpStatus != 202 && initialized.httpStatus != 204) {
		outError = initialized.error.empty() ? "notifications/initialized failed" : initialized.error;
		return false;
	}
	return true;
}

bool InitializeStdioSession(
	StdioMcpSession& session,
	const AIChatMcpServerConfig& server,
	std::string& outError,
	HttpRequestCancellation* cancellation)
{
	outError.clear();
	nlohmann::json params = {
		{"protocolVersion", kMcpProtocolVersion},
		{"capabilities", nlohmann::json::object()},
		{"clientInfo", {
			{"name", "AutoLinker"},
			{"version", AUTOLINKER_VERSION}
		}}
	};
	LogMcpJsonRpcRequest(server, "initialize", params);
	JsonRpcResult init = session.Post(server, "initialize", params, 1, false, cancellation);
	if (!init.ok) {
		outError = init.error.empty() ? "initialize failed" : init.error;
		return false;
	}

	const nlohmann::json result = init.payload.value("result", nlohmann::json::object());
	const std::string protocolVersion = result.value("protocolVersion", std::string());
	if (!protocolVersion.empty() &&
		protocolVersion != kMcpProtocolVersion &&
		protocolVersion != kMcpCompatProtocolVersion) {
		LogMcpLine(std::format("server {} returned MCP protocolVersion {}", server.id, protocolVersion));
	}

	nlohmann::json initializedParams = nlohmann::json::object();
	LogMcpJsonRpcRequest(server, "notifications/initialized", initializedParams);
	JsonRpcResult initialized = session.Post(server, "notifications/initialized", initializedParams, 0, true, cancellation);
	if (!initialized.ok && initialized.httpStatus != 202 && initialized.httpStatus != 204) {
		outError = initialized.error.empty() ? "notifications/initialized failed" : initialized.error;
		return false;
	}
	return true;
}

bool ListToolsFromInitializedTransport(
	const AIChatMcpServerConfig& server,
	const std::function<JsonRpcResult()>& listCall,
	std::vector<AIChatMcpToolInfo>& outTools,
	std::string& outError)
{
	outTools.clear();
	outError.clear();

	JsonRpcResult listed = listCall();
	if (!listed.ok) {
		outError = listed.error.empty() ? "tools/list failed" : listed.error;
		return false;
	}

	const nlohmann::json result = listed.payload.value("result", nlohmann::json::object());
	const nlohmann::json tools = result.value("tools", nlohmann::json::array());
	if (!tools.is_array()) {
		outError = "tools/list result.tools must be array";
		return false;
	}

	for (const auto& item : tools) {
		if (!item.is_object() || !item.contains("name") || !item["name"].is_string()) {
			continue;
		}
		const std::string originalName = item["name"].get<std::string>();
		nlohmann::json normalizedSchema;
		std::string schemaError;
		const nlohmann::json inputSchema = item.value("inputSchema", nlohmann::json::object({ {"type", "object"}, {"properties", nlohmann::json::object()} }));
		if (!NormalizeInputSchema(inputSchema, normalizedSchema, schemaError)) {
			LogMcpLine(std::format("skip MCP tool {}.{}: {}", server.id, originalName, schemaError));
			continue;
		}

		AIChatMcpToolInfo tool;
		tool.serverId = server.id;
		tool.serverName = server.name;
		tool.originalName = originalName;
		tool.description = TruncateUtf8(item.value("description", std::string()), kMaxMcpDescriptionBytes);
		tool.inputSchema = normalizedSchema;
		tool.schemaHash = BuildSchemaHash(normalizedSchema);
		tool.modelName = MakeModelToolName(server, originalName, tool.schemaHash);
		outTools.push_back(std::move(tool));
	}
	return true;
}

bool ListToolsFromServer(
	const AIChatMcpServerConfig& server,
	std::vector<AIChatMcpToolInfo>& outTools,
	std::string& outError,
	HttpRequestCancellation* cancellation)
{
	outTools.clear();
	outError.clear();
	if (IsStdioTransport(server)) {
		StdioMcpSession session;
		if (!session.Start(server, outError)) {
			return false;
		}
		if (!InitializeStdioSession(session, server, outError, cancellation)) {
			return false;
		}
		return ListToolsFromInitializedTransport(
			server,
			[&session, &server, cancellation]() {
				nlohmann::json params = nlohmann::json::object();
				LogMcpJsonRpcRequest(server, "tools/list", params);
				return session.Post(server, "tools/list", params, 2, false, cancellation);
			},
			outTools,
			outError);
	}

	std::string sessionId;
	if (!InitializeHttpSession(server, sessionId, outError, cancellation)) {
		return false;
	}

	return ListToolsFromInitializedTransport(
		server,
		[&server, &sessionId, cancellation]() {
			return PostHttpJsonRpc(server, "tools/list", nlohmann::json::object(), 2, false, sessionId, cancellation);
		},
		outTools,
		outError);
}

nlohmann::json BuildCatalogItem(const AIChatMcpToolInfo& tool)
{
	std::string description = tool.description;
	if (!description.empty()) {
		description += "\n\n";
	}
	description += std::format("External MCP tool from server '{}'. Original MCP tool name: '{}'.", tool.serverName, tool.originalName);
	return {
		{"name", tool.modelName},
		{"description", description},
		{"inputSchema", tool.inputSchema},
		{"x_autolinker_mcp", {
			{"server_id", tool.serverId},
			{"server_name", tool.serverName},
			{"tool_name", tool.originalName},
			{"schema_hash", tool.schemaHash}
		}}
	};
}

std::string BuildConfigFingerprint(const AIChatMcpConfig& config)
{
	AIChatMcpConfig copy = config;
	copy.approvalGrants.clear();
	return AIChatMcpConfigStore::SerializeConfigJson(copy, false);
}

std::vector<AIChatMcpToolInfo> RefreshToolsLocked(const AIChatMcpConfig& config, const std::string& fingerprint)
{
	std::vector<AIChatMcpToolInfo> tools;
	std::unordered_map<std::string, McpToolMapping> mappings;
	for (const auto& server : config.servers) {
		if (!server.enabled) {
			continue;
		}
		if (IsStdioTransport(server)) {
			if (TrimAscii(server.command).empty()) {
				continue;
			}
		}
		else if (TrimAscii(server.url).empty()) {
			continue;
		}
		std::vector<AIChatMcpToolInfo> serverTools;
		std::string error;
		if (!ListToolsFromServer(server, serverTools, error, nullptr)) {
			LogMcpLine(std::format("tools/list failed server={} error={}", server.id, error));
			continue;
		}
		for (auto& tool : serverTools) {
			McpToolMapping mapping;
			mapping.tool = tool;
			mapping.server = server;
			mappings[tool.modelName] = std::move(mapping);
			tools.push_back(std::move(tool));
		}
	}
	g_toolCache.configFingerprint = fingerprint;
	g_toolCache.refreshedAt = std::chrono::steady_clock::now();
	g_toolCache.tools = tools;
	g_toolCache.mappings = std::move(mappings);
	return tools;
}

std::vector<AIChatMcpToolInfo> LoadEnabledToolsInternal(bool forceRefresh)
{
	AIChatMcpConfig config;
	std::string error;
	if (!AIChatMcpConfigStore::Load(config, &error)) {
		LogMcpLine("load MCP config failed: " + error);
		return {};
	}
	const std::string fingerprint = BuildConfigFingerprint(config);
	std::lock_guard<std::mutex> guard(g_toolCache.mutex);
	const auto now = std::chrono::steady_clock::now();
	if (!forceRefresh &&
		g_toolCache.configFingerprint == fingerprint &&
		!g_toolCache.tools.empty() &&
		now - g_toolCache.refreshedAt < kToolCacheTtl) {
		return g_toolCache.tools;
	}
	return RefreshToolsLocked(config, fingerprint);
}

bool FindToolMapping(const std::string& modelToolName, McpToolMapping& outMapping)
{
	{
		std::lock_guard<std::mutex> guard(g_toolCache.mutex);
		const auto it = g_toolCache.mappings.find(modelToolName);
		if (it != g_toolCache.mappings.end()) {
			outMapping = it->second;
			return true;
		}
	}

	LoadEnabledToolsInternal(true);
	std::lock_guard<std::mutex> guard(g_toolCache.mutex);
	const auto it = g_toolCache.mappings.find(modelToolName);
	if (it == g_toolCache.mappings.end()) {
		return false;
	}
	outMapping = it->second;
	return true;
}

nlohmann::json SummarizeMcpContent(const nlohmann::json& content)
{
	nlohmann::json summary = nlohmann::json::array();
	if (!content.is_array()) {
		return summary;
	}
	for (const auto& item : content) {
		if (!item.is_object()) {
			continue;
		}
		const std::string type = item.value("type", std::string());
		if (type == "text") {
			summary.push_back({
				{"type", "text"},
				{"text", TruncateUtf8(item.value("text", std::string()), 12000)}
			});
		}
		else if (type == "image" || type == "audio") {
			summary.push_back({
				{"type", type},
				{"mimeType", item.value("mimeType", std::string())},
				{"omitted", true}
			});
		}
		else if (type == "resource") {
			nlohmann::json resource = item.value("resource", nlohmann::json::object());
			summary.push_back({
				{"type", "resource"},
				{"uri", resource.value("uri", std::string())},
				{"mimeType", resource.value("mimeType", std::string())},
				{"omitted", true}
			});
		}
		else {
			summary.push_back({
				{"type", type.empty() ? "unknown" : type},
				{"omitted", true}
			});
		}
	}
	return summary;
}

nlohmann::json BuildMcpResultEnvelope(
	const McpToolMapping& mapping,
	const std::string& argumentsJsonUtf8,
	const nlohmann::json& result,
	bool toolOk,
	int httpStatus)
{
	nlohmann::json arguments = argumentsJsonUtf8.empty()
		? nlohmann::json::object()
		: nlohmann::json::parse(argumentsJsonUtf8, nullptr, false);
	if (arguments.is_discarded()) {
		arguments = nlohmann::json::object();
	}
	nlohmann::json envelope = {
		{"ok", toolOk},
		{"source", "mcp"},
		{"mcp", {
			{"server_id", mapping.tool.serverId},
			{"server_name", mapping.tool.serverName},
			{"tool_name", mapping.tool.originalName},
			{"model_tool_name", mapping.tool.modelName},
			{"schema_hash", mapping.tool.schemaHash},
			{"http_status", httpStatus}
		}},
		{"arguments", std::move(arguments)}
	};
	if (result.contains("structuredContent")) {
		envelope["structuredContent"] = result["structuredContent"];
	}
	if (result.contains("content")) {
		envelope["content"] = SummarizeMcpContent(result["content"]);
	}
	if (result.contains("isError")) {
		envelope["isError"] = result["isError"];
	}
	return envelope;
}

AIChatMcpExecutionResult BuildErrorExecutionResult(
	const McpToolMapping* mapping,
	const std::string& errorUtf8,
	bool denied,
	bool cancelled,
	int httpStatus)
{
	nlohmann::json envelope = {
		{"ok", false},
		{"source", "mcp"},
		{"error", errorUtf8},
		{"denied", denied},
		{"cancelled", cancelled},
		{"http_status", httpStatus}
	};
	if (mapping != nullptr) {
		envelope["mcp"] = {
			{"server_id", mapping->tool.serverId},
			{"server_name", mapping->tool.serverName},
			{"tool_name", mapping->tool.originalName},
			{"model_tool_name", mapping->tool.modelName},
			{"schema_hash", mapping->tool.schemaHash}
		};
	}
	AIChatMcpExecutionResult result;
	result.ok = false;
	result.denied = denied;
	result.cancelled = cancelled;
	result.httpStatus = httpStatus;
	result.errorUtf8 = errorUtf8;
	result.resultJsonLocal = Utf8ToLocalText(envelope.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace));
	return result;
}

bool IsGrantAllowed(const McpToolMapping& mapping)
{
	AIChatMcpConfig config;
	if (!AIChatMcpConfigStore::Load(config, nullptr)) {
		return false;
	}
	return AIChatMcpConfigStore::HasApprovalGrant(
		config,
		mapping.tool.serverId,
		mapping.tool.originalName,
		mapping.tool.schemaHash);
}

void SaveGrant(const McpToolMapping& mapping)
{
	AIChatMcpConfig config;
	std::string error;
	if (!AIChatMcpConfigStore::Load(config, &error)) {
		LogMcpLine("load MCP config before grant failed: " + error);
		return;
	}
	AIChatMcpConfigStore::UpsertApprovalGrant(
		config,
		mapping.tool.serverId,
		mapping.tool.originalName,
		mapping.tool.schemaHash);
	if (!AIChatMcpConfigStore::Save(config, &error)) {
		LogMcpLine("save MCP approval grant failed: " + error);
	}
}

void SaveServerGrant(const McpToolMapping& mapping)
{
	AIChatMcpConfig config;
	std::string error;
	if (!AIChatMcpConfigStore::Load(config, &error)) {
		LogMcpLine("load MCP config before server grant failed: " + error);
		return;
	}
	AIChatMcpConfigStore::UpsertApprovalGrant(config, mapping.tool.serverId, "*", "*");
	if (!AIChatMcpConfigStore::Save(config, &error)) {
		LogMcpLine("save MCP server approval grant failed: " + error);
	}
}

AIChatMcpExecutionResult CallMcpTool(
	const McpToolMapping& mapping,
	const std::string& argumentsJsonUtf8,
	HttpRequestCancellation* cancellation)
{
	nlohmann::json arguments = nlohmann::json::parse(argumentsJsonUtf8, nullptr, false);
	if (arguments.is_discarded() || !arguments.is_object()) {
		return BuildErrorExecutionResult(&mapping, "MCP tool arguments must be a JSON object", false, false, 0);
	}

	nlohmann::json params = {
		{"name", mapping.tool.originalName},
		{"arguments", arguments}
	};
	JsonRpcResult called;
	if (IsStdioTransport(mapping.server)) {
		std::string error;
		StdioMcpSession session;
		if (!session.Start(mapping.server, error)) {
			return BuildErrorExecutionResult(&mapping, error.empty() ? "stdio process start failed" : error, false, false, 0);
		}
		if (!InitializeStdioSession(session, mapping.server, error, cancellation)) {
			return BuildErrorExecutionResult(&mapping, error.empty() ? "initialize failed" : error, false, false, 0);
		}
		LogMcpJsonRpcRequest(mapping.server, "tools/call", params);
		called = session.Post(mapping.server, "tools/call", params, 2, false, cancellation);
	}
	else {
		std::string sessionId;
		std::string error;
		if (!InitializeHttpSession(mapping.server, sessionId, error, cancellation)) {
			return BuildErrorExecutionResult(&mapping, error.empty() ? "initialize failed" : error, false, false, 0);
		}
		called = PostHttpJsonRpc(mapping.server, "tools/call", params, 2, false, sessionId, cancellation);
	}
	if (!called.ok) {
		return BuildErrorExecutionResult(
			&mapping,
			called.error.empty() ? "tools/call failed" : called.error,
			false,
			called.cancelled,
			called.httpStatus);
	}

	const nlohmann::json result = called.payload.value("result", nlohmann::json::object());
	const bool toolOk = !result.value("isError", false);
	const nlohmann::json envelope = BuildMcpResultEnvelope(mapping, argumentsJsonUtf8, result, toolOk, called.httpStatus);

	AIChatMcpExecutionResult output;
	output.ok = toolOk;
	output.httpStatus = called.httpStatus;
	output.resultJsonLocal = Utf8ToLocalText(envelope.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace));
	if (!toolOk) {
		output.errorUtf8 = "MCP tool returned isError=true";
	}
	return output;
}

} // namespace

namespace AIChatMcpClient {

bool IsMcpModelToolName(const std::string& toolName)
{
	return toolName.rfind(kMcpToolPrefix, 0) == 0;
}

std::vector<AIChatMcpToolInfo> LoadEnabledTools()
{
	return LoadEnabledToolsInternal(false);
}

nlohmann::json AppendMcpToolsToCatalog(const nlohmann::json& baseCatalog)
{
	nlohmann::json catalog = baseCatalog.is_array() ? baseCatalog : nlohmann::json::array();
	const std::vector<AIChatMcpToolInfo> tools = LoadEnabledToolsInternal(false);
	for (const auto& tool : tools) {
		catalog.push_back(BuildCatalogItem(tool));
	}
	return catalog;
}

AIChatMcpExecutionResult ExecuteTool(
	const std::string& modelToolName,
	const std::string& argumentsJsonUtf8,
	const std::function<bool(
		const AIChatMcpApprovalContext& context,
		bool& outAutoAllow,
		bool& outAutoAllowServer)>& approvalCallback,
	const std::function<bool()>& cancelCallback,
	HttpRequestCancellation* cancellation)
{
	McpToolMapping mapping;
	if (!FindToolMapping(modelToolName, mapping)) {
		return BuildErrorExecutionResult(nullptr, "unknown MCP tool: " + modelToolName, false, false, 0);
	}
	if (cancelCallback && cancelCallback()) {
		return BuildErrorExecutionResult(&mapping, "MCP tool execution cancelled before start", false, true, 499);
	}

	if (!IsGrantAllowed(mapping)) {
		if (!approvalCallback) {
			return BuildErrorExecutionResult(&mapping, "MCP tool execution requires approval", true, false, 0);
		}
		AIChatMcpApprovalContext approval;
		approval.serverId = mapping.tool.serverId;
		approval.serverName = mapping.tool.serverName;
		approval.toolName = mapping.tool.originalName;
		approval.modelToolName = mapping.tool.modelName;
		approval.schemaHash = mapping.tool.schemaHash;
		approval.argumentsJsonUtf8 = argumentsJsonUtf8;

		bool autoAllow = false;
		bool autoAllowServer = false;
		if (!approvalCallback(approval, autoAllow, autoAllowServer)) {
			return BuildErrorExecutionResult(&mapping, "MCP tool execution denied by user", true, false, 0);
		}
		if (autoAllowServer) {
			SaveServerGrant(mapping);
		}
		else if (autoAllow) {
			SaveGrant(mapping);
		}
	}

	if (cancelCallback && cancelCallback()) {
		return BuildErrorExecutionResult(&mapping, "MCP tool execution cancelled", false, true, 499);
	}
	return CallMcpTool(mapping, argumentsJsonUtf8, cancellation);
}

std::string BuildSelfTestReportJson()
{
	nlohmann::json report = {
		{"ok", true},
		{"name", "mcp-self-test"}
	};
	nlohmann::json checks = nlohmann::json::array();
	const auto addCheck = [&report, &checks](const std::string& name, bool ok, const std::string& error = std::string()) {
		checks.push_back({ {"name", name}, {"ok", ok}, {"error", error} });
		if (!ok) {
			report["ok"] = false;
		}
	};

	try {
		AIChatMcpConfig parsedConfig;
		std::string error;
		const bool configOk = AIChatMcpConfigStore::ParseConfigJson(AIChatMcpConfigStore::BuildDefaultConfigJson(), parsedConfig, error);
		addCheck("default_config_parse", configOk && parsedConfig.servers.size() >= 2, error);

		nlohmann::json schema = {
			{"type", "object"},
			{"properties", {
				{"text", {{"type", "string"}, {"description", "hello"}}}
			}},
			{"required", nlohmann::json::array({"text"})}
		};
		nlohmann::json normalized;
		const bool schemaOk = NormalizeInputSchema(schema, normalized, error);
		addCheck("schema_normalize", schemaOk && normalized.value("type", "") == "object", error);

		nlohmann::json idaLikeCombinatorSchema = {
			{"type", "object"},
			{"properties", {
				{"queries", {
					{"anyOf", nlohmann::json::array({
						{
							{"type", "array"},
							{"items", {
								{"type", "object"},
								{"properties", {
									{"filter", {{"type", "string"}, {"description", "Name glob/regex"}}},
									{"offset", {{"type", "integer"}}},
									{"count", {{"type", "integer"}}}
								}},
								{"additionalProperties", false}
							}}
						},
						{
							{"type", "object"},
							{"properties", {
								{"filter", {{"type", "string"}}},
								{"offset", {{"type", "integer"}}},
								{"count", {{"type", "integer"}}}
							}},
							{"additionalProperties", false}
						},
						{{"type", "string"}}
					})},
					{"description", "IDA query accepts array, object, or string"}
				}}
			}},
			{"required", nlohmann::json::array({"queries"})}
		};
		const bool combinatorOk = NormalizeInputSchema(idaLikeCombinatorSchema, normalized, error);
		addCheck(
			"schema_combinator_normalize",
			combinatorOk &&
				normalized.contains("properties") &&
				normalized["properties"].contains("queries") &&
				normalized["properties"]["queries"].contains("anyOf"),
			error.empty() ? normalized.dump() : error);

		nlohmann::json refSchema = {
			{"$defs", {
				{"query", {
					{"type", "object"},
					{"properties", {
						{"addr", {{"type", "string"}}},
						{"max_instructions", {{"type", "integer"}, {"maximum", 50000}}}
					}},
					{"required", nlohmann::json::array({"addr"})},
					{"additionalProperties", false}
				}}
			}},
			{"type", "object"},
			{"properties", {
				{"request", {{"$ref", "#/$defs/query"}}}
			}},
			{"required", nlohmann::json::array({"request"})}
		};
		const bool refOk = NormalizeInputSchema(refSchema, normalized, error);
		addCheck(
			"schema_ref_resolved",
			refOk &&
				normalized.contains("properties") &&
				normalized["properties"].contains("request") &&
				normalized["properties"]["request"].contains("properties") &&
				!normalized["properties"]["request"].contains("$ref"),
			error.empty() ? normalized.dump() : error);

		AIChatMcpServerConfig server;
		server.id = "ida-pro";
		server.name = "IDA Pro";
		const std::string schemaHash = BuildSchemaHash(schema);
		const std::string modelName = MakeModelToolName(server, "decompile-function", schemaHash);
		addCheck("model_tool_name", IsMcpModelToolName(modelName) && modelName.size() <= 64, modelName);

		McpToolMapping mapping;
		mapping.tool.serverId = server.id;
		mapping.tool.serverName = server.name;
		mapping.tool.originalName = "echo";
		mapping.tool.modelName = modelName;
		mapping.tool.schemaHash = schemaHash;
		nlohmann::json result = {
			{"content", nlohmann::json::array({
				{{"type", "text"}, {"text", "hello"}}
			})},
			{"structuredContent", {{"value", 1}}}
		};
		const nlohmann::json envelope = BuildMcpResultEnvelope(mapping, R"({"text":"hello"})", result, true, 200);
		addCheck("result_envelope", envelope.value("source", "") == "mcp" && envelope.value("ok", false), envelope.dump());
	}
	catch (const std::exception& ex) {
		addCheck("exception", false, ex.what());
	}
	catch (...) {
		addCheck("exception", false, "unknown");
	}

	report["checks"] = std::move(checks);
	return report.dump(2, ' ', false, nlohmann::json::error_handler_t::replace);
}

} // namespace AIChatMcpClient
