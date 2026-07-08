#include "AIChatMcpConfig.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <fstream>
#include <format>
#include <system_error>
#include <unordered_map>

#include "..\\thirdparty\\json.hpp"

#include "PathHelper.h"

namespace {

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

std::string SanitizeId(std::string text)
{
	text = TrimAscii(text);
	std::string out;
	out.reserve(text.size());
	for (char ch : text) {
		const unsigned char uch = static_cast<unsigned char>(ch);
		if ((uch >= 'a' && uch <= 'z') ||
			(uch >= 'A' && uch <= 'Z') ||
			(uch >= '0' && uch <= '9')) {
			out.push_back(static_cast<char>(std::tolower(uch)));
		}
		else if (ch == '-' || ch == '_' || ch == '.') {
			out.push_back(ch);
		}
		else if (!out.empty() && out.back() != '-') {
			out.push_back('-');
		}
	}
	while (!out.empty() && (out.back() == '-' || out.back() == '_' || out.back() == '.')) {
		out.pop_back();
	}
	if (out.empty()) {
		return "mcp-server";
	}
	return out;
}

bool IsValidHeaderName(const std::string& name)
{
	if (name.empty()) {
		return false;
	}
	for (const unsigned char ch : name) {
		if (ch <= 32 || ch >= 127 || ch == ':') {
			return false;
		}
	}
	return true;
}

// Reject header values containing control characters (notably CR/LF), which
// would otherwise allow HTTP header injection when the value is concatenated
// into the raw request header block.
bool IsValidHeaderValue(const std::string& value)
{
	for (const unsigned char ch : value) {
		if (ch == '\r' || ch == '\n' || ch == '\0') {
			return false;
		}
	}
	return true;
}

std::string NormalizeTransport(std::string text)
{
	text = TrimAscii(text);
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	if (text == "stdio") {
		return "stdio";
	}
	return "streamable_http";
}

bool IsStdioTransport(const AIChatMcpServerConfig& server)
{
	return NormalizeTransport(server.transport) == "stdio";
}

// Whether the JSON explicitly specifies a non-empty transport. When it does not,
// transport is inferred from the presence of a command (stdio) vs url (http).
bool HasExplicitTransport(const nlohmann::json& item)
{
	return item.is_object() && item.contains("transport") &&
		item["transport"].is_string() && !TrimAscii(item["transport"].get<std::string>()).empty();
}

long long CurrentUnixMs()
{
	const auto now = std::chrono::system_clock::now();
	return std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
}

std::string ReadTextFile(const std::filesystem::path& path, std::string& outError)
{
	outError.clear();
	std::ifstream input(path, std::ios::binary);
	if (!input.is_open()) {
		outError = "open config failed";
		return std::string();
	}
	return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
}

bool WriteTextFile(const std::filesystem::path& path, const std::string& text, std::string& outError)
{
	outError.clear();
	std::error_code ec;
	const std::filesystem::path parent = path.parent_path();
	if (!parent.empty()) {
		std::filesystem::create_directories(parent, ec);
		if (ec) {
			outError = "create config directory failed: " + ec.message();
			return false;
		}
	}
	std::ofstream output(path, std::ios::binary | std::ios::trunc);
	if (!output.is_open()) {
		outError = "open config for write failed";
		return false;
	}
	output.write(text.data(), static_cast<std::streamsize>(text.size()));
	return output.good();
}

std::string JsonStringValue(const nlohmann::json& value, const char* key)
{
	if (!value.is_object() || key == nullptr || !value.contains(key) || !value[key].is_string()) {
		return std::string();
	}
	return value[key].get<std::string>();
}

void ReadStringArray(const nlohmann::json& value, std::vector<std::string>& out)
{
	out.clear();
	if (!value.is_array()) {
		return;
	}
	for (const auto& item : value) {
		if (item.is_string()) {
			out.push_back(item.get<std::string>());
		}
	}
}

void ReadEnvConfig(const nlohmann::json& value, std::vector<AIChatMcpEnvConfig>& out)
{
	out.clear();
	if (value.is_object()) {
		for (const auto& item : value.items()) {
			AIChatMcpEnvConfig env;
			env.name = TrimAscii(item.key());
			if (item.value().is_string()) {
				env.value = item.value().get<std::string>();
			}
			else if (!item.value().is_null()) {
				env.value = item.value().dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
			}
			if (!env.name.empty()) {
				out.push_back(std::move(env));
			}
		}
		return;
	}
	if (!value.is_array()) {
		return;
	}
	for (const auto& item : value) {
		if (!item.is_object()) {
			continue;
		}
		AIChatMcpEnvConfig env;
		env.name = TrimAscii(JsonStringValue(item, "name"));
		env.value = JsonStringValue(item, "value");
		if (!env.name.empty()) {
			out.push_back(std::move(env));
		}
	}
}

} // namespace

namespace AIChatMcpConfigStore {

std::filesystem::path GetConfigPath()
{
	return GetAutoLinkerDirectoryPath() / "AIChatMcpConfig.json";
}

std::string BuildDefaultConfigJson()
{
	AIChatMcpConfig config;
	config.servers.push_back({
		"ida-pro",
		"IDA Pro",
		"streamable_http",
		"http://127.0.0.1:8765/mcp",
		"",
		{},
		"",
		false,
		120000,
		{},
		{}
	});
	config.servers.push_back({
		"design-mcp",
		"Design MCP",
		"streamable_http",
		"http://127.0.0.1:8770/mcp",
		"",
		{},
		"",
		false,
		120000,
		{},
		{}
	});
	config.servers.push_back({
		"stdio-example",
		"本地 stdio MCP 示例",
		"stdio",
		"",
		"npx",
		{"-y", "@modelcontextprotocol/server-filesystem", "D:\\git"},
		"",
		false,
		120000,
		{},
		{}
	});
	return SerializeConfigJson(config, true);
}

bool ParseConfigJson(const std::string& jsonText, AIChatMcpConfig& outConfig, std::string& outError)
{
	outError.clear();
	outConfig = {};

	nlohmann::json root = nlohmann::json::parse(jsonText, nullptr, false);
	if (root.is_discarded() || !root.is_object()) {
		outError = "MCP config must be a JSON object";
		return false;
	}

	outConfig.version = root.value("version", 1);
	if (outConfig.version <= 0) {
		outConfig.version = 1;
	}

	if (root.contains("servers")) {
		if (!root["servers"].is_array()) {
			outError = "servers must be an array";
			return false;
		}
		for (const auto& item : root["servers"]) {
			if (!item.is_object()) {
				continue;
			}
			AIChatMcpServerConfig server;
			server.id = SanitizeId(JsonStringValue(item, "id"));
			server.name = TrimAscii(JsonStringValue(item, "name"));
			server.transport = NormalizeTransport(JsonStringValue(item, "transport"));
			server.url = TrimAscii(JsonStringValue(item, "url"));
			server.command = TrimAscii(JsonStringValue(item, "command"));
			if (!HasExplicitTransport(item) && !server.command.empty()) {
				server.transport = "stdio";
			}
			server.workingDirectory = TrimAscii(JsonStringValue(item, "working_directory"));
			if (server.workingDirectory.empty()) {
				server.workingDirectory = TrimAscii(JsonStringValue(item, "cwd"));
			}
			if (item.contains("arguments")) {
				ReadStringArray(item["arguments"], server.arguments);
			}
			else if (item.contains("args")) {
				ReadStringArray(item["args"], server.arguments);
			}
			if (item.contains("env")) {
				ReadEnvConfig(item["env"], server.env);
			}
			server.enabled = item.value("enabled", false);
			server.timeoutMs = (std::clamp)(item.value("timeout_ms", 120000), 1000, 600000);
			if (server.name.empty()) {
				server.name = server.id;
			}
			if (server.enabled) {
				if (IsStdioTransport(server)) {
					if (server.command.empty()) {
						outError = std::format("enabled MCP stdio server '{}' requires command", server.id);
						return false;
					}
				}
				else if (server.url.empty()) {
					outError = std::format("enabled MCP server '{}' requires url", server.id);
					return false;
				}
			}

			if (item.contains("headers")) {
				if (!item["headers"].is_array()) {
					outError = std::format("server '{}' headers must be an array", server.id);
					return false;
				}
				for (const auto& headerItem : item["headers"]) {
					if (!headerItem.is_object()) {
						continue;
					}
					AIChatMcpHeaderConfig header;
					header.name = TrimAscii(JsonStringValue(headerItem, "name"));
					header.value = JsonStringValue(headerItem, "value");
					if (header.name.empty() && header.value.empty()) {
						continue;
					}
					if (!IsValidHeaderName(header.name)) {
						outError = std::format("server '{}' has invalid header name '{}'", server.id, header.name);
						return false;
					}
					if (!IsValidHeaderValue(header.value)) {
						outError = std::format("server '{}' header '{}' has invalid value (control characters not allowed)", server.id, header.name);
						return false;
					}
					server.headers.push_back(std::move(header));
				}
			}
			outConfig.servers.push_back(std::move(server));
		}
	}

	if (root.contains("mcpServers")) {
		if (!root["mcpServers"].is_object()) {
			outError = "mcpServers must be an object";
			return false;
		}
		// Preserve the original block verbatim so a save round-trip never drops
		// servers imported from other clients' configs.
		outConfig.mcpServersRaw = root["mcpServers"];
		for (const auto& namedServer : root["mcpServers"].items()) {
			const nlohmann::json& item = namedServer.value();
			if (!item.is_object()) {
				continue;
			}
			AIChatMcpServerConfig server;
			server.id = SanitizeId(namedServer.key());
			server.name = TrimAscii(JsonStringValue(item, "name"));
			if (server.name.empty()) {
				server.name = namedServer.key();
			}
			server.transport = NormalizeTransport(JsonStringValue(item, "transport"));
			server.url = TrimAscii(JsonStringValue(item, "url"));
			server.command = TrimAscii(JsonStringValue(item, "command"));
			if (!HasExplicitTransport(item) && !server.command.empty()) {
				server.transport = "stdio";
			}
			server.workingDirectory = TrimAscii(JsonStringValue(item, "working_directory"));
			if (server.workingDirectory.empty()) {
				server.workingDirectory = TrimAscii(JsonStringValue(item, "cwd"));
			}
			if (item.contains("arguments")) {
				ReadStringArray(item["arguments"], server.arguments);
			}
			else if (item.contains("args")) {
				ReadStringArray(item["args"], server.arguments);
			}
			if (item.contains("env")) {
				ReadEnvConfig(item["env"], server.env);
			}
			server.enabled = item.value("enabled", true);
			server.timeoutMs = (std::clamp)(item.value("timeout_ms", 120000), 1000, 600000);
			if (server.enabled && IsStdioTransport(server) && server.command.empty()) {
				outError = std::format("enabled MCP stdio server '{}' requires command", server.id);
				return false;
			}
			if (server.enabled && !IsStdioTransport(server) && server.url.empty()) {
				outError = std::format("enabled MCP server '{}' requires url", server.id);
				return false;
			}
			server.readOnly = true;
			outConfig.servers.push_back(std::move(server));
		}
	}

	// Ensure server ids are unique. Distinct raw ids can collapse to the same
	// sanitized id (or both blocks may define the same id); without this a later
	// server would silently overwrite an earlier one's tool routing while still
	// exposing duplicate catalog entries. Suffix collisions with -2, -3, ...
	{
		std::unordered_map<std::string, int> idCounts;
		for (auto& server : outConfig.servers) {
			auto it = idCounts.find(server.id);
			if (it == idCounts.end()) {
				idCounts.emplace(server.id, 1);
				continue;
			}
			std::string candidate;
			do {
				++it->second;
				candidate = server.id + "-" + std::to_string(it->second);
			} while (idCounts.count(candidate) != 0);
			idCounts.emplace(candidate, 1);
			server.id = candidate;
		}
	}

	if (root.contains("approval_grants")) {
		if (!root["approval_grants"].is_array()) {
			outError = "approval_grants must be an array";
			return false;
		}
		for (const auto& item : root["approval_grants"]) {
			if (!item.is_object()) {
				continue;
			}
			AIChatMcpApprovalGrant grant;
			grant.serverId = SanitizeId(JsonStringValue(item, "server_id"));
			grant.toolName = JsonStringValue(item, "tool_name");
			grant.schemaHash = JsonStringValue(item, "schema_hash");
			grant.createdAtUnixMs = item.value("created_at_unix_ms", 0LL);
			grant.updatedAtUnixMs = item.value("updated_at_unix_ms", 0LL);
			if (!grant.serverId.empty() && !grant.toolName.empty() && !grant.schemaHash.empty()) {
				outConfig.approvalGrants.push_back(std::move(grant));
			}
		}
	}

	return true;
}

std::string SerializeConfigJson(const AIChatMcpConfig& config, bool pretty)
{
	nlohmann::json root = nlohmann::json::object();
	root["version"] = config.version <= 0 ? 1 : config.version;
	root["servers"] = nlohmann::json::array();
	for (const auto& server : config.servers) {
		// Read-only servers come from the mcpServers block, which is written back
		// verbatim below; do not duplicate them into the servers array.
		if (server.readOnly) {
			continue;
		}
		nlohmann::json item = {
			{"id", server.id},
			{"name", server.name},
			{"transport", IsStdioTransport(server) ? "stdio" : "streamable_http"},
			{"url", server.url},
			{"command", server.command},
			{"arguments", server.arguments},
			{"working_directory", server.workingDirectory},
			{"enabled", server.enabled},
			{"timeout_ms", server.timeoutMs}
		};
		item["headers"] = nlohmann::json::array();
		for (const auto& header : server.headers) {
			item["headers"].push_back({
				{"name", header.name},
				{"value", header.value}
			});
		}
		item["env"] = nlohmann::json::array();
		for (const auto& env : server.env) {
			item["env"].push_back({
				{"name", env.name},
				{"value", env.value}
			});
		}
		root["servers"].push_back(std::move(item));
	}

	// Preserve the mcpServers block verbatim so imported (read-only) servers are
	// never clobbered by a save originating from our own UI.
	if (config.mcpServersRaw.is_object() && !config.mcpServersRaw.empty()) {
		root["mcpServers"] = config.mcpServersRaw;
	}

	root["approval_grants"] = nlohmann::json::array();
	for (const auto& grant : config.approvalGrants) {
		root["approval_grants"].push_back({
			{"server_id", grant.serverId},
			{"tool_name", grant.toolName},
			{"schema_hash", grant.schemaHash},
			{"created_at_unix_ms", grant.createdAtUnixMs},
			{"updated_at_unix_ms", grant.updatedAtUnixMs}
		});
	}

	return pretty
		? root.dump(4, ' ', false, nlohmann::json::error_handler_t::replace)
		: root.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

std::string SerializeConfigForUi(const AIChatMcpConfig& config)
{
	nlohmann::json root = nlohmann::json::object();
	root["version"] = config.version <= 0 ? 1 : config.version;
	root["servers"] = nlohmann::json::array();
	for (const auto& server : config.servers) {
		nlohmann::json item = {
			{"id", server.id},
			{"name", server.name},
			{"transport", IsStdioTransport(server) ? "stdio" : "streamable_http"},
			{"url", server.url},
			{"command", server.command},
			{"arguments", server.arguments},
			{"working_directory", server.workingDirectory},
			{"enabled", server.enabled},
			{"timeout_ms", server.timeoutMs},
			// Surfaced to the UI so read-only (mcpServers) entries render as
			// non-editable and are excluded from the saved servers array.
			{"read_only", server.readOnly}
		};
		item["headers"] = nlohmann::json::array();
		for (const auto& header : server.headers) {
			item["headers"].push_back({ {"name", header.name}, {"value", header.value} });
		}
		item["env"] = nlohmann::json::array();
		for (const auto& env : server.env) {
			item["env"].push_back({ {"name", env.name}, {"value", env.value} });
		}
		root["servers"].push_back(std::move(item));
	}

	root["approval_grants"] = nlohmann::json::array();
	for (const auto& grant : config.approvalGrants) {
		root["approval_grants"].push_back({
			{"server_id", grant.serverId},
			{"tool_name", grant.toolName},
			{"schema_hash", grant.schemaHash},
			{"created_at_unix_ms", grant.createdAtUnixMs},
			{"updated_at_unix_ms", grant.updatedAtUnixMs}
		});
	}
	return root.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

bool Load(AIChatMcpConfig& outConfig, std::string* outError)
{
	outConfig = {};
	const std::filesystem::path path = GetConfigPath();
	if (!std::filesystem::exists(path)) {
		outConfig.version = 1;
		return true;
	}

	std::string error;
	const std::string text = ReadTextFile(path, error);
	if (!error.empty()) {
		if (outError != nullptr) {
			*outError = error;
		}
		return false;
	}
	if (!ParseConfigJson(text, outConfig, error)) {
		if (outError != nullptr) {
			*outError = error;
		}
		return false;
	}
	return true;
}

bool Save(const AIChatMcpConfig& config, std::string* outError)
{
	std::string error;
	const bool ok = WriteTextFile(GetConfigPath(), SerializeConfigJson(config, true), error);
	if (!ok && outError != nullptr) {
		*outError = error.empty() ? "write config failed" : error;
	}
	return ok;
}

bool SaveJsonText(const std::string& jsonText, std::string* outError, bool preserveMissingMcpServers)
{
	AIChatMcpConfig parsed;
	std::string error;
	if (!ParseConfigJson(jsonText, parsed, error)) {
		if (outError != nullptr) {
			*outError = error;
		}
		return false;
	}
	// The config UI only edits the `servers` array and never emits mcpServers, so
	// it opts in to preserving the on-disk block. The native raw-JSON editor sends
	// the full document, so omitting/clearing mcpServers there is a deliberate
	// deletion and must be honored (preserveMissingMcpServers=false).
	if (preserveMissingMcpServers &&
		(!parsed.mcpServersRaw.is_object() || parsed.mcpServersRaw.empty())) {
		AIChatMcpConfig existing;
		if (Load(existing, nullptr) && existing.mcpServersRaw.is_object() && !existing.mcpServersRaw.empty()) {
			parsed.mcpServersRaw = std::move(existing.mcpServersRaw);
		}
	}
	return Save(parsed, outError);
}

bool HasApprovalGrant(
	const AIChatMcpConfig& config,
	const std::string& serverId,
	const std::string& toolName,
	const std::string& schemaHash)
{
	for (const auto& grant : config.approvalGrants) {
		if (grant.serverId != serverId) {
			continue;
		}
		// Grants are matched per tool and per input-schema hash. A previously
		// stored blanket "*"/"*" grant (legacy behavior) is intentionally NOT
		// honored here: a new tool or a changed schema must be re-approved.
		if (grant.toolName == toolName && grant.schemaHash == schemaHash) {
			return true;
		}
	}
	return false;
}

void UpsertApprovalGrant(
	AIChatMcpConfig& config,
	const std::string& serverId,
	const std::string& toolName,
	const std::string& schemaHash)
{
	const long long now = CurrentUnixMs();
	for (auto& grant : config.approvalGrants) {
		if (grant.serverId == serverId &&
			grant.toolName == toolName &&
			grant.schemaHash == schemaHash) {
			grant.updatedAtUnixMs = now;
			if (grant.createdAtUnixMs == 0) {
				grant.createdAtUnixMs = now;
			}
			return;
		}
	}

	AIChatMcpApprovalGrant grant;
	grant.serverId = serverId;
	grant.toolName = toolName;
	grant.schemaHash = schemaHash;
	grant.createdAtUnixMs = now;
	grant.updatedAtUnixMs = now;
	config.approvalGrants.push_back(std::move(grant));
}

} // namespace AIChatMcpConfigStore
