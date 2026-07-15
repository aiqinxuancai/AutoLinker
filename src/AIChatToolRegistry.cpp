#include "AIChatToolRegistry.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <format>
#include <limits>
#include <string>

namespace AIChatToolRegistry {
namespace {

constexpr std::array<ToolMetadata, 28> kTools = {{
	{"refresh_workspace_mirror", true, false, false, false, false},
	{"update_plan", false, false, false, false, false},
	{"list_files", true, false, true, false, false},
	{"search_code", true, false, true, false, false},
	{"read_file", true, false, true, false, false},
	{"read_files", true, false, true, false, false},
	{"read_code_item", true, false, true, false, false},
	{"read_real_file", true, false, true, false, false},
	{"edit_file", true, false, true, true, false},
	{"multi_edit_file", true, false, true, true, false},
	{"write_file", true, false, true, true, false},
	{"diff_file", true, false, true, false, false},
	{"restore_file_snapshot", true, false, true, true, false},
	{"add_new_file", true, false, false, true, false},
	{"get_current_page_info", true, false, false, false, false},
	{"get_current_eide_info", true, false, false, false, false},
	{"refresh_dependency_catalog", false, true, false, false, false},
	{"search_available_modules", false, true, false, false, false},
	{"search_available_support_libraries", false, true, false, false, false},
	{"list_imported_modules", false, true, false, false, false},
	{"add_module_to_project", false, true, false, true, true},
	{"remove_module_from_project", false, true, false, true, false},
	{"add_support_library_to_project", false, true, false, true, false},
	{"compile_with_output_path", true, false, false, true, false},
	{"run_powershell_command", true, false, false, true, true},
	{"search_web_tavily", true, false, false, false, false},
	{"fetch_url", true, false, false, false, false},
	{"extract_web_document", true, false, false, false, false},
}};

std::string BuildPath(const std::string& parent, const std::string& child)
{
	if (parent.empty()) {
		return child;
	}
	if (!child.empty() && child.front() == '[') {
		return parent + child;
	}
	return parent + "." + child;
}

bool JsonValuesEqual(const nlohmann::json& left, const nlohmann::json& right)
{
	if (left.type() == right.type()) {
		return left == right;
	}
	if (left.is_number() && right.is_number()) {
		return std::fabs(left.get<double>() - right.get<double>()) <=
			(std::numeric_limits<double>::epsilon)();
	}
	return false;
}

bool ValidateSchemaNode(
	const nlohmann::json& value,
	const nlohmann::json& schema,
	const std::string& path,
	int depth,
	std::string& outError)
{
	if (!schema.is_object()) {
		return true;
	}
	if (depth > 24) {
		outError = path + " schema nesting is too deep";
		return false;
	}

	if (schema.contains("enum") && schema["enum"].is_array()) {
		bool matched = false;
		for (const auto& item : schema["enum"]) {
			if (JsonValuesEqual(value, item)) {
				matched = true;
				break;
			}
		}
		if (!matched) {
			outError = path + " is not one of the allowed enum values";
			return false;
		}
	}

	const std::string type = schema.value("type", std::string());
	if (type == "object") {
		if (!value.is_object()) {
			outError = path + " must be an object";
			return false;
		}
		const nlohmann::json properties = schema.value("properties", nlohmann::json::object());
		if (schema.contains("required") && schema["required"].is_array()) {
			for (const auto& item : schema["required"]) {
				if (!item.is_string()) {
					continue;
				}
				const std::string key = item.get<std::string>();
				if (!value.contains(key)) {
					outError = BuildPath(path, key) + " is required";
					return false;
				}
			}
		}
		const bool allowAdditional = !schema.contains("additionalProperties") ||
			!schema["additionalProperties"].is_boolean() ||
			schema["additionalProperties"].get<bool>();
		for (const auto& item : value.items()) {
			const auto propertyIt = properties.find(item.key());
			if (propertyIt == properties.end()) {
				if (!allowAdditional) {
					outError = BuildPath(path, item.key()) + " is not allowed";
					return false;
				}
				continue;
			}
			if (!ValidateSchemaNode(item.value(), *propertyIt, BuildPath(path, item.key()), depth + 1, outError)) {
				return false;
			}
		}
		return true;
	}

	if (type == "array") {
		if (!value.is_array()) {
			outError = path + " must be an array";
			return false;
		}
		const std::size_t minItems = schema.value("minItems", static_cast<std::size_t>(0));
		const std::size_t maxItems = schema.value("maxItems", static_cast<std::size_t>(4096));
		if (value.size() < minItems || value.size() > maxItems) {
			outError = std::format("{} item count must be in [{}, {}]", path, minItems, maxItems);
			return false;
		}
		if (schema.contains("items")) {
			for (std::size_t i = 0; i < value.size(); ++i) {
				if (!ValidateSchemaNode(
						value[i],
						schema["items"],
						BuildPath(path, std::format("[{}]", i)),
						depth + 1,
						outError)) {
					return false;
				}
			}
		}
		return true;
	}

	if (type == "string") {
		if (!value.is_string()) {
			outError = path + " must be a string";
			return false;
		}
		const std::size_t size = value.get_ref<const std::string&>().size();
		const std::size_t minLength = schema.value("minLength", static_cast<std::size_t>(0));
		const std::size_t maxLength = schema.value("maxLength", static_cast<std::size_t>(2 * 1024 * 1024));
		if (size < minLength || size > maxLength) {
			outError = std::format("{} byte length must be in [{}, {}]", path, minLength, maxLength);
			return false;
		}
		return true;
	}

	if (type == "integer") {
		if (!value.is_number_integer()) {
			outError = path + " must be an integer";
			return false;
		}
		const long long number = value.get<long long>();
		if (schema.contains("minimum") && schema["minimum"].is_number() &&
			number < schema["minimum"].get<long long>()) {
			outError = path + " is below minimum";
			return false;
		}
		if (schema.contains("maximum") && schema["maximum"].is_number() &&
			number > schema["maximum"].get<long long>()) {
			outError = path + " is above maximum";
			return false;
		}
		return true;
	}

	if (type == "number") {
		if (!value.is_number()) {
			outError = path + " must be a number";
			return false;
		}
		return true;
	}

	if (type == "boolean") {
		if (!value.is_boolean()) {
			outError = path + " must be a boolean";
			return false;
		}
		return true;
	}

	return true;
}

} // namespace

const ToolMetadata* Find(std::string_view toolName)
{
	const auto it = std::find_if(kTools.begin(), kTools.end(), [toolName](const ToolMetadata& item) {
		return item.name == toolName;
	});
	return it == kTools.end() ? nullptr : &*it;
}

bool IsExternalPublic(std::string_view toolName)
{
	const ToolMetadata* metadata = Find(toolName);
	return metadata != nullptr && metadata->externalPublic;
}

bool IsDependencyManagement(std::string_view toolName)
{
	const ToolMetadata* metadata = Find(toolName);
	return metadata != nullptr && metadata->dependencyManagement;
}

bool RequiresWorkspaceRefresh(std::string_view toolName)
{
	const ToolMetadata* metadata = Find(toolName);
	return metadata != nullptr && metadata->externalPublic && metadata->requiresWorkspaceRefresh;
}

nlohmann::json FilterExternalPublicCatalog(const nlohmann::json& catalog)
{
	nlohmann::json filtered = nlohmann::json::array();
	if (!catalog.is_array()) {
		return filtered;
	}
	for (const auto& item : catalog) {
		if (!item.is_object()) {
			continue;
		}
		const std::string name = item.value("name", std::string());
		if (IsExternalPublic(name)) {
			filtered.push_back(item);
		}
	}
	return filtered;
}

bool ValidateArguments(
	const nlohmann::json& arguments,
	const nlohmann::json& inputSchema,
	std::string& outError)
{
	outError.clear();
	if (!arguments.is_object()) {
		outError = "arguments must be an object";
		return false;
	}
	return ValidateSchemaNode(arguments, inputSchema, "arguments", 0, outError);
}

} // namespace AIChatToolRegistry
