#include "EPackagerTool.h"

#include <exception>
#include <string>

#include "WorkspaceMirror.h"

namespace EPackagerTool {

nlohmann::json Execute(const std::string& argumentsJsonUtf8, bool& outOk)
{
	outOk = false;
	nlohmann::json args;
	try {
		args = argumentsJsonUtf8.empty()
			? nlohmann::json::object()
			: nlohmann::json::parse(argumentsJsonUtf8);
	}
	catch (const std::exception& ex) {
		return {
			{"ok", false},
			{"error", std::string("invalid arguments json: ") + ex.what()}
		};
	}

	if (!args.is_object()) {
		return {
			{"ok", false},
			{"error", "arguments must be an object"}
		};
	}
	const auto pathIt = args.find("file_path");
	if (pathIt == args.end() || !pathIt->is_string() || pathIt->get_ref<const std::string&>().empty()) {
		return {
			{"ok", false},
			{"error", "file_path is required"}
		};
	}

	WorkspaceMirror::ReferenceUnpackResult unpackResult;
	std::string error;
	if (!WorkspaceMirror::UnpackReferenceSource(
			pathIt->get_ref<const std::string&>(),
			unpackResult,
			error)) {
		return {
			{"ok", false},
			{"operation", "unpack"},
			{"error", error.empty() ? "e-packager unpack failed" : error}
		};
	}

	outOk = true;
	return {
		{"ok", true},
		{"operation", "unpack"},
		{"source_file_path", unpackResult.sourcePathUtf8},
		{"output_directory", unpackResult.outputDirectoryUtf8},
		{"file_count", unpackResult.fileCount},
		{"mirror_generation", unpackResult.generation},
		{"read_hint", "Use list_files or search_code to locate files under output_directory, then use read_file/read_files to read them. Do not use read_code_item for unimported code."}
	};
}

} // namespace EPackagerTool
