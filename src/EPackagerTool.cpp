#include "EPackagerTool.h"

#include <exception>
#include <string>

#include "WorkspaceMirror.h"

namespace EPackagerTool {

nlohmann::json BuildResult(const WorkspaceMirror::ReferenceUnpackResult& unpackResult, bool ok, const std::string& errorUtf8)
{
    if (!ok) return {{"ok", false}, {"operation", "unpack"}, {"error", errorUtf8}};

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
