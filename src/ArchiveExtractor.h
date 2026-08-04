#pragma once

#include <filesystem>
#include <string>

// ZIP 归档解压器：优先使用 PowerShell，失败时回退到系统 tar.exe。
namespace ArchiveExtractor {

enum class ExtractionMethod {
	None,
	PowerShell,
	Tar
};

struct ExtractionResult {
	bool ok = false;
	ExtractionMethod method = ExtractionMethod::None;
	std::string primaryError;
	std::string fallbackError;
	std::string error;
};

// 将 ZIP 解压到已创建或可创建的目标目录，并返回实际使用的解压方式。
bool ExtractZip(
	const std::filesystem::path& archivePath,
	const std::filesystem::path& destination,
	const std::filesystem::path& workingDirectory,
	ExtractionResult& outResult,
	int timeoutSeconds = 120);

const char* MethodName(ExtractionMethod method);

} // namespace ArchiveExtractor
