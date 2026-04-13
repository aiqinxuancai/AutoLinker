#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace e2txt {

// 依赖项类型。
enum class DependencyKind {
	ELib,
	ECom,
};

// 依赖项信息。
struct Dependency {
	DependencyKind kind = DependencyKind::ELib;
	std::string name;
	std::string fileName;
	std::string guid;
	std::string versionText;
	std::string path;
	bool reExport = false;
};

// 导出页面信息。
struct Page {
	std::string typeName;
	std::string name;
	std::vector<std::string> lines;
};

// 导出的窗口 XML 信息。
struct FormXml {
	std::string name;
	std::vector<std::string> lines;
};

// e2txt 文档对象。
struct Document {
	std::string sourcePath;
	std::string outputPath;
	std::string projectName;
	std::string versionText;
	std::vector<Dependency> dependencies;
	std::vector<Page> pages;
	std::vector<FormXml> formXmls;
};

// 导出选项。
struct GenerateOptions {
	bool includeImportedPages = false;
};

// 原生 e2txt 生成器。
class Generator {
public:
	bool GenerateDocument(const std::string& inputPath, Document& outDocument, std::string* outError) const;
	// 按内存中的 .e 二进制直接生成文档。
	bool GenerateDocumentFromBytes(
		const std::vector<std::uint8_t>& inputBytes,
		const std::string& sourcePath,
		Document& outDocument,
		std::string* outError) const;
	bool GenerateText(const Document& document, std::string& outText, std::string* outError) const;
	bool GenerateToFile(
		const std::string& inputPath,
		const std::string& outputPath,
		std::string* outSummary,
		std::string* outError,
		const GenerateOptions& options = {}) const;

private:
	bool GenerateDocumentInternal(
		const std::string& inputPath,
		const GenerateOptions& options,
		Document& outDocument,
		std::string* outError) const;
};

// 原生 txt2e 恢复器。
class Restorer {
public:
	bool ParseText(const std::string& inputPath, Document& outDocument, std::string* outError) const;
	bool RestoreToBytes(const Document& document, std::vector<std::uint8_t>& outBytes, std::string* outError) const;
	bool RestoreToFile(
		const std::string& inputPath,
		const std::string& outputPath,
		std::string* outSummary,
		std::string* outError) const;
};

}  // namespace e2txt
