#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace e2txt {

// 依赖项类型。
enum class DependencyKind {
	ELib,
	ECom,
};

// 目录化资源类型。
enum class BundleResourceKind {
	Image,
	Sound,
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

// 目录化源码文件。
struct BundleSourceFile {
	std::string key;
	std::string logicalName;
	std::string relativePath;
	std::string content;
};

// 目录化窗口 XML 文件。
struct BundleFormFile {
	std::string key;
	std::string logicalName;
	std::string relativePath;
	std::string xmlText;
};

// 目录化二进制资源文件。
struct BundleBinaryResource {
	BundleResourceKind kind = BundleResourceKind::Image;
	std::string key;
	std::string logicalName;
	std::string relativePath;
	std::string comment;
	bool isPublic = false;
	std::vector<std::uint8_t> data;
};

// 目录化过滤器节点。
struct BundleFolder {
	int key = 0;
	int parentKey = 0;
	bool expand = true;
	std::string name;
	std::vector<std::string> childKeys;
};

// 窗口与窗口程序集的显式映射。
struct WindowBinding {
	std::string formName;
	std::string className;
};

// 原生方法快照。
struct BundleNativeMethodSnapshot {
	std::string name;
	std::string textDigest;
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::int32_t attr = 0;
	std::vector<std::int32_t> paramIds;
	std::vector<std::int32_t> localIds;
	std::vector<std::uint8_t> lineOffset;
	std::vector<std::uint8_t> blockOffset;
	std::vector<std::uint8_t> methodReference;
	std::vector<std::uint8_t> variableReference;
	std::vector<std::uint8_t> constantReference;
	std::vector<std::uint8_t> expressionData;
};

// 原生源码文件快照。
struct BundleNativeSourceFileSnapshot {
	std::string contentDigest;
	std::string classShapeDigest;
	std::int32_t classId = 0;
	std::int32_t classMemoryAddress = 0;
	std::int32_t formId = 0;
	std::int32_t baseClass = 0;
	std::vector<std::int32_t> classVarIds;
	std::vector<BundleNativeMethodSnapshot> methods;
};

// 原生全局变量快照。
struct BundleNativeGlobalSnapshot {
	std::string name;
	std::string textDigest;
	std::int32_t id = 0;
};

// 原生 DLL 命令快照。
struct BundleNativeDllSnapshot {
	std::string name;
	std::string textDigest;
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::vector<std::int32_t> paramIds;
};

// 原生自定义数据类型快照。
struct BundleNativeStructSnapshot {
	std::string name;
	std::string textDigest;
	std::int32_t id = 0;
	std::int32_t memoryAddress = 0;
	std::vector<std::int32_t> memberIds;
};

// 原生常量快照。
struct BundleNativeConstantSnapshot {
	std::string name;
	std::string key;
	std::string textDigest;
	std::int32_t id = 0;
	std::int32_t pageType = 0;
};

// 原生段快照。
struct NativeSectionSnapshot {
	std::uint32_t key = 0;
	std::string name;
	std::int32_t flags = 0;
	std::vector<std::uint8_t> data;
};

// 原生程序段头快照。
struct BundleNativeProgramHeaderSnapshot {
	std::int32_t versionFlag1 = 0;
	std::int32_t unk1 = 0;
	std::vector<std::uint8_t> unk2_1;
	std::vector<std::uint8_t> unk2_2;
	std::vector<std::uint8_t> unk2_3;
	std::vector<std::string> supportLibraryInfo;
	std::int32_t flag1 = 0;
	std::int32_t flag2 = 0;
	std::vector<std::uint8_t> unk3Op;
	std::vector<std::uint8_t> icon;
	std::string debugCommandLine;
};

// 目录化工程包。
struct ProjectBundle {
	std::string sourcePath;
	std::string projectName;
	bool projectNameStored = false;
	std::string versionText;
	std::int32_t bundleFormatVersion = 0;
	std::vector<Dependency> dependencies;
	std::vector<BundleSourceFile> sourceFiles;
	std::vector<BundleFormFile> formFiles;
	std::string dataTypeText;
	std::string dllDeclareText;
	std::string constantText;
	std::string globalText;
	std::vector<BundleBinaryResource> resources;
	std::int32_t folderAllocatedKey = 0;
	std::vector<BundleFolder> folders;
	std::vector<std::string> rootChildKeys;
	std::vector<WindowBinding> windowBindings;
	std::vector<BundleNativeSourceFileSnapshot> nativeSourceSnapshots;
	std::optional<BundleNativeProgramHeaderSnapshot> nativeProgramHeader;
	std::vector<BundleNativeGlobalSnapshot> nativeGlobalSnapshots;
	std::vector<BundleNativeStructSnapshot> nativeStructSnapshots;
	std::vector<BundleNativeDllSnapshot> nativeDllSnapshots;
	std::vector<BundleNativeConstantSnapshot> nativeConstantSnapshots;
	std::string nativeBundleDigest;
	std::vector<std::uint8_t> nativeSourceBytes;
};

// 计算目录工程包可见内容的稳定摘要。
std::string ComputeBundleDigest(const ProjectBundle& bundle);
// 计算文本内容的稳定摘要。
std::string ComputeTextDigest(const std::string& text);
// 从原生 .e 二进制提取段快照。
bool CaptureNativeSectionSnapshots(
	const std::vector<std::uint8_t>& inputBytes,
	std::vector<NativeSectionSnapshot>& outSnapshots,
	std::string* outError);

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
	// 直接生成目录化工程包。
	bool GenerateBundle(const std::string& inputPath, ProjectBundle& outBundle, std::string* outError) const;
	// 按内存中的 .e 二进制直接生成文档。
	bool GenerateDocumentFromBytes(
		const std::vector<std::uint8_t>& inputBytes,
		const std::string& sourcePath,
		Document& outDocument,
		std::string* outError) const;
	// 按内存中的 .e 二进制直接生成目录化工程包。
	bool GenerateBundleFromBytes(
		const std::vector<std::uint8_t>& inputBytes,
		const std::string& sourcePath,
		ProjectBundle& outBundle,
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
	// 从目录化工程包恢复为 .e 二进制。
	bool RestoreBundleToBytes(const ProjectBundle& bundle, std::vector<std::uint8_t>& outBytes, std::string* outError) const;
	bool RestoreToFile(
		const std::string& inputPath,
		const std::string& outputPath,
		std::string* outSummary,
		std::string* outError) const;
	// 从目录化工程包恢复为 .e 文件。
	bool RestoreBundleToFile(
		const ProjectBundle& bundle,
		const std::string& outputPath,
		std::string* outSummary,
		std::string* outError) const;
};

}  // namespace e2txt
