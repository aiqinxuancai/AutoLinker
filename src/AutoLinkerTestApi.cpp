#include "AutoLinkerTestApi.h"

#include <algorithm>
#include <cstring>
#include <format>
#include <string>
#include <vector>

#include "AutoLinkerVersion.h"
#include "EFolderCodec.h"
#include "EcModulePublicInfoReader.h"
#include "PathHelper.h"
#include "Version.h"
#include "e2txt.h"

namespace {

int CopyStringToBuffer(const std::string& value, char* buffer, int bufferSize)
{
	if (buffer == nullptr || bufferSize <= 0) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const size_t requiredSize = value.size() + 1;
	if (requiredSize > static_cast<size_t>(bufferSize)) {
		return AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL;
	}

	std::memcpy(buffer, value.c_str(), requiredSize);
	return static_cast<int>(value.size());
}

std::string BuildLocalModulePublicInfoDebugText(const e571::ModulePublicInfoDump& dump)
{
	std::string text;
	text += std::format(
		"path={}\r\nsource={}\r\ntrace={}\r\nrecords={}\r\nmodule={}\r\nassembly={}\r\nversion={}\r\n",
		dump.modulePath,
		dump.sourceKind,
		dump.trace,
		dump.records.size(),
		dump.moduleName,
		dump.assemblyName,
		dump.versionText);
	text += "\r\n[Formatted]\r\n";
	text += dump.formattedText;
	text += "\r\n\r\n[Records]\r\n";
	for (size_t i = 0; i < dump.records.size(); ++i) {
		const auto& record = dump.records[i];
		text += std::format(
			"#{} kind={} name={} type={} sig={}\r\n",
			i,
			record.kind,
			record.name,
			record.typeText,
			record.signatureText);
	}
	return text;
}

std::string BuildResourceDataDigest(const e2txt::BundleBinaryResource& resource)
{
	return e2txt::ComputeTextDigest(std::string(resource.data.begin(), resource.data.end()));
}

std::string BuildBundleDigestCompareText(const e2txt::ProjectBundle& fromE, const e2txt::ProjectBundle& fromDir)
{
	std::string text;
	const std::string digestFromE = e2txt::ComputeBundleDigest(fromE);
	const std::string digestFromDir = e2txt::ComputeBundleDigest(fromDir);
	text += std::format(
		"digest_from_e={}\r\ndigest_from_dir={}\r\nmatch={}\r\n",
		digestFromE,
		digestFromDir,
		digestFromE == digestFromDir ? "true" : "false");
	if (digestFromE == digestFromDir) {
		return text;
	}

	const auto appendValueMismatch = [&](const char* label, const std::string& left, const std::string& right) {
		text += std::format("mismatch={}\r\nleft={}\r\nright={}\r\n", label, left, right);
	};

	if (fromE.projectName != fromDir.projectName) {
		appendValueMismatch("projectName", fromE.projectName, fromDir.projectName);
		return text;
	}
	if (fromE.versionText != fromDir.versionText) {
		appendValueMismatch("versionText", fromE.versionText, fromDir.versionText);
		return text;
	}

	if (fromE.dependencies.size() != fromDir.dependencies.size()) {
		text += std::format("mismatch=dependencies.size\r\nleft={}\r\nright={}\r\n", fromE.dependencies.size(), fromDir.dependencies.size());
		return text;
	}
	for (size_t index = 0; index < fromE.dependencies.size(); ++index) {
		const auto& left = fromE.dependencies[index];
		const auto& right = fromDir.dependencies[index];
		if (left.kind != right.kind ||
			left.name != right.name ||
			left.fileName != right.fileName ||
			left.guid != right.guid ||
			left.versionText != right.versionText ||
			left.path != right.path ||
			left.reExport != right.reExport) {
			text += std::format(
				"mismatch=dependencies[{}]\r\nleft_kind={}\r\nright_kind={}\r\nleft_name={}\r\nright_name={}\r\nleft_file={}\r\nright_file={}\r\nleft_guid={}\r\nright_guid={}\r\nleft_version={}\r\nright_version={}\r\nleft_path={}\r\nright_path={}\r\nleft_reExport={}\r\nright_reExport={}\r\n",
				index,
				static_cast<int>(left.kind),
				static_cast<int>(right.kind),
				left.name,
				right.name,
				left.fileName,
				right.fileName,
				left.guid,
				right.guid,
				left.versionText,
				right.versionText,
				left.path,
				right.path,
				left.reExport ? 1 : 0,
				right.reExport ? 1 : 0);
			return text;
		}
	}

	if (fromE.sourceFiles.size() != fromDir.sourceFiles.size()) {
		text += std::format("mismatch=sourceFiles.size\r\nleft={}\r\nright={}\r\n", fromE.sourceFiles.size(), fromDir.sourceFiles.size());
		return text;
	}
	for (size_t index = 0; index < fromE.sourceFiles.size(); ++index) {
		const auto& left = fromE.sourceFiles[index];
		const auto& right = fromDir.sourceFiles[index];
		if (left.key != right.key ||
			left.logicalName != right.logicalName ||
			left.relativePath != right.relativePath ||
			left.content != right.content) {
			text += std::format(
				"mismatch=sourceFiles[{}]\r\nleft_key={}\r\nright_key={}\r\nleft_name={}\r\nright_name={}\r\nleft_relative={}\r\nright_relative={}\r\nleft_digest={}\r\nright_digest={}\r\n",
				index,
				left.key,
				right.key,
				left.logicalName,
				right.logicalName,
				left.relativePath,
				right.relativePath,
				e2txt::ComputeTextDigest(left.content),
				e2txt::ComputeTextDigest(right.content));
			return text;
		}
	}

	if (fromE.formFiles.size() != fromDir.formFiles.size()) {
		text += std::format("mismatch=formFiles.size\r\nleft={}\r\nright={}\r\n", fromE.formFiles.size(), fromDir.formFiles.size());
		return text;
	}
	for (size_t index = 0; index < fromE.formFiles.size(); ++index) {
		const auto& left = fromE.formFiles[index];
		const auto& right = fromDir.formFiles[index];
		if (left.key != right.key ||
			left.logicalName != right.logicalName ||
			left.relativePath != right.relativePath ||
			left.xmlText != right.xmlText) {
			text += std::format(
				"mismatch=formFiles[{}]\r\nleft_key={}\r\nright_key={}\r\nleft_name={}\r\nright_name={}\r\nleft_relative={}\r\nright_relative={}\r\nleft_digest={}\r\nright_digest={}\r\n",
				index,
				left.key,
				right.key,
				left.logicalName,
				right.logicalName,
				left.relativePath,
				right.relativePath,
				e2txt::ComputeTextDigest(left.xmlText),
				e2txt::ComputeTextDigest(right.xmlText));
			return text;
		}
	}

	if (fromE.dataTypeText != fromDir.dataTypeText) {
		appendValueMismatch("dataTypeText.digest", e2txt::ComputeTextDigest(fromE.dataTypeText), e2txt::ComputeTextDigest(fromDir.dataTypeText));
		return text;
	}
	if (fromE.dllDeclareText != fromDir.dllDeclareText) {
		appendValueMismatch("dllDeclareText.digest", e2txt::ComputeTextDigest(fromE.dllDeclareText), e2txt::ComputeTextDigest(fromDir.dllDeclareText));
		return text;
	}
	if (fromE.constantText != fromDir.constantText) {
		appendValueMismatch("constantText.digest", e2txt::ComputeTextDigest(fromE.constantText), e2txt::ComputeTextDigest(fromDir.constantText));
		return text;
	}
	if (fromE.globalText != fromDir.globalText) {
		appendValueMismatch("globalText.digest", e2txt::ComputeTextDigest(fromE.globalText), e2txt::ComputeTextDigest(fromDir.globalText));
		return text;
	}

	if (fromE.resources.size() != fromDir.resources.size()) {
		text += std::format("mismatch=resources.size\r\nleft={}\r\nright={}\r\n", fromE.resources.size(), fromDir.resources.size());
		return text;
	}
	for (size_t index = 0; index < fromE.resources.size(); ++index) {
		const auto& left = fromE.resources[index];
		const auto& right = fromDir.resources[index];
		if (left.kind != right.kind ||
			left.key != right.key ||
			left.logicalName != right.logicalName ||
			left.relativePath != right.relativePath ||
			left.comment != right.comment ||
			left.isPublic != right.isPublic ||
			left.data != right.data) {
			text += std::format(
				"mismatch=resources[{}]\r\nleft_kind={}\r\nright_kind={}\r\nleft_key={}\r\nright_key={}\r\nleft_name={}\r\nright_name={}\r\nleft_relative={}\r\nright_relative={}\r\nleft_comment={}\r\nright_comment={}\r\nleft_public={}\r\nright_public={}\r\nleft_size={}\r\nright_size={}\r\nleft_digest={}\r\nright_digest={}\r\n",
				index,
				static_cast<int>(left.kind),
				static_cast<int>(right.kind),
				left.key,
				right.key,
				left.logicalName,
				right.logicalName,
				left.relativePath,
				right.relativePath,
				left.comment,
				right.comment,
				left.isPublic ? 1 : 0,
				right.isPublic ? 1 : 0,
				left.data.size(),
				right.data.size(),
				BuildResourceDataDigest(left),
				BuildResourceDataDigest(right));
			return text;
		}
	}

	if (fromE.folderAllocatedKey != fromDir.folderAllocatedKey) {
		text += std::format("mismatch=folderAllocatedKey\r\nleft={}\r\nright={}\r\n", fromE.folderAllocatedKey, fromDir.folderAllocatedKey);
		return text;
	}
	if (fromE.rootChildKeys != fromDir.rootChildKeys) {
		text += "mismatch=rootChildKeys\r\n";
		text += std::format("left_count={}\r\nright_count={}\r\n", fromE.rootChildKeys.size(), fromDir.rootChildKeys.size());
		for (size_t index = 0; index < (std::min)(fromE.rootChildKeys.size(), fromDir.rootChildKeys.size()); ++index) {
			if (fromE.rootChildKeys[index] != fromDir.rootChildKeys[index]) {
				text += std::format("first_diff_index={}\r\nleft={}\r\nright={}\r\n", index, fromE.rootChildKeys[index], fromDir.rootChildKeys[index]);
				return text;
			}
		}
		return text;
	}

	if (fromE.folders.size() != fromDir.folders.size()) {
		text += std::format("mismatch=folders.size\r\nleft={}\r\nright={}\r\n", fromE.folders.size(), fromDir.folders.size());
		return text;
	}
	for (size_t index = 0; index < fromE.folders.size(); ++index) {
		const auto& left = fromE.folders[index];
		const auto& right = fromDir.folders[index];
		if (left.key != right.key ||
			left.parentKey != right.parentKey ||
			left.expand != right.expand ||
			left.name != right.name ||
			left.childKeys != right.childKeys) {
			text += std::format(
				"mismatch=folders[{}]\r\nleft_key={}\r\nright_key={}\r\nleft_parent={}\r\nright_parent={}\r\nleft_expand={}\r\nright_expand={}\r\nleft_name={}\r\nright_name={}\r\nleft_child_count={}\r\nright_child_count={}\r\n",
				index,
				left.key,
				right.key,
				left.parentKey,
				right.parentKey,
				left.expand ? 1 : 0,
				right.expand ? 1 : 0,
				left.name,
				right.name,
				left.childKeys.size(),
				right.childKeys.size());
			return text;
		}
	}

	if (fromE.windowBindings.size() != fromDir.windowBindings.size()) {
		text += std::format("mismatch=windowBindings.size\r\nleft={}\r\nright={}\r\n", fromE.windowBindings.size(), fromDir.windowBindings.size());
		return text;
	}
	for (size_t index = 0; index < fromE.windowBindings.size(); ++index) {
		const auto& left = fromE.windowBindings[index];
		const auto& right = fromDir.windowBindings[index];
		if (left.formName != right.formName || left.className != right.className) {
			text += std::format(
				"mismatch=windowBindings[{}]\r\nleft_form={}\r\nright_form={}\r\nleft_class={}\r\nright_class={}\r\n",
				index,
				left.formName,
				right.formName,
				left.className,
				right.className);
			return text;
		}
	}

	text += "mismatch=unknown\r\n";
	return text;
}

}

extern "C" bool AutoLinkerTest_CompareVersion(const char* left, const char* right, int* outResult)
{
	if (left == nullptr || right == nullptr || outResult == nullptr) {
		return false;
	}

	const Version leftVersion(left);
	const Version rightVersion(right);
	if (leftVersion < rightVersion) {
		*outResult = -1;
	}
	else if (leftVersion > rightVersion) {
		*outResult = 1;
	}
	else {
		*outResult = 0;
	}

	return true;
}

extern "C" int AutoLinkerTest_GetLinkerOutFileName(const char* commandLine, char* buffer, int bufferSize)
{
	if (commandLine == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(GetLinkerCommandOutFileName(commandLine), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GetLinkerKrnlnFileName(const char* commandLine, char* buffer, int bufferSize)
{
	if (commandLine == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(GetLinkerCommandKrnlnFileName(commandLine), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_ExtractBetweenDashes(const char* text, char* buffer, int bufferSize)
{
	if (text == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(ExtractBetweenDashes(text), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GetVersionText(char* buffer, int bufferSize)
{
	return CopyStringToBuffer(AUTOLINKER_VERSION, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_DumpLocalModulePublicInfo(const char* modulePath, char* buffer, int bufferSize)
{
	if (modulePath == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	e571::EcModulePublicInfoReader reader;
	e571::ModulePublicInfoDump dump;
	std::string error;
	if (!reader.Load(modulePath, dump, &error)) {
		return CopyStringToBuffer("load_failed: " + error, buffer, bufferSize);
	}

	return CopyStringToBuffer(BuildLocalModulePublicInfoDebugText(dump), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GenerateE2Txt(const char* inputPath, const char* outputPath, char* buffer, int bufferSize)
{
	if (inputPath == nullptr || outputPath == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	e2txt::Generator generator;
	std::string summary;
	std::string error;
	if (!generator.GenerateToFile(inputPath, outputPath, &summary, &error)) {
		return CopyStringToBuffer("generate_failed: " + error, buffer, bufferSize);
	}

	return CopyStringToBuffer(summary, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RestoreE2Txt(const char* inputPath, const char* outputPath, char* buffer, int bufferSize)
{
	if (inputPath == nullptr || outputPath == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	e2txt::Restorer restorer;
	std::string summary;
	std::string error;
	if (!restorer.RestoreToFile(inputPath, outputPath, &summary, &error)) {
		return CopyStringToBuffer("restore_failed: " + error, buffer, bufferSize);
	}

	return CopyStringToBuffer(summary, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_UnpackEProject(const char* inputPath, const char* outputDir, char* buffer, int bufferSize)
{
	if (inputPath == nullptr || outputDir == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	e2txt::Generator generator;
	e2txt::BundleDirectoryCodec codec;
	e2txt::ProjectBundle bundle;
	std::string error;
	if (!generator.GenerateBundle(inputPath, bundle, &error)) {
		return CopyStringToBuffer("unpack_generate_failed: " + error, buffer, bufferSize);
	}
	if (!codec.WriteBundle(bundle, outputDir, &error)) {
		return CopyStringToBuffer("unpack_write_failed: " + error, buffer, bufferSize);
	}

	const std::string summary =
		"source_files=" + std::to_string(bundle.sourceFiles.size()) +
		", form_files=" + std::to_string(bundle.formFiles.size()) +
		", resources=" + std::to_string(bundle.resources.size()) +
		", output=" + std::string(outputDir);
	return CopyStringToBuffer(summary, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_PackEProject(const char* inputDir, const char* outputPath, char* buffer, int bufferSize)
{
	if (inputDir == nullptr || outputPath == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	e2txt::BundleDirectoryCodec codec;
	e2txt::ProjectBundle bundle;
	std::string error;
	if (!codec.ReadBundle(inputDir, bundle, &error)) {
		return CopyStringToBuffer("pack_read_failed: " + error, buffer, bufferSize);
	}

	e2txt::Restorer restorer;
	std::string summary;
	if (!restorer.RestoreBundleToFile(bundle, outputPath, &summary, &error)) {
		return CopyStringToBuffer("pack_restore_failed: " + error, buffer, bufferSize);
	}
	return CopyStringToBuffer(summary, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_CompareBundleDigest(const char* inputPath, const char* inputDir, char* buffer, int bufferSize)
{
	if (inputPath == nullptr || inputDir == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	e2txt::Generator generator;
	e2txt::BundleDirectoryCodec codec;
	e2txt::ProjectBundle bundleFromE;
	e2txt::ProjectBundle bundleFromDir;
	std::string error;
	if (!generator.GenerateBundle(inputPath, bundleFromE, &error)) {
		return CopyStringToBuffer("generate_bundle_failed: " + error, buffer, bufferSize);
	}
	if (!codec.ReadBundle(inputDir, bundleFromDir, &error)) {
		return CopyStringToBuffer("read_bundle_failed: " + error, buffer, bufferSize);
	}
	return CopyStringToBuffer(BuildBundleDigestCompareText(bundleFromE, bundleFromDir), buffer, bufferSize);
}
