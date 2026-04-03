#include "AutoLinkerTestApi.h"

#include <cstring>
#include <format>
#include <string>
#include <vector>

#include "AutoLinkerVersion.h"
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
