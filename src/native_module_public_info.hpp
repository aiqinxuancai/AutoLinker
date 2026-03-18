#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace e571 {

struct ModulePublicInfoParam {
	std::string name;
	std::string typeText;
	std::string flagsText;
	std::string comment;
};

struct ModulePublicInfoRecord {
	int tag = 0;
	unsigned int bodySize = 0;
	std::vector<int> headerInts;
	int payloadOffset = 0;
	std::vector<unsigned char> rawBytes;
	std::vector<std::string> extractedStrings;
	std::string kind;
	std::string name;
	std::string typeText;
	std::string flagsText;
	std::string comment;
	std::string signatureText;
	std::vector<ModulePublicInfoParam> params;
};

struct ModulePublicInfoDump {
	std::string modulePath;
	std::string loaderError;
	std::string trace;
	std::string sourceKind;
	std::string versionText;
	std::string moduleName;
	std::string assemblyName;
	std::string assemblyComment;
	std::string formattedText;
	int nativeResult = -1;
	std::vector<ModulePublicInfoRecord> records;
};

bool LoadModulePublicInfoDump(
	const std::string& modulePath,
	std::uintptr_t moduleBase,
	ModulePublicInfoDump* outDump,
	std::string* outError);

}  // namespace e571
