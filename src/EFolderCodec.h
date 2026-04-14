#pragma once

#include <string>

#include "e2txt.h"

namespace e2txt {

// 负责目录工程包的读写。
class BundleDirectoryCodec {
public:
	// 将工程包写入目录。
	bool WriteBundle(const ProjectBundle& bundle, const std::string& outputDir, std::string* outError) const;

	// 从目录读取工程包。
	bool ReadBundle(const std::string& inputDir, ProjectBundle& outBundle, std::string* outError) const;
};

}  // namespace e2txt
