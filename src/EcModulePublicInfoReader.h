#pragma once

#include <string>

#include "native_module_public_info.hpp"

namespace e571 {

// 基于 EC 文件真实段结构的本地模块公开信息读取器。
class EcModulePublicInfoReader {
public:
	bool Load(const std::string& modulePath, ModulePublicInfoDump& outDump, std::string* outError) const;
};

}  // namespace e571
