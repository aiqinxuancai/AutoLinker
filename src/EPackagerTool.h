#pragma once

#include <string>
#include "WorkspaceMirror.h"

#include "..\\thirdparty\\json.hpp"

// AI e_packager 工具：校验调用参数并把外部易语言源码解包到当前工程镜像。
namespace EPackagerTool {

// 将已由调度器完成的异步解包结果转为公开 JSON。
nlohmann::json BuildResult(const WorkspaceMirror::ReferenceUnpackResult& result, bool ok, const std::string& errorUtf8);

} // namespace EPackagerTool
