#pragma once

#include <string>

#include "..\\thirdparty\\json.hpp"

// AI e_packager 工具：校验调用参数并把外部易语言源码解包到当前工程镜像。
namespace EPackagerTool {

// 执行工具调用，结果始终返回可直接序列化的结构化 JSON。
nlohmann::json Execute(const std::string& argumentsJsonUtf8, bool& outOk);

} // namespace EPackagerTool
