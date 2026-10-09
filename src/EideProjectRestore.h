#pragma once
#include <array>
#include <string>
// e5.95 专用二进制内容恢复；身份在事务开始时捕获，回滚复用同一份。
namespace EideProjectRestore {
struct Identity {
    std::array<std::string, 3> paths;
};
bool IsSupported();
bool CaptureIdentity(Identity& identity, std::string& error);
// 仅在签名、文档状态和路径均验证后返回全局工程序列化对象，不依赖活动代码页。
void* CurrentSerializerObject();
bool Reload(const std::string& binaryPath, const Identity& identity, std::string& error);
}
