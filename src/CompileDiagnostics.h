#pragma once

// 编译诊断契约：仅在错误消息、原生行和源码行映射均可核实时公开位置。
#include <string>
#include "../thirdparty/json.hpp"

namespace CompileDiagnostics {
struct Location {
    std::string file;
    std::string page;
    std::string text;
    std::string mapping;
    int line = -1;
    int ideRow = -1;
    int ideColumn = -1;
    std::string pageCode;
    std::string textSource;
};
nlohmann::json Build(const std::string& output, const Location& location, bool compilationFailed);
// 从真实源码的程序集声明取得名称，避免编译跳转后使用过期页签标题。
std::string GetSourcePageName(const std::string& pageCode);
std::string BuildSelfTestJson();
}
