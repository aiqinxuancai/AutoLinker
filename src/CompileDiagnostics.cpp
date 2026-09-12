#include "CompileDiagnostics.h"
#include <sstream>
#include <algorithm>
#include <vector>

namespace CompileDiagnostics {
namespace {
bool MentionsSourceSymbol(const std::string& message, const std::string& text)
{
    // 当前光标可能仍停留在旧行。仅接受本次错误明确引用且在原生行中出现的符号。
    const auto begin = message.find("“");
    const auto end = begin == std::string::npos ? std::string::npos : message.find("”", begin + 3);
    return end != std::string::npos && end > begin + 3 &&
        text.find(message.substr(begin + 3, end - begin - 3)) != std::string::npos;
}
}
nlohmann::json Build(const std::string& output, const Location& location, bool compilationFailed)
{
    auto result = nlohmann::json::array();
    std::istringstream stream(output);
    std::string line;
    while (std::getline(stream, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        const auto start = line.find_first_not_of(" \t");
        if (start == std::string::npos) continue;
        line.erase(0, start);
        const bool error = line.rfind("错误(", 0) == 0 || line.rfind("错误（", 0) == 0 ||
            line.rfind("语法错误", 0) == 0 || line.rfind("编译错误", 0) == 0;
        const bool warning = line.rfind("警告(", 0) == 0 || line.rfind("警告（", 0) == 0;
        if (!error && !warning) continue;
        result.push_back({{"severity", error ? "error" : "warning"}, {"message", line},
            {"file", nullptr}, {"line", nullptr}, {"location_verified", false},
            {"location_status", "compiler_did_not_provide_verified_source_location"}});
    }
    // 一个当前光标只能代表一个错误，不能将同一行附会到整个错误数组。
    if (compilationFailed && result.size() == 1 && result[0]["severity"] == "error" &&
        location.line > 0 && !location.file.empty() &&
        (location.mapping == "native_range_prefix_verified" || location.mapping == "native_row_unique_text_verified") &&
        MentionsSourceSymbol(result[0]["message"].get<std::string>(), location.text)) {
        auto& diagnostic = result[0];
        diagnostic["file"] = location.file;
        diagnostic["line"] = location.line;
        diagnostic["line_basis"] = "read_real_file_1_based";
        diagnostic["page_name"] = location.page;
        diagnostic["code"] = location.text;
        diagnostic["ide_row"] = location.ideRow;
        diagnostic["ide_column"] = location.ideColumn;
        diagnostic["ide_position_basis"] = "zero_based";
        diagnostic["location_verified"] = true;
        diagnostic["location_status"] = location.mapping;
        diagnostic["code_excerpt"] = std::to_string(location.line) + "\t" + location.text;
        diagnostic["code_source"] = location.textSource;
        std::vector<std::string> sourceLines;
        std::istringstream page(location.pageCode);
        std::string sourceLine;
        while (std::getline(page, sourceLine)) {
            if (!sourceLine.empty() && sourceLine.back() == '\r') sourceLine.pop_back();
            sourceLines.push_back(sourceLine);
        }
        if (location.line <= static_cast<int>(sourceLines.size())) {
            const int first = (std::max)(1, location.line - 2);
            const int last = (std::min)(static_cast<int>(sourceLines.size()), location.line + 2);
            std::string excerpt;
            for (int number = first; number <= last; ++number) {
                excerpt += (number == location.line ? "> " : "  ") + std::to_string(number) +
                    "\t" + sourceLines[number - 1] + "\n";
            }
            diagnostic["code_excerpt"] = excerpt;
            diagnostic["excerpt_start_line"] = first;
            diagnostic["excerpt_end_line"] = last;
            diagnostic["focus_line"] = location.line;
        }
    }
    return result;
}
std::string GetSourcePageName(const std::string& pageCode)
{
    std::istringstream stream(pageCode);
    std::string line;
    const std::string marker = ".程序集";
    while (std::getline(stream, line)) {
        const auto first = line.find_first_not_of(" \t");
        if (first == std::string::npos || line.compare(first, marker.size(), marker) != 0) continue;
        const auto afterMarker = first + marker.size();
        if (afterMarker == line.size() || (line[afterMarker] != ' ' && line[afterMarker] != '\t')) continue;
        const auto begin = line.find_first_not_of(" \t", afterMarker);
        if (begin == std::string::npos) return {};
        const auto end = line.find_first_of(",\r\n\t", begin);
        std::string name = line.substr(begin, end == std::string::npos ? end : end - begin);
        const auto last = name.find_last_not_of(' ');
        return last == std::string::npos ? std::string() : name.substr(0, last + 1);
    }
    return {};
}

std::string BuildSelfTestJson()
{
    Location location{"src/测试.txt", "测试", "输出行 (未定义变量)", "native_range_prefix_verified", 7, 4, 0};
    const std::string error = "错误(30): 找不到指定的变量名称“未定义变量”。\r\n";
    const auto valid = Build(error, location, true);
    const bool mapped = valid.size() == 1 && valid[0].value("location_verified", false) && valid[0]["line"] == 7;
    const bool stale = !Build("错误(30): 找不到指定的变量名称“别的变量”。", location, true)[0].value("location_verified", false);
    const bool multiple = !Build(error + error, location, true)[0].value("location_verified", false);
    const bool success = !Build(error, location, false)[0].value("location_verified", false);
    location.line = -1;
    const bool unmapped = !Build(error, location, true)[0].value("location_verified", false);
    const bool sourceName = GetSourcePageName(".版本 2\r\n\r\n.程序集 第二页, , , 基类\r\n") == "第二页" &&
        GetSourcePageName(".程序集变量 变量, 整数型\r\n").empty();
    return nlohmann::json({{"ok", mapped && stale && multiple && success && unmapped && sourceName},
        {"source_page_name", sourceName},
        {"name", "compile-diagnostics"}, {"mapped", mapped}, {"stale_rejected", stale},
        {"multiple_rejected", multiple}, {"success_rejected", success}, {"unmapped_rejected", unmapped}}).dump();
}
}
