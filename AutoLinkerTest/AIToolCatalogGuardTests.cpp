// 无 IDE 的工具目录防御测试；覆盖重复、冲突、顺序及此次 150 项请求的结构。
#include "../src/AIToolCatalogGuard.h"
#include <iostream>
#include <stdexcept>

using Json = nlohmann::json;
int main(int argc, char** argv)
{
    int checks = 0;
    const auto check = [&checks](bool ok, const char* label) {
        if (!ok) throw std::runtime_error(label);
        ++checks;
    };
    const auto tool = [](const std::string& name) {
        return Json{{"name", name}, {"description", "test"}, {"inputSchema", {{"type", "object"}}}};
    };
    try {
        if (argc > 1) {
            Json request;
            std::cin >> request;
            if (std::string(argv[1]) == "--raw") {
                const auto& catalog = request.at("result").at("tools");
                const auto filtered = AIToolCatalogGuard::Filter(catalog);
                std::cout << "raw input=" << catalog.size() << " output=" << filtered.catalog.size()
                    << " duplicates=" << filtered.duplicateNames.size() << " conflicts="
                    << Json(filtered.conflictNames).dump() << '\n';
                check(!catalog.empty() && filtered.conflictNames.empty() && filtered.invalidCount == 0,
                    "raw IDA tools/list contains no ambiguous tools");
                check(!AIToolCatalogGuard::Filter(filtered.catalog).Changed(), "raw output names unique");
                std::cout << "PASS: current raw tools/list " << catalog.size() << " -> "
                    << filtered.catalog.size() << ", no conflicts\n";
                return 0;
            }
            Json catalog = Json::array();
            for (const auto& item : request.at("tools")) catalog.push_back(item.at("function"));
            const auto filtered = AIToolCatalogGuard::Filter(catalog);
            check(catalog.size() == 150 && filtered.catalog.size() == 92 &&
                filtered.duplicateNames.size() == 58 && filtered.conflictNames.empty(), "captured DeepSeek request");
            check(!AIToolCatalogGuard::Filter(filtered.catalog).Changed(), "captured names unique");
            std::cout << "PASS: captured request 150 -> 92, 58 duplicates removed, no conflicts\n";
            return 0;
        }
        Json a = tool("a"), b = tool("b"), c = tool("c");
        auto result = AIToolCatalogGuard::Filter(Json::array({a, b, a, c, b}));
        check(result.catalog == Json::array({a, b, c}), "exact duplicates and stable order");
        check(result.indices == std::vector<size_t>({0, 1, 3}), "registration mapping indices");
        check(result.duplicateNames.size() == 2 && result.conflictNames.empty(), "duplicate diagnostics");
        check(!AIToolCatalogGuard::Filter(result.catalog).Changed(), "idempotent filtering");
        Json changed = a;
        changed["inputSchema"]["properties"] = {{"new_field", {{"type", "string"}}}};
        result = AIToolCatalogGuard::Filter(Json::array({a, b, changed, a, c}));
        check(result.catalog == Json::array({b, c}), "reject all variants of conflicting schema");
        check(result.conflictNames == std::vector<std::string>({"a"}), "conflict diagnostics");
        check(result.indices == std::vector<size_t>({1, 4}), "no conflicting mapping survives");
        changed = a;
        changed["description"] = "different contract";
        check(AIToolCatalogGuard::Filter(Json::array({a, changed})).catalog.empty(), "description conflict");
        a["route_identity"] = "source_1";
        changed = a;
        changed["route_identity"] = "source_2";
        check(AIToolCatalogGuard::Filter(Json::array({a, changed})).catalog.empty(), "different route rejected");
        result = AIToolCatalogGuard::Filter(Json::array({nullptr, 2, Json::object(),
            {{"name", 7}}, {{"name", ""}}, {{"name", " \t"}}, b}));
        check(result.invalidCount == 6 && result.catalog == Json::array({b}), "invalid names");
        check(AIToolCatalogGuard::Filter(nullptr).invalidCount == 1, "invalid catalog");
        check(!AIToolCatalogGuard::Filter(Json::array()).Changed(), "empty catalog");
        Json catalog = Json::array();
        for (int i = 0; i < 31; ++i) catalog.push_back(tool("builtin_" + std::to_string(i)));
        for (int i = 0; i < 61; ++i) catalog.push_back(tool("mcp_ida_" + std::to_string(i)));
        for (int i = 0; i < 58; ++i) catalog.push_back(tool("mcp_ida_" + std::to_string(i)));
        result = AIToolCatalogGuard::Filter(catalog);
        check(catalog.size() == 150 && result.catalog.size() == 92 && result.duplicateNames.size() == 58,
            "DeepSeek incident structure: 150 to 92");
        Json first = {{"name", "same"}, {"inputSchema", {{"type", "object"}, {"required", Json::array()}}}};
        Json reordered = Json::parse("{\"inputSchema\":{\"required\":[],\"type\":\"object\"},\"name\":\"same\"}");
        result = AIToolCatalogGuard::Filter(Json::array({first, reordered}));
        check(result.catalog.size() == 1 && result.conflictNames.empty(), "structured equality not serialized order");
        std::cout << "PASS: " << checks << " tool catalog guard checks\n";
        return 0;
    } catch (const std::exception& error) {
        std::cerr << "FAIL: " << error.what() << '\n';
        return 1;
    }
}
