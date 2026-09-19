"""离线检查固定 JSON 数据及其缓存：Unicode、并发首次初始化和副本隔离。"""

import json
from pathlib import Path
import re
import shutil
import subprocess

from benchmark_json_compile import ROOT, environment


def main():
    output = ROOT / ".temp/static-json-test"
    output.mkdir(parents=True, exist_ok=True)
    literals = []
    # 检查实际源文件中的固定文本，同时覆盖已经使用 parse 的 request_user_input schema。
    pattern = re.compile(
        r'(?:AUTOLINKER_STATIC_JSON|nlohmann::json::parse)\(\s*(R"([A-Za-z_0-9]*)\(([\s\S]*?)\)\2")')
    for path in sorted((ROOT / "src").glob("*")):
        if path.suffix not in (".cpp", ".h"):
            continue
        try:
            source = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            # 项目部分旧接口头仍是本地编码；不包含此次固定 JSON 定义。
            continue
        for match in pattern.finditer(source):
            value = json.loads(match[3])
            literals.append((str(path.relative_to(ROOT)), match[1], value))
    catalog_path = ROOT / "src/AIChatPublicToolCatalog.inc"
    catalog_source = catalog_path.read_text(encoding="utf-8-sig")
    match = re.search(r'(R"json\(([\s\S]*)\)json")', catalog_source)
    assert match, "missing public tool catalog literal"
    catalog = json.loads(match[2])
    names = [tool["name"] for tool in catalog]
    assert len(names) == len(set(names)) and len(names) >= 31
    assert all(tool["inputSchema"]["type"] == "object" for tool in catalog)
    literals.append((str(catalog_path.relative_to(ROOT)), match[1], catalog))

    code = r'''
#include "StaticJson.h"
#include <atomic>
#include <barrier>
#include <iostream>
#include <thread>
#include <type_traits>
#include <vector>

std::atomic_int parseCount = 0;
const nlohmann::json& Shared() {
    // 测试中故意加入可观测副作用，证明宏参数仅在首次初始化时求值。
    return AUTOLINKER_STATIC_JSON((++parseCount,
        R"({"items":[{"text":"\u4e2d\u6587\ud83d\ude00\u0000end"}],"object":{},"array":[],"null":null,"bool":false,"max":18446744073709551615})"));
}
static_assert(std::is_same_v<decltype(Shared()), const nlohmann::json&>);
bool CheckCache() {
    constexpr int count = 16;
    std::barrier start(count);
    std::vector<const nlohmann::json*> pointers(count);
    std::vector<std::thread> workers;
    for (int i = 0; i < count; ++i) {
        workers.emplace_back([&, i] { start.arrive_and_wait(); pointers[i] = &Shared(); });
    }
    for (auto& worker : workers) worker.join();
    if (parseCount != 1) return false;
    for (auto* pointer : pointers) if (pointer != pointers.front()) return false;
    nlohmann::json copy = Shared();
    copy["items"][0]["text"] = "changed";
    copy["array"].push_back(1);
    const auto& value = Shared();
    return copy != value && value["array"].is_array() && value["array"].empty() &&
        value["object"].is_object() && value["null"].is_null() && value["bool"] == false &&
        value["max"].get<std::uint64_t>() == UINT64_MAX &&
        value["items"][0]["text"].get<std::string>() == std::string("\xE4\xB8\xAD\xE6\x96\x87\xF0\x9F\x98\x80\0end", 14);
}
'''
    for index, (_, literal, _) in enumerate(literals):
        code += f"const nlohmann::json& Value{index}() {{ return AUTOLINKER_STATIC_JSON({literal}); }}\n"
    code += "int main() { if (!CheckCache()) return 1;\n"
    for index in range(len(literals)):
        code += f'std::cout << Value{index}().dump() << "\\n";\n'
    code += "}\n"
    cpp = output / "static_json_test.cpp"
    cpp.write_bytes(b"\xef\xbb\xbf" + code.replace("\n", "\r\n").encode("utf-8"))
    env = environment()
    executable = output / "static_json_test.exe"
    command = [shutil.which("cl", path=env["PATH"]), "/nologo", "/O2", "/MT", "/EHsc",
               "/std:c++20", "/utf-8", f"/I{ROOT / 'src'}", str(cpp),
               f"/Fo{output / 'static_json_test.obj'}", f"/Fe{executable}"]
    result = subprocess.run(command, cwd=output, env=env, capture_output=True)
    (output / "build.log").write_bytes(result.stdout + result.stderr)
    if result.returncode:
        raise RuntimeError(f"Compilation failed; see {output / 'build.log'}")
    values = subprocess.check_output([str(executable)]).splitlines()
    assert len(values) == len(literals)
    for data, (path, _, expected) in zip(values, literals):
        assert json.loads(data) == expected, path
    report = {"ok": True, "fixed_json_literals": len(literals), "public_tools": len(catalog),
              "concurrent_first_initialization": True, "mutation_isolated": True,
              "unicode_nul_and_integer_boundaries": True}
    (output / "result.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
