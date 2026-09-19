"""隔离 AIService 的实际工具目录，比较嵌套初始化列表与 JSON 文本构造成本。"""

import json
import shutil
import statistics
import subprocess
import time

from benchmark_json_compile import ROOT, OUT, environment, write


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    env = environment()
    cl = shutil.which("cl", path=env["PATH"])
    source = (ROOT / "src/AIService.cpp").read_text(encoding="utf-8-sig")
    begin = source.index("nlohmann::json BuildPublicToolCatalog()\n{")
    end = source.index("\nbool IsRealPageReadToolVisible", begin)
    function = source[begin:end]
    # 独立样例统一使用 /utf-8，字符串已为 UTF-8，省略项目的本地编码转换。
    function = function.replace("\tNormalizeJsonStringsToUtf8InPlace(tools);\n", "")
    prefix = '#include "json.hpp"\n#include <iostream>\n'
    suffix = '\nint main() { std::cout << BuildPublicToolCatalog().dump(); }\n'
    write("catalog_initializer.cpp", prefix + function + suffix)
    common = [cl, "/nologo", "/O2", "/MT", "/EHsc", "/std:c++20", "/utf-8", "/DNDEBUG",
              "/Gy", "/Oi", "/Bt+", f"/I{ROOT / 'thirdparty'}"]
    records = []

    def compile_case(name, iteration):
        start = time.perf_counter()
        result = subprocess.run(common + [str(OUT / (name + ".cpp")),
            f"/Fo{OUT / (name + '.obj')}", f"/Fe{OUT / (name + '.exe')}"],
            cwd=OUT, env=env, capture_output=True)
        seconds = time.perf_counter() - start
        (OUT / f"{name}_{iteration}.log").write_bytes(result.stdout + result.stderr)
        if result.returncode:
            raise RuntimeError(f"{name} failed; see log")
        records.append({"case": name, "seconds": seconds})
        print(f"{name}_{iteration}: {seconds:.3f}s", flush=True)
        return subprocess.check_output([str(OUT / (name + ".exe"))])

    catalog = compile_case("catalog_initializer", 0)
    expected = json.loads(catalog)
    (OUT / "public-tool-catalog.json").write_bytes(catalog)
    write("catalog_text.cpp", prefix + 'nlohmann::json BuildPublicToolCatalog() {\n'
          'return nlohmann::json::parse(R"CATALOG(' + catalog.decode("utf-8") + ')CATALOG");\n}\n' + suffix)
    assert json.loads(compile_case("catalog_text", 0)) == expected
    for iteration in [1, 2]:
        order = ["catalog_text", "catalog_initializer"] if iteration == 1 else ["catalog_initializer", "catalog_text"]
        for name in order:
            assert json.loads(compile_case(name, iteration)) == expected
    result = {"records": records, "tool_count": len(expected), "all_outputs_equal": True,
              "median_seconds": {name: statistics.median(r["seconds"] for r in records if r["case"] == name)
                                 for name in ["catalog_initializer", "catalog_text"]}}
    (OUT / "catalog-results.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
