"""比较 Win32 Release 下 JSON 的编译成本；实验产物仅写入 .temp/json-compile。"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import statistics
import subprocess
import time
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / ".temp/json-compile"
VERSION = "0.12.0"
HASHES = {
    "yyjson.h": "175867c5493a5df648cec566717fa1c29aa2f6096f5f0cf1efad0b65e1f6d7b3",
    "yyjson.c": "ac2e9bbb2e2d9149d90878d40506a1d624fa0b33c979a11b61075c54782c6d6a",
}


def write(name, text):
    path = OUT / name
    path.write_bytes(b"\xef\xbb\xbf" + text.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8"))
    return path


def environment():
    vswhere = Path(os.environ["ProgramFiles(x86)"]) / "Microsoft Visual Studio/Installer/vswhere.exe"
    installation = subprocess.check_output([
        str(vswhere), "-latest", "-products", "*", "-requires",
        "Microsoft.VisualStudio.Component.VC.Tools.x86.x64", "-property", "installationPath"
    ], text=True).strip()
    vcvars = Path(installation) / "VC/Auxiliary/Build/vcvarsall.bat"
    result = subprocess.run(f'call "{vcvars}" x64_x86 >nul && set', shell=True,
                            capture_output=True, check=True)
    env = os.environ.copy()
    for line in result.stdout.decode("mbcs").splitlines():
        if "=" in line and not line.startswith("="):
            key, value = line.split("=", 1)
            env[key.upper()] = value
    return env


NLOHMANN = r'''
#include "json.hpp"
#include <iostream>
#include <iterator>
#include <string>
int main() {
    std::string input(std::istreambuf_iterator<char>(std::cin), {});
    auto value = nlohmann::json::parse(input, nullptr, false);
    if (value.is_discarded()) return 2;
    nlohmann::json result = {{"ok", true}, {"payload", value}};
    std::cout << result.dump();
}
'''

YYJSON = r'''
#include "yyjson.h"
#include <cstdlib>
#include <iostream>
#include <iterator>
#include <memory>
#include <string>
int main() {
    std::string input(std::istreambuf_iterator<char>(std::cin), {});
    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(
        yyjson_read(input.data(), input.size(), 0), yyjson_doc_free);
    if (!doc) return 2;
    std::unique_ptr<yyjson_mut_doc, decltype(&yyjson_mut_doc_free)> output(
        yyjson_mut_doc_new(nullptr), yyjson_mut_doc_free);
    if (!output) return 3;
    auto* root = yyjson_mut_obj(output.get());
    auto* payload = yyjson_val_mut_copy(output.get(), yyjson_doc_get_root(doc.get()));
    if (!root || !payload || !yyjson_mut_obj_add_bool(output.get(), root, "ok", true) ||
        !yyjson_mut_obj_add_val(output.get(), root, "payload", payload)) return 3;
    yyjson_mut_doc_set_root(output.get(), root);
    size_t length = 0;
    std::unique_ptr<char, decltype(&std::free)> text(
        yyjson_mut_write(output.get(), 0, &length), std::free);
    if (!text) return 3;
    std::cout.write(text.get(), length);
}
'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    if args.rounds < 1:
        parser.error("--rounds must be positive")
    OUT.mkdir(parents=True, exist_ok=True)
    for name, digest in HASHES.items():
        path = OUT / name
        if not path.exists():
            path.write_bytes(urllib.request.urlopen(
                f"https://raw.githubusercontent.com/ibireme/yyjson/{VERSION}/src/{name}", timeout=60).read())
        if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
            raise RuntimeError(f"SHA256 mismatch: {path}")
    env = environment()
    records = []

    def run(label, command, measured=True):
        command = [shutil.which(command[0], path=env["PATH"]) or command[0], *command[1:]]
        start = time.perf_counter()
        result = subprocess.run(command, cwd=OUT, env=env, capture_output=True)
        seconds = time.perf_counter() - start
        (OUT / f"{label}.log").write_bytes(result.stdout + result.stderr)
        if result.returncode:
            raise RuntimeError(f"{label} failed; see {OUT / (label + '.log')}")
        if measured:
            records.append({"label": label, "seconds": seconds})
        print(f"{label}: {seconds:.3f}s", flush=True)

    common = ["cl", "/nologo", "/c", "/O2", "/MT", "/EHsc", "/std:c++20", "/utf-8",
              "/DNDEBUG", "/DWIN32", "/D_WINDOWS", "/DELIB_FNE_EXPORTS", "/Gy", "/Oi", "/Bt+",
              f"/I{ROOT / 'thirdparty'}", f"/I{ROOT / 'src'}", f"/I{OUT}"]
    write("empty.cpp", "int value() { return 1; }\n")
    write("include_nlohmann.cpp", '#include "json.hpp"\nint value() { return 1; }\n')
    write("include_yyjson.cpp", '// JSON 替代库的纯头文件编译成本。\n#include "yyjson.h"\nint value() { return 1; }\n')
    write("workload_nlohmann.cpp", NLOHMANN)
    write("workload_yyjson.cpp", YYJSON)
    write("pch.cpp", '#include "json.hpp"\n')
    original = (ROOT / "src/CompileConcurrencyManager.cpp").read_text(encoding="utf-8-sig")
    replacement = original.replace('#include "..\\\\thirdparty\\\\json.hpp"',
        '#include "yyjson.h"\n#include <cstdlib>\n#include <memory>\n#include <stdexcept>')
    start = replacement.index("\treturn nlohmann::json({")
    end = replacement.index("\n}", start)
    fields = {
        "ok": "ok", "process_reentry_blocked": "processReentryBlocked",
        "same_path_serialized": "samePathTimedOut && samePathAcquiredAfterRelease",
        "different_path_independent": "differentPathAcquired",
    }
    body = '''    std::unique_ptr<yyjson_mut_doc, decltype(&yyjson_mut_doc_free)> doc(
        yyjson_mut_doc_new(nullptr), yyjson_mut_doc_free);
    if (!doc) throw std::bad_alloc();
    auto* root = yyjson_mut_obj(doc.get());
    if (!root) throw std::bad_alloc();
    yyjson_mut_doc_set_root(doc.get(), root);
    if (!yyjson_mut_obj_add_str(doc.get(), root, "name", "compile-concurrency-manager")) throw std::bad_alloc();
'''
    for key, value in fields.items():
        body += f'    if (!yyjson_mut_obj_add_bool(doc.get(), root, "{key}", {value})) throw std::bad_alloc();\n'
    for key, value in [("first_error", "firstError"), ("same_path_error", "samePathError"),
                       ("other_error", "otherError"), ("retry_error", "retryError")]:
        body += f'    if (!yyjson_mut_obj_add_strn(doc.get(), root, "{key}", {value}.data(), {value}.size())) throw std::bad_alloc();\n'
    body += '''    size_t length = 0;
    std::unique_ptr<char, decltype(&std::free)> output(yyjson_mut_write(doc.get(), 0, &length), std::free);
    if (!output) throw std::runtime_error("JSON serialization failed");
    return std::string(output.get(), length);'''
    write("CompileConcurrencyManager_yyjson.cpp", replacement[:start] + body + replacement[end:])
    write("manager_main.cpp", '#include "CompileConcurrencyManager.h"\n#include <iostream>\n'
          'int main() { std::cout << CompileConcurrencyManager::BuildSelfTestJson(); }\n')

    cases = {name: OUT / (name + ".cpp") for name in [
        "empty", "include_nlohmann", "include_yyjson", "workload_nlohmann", "workload_yyjson"]}
    cases["manager_nlohmann"] = ROOT / "src/CompileConcurrencyManager.cpp"
    cases["manager_yyjson"] = OUT / "CompileConcurrencyManager_yyjson.cpp"
    for iteration in range(args.rounds):
        # 交替顺序，减轻固定顺序和文件缓存对结论的影响；不并发运行计时任务。
        items = list(cases.items())
        if iteration % 2:
            items.reverse()
        for label, source in items:
            run(f"{label}_{iteration}", common + [str(source), f"/Fo{OUT / (label + '.obj')}"])
        run(f"yyjson_library_{iteration}", ["cl", "/nologo", "/c", "/O2", "/MT", "/DNDEBUG", "/TC",
            "/Gy", "/Oi", "/Bt+", str(OUT / "yyjson.c"), f"/Fo{OUT / 'yyjson.obj'}"])

    pch = [f"/Fp{OUT / 'json.pch'}", "/FIjson.hpp"]
    run("pch_create", common + pch + ["/Ycjson.hpp", str(OUT / "pch.cpp"), f"/Fo{OUT / 'pch.obj'}"])
    for iteration in range(args.rounds):
        for label in ["workload_nlohmann", "manager_nlohmann"]:
            run(f"{label}_pch_{iteration}", common + pch + ["/Yujson.hpp", str(cases[label]),
                f"/Fo{OUT / (label + '_pch.obj')}"])

    run("manager_main", common + [str(OUT / "manager_main.cpp"), f"/Fo{OUT / 'manager_main.obj'}"], False)
    for label in ["workload_nlohmann", "workload_yyjson", "manager_nlohmann", "manager_yyjson"]:
        objects = [str(OUT / (label + ".obj"))]
        if label.startswith("manager_"):
            objects.append(str(OUT / "manager_main.obj"))
        if label.endswith("yyjson"):
            objects.append(str(OUT / "yyjson.obj"))
        run(label + "_link", ["link", "/nologo", "/MACHINE:X86", *objects,
            "kernel32.lib", "user32.lib", f"/OUT:{OUT / (label + '.exe')}"], False)

    fixtures = [None, True, [], {}, {"中文": "测试😀\x00尾部", "array": [1, False, None]},
                {"min": -(2**63), "max": 2**64 - 1, "float": 1.25}, {"escape": '\\"\n\t'}]
    for fixture in fixtures:
        payload = json.dumps(fixture, ensure_ascii=False).encode("utf-8")
        for label in ["workload_nlohmann", "workload_yyjson"]:
            output = subprocess.check_output([str(OUT / (label + ".exe"))], input=payload)
            assert json.loads(output) == {"ok": True, "payload": fixture}, label
    for payload in [b'{"bad":', b'{"x":NaN}', b'{"x":"\xff"}', b'{"x":1,}']:
        for label in ["workload_nlohmann", "workload_yyjson"]:
            result = subprocess.run([str(OUT / (label + ".exe"))], input=payload, capture_output=True)
            assert result.returncode == 2, (label, payload, result.returncode)
    reports = [json.loads(subprocess.check_output([str(OUT / (label + ".exe"))]))
               for label in ["manager_nlohmann", "manager_yyjson"]]
    assert reports[0] == reports[1] and reports[0]["ok"], reports
    medians = {}
    for record in records:
        parts = record["label"].rsplit("_", 1)
        key = parts[0] if parts[-1].isdigit() else record["label"]
        medians.setdefault(key, []).append(record["seconds"])
    result = {"yyjson_version": VERSION, "rounds": args.rounds, "records": records,
              "median_seconds": {key: statistics.median(values) for key, values in medians.items()},
              "validation": {"valid_fixtures": len(fixtures), "invalid_fixtures": 4,
                             "manager_reports_equal": True}, "manager_report": reports[0]}
    (OUT / "micro-results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result["median_seconds"], indent=2))
    print("JSON round trips, malformed input rejection and real manager self-tests passed.")


if __name__ == "__main__":
    main()
