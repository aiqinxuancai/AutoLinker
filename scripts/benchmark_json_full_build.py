"""对当前项目进行基线/PCH 对照重编译，不修改生产项目配置。"""

import argparse
import json
from pathlib import Path
import subprocess
import time
import xml.etree.ElementTree as ET

from benchmark_json_compile import ROOT, OUT, environment, write


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mode", choices=["baseline", "pch"], required=True)
    parser.add_argument("--label", required=True)
    parser.add_argument("--catalog-text", action="store_true",
                        help="使用 benchmark_json_catalog.py 导出的工具定义替换临时 AIService 副本")
    args = parser.parse_args()
    if not args.label.replace("-", "").replace("_", "").isalnum():
        parser.error("label may contain only letters, numbers, '-' and '_'")
    OUT.mkdir(parents=True, exist_ok=True)
    env = environment()
    installation = Path(env["VSINSTALLDIR"])
    msbuild = installation / "MSBuild/Current/Bin/amd64/MSBuild.exe"
    ns = {"m": "http://schemas.microsoft.com/developer/msbuild/2003"}
    project = ET.parse(ROOT / "AutoLinker.vcxproj")
    nodes = project.findall("m:ItemGroup/m:ClCompile", ns)
    # 使用已成功基线编译的依赖记录，避免靠字符串搜索遗漏间接包含 JSON 的文件。
    dependency_file = ROOT / "bin/intermediate/Win32/fne_release/AutoLinker.tlog/CL.read.1.tlog"
    dependencies = dependency_file.read_text(encoding="utf-16")
    selected = set()
    for block in dependencies.split("^")[1:]:
        lines = block.splitlines()
        if any(line.upper().endswith("\\THIRDPARTY\\JSON.HPP") or
               ("\\.TEMP\\JSON-COMPILE\\JSON_PCH_" in line.upper() and line.upper().endswith(".PCH"))
               for line in lines[1:]):
            selected.update(line.upper() for line in lines[0].split("|"))
    if str(OUT / "AIService_catalog.cpp").upper() in selected:
        selected.add(str(ROOT / "src/AIService.cpp").upper())
    target = ET.Element("Project", xmlns=ns["m"])
    definition = ET.SubElement(ET.SubElement(target, "ItemDefinitionGroup"), "ClCompile")
    ET.SubElement(definition, "AdditionalOptions").text = "/Bt+ %(AdditionalOptions)"
    groups = {}
    selected_names = []
    items = ET.SubElement(target, "ItemGroup")
    catalog_source = None
    if args.catalog_text:
        catalog = json.loads((OUT / "public-tool-catalog.json").read_bytes())
        source = (ROOT / "src/AIService.cpp").read_text(encoding="utf-8-sig")
        begin = source.index("nlohmann::json BuildPublicToolCatalog()\n{")
        end = source.index("\nbool IsRealPageReadToolVisible", begin)
        # ASCII 转义避免改变原文件的执行字符集；运行时解析还原原始 UTF-8 字符串。
        replacement = ('nlohmann::json BuildPublicToolCatalog() {\n'
                       'return nlohmann::json::parse(R"CATALOG(' +
                       json.dumps(catalog, ensure_ascii=True, separators=(",", ":")) + ')CATALOG");\n}\n')
        catalog_source = write("AIService_catalog.cpp", source[:begin] + replacement + source[end:])
        ET.SubElement(items, "ClCompile", Remove="src\\AIService.cpp")
        item = ET.SubElement(items, "ClCompile", Include=str(catalog_source))
        ET.SubElement(item, "AdditionalIncludeDirectories").text = str(ROOT / "src") + ";%(AdditionalIncludeDirectories)"
    for node in nodes:
        relative = node.attrib["Include"]
        if str((ROOT / relative).resolve()).upper() not in selected:
            continue
        selected_names.append(relative)
        if args.mode != "pch":
            continue
        options = node.findtext("m:AdditionalOptions", "", ns)
        charset = "/utf-8" if "/utf-8" in options else (
            "/execution-charset:utf-8" if "/execution-charset:utf-8" in options else "")
        group = groups.setdefault(charset, f"json_pch_{len(groups)}")
        compiled_path = str(catalog_source) if catalog_source and relative == "src\\AIService.cpp" else relative
        item = ET.SubElement(items, "ClCompile", Update=compiled_path)
        ET.SubElement(item, "PrecompiledHeader").text = "Use"
        ET.SubElement(item, "PrecompiledHeaderFile").text = "json.hpp"
        ET.SubElement(item, "PrecompiledHeaderOutputFile").text = str(OUT / (group + ".pch"))
        ET.SubElement(item, "ForcedIncludeFiles").text = "json.hpp;%(ForcedIncludeFiles)"
    for charset, group in groups.items():
        source = write(group + ".cpp", '#include "json.hpp"\n')
        item = ET.SubElement(items, "ClCompile", Include=str(source))
        ET.SubElement(item, "PrecompiledHeader").text = "Create"
        ET.SubElement(item, "PrecompiledHeaderFile").text = "json.hpp"
        ET.SubElement(item, "PrecompiledHeaderOutputFile").text = str(OUT / (group + ".pch"))
        ET.SubElement(item, "ForcedIncludeFiles").text = "json.hpp;%(ForcedIncludeFiles)"
        ET.SubElement(item, "AdditionalOptions").text = charset + " %(AdditionalOptions)"
    target_file = write(args.label + ".targets", ET.tostring(target, encoding="unicode"))
    (OUT / "json-dependent-files.json").write_text(json.dumps(selected_names, indent=2), encoding="utf-8")
    if not selected_names:
        raise RuntimeError("No JSON dependencies found; first run an ordinary fne_release/Win32 Build")
    command = [str(msbuild), str(ROOT / "AutoLinker.vcxproj"), "/t:Rebuild",
               "/p:Configuration=fne_release;Platform=Win32", "/m", "/v:normal",
               f"/p:ForceImportBeforeCppTargets={target_file}", "/clp:PerformanceSummary"]
    print(f"{args.mode}: {len(selected_names)}/{len(nodes)} JSON-dependent translation units; "
          f"{len(groups)} PCH variants", flush=True)
    start = time.perf_counter()
    with (OUT / (args.label + ".log")).open("wb") as log:
        result = subprocess.run(command, cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT)
    timing = {"mode": args.mode, "seconds": time.perf_counter() - start,
              "exit_code": result.returncode, "json_translation_units": len(selected_names),
              "total_translation_units": len(nodes), "pch_variants": len(groups),
              "catalog_text": args.catalog_text}
    (OUT / (args.label + "-time.json")).write_text(json.dumps(timing, indent=2), encoding="utf-8")
    print(json.dumps(timing, indent=2), flush=True)
    if result.returncode:
        raise RuntimeError(f"Build failed; see {OUT / (args.label + '.log')}")


if __name__ == "__main__":
    main()
