"""验证生产渲染/转义函数经过 LPCWSTR 边界后仍能传输完整历史，可读取真实会话。"""
import argparse
import json
import os
from pathlib import Path
import subprocess
import tempfile


parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--session", type=Path)
parser.add_argument("--source", type=Path, default=Path(__file__).resolve().parents[1] / "src/AIChatFeature.cpp")
args = parser.parse_args()
source = args.source.read_text(encoding="utf-8-sig")


def function_block(start, end):
    begin = source.index(start)
    return source[begin:source.index(end, begin)]


functions = function_block("std::wstring EscapeJsSingleQuotedWide(", "std::string EscapeHtml(")
functions += function_block("std::string RenderPlainTextToHtml(", "std::string RenderMarkdownToHtml(")
parts = ["PE\0\0: after NUL, digit \0" + "7; quotes '\"; slash \\; 中文 😀\u2028\u2029\r\n" +
         "".join(chr(i) for i in range(32)) + "<>&"]
real_nuls = 0
if args.session:
    session = json.loads(args.session.read_text(encoding="utf-8-sig"))
    for message in session["messages"]:
        if not message.get("visible_in_history", True):
            continue
        content = message.get("content", "")
        if content.startswith("@@AUTOLINKER_TOOL_TRANSCRIPT@@\n"):
            transcript = json.loads(content.split("\n", 1)[1])
            content = transcript.get("invocation", "") + "\n" + transcript.get("result_preview", "")
        real_nuls += content.count("\0")
        parts.append(content)
    print(f"Session messages={len(session['messages'])}, visible payload NULs={real_nuls}")

harness = r'''
#include <windows.h>
#include <string>
#include <fstream>
#include <iterator>
#include <cwchar>
/* FUNCTIONS */
std::wstring wide(const std::string& s) {
    std::wstring result(MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0), 0);
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), result.data(), (int)result.size());
    return result;
}
void save(const char* path, const std::wstring& s) {
    std::string result(WideCharToMultiByte(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0, nullptr, nullptr), 0);
    WideCharToMultiByte(CP_UTF8, 0, s.data(), (int)s.size(), result.data(), (int)result.size(), nullptr, nullptr);
    std::ofstream(path, std::ios::binary).write(result.data(), result.size());
}
int main() {
    std::ifstream in("input.txt", std::ios::binary);
    const std::string text((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
    const auto html = wide(RenderPlainTextToHtml(text));
    save("expected.txt", html);
    auto single = L"capture('" + EscapeJsSingleQuotedWide(html) + L"');";
    auto doubleQuoted = L"capture(\"" + EscapeJsDoubleQuotedWide(html) + L"\");";
    // 模拟 ExecuteScript 接收 LPCWSTR 时在首个 NUL 截断的实际接口语义。
    single.resize(wcslen(single.c_str()));
    doubleQuoted.resize(wcslen(doubleQuoted.c_str()));
    save("single.js", single);
    save("double.js", doubleQuoted);
}
'''.replace("/* FUNCTIONS */", functions)
node_test = r'''
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const expected = fs.readFileSync('expected.txt', 'utf8');
for (const file of ['single.js', 'double.js']) {
    let received;
    vm.runInNewContext(fs.readFileSync(file, 'utf8'), { capture: value => { received = value; } });
    assert.equal(received, expected, file + ': payload truncated or changed');
}
console.log('PASS: both JavaScript string forms survived LPCWSTR and preserved the complete rendered payload');
'''
vswhere = Path(os.environ["ProgramFiles(x86)"]) / "Microsoft Visual Studio/Installer/vswhere.exe"
installation = subprocess.check_output([
    str(vswhere), "-latest", "-requires", "Microsoft.VisualStudio.Component.VC.Tools.x86.x64",
    "-property", "installationPath"
], text=True).strip()
with tempfile.TemporaryDirectory(prefix="chat-script-") as directory:
    work = Path(directory)
    (work / "input.txt").write_text("\n".join(parts), encoding="utf-8", newline="")
    (work / "test.cpp").write_text(harness, encoding="utf-8-sig", newline="\r\n")
    (work / "test.cjs").write_text(node_test, encoding="utf-8")
    devcmd = Path(installation) / "Common7/Tools/VsDevCmd.bat"
    command = f'call "{devcmd}" -arch=x86 -host_arch=x64 >nul && cl /nologo /EHsc /std:c++20 /utf-8 test.cpp /Fe:test.exe'
    subprocess.run('cmd.exe /d /s /c "' + command + '"', cwd=work, check=True)
    subprocess.run([str(work / "test.exe")], cwd=work, check=True)
    subprocess.run(["node", "test.cjs"], cwd=work, check=True)
