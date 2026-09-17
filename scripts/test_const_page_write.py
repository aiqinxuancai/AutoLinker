"""在已打开的临时 .e 工程副本中回归 issue #26；会覆盖该副本的常量表。"""

import argparse
import json
from pathlib import Path
import urllib.request


class Mcp:
    def __init__(self, endpoint, output):
        self.endpoint = endpoint
        self.output = output
        self.session = None
        self.sequence = 0

    def rpc(self, method, params):
        self.sequence += 1
        headers = {"Content-Type": "application/json; charset=utf-8"}
        if self.session:
            headers["Mcp-Session-Id"] = self.session
        body = json.dumps({"jsonrpc": "2.0", "id": self.sequence,
                           "method": method, "params": params}, ensure_ascii=False).encode()
        request = urllib.request.Request(self.endpoint, body, headers)
        with urllib.request.urlopen(request, timeout=120) as response:
            self.session = response.headers.get("Mcp-Session-Id", self.session)
            result = json.load(response)
        (self.output / f"{self.sequence:03d}.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        assert "error" not in result, result
        return result["result"]

    def call(self, name, **arguments):
        return self.rpc("tools/call", {"name": name, "arguments": arguments})["structuredContent"]


def source(result):
    assert result["ok"] and not result.get("has_more"), result
    return "\n".join(line.split("\t", 1)[1] for line in result["content"].splitlines())


def declarations(text):
    return [line.strip() for line in text.splitlines() if line.strip().startswith(".常量 ")]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True, type=Path, help="已打开的可丢弃工程副本")
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--endpoint", default="http://127.0.0.1:19207/mcp")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    mcp = Mcp(args.endpoint, args.output)
    mcp.rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                           "clientInfo": {"name": "const-page-regression", "version": "1"}})
    instances = mcp.call("list_instances")["instances"]
    matches = [i for i in instances if Path(i["source_file_path"]).resolve() == args.project.resolve()]
    assert len(matches) == 1, matches
    assert mcp.call("select_instance", instance_id=matches[0]["instance_id"])["ok"]
    assert mcp.call("refresh_workspace_mirror")["ok"]
    file_path = "src/.常量.txt"

    def read():
        result = mcp.call("read_real_file", file_path=file_path)
        source(result)
        return result

    def write(code):
        base = read()
        return mcp.call("write_file", file_path=file_path, full_code=code,
                        expected_base_hash=base["code_hash"])

    def verify(result, expected):
        assert result["ok"] and result["verified"], result
        actual = read()
        assert declarations(source(actual)) == declarations(expected), actual
        assert result["code_hash"] == actual["code_hash"], (result, actual)
        assert result["new_hash"] == actual["code_hash"], result
        return actual

    # 同时覆盖中文、文本、整数和公开属性。
    body = '\n'.join([
        '.常量 测试常量_文本公开, "“你好世界”", 公开, 文本常量公开测试',
        '.常量 测试常量_文本非公开, "“版本说明”", , 文本常量非公开测试',
        '.常量 测试常量_整数公开, "100", 公开, 整数常量公开测试',
        '.常量 测试常量_整数非公开, "12", , 整数常量非公开测试',
    ])
    cases = {
        "missing_header": body,
        "version_lf": '.版本 2\n\n' + body + '\n',
        "version_crlf": ('.版本 2\n\n' + body + '\n').replace('\n', '\r\n'),
        "leading_blank": '\n\n' + body,
        "no_trailing_newline": '.版本 2\n' + body,
        "quoted_comma": body + '\n.常量 逗号文本, "“甲,乙”", 公开, 中文说明',
        "legacy_five_fields": '.版本 2\n.常量 旧格式, "7", , 公开, 旧格式备注',
    }
    for name, code in cases.items():
        verify(write(code), code)
        print(name + ': ok')
    current = verify(write(body), body)
    changed = body.replace('"100"', '"101"')
    edited = mcp.call("edit_file", file_path=file_path, old_text='"100"', new_text='"101"',
                      expected_base_hash=current["code_hash"])
    current = verify(edited, changed)
    edited = mcp.call("multi_edit_file", file_path=file_path,
                      edits=[{"old_text": '"101"', "new_text": '"100"'},
                             {"old_text": '"12"', "new_text": '"13"'}],
                      expected_base_hash=current["code_hash"])
    verify(edited, body.replace('"12"', '"13"'))
    print('edit_file / multi_edit_file: ok')
    current = verify(write(body), body)
    rejected = write(body.replace('公开, 文本常量公开测试', '非法属性, 文本常量公开测试'))
    assert not rejected["ok"] and not rejected.get("verified", False), rejected
    assert rejected["rollback_attempted"] and rejected["rollback_succeeded"], rejected
    assert read()["code_hash"] == current["code_hash"]
    print('verification mismatch / rollback: ok')
    (args.output / 'summary.json').write_text(json.dumps({"ok": True,
        "cases": list(cases) + ["edit_file", "multi_edit_file", "rollback"],
        "final_hash": current["code_hash"], "final_code": source(current)}, ensure_ascii=False, indent=2), encoding='utf-8')


if __name__ == "__main__":
    main()
