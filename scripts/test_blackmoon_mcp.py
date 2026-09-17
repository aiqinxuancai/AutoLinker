"""通过真实 MCP HTTP 协议验证编译和参数纠错；目标必须是已打开的临时控制台工程副本。"""
import argparse
import hashlib
import json
import time
import urllib.request
from pathlib import Path


class Mcp:
    def __init__(self, endpoint, output):
        self.endpoint, self.output = endpoint, output
        self.session, self.sequence = None, 0

    def rpc(self, method, params, notification=False, allow_error=False):
        self.sequence += 1
        payload = {"jsonrpc": "2.0", "method": method, "params": params}
        if not notification:
            payload["id"] = self.sequence
        prefix = self.output / f"{self.sequence:03d}"
        Path(str(prefix) + "-request.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        headers = {"Content-Type": "application/json; charset=utf-8", "Accept": "application/json, text/event-stream"}
        if self.session:
            headers["Mcp-Session-Id"] = self.session
        request = urllib.request.Request(self.endpoint, json.dumps(payload, ensure_ascii=False).encode("utf-8"), headers)
        with urllib.request.urlopen(request, timeout=180) as response:
            self.session = response.headers.get("Mcp-Session-Id", self.session)
            raw = response.read().decode("utf-8")
            Path(str(prefix) + "-http.json").write_text(json.dumps({"status": response.status, "headers": dict(response.headers)}, indent=2), encoding="utf-8")
        Path(str(prefix) + "-response.json").write_text(raw, encoding="utf-8")
        if notification:
            return None
        result = json.loads(raw)
        if allow_error and "error" in result:
            return {"protocol_error": result["error"]}
        assert "error" not in result, result
        return result["result"]

    def call(self, name, _allow_error=False, **arguments):
        return self.rpc("tools/call", {"name": name, "arguments": arguments}, allow_error=_allow_error)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--endpoint", default="http://127.0.0.1:19207/mcp")
    parser.add_argument("--config", type=Path, default=Path(r"C:\ProgramData\BlackMoon\Config.json"))
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    baseline = hashlib.sha256(args.config.read_bytes()).hexdigest()
    mcp = Mcp(args.endpoint, args.output)
    mcp.rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "blackmoon-mcp-regression", "version": "1"}})
    mcp.rpc("notifications/initialized", {}, notification=True)
    matches = []
    for _ in range(30):
        instances = mcp.call("list_instances")["structuredContent"]["instances"]
        matches = [i for i in instances if Path(i["source_file_path"]).resolve() == args.project.resolve()]
        if matches:
            break
        time.sleep(1)
    assert len(matches) == 1, instances
    assert mcp.call("select_instance", instance_id=matches[0]["instance_id"])["structuredContent"]["ok"]
    catalog = mcp.rpc("tools/list", {})
    tool = next(t for t in catalog["tools"] if t["name"] == "compile_with_output_path")
    assert "hint" in tool["description"] and "omit static_compile" in tool["description"]
    info = mcp.call("get_current_eide_info")["structuredContent"]
    assert info["project_type"] == "win_console_exe" and info["blackmoon_plugin_loaded"], info
    assert set(info["blackmoon_supported_compile_modes"]) == {"asm", "cxx", "vcxx"}, info
    cases = [
        ("ordinary", {"compiler": "eide", "static_compile": True}, None),
        ("asm", {"compiler": "blackmoon", "blackmoon_mode": "asm"}, None),
        ("cxx", {"compiler": "blackmoon", "blackmoon_mode": "cxx"}, None),
        ("vcxx", {"compiler": "blackmoon", "blackmoon_mode": "vcxx"}, None),
        ("默认中文路径", {"compiler": "blackmoon"}, None),
        ("ordinary-stray-mode", {"compiler": "eide", "blackmoon_mode": "asm"}, "blackmoon_mode_requires_blackmoon"),
        ("implicit-stray-mode", {"blackmoon_mode": "vcxx"}, "blackmoon_mode_requires_blackmoon"),
        ("blackmoon-static-true", {"compiler": "blackmoon", "static_compile": True}, "blackmoon_static_compile_not_supported"),
        ("blackmoon-static-false", {"compiler": "blackmoon", "static_compile": False}, "blackmoon_static_compile_not_supported"),
        ("wrong-target", {"compiler": "blackmoon", "target": "win_dll"}, "blackmoon_target_override_not_supported"),
        ("wrong-mode", {"compiler": "blackmoon", "blackmoon_mode": "invalid"}, "rpc:blackmoon_mode"),
        ("wrong-compiler", {"compiler": "invalid"}, "rpc:compiler"),
        ("wrong-type", {"compiler": "blackmoon", "blackmoon_mode": 2}, "rpc:blackmoon_mode"),
    ]
    rows = []
    for name, options, expected in cases:
        output = args.output / (name + ".exe")
        arguments = {"target": "auto", "output_path": str(output.resolve()), **options}
        response = mcp.call("compile_with_output_path", _allow_error=True, **arguments)
        result = response.get("structuredContent", {})
        text = "\n".join(c["text"] for c in response.get("content", []) if c["type"] == "text")
        checks = {}
        if expected and expected.startswith("rpc:"):
            protocol_error = response.get("protocol_error", {})
            checks["schema_rejection"] = protocol_error.get("code") == -32602
            checks["field_identified"] = expected[4:] in protocol_error.get("message", "")
            checks["no_artifact"] = not output.exists()
            result = {"error": protocol_error}
        elif expected:
            checks["error_code"] = result.get("error") == expected
            checks["failed_status"] = result.get("ok") is False and response.get("isError") is True
            checks["actionable_hint"] = bool(result.get("hint"))
            checks["text_client_receives_hint"] = result.get("hint", "MISSING") in text
            checks["no_artifact"] = not output.exists()
        else:
            checks["success"] = result.get("ok") is True and not response.get("isError", False)
            checks["artifact_updated"] = result.get("artifact_verified") is True and result.get("output_file_modified_after_compile") is True and output.exists()
            if options["compiler"] == "blackmoon":
                mode = options.get("blackmoon_mode", "vcxx")
                profile = {"asm": "汇编模式", "cxx": "C/C++模式", "vcxx": "VC++模式"}[mode]
                checks["actual_mode"] = "当前配置: " + profile in result.get("output_window_text", "")
                checks["mode_returned"] = result.get("blackmoon_mode") == mode
                checks["config_restored"] = result.get("config_restored") is True
                checks["static_not_claimed"] = result.get("static_compile", "MISSING") is None
                checks["path_injected"] = result.get("output_path_injected") is True
        checks["config_hash_unchanged"] = hashlib.sha256(args.config.read_bytes()).hexdigest() == baseline
        checks["no_pending_recovery"] = not Path(str(args.config) + ".autolinker-recovery.json").exists()
        row = {"name": name, "passed": all(checks.values()), "checks": checks, "error": result.get("error"), "request_number": mcp.sequence}
        rows.append(row)
        print(json.dumps(row, ensure_ascii=False), flush=True)
        (args.output / "summary.json").write_text(json.dumps({"transport": "MCP HTTP JSON-RPC tools/call", "instance": matches[0], "config_sha256": baseline, "cases": rows, "ok": all(r["passed"] for r in rows)}, ensure_ascii=False, indent=2), encoding="utf-8")
        if not checks["config_hash_unchanged"]:
            raise AssertionError("Configuration changed: stopping tests and preserving evidence")
    assert all(r["passed"] for r in rows), rows


if __name__ == "__main__":
    main()
