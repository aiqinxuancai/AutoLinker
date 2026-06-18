import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const inputCssPath = path.join(root, "src", "webview", "ai_config_dialog.tailwind.css");
const tempCssPath = path.join(os.tmpdir(), `autolinker-webview-${process.pid}.css`);
const prelinePath = path.join(root, "node_modules", "preline", "dist", "dropdown.js");

// Templates that share the same inlined Tailwind CSS + Preline bundle.
const templates = [
	{ src: "ai_config_dialog.src.html", out: "ai_config_dialog.html" },
	{ src: "linker_config_dialog.src.html", out: "linker_config_dialog.html" }
];

execFileSync(process.execPath, [
	path.join(root, "node_modules", "tailwindcss", "lib", "cli.js"),
	"-c", path.join(root, "tailwind.config.js"),
	"-i", inputCssPath,
	"-o", tempCssPath,
	"--minify"
], {
	cwd: root,
	stdio: "inherit"
});

const css = fs.readFileSync(tempCssPath, "utf8");
const preline = fs.existsSync(prelinePath) ? fs.readFileSync(prelinePath, "utf8") : "";

for (const entry of templates) {
	const templatePath = path.join(root, "src", "webview", entry.src);
	if (!fs.existsSync(templatePath)) {
		continue;
	}
	const outputHtmlPath = path.join(root, "src", "webview", entry.out);
	const template = fs.readFileSync(templatePath, "utf8");
	let html = template
		.replace("/*__AUTOLINKER_AI_CONFIG_CSS__*/", css)
		.replace("/*__AUTOLINKER_PRELINE_JS__*/", preline);
	html = html.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\r\n");
	fs.writeFileSync(outputHtmlPath, `\uFEFF${html}`, "utf8");
}

fs.rmSync(tempCssPath, { force: true });
