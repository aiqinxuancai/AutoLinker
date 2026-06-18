module.exports = {
  content: [
    "./src/webview/ai_config_dialog.src.html",
    "./src/webview/linker_config_dialog.src.html"
  ],
  safelist: ["block"],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Microsoft YaHei UI"', '"Segoe UI"', "Arial", "sans-serif"],
        mono: ["Consolas", '"Courier New"', "monospace"]
      }
    }
  },
  plugins: []
};
