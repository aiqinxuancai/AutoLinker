module.exports = {
  content: [
    "./src/webview/ai_config_dialog.src.html",
    "./src/webview/linker_config_dialog.src.html",
    "./src/webview/ec_switch_config_dialog.src.html",
    "./src/webview/force_link_lib_config_dialog.src.html",
    "./src/webview/project_agents_config_dialog.src.html",
    "./src/webview/ai_chat_theme_config_dialog.src.html",
    "./src/webview/ai_chat_mcp_config_dialog.src.html",
    "./src/webview/ai_skill_config_dialog.src.html",
    "./src/webview/ai_other_settings.src.html",
    "./src/webview/settings_toast.partial.html"
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
