#pragma once

#include <cstddef>
#include <string>
#include <vector>

// AI 对话 Markdown 表格渲染器：识别 GFM 表格并生成安全的表格结构。
namespace AIChatMarkdownTableRenderer {

using InlineRenderer = std::string (*)(const std::string& text);

// 尝试从 headerLineIndex 开始渲染表格，成功时返回最后消费的行号。
bool TryRenderTable(
	const std::vector<std::string>& lines,
	size_t headerLineIndex,
	InlineRenderer inlineRenderer,
	std::string& html,
	size_t& lastRenderedLineIndex);

} // namespace AIChatMarkdownTableRenderer
