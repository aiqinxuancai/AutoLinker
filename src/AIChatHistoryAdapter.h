#pragma once

#include "..\thirdparty\json.hpp"

// 在发送边界转换混合协议历史，不修改会话中保存的原始记录。
namespace AIChatHistoryAdapter {
	// 合并 Responses 助手输出和并行工具调用，重建 Chat Completions 的工具结果。
	void ConvertResponsesItemsToChat(nlohmann::json& messages, bool includeReasoningContent);
}
