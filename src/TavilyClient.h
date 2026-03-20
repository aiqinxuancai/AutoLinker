#pragma once

#include <string>

// Tavily 搜索结果。
struct TavilySearchResult {
	bool ok = false;
	int httpStatus = 0;
	std::string error;
	std::string normalizedResultJsonUtf8;
};

// Tavily 搜索客户端。
class TavilyClient {
public:
	static TavilySearchResult Search(
		const std::string& apiKey,
		const std::string& queryUtf8,
		int maxResults,
		const std::string& topicUtf8);
};
