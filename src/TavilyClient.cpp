#include "TavilyClient.h"

#include <algorithm>
#include <string>

#include "..\\thirdparty\\json.hpp"

#include "WinINetUtil.h"

namespace {
constexpr const char* kTavilySearchEndpoint = "https://api.tavily.com/search";

std::string TrimAscii(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && static_cast<unsigned char>(text[begin]) <= 0x20) {
		++begin;
	}
	while (end > begin && static_cast<unsigned char>(text[end - 1]) <= 0x20) {
		--end;
	}
	return text.substr(begin, end - begin);
}
} // namespace

TavilySearchResult TavilyClient::Search(
	const std::string& apiKey,
	const std::string& queryUtf8,
	int maxResults,
	const std::string& topicUtf8)
{
	TavilySearchResult result = {};
	const std::string trimmedKey = TrimAscii(apiKey);
	const std::string trimmedQuery = TrimAscii(queryUtf8);
	const std::string trimmedTopic = TrimAscii(topicUtf8);

	if (trimmedKey.empty()) {
		result.error = "tavily api key is not configured";
		return result;
	}
	if (trimmedQuery.empty()) {
		result.error = "query is required";
		return result;
	}

	nlohmann::json requestBody;
	requestBody["api_key"] = trimmedKey;
	requestBody["query"] = trimmedQuery;
	requestBody["topic"] = trimmedTopic.empty() ? "general" : trimmedTopic;
	requestBody["search_depth"] = "basic";
	requestBody["max_results"] = (std::clamp)(maxResults, 1, 10);
	requestBody["include_answer"] = false;
	requestBody["include_images"] = false;
	requestBody["include_raw_content"] = false;

	const auto [responseBody, statusCode] = PerformPostRequest(
		kTavilySearchEndpoint,
		requestBody.dump(),
		"Content-Type: application/json\r\n",
		90000,
		false,
		false);
	result.httpStatus = statusCode;

	if (statusCode < 200 || statusCode >= 300) {
		result.error = "Tavily HTTP request failed";
		if (!responseBody.empty()) {
			result.error += ": " + responseBody;
		}
		return result;
	}

	nlohmann::json parsed;
	try {
		parsed = nlohmann::json::parse(responseBody);
	}
	catch (const std::exception& ex) {
		result.error = std::string("failed to parse Tavily response: ") + ex.what();
		return result;
	}

	if (parsed.contains("error") && parsed["error"].is_string()) {
		result.error = parsed["error"].get<std::string>();
		return result;
	}

	nlohmann::json normalized;
	normalized["ok"] = true;
	normalized["query"] = trimmedQuery;
	normalized["topic"] = requestBody["topic"];
	normalized["http_status"] = statusCode;
	normalized["results"] = nlohmann::json::array();

	if (parsed.contains("results") && parsed["results"].is_array()) {
		for (const auto& item : parsed["results"]) {
			nlohmann::json row;
			row["title"] = item.value("title", "");
			row["url"] = item.value("url", "");
			row["content"] = item.value("content", "");
			row["score"] = item.value("score", 0.0);
			row["site_name"] = item.value("site_name", "");
			normalized["results"].push_back(std::move(row));
		}
	}

	if (parsed.contains("response_time")) {
		normalized["response_time"] = parsed["response_time"];
	}

	result.ok = true;
	result.normalizedResultJsonUtf8 = normalized.dump();
	return result;
}
