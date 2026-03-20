#pragma once

#include <string>

// HTTP 文本抓取结果。
struct HttpFetchResult {
	bool ok = false;
	std::string url;
	std::string finalUrl;
	int httpStatus = 0;
	std::string contentType;
	size_t contentLength = 0;
	std::string bodyText;
	bool bodyTruncated = false;
	bool isHtml = false;
	bool isJson = false;
	bool isMarkdown = false;
	std::string error;
};

// Web 文档抓取客户端。
class WebDocumentClient {
public:
	static HttpFetchResult FetchTextUrl(
		const std::string& urlUtf8,
		int timeoutSeconds,
		size_t maxBytes);
};
