#pragma once

#include <functional>
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
	// 抓取文本，并允许调用方在关闭或请求取消时提前终止读取。
	static HttpFetchResult FetchTextUrl(
		const std::string& urlUtf8,
		int timeoutSeconds,
		size_t maxBytes,
		const std::function<bool()>& cancelCallback);
	// 构建无需联网的 URL 安全与 UTF-8 截断自检报告。
	static std::string BuildSelfTestReportJson();
};
