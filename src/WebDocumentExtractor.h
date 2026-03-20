#pragma once

#include <string>
#include <vector>

#include "WebDocumentClient.h"

// 抽取出的文档链接。
struct ExtractedWebDocumentLink {
	std::string text;
	std::string url;
};

// 抽取出的网页文档。
struct ExtractedWebDocument {
	bool ok = false;
	std::string url;
	int httpStatus = 0;
	std::string contentType;
	std::string title;
	std::string plainText;
	std::string excerpt;
	std::vector<ExtractedWebDocumentLink> links;
	std::string error;
};

// Web 文档正文提取器。
class WebDocumentExtractor {
public:
	static ExtractedWebDocument Extract(const HttpFetchResult& fetchResult);
};
