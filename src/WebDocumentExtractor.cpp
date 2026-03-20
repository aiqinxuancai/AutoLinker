#include "WebDocumentExtractor.h"

#include <algorithm>
#include <cctype>
#include <string>

#include "..\\thirdparty\\json.hpp"

namespace {
constexpr size_t kMaxLinks = 20;
constexpr size_t kMaxPlainTextChars = 30000;
constexpr size_t kExcerptChars = 600;

std::string ToLowerAscii(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string TrimAscii(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

bool StartsWithInsensitiveAt(const std::string& text, size_t pos, const std::string& needle)
{
	if (pos + needle.size() > text.size()) {
		return false;
	}
	for (size_t i = 0; i < needle.size(); ++i) {
		if (std::tolower(static_cast<unsigned char>(text[pos + i])) !=
			std::tolower(static_cast<unsigned char>(needle[i]))) {
			return false;
		}
	}
	return true;
}

std::string DecodeHtmlEntities(const std::string& text)
{
	std::string out;
	out.reserve(text.size());
	for (size_t i = 0; i < text.size(); ++i) {
		if (text[i] != '&') {
			out.push_back(text[i]);
			continue;
		}

		const size_t semi = text.find(';', i + 1);
		if (semi == std::string::npos || semi - i > 16) {
			out.push_back(text[i]);
			continue;
		}

		const std::string entity = text.substr(i + 1, semi - i - 1);
		if (entity == "amp") {
			out.push_back('&');
		}
		else if (entity == "lt") {
			out.push_back('<');
		}
		else if (entity == "gt") {
			out.push_back('>');
		}
		else if (entity == "quot") {
			out.push_back('"');
		}
		else if (entity == "apos" || entity == "#39") {
			out.push_back('\'');
		}
		else if (entity == "nbsp") {
			out.push_back(' ');
		}
		else if (!entity.empty() && entity[0] == '#') {
			try {
				unsigned int codePoint = 0;
				if (entity.size() > 1 && (entity[1] == 'x' || entity[1] == 'X')) {
					codePoint = static_cast<unsigned int>(std::stoul(entity.substr(2), nullptr, 16));
				}
				else {
					codePoint = static_cast<unsigned int>(std::stoul(entity.substr(1), nullptr, 10));
				}
				if (codePoint <= 0x7F) {
					out.push_back(static_cast<char>(codePoint));
				}
				else if (codePoint <= 0x7FF) {
					out.push_back(static_cast<char>(0xC0 | ((codePoint >> 6) & 0x1F)));
					out.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
				}
				else if (codePoint <= 0xFFFF) {
					out.push_back(static_cast<char>(0xE0 | ((codePoint >> 12) & 0x0F)));
					out.push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
					out.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
				}
				else {
					out.push_back(' ');
				}
			}
			catch (...) {
				out += "&" + entity + ";";
			}
		}
		else {
			out += "&" + entity + ";";
		}

		i = semi;
	}
	return out;
}

std::string StripTags(const std::string& html)
{
	std::string out;
	out.reserve(html.size());
	bool inTag = false;
	for (char ch : html) {
		if (ch == '<') {
			inTag = true;
			continue;
		}
		if (ch == '>') {
			inTag = false;
			continue;
		}
		if (!inTag) {
			out.push_back(ch);
		}
	}
	return DecodeHtmlEntities(out);
}

std::string RemoveTagBlocks(const std::string& html, const std::string& tagName)
{
	std::string lower = ToLowerAscii(html);
	std::string out;
	size_t pos = 0;
	const std::string openTag = "<" + tagName;
	const std::string closeTag = "</" + tagName;
	while (pos < html.size()) {
		const size_t openPos = lower.find(openTag, pos);
		if (openPos == std::string::npos) {
			out += html.substr(pos);
			break;
		}
		out += html.substr(pos, openPos - pos);
		const size_t closePos = lower.find(closeTag, openPos);
		if (closePos == std::string::npos) {
			break;
		}
		const size_t closeEnd = lower.find('>', closePos);
		if (closeEnd == std::string::npos) {
			break;
		}
		pos = closeEnd + 1;
	}
	return out;
}

std::string RemoveHtmlComments(const std::string& html)
{
	std::string out;
	size_t pos = 0;
	while (pos < html.size()) {
		const size_t begin = html.find("<!--", pos);
		if (begin == std::string::npos) {
			out += html.substr(pos);
			break;
		}
		out += html.substr(pos, begin - pos);
		const size_t end = html.find("-->", begin + 4);
		if (end == std::string::npos) {
			break;
		}
		pos = end + 3;
	}
	return out;
}

std::string ExtractTitle(const std::string& html)
{
	const std::string lower = ToLowerAscii(html);
	const size_t open = lower.find("<title");
	if (open == std::string::npos) {
		return std::string();
	}
	const size_t openEnd = lower.find('>', open);
	if (openEnd == std::string::npos) {
		return std::string();
	}
	const size_t close = lower.find("</title>", openEnd + 1);
	if (close == std::string::npos) {
		return std::string();
	}
	return TrimAscii(StripTags(html.substr(openEnd + 1, close - openEnd - 1)));
}

bool IsBlockTag(const std::string& lowerTag)
{
	static const char* kTags[] = {
		"br", "p", "/p", "div", "/div", "section", "/section", "article", "/article",
		"main", "/main", "header", "/header", "footer", "/footer", "aside", "/aside",
		"h1", "/h1", "h2", "/h2", "h3", "/h3", "h4", "/h4", "h5", "/h5", "h6", "/h6",
		"li", "/li", "ul", "/ul", "ol", "/ol", "pre", "/pre", "tr", "/tr", "table", "/table"
	};
	for (const char* tag : kTags) {
		if (lowerTag == tag) {
			return true;
		}
	}
	return false;
}

std::string HtmlToPlainText(const std::string& html)
{
	std::string cleaned = RemoveHtmlComments(html);
	cleaned = RemoveTagBlocks(cleaned, "script");
	cleaned = RemoveTagBlocks(cleaned, "style");
	cleaned = RemoveTagBlocks(cleaned, "noscript");

	std::string out;
	out.reserve(cleaned.size());
	size_t i = 0;
	while (i < cleaned.size()) {
		if (cleaned[i] != '<') {
			out.push_back(cleaned[i]);
			++i;
			continue;
		}

		const size_t end = cleaned.find('>', i + 1);
		if (end == std::string::npos) {
			break;
		}

		std::string tag = TrimAscii(cleaned.substr(i + 1, end - i - 1));
		if (!tag.empty() && tag[0] == '!') {
			i = end + 1;
			continue;
		}
		if (!tag.empty() && tag[0] == '/') {
			tag = "/" + ToLowerAscii(TrimAscii(tag.substr(1)));
		}
		else {
			const size_t space = tag.find_first_of(" \t\r\n/");
			tag = ToLowerAscii(space == std::string::npos ? tag : tag.substr(0, space));
		}

		if (tag == "li") {
			out += "\n- ";
		}
		else if (tag == "td" || tag == "th") {
			out += "\t";
		}
		else if (IsBlockTag(tag)) {
			out.push_back('\n');
		}

		i = end + 1;
	}

	out = DecodeHtmlEntities(out);
	std::string normalized;
	normalized.reserve(out.size());
	bool previousSpace = false;
	int newlineRun = 0;
	for (char ch : out) {
		if (ch == '\r') {
			continue;
		}
		if (ch == '\n') {
			while (!normalized.empty() && normalized.back() == ' ') {
				normalized.pop_back();
			}
			if (newlineRun < 2) {
				normalized.push_back('\n');
			}
			++newlineRun;
			previousSpace = false;
			continue;
		}
		if (std::isspace(static_cast<unsigned char>(ch)) != 0) {
			if (!previousSpace && newlineRun == 0) {
				normalized.push_back(' ');
				previousSpace = true;
			}
			continue;
		}
		newlineRun = 0;
		previousSpace = false;
		normalized.push_back(ch);
	}
	return TrimAscii(normalized);
}

std::vector<ExtractedWebDocumentLink> ExtractLinks(const std::string& html)
{
	std::vector<ExtractedWebDocumentLink> links;
	std::string lower = ToLowerAscii(html);
	size_t pos = 0;
	while (pos < html.size() && links.size() < kMaxLinks) {
		const size_t open = lower.find("<a", pos);
		if (open == std::string::npos) {
			break;
		}
		const size_t tagEnd = html.find('>', open + 2);
		if (tagEnd == std::string::npos) {
			break;
		}
		const size_t close = lower.find("</a>", tagEnd + 1);
		if (close == std::string::npos) {
			break;
		}

		const std::string tag = html.substr(open, tagEnd - open + 1);
		const std::string tagLower = ToLowerAscii(tag);
		const size_t hrefPos = tagLower.find("href=");
		if (hrefPos != std::string::npos) {
			size_t valuePos = hrefPos + 5;
			while (valuePos < tag.size() && std::isspace(static_cast<unsigned char>(tag[valuePos])) != 0) {
				++valuePos;
			}

			std::string href;
			if (valuePos < tag.size() && (tag[valuePos] == '"' || tag[valuePos] == '\'')) {
				const char quote = tag[valuePos];
				++valuePos;
				const size_t hrefEnd = tag.find(quote, valuePos);
				if (hrefEnd != std::string::npos) {
					href = tag.substr(valuePos, hrefEnd - valuePos);
				}
			}
			else {
				const size_t hrefEnd = tag.find_first_of(" \t\r\n>", valuePos);
				href = tag.substr(valuePos, hrefEnd == std::string::npos ? std::string::npos : hrefEnd - valuePos);
			}

			href = TrimAscii(DecodeHtmlEntities(href));
			const std::string anchorText = TrimAscii(StripTags(html.substr(tagEnd + 1, close - tagEnd - 1)));
			const std::string hrefLower = ToLowerAscii(href);
			if (!href.empty() &&
				!hrefLower.starts_with("javascript:") &&
				!hrefLower.starts_with("mailto:") &&
				!hrefLower.starts_with("#")) {
				links.push_back(ExtractedWebDocumentLink{ anchorText, href });
			}
		}

		pos = close + 4;
	}
	return links;
}

std::string MakeExcerpt(const std::string& plainText)
{
	if (plainText.size() <= kExcerptChars) {
		return plainText;
	}
	return TrimAscii(plainText.substr(0, kExcerptChars)) + "...";
}
} // namespace

ExtractedWebDocument WebDocumentExtractor::Extract(const HttpFetchResult& fetchResult)
{
	ExtractedWebDocument out = {};
	out.url = fetchResult.finalUrl.empty() ? fetchResult.url : fetchResult.finalUrl;
	out.httpStatus = fetchResult.httpStatus;
	out.contentType = fetchResult.contentType;

	if (!fetchResult.ok) {
		out.error = fetchResult.error.empty() ? "fetch failed" : fetchResult.error;
		return out;
	}

	if (fetchResult.isJson) {
		std::string formatted = fetchResult.bodyText;
		try {
			formatted = nlohmann::json::parse(fetchResult.bodyText).dump(2);
		}
		catch (...) {
		}
		out.title = "JSON Document";
		out.plainText = formatted;
	}
	else if (fetchResult.isHtml) {
		out.title = ExtractTitle(fetchResult.bodyText);
		out.links = ExtractLinks(fetchResult.bodyText);
		out.plainText = HtmlToPlainText(fetchResult.bodyText);
	}
	else {
		out.title.clear();
		out.plainText = TrimAscii(fetchResult.bodyText);
	}

	if (out.plainText.size() > kMaxPlainTextChars) {
		out.plainText.resize(kMaxPlainTextChars);
		out.plainText = TrimAscii(out.plainText) + "...";
	}
	out.excerpt = MakeExcerpt(out.plainText);
	out.ok = true;
	return out;
}
