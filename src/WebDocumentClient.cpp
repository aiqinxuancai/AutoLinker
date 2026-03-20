#include "WebDocumentClient.h"

#include <Windows.h>
#include <wininet.h>

#include <algorithm>
#include <cctype>
#include <cstring>
#include <string>

#pragma comment(lib, "wininet.lib")

namespace {
constexpr size_t kDefaultMaxBytes = 512 * 1024;
constexpr size_t kMaxAllowedBytes = 2 * 1024 * 1024;
constexpr size_t kMaxReturnedChars = 40000;

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

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0) > 0;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int len = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (len <= 0) {
		return std::string();
	}

	std::string out(static_cast<size_t>(len), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		out.data(),
		len,
		nullptr,
		nullptr) <= 0) {
		return std::string();
	}
	return out;
}

std::string ConvertToUtf8(const std::string& text, UINT codePage, DWORD flags = 0)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(
		codePage,
		flags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::string();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		codePage,
		flags,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return std::string();
	}
	return WideToUtf8(wide);
}

std::string DecodeUtf16ToUtf8(const std::string& bytes, bool bigEndian)
{
	if (bytes.size() < 2) {
		return std::string();
	}

	const size_t charCount = bytes.size() / 2;
	std::wstring wide;
	wide.reserve(charCount);
	for (size_t i = 0; i + 1 < bytes.size(); i += 2) {
		unsigned char b0 = static_cast<unsigned char>(bytes[i]);
		unsigned char b1 = static_cast<unsigned char>(bytes[i + 1]);
		wchar_t ch = bigEndian
			? static_cast<wchar_t>((b0 << 8) | b1)
			: static_cast<wchar_t>((b1 << 8) | b0);
		wide.push_back(ch);
	}
	return WideToUtf8(wide);
}

UINT DetectCharsetCodePage(const std::string& contentTypeLower)
{
	const size_t pos = contentTypeLower.find("charset=");
	if (pos == std::string::npos) {
		return 0;
	}

	std::string charset = TrimAscii(contentTypeLower.substr(pos + 8));
	const size_t semi = charset.find(';');
	if (semi != std::string::npos) {
		charset = TrimAscii(charset.substr(0, semi));
	}
	if (!charset.empty() && (charset.front() == '"' || charset.front() == '\'')) {
		charset.erase(charset.begin());
	}
	if (!charset.empty() && (charset.back() == '"' || charset.back() == '\'')) {
		charset.pop_back();
	}

	if (charset == "utf-8" || charset == "utf8") {
		return CP_UTF8;
	}
	if (charset == "gbk" || charset == "gb2312") {
		return 936;
	}
	if (charset == "gb18030") {
		return 54936;
	}
	if (charset == "big5") {
		return 950;
	}
	if (charset == "shift_jis" || charset == "sjis") {
		return 932;
	}
	if (charset == "euc-kr") {
		return 51949;
	}
	if (charset == "iso-8859-1" || charset == "latin1") {
		return 28591;
	}
	return 0;
}

bool LooksLikeTextContentType(const std::string& contentTypeLower)
{
	return contentTypeLower.starts_with("text/") ||
		contentTypeLower.find("json") != std::string::npos ||
		contentTypeLower.find("xml") != std::string::npos ||
		contentTypeLower.find("javascript") != std::string::npos;
}

bool LooksLikeMarkdownTypeOrUrl(const std::string& contentTypeLower, const std::string& urlLower)
{
	return contentTypeLower.find("markdown") != std::string::npos ||
		urlLower.ends_with(".md") ||
		urlLower.ends_with(".markdown");
}

bool LooksLikeBinary(const std::string& bytes)
{
	if (bytes.empty()) {
		return false;
	}
	size_t suspicious = 0;
	const size_t sample = (std::min)(bytes.size(), static_cast<size_t>(512));
	for (size_t i = 0; i < sample; ++i) {
		const unsigned char ch = static_cast<unsigned char>(bytes[i]);
		if (ch == 0) {
			return true;
		}
		if (ch < 0x09 || (ch > 0x0D && ch < 0x20)) {
			++suspicious;
		}
	}
	return suspicious > sample / 10;
}

std::string StripUtf8Bom(std::string text)
{
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		text.erase(0, 3);
	}
	return text;
}

std::string DecodeBodyToUtf8(const std::string& bytes, const std::string& contentTypeLower)
{
	if (bytes.empty()) {
		return std::string();
	}

	if (bytes.size() >= 2) {
		const unsigned char b0 = static_cast<unsigned char>(bytes[0]);
		const unsigned char b1 = static_cast<unsigned char>(bytes[1]);
		if (b0 == 0xFF && b1 == 0xFE) {
			return DecodeUtf16ToUtf8(bytes.substr(2), false);
		}
		if (b0 == 0xFE && b1 == 0xFF) {
			return DecodeUtf16ToUtf8(bytes.substr(2), true);
		}
	}

	const UINT hintedCodePage = DetectCharsetCodePage(contentTypeLower);
	if (hintedCodePage == CP_UTF8 && IsValidUtf8(bytes)) {
		return StripUtf8Bom(bytes);
	}
	if (hintedCodePage == 1200 || hintedCodePage == 1201) {
		return DecodeUtf16ToUtf8(bytes, hintedCodePage == 1201);
	}
	if (hintedCodePage != 0) {
		const std::string decoded = ConvertToUtf8(bytes, hintedCodePage, hintedCodePage == CP_UTF8 ? MB_ERR_INVALID_CHARS : 0);
		if (!decoded.empty()) {
			return StripUtf8Bom(decoded);
		}
	}
	if (IsValidUtf8(bytes)) {
		return StripUtf8Bom(bytes);
	}
	{
		const std::string decoded = ConvertToUtf8(bytes, CP_ACP, 0);
		if (!decoded.empty()) {
			return decoded;
		}
	}
	return bytes;
}

bool QueryStringInfo(HINTERNET hRequest, DWORD infoLevel, std::string& outText)
{
	outText.clear();
	DWORD len = 0;
	if (HttpQueryInfoA(hRequest, infoLevel, nullptr, &len, nullptr) == FALSE && GetLastError() != ERROR_INSUFFICIENT_BUFFER) {
		return false;
	}
	if (len == 0) {
		return false;
	}

	std::string buffer(static_cast<size_t>(len), '\0');
	if (HttpQueryInfoA(hRequest, infoLevel, buffer.data(), &len, nullptr) == FALSE || len == 0) {
		return false;
	}
	if (!buffer.empty() && buffer.back() == '\0') {
		buffer.pop_back();
	}
	outText = buffer;
	return true;
}
} // namespace

HttpFetchResult WebDocumentClient::FetchTextUrl(
	const std::string& urlUtf8,
	int timeoutSeconds,
	size_t maxBytes)
{
	HttpFetchResult result = {};
	result.url = urlUtf8;
	result.finalUrl = urlUtf8;

	const std::string trimmedUrl = TrimAscii(urlUtf8);
	if (trimmedUrl.empty()) {
		result.error = "url is required";
		return result;
	}

	const size_t boundedMaxBytes = (std::clamp)(
		maxBytes == 0 ? kDefaultMaxBytes : maxBytes,
		static_cast<size_t>(4096),
		kMaxAllowedBytes);
	int timeoutMs = (std::clamp)(timeoutSeconds <= 0 ? 60 : timeoutSeconds, 1, 300) * 1000;

	HINTERNET hInternet = nullptr;
	HINTERNET hConnect = nullptr;
	HINTERNET hRequest = nullptr;

	URL_COMPONENTSA urlComp = {};
	urlComp.dwStructSize = sizeof(urlComp);
	char hostName[256] = {};
	char urlPath[4096] = {};
	char extraInfo[1024] = {};
	urlComp.lpszHostName = hostName;
	urlComp.dwHostNameLength = static_cast<DWORD>(sizeof(hostName));
	urlComp.lpszUrlPath = urlPath;
	urlComp.dwUrlPathLength = static_cast<DWORD>(sizeof(urlPath));
	urlComp.lpszExtraInfo = extraInfo;
	urlComp.dwExtraInfoLength = static_cast<DWORD>(sizeof(extraInfo));

	if (!InternetCrackUrlA(trimmedUrl.c_str(), static_cast<DWORD>(trimmedUrl.size()), 0, &urlComp)) {
		result.error = "invalid url";
		return result;
	}

	DWORD flags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_DONT_CACHE;
	if (urlComp.nScheme == INTERNET_SCHEME_HTTPS) {
		flags |= INTERNET_FLAG_SECURE | INTERNET_FLAG_IGNORE_CERT_CN_INVALID | INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;
	}
	else {
		flags |= INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS;
	}

	hInternet = InternetOpenA("AutoLinker-WebDocumentClient", INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0);
	if (hInternet == nullptr) {
		result.error = "InternetOpen failed";
		return result;
	}

	InternetSetOption(hInternet, INTERNET_OPTION_CONNECT_TIMEOUT, &timeoutMs, sizeof(timeoutMs));
	InternetSetOption(hInternet, INTERNET_OPTION_SEND_TIMEOUT, &timeoutMs, sizeof(timeoutMs));
	InternetSetOption(hInternet, INTERNET_OPTION_RECEIVE_TIMEOUT, &timeoutMs, sizeof(timeoutMs));

	hConnect = InternetConnectA(
		hInternet,
		urlComp.lpszHostName,
		urlComp.nPort,
		nullptr,
		nullptr,
		INTERNET_SERVICE_HTTP,
		0,
		0);
	if (hConnect == nullptr) {
		InternetCloseHandle(hInternet);
		result.error = "InternetConnect failed";
		return result;
	}

	std::string requestPath = urlComp.lpszUrlPath != nullptr ? urlComp.lpszUrlPath : "/";
	if (requestPath.empty()) {
		requestPath = "/";
	}
	if (urlComp.lpszExtraInfo != nullptr && *urlComp.lpszExtraInfo != '\0') {
		requestPath += urlComp.lpszExtraInfo;
	}

	hRequest = HttpOpenRequestA(
		hConnect,
		"GET",
		requestPath.c_str(),
		nullptr,
		nullptr,
		nullptr,
		flags,
		0);
	if (hRequest == nullptr) {
		InternetCloseHandle(hConnect);
		InternetCloseHandle(hInternet);
		result.error = "HttpOpenRequest failed";
		return result;
	}

	std::string headers =
		"Accept: text/html,text/plain,text/markdown,application/xhtml+xml,application/json;q=0.9,*/*;q=0.5\r\n"
		"User-Agent: AutoLinker/0.0.0\r\n";
	if (!HttpSendRequestA(hRequest, headers.c_str(), static_cast<DWORD>(headers.size()), nullptr, 0)) {
		InternetCloseHandle(hRequest);
		InternetCloseHandle(hConnect);
		InternetCloseHandle(hInternet);
		result.error = "HttpSendRequest failed";
		return result;
	}

	std::string statusCodeText;
	if (QueryStringInfo(hRequest, HTTP_QUERY_STATUS_CODE, statusCodeText)) {
		try {
			result.httpStatus = std::stoi(statusCodeText);
		}
		catch (...) {
			result.httpStatus = 0;
		}
	}

	std::string contentType;
	QueryStringInfo(hRequest, HTTP_QUERY_CONTENT_TYPE, contentType);
	result.contentType = contentType;
	const std::string contentTypeLower = ToLowerAscii(contentType);

	std::string contentLengthText;
	if (QueryStringInfo(hRequest, HTTP_QUERY_CONTENT_LENGTH, contentLengthText)) {
		try {
			result.contentLength = static_cast<size_t>(std::stoull(contentLengthText));
		}
		catch (...) {
			result.contentLength = 0;
		}
	}

	std::string rawBody;
	rawBody.reserve((std::min)(boundedMaxBytes, static_cast<size_t>(65536)));
	char buffer[4096] = {};
	DWORD bytesRead = 0;
	while (InternetReadFile(hRequest, buffer, sizeof(buffer), &bytesRead) != FALSE && bytesRead > 0) {
		const size_t remain = boundedMaxBytes > rawBody.size() ? (boundedMaxBytes - rawBody.size()) : 0;
		const size_t copyBytes = (std::min)(remain, static_cast<size_t>(bytesRead));
		if (copyBytes > 0) {
			rawBody.append(buffer, copyBytes);
		}
		if (copyBytes < bytesRead || rawBody.size() >= boundedMaxBytes) {
			result.bodyTruncated = true;
			break;
		}
	}

	InternetCloseHandle(hRequest);
	InternetCloseHandle(hConnect);
	InternetCloseHandle(hInternet);

	if (result.contentLength == 0) {
		result.contentLength = rawBody.size();
	}

	const std::string urlLower = ToLowerAscii(trimmedUrl);
	result.isHtml =
		contentTypeLower.find("text/html") != std::string::npos ||
		contentTypeLower.find("application/xhtml+xml") != std::string::npos ||
		urlLower.ends_with(".html") ||
		urlLower.ends_with(".htm");
	result.isJson = contentTypeLower.find("json") != std::string::npos || urlLower.ends_with(".json");
	result.isMarkdown = LooksLikeMarkdownTypeOrUrl(contentTypeLower, urlLower);

	if (!LooksLikeTextContentType(contentTypeLower) && !result.isHtml && !result.isMarkdown && !result.isJson) {
		if (LooksLikeBinary(rawBody)) {
			result.error = "unsupported content type";
			return result;
		}
	}

	result.bodyText = DecodeBodyToUtf8(rawBody, contentTypeLower);
	if (result.bodyText.size() > kMaxReturnedChars) {
		result.bodyText.resize(kMaxReturnedChars);
		result.bodyTruncated = true;
	}

	if (result.httpStatus < 200 || result.httpStatus >= 300) {
		result.error = "HTTP request failed";
		return result;
	}

	result.ok = true;
	return result;
}
