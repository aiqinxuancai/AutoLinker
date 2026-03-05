#include "WinINetUtil.h"

#include <cstring>

#pragma comment(lib, "wininet.lib")

namespace {
std::pair<std::string, int> PerformPostRequestCore(
    const std::string& url,
    const std::string& postData,
    const std::string& customHeaders,
    int timeout,
    bool autoCookies,
    bool neverRedirect,
    const std::function<bool(const std::string& chunk)>* onChunk)
{
    HINTERNET hInternet = nullptr;
    HINTERNET hConnect = nullptr;
    HINTERNET hRequest = nullptr;

    std::string response;
    DWORD statusCode = 0;

    DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_DONT_CACHE |
        INTERNET_FLAG_IGNORE_CERT_CN_INVALID | INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;

    if (autoCookies || neverRedirect) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;
    }
    if (autoCookies) {
        dwFlags |= INTERNET_FLAG_NO_COOKIES;
    }

    hInternet = InternetOpenA("HttpPostApp", INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0);
    if (!hInternet) {
        return std::make_pair("Error in InternetOpen", 0);
    }

    InternetSetOption(hInternet, INTERNET_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));

    URL_COMPONENTSA urlComp = {};
    urlComp.dwStructSize = sizeof(urlComp);

    char hostName[256] = {};
    char urlPath[2048] = {};

    urlComp.lpszHostName = hostName;
    urlComp.dwHostNameLength = static_cast<DWORD>(sizeof(hostName));
    urlComp.lpszUrlPath = urlPath;
    urlComp.dwUrlPathLength = static_cast<DWORD>(sizeof(urlPath));

    if (!InternetCrackUrlA(url.c_str(), static_cast<DWORD>(url.length()), 0, &urlComp)) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetCrackUrl", 0);
    }

    if (urlComp.nScheme == INTERNET_SCHEME_HTTPS) {
        dwFlags |= INTERNET_FLAG_SECURE;
    }
    else {
        dwFlags |= INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS;
    }

    hConnect = InternetConnectA(
        hInternet,
        urlComp.lpszHostName,
        urlComp.nPort,
        nullptr,
        nullptr,
        INTERNET_SERVICE_HTTP,
        0,
        0);
    if (!hConnect) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetConnect", 0);
    }

    hRequest = HttpOpenRequestA(hConnect, "POST", urlComp.lpszUrlPath, nullptr, nullptr, nullptr, dwFlags, 0);
    if (!hRequest) {
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpOpenRequest", 0);
    }

    if (!HttpSendRequestA(
        hRequest,
        customHeaders.c_str(),
        static_cast<DWORD>(customHeaders.length()),
        const_cast<char*>(postData.data()),
        static_cast<DWORD>(postData.size()))) {
        InternetCloseHandle(hRequest);
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpSendRequest", 0);
    }

    char statusCodeStr[32] = {};
    DWORD length = static_cast<DWORD>(sizeof(statusCodeStr));
    if (HttpQueryInfoA(hRequest, HTTP_QUERY_STATUS_CODE, &statusCodeStr, &length, nullptr)) {
        try {
            statusCode = static_cast<DWORD>(std::stoi(statusCodeStr));
        }
        catch (...) {
            statusCode = 0;
        }
    }

    char buffer[4096] = {};
    DWORD bytesRead = 0;
    while (InternetReadFile(hRequest, buffer, sizeof(buffer), &bytesRead) && bytesRead > 0) {
        response.append(buffer, bytesRead);
        if (onChunk != nullptr && *onChunk) {
            if (!(*onChunk)(std::string(buffer, bytesRead))) {
                break;
            }
        }
    }

    InternetCloseHandle(hRequest);
    InternetCloseHandle(hConnect);
    InternetCloseHandle(hInternet);

    return std::make_pair(response, static_cast<int>(statusCode));
}
} // namespace

std::pair<std::string, int> PerformPostRequest(
    const std::string& url,
    const std::string& postData,
    const std::string& customHeaders,
    int timeout,
    bool AutoCookies,
    bool NeverRedirect)
{
    return PerformPostRequestCore(
        url,
        postData,
        customHeaders,
        timeout,
        AutoCookies,
        NeverRedirect,
        nullptr);
}

std::pair<std::string, int> PerformPostRequestStreaming(
    const std::string& url,
    const std::string& postData,
    const std::function<bool(const std::string& chunk)>& onChunk,
    const std::string& customHeaders,
    int timeout,
    bool AutoCookies,
    bool NeverRedirect)
{
    return PerformPostRequestCore(
        url,
        postData,
        customHeaders,
        timeout,
        AutoCookies,
        NeverRedirect,
        &onChunk);
}

std::pair<std::string, int> PerformGetRequest(const std::string& url, const std::string& customHeaders, int timeout, bool AutoCookies, bool NeverRedirect) {
    HINTERNET hInternet, hConnect, hRequest;
    std::string response;
    DWORD statusCode = 0;
    DWORD length = sizeof(statusCode);
    char statusCodeStr[32] = { 0 };  // Buffer for the status code string

    DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_DONT_CACHE | INTERNET_FLAG_IGNORE_CERT_CN_INVALID | INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;

    // 根据设置处理重定向和 Cookies
    if (AutoCookies) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT | INTERNET_FLAG_NO_COOKIES;
    }
    if (NeverRedirect) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;
    }

    hInternet = InternetOpenA("HttpGetApp", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);

    if (!hInternet) {
        return std::make_pair("Error in InternetOpen", 0);
    }

    InternetSetOption(hInternet, INTERNET_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));

    URL_COMPONENTSA urlComp;
    memset(&urlComp, 0, sizeof(urlComp));
    urlComp.dwStructSize = sizeof(urlComp);

    char hostName[256] = {};
    char urlPath[2048] = {};
    urlComp.lpszHostName = hostName;
    urlComp.dwHostNameLength = static_cast<DWORD>(sizeof(hostName));
    urlComp.lpszUrlPath = urlPath;
    urlComp.dwUrlPathLength = static_cast<DWORD>(sizeof(urlPath));

    // 解析 URL
    if (!InternetCrackUrlA(url.c_str(), static_cast<DWORD>(url.length()), 0, &urlComp)) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetCrackUrl", 0);
    }

    // 根据 URL 是 HTTP 还是 HTTPS 设置标志
    if (urlComp.nScheme == INTERNET_SCHEME_HTTPS) {
        dwFlags |= INTERNET_FLAG_SECURE;
    }
    else {
        dwFlags |= INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS;
    }

    // 连接到 HTTP 服务器
    hConnect = InternetConnectA(hInternet, urlComp.lpszHostName, urlComp.nPort, NULL, NULL, INTERNET_SERVICE_HTTP, 0, 0);
    if (!hConnect) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetConnect", 0);
    }

    // 创建 HTTP 请求句柄
    hRequest = HttpOpenRequestA(hConnect, "GET", urlComp.lpszUrlPath, NULL, NULL, NULL, dwFlags, 0);
    if (!hRequest) {
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpOpenRequest", 0);
    }

    // 准备自定义头部和发送请求
    std::string headers = customHeaders;
    if (!HttpSendRequestA(hRequest, headers.c_str(), static_cast<DWORD>(headers.length()), NULL, 0)) {
        InternetCloseHandle(hRequest);
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpSendRequest", 0);
    }

    length = static_cast<DWORD>(sizeof(statusCodeStr));
    if (HttpQueryInfoA(hRequest, HTTP_QUERY_STATUS_CODE, &statusCodeStr, &length, NULL)) {
        statusCode = std::stoi(statusCodeStr);
    }
    else {
        statusCode = 0;
    }

    char buffer[1024];
    DWORD bytesRead;
    while (InternetReadFile(hRequest, buffer, 1024, &bytesRead) && bytesRead > 0) {
        response.append(buffer, bytesRead);
    }

    InternetCloseHandle(hRequest);
    InternetCloseHandle(hConnect);
    InternetCloseHandle(hInternet);

    // 返回响应和状态码
    return std::make_pair(response, static_cast<int>(statusCode));
}
