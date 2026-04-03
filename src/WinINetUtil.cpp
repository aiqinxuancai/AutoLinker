#include "WinINetUtil.h"

#include <cstring>

#pragma comment(lib, "wininet.lib")

namespace {
constexpr int kHttpStatusCancelled = 499;
constexpr const char* kCancelledResponseText = "Request cancelled";
} // namespace

void HttpRequestCancellation::CloseHandleLocked(HINTERNET& handle)
{
	if (handle == nullptr) {
		return;
	}
	HINTERNET closingHandle = handle;
	handle = nullptr;
	InternetCloseHandle(closingHandle);
}

void HttpRequestCancellation::AttachHandleLocked(HINTERNET& slot, HINTERNET handle)
{
	slot = handle;
	if (cancelled_.load()) {
		CloseHandleLocked(slot);
	}
}

void HttpRequestCancellation::CloseRegisteredHandleLocked(HINTERNET& slot, HINTERNET handle)
{
	if (handle == nullptr || slot != handle) {
		return;
	}
	CloseHandleLocked(slot);
}

void HttpRequestCancellation::Cancel()
{
	cancelled_.store(true);
	std::lock_guard<std::mutex> guard(mutex_);
	CloseHandleLocked(requestHandle_);
	CloseHandleLocked(connectionHandle_);
	CloseHandleLocked(internetHandle_);
}

bool HttpRequestCancellation::IsCancelled() const
{
	return cancelled_.load();
}

void HttpRequestCancellation::AttachInternetHandle(HINTERNET handle)
{
	std::lock_guard<std::mutex> guard(mutex_);
	AttachHandleLocked(internetHandle_, handle);
}

void HttpRequestCancellation::AttachConnectionHandle(HINTERNET handle)
{
	std::lock_guard<std::mutex> guard(mutex_);
	AttachHandleLocked(connectionHandle_, handle);
}

void HttpRequestCancellation::AttachRequestHandle(HINTERNET handle)
{
	std::lock_guard<std::mutex> guard(mutex_);
	AttachHandleLocked(requestHandle_, handle);
}

void HttpRequestCancellation::CloseRegisteredInternetHandle(HINTERNET handle)
{
	std::lock_guard<std::mutex> guard(mutex_);
	CloseRegisteredHandleLocked(internetHandle_, handle);
}

void HttpRequestCancellation::CloseRegisteredConnectionHandle(HINTERNET handle)
{
	std::lock_guard<std::mutex> guard(mutex_);
	CloseRegisteredHandleLocked(connectionHandle_, handle);
}

void HttpRequestCancellation::CloseRegisteredRequestHandle(HINTERNET handle)
{
	std::lock_guard<std::mutex> guard(mutex_);
	CloseRegisteredHandleLocked(requestHandle_, handle);
}

namespace {
std::pair<std::string, int> PerformPostRequestCore(
    const std::string& url,
    const std::string& postData,
    const std::string& customHeaders,
    int timeout,
    bool autoCookies,
    bool neverRedirect,
    const std::function<bool(const std::string& chunk)>* onChunk,
    HttpRequestCancellation* cancellation)
{
    HINTERNET hInternet = nullptr;
    HINTERNET hConnect = nullptr;
    HINTERNET hRequest = nullptr;

    std::string response;
    DWORD statusCode = 0;
    const auto closeRequest = [&]() {
        if (cancellation != nullptr) {
            cancellation->CloseRegisteredRequestHandle(hRequest);
            hRequest = nullptr;
            return;
        }
        if (hRequest != nullptr) {
            InternetCloseHandle(hRequest);
            hRequest = nullptr;
        }
    };
    const auto closeConnect = [&]() {
        if (cancellation != nullptr) {
            cancellation->CloseRegisteredConnectionHandle(hConnect);
            hConnect = nullptr;
            return;
        }
        if (hConnect != nullptr) {
            InternetCloseHandle(hConnect);
            hConnect = nullptr;
        }
    };
    const auto closeInternet = [&]() {
        if (cancellation != nullptr) {
            cancellation->CloseRegisteredInternetHandle(hInternet);
            hInternet = nullptr;
            return;
        }
        if (hInternet != nullptr) {
            InternetCloseHandle(hInternet);
            hInternet = nullptr;
        }
    };
    const auto cleanupAll = [&]() {
        closeRequest();
        closeConnect();
        closeInternet();
    };
    const auto cancelledResult = [&]() -> std::pair<std::string, int> {
        cleanupAll();
        return std::make_pair(std::string(kCancelledResponseText), kHttpStatusCancelled);
    };

    DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_DONT_CACHE |
        INTERNET_FLAG_IGNORE_CERT_CN_INVALID | INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;

    if (autoCookies || neverRedirect) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;
    }
    if (autoCookies) {
        dwFlags |= INTERNET_FLAG_NO_COOKIES;
    }

    if (cancellation != nullptr && cancellation->IsCancelled()) {
        return cancelledResult();
    }

    hInternet = InternetOpenA("HttpPostApp", INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0);
    if (!hInternet) {
        return std::make_pair("Error in InternetOpen", 0);
    }
    if (cancellation != nullptr) {
        cancellation->AttachInternetHandle(hInternet);
        if (cancellation->IsCancelled()) {
            return cancelledResult();
        }
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
        closeInternet();
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
        if (cancellation != nullptr && cancellation->IsCancelled()) {
            return cancelledResult();
        }
        closeInternet();
        return std::make_pair("Error in InternetConnect", 0);
    }
    if (cancellation != nullptr) {
        cancellation->AttachConnectionHandle(hConnect);
        if (cancellation->IsCancelled()) {
            return cancelledResult();
        }
    }

    hRequest = HttpOpenRequestA(hConnect, "POST", urlComp.lpszUrlPath, nullptr, nullptr, nullptr, dwFlags, 0);
    if (!hRequest) {
        if (cancellation != nullptr && cancellation->IsCancelled()) {
            return cancelledResult();
        }
        closeConnect();
        closeInternet();
        return std::make_pair("Error in HttpOpenRequest", 0);
    }
    if (cancellation != nullptr) {
        cancellation->AttachRequestHandle(hRequest);
        if (cancellation->IsCancelled()) {
            return cancelledResult();
        }
    }

    if (!HttpSendRequestA(
        hRequest,
        customHeaders.c_str(),
        static_cast<DWORD>(customHeaders.length()),
        const_cast<char*>(postData.data()),
        static_cast<DWORD>(postData.size()))) {
        if (cancellation != nullptr && cancellation->IsCancelled()) {
            return cancelledResult();
        }
        cleanupAll();
        return std::make_pair("Error in HttpSendRequest", 0);
    }
    if (cancellation != nullptr && cancellation->IsCancelled()) {
        return cancelledResult();
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
    while (true) {
        if (cancellation != nullptr && cancellation->IsCancelled()) {
            return cancelledResult();
        }
        if (!InternetReadFile(hRequest, buffer, sizeof(buffer), &bytesRead)) {
            if (cancellation != nullptr && cancellation->IsCancelled()) {
                return cancelledResult();
            }
            break;
        }
        if (bytesRead == 0) {
            break;
        }
        response.append(buffer, bytesRead);
        if (onChunk != nullptr && *onChunk) {
            if (!(*onChunk)(std::string(buffer, bytesRead))) {
                break;
            }
        }
        if (cancellation != nullptr && cancellation->IsCancelled()) {
            return cancelledResult();
        }
    }

    cleanupAll();

    return std::make_pair(response, static_cast<int>(statusCode));
}
} // namespace

std::pair<std::string, int> PerformPostRequest(
    const std::string& url,
    const std::string& postData,
    const std::string& customHeaders,
    int timeout,
    bool AutoCookies,
    bool NeverRedirect,
    HttpRequestCancellation* cancellation)
{
    return PerformPostRequestCore(
        url,
        postData,
        customHeaders,
        timeout,
        AutoCookies,
        NeverRedirect,
        nullptr,
        cancellation);
}

std::pair<std::string, int> PerformPostRequestStreaming(
    const std::string& url,
    const std::string& postData,
    const std::function<bool(const std::string& chunk)>& onChunk,
    const std::string& customHeaders,
    int timeout,
    bool AutoCookies,
    bool NeverRedirect,
    HttpRequestCancellation* cancellation)
{
    return PerformPostRequestCore(
        url,
        postData,
        customHeaders,
        timeout,
        AutoCookies,
        NeverRedirect,
        &onChunk,
        cancellation);
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
