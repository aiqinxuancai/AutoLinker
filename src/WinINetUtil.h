#pragma once

#include <windows.h>
#include <wininet.h>

#include <atomic>
#include <cstdint>
#include <functional>
#include <mutex>
#include <string>
#include <vector>

// HTTP request cancellation context.
class HttpRequestCancellation {
public:
	void Cancel();
	bool IsCancelled() const;

	// Register active WinINet handles for cancellation.
	void AttachInternetHandle(HINTERNET handle);
	void AttachConnectionHandle(HINTERNET handle);
	void AttachRequestHandle(HINTERNET handle);
	void CloseRegisteredInternetHandle(HINTERNET handle);
	void CloseRegisteredConnectionHandle(HINTERNET handle);
	void CloseRegisteredRequestHandle(HINTERNET handle);

private:
	void CloseHandleLocked(HINTERNET& handle);
	void AttachHandleLocked(HINTERNET& slot, HINTERNET handle);
	void CloseRegisteredHandleLocked(HINTERNET& slot, HINTERNET handle);

	std::atomic_bool cancelled_{ false };
	mutable std::mutex mutex_;
	HINTERNET internetHandle_ = nullptr;
	HINTERNET connectionHandle_ = nullptr;
	HINTERNET requestHandle_ = nullptr;
};

// HTTP response header entry.
struct HttpResponseHeaderEntry {
	std::string name;
	std::string value;
};

// HTTP response with body, status and headers.
struct HttpResponseDetails {
	std::string body;
	int statusCode = 0;
	std::vector<HttpResponseHeaderEntry> headers;

	std::string GetHeaderValue(const std::string& name) const;
};

// HTTP 下载进度回调，参数依次为已下载字节数和响应总字节数；总大小未知时为 0。
using HttpDownloadProgressCallback =
	std::function<void(std::uint64_t downloadedBytes, std::uint64_t totalBytes)>;

// Execute HTTP POST.
std::pair<std::string, int> PerformPostRequest(
	const std::string& url,
	const std::string& postData,
	const std::string& customHeaders = "",
	int timeout = 200000,
	bool AutoCookies = true,
	bool NeverRedirect = true,
	HttpRequestCancellation* cancellation = nullptr);

// Execute HTTP POST and return response headers.
HttpResponseDetails PerformPostRequestDetailed(
	const std::string& url,
	const std::string& postData,
	const std::string& customHeaders = "",
	int timeout = 200000,
	bool AutoCookies = true,
	bool NeverRedirect = true,
	HttpRequestCancellation* cancellation = nullptr);

// Execute streaming HTTP POST.
std::pair<std::string, int> PerformPostRequestStreaming(
	const std::string& url,
	const std::string& postData,
	const std::function<bool(const std::string& chunk)>& onChunk,
	const std::string& customHeaders = "",
	int timeout = 200000,
	bool AutoCookies = true,
	bool NeverRedirect = true,
	HttpRequestCancellation* cancellation = nullptr);

// Execute HTTP GET.
std::pair<std::string, int> PerformGetRequest(
	const std::string& url,
	const std::string& customHeaders = "",
	int timeout = 200000,
	bool AutoCookies = true,
	bool NeverRedirect = true);

// 执行 HTTP GET，并报告下载进度。
std::pair<std::string, int> PerformGetRequest(
	const std::string& url,
	const std::string& customHeaders,
	int timeout,
	bool AutoCookies,
	bool NeverRedirect,
	const HttpDownloadProgressCallback& onProgress);
