#pragma once

#include <windows.h>
#include <wininet.h>

#include <atomic>
#include <cstdint>
#include <functional>
#include <memory>
#include <mutex>
#include <string>
#include <vector>

// HTTP request cancellation context.
class HttpRequestCancellation {
public:
	HttpRequestCancellation();
	HttpRequestCancellation(const HttpRequestCancellation&) = delete;
	HttpRequestCancellation& operator=(const HttpRequestCancellation&) = delete;

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
	struct State;
	static void CloseHandleLocked(HINTERNET& handle);
	static void AttachHandleLocked(State& state, HINTERNET& slot, HINTERNET handle);
	static void CloseRegisteredHandleLocked(HINTERNET& slot, HINTERNET handle);

	// 取消线程与请求线程共享句柄状态，保证异步关闭期间对象销毁也安全。
	std::shared_ptr<State> state_;
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
	HttpRequestCancellation* cancellation = nullptr,
	// 流式响应的独立接收空闲超时；0 表示沿用请求超时。
	int streamIdleTimeout = 300000);

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
