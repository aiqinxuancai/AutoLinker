#include "WinINetUtil.h"

#pragma comment(lib, "wininet.lib")

std::pair<std::string, int> PerformPostRequest(const std::string& url, const std::string& postData,
    const std::string& customHeaders, int timeout, bool AutoCookies, bool NeverRedirect) {
    HINTERNET hInternet, hConnect, hRequest;
    std::string response;
    DWORD statusCode = 0;
    DWORD length = sizeof(statusCode);
    char statusCodeStr[32] = { 0 };  // Buffer for the status code string


    DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_DONT_CACHE | INTERNET_FLAG_IGNORE_CERT_CN_INVALID | INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;

    // �������ô����ض���� Cookies
    if (AutoCookies == 1 || NeverRedirect) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;
    }
    if (AutoCookies == 1) {
        dwFlags |= INTERNET_FLAG_NO_COOKIES;
    }

    hInternet = InternetOpen("HttpPostApp", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);

    if (!hInternet) {
        return std::make_pair("Error in InternetOpen", 0);
    }

    InternetSetOption(hInternet, INTERNET_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));

    URL_COMPONENTS urlComp;
    memset(&urlComp, 0, sizeof(urlComp));
    urlComp.dwStructSize = sizeof(urlComp);

    char hostName[256], urlPath[256];

    //Ϊ�˲���IDE��������ӵĴ���
    hostName[sizeof(hostName) - 1] = '\0';
    urlPath[sizeof(urlPath) - 1] = '\0';

    urlComp.lpszHostName = hostName;
    urlComp.dwHostNameLength = 256;
    urlComp.lpszUrlPath = urlPath;
    urlComp.dwUrlPathLength = 256;

   

    // ���� URL
    if (!InternetCrackUrl(url.c_str(), url.length(), 0, &urlComp)) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetCrackUrl", 0);
    }

    // ���� URL �� HTTP ���� HTTPS ���ñ�־
    if (urlComp.nScheme == INTERNET_SCHEME_HTTPS) {
        dwFlags |= INTERNET_FLAG_SECURE;
    }
    else {
        dwFlags |= INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS;
    }

    // ���ӵ� HTTP ������
    hConnect = InternetConnect(hInternet, urlComp.lpszHostName, urlComp.nPort, NULL, NULL, INTERNET_SERVICE_HTTP, 0, 0);
    if (!hConnect) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetConnect", 0);
    }

    // ���� HTTP ������
    hRequest = HttpOpenRequest(hConnect, "POST", urlComp.lpszUrlPath, NULL, NULL, NULL, dwFlags, 0);
    if (!hRequest) {
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpOpenRequest", 0);
    }

    // ׼���Զ���ͷ���ͷ�������
    //std::string headers = "Content-Type: application/x-www-form-urlencoded\r\n" + customHeaders;
    std::string headers = customHeaders;
    if (!HttpSendRequest(hRequest, headers.c_str(), headers.length(), const_cast<char*>(postData.c_str()), postData.size())) {
        InternetCloseHandle(hRequest);
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpSendRequest", 0);
    }

    if (HttpQueryInfo(hRequest, HTTP_QUERY_STATUS_CODE, &statusCodeStr, &length, NULL)) {
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

    // ������Ӧ��״̬��
    return std::make_pair(response, static_cast<int>(statusCode));
}


std::pair<std::string, int> PerformGetRequest(const std::string& url, const std::string& customHeaders, int timeout, bool AutoCookies, bool NeverRedirect) {
    HINTERNET hInternet, hConnect, hRequest;
    std::string response;
    DWORD statusCode = 0;
    DWORD length = sizeof(statusCode);
    char statusCodeStr[32] = { 0 };  // Buffer for the status code string

    DWORD dwFlags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_DONT_CACHE | INTERNET_FLAG_IGNORE_CERT_CN_INVALID | INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;

    // �������ô����ض���� Cookies
    if (AutoCookies) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT | INTERNET_FLAG_NO_COOKIES;
    }
    if (NeverRedirect) {
        dwFlags |= INTERNET_FLAG_NO_AUTO_REDIRECT;
    }

    hInternet = InternetOpen("HttpGetApp", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);

    if (!hInternet) {
        return std::make_pair("Error in InternetOpen", 0);
    }

    InternetSetOption(hInternet, INTERNET_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
    InternetSetOption(hInternet, INTERNET_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));

    URL_COMPONENTS urlComp;
    memset(&urlComp, 0, sizeof(urlComp));
    urlComp.dwStructSize = sizeof(urlComp);

    char hostName[256], urlPath[256];
    urlComp.lpszHostName = hostName;
    urlComp.dwHostNameLength = sizeof(hostName);
    urlComp.lpszUrlPath = urlPath;
    urlComp.dwUrlPathLength = sizeof(urlPath);

    // ���� URL
    if (!InternetCrackUrl(url.c_str(), url.length(), 0, &urlComp)) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetCrackUrl", 0);
    }

    // ���� URL �� HTTP ���� HTTPS ���ñ�־
    if (urlComp.nScheme == INTERNET_SCHEME_HTTPS) {
        dwFlags |= INTERNET_FLAG_SECURE;
    }
    else {
        dwFlags |= INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS;
    }

    // ���ӵ� HTTP ������
    hConnect = InternetConnect(hInternet, urlComp.lpszHostName, urlComp.nPort, NULL, NULL, INTERNET_SERVICE_HTTP, 0, 0);
    if (!hConnect) {
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in InternetConnect", 0);
    }

    // ���� HTTP ������
    hRequest = HttpOpenRequest(hConnect, "GET", urlComp.lpszUrlPath, NULL, NULL, NULL, dwFlags, 0);
    if (!hRequest) {
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpOpenRequest", 0);
    }

    // ׼���Զ���ͷ���ͷ�������
    std::string headers = customHeaders;
    if (!HttpSendRequest(hRequest, headers.c_str(), headers.length(), NULL, 0)) {
        InternetCloseHandle(hRequest);
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return std::make_pair("Error in HttpSendRequest", 0);
    }

    if (HttpQueryInfo(hRequest, HTTP_QUERY_STATUS_CODE, &statusCodeStr, &length, NULL)) {
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

    // ������Ӧ��״̬��
    return std::make_pair(response, static_cast<int>(statusCode));
}
