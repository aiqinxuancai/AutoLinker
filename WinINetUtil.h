#pragma once
#include <windows.h>
#include <wininet.h>
#include <string>
std::pair<std::string, int> PerformPostRequest(const std::string& url, const std::string& postData, const std::string& customHeaders = "",
	int timeout = 200000, bool AutoCookies = true, bool NeverRedirect = true);;


std::pair<std::string, int> PerformGetRequest(const std::string& url, const std::string& customHeaders = "", int timeout = 200000, bool AutoCookies = true, bool NeverRedirect = true);