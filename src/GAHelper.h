#ifndef GAHELPER_H
#define GAHELPER_H

#include <string>
#include <mutex>
#include <json.hpp>

class GAHelper {
private:
    static GAHelper* instance;
    static std::mutex mtx;
    std::string userAgent;
    inline static const std::string GAUrl = "https://www.google-analytics.com/mp/collect?api_secret=O5FOlVSiSmGmuGtj2xSZcQ&measurement_id=G-SDSDNLC0JV";
    std::string cid; // 假设这是从某个配置类中获取的

    GAHelper();

public:
    GAHelper(const GAHelper&) = delete;
    GAHelper& operator=(const GAHelper&) = delete;

    static GAHelper* GetInstance();

    // 异步请求页面视图
    void RequestPageViewAsync(const std::string& page, const std::string& title = "");

    // PerformPostRequest的声明
    std::pair<std::string, int> PerformPostRequest(const std::string& url, const nlohmann::json& postData, const std::string& customHeaders = "", int timeout = 200000, bool autoCookies = true, bool neverRedirect = true);
};

#endif // GAHELPER_H
