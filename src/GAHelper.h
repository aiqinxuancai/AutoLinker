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
    std::string cid; // �������Ǵ�ĳ���������л�ȡ��

    GAHelper();

public:
    GAHelper(const GAHelper&) = delete;
    GAHelper& operator=(const GAHelper&) = delete;

    static GAHelper* GetInstance();

    // �첽����ҳ����ͼ
    void RequestPageViewAsync(const std::string& page, const std::string& title = "");

    // PerformPostRequest������
    std::pair<std::string, int> PerformPostRequest(const std::string& url, const nlohmann::json& postData, const std::string& customHeaders = "", int timeout = 200000, bool autoCookies = true, bool neverRedirect = true);
};

#endif // GAHELPER_H
