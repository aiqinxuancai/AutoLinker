// 独立更新协议测试：链接 ReleaseClient 和 DownloadProgressReporter，使用模拟 HTTP。
#include "../src/AutoLinkerReleaseClient.h"
#include "../src/WinINetUtil.h"
#include "../thirdparty/json.hpp"
#include <deque>
#include <fstream>
#include <iostream>
#include <iterator>
#include <stdexcept>

namespace {
struct Reply {
    std::string url;
    int status;
    std::string body;
    std::string finalUrl;
};
std::deque<Reply> replies;
int checks = 0;
void Check(bool ok, const char* message)
{
    if (!ok) throw std::runtime_error(message);
    ++checks;
}
std::string ReadFile(const char* path)
{
    std::ifstream stream(path, std::ios::binary);
    if (!stream) throw std::runtime_error("Cannot read fixture");
    return {std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>()};
}
}

void OutputStringToELog(const std::string&) {}

std::pair<std::string, int> PerformGetRequest(
    const std::string& url, const std::string&, int, bool, bool,
    const HttpDownloadProgressCallback&, std::string* finalUrl)
{
    Check(!replies.empty(), "Unexpected HTTP request (or forbidden API fallback)");
    auto reply = replies.front();
    replies.pop_front();
    Check(reply.url == url, "Unexpected request URL");
    if (finalUrl) *finalUrl = reply.finalUrl;
    return {reply.body, reply.status};
}

std::pair<std::string, int> PerformGetRequest(
    const std::string& url, const std::string& headers, int timeout, bool cookies, bool redirect)
{
    return PerformGetRequest(url, headers, timeout, cookies, redirect, {}, nullptr);
}

int main(int argc, char** argv)
{
    try {
        Check(argc == 4, "Usage: test autolinker.json archive.zip epackager.json");
        for (auto component : {ReleaseComponent::AutoLinker, ReleaseComponent::EPackager}) {
            const bool packager = component == ReleaseComponent::EPackager;
            const std::string repo = packager ? "e-packager" : "AutoLinker";
            const std::string base = "https://github.com/aiqinxuancai/" + repo + "/releases/";
            const std::string assetName = packager ? "e-packager-v5.9.1-windows-win32.zip" : "AutoLinker-5.9.1.zip";
            const auto manifestText = ReadFile(argv[packager ? 3 : 1]);
            const auto archive = ReadFile(argv[2]);
            const auto manifest = nlohmann::json::parse(manifestText);
            AutoLinkerReleaseInfo info;
            std::string error;
            const auto fetch = [&](const std::string& body, int status = 200) {
                replies = {{base + "latest/download/update.json", status, body, {}}};
                return AutoLinkerReleaseClient::FetchLatest(info, error, component);
            };
            Check(fetch(manifestText), "Generated manifest must be accepted");
            Check(info.tag == "v5.9.1" && info.asset.name == assetName, "Version naming");
            Check(info.asset.size == archive.size() && info.asset.digest.size() == 71, "Integrity metadata");
            Check(info.asset.downloadUrl == base + "download/v5.9.1/" + assetName, "Pinned download URL");
            const auto validInfo = info;

            for (const auto& bad : {"null", "[]", "{}", "<html>unavailable</html>"}) {
                Check(!fetch(bad) && !error.empty() && replies.empty(), "Malformed manifest must fail closed");
            }
            const auto reject = [&](nlohmann::json bad) {
                Check(!fetch(bad.dump()) && !error.empty() && replies.empty(), "Invalid field must be rejected");
            };
            auto bad = manifest; bad["schema_version"] = 2; reject(bad);
            bad = manifest; bad["tag"] = "../other"; reject(bad);
            bad = manifest; bad["tag"] = 12; reject(bad);
            bad = manifest; bad["asset"]["name"] = "AutoLinkerSymbols-5.9.1.pdb.zip"; reject(bad);
            for (auto size : {nlohmann::json(-1), nlohmann::json(0), nlohmann::json(1.5), nlohmann::json("100")}) {
                bad = manifest; bad["asset"]["size"] = size; reject(bad);
            }
            bad = manifest; bad["asset"]["sha256"] = std::string(64, 'z'); reject(bad);
            bad = manifest; bad["asset"].erase("sha256"); reject(bad);
            Check(!fetch("Forbidden", 403) && replies.empty(), "403 must not downgrade integrity");
            Check(!fetch("Unavailable", 503) && replies.empty(), "503 must not downgrade integrity");

            replies = {{base + "latest/download/update.json", 404, "", ""},
                {base + "latest", 200, "HTML is not parsed", base + "tag/5.1.19"}};
            Check(AutoLinkerReleaseClient::FetchLatest(info, error, component) && replies.empty(), "Legacy redirect fallback");
            Check(info.tag == "5.1.19" && info.asset.digest.empty() && info.asset.size == 0, "Legacy metadata reset");
            replies = {{base + "latest/download/update.json", 404, "", ""},
                {base + "latest", 200, "", "https://example.com/releases/tag/5.1.19"}};
            Check(!AutoLinkerReleaseClient::FetchLatest(info, error, component), "Reject foreign legacy redirect");

            const std::string accelerated = "https://github-fast.apptest.dev/aiqinxuancai/" + repo + "/releases/download/v5.9.1/" + assetName;
            std::string bytes;
            replies = {{accelerated, 200, archive, ""}};
            Check(AutoLinkerReleaseClient::DownloadArchive(validInfo.asset, bytes, error, component) && bytes == archive,
                "Generated SHA256 and size must validate real archive");
            auto corrupted = archive; corrupted.back() ^= 1;
            replies = {{accelerated, 200, corrupted, ""}, {validInfo.asset.downloadUrl, 200, archive, ""}};
            Check(AutoLinkerReleaseClient::DownloadArchive(validInfo.asset, bytes, error, component) && bytes == archive && replies.empty(),
                "Corrupt mirror must fall back to verified original");
            replies = {{accelerated, 200, corrupted, ""}, {validInfo.asset.downloadUrl, 200, corrupted, ""}};
            Check(!AutoLinkerReleaseClient::DownloadArchive(validInfo.asset, bytes, error, component), "Reject checksum mismatch");
            replies = {{accelerated, 200, archive.substr(0, 4), ""}, {validInfo.asset.downloadUrl, 200, "<html>", ""}};
            Check(!AutoLinkerReleaseClient::DownloadArchive(validInfo.asset, bytes, error, component), "Reject truncated and non-ZIP responses");
        }
        std::cout << checks << " checks passed\n";
        return 0;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
}
