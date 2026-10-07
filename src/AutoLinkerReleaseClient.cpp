#include "AutoLinkerReleaseClient.h"

#include <Windows.h>
#include <bcrypt.h>

#include <algorithm>
#include <cctype>
#include <utility>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "DownloadProgressReporter.h"
#include "Global.h"
#include "WinINetUtil.h"

#pragma comment(lib, "bcrypt.lib")

namespace {

using json = nlohmann::json;

constexpr const char* kReleaseBaseUrl = "https://github.com/aiqinxuancai/AutoLinker/releases/";
constexpr const char* kUpdateManifestUrl =
	"https://github.com/aiqinxuancai/AutoLinker/releases/latest/download/update.json";
constexpr const char* kGitHubBaseUrl = "https://github.com/";
constexpr const char* kGitHubAcceleratorBaseUrl = "https://github-fast.apptest.dev/";
constexpr const char* kGitHubHeaders =
	"User-Agent: AutoLinker\r\n"
	"Accept: */*\r\n";
constexpr int kReleaseRequestTimeoutMs = 60000;
constexpr int kDownloadTimeoutMs = 300000;

std::string LocalFromUtf8(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int wideLength = MultiByteToWideChar(
		CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLength <= 0) {
		return text;
	}
	std::wstring wide(static_cast<size_t>(wideLength), L'\0');
	MultiByteToWideChar(
		CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), wide.data(), wideLength);
	const int localLength = WideCharToMultiByte(
		CP_ACP, 0, wide.data(), wideLength, nullptr, 0, nullptr, nullptr);
	if (localLength <= 0) {
		return text;
	}
	std::string local(static_cast<size_t>(localLength), '\0');
	WideCharToMultiByte(CP_ACP, 0, wide.data(), wideLength, local.data(), localLength, nullptr, nullptr);
	return local;
}

void OutputUpdateLog(const std::string& text)
{
	OutputStringToELog(LocalFromUtf8(text));
}

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string BuildUpdateAssetName(const std::string& releaseTag)
{
	std::string versionLabel = releaseTag;
	if (!versionLabel.empty() && (versionLabel.front() == 'v' || versionLabel.front() == 'V')) {
		versionLabel.erase(versionLabel.begin());
	}
	return "AutoLinker-" + versionLabel + ".zip";
}

std::string BuildAcceleratedGitHubUrl(const std::string& url)
{
	const std::string baseUrl = kGitHubBaseUrl;
	if (url.rfind(baseUrl, 0) != 0) {
		return url;
	}
	return std::string(kGitHubAcceleratorBaseUrl) + url.substr(baseUrl.size());
}

std::string DescribeHttpFailure(const std::pair<std::string, int>& response)
{
	std::string error = response.second == 0
		? std::string("未收到 HTTP 响应")
		: "HTTP " + std::to_string(response.second);
	if (!response.first.empty()) {
		error += ": " + response.first.substr(0, (std::min<size_t>)(response.first.size(), 300));
	}
	return error;
}

bool IsZipBytes(const std::string& bytes)
{
	if (bytes.size() < 4 || bytes[0] != 'P' || bytes[1] != 'K') {
		return false;
	}
	return (bytes[2] == '\x03' && bytes[3] == '\x04') ||
		(bytes[2] == '\x05' && bytes[3] == '\x06') ||
		(bytes[2] == '\x07' && bytes[3] == '\x08');
}

bool ComputeSha256(const std::string& bytes, std::string& outHex)
{
	outHex.clear();
	BCRYPT_ALG_HANDLE algorithm = nullptr;
	BCRYPT_HASH_HANDLE hash = nullptr;
	DWORD objectLength = 0;
	DWORD hashLength = 0;
	DWORD resultLength = 0;
	std::vector<unsigned char> hashObject;
	std::vector<unsigned char> hashBytes;

	bool ok = BCryptOpenAlgorithmProvider(&algorithm, BCRYPT_SHA256_ALGORITHM, nullptr, 0) >= 0;
	if (ok) {
		ok = BCryptGetProperty(
			algorithm,
			BCRYPT_OBJECT_LENGTH,
			reinterpret_cast<PUCHAR>(&objectLength),
			sizeof(objectLength),
			&resultLength,
			0) >= 0;
	}
	if (ok) {
		ok = BCryptGetProperty(
			algorithm,
			BCRYPT_HASH_LENGTH,
			reinterpret_cast<PUCHAR>(&hashLength),
			sizeof(hashLength),
			&resultLength,
			0) >= 0;
	}
	if (ok) {
		hashObject.resize(objectLength);
		hashBytes.resize(hashLength);
		ok = BCryptCreateHash(
			algorithm,
			&hash,
			hashObject.data(),
			static_cast<ULONG>(hashObject.size()),
			nullptr,
			0,
			0) >= 0;
	}
	if (ok) {
		ok = bytes.size() <= ULONG_MAX &&
			BCryptHashData(
				hash,
				reinterpret_cast<PUCHAR>(const_cast<char*>(bytes.data())),
				static_cast<ULONG>(bytes.size()),
				0) >= 0 &&
			BCryptFinishHash(hash, hashBytes.data(), static_cast<ULONG>(hashBytes.size()), 0) >= 0;
	}

	if (hash != nullptr) {
		BCryptDestroyHash(hash);
	}
	if (algorithm != nullptr) {
		BCryptCloseAlgorithmProvider(algorithm, 0);
	}
	if (!ok) {
		return false;
	}

	static constexpr char kHex[] = "0123456789abcdef";
	outHex.reserve(hashBytes.size() * 2);
	for (const unsigned char byte : hashBytes) {
		outHex.push_back(kHex[byte >> 4]);
		outHex.push_back(kHex[byte & 0x0F]);
	}
	return true;
}

bool ValidateDownloadedArchive(
	const std::string& bytes,
	const AutoLinkerReleaseAsset& asset,
	std::string& outError)
{
	if (!IsZipBytes(bytes)) {
		outError = "响应内容不是 ZIP 文件（" + std::to_string(bytes.size()) + " 字节）";
		return false;
	}
	if (asset.size > 0 && bytes.size() != asset.size) {
		outError = "文件大小不匹配，预期 " + std::to_string(asset.size) +
			" 字节，实际 " + std::to_string(bytes.size()) + " 字节";
		return false;
	}

	constexpr const char* digestPrefix = "sha256:";
	if (asset.digest.rfind(digestPrefix, 0) == 0) {
		std::string actualDigest;
		if (!ComputeSha256(bytes, actualDigest)) {
			outError = "无法计算更新包 SHA-256";
			return false;
		}
		const std::string expectedDigest = ToLowerAscii(
			asset.digest.substr(std::char_traits<char>::length(digestPrefix)));
		if (actualDigest != expectedDigest) {
			outError = "更新包 SHA-256 校验失败";
			return false;
		}
	}
	return true;
}

// 标签只允许版本号使用的 ASCII 字符，禁止路径、转义和查询参数。
bool IsSafeReleaseTag(const std::string& tag)
{
	if (tag.empty() || tag.size() > 128) {
		return false;
	}
	const size_t start = (tag.front() == 'v' || tag.front() == 'V') ? 1 : 0;
	return start < tag.size() && tag[start] >= '0' && tag[start] <= '9' &&
		std::all_of(tag.begin(), tag.end(), [](unsigned char ch) {
			return (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'z') ||
				(ch >= 'A' && ch <= 'Z') || ch == '.' || ch == '-' || ch == '+';
		});
}

std::string BuildReleaseDownloadUrl(const AutoLinkerReleaseInfo& info)
{
	return std::string(kReleaseBaseUrl) + "download/" + info.tag + "/" + info.asset.name;
}

bool ParseUpdateManifest(const std::string& body, AutoLinkerReleaseInfo& outInfo, std::string& outError)
{
	const json manifest = json::parse(body, nullptr, false);
	if (!manifest.is_object() || !manifest.contains("schema_version") ||
		!manifest["schema_version"].is_number_unsigned() || manifest["schema_version"] != 1 ||
		!manifest.contains("tag") || !manifest["tag"].is_string() ||
		!manifest.contains("asset") || !manifest["asset"].is_object()) {
		outError = "AutoLinker 更新清单格式无效或版本不受支持";
		return false;
	}
	AutoLinkerReleaseInfo info;
	info.tag = manifest["tag"].get<std::string>();
	const auto& asset = manifest["asset"];
	if (!IsSafeReleaseTag(info.tag) ||
		!asset.contains("name") || !asset["name"].is_string() ||
		!asset.contains("size") || !asset["size"].is_number_unsigned() ||
		!asset.contains("sha256") || !asset["sha256"].is_string()) {
		outError = "AutoLinker 更新清单缺少有效的版本号、文件名、大小或 SHA-256";
		return false;
	}
	info.asset.name = asset["name"].get<std::string>();
	info.asset.size = asset["size"].get<unsigned long long>();
	const std::string hash = ToLowerAscii(asset["sha256"].get<std::string>());
	if (info.asset.name != BuildUpdateAssetName(info.tag) || info.asset.size == 0 ||
		hash.size() != 64 || !std::all_of(hash.begin(), hash.end(), [](unsigned char ch) {
			return (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f');
		})) {
		outError = "AutoLinker 更新清单的更新包名称、大小或 SHA-256 不合法";
		return false;
	}
	info.asset.digest = "sha256:" + hash;
	info.asset.downloadUrl = BuildReleaseDownloadUrl(info);
	outInfo = std::move(info);
	return true;
}

bool FetchLegacyRelease(AutoLinkerReleaseInfo& outInfo, std::string& outError)
{
	std::string finalUrl;
	const auto response = PerformGetRequest(
		std::string(kReleaseBaseUrl) + "latest", kGitHubHeaders,
		kReleaseRequestTimeoutMs, false, false, {}, &finalUrl);
	if (response.second != 200) {
		outError = "读取旧版 AutoLinker Release 页面失败：" + DescribeHttpFailure(response);
		return false;
	}
	const std::string prefix = std::string(kReleaseBaseUrl) + "tag/";
	AutoLinkerReleaseInfo info;
	if (finalUrl.rfind(prefix, 0) != 0 || !IsSafeReleaseTag(finalUrl.substr(prefix.size()))) {
		outError = "无法从 AutoLinker Release 页面重定向地址获取有效版本号";
		return false;
	}
	info.tag = finalUrl.substr(prefix.size());
	info.asset.name = BuildUpdateAssetName(info.tag);
	info.asset.downloadUrl = BuildReleaseDownloadUrl(info);
	OutputUpdateLog("[AutoLinker更新] 此 Release 尚无 update.json，使用旧版兼容方式（无文件大小和 SHA-256 校验）");
	outInfo = std::move(info);
	return true;
}

} // namespace

bool AutoLinkerReleaseClient::FetchLatest(AutoLinkerReleaseInfo& outInfo, std::string& outError)
{
	outError.clear();
	const auto response = PerformGetRequest(
		kUpdateManifestUrl, kGitHubHeaders, kReleaseRequestTimeoutMs, false, false);
	if (response.second == 404) {
		return FetchLegacyRelease(outInfo, outError);
	}
	if (response.second != 200) {
		outError = "读取 AutoLinker 更新清单失败：" + DescribeHttpFailure(response);
		return false;
	}
	return ParseUpdateManifest(response.first, outInfo, outError);
}

bool AutoLinkerReleaseClient::DownloadArchive(
	const AutoLinkerReleaseAsset& asset,
	std::string& outBytes,
	std::string& outError)
{
	const auto download = [&asset](const std::string& url) {
		DownloadProgressReporter progress("[AutoLinker更新]", asset.size);
		return PerformGetRequest(
			url,
			kGitHubHeaders,
			kDownloadTimeoutMs,
			false,
			false,
			[&progress](std::uint64_t downloadedBytes, std::uint64_t totalBytes) {
				if (const auto message = progress.Update(downloadedBytes, totalBytes)) {
					OutputUpdateLog(*message);
				}
			});
	};

	const std::string acceleratedUrl = BuildAcceleratedGitHubUrl(asset.downloadUrl);
	OutputUpdateLog("[AutoLinker更新] 通过 GitHub 加速地址下载：" + acceleratedUrl);
	auto response = download(acceleratedUrl);
	std::string validationError;
	if (response.second == 200 && ValidateDownloadedArchive(response.first, asset, validationError)) {
		outBytes = std::move(response.first);
		return true;
	}

	const std::string acceleratedError = response.second == 200
		? validationError
		: DescribeHttpFailure(response);
	OutputUpdateLog("[AutoLinker更新] 加速地址下载失败，将尝试原始 GitHub 地址：" + acceleratedError);
	response = download(asset.downloadUrl);
	validationError.clear();
	if (response.second == 200 && ValidateDownloadedArchive(response.first, asset, validationError)) {
		outBytes = std::move(response.first);
		return true;
	}

	const std::string githubError = response.second == 200
		? validationError
		: DescribeHttpFailure(response);
	outError = "加速地址下载失败（" + acceleratedError + "）；原始 GitHub 下载失败（" + githubError + "）";
	return false;
}
