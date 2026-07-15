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

constexpr const char* kLatestReleaseApi =
	"https://api.github.com/repos/aiqinxuancai/AutoLinker/releases/latest";
constexpr const char* kGitHubBaseUrl = "https://github.com/";
constexpr const char* kGitHubAcceleratorBaseUrl = "https://github-fast.apptest.dev/";
constexpr const char* kGitHubHeaders =
	"User-Agent: AutoLinker\r\n"
	"Accept: application/vnd.github+json\r\n";
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

bool EndsWithInsensitive(const std::string& text, const std::string& suffix)
{
	return text.size() >= suffix.size() &&
		ToLowerAscii(text.substr(text.size() - suffix.size())) == ToLowerAscii(suffix);
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

} // namespace

bool AutoLinkerReleaseClient::FetchLatest(AutoLinkerReleaseInfo& outInfo, std::string& outError)
{
	const auto response = PerformGetRequest(
		kLatestReleaseApi,
		kGitHubHeaders,
		kReleaseRequestTimeoutMs,
		false,
		false);
	if (response.second != 200) {
		outError = "读取 AutoLinker 最新 Release 失败：" + DescribeHttpFailure(response);
		return false;
	}

	const json release = json::parse(response.first, nullptr, false);
	if (!release.is_object()) {
		outError = "GitHub 返回的 AutoLinker Release JSON 无效";
		return false;
	}

	AutoLinkerReleaseInfo info;
	info.tag = release.value("tag_name", std::string());
	const json assets = release.value("assets", json::array());
	for (const auto& asset : assets) {
		if (!asset.is_object()) {
			continue;
		}
		const std::string name = asset.value("name", std::string());
		const std::string lowered = ToLowerAscii(name);
		if (lowered.rfind("autolinker-", 0) != 0 || !EndsWithInsensitive(name, ".zip")) {
			continue;
		}
		info.asset.name = name;
		info.asset.downloadUrl = asset.value("browser_download_url", std::string());
		info.asset.digest = asset.value("digest", std::string());
		info.asset.size = asset.value("size", 0ULL);
		break;
	}

	if (info.tag.empty() || info.asset.name.empty() || info.asset.downloadUrl.empty()) {
		outError = "最新 Release 中未找到 AutoLinker-*.zip 更新包";
		return false;
	}

	outInfo = std::move(info);
	return true;
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
