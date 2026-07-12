#pragma once

#include <string>

// AutoLinker Release 资源信息。
struct AutoLinkerReleaseAsset {
	std::string name;
	std::string downloadUrl;
	std::string digest;
	unsigned long long size = 0;
};

// AutoLinker 最新 Release 信息。
struct AutoLinkerReleaseInfo {
	std::string tag;
	AutoLinkerReleaseAsset asset;
};

// GitHub Release 客户端：读取最新版本并下载、校验更新包。
class AutoLinkerReleaseClient {
public:
	static bool FetchLatest(AutoLinkerReleaseInfo& outInfo, std::string& outError);
	static bool DownloadArchive(
		const AutoLinkerReleaseAsset& asset,
		std::string& outBytes,
		std::string& outError);
};
