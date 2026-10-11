#pragma once

#include <string>

// 使用固定仓库与附件命名规则的更新组件。
enum class ReleaseComponent { AutoLinker, EPackager };

// Release 资源信息。
struct AutoLinkerReleaseAsset {
	std::string name;
	std::string downloadUrl;
	std::string digest;
	unsigned long long size = 0;
};

// 组件最新 Release 信息。
struct AutoLinkerReleaseInfo {
	std::string tag;
	AutoLinkerReleaseAsset asset;
};

// GitHub Release 客户端：读取 update.json 清单并下载、校验更新包，无需 API Token。
class AutoLinkerReleaseClient {
public:
	static bool FetchLatest(AutoLinkerReleaseInfo& outInfo, std::string& outError,
		ReleaseComponent component = ReleaseComponent::AutoLinker);
	static bool DownloadArchive(
		const AutoLinkerReleaseAsset& asset,
		std::string& outBytes,
		std::string& outError,
		ReleaseComponent component = ReleaseComponent::AutoLinker);
};
