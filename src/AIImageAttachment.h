#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <vector>

// AI 图片附件的持久化元数据。
struct AIImageAttachment {
	std::string id;
	std::string fileNameLocal;
	std::string mimeType;
	std::string assetPathLocal;
	std::string sourcePathLocal;
	std::string detail = "auto";
	std::uint64_t byteSize = 0;
	unsigned int width = 0;
	unsigned int height = 0;
};

// 图片输入处理结果。
struct AIImagePrepareResult {
	bool ok = false;
	AIImageAttachment attachment;
	std::string errorLocal;
};

// 负责校验、缩放、快照和编码 AI 图片附件。
class AIImageAttachmentManager {
public:
	static constexpr size_t kMaxAttachmentCount = 10;
	static constexpr std::uint64_t kMaxInputBytes = 20ull * 1024ull * 1024ull;
	static constexpr std::uint64_t kMaxPreparedBytes = 4ull * 1024ull * 1024ull;
	static constexpr unsigned int kMaxDimension = 2048;

	// 将磁盘图片处理并快照到会话资源目录。
	static AIImagePrepareResult PrepareFile(
		const std::filesystem::path& sourcePath,
		const std::filesystem::path& assetDirectory,
		const std::string& detail = "auto");

	// 将 WebView 传入的 data URL 处理并快照到会话资源目录。
	static AIImagePrepareResult PrepareDataUrl(
		const std::string& dataUrlUtf8,
		const std::string& fileNameUtf8,
		const std::filesystem::path& assetDirectory,
		const std::string& detail = "auto");

	// 读取快照并构建供应商请求使用的 Base64 data URL。
	static bool BuildDataUrl(
		const AIImageAttachment& attachment,
		std::string& outDataUrlUtf8,
		std::string& outErrorLocal);

	// 判断路径是否为可尝试读取的图片扩展名。
	static bool IsSupportedImagePath(const std::filesystem::path& path);

	// 构建图片处理与数据 URL 的内部自测报告。
	static std::string BuildSelfTestJson();
};
