#pragma once

#include <chrono>
#include <cstdint>
#include <optional>
#include <string>

// 下载进度报告器：按时间或百分比阈值生成适合日志输出的进度文本。
class DownloadProgressReporter {
public:
	explicit DownloadProgressReporter(std::string logPrefix, std::uint64_t expectedTotalBytes = 0);

	// 更新已下载大小；需要输出日志时返回 UTF-8 文本，否则返回空值。
	std::optional<std::string> Update(
		std::uint64_t downloadedBytes,
		std::uint64_t responseTotalBytes = 0);

private:
	std::string logPrefix_;
	std::uint64_t totalBytes_ = 0;
	std::uint64_t lastReportedBytes_ = 0;
	unsigned int nextPercentThreshold_ = 10;
	bool hasReported_ = false;
	std::chrono::steady_clock::time_point lastReportedAt_ = {};
};
