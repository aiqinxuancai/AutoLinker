#include "DownloadProgressReporter.h"

#include <algorithm>
#include <iomanip>
#include <iterator>
#include <sstream>
#include <utility>

namespace {

constexpr auto kProgressReportInterval = std::chrono::seconds(5);

std::string FormatByteSize(std::uint64_t bytes)
{
	static constexpr const char* kUnits[] = { "B", "KB", "MB", "GB", "TB" };
	double value = static_cast<double>(bytes);
	size_t unitIndex = 0;
	while (value >= 1024.0 && unitIndex + 1 < std::size(kUnits)) {
		value /= 1024.0;
		++unitIndex;
	}
	if (unitIndex == 0) {
		return std::to_string(bytes) + " " + kUnits[unitIndex];
	}
	std::ostringstream output;
	output << std::fixed << std::setprecision(2) << value << " " << kUnits[unitIndex];
	return output.str();
}

} // namespace

DownloadProgressReporter::DownloadProgressReporter(
	std::string logPrefix,
	std::uint64_t expectedTotalBytes)
	: logPrefix_(std::move(logPrefix)),
	  totalBytes_(expectedTotalBytes)
{
}

std::optional<std::string> DownloadProgressReporter::Update(
	std::uint64_t downloadedBytes,
	std::uint64_t responseTotalBytes)
{
	const bool totalChanged = responseTotalBytes > 0 && responseTotalBytes != totalBytes_;
	if (responseTotalBytes > 0) {
		totalBytes_ = responseTotalBytes;
	}

	unsigned int percentage = 0;
	if (totalBytes_ > 0) {
		const std::uint64_t boundedBytes = (std::min)(downloadedBytes, totalBytes_);
		percentage = static_cast<unsigned int>(
			static_cast<long double>(boundedBytes) * 100.0L / static_cast<long double>(totalBytes_));
	}

	const auto now = std::chrono::steady_clock::now();
	const bool intervalReached = hasReported_ && now - lastReportedAt_ >= kProgressReportInterval;
	const bool percentageReached = totalBytes_ > 0 && percentage >= nextPercentThreshold_;
	const bool completed = totalBytes_ > 0 && downloadedBytes >= totalBytes_ && lastReportedBytes_ < totalBytes_;
	if (hasReported_ && !totalChanged && !intervalReached && !percentageReached && !completed) {
		return std::nullopt;
	}

	hasReported_ = true;
	lastReportedAt_ = now;
	lastReportedBytes_ = downloadedBytes;
	while (nextPercentThreshold_ <= percentage && nextPercentThreshold_ <= 100) {
		nextPercentThreshold_ += 10;
	}

	if (totalBytes_ == 0) {
		return logPrefix_ + " 下载进度：已下载 " + FormatByteSize(downloadedBytes) + " / 总大小未知";
	}
	return logPrefix_ + " 下载进度：" + std::to_string(percentage) + "%（" +
		FormatByteSize(downloadedBytes) + " / " + FormatByteSize(totalBytes_) + "）";
}
