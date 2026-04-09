// Logger.cpp - 统一日志类实现
#include "Logger.h"

#include "Global.h"

#include <Windows.h>

#include <format>
#include <string>

Logger& Logger::Instance()
{
	static Logger instance;
	return instance;
}

void Logger::Open(const std::string& filePath)
{
	std::lock_guard<std::mutex> lock(m_mutex);
	if (m_file.is_open()) {
		m_file.close();
	}
	m_file.open(filePath, std::ios::trunc | std::ios::binary);
	if (!m_file.is_open()) {
		return;
	}
	// 写入 UTF-8 BOM
	static const char kBom[] = "\xEF\xBB\xBF";
	m_file.write(kBom, 3);
	m_file.flush();
}

std::string Logger::BuildTimestamp()
{
	SYSTEMTIME localTime = {};
	GetLocalTime(&localTime);

	TIME_ZONE_INFORMATION tzInfo = {};
	const DWORD tzResult = GetTimeZoneInformation(&tzInfo);

	// Bias 是 UTC 以西的分钟数（东八区 Bias = -480）
	LONG biasMinutes = tzInfo.Bias;
	if (tzResult == TIME_ZONE_ID_DAYLIGHT) {
		biasMinutes += tzInfo.DaylightBias;
	}
	else if (tzResult == TIME_ZONE_ID_STANDARD) {
		biasMinutes += tzInfo.StandardBias;
	}

	// 转为相对 UTC 的偏移量（东正西负）
	const int offsetMinutes = -static_cast<int>(biasMinutes);
	const int offsetHours   = offsetMinutes / 60;
	const int offsetMins    = std::abs(offsetMinutes % 60);
	const char sign = (offsetHours >= 0) ? '+' : '-';

	return std::format(
		"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03} {}{:02}{:02}",
		localTime.wYear,
		localTime.wMonth,
		localTime.wDay,
		localTime.wHour,
		localTime.wMinute,
		localTime.wSecond,
		localTime.wMilliseconds,
		sign,
		std::abs(offsetHours),
		offsetMins);
}

std::string Logger::Utf8ToGbk(const std::string& text)
{
	if (text.empty()) {
		return text;
	}

	const int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		// 不是合法 UTF-8，原样返回
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return text;
	}

	constexpr UINT kGbk = 936;
	const int gbkLen = WideCharToMultiByte(
		kGbk, 0, wide.data(), wideLen, nullptr, 0, nullptr, nullptr);
	if (gbkLen <= 0) {
		return text;
	}

	std::string gbk(static_cast<size_t>(gbkLen), '\0');
	if (WideCharToMultiByte(
		kGbk, 0, wide.data(), wideLen, gbk.data(), gbkLen, nullptr, nullptr) <= 0) {
		return text;
	}
	return gbk;
}

void Logger::Write(const std::string& category, const std::string& message)
{
	std::lock_guard<std::mutex> lock(m_mutex);
	if (!m_file.is_open()) {
		return;
	}
	const std::string line = std::format(
		"[{}] [{}] {}\r\n",
		BuildTimestamp(),
		category,
		message);
	m_file.write(line.data(), static_cast<std::streamsize>(line.size()));
	m_file.flush();
}

void Logger::WriteAndIde(const std::string& category, const std::string& message)
{
	Write(category, message);
	// 输出至 IDE 日志窗口，需 GBK 编码
	OutputStringToELog(Utf8ToGbk("[" + category + "] " + message));
}
