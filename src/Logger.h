// Logger.h - 统一日志类，全项目共享单例，避免到处造轮子
// 日志文件格式：UTF-8 with BOM，每行以本机时区毫秒精度时间戳开头
#pragma once

#include <fstream>
#include <mutex>
#include <string>

// 统一日志类，将日志写入文件（UTF-8）并可选输出至 IDE 日志窗口（GBK）
class Logger {
public:
	static Logger& Instance();

	// 打开日志文件（每次启动覆盖旧内容，写入 UTF-8 BOM）
	// filePath 为完整路径，父目录须已存在
	void Open(const std::string& filePath);

	// 写入一条日志到文件（category 如 "MCP"、"Tool"、"Init"）
	// 格式：[时间戳] [category] message\r\n
	void Write(const std::string& category, const std::string& message);

	// 写入一条日志到文件，同时通过 OutputStringToELog 输出至 IDE 日志窗口（GBK）
	// message 须为 UTF-8 编码
	void WriteAndIde(const std::string& category, const std::string& message);

	// 构建本机时区毫秒精度时间戳字符串，可供外部使用（如启动追踪）
	// 格式示例：2026-04-09 09:20:41.501 +0800
	static std::string BuildTimestamp();

private:
	Logger() = default;

	// 将 UTF-8 文本转换为 GBK，供 IDE 日志窗口使用
	static std::string Utf8ToGbk(const std::string& text);

	std::mutex m_mutex;
	std::ofstream m_file;
};
