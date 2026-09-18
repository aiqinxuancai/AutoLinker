#pragma once

#include <cstdint>
#include <span>
#include <string>

// 获取当前 IDE 磁盘文件身份；首次调用计算 MD5，后续复用结果。
const std::string& GetIdeBinaryIdentity();

// 安全读取签名处内存，返回地址、预期字节、实际字节及读取状态。
std::string BuildIdeSignatureDiagnostics(
	std::uintptr_t moduleBase,
	std::uintptr_t rva,
	std::span<const std::uint8_t> expected);
