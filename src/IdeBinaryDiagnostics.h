#pragma once

#include <cstddef>
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

// 固定版本的 x86 函数档案；字节基于首选装载地址，补丁窗口必须结束于原始指令边界。
struct IdeEntryProfile {
	std::uint32_t timeDateStamp;
	std::uint32_t sizeOfImage;
	std::uintptr_t rva;
	std::span<const std::uint8_t> expected;
	size_t patchWindow;
};

// 分别验证磁盘版本、未覆盖的函数体和入口跳转链；通过后仍调用原入口。
// 仅允许有界 x86 跳转链终止于已加载的其他映像，拒绝未知私有代码终点。
bool ValidateIdeCallableEntry(
	std::uintptr_t moduleBase,
	const IdeEntryProfile& profile,
	std::string& trace);
