// 独立 Win32 回归测试：与 IdeBinaryDiagnostics.cpp 一同编译，链接 /DYNAMICBASE:NO。
// 只修改本测试映像的专用字节区，不执行模拟 Hook，也不启动或修改 IDE。
// 在本目录的 x86 Native Tools 命令行运行：
// cl /std:c++20 /EHsc /utf-8 /O2 /W4 IdeEntryValidationTest.cpp ../src/IdeBinaryDiagnostics.cpp /link /DYNAMICBASE:NO
#include "../src/IdeBinaryDiagnostics.h"
#include <Windows.h>
#include <array>
#include <cstring>
#include <iostream>

#pragma section(".entrytest", read, execute)
__declspec(allocate(".entrytest")) unsigned char kEntryFixture[64] = {
	0x6A, 0xFF, 0x68, 0xD8, 0xF1, 0x5E, 0x00, 0x64,
	0xA1, 0x00, 0x00, 0x00, 0x00, 0x50, 0x64, 0x89,
	0x25, 0x00, 0x00, 0x00, 0x00, 0x81, 0xEC, 0x90,
	0x00, 0x00, 0x00, 0x8B, 0x84, 0x24, 0xA0, 0x00,
	0x00, 0x00, 0x55, 0x56, 0x57, 0x8B, 0xF1, 0x50,
	0x89, 0x74, 0x24, 0x14, 0xE8, 0x5F, 0x12, 0x00,
	0x00, 0x6A, 0x00, 0x6A, 0x00, 0x6A, 0x00, 0x6A,
	0x00, 0x6A, 0x00, 0x8D, 0x8C, 0x24, 0x88, 0x00,
};

int main()
{
	if (sizeof(void*) != 4) return 2;
	const auto base = reinterpret_cast<std::uintptr_t>(GetModuleHandleW(nullptr));
	const auto dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
	const auto nt = reinterpret_cast<const IMAGE_NT_HEADERS32*>(base + dos->e_lfanew);
	std::array<std::uint8_t, 64> original{};
	std::memcpy(original.data(), kEntryFixture, original.size());
	auto patched = original;
	auto* entry = patched.data();
	const auto address = reinterpret_cast<std::uintptr_t>(kEntryFixture);
	DWORD protection = 0;
	if (!VirtualProtect(kEntryFixture, original.size(), PAGE_EXECUTE_READWRITE, &protection)) return 2;
	const IdeEntryProfile profile{ nt->FileHeader.TimeDateStamp, nt->OptionalHeader.SizeOfImage,
		address - base, original, 7 };
	unsigned passed = 0, failed = 0;
	auto check = [&](const char* name, bool expected, const char* reason, const IdeEntryProfile& candidate) {
		SIZE_T written = 0;
		if (!WriteProcessMemory(GetCurrentProcess(), kEntryFixture, patched.data(), patched.size(), &written) || written != patched.size()) {
			++failed;
			return;
		}
		std::string trace;
		const bool actual = ValidateIdeCallableEntry(base, candidate, trace);
		const bool ok = actual == expected && trace.find(reason) != std::string::npos;
		(ok ? passed : failed)++;
		std::cout << (ok ? "PASS " : "FAIL ") << name << ": " << trace << '\n';
	};
	check("original", true, "entry=original", profile);
	auto wrongVersion = profile;
	wrongVersion.timeDateStamp ^= 1;
	check("wrong_version", false, "version_or_disk_signature", wrongVersion);
	auto wrongBytes = original;
	wrongBytes[40] ^= 1;
	auto wrongDisk = profile;
	wrongDisk.expected = wrongBytes;
	check("wrong_disk_signature", false, "version_or_disk_signature", wrongDisk);

	auto* stub = static_cast<unsigned char*>(VirtualAlloc(nullptr, 4096, MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE));
	if (!stub) return 2;
	const auto target = reinterpret_cast<std::uintptr_t>(GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "GetCurrentProcessId"));
	if (!target) return 2;
	auto reset = [&] { std::memcpy(entry, original.data(), original.size()); };
	auto jump = [&](unsigned char* code, std::uintptr_t destination) {
		code[0] = 0xE9;
		const auto delta = static_cast<std::uint32_t>(destination - (code == entry ? address : reinterpret_cast<std::uintptr_t>(code)) - 5);
		std::memcpy(code + 1, &delta, sizeof(delta));
	};
	jump(entry, target);
	check("e9_original_padding", true, "entry=hook_chain", profile);
	entry[5] = entry[6] = 0x90;
	check("e9_nop_padding", true, "hook_module=", profile);
	entry[5] = entry[6] = 0xCC;
	check("e9_int3_padding", true, "entry=hook_chain", profile);
	entry[6] = 0x41;
	check("unexpected_padding", false, "unexpected_padding", profile);
	reset();
	jump(entry, target);
	entry[40] ^= 1;
	check("changed_body", false, "body_changed", profile);
	reset();
	entry[0] = 0xCC;
	check("unknown_patch", false, "unsupported_patch", profile);
	reset();
	jump(entry, 1);
	check("unmapped_target", false, "target_not_executable", profile);
	jump(entry, address);
	check("cycle", false, "jump_cycle", profile);
	jump(entry, address + 7);
	check("target_inside_ide", false, "unverified_target_owner", profile);
	stub[0] = 0xC3;
	jump(entry, reinterpret_cast<std::uintptr_t>(stub));
	check("private_terminal", false, "unverified_target_owner", profile);
	jump(stub, target);
	check("private_jump_stub", true, "entry=hook_chain", profile);
	stub[0] = 0xEB;
	stub[1] = 5;
	jump(stub + 7, target);
	check("short_jump_chain", true, "entry=hook_chain", profile);
	jump(stub, address);
	check("two_node_cycle", false, "jump_cycle", profile);
	jump(stub, target);
	DWORD stubProtection = 0;
	VirtualProtect(stub, 4096, PAGE_READWRITE, &stubProtection);
	check("non_executable_target", false, "target_not_executable", profile);
	VirtualProtect(stub, 4096, stubProtection, &stubProtection);
	for (unsigned i = 0; i < 10; ++i) jump(stub + i * 8, reinterpret_cast<std::uintptr_t>(stub + (i + 1) * 8));
	jump(stub + 80, target);
	check("bounded_chain", false, "jump_depth_limit", profile);
	reset();
	entry[0] = 0x68;
	const auto target32 = static_cast<std::uint32_t>(target);
	std::memcpy(entry + 1, &target32, 4);
	entry[5] = 0xC3;
	check("push_ret", true, "entry=hook_chain", profile);
	reset();
	entry[0] = 0xFF;
	entry[1] = 0x25;
	const auto slot = static_cast<std::uint32_t>(reinterpret_cast<std::uintptr_t>(stub));
	std::memcpy(entry + 2, &slot, 4);
	std::memcpy(stub, &target32, 4);
	check("indirect_jump", true, "entry=hook_chain", profile);
	const std::uint32_t invalidSlot = 1;
	std::memcpy(entry + 2, &invalidSlot, 4);
	check("invalid_indirect_slot", false, "unsupported_patch", profile);
	reset();
	WriteProcessMemory(GetCurrentProcess(), kEntryFixture, original.data(), original.size(), nullptr);
	VirtualProtect(kEntryFixture, original.size(), protection, &protection);
	FlushInstructionCache(GetCurrentProcess(), kEntryFixture, original.size());
	VirtualFree(stub, 0, MEM_RELEASE);
	std::cout << "passed=" << passed << " failed=" << failed << '\n';
	return failed ? 1 : 0;
}
