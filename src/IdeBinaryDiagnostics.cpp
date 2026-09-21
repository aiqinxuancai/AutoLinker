#include "IdeBinaryDiagnostics.h"

#include <Windows.h>
#include <bcrypt.h>

#include <array>
#include <algorithm>
#include <cstring>
#include <format>
#include <vector>

#pragma comment(lib, "bcrypt.lib")

namespace {

std::string HexBytes(std::span<const std::uint8_t> bytes)
{
	std::string text;
	text.reserve(bytes.size() * 2);
	constexpr char digits[] = "0123456789abcdef";
	for (const auto byte : bytes) {
		text += digits[byte >> 4];
		text += digits[byte & 15];
	}
	return text;
}

struct HashResources {
	HANDLE file = INVALID_HANDLE_VALUE;
	BCRYPT_ALG_HANDLE algorithm = nullptr;
	BCRYPT_HASH_HANDLE hash = nullptr;
	~HashResources()
	{
		if (hash != nullptr) BCryptDestroyHash(hash);
		if (algorithm != nullptr) BCryptCloseAlgorithmProvider(algorithm, 0);
		if (file != INVALID_HANDLE_VALUE) CloseHandle(file);
	}
};

std::string BuildIdentity()
{
	std::wstring path(32768, L'\0');
	const DWORD length = GetModuleFileNameW(nullptr, path.data(), static_cast<DWORD>(path.size()));
	if (length == 0 || length >= path.size()) {
		return "ide_identity_error=get_module_path|win32_error=" + std::to_string(GetLastError());
	}
	path.resize(length);
	const int utf8Length = WideCharToMultiByte(CP_UTF8, 0, path.data(), length, nullptr, 0, nullptr, nullptr);
	std::string utf8Path(utf8Length, '\0');
	WideCharToMultiByte(CP_UTF8, 0, path.data(), length, utf8Path.data(), utf8Length, nullptr, nullptr);
	std::string result = "ide_path=" + utf8Path;
	HashResources resources;
	resources.file = CreateFileW(path.c_str(), GENERIC_READ,
		FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING,
		FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
	if (resources.file == INVALID_HANDLE_VALUE) {
		return result + "|ide_identity_error=open_file|win32_error=" + std::to_string(GetLastError());
	}
	LARGE_INTEGER size{};
	if (!GetFileSizeEx(resources.file, &size)) {
		return result + "|ide_identity_error=file_size|win32_error=" + std::to_string(GetLastError());
	}
	result += "|ide_file_size=" + std::to_string(size.QuadPart);
	NTSTATUS status = BCryptOpenAlgorithmProvider(&resources.algorithm, BCRYPT_MD5_ALGORITHM, nullptr, 0);
	if (status >= 0) {
		status = BCryptCreateHash(resources.algorithm, &resources.hash, nullptr, 0, nullptr, 0, 0);
	}
	if (status < 0) {
		return result + std::format("|ide_identity_error=hash_init|ntstatus=0x{:08X}", static_cast<ULONG>(status));
	}
	std::array<std::uint8_t, 65536> buffer{};
	for (;;) {
		DWORD bytes = 0;
		if (!ReadFile(resources.file, buffer.data(), static_cast<DWORD>(buffer.size()), &bytes, nullptr)) {
			return result + "|ide_identity_error=read_file|win32_error=" + std::to_string(GetLastError());
		}
		if (bytes == 0) break;
		status = BCryptHashData(resources.hash, buffer.data(), bytes, 0);
		if (status < 0) {
			return result + std::format("|ide_identity_error=hash_data|ntstatus=0x{:08X}", static_cast<ULONG>(status));
		}
	}
	std::array<std::uint8_t, 16> digest{};
	status = BCryptFinishHash(resources.hash, digest.data(), static_cast<ULONG>(digest.size()), 0);
	if (status < 0) {
		return result + std::format("|ide_identity_error=hash_finish|ntstatus=0x{:08X}", static_cast<ULONG>(status));
	}
	return result + "|ide_md5=" + HexBytes(digest);
}

} // namespace

const std::string& GetIdeBinaryIdentity()
{
	static const std::string identity = BuildIdentity();
	return identity;
}

std::string BuildIdeSignatureDiagnostics(
	std::uintptr_t moduleBase,
	std::uintptr_t rva,
	std::span<const std::uint8_t> expected)
{
	const std::uintptr_t address = moduleBase + rva;
	std::vector<std::uint8_t> actual(expected.size());
	SIZE_T bytesRead = 0;
	const BOOL ok = ReadProcessMemory(GetCurrentProcess(), reinterpret_cast<const void*>(address),
		actual.data(), actual.size(), &bytesRead);
	const DWORD error = ok ? ERROR_SUCCESS : GetLastError();
	actual.resize(bytesRead);
	// MCP 的旧 trace 使用本地编码；这里只携带 ASCII 指纹，UTF-8 路径留在启动日志。
	const std::string& identity = GetIdeBinaryIdentity();
	const size_t md5Offset = identity.rfind("|ide_md5=");
	const std::string fingerprint = md5Offset == std::string::npos
		? "|ide_md5=unavailable" : identity.substr(md5Offset);
	return std::format("module_base=0x{:X}|signature_rva=0x{:X}|signature_address=0x{:X}",
		moduleBase, rva, address) +
		"|expected_bytes=" + HexBytes(expected) + "|actual_bytes=" + HexBytes(actual) +
		"|signature_read_ok=" + std::to_string(ok != FALSE) +
		"|signature_bytes_read=" + std::to_string(bytesRead) +
		"|signature_read_error=" + std::to_string(error) + fingerprint;
}

namespace {

bool ReadMemory(std::uintptr_t address, void* output, size_t size)
{
	SIZE_T read = 0;
	return ReadProcessMemory(GetCurrentProcess(), reinterpret_cast<const void*>(address),
		output, size, &read) && read == size;
}

bool ReadFileAt(HANDLE file, std::uint64_t offset, void* output, DWORD size)
{
	LARGE_INTEGER position{};
	position.QuadPart = static_cast<LONGLONG>(offset);
	DWORD read = 0;
	return SetFilePointerEx(file, position, nullptr, FILE_BEGIN) &&
		ReadFile(file, output, size, &read, nullptr) && read == size;
}

bool MatchProfileHeaders(const IMAGE_NT_HEADERS32& nt, const IdeEntryProfile& profile)
{
	return nt.Signature == IMAGE_NT_SIGNATURE && nt.FileHeader.Machine == IMAGE_FILE_MACHINE_I386 &&
		nt.OptionalHeader.Magic == IMAGE_NT_OPTIONAL_HDR32_MAGIC &&
		nt.FileHeader.SizeOfOptionalHeader == sizeof(IMAGE_OPTIONAL_HEADER32) &&
		nt.FileHeader.TimeDateStamp == profile.timeDateStamp &&
		nt.OptionalHeader.SizeOfImage == profile.sizeOfImage;
}

bool MatchDiskEntry(std::uintptr_t moduleBase, const IdeEntryProfile& profile)
{
	std::wstring path(32768, L'\0');
	const DWORD length = GetModuleFileNameW(reinterpret_cast<HMODULE>(moduleBase),
		path.data(), static_cast<DWORD>(path.size()));
	if (length == 0 || length >= path.size()) return false;
	path.resize(length);
	HashResources resources;
	resources.file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_DELETE,
		nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
	if (resources.file == INVALID_HANDLE_VALUE) return false;
	IMAGE_DOS_HEADER dos{};
	IMAGE_NT_HEADERS32 nt{};
	if (!ReadFileAt(resources.file, 0, &dos, sizeof(dos)) || dos.e_magic != IMAGE_DOS_SIGNATURE ||
		dos.e_lfanew <= 0 || !ReadFileAt(resources.file, dos.e_lfanew, &nt, sizeof(nt)) ||
		!MatchProfileHeaders(nt, profile) || nt.FileHeader.NumberOfSections > 96) return false;
	// 此档案使用未重定位的指令字节；不把重定位差异误判为入口 Hook。
	if (moduleBase != nt.OptionalHeader.ImageBase) return false;
	for (unsigned i = 0; i < nt.FileHeader.NumberOfSections; ++i) {
		IMAGE_SECTION_HEADER section{};
		const auto offset = static_cast<std::uint64_t>(dos.e_lfanew) + sizeof(nt) + i * sizeof(section);
		if (!ReadFileAt(resources.file, offset, &section, sizeof(section))) return false;
		if (profile.rva < section.VirtualAddress) continue;
		const auto delta = profile.rva - section.VirtualAddress;
		if (delta > section.SizeOfRawData || profile.expected.size() > section.SizeOfRawData - delta) continue;
		if (!(section.Characteristics & IMAGE_SCN_MEM_EXECUTE)) return false;
		std::vector<std::uint8_t> bytes(profile.expected.size());
		return ReadFileAt(resources.file, static_cast<std::uint64_t>(section.PointerToRawData) + delta,
			bytes.data(), static_cast<DWORD>(bytes.size())) &&
			std::equal(bytes.begin(), bytes.end(), profile.expected.begin());
	}
	return false;
}

bool IsExecutable(std::uintptr_t address)
{
	MEMORY_BASIC_INFORMATION info{};
	if (!VirtualQuery(reinterpret_cast<void*>(address), &info, sizeof(info)) ||
		info.State != MEM_COMMIT || (info.Protect & (PAGE_GUARD | PAGE_NOACCESS))) return false;
	const DWORD access = info.Protect & 0xFF;
	return access == PAGE_EXECUTE || access == PAGE_EXECUTE_READ ||
		access == PAGE_EXECUTE_READWRITE || access == PAGE_EXECUTE_WRITECOPY;
}

// 返回 0 表示普通指令，-1 表示无法安全解码，正数表示跳转指令长度。
int DecodeEntryJump(std::uintptr_t address, std::uintptr_t& target)
{
	std::array<std::uint8_t, 6> bytes{};
	if (!ReadMemory(address, bytes.data(), 2)) return -1;
	if (bytes[0] == 0xE9) {
		if (!ReadMemory(address, bytes.data(), 5)) return -1;
		std::int32_t displacement = 0;
		std::memcpy(&displacement, bytes.data() + 1, sizeof(displacement));
		target = static_cast<std::uint32_t>(address + 5 + displacement);
		return 5;
	}
	if (bytes[0] == 0xEB) {
		target = static_cast<std::uint32_t>(address + 2 + static_cast<std::int8_t>(bytes[1]));
		return 2;
	}
	if (bytes[0] == 0xFF && bytes[1] == 0x25) {
		if (!ReadMemory(address, bytes.data(), 6)) return -1;
		std::uint32_t slot = 0, value = 0;
		std::memcpy(&slot, bytes.data() + 2, sizeof(slot));
		if (!ReadMemory(slot, &value, sizeof(value))) return -1;
		target = value;
		return 6;
	}
	if (bytes[0] == 0x68) {
		if (!ReadMemory(address, bytes.data(), 6)) return -1;
		if (bytes[5] != 0xC3) return 0;
		std::uint32_t value = 0;
		std::memcpy(&value, bytes.data() + 1, sizeof(value));
		target = value;
		return 6;
	}
	return 0;
}

bool ValidateJumpChain(std::uintptr_t entry, std::uintptr_t moduleBase, size_t patchWindow,
	std::span<const std::uint8_t> actual, std::span<const std::uint8_t> expected, std::string& trace)
{
	std::array<std::uintptr_t, 9> visited{};
	std::uintptr_t current = entry;
	for (size_t depth = 0; depth < visited.size(); ++depth) {
		if (std::find(visited.begin(), visited.begin() + depth, current) != visited.begin() + depth) {
			trace += "|entry_rejected=jump_cycle";
			return false;
		}
		visited[depth] = current;
		if (!IsExecutable(current)) {
			trace += "|entry_rejected=target_not_executable";
			return false;
		}
		std::uintptr_t target = 0;
		const int length = DecodeEntryJump(current, target);
		if (length < 0 || (depth == 0 && (length == 0 || static_cast<size_t>(length) > patchWindow))) {
			trace += "|entry_rejected=unsupported_patch";
			return false;
		}
		if (depth == 0) {
			for (size_t i = length; i < patchWindow; ++i) {
				if (actual[i] != expected[i] && actual[i] != 0x90 && actual[i] != 0xCC) {
					trace += "|entry_rejected=unexpected_padding";
					return false;
				}
			}
		}
		if (length == 0) {
			HMODULE owner = nullptr;
			if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
				reinterpret_cast<LPCWSTR>(current), &owner) || reinterpret_cast<std::uintptr_t>(owner) == moduleBase) {
				trace += "|entry_rejected=unverified_target_owner";
				return false;
			}
			std::array<wchar_t, 32768> path{};
			const DWORD size = GetModuleFileNameW(owner, path.data(), static_cast<DWORD>(path.size()));
			if (size == 0 || size >= path.size()) return false;
			const int count = WideCharToMultiByte(CP_UTF8, 0, path.data(), size, nullptr, 0, nullptr, nullptr);
			std::string utf8(count, '\0');
			WideCharToMultiByte(CP_UTF8, 0, path.data(), size, utf8.data(), count, nullptr, nullptr);
			trace += std::format("|entry=hook_chain|hook_target=0x{:X}|hook_module=", current) + utf8;
			return true;
		}
		trace += std::format("|jump=0x{:X}->0x{:X}", current, target);
		current = target;
	}
	trace += "|entry_rejected=jump_depth_limit";
	return false;
}

} // namespace

bool ValidateIdeCallableEntry(std::uintptr_t moduleBase, const IdeEntryProfile& profile, std::string& trace)
{
	trace = "ide_entry_probe";
	if (sizeof(void*) != 4 || profile.sizeOfImage < sizeof(IMAGE_NT_HEADERS32) ||
		profile.expected.size() < 32 || profile.expected.size() > 256 ||
		profile.patchWindow < 5 || profile.patchWindow > 16 || profile.patchWindow >= profile.expected.size() ||
		profile.rva >= profile.sizeOfImage || profile.expected.size() > profile.sizeOfImage - profile.rva) {
		trace += "|entry_rejected=invalid_profile";
		return false;
	}
	IMAGE_DOS_HEADER dos{};
	IMAGE_NT_HEADERS32 nt{};
	if (!ReadMemory(moduleBase, &dos, sizeof(dos)) || dos.e_magic != IMAGE_DOS_SIGNATURE ||
		dos.e_lfanew <= 0 || static_cast<std::uint32_t>(dos.e_lfanew) > profile.sizeOfImage - sizeof(nt) ||
		!ReadMemory(moduleBase + dos.e_lfanew, &nt, sizeof(nt)) || !MatchProfileHeaders(nt, profile) ||
		!MatchDiskEntry(moduleBase, profile)) {
		trace += "|entry_rejected=version_or_disk_signature";
		return false;
	}
	trace += "|disk_profile=matched";
	const auto entry = moduleBase + profile.rva;
	std::vector<std::uint8_t> actual(profile.expected.size());
	if (!IsExecutable(entry) || !ReadMemory(entry, actual.data(), actual.size())) {
		trace += "|entry_rejected=unreadable_or_non_executable";
		return false;
	}
	if (std::equal(actual.begin(), actual.end(), profile.expected.begin())) {
		trace += "|entry=original";
		return true;
	}
	if (!std::equal(actual.begin() + profile.patchWindow, actual.end(), profile.expected.begin() + profile.patchWindow)) {
		trace += "|entry_rejected=body_changed";
		return false;
	}
	return ValidateJumpChain(entry, moduleBase, profile.patchWindow, actual, profile.expected, trace);
}
