#include "IdeBinaryDiagnostics.h"

#include <Windows.h>
#include <bcrypt.h>

#include <array>
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
