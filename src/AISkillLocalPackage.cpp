#include "AISkillLocalPackage.h"

#include <Windows.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <cwctype>
#include <fstream>
#include <format>
#include <system_error>

#include "PowerShellToolRunner.h"

namespace {

constexpr unsigned long long kMaxArchiveBytes = 100ULL * 1024 * 1024;
constexpr unsigned long long kMaxExpandedBytes = 512ULL * 1024 * 1024;
constexpr unsigned long long kMaxArchiveEntries = 10000;

std::atomic_ullong g_localPackageCounter = 1;

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int size = WideCharToMultiByte(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return {};
	}
	std::string result(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr) <= 0) {
		return {};
	}
	return result;
}

std::string PathToUtf8(const std::filesystem::path& path)
{
	return WideToUtf8(path.wstring());
}

std::filesystem::path CanonicalOrAbsolute(const std::filesystem::path& path, std::error_code& ec)
{
	ec.clear();
	auto normalized = std::filesystem::weakly_canonical(path, ec);
	if (!ec) {
		return normalized;
	}
	ec.clear();
	normalized = std::filesystem::absolute(path, ec);
	return normalized.lexically_normal();
}

bool IsPathInside(const std::filesystem::path& candidate, const std::filesystem::path& root)
{
	std::error_code ec;
	const auto normalizedCandidate = CanonicalOrAbsolute(candidate, ec);
	if (ec) {
		return false;
	}
	const auto normalizedRoot = CanonicalOrAbsolute(root, ec);
	if (ec) {
		return false;
	}
	auto candidateIt = normalizedCandidate.begin();
	for (auto rootIt = normalizedRoot.begin(); rootIt != normalizedRoot.end(); ++rootIt, ++candidateIt) {
		if (candidateIt == normalizedCandidate.end() || _wcsicmp(rootIt->c_str(), candidateIt->c_str()) != 0) {
			return false;
		}
	}
	return true;
}

std::string PowerShellLiteral(const std::filesystem::path& path)
{
	const std::string value = PathToUtf8(path);
	std::string escaped;
	escaped.reserve(value.size() + 8);
	for (const char ch : value) {
		escaped.push_back(ch);
		if (ch == '\'') {
			escaped.push_back('\'');
		}
	}
	return "'" + escaped + "'";
}

std::string DescribePowerShellFailure(const PowerShellRunResult& result, const std::string& fallback)
{
	if (!result.stdErr.empty()) {
		return result.stdErr;
	}
	if (!result.stdOut.empty()) {
		return result.stdOut;
	}
	if (!result.error.empty()) {
		return result.error;
	}
	return fallback;
}

bool HasZipSignature(const std::filesystem::path& path)
{
	std::ifstream input(path, std::ios::binary);
	char signature[2] = {};
	return input.is_open() && input.read(signature, sizeof(signature)) && signature[0] == 'P' && signature[1] == 'K';
}

bool PreflightZip(const std::filesystem::path& archive, const std::filesystem::path& workingDirectory, std::string& outError)
{
	const std::string command =
		"Add-Type -AssemblyName System.IO.Compression.FileSystem; "
		"$zip=[System.IO.Compression.ZipFile]::OpenRead(" + PowerShellLiteral(archive) + "); "
		"try { "
		"if ($zip.Entries.Count -gt " + std::to_string(kMaxArchiveEntries) + ") { throw 'ZIP_ENTRY_LIMIT' }; "
		"[Int64]$total=0; "
		"foreach ($entry in $zip.Entries) { "
		"$name=$entry.FullName.Replace('\\','/'); "
		"if ([System.IO.Path]::IsPathRooted($name) -or $name.StartsWith('/') -or $name.StartsWith('//') -or $name -match '^[A-Za-z]:') { throw 'ZIP_ABSOLUTE_PATH' }; "
		"foreach ($part in $name.Split('/')) { if ($part -match '^\\.\\.[ .]*$') { throw 'ZIP_TRAVERSAL' } }; "
		"if ($entry.Length -gt " + std::to_string(kMaxExpandedBytes) + " -or $total -gt (" +
		std::to_string(kMaxExpandedBytes) + "-$entry.Length)) { throw 'ZIP_EXPANDED_LIMIT' }; "
		"$total += $entry.Length "
		"} "
		"} finally { $zip.Dispose() }";
	const PowerShellRunResult result = PowerShellToolRunner::Run(command, PathToUtf8(workingDirectory), 60);
	if (result.ok) {
		return true;
	}
	const std::string detail = DescribePowerShellFailure(result, "ZIP 预检失败");
	if (detail.find("ZIP_ENTRY_LIMIT") != std::string::npos) {
		outError = "ZIP 文件条目超过 10000 个限制";
	}
	else if (detail.find("ZIP_EXPANDED_LIMIT") != std::string::npos) {
		outError = "ZIP 声明的解压后总大小超过 512 MiB 限制";
	}
	else if (detail.find("ZIP_ABSOLUTE_PATH") != std::string::npos ||
		detail.find("ZIP_TRAVERSAL") != std::string::npos) {
		outError = "ZIP 包含绝对路径或上级目录穿越条目";
	}
	else {
		outError = "无法检查 ZIP 文件：" + detail;
	}
	return false;
}

bool ExtractZip(
	const std::filesystem::path& archive,
	const std::filesystem::path& destination,
	std::string& outError)
{
	const std::string command = "Expand-Archive -LiteralPath " + PowerShellLiteral(archive) +
		" -DestinationPath " + PowerShellLiteral(destination) + " -Force";
	const PowerShellRunResult result = PowerShellToolRunner::Run(
		command, PathToUtf8(destination.parent_path()), 120);
	if (!result.ok) {
		outError = "解压 ZIP 失败：" + DescribePowerShellFailure(result, "未知错误");
		return false;
	}
	return true;
}

bool IsZipExtension(const std::filesystem::path& path)
{
	std::wstring extension = path.extension().wstring();
	std::transform(extension.begin(), extension.end(), extension.begin(), [](const wchar_t ch) {
		return static_cast<wchar_t>(std::towlower(ch));
	});
	return extension == L".zip";
}

} // namespace

namespace AISkillLocalPackage {

bool Prepare(
	const std::filesystem::path& sourcePath,
	const std::filesystem::path& cacheRoot,
	AISkillPreparedLocalSource& outSource,
	std::string& outError)
{
	outSource = {};
	outError.clear();
	if (sourcePath.empty()) {
		outError = "请选择本地技能目录、SKILL.md 或 ZIP 文件";
		return false;
	}
	std::error_code ec;
	const auto normalized = CanonicalOrAbsolute(sourcePath, ec);
	if (ec || !std::filesystem::exists(normalized, ec)) {
		outError = "本地来源不存在：" + PathToUtf8(sourcePath);
		return false;
	}
	outSource.sourcePath = normalized;
	if (std::filesystem::is_directory(normalized, ec)) {
		outSource.kind = AISkillLocalSourceKind::Directory;
		outSource.scanRoot = normalized;
		return true;
	}
	if (!std::filesystem::is_regular_file(normalized, ec)) {
		outError = "本地来源必须是目录、SKILL.md 或 ZIP 文件";
		return false;
	}
	if (_wcsicmp(normalized.filename().c_str(), L"SKILL.md") == 0) {
		outSource.kind = AISkillLocalSourceKind::SkillFile;
		outSource.scanRoot = normalized.parent_path();
		outSource.directSkillFile = normalized;
		return true;
	}
	if (!IsZipExtension(normalized)) {
		outError = "仅支持目录、SKILL.md 或 .zip 文件";
		return false;
	}
	const auto archiveSize = std::filesystem::file_size(normalized, ec);
	if (ec) {
		outError = "无法读取 ZIP 文件大小：" + ec.message();
		return false;
	}
	if (archiveSize > kMaxArchiveBytes) {
		outError = "ZIP 文件超过 100 MiB 限制";
		return false;
	}
	if (!HasZipSignature(normalized)) {
		outError = "文件扩展名为 .zip，但缺少有效的 PK 签名";
		return false;
	}
	outSource.kind = AISkillLocalSourceKind::ZipArchive;
	outSource.stagingRoot = cacheRoot /
		std::format(L"SkillLocal.{}.{}.{}", GetCurrentProcessId(), GetTickCount64(), g_localPackageCounter.fetch_add(1));
	outSource.scanRoot = outSource.stagingRoot / "extract";
	std::filesystem::create_directories(outSource.scanRoot, ec);
	if (ec) {
		outError = "创建 ZIP 临时目录失败：" + ec.message();
		Cleanup(outSource, cacheRoot);
		return false;
	}
	if (!PreflightZip(normalized, outSource.stagingRoot, outError) ||
		!ExtractZip(normalized, outSource.scanRoot, outError)) {
		Cleanup(outSource, cacheRoot);
		return false;
	}
	return true;
}

void Cleanup(AISkillPreparedLocalSource& source, const std::filesystem::path& cacheRoot)
{
	if (source.stagingRoot.empty()) {
		return;
	}
	std::error_code ec;
	const auto target = CanonicalOrAbsolute(source.stagingRoot, ec);
	if (!ec && IsPathInside(target, cacheRoot)) {
		std::filesystem::remove_all(target, ec);
	}
	source.stagingRoot.clear();
	source.scanRoot.clear();
}

const char* KindName(const AISkillLocalSourceKind kind)
{
	switch (kind) {
	case AISkillLocalSourceKind::Directory:
		return "directory";
	case AISkillLocalSourceKind::SkillFile:
		return "skill_file";
	case AISkillLocalSourceKind::ZipArchive:
		return "zip";
	default:
		return "unknown";
	}
}

} // namespace AISkillLocalPackage
