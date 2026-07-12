#include "AutoLinkerUpdateInstaller.h"

#include <fstream>
#include <string>
#include <vector>

namespace {

std::string Utf8FromWide(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (length <= 0) {
		return {};
	}
	std::string result(static_cast<size_t>(length), '\0');
	if (WideCharToMultiByte(
			CP_UTF8,
			0,
			text.data(),
			static_cast<int>(text.size()),
			result.data(),
			length,
			nullptr,
			nullptr) <= 0) {
		return {};
	}
	return result;
}

std::string EscapePowerShellSingleQuoted(const std::string& text)
{
	std::string escaped;
	escaped.reserve(text.size() + 8);
	for (const char ch : text) {
		escaped.push_back(ch);
		if (ch == '\'') {
			escaped.push_back('\'');
		}
	}
	return escaped;
}

std::string PowerShellLiteral(const std::filesystem::path& path)
{
	return "'" + EscapePowerShellSingleQuoted(Utf8FromWide(path.wstring())) + "'";
}

std::wstring QuoteCommandLineArgument(const std::wstring& argument)
{
	std::wstring quoted = L"\"";
	size_t backslashes = 0;
	for (const wchar_t ch : argument) {
		if (ch == L'\\') {
			++backslashes;
			continue;
		}
		if (ch == L'\"') {
			quoted.append(backslashes * 2 + 1, L'\\');
			quoted.push_back(ch);
			backslashes = 0;
			continue;
		}
		quoted.append(backslashes, L'\\');
		backslashes = 0;
		quoted.push_back(ch);
	}
	quoted.append(backslashes * 2, L'\\');
	quoted.push_back(L'\"');
	return quoted;
}

bool WriteUtf8BomFile(const std::filesystem::path& path, const std::string& text, std::string& outError)
{
	std::ofstream output(path, std::ios::binary | std::ios::trunc);
	if (!output.is_open()) {
		outError = "无法创建退出后更新脚本";
		return false;
	}
	static constexpr unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
	output.write(reinterpret_cast<const char*>(bom), sizeof(bom));
	output.write(text.data(), static_cast<std::streamsize>(text.size()));
	if (!output.good()) {
		outError = "写入退出后更新脚本失败";
		return false;
	}
	return true;
}

std::string BuildUpdaterScript(const AutoLinkerUpdateInstallRequest& request)
{
	std::string script;
	script += "$ErrorActionPreference = 'Stop'\r\n";
	script += "$pidToWait = " + std::to_string(request.processId) + "\r\n";
	script += "$sourceFne = " + PowerShellLiteral(request.stagedFne) + "\r\n";
	script += "$targetFne = " + PowerShellLiteral(request.targetFne) + "\r\n";
	script += "$stagingRoot = " + PowerShellLiteral(request.stagingRoot) + "\r\n";
	script += "$logPath = " + PowerShellLiteral(request.logPath) + "\r\n";
	script += "$targetVersion = '" + EscapePowerShellSingleQuoted(request.targetVersion) + "'\r\n";
	script += "function Write-UpdateLog([string]$message) {\r\n";
	script += "  $line = ('[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date), $message)\r\n";
	script += "  Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8\r\n";
	script += "}\r\n";
	script += "try {\r\n";
	script += "  Write-UpdateLog ('Waiting for IDE process ' + $pidToWait + ' to exit.')\r\n";
	script += "  $deadline = (Get-Date).AddMinutes(10)\r\n";
	script += "  while (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) {\r\n";
	script += "    if ((Get-Date) -ge $deadline) { throw 'IDE did not exit within 10 minutes; update cancelled.' }\r\n";
	script += "    Start-Sleep -Milliseconds 500\r\n";
	script += "  }\r\n";
	script += "  $targetDirectory = Split-Path -Parent $targetFne\r\n";
	script += "  $temporaryTarget = Join-Path $targetDirectory 'AutoLinker.fne.update.tmp'\r\n";
	script += "  $backupTarget = Join-Path $targetDirectory 'AutoLinker.fne.update.bak'\r\n";
	script += "  $updated = $false\r\n";
	script += "  for ($attempt = 1; $attempt -le 20 -and -not $updated; $attempt++) {\r\n";
	script += "    try {\r\n";
	script += "      Remove-Item -LiteralPath $temporaryTarget -Force -ErrorAction SilentlyContinue\r\n";
	script += "      Copy-Item -LiteralPath $sourceFne -Destination $temporaryTarget -Force\r\n";
	script += "      if (Test-Path -LiteralPath $targetFne) {\r\n";
	script += "        Remove-Item -LiteralPath $backupTarget -Force -ErrorAction SilentlyContinue\r\n";
	script += "        [System.IO.File]::Replace($temporaryTarget, $targetFne, $backupTarget, $true)\r\n";
	script += "      } else {\r\n";
	script += "        Move-Item -LiteralPath $temporaryTarget -Destination $targetFne -Force\r\n";
	script += "      }\r\n";
	script += "      $updated = $true\r\n";
	script += "    } catch {\r\n";
	script += "      if ($attempt -ge 20) { throw }\r\n";
	script += "      Start-Sleep -Milliseconds 500\r\n";
	script += "    }\r\n";
	script += "  }\r\n";
	script += "  Remove-Item -LiteralPath $backupTarget -Force -ErrorAction SilentlyContinue\r\n";
	script += "  Write-UpdateLog ('AutoLinker updated successfully to ' + $targetVersion + '.')\r\n";
	script += "  Add-Type -AssemblyName System.Windows.Forms\r\n";
	script += "  [System.Windows.Forms.MessageBox]::Show(('AutoLinker 已成功更新到 ' + $targetVersion + '。请重新打开易语言 IDE。'), 'AutoLinker 更新', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null\r\n";
	script += "} catch {\r\n";
	script += "  $failure = $_.Exception.Message\r\n";
	script += "  Write-UpdateLog ('Update failed: ' + $failure)\r\n";
	script += "  Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue\r\n";
	script += "  [System.Windows.Forms.MessageBox]::Show(('AutoLinker 更新失败：' + $failure), 'AutoLinker 更新', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null\r\n";
	script += "} finally {\r\n";
	script += "  Start-Sleep -Milliseconds 500\r\n";
	script += "  Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue\r\n";
	script += "}\r\n";
	return script;
}

} // namespace

bool AutoLinkerUpdateInstaller::Launch(
	const AutoLinkerUpdateInstallRequest& request,
	PROCESS_INFORMATION& outProcess,
	std::string& outError)
{
	const std::filesystem::path scriptPath = request.stagingRoot / L"apply-update.ps1";
	if (!WriteUtf8BomFile(scriptPath, BuildUpdaterScript(request), outError)) {
		return false;
	}

	std::wstring commandLine =
		L"powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ";
	commandLine += QuoteCommandLineArgument(scriptPath.wstring());
	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');

	STARTUPINFOW startupInfo = {};
	startupInfo.cb = sizeof(startupInfo);
	startupInfo.dwFlags = STARTF_USESHOWWINDOW;
	startupInfo.wShowWindow = SW_HIDE;
	outProcess = {};
	if (CreateProcessW(
			nullptr,
			mutableCommandLine.data(),
			nullptr,
			nullptr,
			FALSE,
			CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP,
			nullptr,
			scriptPath.parent_path().c_str(),
			&startupInfo,
			&outProcess) == FALSE) {
		outError = "启动退出后更新器失败，Win32 error=" + std::to_string(GetLastError());
		return false;
	}
	return true;
}
