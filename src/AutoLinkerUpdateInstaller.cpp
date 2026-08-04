#include "AutoLinkerUpdateInstaller.h"

#include <filesystem>
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

std::string LocalFromWide(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(
		CP_ACP,
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
		CP_ACP,
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

bool WriteLocalTextFile(const std::filesystem::path& path, const std::string& text, std::string& outError)
{
	std::ofstream output(path, std::ios::binary | std::ios::trunc);
	if (!output.is_open()) {
		outError = "无法创建退出后 CMD 更新脚本";
		return false;
	}
	output.write(text.data(), static_cast<std::streamsize>(text.size()));
	if (!output.good()) {
		outError = "写入退出后 CMD 更新脚本失败";
		return false;
	}
	return true;
}

std::string EscapeCmdSetValue(std::string text)
{
	std::string escaped;
	escaped.reserve(text.size() + 16);
	for (const char ch : text) {
		switch (ch) {
		case '%':
			escaped += "%%";
			break;
		case '^':
		case '&':
		case '|':
		case '<':
		case '>':
		case '(':
		case ')':
			escaped.push_back('^');
			escaped.push_back(ch);
			break;
		default:
			escaped.push_back(ch);
			break;
		}
	}
	return escaped;
}

std::string BuildUpdaterScript(const AutoLinkerUpdateInstallRequest& request)
{
	std::string script;
	script += "$ErrorActionPreference = 'Stop'\r\n";
	script += "$updateCompleted = $false\r\n";
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
	script += "  $updateCompleted = $true\r\n";
	script += "  Add-Type -AssemblyName System.Windows.Forms\r\n";
	script += "  [System.Windows.Forms.MessageBox]::Show(('AutoLinker 已成功更新到 ' + $targetVersion + '。请重新打开易语言 IDE。'), 'AutoLinker 更新', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null\r\n";
	script += "} catch {\r\n";
	script += "  $failure = $_.Exception.Message\r\n";
	script += "  Write-UpdateLog ('PowerShell update failed; CMD fallback will be attempted: ' + $failure)\r\n";
	script += "  if ($updateCompleted) { exit 0 }\r\n";
	script += "  exit 1\r\n";
	script += "} finally {\r\n";
	script += "  if ($updateCompleted) {\r\n";
	script += "    Start-Sleep -Milliseconds 500\r\n";
	script += "    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue\r\n";
	script += "  }\r\n";
	script += "}\r\n";
	return script;
}

std::string BuildCmdFallbackScript(const AutoLinkerUpdateInstallRequest& request)
{
	const std::filesystem::path targetDirectory = request.targetFne.parent_path();
	const std::filesystem::path temporaryTarget = targetDirectory / L"AutoLinker.fne.update.tmp";
	const std::filesystem::path backupTarget = targetDirectory / L"AutoLinker.fne.update.bak";
	const auto setPath = [](const std::filesystem::path& path) {
		return EscapeCmdSetValue(LocalFromWide(path.wstring()));
	};

	std::string script;
	script += "@echo off\r\n";
	script += "setlocal EnableExtensions DisableDelayedExpansion\r\n";
	script += "set \"pidToWait=" + std::to_string(request.processId) + "\"\r\n";
	script += "set \"sourceFne=" + setPath(request.stagedFne) + "\"\r\n";
	script += "set \"targetFne=" + setPath(request.targetFne) + "\"\r\n";
	script += "set \"temporaryTarget=" + setPath(temporaryTarget) + "\"\r\n";
	script += "set \"backupTarget=" + setPath(backupTarget) + "\"\r\n";
	script += "set \"logPath=" + setPath(request.logPath) + "\"\r\n";
	script += "set /a waitCount=0\r\n";
	script += "call :log CMD fallback updater started.\r\n";
	script += ":waitForIde\r\n";
	script += "\"%SystemRoot%\\System32\\tasklist.exe\" /FI \"PID eq %pidToWait%\" /NH 2>nul | \"%SystemRoot%\\System32\\findstr.exe\" /R /C:\"[ ]%pidToWait%[ ]\" >nul\r\n";
	script += "if errorlevel 1 goto ideExited\r\n";
	script += "set /a waitCount+=1\r\n";
	script += "if %waitCount% GEQ 1200 goto waitTimedOut\r\n";
	script += "\"%SystemRoot%\\System32\\timeout.exe\" /t 1 /nobreak >nul 2>&1\r\n";
	script += "goto waitForIde\r\n";
	script += ":waitTimedOut\r\n";
	script += "call :log IDE did not exit within 20 minutes; CMD fallback cancelled.\r\n";
	script += "exit /b 1\r\n";
	script += ":ideExited\r\n";
	script += "if not exist \"%sourceFne%\" goto sourceMissing\r\n";
	script += "set /a attempt=0\r\n";
	script += ":retryUpdate\r\n";
	script += "set /a attempt+=1\r\n";
	script += "del /f /q \"%temporaryTarget%\" >nul 2>&1\r\n";
	script += "copy /b /y \"%sourceFne%\" \"%temporaryTarget%\" >nul 2>&1\r\n";
	script += "if errorlevel 1 goto retryDelay\r\n";
	script += "if not exist \"%targetFne%\" goto installNewFile\r\n";
	script += "del /f /q \"%backupTarget%\" >nul 2>&1\r\n";
	script += "move /y \"%targetFne%\" \"%backupTarget%\" >nul 2>&1\r\n";
	script += "if errorlevel 1 goto retryDelay\r\n";
	script += ":installNewFile\r\n";
	script += "move /y \"%temporaryTarget%\" \"%targetFne%\" >nul 2>&1\r\n";
	script += "if errorlevel 1 goto restoreAndRetry\r\n";
	script += "if not exist \"%targetFne%\" goto restoreAndRetry\r\n";
	script += "del /f /q \"%backupTarget%\" >nul 2>&1\r\n";
	script += "call :log AutoLinker updated successfully by CMD fallback.\r\n";
	script += "exit /b 0\r\n";
	script += ":restoreAndRetry\r\n";
	script += "if exist \"%backupTarget%\" (\r\n";
	script += "  del /f /q \"%targetFne%\" >nul 2>&1\r\n";
	script += "  move /y \"%backupTarget%\" \"%targetFne%\" >nul 2>&1\r\n";
	script += ")\r\n";
	script += ":retryDelay\r\n";
	script += "if %attempt% GEQ 20 goto updateFailed\r\n";
	script += "\"%SystemRoot%\\System32\\timeout.exe\" /t 1 /nobreak >nul 2>&1\r\n";
	script += "goto retryUpdate\r\n";
	script += ":sourceMissing\r\n";
	script += "call :log Staged AutoLinker.fne is missing; manual update is required.\r\n";
	script += "exit /b 1\r\n";
	script += ":updateFailed\r\n";
	script += "if exist \"%backupTarget%\" (\r\n";
	script += "  del /f /q \"%targetFne%\" >nul 2>&1\r\n";
	script += "  move /y \"%backupTarget%\" \"%targetFne%\" >nul 2>&1\r\n";
	script += ")\r\n";
	script += "call :log CMD fallback failed after 20 attempts; original file was restored when possible.\r\n";
	script += "exit /b 1\r\n";
	script += ":log\r\n";
	script += ">>\"%logPath%\" echo [%date% %time%] %*\r\n";
	script += "exit /b 0\r\n";
	return script;
}

std::string BuildBootstrapScript()
{
	std::string script;
	script += "@echo off\r\n";
	script += "setlocal EnableExtensions DisableDelayedExpansion\r\n";
	script += "set \"powerShellPath=%SystemRoot%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe\"\r\n";
	script += "if not exist \"%powerShellPath%\" set \"powerShellPath=%SystemRoot%\\Sysnative\\WindowsPowerShell\\v1.0\\powershell.exe\"\r\n";
	script += "if not exist \"%powerShellPath%\" goto cmdFallback\r\n";
	script += "\"%powerShellPath%\" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"%~dp0apply-update.ps1\"\r\n";
	script += "if not errorlevel 1 exit /b 0\r\n";
	script += ":cmdFallback\r\n";
	script += "call \"%~dp0apply-update-fallback.cmd\"\r\n";
	script += "exit /b %errorlevel%\r\n";
	return script;
}

std::filesystem::path ResolveCmdExecutable()
{
	std::vector<wchar_t> buffer(MAX_PATH);
	for (;;) {
		const UINT length = GetSystemDirectoryW(buffer.data(), static_cast<UINT>(buffer.size()));
		if (length == 0) {
			return {};
		}
		if (length < buffer.size()) {
			const std::filesystem::path candidate =
				std::filesystem::path(std::wstring(buffer.data(), length)) / L"cmd.exe";
			std::error_code ec;
			return std::filesystem::is_regular_file(candidate, ec) && !ec ? candidate : std::filesystem::path();
		}
		buffer.resize(static_cast<size_t>(length) + 1);
	}
}

} // namespace

bool AutoLinkerUpdateInstaller::Launch(
	const AutoLinkerUpdateInstallRequest& request,
	PROCESS_INFORMATION& outProcess,
	std::string& outError)
{
	const std::filesystem::path powerShellScriptPath = request.stagingRoot / L"apply-update.ps1";
	const std::filesystem::path fallbackScriptPath = request.stagingRoot / L"apply-update-fallback.cmd";
	const std::filesystem::path bootstrapScriptPath = request.stagingRoot / L"apply-update-bootstrap.cmd";
	if (!WriteUtf8BomFile(powerShellScriptPath, BuildUpdaterScript(request), outError) ||
		!WriteLocalTextFile(fallbackScriptPath, BuildCmdFallbackScript(request), outError) ||
		!WriteLocalTextFile(bootstrapScriptPath, BuildBootstrapScript(), outError)) {
		return false;
	}

	const std::filesystem::path cmdPath = ResolveCmdExecutable();
	if (cmdPath.empty()) {
		outError = "未找到系统 cmd.exe，无法启动退出后更新器；请关闭 IDE 后手动替换 AutoLinker.fne";
		return false;
	}
	std::wstring commandLine = QuoteCommandLineArgument(cmdPath.wstring());
	commandLine += L" /d /c call ";
	commandLine += QuoteCommandLineArgument(bootstrapScriptPath.wstring());
	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');

	STARTUPINFOW startupInfo = {};
	startupInfo.cb = sizeof(startupInfo);
	startupInfo.dwFlags = STARTF_USESHOWWINDOW;
	startupInfo.wShowWindow = SW_HIDE;
	outProcess = {};
	if (CreateProcessW(
			cmdPath.c_str(),
			mutableCommandLine.data(),
			nullptr,
			nullptr,
			FALSE,
			CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP,
			nullptr,
			bootstrapScriptPath.parent_path().c_str(),
			&startupInfo,
			&outProcess) == FALSE) {
		outError = "启动退出后更新器失败，Win32 error=" + std::to_string(GetLastError());
		return false;
	}
	return true;
}
