# 编译真实镜像实现并模拟源码变化回调；超时视为死锁回归。
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (!$installation) { throw 'Visual C++ installation not found' }
$vcvars = Join-Path $installation 'VC\Auxiliary\Build\vcvars32.bat'
$output = Join-Path $repo '.tmp\workspace-mirror-reentry'
$null = New-Item -ItemType Directory -Force -Path $output
$buildFile = Join-Path $output 'build.cmd'
$testSource = Join-Path $repo 'AutoLinkerTest\WorkspaceMirrorReentryTests.cpp'
$mirrorSource = Join-Path $repo 'src\WorkspaceMirror.cpp'
@"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b 1
cl /nologo /EHsc /std:c++20 /utf-8 /MD /DWIN32 /D_WINDOWS "$testSource" "$mirrorSource" /Fe:WorkspaceMirrorReentryTests.exe /link user32.lib
exit /b %errorlevel%
"@ | Set-Content -LiteralPath $buildFile -Encoding Default
Push-Location $output
try {
    & $env:ComSpec /d /c $buildFile
    if ($LASTEXITCODE -ne 0) { throw 'Regression test build failed' }
    $process = Start-Process -FilePath (Join-Path $output 'WorkspaceMirrorReentryTests.exe') -WorkingDirectory $output -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $output 'stdout.log') -RedirectStandardError (Join-Path $output 'stderr.log')
    if (!$process.WaitForExit(30000)) {
        $process.Kill()
        $process.WaitForExit()
        throw 'Regression test timed out: possible mirror deadlock'
    }
    Get-Content -LiteralPath (Join-Path $output 'stdout.log')
    Get-Content -LiteralPath (Join-Path $output 'stderr.log')
    if ($process.ExitCode -ne 0) { throw "Regression test failed: $($process.ExitCode)" }
} finally {
    Pop-Location
}
