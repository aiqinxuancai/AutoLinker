param(
    [string]$IdePath = 'C:\Users\aiqin\OneDrive\e5.6\e5.95.exe',
    [string]$ConfigPath = 'C:\ProgramData\BlackMoon\Config.json'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$launcher = Join-Path $repo 'bin\fne_release\AutoLinkerTest.exe'
$root = Join-Path $repo ('.tmp\blackmoon-regression-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$baseline = (Get-FileHash -LiteralPath $ConfigPath).Hash
$source = Join-Path $root 'console.e'
$otherSource = Join-Path $root 'window.e'
Copy-Item -LiteralPath (Join-Path $repo 'eproj\e-console-exe-new-proj.e') -Destination $source
Copy-Item -LiteralPath (Join-Path $repo 'eproj\e-window-exe-new-proj.e') -Destination $otherSource
$results = [Collections.Generic.List[object]]::new()

function Read-Case([string]$name, [string]$expectedError = '') {
    $result = Get-Content -LiteralPath (Join-Path $root "$name.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $compile = $result.compile_result
    if ($expectedError) {
        $passed = !$result.ok -and $compile.error -eq $expectedError -and [bool]$compile.hint
    } else {
        $passed = $result.ok -and $compile.artifact_verified -and $compile.output_file_modified_after_compile
        if ($compile.compiler -eq 'blackmoon') {
            $passed = $passed -and $compile.config_restored -and $compile.output_path_injected -and ($null -eq $compile.static_compile)
            $profiles = @{asm='汇编模式'; cxx='C/C++模式'; vcxx='VC++模式'}
            $passed = $passed -and $compile.output_window_text.Contains('当前配置: ' + $profiles[$compile.blackmoon_mode])
        }
    }
    $results.Add([pscustomobject]@{name=$name; passed=[bool]$passed; error=$compile.error; hint=$compile.hint; compiler=$compile.compiler; mode=$compile.blackmoon_mode})
    Write-Output ("{0}: passed={1}, error={2}" -f $name, $passed, $compile.error)
}

function Invoke-Case([string]$name, [string[]]$options, [string]$expectedError = '') {
    $output = Join-Path $root "$name.exe"
    $arguments = @('headless-compile', $IdePath, $source, $output, '--result', (Join-Path $root "$name.json"), '--timeout', '180') + $options
    & $launcher @arguments *> (Join-Path $root "$name.log")
    Read-Case $name $expectedError
    if ((Get-FileHash -LiteralPath $ConfigPath).Hash -ne $baseline) { throw "Configuration not restored after $name; evidence: $root" }
}

& $launcher blackmoon-self-test > (Join-Path $root 'self-test.json')
if ($LASTEXITCODE -ne 0) { throw "Self-test failed: $root" }
foreach ($mode in @('asm','cxx','vcxx')) { Invoke-Case $mode @('--blackmoon', $mode) }
Invoke-Case 'default-vcxx' @('--compiler', 'blackmoon')
Invoke-Case 'ordinary' @('--static')
Invoke-Case 'stray-mode' @('--blackmoon','asm','--compiler','eide') 'blackmoon_mode_requires_blackmoon'
Invoke-Case 'blackmoon-static' @('--blackmoon','vcxx','--static') 'blackmoon_static_compile_not_supported'
Invoke-Case 'wrong-mode' @('--blackmoon','invalid') 'unsupported_blackmoon_mode'
Invoke-Case 'wrong-compiler' @('--compiler','invalid') 'unsupported_compiler'
Invoke-Case 'wrong-target' @('--blackmoon','vcxx','--target','win_dll') 'blackmoon_target_override_not_supported'

# 不同源文件绕过启动器的同项目互斥，验证共享黑月配置下的跨 IDE 并发。
$jobs = @()
foreach ($case in @(@('parallel-asm', $source, 'asm'), @('parallel-vcxx', $otherSource, 'vcxx'))) {
    $name = $case[0]
    $parts = @('headless-compile', $IdePath, $case[1], (Join-Path $root "$name.exe"), '--blackmoon', $case[2], '--result', (Join-Path $root "$name.json"), '--timeout', '180')
    $commandLine = ($parts | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $jobs += Start-Process -FilePath $launcher -ArgumentList $commandLine -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root "$name.log") -RedirectStandardError (Join-Path $root "$name.stderr.log")
}
foreach ($job in $jobs) { $job.WaitForExit() }
Read-Case 'parallel-asm'
Read-Case 'parallel-vcxx'
$restored = (Get-FileHash -LiteralPath $ConfigPath).Hash -eq $baseline
$results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $root 'summary.json') -Encoding UTF8
Write-Output "Evidence: $root"
if (!$restored -or ($results | Where-Object { !$_.passed })) { throw 'BlackMoon regression failed; inspect summary.json and per-case results.' }
Write-Output 'All BlackMoon regression cases passed; original config hash preserved.'
