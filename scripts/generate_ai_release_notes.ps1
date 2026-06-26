param(
	[string]$OutputPath = "release_notes.md"
)

$ErrorActionPreference = "Stop"

function Get-EnvValue([string]$Name, [string]$DefaultValue = "") {
	$value = [Environment]::GetEnvironmentVariable($Name)
	if ([string]::IsNullOrWhiteSpace($value)) {
		return $DefaultValue
	}
	return $value.Trim()
}

function Limit-Text([string]$Text, [int]$MaxLength) {
	if ([string]::IsNullOrEmpty($Text) -or $Text.Length -le $MaxLength) {
		return $Text
	}
	return $Text.Substring(0, $MaxLength) + "`n...（内容过长，已截断）"
}

function Invoke-GitText([string[]]$ArgsList) {
	$output = & git @ArgsList 2>$null
	if ($LASTEXITCODE -ne 0) {
		return ""
	}
	return ($output -join "`n").Trim()
}

function Resolve-ChatCompletionsEndpoint([string]$BaseUrl) {
	$url = $BaseUrl.Trim()
	while ($url.EndsWith("/")) {
		$url = $url.Substring(0, $url.Length - 1)
	}
	if ($url.EndsWith("/chat/completions", [StringComparison]::OrdinalIgnoreCase)) {
		return $url
	}
	if ($url.EndsWith("/v1", [StringComparison]::OrdinalIgnoreCase)) {
		return "$url/chat/completions"
	}
	return "$url/v1/chat/completions"
}

function New-FallbackNotes(
	[string]$ReleaseName,
	[string]$CurrentTag,
	[string]$PreviousTag,
	[string]$CommitSubjects,
	[string]$DiffStat
) {
	$rangeText = if ([string]::IsNullOrWhiteSpace($PreviousTag)) {
		"首次发布或未找到上一个 tag。"
	}
	else {
		"$PreviousTag ... $CurrentTag"
	}

	$notes = @()
	$notes += "# $ReleaseName"
	$notes += ""
	$notes += "## 更新摘要"
	$notes += ""
	$notes += "本版本根据提交记录自动生成发布说明。AI 发布摘要未启用或生成失败，因此保留原始变更列表。"
	$notes += ""
	$notes += "变更范围：$rangeText"
	$notes += ""
	if (-not [string]::IsNullOrWhiteSpace($DiffStat)) {
		$notes += "## 文件变更统计"
		$notes += ""
		$notes += '```text'
		$notes += $DiffStat
		$notes += '```'
		$notes += ""
	}
	$notes += "## 完整变更"
	$notes += ""
	if ([string]::IsNullOrWhiteSpace($CommitSubjects)) {
		$notes += "- 未找到提交记录。"
	}
	else {
		$notes += $CommitSubjects
	}
	return ($notes -join "`r`n") + "`r`n"
}

function Invoke-AiReleaseNotes(
	[string]$ApiKey,
	[string]$BaseUrl,
	[string]$Model,
	[string]$ReleaseName,
	[string]$CurrentTag,
	[string]$PreviousTag,
	[string]$CommitDetails,
	[string]$CommitSubjects,
	[string]$DiffStat,
	[string]$ChangedFiles
) {
	$endpoint = Resolve-ChatCompletionsEndpoint $BaseUrl
	$rangeText = if ([string]::IsNullOrWhiteSpace($PreviousTag)) {
		"无上一个 tag，按当前 tag 可见提交生成。"
	}
	else {
		"$PreviousTag ... $CurrentTag"
	}

	$prompt = @"
你是 AutoLinker 项目的发布说明撰写助手。请根据提交记录、文件变更统计和文件列表，生成中文 GitHub Release Notes。

要求：
1. 不要编造提交记录中没有体现的功能。
2. 面向用户描述，不要写成开发流水账。
3. 按实际内容归类，可使用这些标题：更新摘要、新功能、问题修复、体验优化、构建与发布、内部调整、完整变更。
4. 没有内容的分类直接省略。
5. “完整变更”必须保留原始 commit 标题列表。
6. 输出 Markdown，不要包裹代码块。

Release: $ReleaseName
Tag: $CurrentTag
Range: $rangeText

提交详情：
$CommitDetails

文件变更统计：
$DiffStat

变更文件：
$ChangedFiles

原始 commit 标题：
$CommitSubjects
"@

	$body = @{
		model = $Model
		temperature = 0.2
		messages = @(
			@{
				role = "system"
				content = "你负责把 GitHub 提交记录整理成准确、克制、面向用户的中文发布说明。"
			},
			@{
				role = "user"
				content = $prompt
			}
		)
	} | ConvertTo-Json -Depth 8

	$response = Invoke-RestMethod `
		-Method Post `
		-Uri $endpoint `
		-Headers @{
			"Authorization" = "Bearer $ApiKey"
			"Content-Type" = "application/json"
		} `
		-Body $body `
		-TimeoutSec 120

	$content = $response.choices[0].message.content
	if ([string]::IsNullOrWhiteSpace($content)) {
		throw "AI response content is empty."
	}
	return $content.Trim() + "`r`n"
}

$currentTag = Get-EnvValue "RELEASE_TAG" (Get-EnvValue "GITHUB_REF_NAME")
$releaseName = Get-EnvValue "RELEASE_NAME" $currentTag
if ([string]::IsNullOrWhiteSpace($currentTag)) {
	throw "RELEASE_TAG or GITHUB_REF_NAME is required."
}

Invoke-GitText @("fetch", "--tags", "--force") | Out-Null
$previousTag = Invoke-GitText @("describe", "--tags", "--abbrev=0", "$currentTag^")
$range = if ([string]::IsNullOrWhiteSpace($previousTag)) { $currentTag } else { "$previousTag..$currentTag" }

$commitSubjects = Invoke-GitText @("log", $range, "--pretty=format:- %s")
$commitDetails = Invoke-GitText @("log", $range, "--pretty=format:commit %h%nsubject: %s%nauthor: %an%nbody:%n%b%n---")
$diffStat = if ([string]::IsNullOrWhiteSpace($previousTag)) {
	Invoke-GitText @("show", "--stat", "--oneline", "--no-renames", $currentTag)
}
else {
	Invoke-GitText @("diff", "--stat", "--no-renames", $previousTag, $currentTag)
}
$changedFiles = if ([string]::IsNullOrWhiteSpace($previousTag)) {
	Invoke-GitText @("show", "--name-only", "--pretty=format:", $currentTag)
}
else {
	Invoke-GitText @("diff", "--name-only", $previousTag, $currentTag)
}

$commitDetails = Limit-Text $commitDetails 24000
$commitSubjects = Limit-Text $commitSubjects 8000
$diffStat = Limit-Text $diffStat 12000
$changedFiles = Limit-Text $changedFiles 12000

$apiKey = Get-EnvValue "AI_NOTES_API_KEY"
$baseUrl = Get-EnvValue "AI_NOTES_BASE_URL" "https://api.openai.com/v1"
$model = Get-EnvValue "AI_NOTES_MODEL"

$notes = ""
if ([string]::IsNullOrWhiteSpace($apiKey) -or [string]::IsNullOrWhiteSpace($model)) {
	Write-Warning "AI_NOTES_API_KEY or AI_NOTES_MODEL is empty. Falling back to non-AI release notes."
	$notes = New-FallbackNotes $releaseName $currentTag $previousTag $commitSubjects $diffStat
}
else {
	try {
		$notes = Invoke-AiReleaseNotes `
			-ApiKey $apiKey `
			-BaseUrl $baseUrl `
			-Model $model `
			-ReleaseName $releaseName `
			-CurrentTag $currentTag `
			-PreviousTag $previousTag `
			-CommitDetails $commitDetails `
			-CommitSubjects $commitSubjects `
			-DiffStat $diffStat `
			-ChangedFiles $changedFiles
	}
	catch {
		Write-Warning "AI release notes generation failed: $($_.Exception.Message)"
		$notes = New-FallbackNotes $releaseName $currentTag $previousTag $commitSubjects $diffStat
	}
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($outputFullPath, $notes, $utf8Bom)
Write-Output "Release notes written to $outputFullPath"
