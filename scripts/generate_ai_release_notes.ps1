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

function Resolve-TagCommit([string]$TagName) {
	if ([string]::IsNullOrWhiteSpace($TagName)) {
		return ""
	}
	$commit = Invoke-GitText @("rev-parse", "$TagName^{}")
	if (-not [string]::IsNullOrWhiteSpace($commit)) {
		return $commit
	}
	return Invoke-GitText @("rev-parse", $TagName)
}

function Resolve-PreviousTag([string]$CurrentTag, [string]$CurrentCommit) {
	if ([string]::IsNullOrWhiteSpace($CurrentCommit)) {
		return ""
	}
	$parentRef = "$CurrentCommit^"
	$previous = Invoke-GitText @("describe", "--tags", "--abbrev=0", $parentRef)
	if (-not [string]::IsNullOrWhiteSpace($previous) -and $previous -ne $CurrentTag) {
		return $previous
	}

	$allTagsText = Invoke-GitText @("tag", "--merged", $parentRef, "--sort=-creatordate")
	if ([string]::IsNullOrWhiteSpace($allTagsText)) {
		return ""
	}
	foreach ($tag in ($allTagsText -split "`n")) {
		$trimmed = $tag.Trim()
		if (-not [string]::IsNullOrWhiteSpace($trimmed) -and $trimmed -ne $CurrentTag) {
			return $trimmed
		}
	}
	return ""
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

function Get-CommitSubjectLines([string]$CommitSubjectsText) {
	if ([string]::IsNullOrWhiteSpace($CommitSubjectsText)) {
		return @()
	}

	$subjects = @()
	foreach ($line in ($CommitSubjectsText -split "`r?`n")) {
		if (-not [string]::IsNullOrWhiteSpace($line)) {
			$subjects += $line.Trim()
		}
	}
	return $subjects
}

function New-ReleaseNotesFromCategories(
	[string]$ReleaseName,
	[string[]]$CommitSubjects,
	[hashtable]$CategoriesById
) {
	$allowedCategories = @("新功能", "改进", "修正")
	$grouped = @{
		"新功能" = New-Object System.Collections.Generic.List[string]
		"改进" = New-Object System.Collections.Generic.List[string]
		"修正" = New-Object System.Collections.Generic.List[string]
	}

	for ($index = 0; $index -lt $CommitSubjects.Count; $index++) {
		$id = $index + 1
		$category = "改进"
		if ($CategoriesById.ContainsKey($id) -and ($allowedCategories -contains $CategoriesById[$id])) {
			$category = $CategoriesById[$id]
		}
		$grouped[$category].Add($CommitSubjects[$index]) | Out-Null
	}

	$notes = @()
	$notes += "# $ReleaseName"
	$notes += ""

	if ($CommitSubjects.Count -eq 0) {
		$notes += "## 改进"
		$notes += ""
		$notes += "- 未找到可整理的变更。"
		return ($notes -join "`r`n") + "`r`n"
	}

	foreach ($category in $allowedCategories) {
		if ($grouped[$category].Count -eq 0) {
			continue
		}
		$notes += "## $category"
		$notes += ""
		foreach ($subject in $grouped[$category]) {
			$notes += "- $subject"
		}
		$notes += ""
	}

	return ($notes -join "`r`n").TrimEnd() + "`r`n"
}

function New-FallbackNotes(
	[string]$ReleaseName,
	[string]$CurrentTag,
	[string]$PreviousTag,
	[string[]]$CommitSubjects
) {
	return New-ReleaseNotesFromCategories `
		-ReleaseName $ReleaseName `
		-CommitSubjects $CommitSubjects `
		-CategoriesById @{}
}

function ConvertFrom-AiJsonContent([string]$Content) {
	$text = $Content.Trim()
	$text = $text -replace "^\s*```(?:json)?\s*", ""
	$text = $text -replace "\s*```\s*$", ""

	$start = $text.IndexOf("{")
	$end = $text.LastIndexOf("}")
	if ($start -ge 0 -and $end -gt $start) {
		$text = $text.Substring($start, $end - $start + 1)
	}

	return $text | ConvertFrom-Json
}

function Invoke-AiCommitClassification(
	[string]$ApiKey,
	[string]$BaseUrl,
	[string]$Model,
	[string]$ReleaseName,
	[string]$CurrentTag,
	[string]$PreviousTag,
	[string[]]$CommitSubjects
) {
	$endpoint = Resolve-ChatCompletionsEndpoint $BaseUrl
	$rangeText = "$PreviousTag ... $CurrentTag"
	$indexedSubjects = New-Object System.Collections.Generic.List[string]
	for ($index = 0; $index -lt $CommitSubjects.Count; $index++) {
		$indexedSubjects.Add("$($index + 1). $($CommitSubjects[$index])") | Out-Null
	}
	$indexedSubjectText = $indexedSubjects -join "`n"

	$prompt = @"
你是 AutoLinker 项目的发布说明分类助手。请只判断每条 commit 标题属于哪个分类。

要求：
1. 只能返回 JSON，不要返回 Markdown，不要解释。
2. category 只能是：新功能、改进、修正。
3. 必须按 id 分类；不要改写、扩写、翻译、合并或拆分标题。
4. 无法明确判断时归类为“改进”。
5. 返回格式必须是：{"items":[{"id":1,"category":"改进"}]}。

Release: $ReleaseName
Tag: $CurrentTag
Range: $rangeText

原始 commit 标题列表：
$indexedSubjectText
"@

	$body = @{
		model = $Model
		temperature = 0.2
		messages = @(
			@{
				role = "system"
				content = "你只负责把 GitHub commit 标题 id 分类为新功能、改进或修正，并且只返回 JSON。"
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

	$result = ConvertFrom-AiJsonContent $content
	$categoriesById = @{}
	foreach ($item in @($result.items)) {
		try {
			$id = [int]$item.id
		}
		catch {
			continue
		}

		$category = [string]$item.category
		if ($id -lt 1 -or $id -gt $CommitSubjects.Count) {
			continue
		}
		if (@("新功能", "改进", "修正") -notcontains $category) {
			continue
		}
		if (-not $categoriesById.ContainsKey($id)) {
			$categoriesById[$id] = $category
		}
	}

	return $categoriesById
}

$currentTag = Get-EnvValue "RELEASE_TAG" (Get-EnvValue "GITHUB_REF_NAME")
$releaseName = Get-EnvValue "RELEASE_NAME" $currentTag
if ([string]::IsNullOrWhiteSpace($currentTag)) {
	throw "RELEASE_TAG or GITHUB_REF_NAME is required."
}

if ((Invoke-GitText @("rev-parse", "--is-shallow-repository")) -eq "true") {
	Invoke-GitText @("fetch", "--tags", "--force", "--prune", "--unshallow") | Out-Null
}
Invoke-GitText @("fetch", "--tags", "--force", "--prune") | Out-Null

$currentCommit = Resolve-TagCommit $currentTag
if ([string]::IsNullOrWhiteSpace($currentCommit)) {
	throw "Unable to resolve current tag commit: $currentTag"
}

$previousTag = Resolve-PreviousTag $currentTag $currentCommit
$range = if ([string]::IsNullOrWhiteSpace($previousTag)) { "$currentCommit^..$currentCommit" } else { "$previousTag..$currentCommit" }
Write-Output "Release notes current_tag=$currentTag current_commit=$currentCommit previous_tag=$previousTag range=$range"

$commitSubjectsText = Invoke-GitText @("log", $range, "--pretty=format:%s")
$commitSubjects = Get-CommitSubjectLines $commitSubjectsText

$apiKey = Get-EnvValue "AI_NOTES_API_KEY"
$baseUrl = Get-EnvValue "AI_NOTES_BASE_URL" "https://api.openai.com/v1"
$model = Get-EnvValue "AI_NOTES_MODEL"

$notes = ""
if ([string]::IsNullOrWhiteSpace($previousTag)) {
	Write-Warning "Previous tag was not found. Falling back to non-AI release notes to avoid summarizing the whole repository."
	$notes = New-FallbackNotes `
		-ReleaseName $releaseName `
		-CurrentTag $currentTag `
		-PreviousTag $previousTag `
		-CommitSubjects $commitSubjects
}
elseif ([string]::IsNullOrWhiteSpace($apiKey) -or [string]::IsNullOrWhiteSpace($model)) {
	Write-Warning "AI_NOTES_API_KEY or AI_NOTES_MODEL is empty. Falling back to non-AI release notes."
	$notes = New-FallbackNotes `
		-ReleaseName $releaseName `
		-CurrentTag $currentTag `
		-PreviousTag $previousTag `
		-CommitSubjects $commitSubjects
}
else {
	try {
		$categoriesById = Invoke-AiCommitClassification `
			-ApiKey $apiKey `
			-BaseUrl $baseUrl `
			-Model $model `
			-ReleaseName $releaseName `
			-CurrentTag $currentTag `
			-PreviousTag $previousTag `
			-CommitSubjects $commitSubjects
		$notes = New-ReleaseNotesFromCategories `
			-ReleaseName $releaseName `
			-CommitSubjects $commitSubjects `
			-CategoriesById $categoriesById
	}
	catch {
		Write-Warning "AI release notes generation failed: $($_.Exception.Message)"
		$notes = New-FallbackNotes `
			-ReleaseName $releaseName `
			-CurrentTag $currentTag `
			-PreviousTag $previousTag `
			-CommitSubjects $commitSubjects
	}
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($outputFullPath, $notes, $utf8Bom)
Write-Output "Release notes written to $outputFullPath"
