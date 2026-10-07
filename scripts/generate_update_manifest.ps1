param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [Parameter(Mandatory = $true)][string]$AssetPath,
    [string]$OutputPath = 'update.json'
)

$ErrorActionPreference = 'Stop'
# 清单只记录版本和附件信息，客户端根据固定仓库构造下载地址。
if ($Tag.Length -gt 128 -or $Tag -cnotmatch '^[vV]?[0-9][A-Za-z0-9.+-]*$') {
    throw "Invalid release tag: $Tag"
}
$versionLabel = $Tag -replace '^[vV]', ''
$asset = Get-Item -LiteralPath $AssetPath
if ($asset.PSIsContainer -or $asset.Name -cne "AutoLinker-$versionLabel.zip" -or $asset.Length -le 0) {
    throw "Invalid update archive: $AssetPath"
}
$manifest = [ordered]@{
    schema_version = 1
    tag = $Tag
    asset = [ordered]@{
        name = $asset.Name
        size = $asset.Length
        sha256 = (Get-FileHash -LiteralPath $asset.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}
$json = $manifest | ConvertTo-Json -Depth 3
[System.IO.File]::WriteAllText($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath), $json, [System.Text.UTF8Encoding]::new($false))
