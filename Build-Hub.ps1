# =============================================================================
#  Build-Hub.ps1  -  builds the NPI HUB (local preview folder .\hub\)
#
#  Each tab in hub-tabs.json is a finished, self-contained dashboard .html file. The hub shows every
#  one in its own frame, so a dashboard keeps ITS OWN code, data, filters and logic untouched
#  (no id/CSS/JS clashes are possible). Files are copied BYTE-FOR-BYTE and hash-verified.
#
#  To add a new dashboard as a tab:  add one entry to hub-tabs.json, run this script. Nothing else changes.
#
#  Output : .\hub\index.html + one .html per tab + hub-manifest.json   (preview only)
#  This script never runs git, never touches index.html / dashboard-final.html, never changes a source dashboard.
# =============================================================================
param([string]$OutDir = '')

$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutDir) { $OutDir = Join-Path $dir 'hub' }
$tabsFile = Join-Path $dir 'hub-tabs.json'
$tplFile  = Join-Path $dir 'hub.template.html'

if (-not (Test-Path $tabsFile)) { throw "Missing $tabsFile" }
if (-not (Test-Path $tplFile))  { throw "Missing $tplFile" }
$tabs = @(Get-Content $tabsFile -Raw | ConvertFrom-Json | ForEach-Object { $_ })
if ($tabs.Count -lt 1) { throw 'hub-tabs.json has no tabs' }

$seenId = @{}; $seenFile = @{}
foreach ($t in $tabs) {
    foreach ($k in 'id','label','source','file') { if (-not $t.$k) { throw "A tab entry is missing '$k'" } }
    if ($t.id -notmatch '^[a-z0-9][a-z0-9-]*$')   { throw "Tab id '$($t.id)' must be lowercase letters, digits and hyphens" }
    if ($t.file -notmatch '^[A-Za-z0-9._-]+\.html$' -or $t.file -eq 'index.html') { throw "Tab file '$($t.file)' must be a plain name ending in .html (not index.html)" }
    if ($seenId.ContainsKey($t.id))     { throw "Duplicate tab id '$($t.id)'" };     $seenId[$t.id] = 1
    if ($seenFile.ContainsKey($t.file)) { throw "Duplicate tab file '$($t.file)'" }; $seenFile[$t.file] = 1
}

# build in a temp folder; only swap into place when everything verified
$tmp = $OutDir + '.building'
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
New-Item $tmp -ItemType Directory | Out-Null

$manifest = @()
foreach ($t in $tabs) {
    $src = if ([IO.Path]::IsPathRooted($t.source)) { $t.source } else { Join-Path $dir $t.source }
    if (-not (Test-Path -LiteralPath $src)) { throw "Source dashboard for tab '$($t.label)' not found: $src" }
    $len = (Get-Item -LiteralPath $src).Length
    if ($len -lt 1000) { throw "Source for '$($t.label)' is suspiciously small ($len bytes): $src" }
    $head = [IO.File]::ReadAllText($src, [Text.Encoding]::UTF8)
    if ($head -notmatch '(?i)</html>\s*$') { throw "Source for '$($t.label)' does not end with </html> (incomplete file?): $src" }
    $dst = Join-Path $tmp $t.file
    [IO.File]::Copy($src, $dst, $true)
    $h1 = (Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash
    $h2 = (Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash
    if ($h1 -ne $h2) { throw "Copy of '$($t.label)' is not byte-identical to its source" }
    $manifest += [pscustomobject]@{ id = $t.id; label = $t.label; file = $t.file; source = $src; sha256 = $h1; bytes = $len; sourceModified = (Get-Item -LiteralPath $src).LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') }
    Write-Host ("  tab '{0}': {1:N2} MB  sha256 {2}..." -f $t.label, ($len / 1MB), $h1.Substring(0, 12)) -ForegroundColor Green
}

# hub shell
$js = '[' + ((@($tabs | ForEach-Object { '{"id":"' + $_.id + '","label":"' + ($_.label -replace '\\','\\' -replace '"','\"') + '","src":"' + $_.file + '"}' })) -join ',') + ']'
$tpl = [IO.File]::ReadAllText($tplFile, [Text.Encoding]::UTF8)
if (-not $tpl.Contains('/*{{HUB_TABS}}*/[]')) { throw 'Placeholder /*{{HUB_TABS}}*/[] not found in hub.template.html' }
$html = $tpl.Replace('/*{{HUB_TABS}}*/[]', $js)
[IO.File]::WriteAllText((Join-Path $tmp 'index.html'), $html, (New-Object Text.UTF8Encoding($false)))
$manifest | ConvertTo-Json -Depth 3 | Set-Content (Join-Path $tmp 'hub-manifest.json') -Encoding UTF8

# swap into place
$swapped = $false
if (Test-Path $OutDir) {
    try { Remove-Item $OutDir -Recurse -Force -ErrorAction Stop }
    catch {
        # the folder itself can be locked (open in Explorer / a terminal / an editor); replace its contents instead
        Get-ChildItem $OutDir -Force | Remove-Item -Recurse -Force
        Get-ChildItem $tmp -Force | Move-Item -Destination $OutDir -Force
        Remove-Item $tmp -Recurse -Force
        $swapped = $true
    }
}
if (-not $swapped) { Rename-Item $tmp (Split-Path $OutDir -Leaf) }
Write-Host ("`nNPI HUB built: {0}\index.html  ({1} tabs)" -f $OutDir, $tabs.Count) -ForegroundColor Cyan
Write-Host 'Preview only. Nothing was committed or pushed; the live dashboard is untouched.' -ForegroundColor Cyan
