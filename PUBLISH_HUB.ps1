# =============================================================================
#  PUBLISH_HUB.ps1  -  publish the reviewed NPI HUB to GitHub Pages
#  Refuses to run without -Approved, and refuses unless check-hub.ps1 passed for the CURRENT hub\ build
#  (same manifest, same tab-file hashes). Copies hub\* into the repo root, commits, merges to main, pushes,
#  then checks the live site.
#  Revert: git revert <that commit>   (or tag pre-npi-hub = ESS-only dashboard before the hub)
# =============================================================================
param([switch]$Approved, [ValidateSet('ess','npi','all')][string]$Tab = 'all')

$ErrorActionPreference = 'Stop'
$dir = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
Set-Location $dir

if (-not $Approved) { Write-Host "Not approved. Review hub\index.html, then run:  .\PUBLISH_HUB.ps1 -Tab ess|npi|all -Approved" -ForegroundColor Yellow; exit 2 }

function Invoke-Git {
    $ErrorActionPreference = 'Continue'      # git writes progress to stderr; judge by exit code only
    $out = & git.exe @args 2>&1
    $out | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') failed (exit $LASTEXITCODE)" }
}

# gate: checks must have passed for this exact hub build
$rep = Get-Content "$dir\hub-check-report.json" -Raw | ConvertFrom-Json
$man = @(Get-Content "$dir\hub\hub-manifest.json" -Raw | ConvertFrom-Json | ForEach-Object { $_ })
if (-not $rep.pass -or [int64]$rep.manifestTicks -ne (Get-Item "$dir\hub\hub-manifest.json").LastWriteTimeUtc.Ticks) {
    throw "check-hub.ps1 has not passed for the current hub build - run REFRESH_HUB.ps1 (or Build-Hub.ps1 + check-hub.ps1) first."
}
foreach ($t in $man) {
    $h = (Get-FileHash "$dir\hub\$($t.file)" -Algorithm SHA256).Hash
    if ($h -ne $t.sha256) { throw "hub\$($t.file) does not match the manifest - rebuild the hub." }
}

$idMap = @{ ess = 'ess-installs'; npi = 'daily-npi-eo-orders' }
$sel = if ($Tab -eq 'all') { $man } else { @($man | Where-Object { $_.id -eq $idMap[$Tab] }) }
if (-not $sel) { throw "Tab '$Tab' is not in the hub." }
foreach ($t in ($man | Where-Object { $sel -notcontains $_ })) { if (-not (Test-Path "$dir\$($t.file)")) { throw "Cannot publish only '$Tab': $($t.file) has never been published. Use -Tab all." } }

$branch = (& git.exe branch --show-current).Trim()
Write-Host "Publishing from branch '$branch'..." -ForegroundColor Cyan

# hub shell -> index.html, each tab file next to it; dashboard-final.html stays the standalone ESS dashboard
Copy-Item "$dir\hub\index.html" "$dir\index.html" -Force
foreach ($t in $sel) { Copy-Item "$dir\hub\$($t.file)" "$dir\$($t.file)" -Force }
$ess = $sel | Where-Object { $_.id -eq 'ess-installs' }
if ($ess) { Copy-Item "$dir\hub\$($ess.file)" "$dir\dashboard-final.html" -Force }

$files = @('index.html','dashboard-final.html','dashboard-ess.html','hub-tabs.json','hub.template.html','Build-Hub.ps1','check-hub.ps1','REFRESH_HUB.ps1','PUBLISH_HUB.ps1','HUB_README.md',
           'ess-build-summary.json','ess-query-summary.json','UPDATE_DASHBOARD_ESS.ps1','UPDATE_DASHBOARD.ps1','REFRESH_ESS.ps1','PUBLISH_ESS.ps1','check-ess.ps1','build-ess-template.ps1','dashboard-ess.template.html','dump-ess.ps1') + ($man | ForEach-Object { $_.file })
Invoke-Git add @($files | Select-Object -Unique | Where-Object { Test-Path "$dir\$_" })
$pending = (& git.exe status --porcelain --untracked-files=no) | Where-Object { $_ }
if ($pending) { Invoke-Git commit -m ("NPI HUB publish " + (Get-Date -Format 'yyyy-MM-dd HH:mm') + " (" + (($sel | ForEach-Object { $_.label }) -join ' + ') + ")") } else { Write-Host "  (nothing new to commit)" }
if ($branch -ne 'main') { Invoke-Git checkout main; Invoke-Git merge --no-edit $branch }
Invoke-Git push origin main

$a = (& git.exe rev-parse main).Trim(); $b = (& git.exe rev-parse origin/main).Trim()
if ($a -ne $b) { throw "main != origin/main after push" }
Write-Host "`nPushed. main == origin/main ($($a.Substring(0,7)))" -ForegroundColor Green

# verify the live site serves the hub and each tab file (GitHub Pages can take a few minutes)
$base = 'https://murugan-9.github.io/brand-installs-dashboard'
$want = @{ 'index.html' = (Get-FileHash "$dir\hub\index.html" -Algorithm SHA256).Hash }
foreach ($t in $sel) { $want[$t.file] = $t.sha256 }
$ok = $false
for ($i = 1; $i -le 14 -and -not $ok; $i++) {
    $ok = $true
    foreach ($k in $want.Keys) {
        try {
            $r = Invoke-WebRequest -Uri "$base/$k`?v=$([guid]::NewGuid().ToString('N'))" -UseBasicParsing -TimeoutSec 60
            $ms = New-Object IO.MemoryStream; $bytes = $r.RawContentStream.ToArray()
            $sha = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($bytes)).Replace('-','')
            if ($sha -ne $want[$k]) { $ok = $false }
        } catch { $ok = $false }
    }
    if (-not $ok) { Write-Host "  live site not updated yet (try $i/14)..."; Start-Sleep -Seconds 20 }
}
if ($ok) { Write-Host "LIVE VERIFIED: hub + $($sel.Count) tab file(s) match the local build." -ForegroundColor Green }
else     { Write-Host "Live site did not match after ~5 min - pushed OK, check $base/ again shortly (Ctrl+Shift+R)." -ForegroundColor Yellow }
Write-Host "$base/" -ForegroundColor Cyan

