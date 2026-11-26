# =============================================================================
#  REFRESH_HUB.ps1  -  refresh one or all NPI HUB tabs, rebuild the hub, run the checks.
#  It NEVER runs git and never touches the live site. Publishing is PUBLISH_HUB.ps1 -Approved.
#
#  -Tab ess   ESS Installs         : runs REFRESH_ESS.ps1 (DB2 + P11)         - on request
#  -Tab npi   Daily NPI EO Orders  : runs NPI Tracker\build\Build-Dashboard.ps1 - daily, after you upload the Masterfile
#  -Tab all   both (default)
#  -Master    NPI Masterfile to use (default: newest Masterfile_*.xlsx in the NPI Tracker folder, backups excluded)
#  -SkipQuery passed on to REFRESH_ESS.ps1 (reuse the xlsx already in the ESS folder)
#  -RebuildOnly  do not refresh any data, just re-copy the current dashboards into the hub and check
# =============================================================================
param([ValidateSet('ess','npi','all')][string]$Tab = 'all', [string]$Master = '', [switch]$SkipQuery, [switch]$RebuildOnly)

$ErrorActionPreference = 'Stop'
$dir = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$npi = "C:\Users\MuruganVenugopal\.bob\playground\NPI Tracker"
Set-Location $dir
$ps = 'powershell.exe'

if (-not $RebuildOnly) {
    if ($Tab -in 'npi','all') {
        Write-Host "`n=== Daily NPI EO Orders: rebuild from the Masterfile ===" -ForegroundColor Cyan
        if (-not $Master) {
            $m = Get-ChildItem $npi -Filter 'Masterfile*.xlsx' | Where-Object { $_.Name -notmatch 'backup' -and $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if (-not $m) { throw "No Masterfile*.xlsx found in $npi" }
            $Master = $m.FullName
        }
        if (-not (Test-Path -LiteralPath $Master)) { throw "Masterfile not found: $Master" }
        Write-Host "Masterfile: $Master  (modified $((Get-Item -LiteralPath $Master).LastWriteTime))"
        & $ps -NoProfile -ExecutionPolicy Bypass -File "$npi\build\Build-Dashboard.ps1" -Master $Master
        if ($LASTEXITCODE -ne 0) { throw "NPI Build-Dashboard.ps1 failed (exit $LASTEXITCODE) - see $npi\build\build-log.txt. The previous NPI dashboard is unchanged." }
    }
    if ($Tab -in 'ess','all') {
        Write-Host "`n=== ESS Installs: DB2 refresh ===" -ForegroundColor Cyan
        $a = @('-NoProfile','-ExecutionPolicy','Bypass','-File',"$dir\REFRESH_ESS.ps1")
        if ($SkipQuery) { $a += '-SkipQuery' }
        & $ps @a
        if ($LASTEXITCODE -ne 0) { throw "REFRESH_ESS.ps1 failed (exit $LASTEXITCODE). The hub was not rebuilt." }
    }
}

Write-Host "`n=== Rebuild the hub ===" -ForegroundColor Cyan
& $ps -NoProfile -ExecutionPolicy Bypass -File "$dir\Build-Hub.ps1"
if ($LASTEXITCODE -ne 0) { throw "Build-Hub.ps1 failed (exit $LASTEXITCODE)" }

Write-Host "`n=== Check the hub ===" -ForegroundColor Cyan
& $ps -NoProfile -ExecutionPolicy Bypass -File "$dir\check-hub.ps1"
if ($LASTEXITCODE -ne 0) { Write-Host "`nHUB CHECKS FAILED - do not publish." -ForegroundColor Red; exit 1 }

Write-Host "`nCHECKS PASSED. Preview: $dir\hub\index.html" -ForegroundColor Green
Write-Host "NOTHING HAS BEEN PUSHED. After you review and approve, publish with:" -ForegroundColor Cyan
Write-Host "  .\PUBLISH_HUB.ps1 -Approved" -ForegroundColor Yellow
