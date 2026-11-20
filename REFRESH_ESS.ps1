# =============================================================================
#  REFRESH_ESS.ps1  -  ONE-COMMAND REFRESH (stops before GitHub)
#    1. Backs up the 17 current ESS output xlsx files
#    2. Runs Run_All_17.ps1 (DB2 -> 17 xlsx; ~1h40m)         [skip with -SkipQuery]
#    3. Verifies the run in Run_All_17.log (restores backup on failure)
#    4. Dumps all 18 ESS xlsx and builds dashboard-ess.html (UPDATE_DASHBOARD_ESS.ps1)
#    5. Runs the automated checks (check-ess.ps1)
#    6. Prints a review summary and STOPS.
#  It never runs git. Publishing is PUBLISH_ESS.ps1 -Approved, run only after you approve.
#
#  P11 (2025).xlsx is NOT produced by Run_All_17 - upload it to the ESS folder first.
#  Params:  -SkipQuery   reuse the xlsx already in the ESS folder
#           -Only 'A','B' pass through to Run_All_17.ps1 (refresh just those programs)
# =============================================================================
param([switch]$SkipQuery, [string[]]$Only = @())

$ErrorActionPreference = 'Stop'
$dir    = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$ess    = "C:\Users\MuruganVenugopal\.bob\ESS"
$log    = "$ess\Run_All_17.log"
$sumF   = "$dir\ess-build-summary.json"
$t0     = Get-Date
Set-Location $dir

Write-Host "`n==================== ESS REFRESH ====================" -ForegroundColor Cyan

# -- P11 (2025) manual upload check -------------------------------------------
$p11 = Get-Item "$ess\P11 (2025).xlsx" -ErrorAction SilentlyContinue
if (-not $p11) { throw "P11 (2025).xlsx is missing from $ess - upload it first." }
$age = [math]::Round(((Get-Date) - $p11.LastWriteTime).TotalDays, 1)
Write-Host ("P11 (2025).xlsx (manual upload) last modified {0}  ({1} days ago)" -f $p11.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), $age) -ForegroundColor $(if ($age -gt 7) { 'Yellow' } else { 'Green' })
if ($age -gt 7) { Write-Host "  WARNING: P11 file is older than 7 days - upload a fresh one if you have it." -ForegroundColor Yellow }

# keep the previous build summary so we can show changes
if (Test-Path $sumF) { Copy-Item $sumF "$dir\ess-build-summary.prev.json" -Force }

# -- Steps 1-3: refresh the 17 output files -----------------------------------
if ($SkipQuery) {
    Write-Host "`n[1-3] SKIPPED query run (-SkipQuery) - using xlsx already in ESS folder" -ForegroundColor Yellow
} else {
    $bk = "$ess\Old Files\pre-refresh_$($t0.ToString('yyyyMMdd_HHmmss'))"
    Write-Host "`n[1/5] Backing up current output files -> $bk" -ForegroundColor White
    New-Item $bk -ItemType Directory -Force | Out-Null
    Get-ChildItem "$ess\*.xlsx" | Where-Object { $_.Name -ne 'P11 (2025).xlsx' } | Copy-Item -Destination $bk -Force
    # keep only the 2 most recent pre-refresh backups
    Get-ChildItem "$ess\Old Files" -Directory -Filter 'pre-refresh_*' | Sort-Object Name -Descending | Select-Object -Skip 2 |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force }

    Write-Host "[2/5] Running Run_All_17.ps1 (this takes ~1h40m)..." -ForegroundColor White

    # Runs Run_All_17.ps1 (optionally for ONE program) and returns @{ Name = status } from the SUMMARY it wrote.
    # Each call is a new powershell process = a fresh DB2 connection.
    function Invoke-Query([string]$one) {
        $logBefore = if (Test-Path $log) { @(Get-Content $log).Count } else { 0 }
        $a = @('-ExecutionPolicy','Bypass','-File',"$ess\Run_All_17.ps1")
        if ($one) { $a += '-Only'; $a += $one }
        & powershell.exe @a
        $new = @(Get-Content $log | Select-Object -Skip $logBefore)
        $st = @{}; $inSum = $false
        foreach ($l in $new) {
            if ($l -match '=== SUMMARY ===') { $inSum = $true; continue }
            if (-not $inSum) { continue }
            $m = [regex]::Match($l, '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d (.*)$'); if (-not $m.Success) { continue }
            $body = $m.Groups[1].Value
            if ($body -match '^(File|----|===)') { continue }
            # Format-Table pads columns to the widest value, so in a one-row summary the name sits only ONE space
            # from the numbers ("z17 - 9175 61998 ..."). Strip the trailing numeric columns + status instead.
            $mm = [regex]::Match($body, '^(.+?)\s+\d+\s+(?:\d+\s+)?\d+\s+(?:\d+\s+)?(OK|FAIL.*)$')
            if (-not $mm.Success) { continue }
            $st[$mm.Groups[1].Value.Trim()] = if ($mm.Groups[2].Value -eq 'OK') { 'OK' } else { 'FAIL' }
        }
        return @{ Status = $st; Lines = $new }
    }

    $status = @{}; $lastLines = @()
    if ($Only.Count) { foreach ($o in $Only) { $r = Invoke-Query $o; $r.Status.GetEnumerator() | ForEach-Object { $status[$_.Key] = $_.Value }; $lastLines = $r.Lines } }
    else             { $r = Invoke-Query '';  $status = $r.Status; $lastLines = $r.Lines }

    if ($status.Count -eq 0) {
        # no summary at all = aborted before/at start (VPN down, no DB2 ...) - nothing trustworthy was written
        Write-Host "`nRun_All_17 produced no summary (aborted). Restoring previous output files..." -ForegroundColor Red
        Get-ChildItem "$bk\*.xlsx" | Copy-Item -Destination $ess -Force
        $lastLines | Select-Object -Last 15 | ForEach-Object { Write-Host "  $_" }
        throw "Query step failed - dashboard NOT rebuilt, previous data restored."
    }

    # A program that failed (e.g. DB2 connection dropped after a long job) is retried on its own, fresh connection.
    foreach ($attempt in 1..2) {
        $failed = @($status.GetEnumerator() | Where-Object { $_.Value -ne 'OK' } | ForEach-Object { $_.Key })
        if (-not $failed.Count) { break }
        Write-Host ("`nRetry {0}/2 for: {1}" -f $attempt, ($failed -join ', ')) -ForegroundColor Yellow
        foreach ($f in $failed) { $r = Invoke-Query $f; if ($r.Status.ContainsKey($f)) { $status[$f] = $r.Status[$f] } else { $status[$f] = 'FAIL' } }
    }
    $failed = @($status.GetEnumerator() | Where-Object { $_.Value -ne 'OK' } | ForEach-Object { $_.Key })
    if (-not $Only.Count -and $status.Count -ne 17) { $failed += "(only $($status.Count) of 17 programs reported)" }
    if ($failed.Count) {
        Write-Host "`nStill failing after retries: $($failed -join ', ')" -ForegroundColor Red
        foreach ($f in $failed) { $bf = "$bk\$f.xlsx"; if (Test-Path $bf) { Copy-Item $bf $ess -Force; Write-Host "  restored previous $f.xlsx" -ForegroundColor Red } }
        throw "Query step failed for: $($failed -join ', ') - dashboard NOT rebuilt (other programs keep their fresh files, failed ones were restored)."
    }
    Write-Host "[3/5] Run_All_17 finished: $($status.Count) programs saved and verified" -ForegroundColor Green
}

# -- Step 4: dump + build dashboard ------------------------------------------
Write-Host "`n[4/5] Dumping ESS files and building dashboard-ess.html..." -ForegroundColor White
& powershell.exe -ExecutionPolicy Bypass -File "$dir\UPDATE_DASHBOARD_ESS.ps1"
if ($LASTEXITCODE -ne 0) { throw "UPDATE_DASHBOARD_ESS.ps1 failed" }

# -- Step 5: checks -----------------------------------------------------------
Write-Host "`n[5/5] Running automated checks..." -ForegroundColor White
& powershell.exe -ExecutionPolicy Bypass -File "$dir\check-ess.ps1"
$checksOk = ($LASTEXITCODE -eq 0)

# -- Review summary -----------------------------------------------------------
Write-Host "`n==================== REVIEW ====================" -ForegroundColor Cyan
$cur  = @(Get-Content $sumF -Raw | ConvertFrom-Json | ForEach-Object { $_ })
$prev = @{}
if (Test-Path "$dir\ess-build-summary.prev.json") { @(Get-Content "$dir\ess-build-summary.prev.json" -Raw | ConvertFrom-Json | ForEach-Object { $_ }) | ForEach-Object { $prev[$_.program] = $_.outRows } }
"{0,-26}{1,8}{2,10}{3,8}" -f 'Program','Rows','Previous','Change'
foreach ($p in $cur) {
    $pv = if ($prev.ContainsKey($p.program)) { $prev[$p.program] } else { $null }
    $chg = if ($null -eq $pv) { 'new' } else { '{0:+#;-#;0}' -f ($p.outRows - $pv) }
    "{0,-26}{1,8}{2,10}{3,8}" -f $p.program, $p.outRows, $(if ($null -eq $pv) { '-' } else { $pv }), $chg
}
"{0,-26}{1,8}" -f 'TOTAL', ($cur | Measure-Object outRows -Sum).Sum
Write-Host ("`nElapsed: {0:N0} min" -f ((Get-Date) - $t0).TotalMinutes)
if ($checksOk) {
    Write-Host "`nCHECKS PASSED. Preview: $dir\dashboard-ess.html" -ForegroundColor Green
    Write-Host "NOTHING HAS BEEN PUSHED. After you review and approve, publish with:" -ForegroundColor Cyan
    Write-Host "  .\PUBLISH_ESS.ps1 -Approved" -ForegroundColor Yellow
} else {
    Write-Host "`nCHECKS FAILED - see FAIL lines above. Do not publish." -ForegroundColor Red
    exit 1
}
