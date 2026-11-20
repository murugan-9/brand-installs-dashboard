# =============================================================================
#  REFRESH_ESS.ps1  -  ONE-COMMAND REFRESH (stops before GitHub)
#    1. Backs up the 17 current ESS output xlsx files
#    2. Runs Run_All_17.ps1 for all 17 programs IN PARALLEL (4 processes, 4 DB2 connections):
#         worker 1 = the 14 small programs, one after another
#         worker 2 = FlashSystem   worker 3 = z17 - 9175   worker 4 = DS8K
#       (use -Sequential for the old single-process behaviour)
#    3. Verifies each program; any that failed is retried on its own fresh connection (twice),
#       and if it still fails its previous xlsx is restored.
#    4. Dumps all 18 ESS xlsx and builds dashboard-ess.html (UPDATE_DASHBOARD_ESS.ps1)
#    5. Runs the automated checks (check-ess.ps1)
#    6. Prints a review summary and STOPS.
#  It never runs git. Publishing is PUBLISH_ESS.ps1 -Approved, run only after you approve.
#
#  P11 (2025).xlsx is NOT produced by Run_All_17 - upload it to the ESS folder first.
#  Params:  -SkipQuery    reuse the xlsx already in the ESS folder
#           -Sequential   one process, programs one after another (the original behaviour)
#           -Only 'A','B' refresh just those programs
#  Run_All_17.ps1 itself is never modified.
# =============================================================================
param([switch]$SkipQuery, [switch]$Sequential, [string[]]$Only = @())

$ErrorActionPreference = 'Stop'
$dir    = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$ess    = "C:\Users\MuruganVenugopal\.bob\ESS"
$sumF   = "$dir\ess-build-summary.json"
$qsF    = "$dir\ess-query-summary.json"
$workDir = "$dir\ess-workers"
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

# Runs one batch of Run_All_17.ps1 processes (one per group) at the same time and waits for all of them.
# Each process has its OWN stdout file, so results never depend on the shared, interleaved Run_All_17.log.
# Returns @{ 'Program' = @{ Status='OK'|'FAIL'; Main; DB2Main; Bottom; DB2Bottom } } for every program asked for.
function Invoke-Workers($groups, [string]$tag) {
    $w = @()
    for ($g = 0; $g -lt $groups.Count; $g++) {
        $list = (@($groups[$g]) | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ','
        $out  = "$workDir\$tag-w$($g + 1).out.log"; $err = "$workDir\$tag-w$($g + 1).err.log"
        $cmd  = "& '$ess\Run_All_17.ps1' -Only $list"
        $p = Start-Process powershell.exe -ArgumentList @('-ExecutionPolicy', 'Bypass', '-Command', "`"$cmd`"") `
             -RedirectStandardOutput $out -RedirectStandardError $err -WindowStyle Hidden -PassThru
        $w += [pscustomobject]@{ N = $g + 1; P = $p; Names = @($groups[$g]); Out = $out; Start = Get-Date; End = $null }
        Write-Host ("  worker {0} started (pid {1}): {2}" -f ($g + 1), $p.Id, $(if ($groups[$g].Count -gt 3) { "$($groups[$g].Count) programs" } else { @($groups[$g]) -join ', ' }))
    }
    $beat = Get-Date
    while (@($w | Where-Object { -not $_.P.HasExited }).Count -gt 0) {
        Start-Sleep -Seconds 15
        foreach ($x in $w) { if ($x.P.HasExited -and -not $x.End) { $x.End = Get-Date; Write-Host ("  worker {0} finished after {1:N0} min" -f $x.N, ($x.End - $x.Start).TotalMinutes) -ForegroundColor DarkGray } }
        if (((Get-Date) - $beat).TotalMinutes -ge 5) {
            $beat = Get-Date
            $line = foreach ($x in $w) {
                $done = 0; try { $done = @(Get-Content $x.Out -ErrorAction Stop | Select-String ' - COMPLETE').Count } catch {}
                "w$($x.N):$(if ($x.P.HasExited) { 'done' } else { "$done/$($x.Names.Count)" })"
            }
            Write-Host ("  [{0:HH:mm}] progress  {1}" -f (Get-Date), ($line -join '  ')) -ForegroundColor DarkGray
        }
    }
    foreach ($x in $w) { if (-not $x.End) { $x.End = Get-Date } }

    $res = @{}
    foreach ($x in $w) {
        $lines = @(); try { $lines = @(Get-Content $x.Out -ErrorAction Stop) } catch {}
        $inSum = $false
        foreach ($l in $lines) {
            if ($l -match '=== SUMMARY ===') { $inSum = $true; continue }
            if (-not $inSum) { continue }
            $m = [regex]::Match($l, '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d (.*)$'); if (-not $m.Success) { continue }
            $body = $m.Groups[1].Value
            if ($body -match '^(File|----|===)') { continue }
            # Format-Table pads columns to the widest value, so a short name can sit ONE space from the numbers.
            $ok = [regex]::Match($body, '^(.+?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+OK\s*$')
            if ($ok.Success) {
                $res[$ok.Groups[1].Value.Trim()] = @{ Status = 'OK'; Main = [int]$ok.Groups[2].Value; DB2Main = [int]$ok.Groups[3].Value; Bottom = [int]$ok.Groups[4].Value; DB2Bottom = [int]$ok.Groups[5].Value }
                continue
            }
            $bad = [regex]::Match($body, '^(.+?)\s+\d+\s+(?:\d+\s+)?\d+\s+(?:\d+\s+)?FAIL')
            if ($bad.Success) { $res[$bad.Groups[1].Value.Trim()] = @{ Status = 'FAIL' } }
        }
    }
    foreach ($x in $w) { foreach ($n in $x.Names) { if (-not $res.ContainsKey($n)) { $res[$n] = @{ Status = 'FAIL' } } } }   # no summary row = failed
    return @{ Results = $res; Workers = $w }
}

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

    # program list is read from Run_All_17.ps1 itself, so it can never drift
    $names = @(Get-Content "$ess\Run_All_17.ps1" | ForEach-Object { [regex]::Match($_, "File='([^']+)'") } | Where-Object { $_.Success } | ForEach-Object { $_.Groups[1].Value })
    if ($names.Count -ne 17) { throw "Expected 17 programs in Run_All_17.ps1, found $($names.Count)" }
    if ($Only.Count) { $names = @($names | Where-Object { $Only -contains $_ }); if (-not $names.Count) { throw "-Only matched no programs" } }

    $slow = @('z17 - 9175', 'FlashSystem', 'DS8K')     # the long jobs each get their own process + connection
    $groups = New-Object 'System.Collections.Generic.List[string[]]'
    if ($Sequential) { $groups.Add([string[]]$names) }
    else {
        $fast = @($names | Where-Object { $slow -notcontains $_ })
        if ($fast.Count) { $groups.Add([string[]]$fast) }
        foreach ($s in $slow) { if ($names -contains $s) { $groups.Add([string[]]@($s)) } }
    }

    New-Item $workDir -ItemType Directory -Force | Out-Null
    Remove-Item "$workDir\*" -Force -ErrorAction SilentlyContinue
    Write-Host ("[2/5] Running Run_All_17.ps1: {0} programs in {1} process(es)..." -f $names.Count, $groups.Count) -ForegroundColor White
    $r = Invoke-Workers $groups 'run'
    $status = $r.Results
    Write-Host "  timings:" -ForegroundColor DarkGray
    foreach ($x in $r.Workers) { Write-Host ("    worker {0}: {1,5:N0} min  ({2})" -f $x.N, ($x.End - $x.Start).TotalMinutes, $(if ($x.Names.Count -gt 3) { "$($x.Names.Count) programs" } else { $x.Names -join ', ' })) -ForegroundColor DarkGray }

    # A program that failed (e.g. DB2 connection dropped on a long query) is retried on its own, fresh connection.
    foreach ($attempt in 1..2) {
        $failed = @($names | Where-Object { $status[$_].Status -ne 'OK' })
        if (-not $failed.Count) { break }
        Write-Host ("`nRetry {0}/2 for: {1}" -f $attempt, ($failed -join ', ')) -ForegroundColor Yellow
        $rg = New-Object 'System.Collections.Generic.List[string[]]'
        if ($Sequential) { $rg.Add([string[]]$failed) } else { foreach ($f in $failed) { $rg.Add([string[]]@($f)) } }
        $rr = Invoke-Workers $rg "retry$attempt"
        foreach ($k in $rr.Results.Keys) { $status[$k] = $rr.Results[$k] }
    }

    # persist authoritative per-program results (merged with earlier runs for partial -Only refreshes)
    $qs = @{}
    if (Test-Path $qsF) { foreach ($e in @(Get-Content $qsF -Raw | ConvertFrom-Json | ForEach-Object { $_ })) { $qs[$e.program] = $e } }
    foreach ($n in $names) {
        $s = $status[$n]
        $qs[$n] = [pscustomobject]@{ program = $n; status = $s.Status; main = $s.Main; db2Main = $s.DB2Main; bottom = $s.Bottom; db2Bottom = $s.DB2Bottom; at = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
    }
    @($qs.Values) | ConvertTo-Json -Depth 3 | Set-Content $qsF -Encoding UTF8

    $failed = @($names | Where-Object { $status[$_].Status -ne 'OK' })
    if ($failed.Count) {
        Write-Host "`nStill failing after retries: $($failed -join ', ')" -ForegroundColor Red
        foreach ($f in $failed) { $bf = "$bk\$f.xlsx"; if (Test-Path $bf) { Copy-Item $bf $ess -Force; Write-Host "  restored previous $f.xlsx" -ForegroundColor Red } }
        throw "Query step failed for: $($failed -join ', ') - dashboard NOT rebuilt (other programs keep their fresh files, failed ones were restored)."
    }
    Write-Host ("[3/5] Run_All_17 finished: {0} programs saved and verified  ({1:N0} min)" -f $names.Count, ((Get-Date) - $t0).TotalMinutes) -ForegroundColor Green
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
