# =============================================================================
#  check-ess.ps1  -  automated validation of dashboard-ess.html
#  Exit code 0 = all checks passed, 1 = at least one FAIL.  Writes ess-check-report.json.
#  Never modifies the dashboard, never touches git.
# =============================================================================
$ErrorActionPreference = 'Stop'
$dir      = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$essLog   = "C:\Users\MuruganVenugopal\.bob\ESS\Run_All_17.log"
$dashFile = "$dir\dashboard-ess.html"
$sumFile  = "$dir\ess-build-summary.json"
$report   = "$dir\ess-check-report.json"

$fails = [System.Collections.Generic.List[string]]::new()
function Bad($m) { $script:fails.Add($m); Write-Host "  FAIL: $m" -ForegroundColor Red }
function Good($m) { Write-Host "  OK:   $m" -ForegroundColor Green }

Write-Host "`n=== CHECKS ===" -ForegroundColor Cyan
if (-not (Test-Path $dashFile)) { throw "dashboard-ess.html not found" }
if (-not (Test-Path $sumFile))  { throw "ess-build-summary.json not found" }
$sum = @(Get-Content $sumFile -Raw | ConvertFrom-Json | ForEach-Object { $_ })

# 1. Every program: dashboard rows == ESS sheet rows
Write-Host "1. Row counts: dashboard vs ESS sheet" -ForegroundColor White
$bad1 = @($sum | Where-Object { $_.outRows -ne $_.essRows })
if ($bad1.Count) { $bad1 | ForEach-Object { Bad "$($_.program): dashboard $($_.outRows) vs ESS $($_.essRows)" } } else { Good "all $($sum.Count) programs match" }

# 2. Cross-check the xlsx contents against the DB2 counts recorded when the files were generated.
#    Preferred source: ess-query-summary.json (written by REFRESH_ESS.ps1 from each worker's own output - safe
#    with parallel workers). Fallback: the cumulative Run_All_17.log (latest summary row per program).
$qsFile = "$dir\ess-query-summary.json"
$latest = [ordered]@{}   # name -> @{ status; main; db2Main; bottom; db2Bottom }
if (Test-Path $qsFile) {
    Write-Host "2. ESS files vs DB2 counts (ess-query-summary.json)" -ForegroundColor White
    foreach ($e in @(Get-Content $qsFile -Raw | ConvertFrom-Json | ForEach-Object { $_ })) {
        $latest[$e.program] = @{ status = $e.status; main = $e.main; db2Main = $e.db2Main; bottom = $e.bottom; db2Bottom = $e.db2Bottom }
    }
} else {
    Write-Host "2. ESS files vs DB2 counts in Run_All_17.log" -ForegroundColor White
    $rx = '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d (.+?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(OK|FAIL.*)$'
    foreach ($l in @(Get-Content $essLog)) {
        $m = [regex]::Match($l, $rx); if (-not $m.Success) { continue }
        $latest[$m.Groups[1].Value.Trim()] = @{ status = $(if ($m.Groups[6].Value -eq 'OK') { 'OK' } else { 'FAIL' }); main = [int]$m.Groups[2].Value; db2Main = [int]$m.Groups[3].Value; bottom = [int]$m.Groups[4].Value; db2Bottom = [int]$m.Groups[5].Value }
    }
}
if ($latest.Count -eq 0) { Bad "no DB2 count records found" }
else {
    $n = 0
    foreach ($name in $latest.Keys) {
        $q = $latest[$name]; $n++
        $p = $sum | Where-Object { $_.program -eq $name } | Select-Object -First 1
        if (-not $p) { Bad "'$name' was queried but no dump exists"; continue }
        if ($q.status -ne 'OK')                       { Bad "$name : query status $($q.status)" }
        elseif ([int]$q.db2Bottom -ne $p.essRows)     { Bad "$name : DB2 bottom $($q.db2Bottom) vs file $($p.essRows)" }
        elseif ([int]$q.db2Main -ne $p.svlRows)       { Bad "$name : DB2 main $($q.db2Main) vs file $($p.svlRows)" }
        elseif ([int]$q.main -ne [int]$q.db2Main -or [int]$q.bottom -ne [int]$q.db2Bottom) { Bad "$name : fetched rows differ from DB2 COUNT(*)" }
    }
    if ($n -ne 17) { Bad "records for $n programs, expected 17" } else { Good "17 programs match DB2 counts (P11 2025 is uploaded manually)" }
}

# 3. Parse dashboard data
Write-Host "3. Dashboard data integrity" -ForegroundColor White
$html = [IO.File]::ReadAllText($dashFile)
$m = [regex]::Match($html, 'var BRAND_ROWS=\{(.*?)\};\r?\n')
if (-not $m.Success) { throw "BRAND_ROWS not found in dashboard-ess.html" }
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = [int]::MaxValue; $ser.RecursionLimit = 100
$obj = $ser.DeserializeObject('{' + $m.Groups[1].Value + '}')
if ($obj.Count -ne $sum.Count) { Bad "dashboard has $($obj.Count) programs, build had $($sum.Count)" } else { Good "$($obj.Count) programs in dashboard" }
$allowed = @('Complete','IBM','Not Installed','No Entitlement','Contract Expired','Invalid','')
$ctys = @{}; $badDate = 0; $badCols = 0; $badCmt = @{}; $total = 0; $emptyCmt = 0
foreach ($k in $obj.Keys) {
    foreach ($r in $obj[$k]) {
        $total++
        if ($r.Count -ne 28) { $badCols++; continue }
        foreach ($i in 16,17,18,21,22) { $v = [string]$r[$i]; if ($v -ne '' -and $v -notmatch '^\d{4}-\d{2}-\d{2}$') { $badDate++ } }
        $c = [string]$r[26]; if ($allowed -notcontains $c) { $badCmt[$c] = 1 }; if ($c -eq '') { $emptyCmt++ }
        $ctys[[string]$r[3]] = 1
    }
}
if ($badCols) { Bad "$badCols rows without 28 columns" } else { Good "all $total rows have 28 columns" }
if ($badDate) { Bad "$badDate date cells not in yyyy-MM-dd" } else { Good "all date cells are yyyy-MM-dd (or empty)" }
if ($badCmt.Count) { Bad ("unexpected comment values: " + ($badCmt.Keys -join ', ')) } else { Good "comment values all valid ($emptyCmt rows with blank comment)" }
$cm = [regex]::Match($html, 'var CTY_FALLBACK=\{(.*?)\n\};', 'Singleline')
$known = @{}; [regex]::Matches($cm.Groups[1].Value, '(\d+)\s*:') | ForEach-Object { $known[$_.Groups[1].Value] = 1 }
$missing = @($ctys.Keys | Where-Object { $_ -and -not $known.ContainsKey($_) } | Sort-Object)
if ($missing.Count) { Bad ("IBMCTY codes missing from CTY_FALLBACK (red banner would show): " + ($missing -join ',')) } else { Good "no missing IBMCTY codes ($($known.Count) mapped)" }

# 4. Headless-browser load test (every brand + every tab)
Write-Host "4. Headless Edge load test" -ForegroundColor White
$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
if (-not (Test-Path $edge)) { Write-Host "  SKIP: Edge not found" -ForegroundColor Yellow }
else {
    $test = @'
<script>
window.__errs=[];window.addEventListener('error',function(e){window.__errs.push(String(e.message)+' @'+e.lineno);});
window.addEventListener('load',function(){setTimeout(function(){
 var out=[];var rs=document.querySelectorAll('.brk');
 for(var i=0;i<rs.length;i++){try{rs[i].checked=true;rs[i].dispatchEvent(new Event('change',{bubbles:true}));
   out.push(rs[i].value+' | '+document.getElementById('hdrRecs').textContent);}catch(e){out.push(rs[i].value+' ERR '+e.message);}}
 var tabs=document.querySelectorAll('.tab');for(var t=0;t<tabs.length;t++){try{tabs[t].click();out.push('tab'+t+' ok');}catch(e){out.push('tab'+t+' ERR '+e.message);}}
 var d=document.createElement('pre');d.id='testout';d.textContent='RADIOS='+rs.length+'\n'+out.join('\n')+'\nERRORS='+window.__errs.length+'\n'+window.__errs.join('\n');document.body.appendChild(d);
},500);});
</script></body>
'@
    $tmp = "$dir\ess-test.tmp.html"
    $i = $html.LastIndexOf('</body>')
    [IO.File]::WriteAllText($tmp, $html.Substring(0, $i) + $test + $html.Substring($i + 7), (New-Object Text.UTF8Encoding($false)))
    $ErrorActionPreference = 'Continue'
    try {
        $dom = & $edge --headless --disable-gpu --virtual-time-budget=20000 --dump-dom ("file:///" + ($tmp -replace '\\','/')) 2>$null | Out-String
    } finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue; $ErrorActionPreference = 'Stop' }
    $o = [regex]::Match($dom, '<pre id="testout">(.*?)</pre>', 'Singleline')
    if (-not $o.Success) { Bad "browser test produced no output" }
    else {
        $txt = $o.Groups[1].Value
        if ($txt -notmatch 'ERRORS=0') { Bad "JavaScript errors: $(($txt -split 'ERRORS=')[1])" } else { Good "0 JavaScript errors" }
        if ($txt -match 'ERR ') { Bad "a brand/tab raised an exception" }
        if ($txt -notmatch "RADIOS=$($sum.Count + 1)\b") { Bad "expected $($sum.Count + 1) brand options" } else { Good "$($sum.Count + 1) brand options (All + $($sum.Count) programs)" }
        $tot = ($sum | Measure-Object outRows -Sum).Sum
        foreach ($p in $sum) {
            $mm = [regex]::Match($txt, [regex]::Escape($p.key) + ' \| ([\d,]+) records')
            if (-not $mm.Success -or [int]($mm.Groups[1].Value -replace ',','') -ne $p.outRows) { Bad "$($p.program): header shows '$($mm.Groups[1].Value)' expected $($p.outRows)" }
        }
        $ma = [regex]::Match($txt, 'ALL \| ([\d,]+) records')
        if (-not $ma.Success -or [int]($ma.Groups[1].Value -replace ',','') -ne $tot) { Bad "All Brands shows '$($ma.Groups[1].Value)' expected $tot" } else { Good "header counts correct for every brand (All = $tot)" }
    }
}

$pass = ($fails.Count -eq 0)
[pscustomobject]@{
    pass = $pass; checkedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    dashboardTicks = (Get-Item $dashFile).LastWriteTimeUtc.Ticks
    totalRows = $total; failures = @($fails)
} | ConvertTo-Json | Set-Content $report -Encoding UTF8
Write-Host ""
if ($pass) { Write-Host "ALL CHECKS PASSED" -ForegroundColor Green; exit 0 }
else       { Write-Host "$($fails.Count) CHECK(S) FAILED - do not publish" -ForegroundColor Red; exit 1 }
