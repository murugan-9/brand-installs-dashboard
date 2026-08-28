# =============================================================================
#  UPDATE_DASHBOARD.ps1  —  ONE-CLICK WEEKLY REFRESH  (fast JSON-dump method)
#  Drop new P11, FLASHSYSTEM, z Mid Range and P11 Balcones xlsx files into the
#  playground folder, then run this script. It does everything automatically:
#    1. Finds the pre-built JSON dumps for the latest brand xlsx files
#    2. Builds ALL_ROWS, V1/V2/V3 data for all brands (PowerShell, no Excel COM)
#    3. Injects the new data directly into dashboard-final.html
#    4. Copies dashboard-final.html to index.html
#    5. Commits and pushes to GitHub Pages
#
#  FAST because it reads JSON (already dumped by Bob) — NOT Excel COM.
#  Run time: ~30-60 seconds instead of 10+ minutes.
# =============================================================================

$ErrorActionPreference = "Stop"
$dashDir  = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$dumpBase = "$dashDir\.bob\tmp\xlsx-dumps"
$dashFile = "$dashDir\dashboard-final.html"
$indexFile= "$dashDir\index.html"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   BRAND INSTALLS DASHBOARD REFRESH    " -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 0: Backup current dashboard-final.html ───────────────────────────
$backupFile = "$dashDir\dashboard-final.bak_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
if (Test-Path $dashFile) {
    Copy-Item $dashFile $backupFile -Force
    Write-Host "Backup created: $(Split-Path $backupFile -Leaf)" -ForegroundColor DarkGray
    # Keep only the 2 most recent backups — delete older ones
    $oldBaks = Get-ChildItem "$dashDir\dashboard-final.bak_*.html" |
               Sort-Object LastWriteTime -Descending | Select-Object -Skip 2
    foreach ($b in $oldBaks) { Remove-Item $b.FullName -Force; Write-Host "  Removed old backup: $($b.Name)" -ForegroundColor DarkGray }
}
Write-Host ""

# ── Step 1: Find latest JSON dump folders ─────────────────────────────────
Write-Host "Step 1/5  Finding latest JSON dump folders..." -ForegroundColor White

# NOTE: P11 Balcones dumps start with "P11 Balcones" so we exclude them from the plain P11 match
$p11Dump      = Get-ChildItem "$dumpBase\P11 -*" -Directory -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $p11Dump) {
    $p11Dump  = Get-ChildItem "$dumpBase\P11*" -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notmatch '^P11 Balcones' } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
}
$fsDump       = Get-ChildItem "$dumpBase\FLASHSYSTEM*" -Directory -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
$zmidDump     = Get-ChildItem "$dumpBase\z Mid Range*" -Directory -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
$balconesDump = Get-ChildItem "$dumpBase\P11 Balcones*" -Directory -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $p11Dump)      { Write-Host "ERROR: No P11 dump found in $dumpBase" -ForegroundColor Red; Write-Host "Please ask Bob to dump the new P11 xlsx first." -ForegroundColor Yellow; pause; exit 1 }
if (-not $fsDump)       { Write-Host "ERROR: No FLASHSYSTEM dump found in $dumpBase" -ForegroundColor Red; Write-Host "Please ask Bob to dump the new FLASHSYSTEM xlsx first." -ForegroundColor Yellow; pause; exit 1 }
if (-not $zmidDump)     { Write-Host "WARNING: No z Mid Range dump found — ZMID data will be empty" -ForegroundColor Yellow }
if (-not $balconesDump) { Write-Host "WARNING: No P11 Balcones dump found — BALCONES data will be empty" -ForegroundColor Yellow }

Write-Host "  P11           : $($p11Dump.Name)"          -ForegroundColor Green
Write-Host "  FS7600        : $($fsDump.Name)"           -ForegroundColor Green
Write-Host "  z Mid Range   : $(if($zmidDump){$zmidDump.Name}else{'(not found)'})"     -ForegroundColor $(if($zmidDump){'Green'}else{'Yellow'})
Write-Host "  P11 Balcones  : $(if($balconesDump){$balconesDump.Name}else{'(not found)'})" -ForegroundColor $(if($balconesDump){'Green'}else{'Yellow'})
Write-Host ""

# ── Step 2: Load JSON sheets ───────────────────────────────────────────────
Write-Host "Step 2/5  Loading JSON data..." -ForegroundColor White

function Load-Sheet($dumpDir, $filename) {
    $path = "$dumpDir\$filename"
    if (-not (Test-Path $path)) { throw "Sheet file not found: $path" }
    return Get-Content $path -Raw | ConvertFrom-Json
}

$p11Ess = Load-Sheet $p11Dump.FullName "ESS_Installs.json"
$p11Svl = Load-Sheet $p11Dump.FullName "SNs_Config____SVL.json"
$fsEss  = Load-Sheet $fsDump.FullName  "ESS_Installs.json"
$fsSvl  = Load-Sheet $fsDump.FullName  "SNs_Config____SVL.json"

# z Mid Range — graceful empty fallback if dump not found
if ($zmidDump) {
    $zmidEss = Load-Sheet $zmidDump.FullName "ESS_Installs.json"
    $zmidSvl = Load-Sheet $zmidDump.FullName "SNs_Config____SVL.json"
} else {
    $zmidEss = [PSCustomObject]@{ headers=@(); rows=@() }
    $zmidSvl = [PSCustomObject]@{ headers=@(); rows=@() }
}

# P11 Balcones — graceful empty fallback if dump not found
if ($balconesDump) {
    $balconesEss = Load-Sheet $balconesDump.FullName "ESS_Installs.json"
    $balconesSvl = Load-Sheet $balconesDump.FullName "SNs_Config____SVL.json"
} else {
    $balconesEss = [PSCustomObject]@{ headers=@(); rows=@() }
    $balconesSvl = [PSCustomObject]@{ headers=@(); rows=@() }
}

Write-Host "  P11          ESS=$($p11Ess.rows.Count)  SVL=$($p11Svl.rows.Count)"          -ForegroundColor Green
Write-Host "  FS7600       ESS=$($fsEss.rows.Count)  SVL=$($fsSvl.rows.Count)"           -ForegroundColor Green
Write-Host "  z Mid Range  ESS=$($zmidEss.rows.Count)  SVL=$($zmidSvl.rows.Count)"       -ForegroundColor $(if($zmidDump){'Green'}else{'Yellow'})
Write-Host "  P11 Balcones ESS=$($balconesEss.rows.Count)  SVL=$($balconesSvl.rows.Count)" -ForegroundColor $(if($balconesDump){'Green'}else{'Yellow'})
Write-Host ""

# ── Step 3: Build dashboard data ──────────────────────────────────────────
Write-Host "Step 3/5  Building dashboard data..." -ForegroundColor White

function esc($v) {
    return ([string]$v).Trim() -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' '
}
function fmtDate($v) {
    if ($null -eq $v -or [string]$v -eq "" -or [string]$v -eq "0001-01-01") { return "" }
    try { return ([datetime]::Parse([string]$v)).ToString("yyyy-MM-dd") } catch { return [string]$v }
}
function parseDate($v) {
    if ($null -eq $v -or [string]$v -eq "" -or [string]$v -eq "0001-01-01") { return $null }
    try { return [datetime]::Parse([string]$v) } catch { return $null }
}
function deriveIbmClient($cname) {
    $cu = ([string]$cname).Trim().ToUpper()
    if ($cu -match '\bIBM\b' -or $cu -match 'INTERNATIONAL BUSINESS MACHINE') { return "IBM" }
    return "CLIENT"
}
function svclvdPriority($v) {
    $v = ([string]$v).Trim().ToUpper()
    if ($v -match '^2H COMMITTED FIX')        { return 0 }
    if ($v -match '^4H COMMITTED FIX')        { return 1 }
    if ($v -match '^6H COMMITTED FIX')        { return 2 }
    if ($v -match '^8H COMMITTED FIX')        { return 3 }
    if ($v -match '^12H COMMITTED FIX')       { return 4 }
    if ($v -match '^24H COMMITTED FIX')       { return 5 }
    if ($v -match '^48H COMMITTED FIX')       { return 6 }
    if ($v -match '^72H COMMITTED FIX')       { return 7 }
    if ($v -match '^SD SAME DAY')             { return 8 }
    if ($v -match '^ON-SITE REPAIR.*ORT=SD')  { return 9 }
    if ($v -match '^NBD NEXT BUSINESS DAY')   { return 10 }
    if ($v -match '^ON-SITE REPAIR.*ORT=NBD') { return 11 }
    if ($v -eq ',')                           { return 12 }
    return 99
}
function deriveComment($vals) {
    $status = ([string]($vals["STATUS"])).Trim().ToUpper()
    $ibc    = ([string]($vals["IBM/CLIENT"])).Trim().ToUpper()
    if ($ibc -eq "IBM") { return "IBM" }
    function blankV($v) { return ($null -eq $v -or ([string]$v).Trim() -eq "" -or ([string]$v).Trim() -eq "0001-01-01") }
    $hasEnt = (-not (blankV $vals["SVCLVC"])) -or (-not (blankV $vals["SVCLVD"])) -or
              (-not (blankV $vals["SVCSTA"])) -or (-not (blankV $vals["SVCEND"]))
    if ($status -eq "IN" -and (-not (blankV $vals["SVCEND"]))) {
        $svcend = parseDate $vals["SVCEND"]
        if ($null -ne $svcend -and $svcend -lt (Get-Date)) { return "Contract Expired" }
    }
    $shipd = parseDate $vals["SHIPD"]; $instd = parseDate $vals["INSTD"]
    $oldDate = ($null -ne $shipd -and $shipd.Year -lt 2025) -or ($null -ne $instd -and $instd.Year -lt 2025)
    if ($oldDate -and (-not $hasEnt)) { return "Invalid" }
    if ($hasEnt)                       { return "Complete" }
    if ($status -eq "IN")              { return "No Entitlement" }
    if ($status -eq "SH")              { return "Not Installed" }
    return ""
}

function Build-DashData($ess, $svl, $label) {
    $eH = $ess.headers; $sH = $svl.headers
    $eR = $ess.rows;    $sR = $svl.rows
    $eI = @{}; for ($i=0;$i -lt $eH.Count;$i++){$eI[$eH[$i]]=$i}
    $sI = @{}; for ($i=0;$i -lt $sH.Count;$i++){$sI[$sH[$i]]=$i}

    $svlMap = @{}
    foreach ($r in $sR) {
        $s = ([string]$r[$sI["SERIAL"]]).Trim()
        if ($s) {
            if (-not $svlMap.ContainsKey($s)){$svlMap[$s]=[System.Collections.Generic.List[object]]::new()}
            $svlMap[$s].Add($r)
        }
    }
    function getBest($rs) {
        if (-not $rs){return $null}; $b=$null; $bp=999
        foreach ($r in $rs){$p=svclvdPriority $r[$sI["SVCLVD"]];if($p -ne 99 -and $p -lt $bp){$bp=$p;$b=$r}}
        return $b
    }

    $COLS     = @("ILPRID","SERIAL","STATUS","IBMCTY","Country","Region","GEO","MDL","PRDTYP","ISO","CNAME","CITY","ZIP","ADDR","RGN","INSTT","INSTD","SHIPD","WED","SVCLVD","SVCLVC","SVCSTA","SVCEND","CONTRNO","BGNAM","TOM")
    $dateCols = @("INSTD","SHIPD","WED","SVCSTA","SVCEND")
    $numCols  = @("IBMCTY","PRDTYP")
    $missCols = @("Country","Region","GEO")
    $svlCols  = @("SVCLVD","SVCLVC","SVCSTA","SVCEND","CONTRNO")
    $compact  = [System.Collections.Generic.List[string]]::new()
    $ctyIN=@{};$ctySH=@{};$ctyMtIN=@{};$ctyMtSH=@{};$ciCustSvc=@{};$csCustSvc=@{}
    function Inc($h,$k){if($h.ContainsKey($k)){$h[$k]++}else{$h[$k]=1}}

    foreach ($row in $eR) {
        $serial  = ([string]$row[$eI["SERIAL"]]).Trim()
        $bestSvl = if ($svlMap.ContainsKey($serial)){getBest $svlMap[$serial]}else{$null}
        $res=@{}
        foreach ($col in $COLS) {
            if     ($missCols -contains $col) { $res[$col]="" }
            elseif ($svlCols  -contains $col) { $res[$col]=if($null -ne $bestSvl -and $sI.ContainsKey($col)){$bestSvl[$sI[$col]]}else{""} }
            elseif ($eI.ContainsKey($col))    { $res[$col]=$row[$eI[$col]] }
            else                              { $res[$col]="" }
        }
        $ibc=$res["IBM/CLIENT"]=deriveIbmClient $res["CNAME"]
        $res["INSTD"] =if($eI.ContainsKey("INSTD")){$row[$eI["INSTD"]]}else{""}
        $res["SHIPD"] =if($eI.ContainsKey("SHIPD")){$row[$eI["SHIPD"]]}else{""}
        $res["STATUS"]=if($eI.ContainsKey("STATUS")){$row[$eI["STATUS"]]}else{""}
        $parts=[System.Collections.Generic.List[string]]::new()
        foreach ($col in $COLS) {
            if     ($missCols  -contains $col) { $parts.Add('""') }
            elseif ($dateCols  -contains $col) { $parts.Add('"'+(fmtDate $res[$col])+'"') }
            elseif ($numCols   -contains $col) { $n=0;if([int]::TryParse([string]$res[$col],[ref]$n)){$parts.Add([string]$n)}else{$parts.Add('"'+(esc $res[$col])+'"')} }
            else                               { $parts.Add('"'+(esc $res[$col])+'"') }
        }
        $parts.Add('"'+(esc (deriveComment $res))+'"'); $parts.Add('"'+$ibc+'"')
        $compact.Add('['+($parts -join ',')+']')
        $st=([string]$row[$eI["STATUS"]]).Trim(); $cty=([string]$row[$eI["IBMCTY"]]).Trim()
        $mt=([string]$row[$eI["PRDTYP"]]).Trim(); $cu=([string]$row[$eI["CNAME"]]).Trim()
        $sv=if($null -ne $bestSvl -and $sI.ContainsKey("SVCLVD")){([string]$bestSvl[$sI["SVCLVD"]]).Trim()}else{""}
        if     ($st -eq "IN") { Inc $ctyIN $cty; Inc $ctyMtIN "$cty|$mt"; Inc $ciCustSvc "$cty|$mt|$cu|$sv" }
        elseif ($st -eq "SH") { Inc $ctySH $cty; Inc $ctyMtSH "$cty|$mt"; Inc $csCustSvc "$cty|$mt|$cu|$sv" }
    }
    $v1 = ($ctyIN.Keys+$ctySH.Keys) | Select-Object -Unique | Sort-Object {[int]$_} | ForEach-Object {
        $c=$_; $i=if($ctyIN.ContainsKey($c)){$ctyIN[$c]}else{0}; $s=if($ctySH.ContainsKey($c)){$ctySH[$c]}else{0}
        '{"c":'+$c+',"i":'+$i+',"s":'+$s+'}'
    }
    $v2 = ($ctyMtIN.Keys+$ctyMtSH.Keys) | Select-Object -Unique | Sort-Object | ForEach-Object {
        $k=$_; $p=$k -split '\|'
        $i=if($ctyMtIN.ContainsKey($k)){$ctyMtIN[$k]}else{0}; $s=if($ctyMtSH.ContainsKey($k)){$ctyMtSH[$k]}else{0}
        '{"c":'+$p[0]+',"m":'+$p[1]+',"i":'+$i+',"s":'+$s+'}'
    }
    $v3 = ($ciCustSvc.Keys+$csCustSvc.Keys) | Select-Object -Unique | Sort-Object | ForEach-Object {
        $k=$_; $p=$k -split '\|'
        $i=if($ciCustSvc.ContainsKey($k)){$ciCustSvc[$k]}else{0}; $s=if($csCustSvc.ContainsKey($k)){$csCustSvc[$k]}else{0}
        '{"c":'+$p[0]+',"m":'+$p[1]+',"cu":"'+(esc $p[2])+'","sv":"'+(esc ($p[3..($p.Count-1)] -join '|'))+'","i":'+$i+',"s":'+$s+'}'
    }
    Write-Host "  $label`: V1=$($v1.Count) V2=$($v2.Count) V3=$($v3.Count) Rows=$($compact.Count)" -ForegroundColor Green
    return @{ V1='['+($v1 -join ',')+']'; V2='['+($v2 -join ',')+']'; V3='['+($v3 -join ',')+']'; ROWS='['+($compact -join ',')+']' }
}

$p11Data      = Build-DashData $p11Ess      $p11Svl      "P11"
$fsData       = Build-DashData $fsEss       $fsSvl       "FS7600"
$zmidData     = Build-DashData $zmidEss     $zmidSvl     "ZMID"
$balconesData = Build-DashData $balconesEss $balconesSvl "BALCONES"
Write-Host ""

# ── Step 4: Inject into dashboard-final.html ──────────────────────────────
Write-Host "Step 4/5  Injecting data into dashboard-final.html..." -ForegroundColor White

# Line-based replacement handles JSON arrays with nested brackets reliably
$varMap = @{
    "ALL_ROWS"          = ("var ALL_ROWS="          + $p11Data.ROWS      + ";")
    "V1"                = ("var V1="                + $p11Data.V1        + ";")
    "V2"                = ("var V2="                + $p11Data.V2        + ";")
    "V3"                = ("var V3="                + $p11Data.V3        + ";")
    "ALL_ROWS_FS7600"   = ("var ALL_ROWS_FS7600="   + $fsData.ROWS       + ";")
    "V1_FS7600"         = ("var V1_FS7600="         + $fsData.V1         + ";")
    "V2_FS7600"         = ("var V2_FS7600="         + $fsData.V2         + ";")
    "V3_FS7600"         = ("var V3_FS7600="         + $fsData.V3         + ";")
    "ALL_ROWS_ZMID"     = ("var ALL_ROWS_ZMID="     + $zmidData.ROWS     + ";")
    "V1_ZMID"           = ("var V1_ZMID="           + $zmidData.V1       + ";")
    "V2_ZMID"           = ("var V2_ZMID="           + $zmidData.V2       + ";")
    "V3_ZMID"           = ("var V3_ZMID="           + $zmidData.V3       + ";")
    "ALL_ROWS_BALCONES" = ("var ALL_ROWS_BALCONES=" + $balconesData.ROWS + ";")
    "V1_BALCONES"       = ("var V1_BALCONES="       + $balconesData.V1   + ";")
    "V2_BALCONES"       = ("var V2_BALCONES="       + $balconesData.V2   + ";")
    "V3_BALCONES"       = ("var V3_BALCONES="       + $balconesData.V3   + ";")
}
$lines = [System.IO.File]::ReadAllLines($dashFile, [System.Text.Encoding]::UTF8)
$replaced = @{}
for ($i = 0; $i -lt $lines.Length; $i++) {
    foreach ($key in $varMap.Keys) {
        if ($lines[$i] -match ("^var " + $key + "=")) {
            $lines[$i] = $varMap[$key]
            $replaced[$key] = $true
        }
    }
}
foreach ($key in $varMap.Keys) {
    if (-not $replaced.ContainsKey($key)) { Write-Host ("  WARNING: var " + $key + " not found in HTML") -ForegroundColor Yellow }
}
[System.IO.File]::WriteAllLines($dashFile, $lines, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "  dashboard-final.html updated" -ForegroundColor Green
Copy-Item $dashFile $indexFile -Force
Write-Host "  index.html updated" -ForegroundColor Green
Write-Host ""

# ── Step 5: Commit and push ────────────────────────────────────────────────
Write-Host "Step 5/5  Publishing to GitHub Pages..." -ForegroundColor White
Set-Location $dashDir
$today = Get-Date -Format "yyyy-MM-dd HH:mm"
git add index.html dashboard-final.html
git commit -m "Dashboard refresh $today"
git push origin main

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  DONE! Dashboard live at:" -ForegroundColor Green
Write-Host "  https://murugan-9.github.io/brand-installs-dashboard/" -ForegroundColor Yellow
Write-Host "  Wait 1-2 mins then press Ctrl+Shift+R" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
pause
