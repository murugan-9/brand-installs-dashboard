# =============================================================================
#  UPDATE_DASHBOARD_ESS.ps1  -  18-program refresh from the ESS folder
#  Source : C:\Users\MuruganVenugopal\.bob\ESS\*.xlsx   (17 query outputs + P11 (2025))
#  Output : dashboard-ess.html  (LOCAL PREVIEW ONLY)
#
#  This script NEVER touches dashboard-final.html / index.html and NEVER runs git.
#  Publishing is a separate, explicit step done only after the user has reviewed the result.
#
#  Usage:  .\UPDATE_DASHBOARD_ESS.ps1            (re-dumps all ESS xlsx first, then builds)
#          .\UPDATE_DASHBOARD_ESS.ps1 -SkipDump  (reuse the existing ess-staging dumps)
# =============================================================================
param([switch]$SkipDump)

$ErrorActionPreference = "Stop"
$dashDir  = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$stage    = "C:\Users\MuruganVenugopal\.bob\playground\ess-staging"
$template = "$dashDir\dashboard-ess.template.html"
$outFile  = "$dashDir\dashboard-ess.html"
$summary  = "$dashDir\ess-build-summary.json"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   ESS DASHBOARD BUILD (18 programs)    " -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# -- Step 1: dump ESS xlsx -> JSON staging ------------------------------------
if (-not $SkipDump) {
    Write-Host "Step 1/4  Dumping ESS xlsx files..." -ForegroundColor White
    & "$dashDir\dump-ess.ps1"
} else {
    Write-Host "Step 1/4  Skipped dump (using existing staging)" -ForegroundColor Yellow
}
if (-not (Test-Path $template)) { throw "Template missing: $template  (run build-ess-template.ps1)" }

# -- Helpers (same permanent logic as UPDATE_DASHBOARD.ps1, blankV also treats 01-01-0001 as no entitlement) ----
function esc($v) {
    return ([string]$v).Trim() -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' '
}
$script:dateFormats = [string[]]@(
    'yyyy-MM-dd', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:ss.fff', "yyyy-MM-ddTHH:mm:ss.fffZ", "yyyy-MM-ddTHH:mm:ssZ",
    'yyyy-MM-dd HH:mm:ss', 'dd-MM-yyyy', 'dd/MM/yyyy', 'dd-MM-yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm:ss'
)
function isBlankDate($s) {
    return ($s -eq "" -or $s -eq "0001-01-01")
}
function toDate($v) {
    if ($null -eq $v) { return $null }
    $s = ([string]$v).Trim()
    if (isBlankDate $s) { return $null }
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    if ($s -match '^\d{5,7}(\.\d+)?$') {
        try { return [datetime]::FromOADate([double]::Parse($s, $inv)) } catch { return $null }
    }
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    $out = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, $script:dateFormats, $inv, $styles, [ref]$out)) { return $out }
    return $null
}
function fmtDate($v) {
    $d = toDate $v
    if ($null -ne $d) { return $d.ToString("yyyy-MM-dd") }
    if ($null -eq $v -or (isBlankDate ([string]$v).Trim())) { return "" }
    return ([string]$v).Trim()
}
function parseDate($v) { return toDate $v }
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
function blankV($v) {
    if ($null -eq $v) { return $true }
    $s = ([string]$v).Trim()
    return ((isBlankDate $s) -or $s -eq "01-01-0001")
}
function deriveComment($vals) {
    $status = ([string]($vals["STATUS"])).Trim().ToUpper()
    $ibc    = ([string]($vals["IBM/CLIENT"])).Trim().ToUpper()
    if ($ibc -eq "IBM") { return "IBM" }
    $hasEnt = (-not (blankV $vals["SVCLVC"])) -or (-not (blankV $vals["SVCLVD"])) -or
              (-not (blankV $vals["SVCSTA"])) -or (-not (blankV $vals["SVCEND"]))
    if ($status -eq "IN" -and (-not (blankV $vals["SVCEND"]))) {
        $svcend = parseDate $vals["SVCEND"]
        if ($null -ne $svcend -and $svcend -lt (Get-Date)) { return "Contract Expired" }
    }
    if ($hasEnt) { return "Complete" }
    if ($status -eq "SH") {
        $shipd = parseDate $vals["SHIPD"]
        if ($null -ne $shipd -and $shipd.Year -lt 2025) { return "Invalid" }
        return "Not Installed"
    }
    if ($status -eq "IN") {
        $instd = parseDate $vals["INSTD"]
        if ($null -ne $instd -and $instd.Year -lt 2025) { return "Invalid" }
        return "No Entitlement"
    }
    return ""
}

function Load-Json($path) {
    if (-not (Test-Path $path)) { throw "Not found: $path" }
    return (Get-Content $path -Raw | ConvertFrom-Json)
}

# Returns @{ ROWS = '[[...],...]'; COUNT = n; COMMENTS = @{...} }
function Build-Rows($ess, $svl) {
    $eH = $ess.headers; $sH = $svl.headers
    $eR = $ess.rows;    $sR = $svl.rows
    $eI = @{}; for ($i=0;$i -lt $eH.Count;$i++){$eI[$eH[$i]]=$i}
    $sI = @{}; for ($i=0;$i -lt $sH.Count;$i++){$sI[$sH[$i]]=$i}

    $svlMap = @{}
    if ($sH.Count -gt 0) {
        $serIdx = $sI["SERIAL"]
        foreach ($r in $sR) {
            $s = ([string]$r[$serIdx]).Trim()
            if ($s) {
                if (-not $svlMap.ContainsKey($s)){$svlMap[$s]=[System.Collections.Generic.List[object]]::new()}
                $svlMap[$s].Add($r)
            }
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
    $cmts     = @{}

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
        $cm = deriveComment $res
        if ($cmts.ContainsKey($cm)) { $cmts[$cm]++ } else { $cmts[$cm] = 1 }
        $parts.Add('"'+(esc $cm)+'"'); $parts.Add('"'+$ibc+'"')
        $compact.Add('['+($parts -join ',')+']')
    }
    return @{ ROWS = '['+($compact -join ',')+']'; COUNT = $compact.Count; COMMENTS = $cmts }
}

# -- Program registry: one entry per ESS xlsx --------------------------------
# Legacy brands keep their original keys/classes so existing views and CSS are unchanged.
$legacy = @{
    'P11 (2025)'          = @{ key='P11';      cls='bp-p11'      }
    'FlashSystem'         = @{ key='FS7600';   cls='bp-fs'       }
    'z Mid Range (2026)'  = @{ key='ZMID';     cls='bp-zmid'     }
    'P11 Balcones (2026)' = @{ key='BALCONES'; cls='bp-balcones' }
}

# Label shown on the dashboard when it differs from the source file/folder name (the DB2 query, staging and checks keep the source name)
$displayNames = @{ 'FlashSystem C200' = 'FlashSystem C300' }
Write-Host "Step 2/4  Discovering programs in staging..." -ForegroundColor White
$dirs = Get-ChildItem $stage -Directory | Sort-Object Name
$progs = @()
$hue = 0; $i = 0
foreach ($d in $dirs) {
    $name = $d.Name
    if ($legacy.ContainsKey($name)) { $key = $legacy[$name].key; $cls = $legacy[$name].cls; $css = "" }
    else {
        $key = 'E' + ($name -replace '[^A-Za-z0-9]','').ToUpper()
        $cls = 'bp-' + ($name -replace '[^A-Za-z0-9]','').ToLower()
        $h = [int](($i * 47 + 20) % 360)
        $css = ".brand-pill.$cls{background:hsla($h,70%,55%,.15);border-color:hsl($h,70%,38%);color:hsl($h,90%,78%)}"
    }
    $i++
    $disp = if ($displayNames.ContainsKey($name)) { $displayNames[$name] } else { $name }
    $progs += [pscustomobject]@{ name=$name; disp=$disp; key=$key; cls=$cls; css=$css; dir=$d.FullName; file=($disp -replace '[^A-Za-z0-9]+','_').Trim('_') }
}
# P11 first (user's headline program) then alphabetical
$progs = @($progs | Sort-Object @{e={ if ($_.key -eq 'P11') {0} else {1} }}, name)
Write-Host "  $($progs.Count) programs" -ForegroundColor Green

# -- Step 3: build rows ------------------------------------------------------
Write-Host "Step 3/4  Building rows..." -ForegroundColor White
$sbRows   = New-Object System.Text.StringBuilder
$brandsJs = [System.Collections.Generic.List[string]]::new()
$opts     = [System.Collections.Generic.List[string]]::new()
$css      = [System.Collections.Generic.List[string]]::new()
$sum      = [System.Collections.Generic.List[object]]::new()
$n = 0
foreach ($p in $progs) {
    $n++
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ess = Load-Json "$($p.dir)\ESS_Installs.json"
    $svlPath = "$($p.dir)\SNs_Config_SVL.json"
    $svl = if (Test-Path $svlPath) { Load-Json $svlPath } else { [pscustomobject]@{ headers=@(); rows=@() } }
    $d = Build-Rows $ess $svl
    if ($sbRows.Length -gt 0) { [void]$sbRows.Append(',') }
    [void]$sbRows.Append('"' + $p.key + '":' + $d.ROWS)
    $brandsJs.Add('{"key":"' + $p.key + '","label":"' + (esc $p.disp) + '","cls":"' + $p.cls + '","file":"' + (esc $p.file) + '"}')
    $lblHtml = [System.Net.WebUtility]::HtmlEncode($p.disp)
    $opts.Add('<div class="cmt-opt"><input type="radio" name="brand" id="brx' + $n + '" class="brk" value="' + $p.key + '"><label for="brx' + $n + '">' + $lblHtml + '</label></div>')
    if ($p.css) { $css.Add($p.css) }
    $sum.Add([pscustomobject]@{ program=$p.name; key=$p.key; essRows=$ess.rows.Count; svlRows=$svl.rows.Count; outRows=$d.COUNT; comments=$d.COMMENTS })
    Write-Host ("  [{0}/{1}] {2,-24} ESS={3,5} SVL={4,6} -> rows={5,5}  ({6:N0}s)" -f $n,$progs.Count,$p.name,$ess.rows.Count,$svl.rows.Count,$d.COUNT,$sw.Elapsed.TotalSeconds) -ForegroundColor Green
    $ess = $null; $svl = $null; $d = $null
    [GC]::Collect()
}

# -- Step 4: fill template, write local output -------------------------------
Write-Host "Step 4/4  Writing dashboard-ess.html (local only)..." -ForegroundColor White
$html = [System.IO.File]::ReadAllText($template, [System.Text.Encoding]::UTF8)
# Last refresh date shown under the title: when the DB2 data was last pulled (ess-query-summary.json), else newest ESS xlsx
$lastRefresh = $null
$qsum = "$dashDir\ess-query-summary.json"
if (Test-Path $qsum) {
    $ats = @(Get-Content $qsum -Raw | ConvertFrom-Json | ForEach-Object { $_ } | ForEach-Object { $_.at } | Where-Object { $_ })
    if ($ats.Count) { $lastRefresh = ($ats | ForEach-Object { [datetime]$_ } | Measure-Object -Maximum).Maximum }
}
if (-not $lastRefresh) { $lastRefresh = (Get-ChildItem 'C:\Users\MuruganVenugopal\.bob\ESS\*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime }
$dn = $lastRefresh.Day
$sfx = if ($dn -in 11,12,13) { 'th' } else { switch ($dn % 10) { 1 {'st'} 2 {'nd'} 3 {'rd'} default {'th'} } }
$lastRefreshTxt = "$dn$sfx " + $lastRefresh.ToString('MMM', [System.Globalization.CultureInfo]::InvariantCulture)
Write-Host "  Last refresh date shown: $lastRefreshTxt" -ForegroundColor White
$html = $html.Replace('{{PROG_COUNT}}', "$($progs.Count)")
$html = $html.Replace('{{LAST_REFRESH}}', $lastRefreshTxt)
$html = $html.Replace('/*{{BRAND_CSS}}*/', ($css -join "`n"))
$html = $html.Replace('{{BRAND_OPTIONS}}', ($opts -join "`n        "))
if (-not $html.Contains('var BRANDS=[];') -or -not $html.Contains('var BRAND_ROWS={};')) { throw "BRANDS/BRAND_ROWS placeholders not found in template" }
$html = $html.Replace('var BRANDS=[];', 'var BRANDS=[' + ($brandsJs -join ',') + '];')
$html = $html.Replace('var BRAND_ROWS={};', 'var BRAND_ROWS={' + $sbRows.ToString() + '};')
[System.IO.File]::WriteAllText($outFile, $html, (New-Object System.Text.UTF8Encoding($false)))
$sum | ConvertTo-Json -Depth 5 | Set-Content $summary -Encoding UTF8
$mb = [math]::Round((Get-Item $outFile).Length / 1MB, 2)
Write-Host ""
Write-Host "DONE. $outFile  ($mb MB)" -ForegroundColor Cyan
Write-Host "Nothing was committed or pushed. dashboard-final.html / index.html are untouched." -ForegroundColor Cyan
