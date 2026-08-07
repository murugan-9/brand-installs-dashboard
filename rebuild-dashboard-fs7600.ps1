
# ── Paths ─────────────────────────────────────────────────────────────────────
$FS_DUMP_DIR = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data\.bob\tmp\xlsx-dumps\FLASHSYSTEM - 7th Aug-88a95f8558463365"
$OUT_DIR     = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"

# ── Load both sheets ──────────────────────────────────────────────────────────
$essSheet = Get-Content "$FS_DUMP_DIR\ESS_Installs.json"      -Raw | ConvertFrom-Json
$svlSheet = Get-Content "$FS_DUMP_DIR\SNs_Config____SVL.json" -Raw | ConvertFrom-Json

$essHdr  = $essSheet.headers
$svlHdr  = $svlSheet.headers
$essRows = $essSheet.rows
$svlRows = $svlSheet.rows

$eIdx = @{}; for ($i=0;$i -lt $essHdr.Count;$i++){$eIdx[$essHdr[$i]]=$i}
$sIdx = @{}; for ($i=0;$i -lt $svlHdr.Count;$i++){$sIdx[$svlHdr[$i]]=$i}

Write-Host "ESS headers ($($essHdr.Count)): $($essHdr -join ', ')"
Write-Host "SVL headers ($($svlHdr.Count)): $($svlHdr -join ', ')"
Write-Host "ESS rows: $($essRows.Count)   SVL rows: $($svlRows.Count)"

# ── Helpers ───────────────────────────────────────────────────────────────────
function Inc($ht,$k){ if($ht.ContainsKey($k)){$ht[$k]++}else{$ht[$k]=1} }

function fmtDate($v){
    if($null -eq $v -or [string]$v -eq "" -or [string]$v -eq "0001-01-01"){return ""}
    try{return ([datetime]::Parse([string]$v)).ToString("yyyy-MM-dd")}catch{return [string]$v}
}

function parseDate($v){
    if($null -eq $v -or [string]$v -eq "" -or [string]$v -eq "0001-01-01"){return $null}
    try{return [datetime]::Parse([string]$v)}catch{return $null}
}

function esc($v){ return ([string]$v).Trim() -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' }

function deriveIbmClient($cname){
    $cu = ([string]$cname).Trim().ToUpper()
    if ($cu -match '\bIBM\b' -or $cu -match 'INTERNATIONAL BUSINESS MACHINE') { return "IBM" }
    return "CLIENT"
}

function svclvdPriority($svclvd){
    $v = ([string]$svclvd).Trim().ToUpper()
    if ($v -match '^2H COMMITTED FIX')  { return 0 }
    if ($v -match '^4H COMMITTED FIX')  { return 1 }
    if ($v -match '^6H COMMITTED FIX')  { return 2 }
    if ($v -match '^8H COMMITTED FIX')  { return 3 }
    if ($v -match '^12H COMMITTED FIX') { return 4 }
    if ($v -match '^24H COMMITTED FIX') { return 5 }
    if ($v -match '^72H COMMITTED FIX') { return 6 }
    if ($v -match '^SD SAME DAY')        { return 7 }
    if ($v -match '^ON-SITE REPAIR.*ORT=SD') { return 8 }
    if ($v -match '^NBD NEXT BUSINESS DAY') { return 9 }
    if ($v -match '^ON-SITE REPAIR.*ORT=NBD') { return 10 }
    if ($v -eq ',')                      { return 11 }
    return 99
}

Write-Host "Building SVL lookup by SERIAL..."
$svlBySerial = @{}
foreach ($row in $svlRows) {
    $serial = ([string]$row[$sIdx["SERIAL"]]).Trim()
    if ($serial -eq "") { continue }
    if (-not $svlBySerial.ContainsKey($serial)) {
        $svlBySerial[$serial] = [System.Collections.Generic.List[object]]::new()
    }
    $svlBySerial[$serial].Add($row)
}
Write-Host "Unique SERIALs in SVL: $($svlBySerial.Count)"

function pickBestSvlRow($rows) {
    if ($null -eq $rows -or $rows.Count -eq 0) { return $null }
    $best      = $null
    $bestPri   = 999
    foreach ($row in $rows) {
        $pri = svclvdPriority $row[$sIdx["SVCLVD"]]
        if ($pri -eq 99) { continue }
        if ($pri -lt $bestPri) { $bestPri = $pri; $best = $row }
    }
    return $best
}

function deriveComment($vals) {
    $status = ([string]($vals["STATUS"])).Trim().ToUpper()
    $ibc    = ([string]($vals["IBM/CLIENT"])).Trim().ToUpper()
    if ($ibc -eq "IBM") { return "IBM" }
    function blankV($v){ return ($null -eq $v -or ([string]$v).Trim() -eq "" -or ([string]$v).Trim() -eq "0001-01-01") }
    $hasEnt = (-not (blankV $vals["SVCLVC"])) -or (-not (blankV $vals["SVCLVD"])) -or `
              (-not (blankV $vals["SVCSTA"])) -or (-not (blankV $vals["SVCEND"]))
    if ($status -eq "IN" -and (-not (blankV $vals["SVCEND"]))) {
        $svcend = parseDate $vals["SVCEND"]
        if ($null -ne $svcend -and $svcend -lt (Get-Date)) { return "Contract Expired" }
    }
    $shipd = parseDate $vals["SHIPD"]
    $instd = parseDate $vals["INSTD"]
    $oldDate = ($null -ne $shipd -and $shipd.Year -lt 2025) -or `
               ($null -ne $instd -and $instd.Year -lt 2025)
    if ($oldDate -and (-not $hasEnt)) { return "Invalid" }
    if ($hasEnt) { return "Complete" }
    if ($status -eq "IN" -and (-not $hasEnt)) { return "No Entitlement" }
    if ($status -eq "SH" -and (-not $hasEnt)) { return "Not Installed" }
    return ""
}

$COLS        = @("ILPRID","SERIAL","STATUS","IBMCTY","Country","Region","GEO","MDL","PRDTYP","ISO","CNAME","CITY","ZIP","ADDR","RGN","INSTT","INSTD","SHIPD","WED","SVCLVD","SVCLVC","SVCSTA","SVCEND","CONTRNO","BGNAM","TOM")
$dateCols    = @("INSTD","SHIPD","WED","SVCSTA","SVCEND")
$numCols     = @("IBMCTY","PRDTYP")
$missingCols = @("Country","Region","GEO")
$svlCols     = @("SVCLVD","SVCLVC","SVCSTA","SVCEND","CONTRNO")

Write-Host "Building compact rows..."
$compactRows = [System.Collections.Generic.List[string]]::new()

foreach ($row in $essRows) {
    $serial  = ([string]$row[$eIdx["SERIAL"]]).Trim()
    $bestSvl = if ($svlBySerial.ContainsKey($serial)) { pickBestSvlRow $svlBySerial[$serial] } else { $null }

    $resolved = @{}
    foreach ($col in $COLS) {
        if ($missingCols -contains $col) {
            $resolved[$col] = ""
        } elseif ($svlCols -contains $col) {
            $resolved[$col] = if ($null -ne $bestSvl -and $sIdx.ContainsKey($col)) { $bestSvl[$sIdx[$col]] } else { "" }
        } elseif ($eIdx.ContainsKey($col)) {
            $resolved[$col] = $row[$eIdx[$col]]
        } else {
            $resolved[$col] = ""
        }
    }

    $ibcOut = deriveIbmClient $resolved["CNAME"]
    $resolved["IBM/CLIENT"] = $ibcOut
    $resolved["INSTD"]      = if ($eIdx.ContainsKey("INSTD")) { $row[$eIdx["INSTD"]] } else { "" }
    $resolved["SHIPD"]      = if ($eIdx.ContainsKey("SHIPD")) { $row[$eIdx["SHIPD"]] } else { "" }
    $resolved["STATUS"]     = if ($eIdx.ContainsKey("STATUS")) { $row[$eIdx["STATUS"]] } else { "" }

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($col in $COLS) {
        if ($missingCols -contains $col) {
            $parts.Add('""')
        } elseif ($dateCols -contains $col) {
            $parts.Add('"' + (fmtDate $resolved[$col]) + '"')
        } elseif ($numCols -contains $col) {
            $n = 0
            if ([int]::TryParse([string]$resolved[$col],[ref]$n)){$parts.Add([string]$n)}else{$parts.Add('"'+(esc $resolved[$col])+'"')}
        } else {
            $parts.Add('"'+(esc $resolved[$col])+'"')
        }
    }
    $parts.Add('"'+(esc (deriveComment $resolved))+'"')
    $parts.Add('"'+$ibcOut+'"')
    $compactRows.Add('['+($parts -join ',')+']')
}

$ctyIN=@{};$ctySH=@{}
$ctyMtIN=@{};$ctyMtSH=@{}
$ctyMtCustSvcIN=@{};$ctyMtCustSvcSH=@{}

foreach ($row in $essRows) {
    $st      = ([string]$row[$eIdx["STATUS"]]).Trim()
    $cty     = ([string]$row[$eIdx["IBMCTY"]]).Trim()
    $mt      = ([string]$row[$eIdx["PRDTYP"]]).Trim()
    $cu      = ([string]$row[$eIdx["CNAME"]]).Trim()
    $serial  = ([string]$row[$eIdx["SERIAL"]]).Trim()
    $bestSvl = if ($svlBySerial.ContainsKey($serial)) { pickBestSvlRow $svlBySerial[$serial] } else { $null }
    $sv      = if ($null -ne $bestSvl -and $sIdx.ContainsKey("SVCLVD")) { ([string]$bestSvl[$sIdx["SVCLVD"]]).Trim() } else { "" }
    $k2 = "$cty|$mt";  $k3 = "$cty|$mt|$cu|$sv"
    if ($st -eq "IN") { Inc $ctyIN $cty; Inc $ctyMtIN $k2; Inc $ctyMtCustSvcIN $k3 }
    elseif ($st -eq "SH") { Inc $ctySH $cty; Inc $ctyMtSH $k2; Inc $ctyMtCustSvcSH $k3 }
}

$allCty = ($ctyIN.Keys+$ctySH.Keys)|Select-Object -Unique|Sort-Object {[int]$_}
$v1 = $allCty | ForEach-Object {
    $c=$_; $i=if($ctyIN.ContainsKey($c)){$ctyIN[$c]}else{0}; $s=if($ctySH.ContainsKey($c)){$ctySH[$c]}else{0}
    '{"c":'+$c+',"i":'+$i+',"s":'+$s+'}'
}

$allK2 = ($ctyMtIN.Keys+$ctyMtSH.Keys)|Select-Object -Unique|Sort-Object
$v2 = $allK2 | ForEach-Object {
    $k=$_; $p=$k -split '\|'; $c=$p[0]; $m=$p[1]
    $i=if($ctyMtIN.ContainsKey($k)){$ctyMtIN[$k]}else{0}; $s=if($ctyMtSH.ContainsKey($k)){$ctyMtSH[$k]}else{0}
    '{"c":'+$c+',"m":'+$m+',"i":'+$i+',"s":'+$s+'}'
}

$allK3 = ($ctyMtCustSvcIN.Keys+$ctyMtCustSvcSH.Keys)|Select-Object -Unique|Sort-Object
$v3 = $allK3 | ForEach-Object {
    $k=$_; $p=$k -split '\|'
    $c=$p[0]; $m=$p[1]; $cu=(esc $p[2]); $sv=(esc ($p[3..($p.Count-1)] -join '|'))
    $i=if($ctyMtCustSvcIN.ContainsKey($k)){$ctyMtCustSvcIN[$k]}else{0}
    $s=if($ctyMtCustSvcSH.ContainsKey($k)){$ctyMtCustSvcSH[$k]}else{0}
    '{"c":'+$c+',"m":'+$m+',"cu":"'+$cu+'","sv":"'+$sv+'","i":'+$i+',"s":'+$s+'}'
}

$v1Json   = '['+($v1  -join ',')+']'
$v2Json   = '['+($v2  -join ',')+']'
$v3Json   = '['+($v3  -join ',')+']'
$rowsJson = '['+($compactRows -join ',')+']'

Write-Host "V1:$($v1.Count)  V2:$($v2.Count)  V3:$($v3.Count)  Rows:$($compactRows.Count)"

$v1Json   | Set-Content "$OUT_DIR\v1_fs7600.json"    -Encoding UTF8 -NoNewline
$v2Json   | Set-Content "$OUT_DIR\v2_fs7600.json"    -Encoding UTF8 -NoNewline
$v3Json   | Set-Content "$OUT_DIR\v3_fs7600.json"    -Encoding UTF8 -NoNewline
$rowsJson | Set-Content "$OUT_DIR\rows2_fs7600.json" -Encoding UTF8 -NoNewline

Write-Host "All FS7600 data files written."
