$ErrorActionPreference = "Stop"
$dashDir  = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$dumpBase = "$dashDir\.bob\tmp\xlsx-dumps"
$dashFile = "$dashDir\dashboard-final.html"
$indexFile= "$dashDir\index.html"
$logFile  = "$dashDir\refresh-29sep.log"

"$(Get-Date) START" | Out-File $logFile

function Load-Sheet($d,$f){return Get-Content "$d\$f" -Raw|ConvertFrom-Json}
function Load-Svl($d){if(Test-Path "$d\SNs_Config____SVL.json"){return Get-Content "$d\SNs_Config____SVL.json" -Raw|ConvertFrom-Json};return Get-Content "$d\SNs_Config_SVL.json" -Raw|ConvertFrom-Json}

$p11Dump=Get-ChildItem "$dumpBase\P11 (2025)*" -Directory|Sort-Object LastWriteTime -Descending|Select-Object -First 1
if(-not $p11Dump){$p11Dump=Get-ChildItem "$dumpBase\P11*" -Directory|Where-Object{$_.Name -notmatch '^P11 Balcones'}|Sort-Object LastWriteTime -Descending|Select-Object -First 1}
$fsDump      =Get-ChildItem "$dumpBase\FLASHSYSTEM*" -Directory|Sort-Object LastWriteTime -Descending|Select-Object -First 1
$zmidDump    =Get-ChildItem "$dumpBase\z Mid Range*" -Directory|Sort-Object LastWriteTime -Descending|Select-Object -First 1
$balconesDump=Get-ChildItem "$dumpBase\P11 Balcones*"-Directory|Sort-Object LastWriteTime -Descending|Select-Object -First 1

"P11: $($p11Dump.Name)" | Add-Content $logFile
"FS:  $($fsDump.Name)"  | Add-Content $logFile
"ZM:  $($zmidDump.Name)"| Add-Content $logFile
"BC:  $($balconesDump.Name)"| Add-Content $logFile

$p11Ess=Load-Sheet $p11Dump.FullName "ESS_Installs.json"; $p11Svl=Load-Svl $p11Dump.FullName
$fsEss =Load-Sheet $fsDump.FullName  "ESS_Installs.json"; $fsSvl =Load-Svl $fsDump.FullName
$zmidEss=Load-Sheet $zmidDump.FullName "ESS_Installs.json"; $zmidSvl=Load-Svl $zmidDump.FullName
$balconesEss=Load-Sheet $balconesDump.FullName "ESS_Installs.json"; $balconesSvl=Load-Svl $balconesDump.FullName
"$(Get-Date) Loaded JSON" | Add-Content $logFile

function esc($v){return([string]$v).Trim()-replace'\\','\\'-replace'"','\"'-replace"`r`n",' '-replace"`n",' '}
$script:dateFormats=[string[]]@('yyyy-MM-dd','yyyy-MM-ddTHH:mm:ss','yyyy-MM-ddTHH:mm:ss.fff','yyyy-MM-ddTHH:mm:ss.fffZ','yyyy-MM-ddTHH:mm:ssZ','yyyy-MM-dd HH:mm:ss','dd-MM-yyyy','dd/MM/yyyy','dd-MM-yyyy HH:mm:ss','dd/MM/yyyy HH:mm:ss')
function toDate($v){if($null -eq $v){return $null};$s=([string]$v).Trim();if($s -eq "" -or $s -eq "0001-01-01"){return $null};$inv=[System.Globalization.CultureInfo]::InvariantCulture;if($s -match '^\d{5,7}(\.\d+)?$'){try{return[datetime]::FromOADate([double]::Parse($s,$inv))}catch{return $null}};$st=[System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal;$o=[datetime]::MinValue;if([datetime]::TryParseExact($s,$script:dateFormats,$inv,$st,[ref]$o)){return $o};return $null}
function fmtDate($v){$d=toDate $v;if($null -ne $d){return $d.ToString("yyyy-MM-dd")};if($null -eq $v -or ([string]$v).Trim() -eq "0001-01-01"){return ""};return ([string]$v).Trim()}
function parseDate($v){return toDate $v}
function deriveIbmClient($c){$cu=([string]$c).Trim().ToUpper();if($cu -match '\bIBM\b' -or $cu -match 'INTERNATIONAL BUSINESS MACHINE'){return "IBM"};return "CLIENT"}
function svclvdPriority($v){$v=([string]$v).Trim().ToUpper();if($v -match '^2H COMMITTED FIX'){return 0};if($v -match '^4H COMMITTED FIX'){return 1};if($v -match '^6H COMMITTED FIX'){return 2};if($v -match '^8H COMMITTED FIX'){return 3};if($v -match '^12H COMMITTED FIX'){return 4};if($v -match '^24H COMMITTED FIX'){return 5};if($v -match '^48H COMMITTED FIX'){return 6};if($v -match '^72H COMMITTED FIX'){return 7};if($v -match '^SD SAME DAY'){return 8};if($v -match '^ON-SITE REPAIR.*ORT=SD'){return 9};if($v -match '^NBD NEXT BUSINESS DAY'){return 10};if($v -match '^ON-SITE REPAIR.*ORT=NBD'){return 11};if($v -eq ','){return 12};return 99}
function deriveComment($vals){$status=([string]($vals["STATUS"])).Trim().ToUpper();$ibc=([string]($vals["IBM/CLIENT"])).Trim().ToUpper();if($ibc -eq "IBM"){return "IBM"};function blankV($v){return($null -eq $v -or ([string]$v).Trim() -eq "" -or ([string]$v).Trim() -eq "0001-01-01")};$hasEnt=(-not(blankV $vals["SVCLVC"]))-or(-not(blankV $vals["SVCLVD"]))-or(-not(blankV $vals["SVCSTA"]))-or(-not(blankV $vals["SVCEND"]));if($status -eq "IN" -and(-not(blankV $vals["SVCEND"]))){$se=parseDate $vals["SVCEND"];if($null -ne $se -and $se -lt(Get-Date)){return "Contract Expired"}};if($hasEnt){return "Complete"};if($status -eq "SH"){$sh=parseDate $vals["SHIPD"];if($null -ne $sh -and $sh.Year -lt 2025){return "Invalid"};return "Not Installed"};if($status -eq "IN"){$ins=parseDate $vals["INSTD"];if($null -ne $ins -and $ins.Year -lt 2025){return "Invalid"};return "No Entitlement"};return ""}
function Build-DashData($ess,$svl,$lbl){
    $eH=$ess.headers;$sH=$svl.headers;$eR=$ess.rows;$sR=$svl.rows
    $eI=@{};for($i=0;$i -lt $eH.Count;$i++){$eI[$eH[$i]]=$i}
    $sI=@{};for($i=0;$i -lt $sH.Count;$i++){$sI[$sH[$i]]=$i}
    $svlMap=@{}
    foreach($r in $sR){$s=([string]$r[$sI["SERIAL"]]).Trim();if($s){if(-not $svlMap.ContainsKey($s)){$svlMap[$s]=[System.Collections.Generic.List[object]]::new()};$svlMap[$s].Add($r)}}
    function getBest($rs){if(-not $rs){return $null};$b=$null;$bp=999;foreach($r in $rs){$p=svclvdPriority $r[$sI["SVCLVD"]];if($p -ne 99 -and $p -lt $bp){$bp=$p;$b=$r}};return $b}
    $COLS=@("ILPRID","SERIAL","STATUS","IBMCTY","Country","Region","GEO","MDL","PRDTYP","ISO","CNAME","CITY","ZIP","ADDR","RGN","INSTT","INSTD","SHIPD","WED","SVCLVD","SVCLVC","SVCSTA","SVCEND","CONTRNO","BGNAM","TOM")
    $dateCols=@("INSTD","SHIPD","WED","SVCSTA","SVCEND");$numCols=@("IBMCTY","PRDTYP");$missCols=@("Country","Region","GEO");$svlCols=@("SVCLVD","SVCLVC","SVCSTA","SVCEND","CONTRNO")
    $compact=[System.Collections.Generic.List[string]]::new()
    $ctyIN=@{};$ctySH=@{};$ctyMtIN=@{};$ctyMtSH=@{};$ciCustSvc=@{};$csCustSvc=@{}
    function Inc($h,$k){if($h.ContainsKey($k)){$h[$k]++}else{$h[$k]=1}}
    foreach($row in $eR){
        $serial=([string]$row[$eI["SERIAL"]]).Trim();$bestSvl=if($svlMap.ContainsKey($serial)){getBest $svlMap[$serial]}else{$null}
        $res=@{}
        foreach($col in $COLS){if($missCols -contains $col){$res[$col]=""}elseif($svlCols -contains $col){$res[$col]=if($null -ne $bestSvl -and $sI.ContainsKey($col)){$bestSvl[$sI[$col]]}else{""}}elseif($eI.ContainsKey($col)){$res[$col]=$row[$eI[$col]]}else{$res[$col]=""}}
        $ibc=$res["IBM/CLIENT"]=deriveIbmClient $res["CNAME"]
        $res["INSTD"]=if($eI.ContainsKey("INSTD")){$row[$eI["INSTD"]]}else{""}
        $res["SHIPD"]=if($eI.ContainsKey("SHIPD")){$row[$eI["SHIPD"]]}else{""}
        $res["STATUS"]=if($eI.ContainsKey("STATUS")){$row[$eI["STATUS"]]}else{""}
        $parts=[System.Collections.Generic.List[string]]::new()
        foreach($col in $COLS){if($missCols -contains $col){$parts.Add('""')}elseif($dateCols -contains $col){$parts.Add('"'+(fmtDate $res[$col])+'"')}elseif($numCols -contains $col){$n=0;if([int]::TryParse([string]$res[$col],[ref]$n)){$parts.Add([string]$n)}else{$parts.Add('"'+(esc $res[$col])+'"')}}else{$parts.Add('"'+(esc $res[$col])+'"')}}
        $parts.Add('"'+(esc(deriveComment $res))+'"');$parts.Add('"'+$ibc+'"')
        $compact.Add('['+($parts -join ',')+']')
        $st=([string]$row[$eI["STATUS"]]).Trim();$cty=([string]$row[$eI["IBMCTY"]]).Trim();$mt=([string]$row[$eI["PRDTYP"]]).Trim();$cu=([string]$row[$eI["CNAME"]]).Trim()
        $sv=if($null -ne $bestSvl -and $sI.ContainsKey("SVCLVD")){([string]$bestSvl[$sI["SVCLVD"]]).Trim()}else{""}
        if($st -eq "IN"){Inc $ctyIN $cty;Inc $ctyMtIN "$cty|$mt";Inc $ciCustSvc "$cty|$mt|$cu|$sv"}elseif($st -eq "SH"){Inc $ctySH $cty;Inc $ctyMtSH "$cty|$mt";Inc $csCustSvc "$cty|$mt|$cu|$sv"}
    }
    $v1=($ctyIN.Keys+$ctySH.Keys)|Select-Object -Unique|Sort-Object{[int]$_}|ForEach-Object{$c=$_;$i=if($ctyIN.ContainsKey($c)){$ctyIN[$c]}else{0};$s=if($ctySH.ContainsKey($c)){$ctySH[$c]}else{0};'{"c":'+$c+',"i":'+$i+',"s":'+$s+'}'}
    $v2=($ctyMtIN.Keys+$ctyMtSH.Keys)|Select-Object -Unique|Sort-Object|ForEach-Object{$k=$_;$p=$k -split '\|';$i=if($ctyMtIN.ContainsKey($k)){$ctyMtIN[$k]}else{0};$s=if($ctyMtSH.ContainsKey($k)){$ctyMtSH[$k]}else{0};'{"c":'+$p[0]+',"m":'+$p[1]+',"i":'+$i+',"s":'+$s+'}'}
    $v3=($ciCustSvc.Keys+$csCustSvc.Keys)|Select-Object -Unique|Sort-Object|ForEach-Object{$k=$_;$p=$k -split '\|';$i=if($ciCustSvc.ContainsKey($k)){$ciCustSvc[$k]}else{0};$s=if($csCustSvc.ContainsKey($k)){$csCustSvc[$k]}else{0};'{"c":'+$p[0]+',"m":'+$p[1]+',"cu":"'+(esc $p[2])+'","sv":"'+(esc($p[3..($p.Count-1)] -join '|'))+'","i":'+$i+',"s":'+$s+'}'}
    "$lbl done: V1=$($v1.Count) V2=$($v2.Count) V3=$($v3.Count) Rows=$($compact.Count)" | Add-Content $logFile
    return @{V1='['+($v1 -join ',')+']';V2='['+($v2 -join ',')+']';V3='['+($v3 -join ',')+']';ROWS='['+($compact -join ',')+']'}
}

$zmidData    =Build-DashData $zmidEss     $zmidSvl     "ZMID"
$balconesData=Build-DashData $balconesEss $balconesSvl "BALCONES"
$fsData      =Build-DashData $fsEss       $fsSvl       "FS7600"
$p11Data     =Build-DashData $p11Ess      $p11Svl      "P11"
"$(Get-Date) Build complete" | Add-Content $logFile

$varMap=@{
    "ALL_ROWS"="var ALL_ROWS="+$p11Data.ROWS+";"; "V1"="var V1="+$p11Data.V1+";"; "V2"="var V2="+$p11Data.V2+";"; "V3"="var V3="+$p11Data.V3+";"
    "ALL_ROWS_FS7600"="var ALL_ROWS_FS7600="+$fsData.ROWS+";"; "V1_FS7600"="var V1_FS7600="+$fsData.V1+";"; "V2_FS7600"="var V2_FS7600="+$fsData.V2+";"; "V3_FS7600"="var V3_FS7600="+$fsData.V3+";"
    "ALL_ROWS_ZMID"="var ALL_ROWS_ZMID="+$zmidData.ROWS+";"; "V1_ZMID"="var V1_ZMID="+$zmidData.V1+";"; "V2_ZMID"="var V2_ZMID="+$zmidData.V2+";"; "V3_ZMID"="var V3_ZMID="+$zmidData.V3+";"
    "ALL_ROWS_BALCONES"="var ALL_ROWS_BALCONES="+$balconesData.ROWS+";"; "V1_BALCONES"="var V1_BALCONES="+$balconesData.V1+";"; "V2_BALCONES"="var V2_BALCONES="+$balconesData.V2+";"; "V3_BALCONES"="var V3_BALCONES="+$balconesData.V3+";"
}
$lines=[System.IO.File]::ReadAllLines($dashFile,[System.Text.Encoding]::UTF8)
$replaced=@{}
for($i=0;$i -lt $lines.Length;$i++){foreach($key in $varMap.Keys){if($lines[$i] -match ("^var "+$key+"=")){$lines[$i]=$varMap[$key];$replaced[$key]=$true}}}
[System.IO.File]::WriteAllLines($dashFile,$lines,(New-Object System.Text.UTF8Encoding($false)))
Copy-Item $dashFile $indexFile -Force
"$(Get-Date) Injected OK - replaced $($replaced.Count)/16" | Add-Content $logFile

Set-Location $dashDir
$today=Get-Date -Format "yyyy-MM-dd HH:mm"
git add index.html dashboard-final.html
git commit -m "Dashboard refresh 29th Sep — P11(5328) FS7600(2308) ZMID(47) BALCONES(77) — $today"
git push origin main
"$(Get-Date) PUSHED" | Add-Content $logFile