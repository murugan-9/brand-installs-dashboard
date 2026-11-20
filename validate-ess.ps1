$ErrorActionPreference='Stop'
$dir = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$old = [IO.File]::ReadAllLines("$dir\dashboard-final.html")
$new = [IO.File]::ReadAllText("$dir\dashboard-ess.html")
$map = @{ P11='ALL_ROWS'; FS7600='ALL_ROWS_FS7600'; ZMID='ALL_ROWS_ZMID'; BALCONES='ALL_ROWS_BALCONES' }
foreach ($k in $map.Keys) {
  $prefix = 'var ' + $map[$k] + '='
  $line = $old | ? { $_.StartsWith($prefix) } | select -First 1
  $oldRows = $line.Substring($prefix.Length).TrimEnd(';')
  $needle = '"' + $k + '":' + $oldRows
  $same = $new.Contains($needle)
  "{0,-9} legacy rows byte-identical in new dashboard: {1}  (len {2})" -f $k,$same,$oldRows.Length
}
# date cells: any non yyyy-MM-dd / empty value in date columns (idx 16,17,18,21,22)
$m = [regex]::Match($new, 'var BRAND_ROWS=\{(.*?)\};\r?\n')
"BRAND_ROWS length: $($m.Groups[1].Length)"
$json = '{' + $m.Groups[1].Value + '}'
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = [int]::MaxValue; $ser.RecursionLimit = 100
$obj = $ser.DeserializeObject($json)
$bad = @{}; $cmt = @{}; $ctys = @{}; $tot=0
foreach ($k in $obj.Keys) {
  $rows = $obj[$k]; $tot += $rows.Count
  foreach ($r in $rows) {
    if ($r.Count -ne 28) { "BAD COLCOUNT in $k : $($r.Count)" }
    foreach ($i in 16,17,18,21,22) { $v=[string]$r[$i]; if ($v -ne '' -and $v -notmatch '^\d{4}-\d{2}-\d{2}$') { $bad["$k|$i|$v"]=1 } elseif ($v -match '^0001') { $bad["$k|$i|$v"]=1 } }
    $ck="$k|"+$r[26]; if ($cmt.ContainsKey($ck)) {$cmt[$ck]++} else {$cmt[$ck]=1}
    $ctys[[string]$r[3]] = 1
  }
}
"Total rows: $tot"
"Bad date cells: $($bad.Count)"; $bad.Keys | select -First 10
"--- comments ---"
$cmt.GetEnumerator() | sort Name | % { "{0,-40} {1}" -f $_.Name,$_.Value }
# missing CTY
$cf = ($old | ? { $_ -match 'CTY_FALLBACK\s*=' } | select -First 1)
$mm = [regex]::Match($new, 'var CTY_FALLBACK=\{(.*?)\n\};', 'Singleline'); $line = $mm.Groups[1].Value
$known = @{}; [regex]::Matches($line,'(\d+)\s*:') | % { $known[$_.Groups[1].Value]=1 }
"CTY_FALLBACK entries parsed: $($known.Count)"
$missing = $ctys.Keys | ? { $_ -and -not $known.ContainsKey($_) } | sort
"Missing IBMCTY codes: $($missing -join ',')"
