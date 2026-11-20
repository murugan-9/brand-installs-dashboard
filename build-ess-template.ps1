# build-ess-template.ps1
# One-time (re-runnable) patch: turns the 4-brand dashboard-final.html into a data-driven
# template (dashboard-ess.template.html) that can show any number of programs.
# Reads dashboard-final.html, NEVER modifies it. Every patch asserts that its anchor was found.
param(
    [string]$Source   = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data\dashboard-final.html",
    [string]$Template = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data\dashboard-ess.template.html"
)
$ErrorActionPreference = "Stop"
$L = [System.Collections.Generic.List[string]]::new()
$L.AddRange([string[]][System.IO.File]::ReadAllLines($Source, [System.Text.Encoding]::UTF8))

function Find-Line([string]$pattern, [int]$from = 0) {
    for ($i = $from; $i -lt $L.Count; $i++) { if ($L[$i] -match $pattern) { return $i } }
    throw "Anchor not found: $pattern"
}

# 1. Title + header text -> tokens filled in by the refresh script
$i = Find-Line '^<title>Serial Status Dashboard'
$L[$i] = '<title>Serial Status Dashboard &#x2013; {{TITLE_SUFFIX}}</title>'
$i = Find-Line '^\s*<h1>Serial Status Dashboard'
$L[$i] = '    <h1>Serial Status Dashboard &mdash; {{TITLE_SUFFIX}}</h1>'

# 2. Brand dropdown: scrollable (18+ entries), per-brand pill CSS injected after the last legacy pill rule
$i = Find-Line '^\.hdr-brand-drop\{'
if ($L[$i] -notmatch 'min-width:200px\}$') { throw "hdr-brand-drop rule has unexpected shape" }
$L[$i] = $L[$i] -replace 'min-width:200px\}$', 'min-width:200px;max-height:70vh;overflow-y:auto}'
$i = Find-Line '^\.brand-pill\.bp-balcones\{'
$L.Insert($i + 1, '/*{{BRAND_CSS}}*/')

# 3. Static brand radios br1..br4 -> generated
$first = Find-Line 'id="br1" class="brk"'
$last  = Find-Line 'id="br4" class="brk"'
if ($last -ne $first + 3) { throw "Expected 4 consecutive brand radios" }
$L.RemoveRange($first, 4)
$L.Insert($first, '        {{BRAND_OPTIONS}}')

# 4. Per-brand data vars (ALL_ROWS*, V1*/V2*/V3*) -> two registry vars the refresh script fills
$removed = 0
for ($i = $L.Count - 1; $i -ge 0; $i--) {
    if ($L[$i] -match '^var (ALL_ROWS|V1|V2|V3)(_FS7600|_ZMID|_BALCONES)?=') { $L.RemoveAt($i); $removed++ }
}
if ($removed -ne 16) { throw "Expected to remove 16 legacy data vars, removed $removed" }
$i = Find-Line '^var COLS=\['
$L.Insert($i, 'var BRANDS=[];')
$L.Insert($i + 1, 'var BRAND_ROWS={};')

# 5. MASTER_ROWS built from the registry
$s = Find-Line '^var ALL_ROWS_P11 = ALL_ROWS;'
$e = Find-Line '^\}\(\)\);' $s
$L.RemoveRange($s, $e - $s + 1)
$L.InsertRange($s, [string[]]@(
'var MASTER_ROWS = (function(){',
'  var m = [];',
'  BRANDS.forEach(function(b){ var a=BRAND_ROWS[b.key]||[]; for(var i=0;i<a.length;i++) m.push(a[i].concat([b.key])); });',
'  return m;',
'}());'))

# 6. Per-brand first-seen-ISO index + brand-aware lookup
$s = Find-Line '^var _isoP11=\{\}'
$fn = Find-Line '^function getCtyMetaByBrand' $s
$e = Find-Line '^\}$' $fn
$L.RemoveRange($s, $e - $s + 1)
$L.InsertRange($s, [string[]]@(
'var _isoBy={};',
'BRANDS.forEach(function(b){ var m=_isoBy[b.key]={}; (BRAND_ROWS[b.key]||[]).forEach(function(r){ if(r[3]&&!m[r[3]]) m[r[3]]=String(r[9]||'''').trim(); }); });',
'// For summary views: pick ISO by active brand, then P11, then any brand, then resolve via getCtyMeta',
'function getCtyMetaByBrand(ibmcty){',
'  var brand=getActiveBrand();',
'  var iso='''';',
'  if(brand!==''ALL''&&_isoBy[brand]&&_isoBy[brand][ibmcty]) iso=_isoBy[brand][ibmcty];',
'  if(!iso&&_isoBy[''P11'']&&_isoBy[''P11''][ibmcty]) iso=_isoBy[''P11''][ibmcty];',
'  for(var i=0;i<BRANDS.length&&!iso;i++){ var mm=_isoBy[BRANDS[i].key]; if(mm&&mm[ibmcty]) iso=mm[ibmcty]; }',
'  return getCtyMeta(ibmcty, iso);',
'}'))

# 7. filteredRows source selection
$s = Find-Line '^\s*var src = brand===''P11'''
$e = Find-Line 'MASTER_ROWS;\s*$' $s
$L.RemoveRange($s, $e - $s + 1)
$L.Insert($s, '  var src = brand===''ALL'' ? MASTER_ROWS : (BRAND_ROWS[brand]||[]);')

# 8. Download file name from the registry (legacy names preserved; All Brands no longer mislabelled P11_2025)
$i = Find-Line '^\s*_modalTitle=\(_ab===''FS7600'''
$L[$i] = '  _modalTitle=(function(){var b=BRANDS.filter(function(x){return x.key===_ab;})[0];return b?b.file:''ALL_Programs'';}())+''_''+lab+''_Records'';'

# 9. Brand pill text/colour from the registry
$s = Find-Line '^\s*if\(bv===''P11''\)'
$e = Find-Line 'bpill\.textContent=''All Brands''' $s
$L.RemoveRange($s, $e - $s + 1)
$L.InsertRange($s, [string[]]@(
'          var _bb=BRANDS.filter(function(x){return x.key===bv;})[0];',
'          if(_bb){bpill.textContent=_bb.label;bpill.className=''brand-pill ''+_bb.cls;}',
'          else   {bpill.textContent=''All Brands'';bpill.className=''brand-pill bp-all'';}'))

# 10. Missing-country check runs over every program
$i = Find-Line '^\s*var allRows = \(ALL_ROWS\|\|\[\]\)'
$L[$i] = '  var allRows = MASTER_ROWS;'

# Final sanity: no legacy brand-array references may remain
$left = @(); for ($i = 0; $i -lt $L.Count; $i++) { if ($L[$i] -match 'ALL_ROWS|_isoP11|_isoFS|_isoZMID|_isoBALCONES|ALL_ROWS_P11') { $left += ("{0}: {1}" -f ($i+1), $L[$i].Substring(0, [Math]::Min(100, $L[$i].Length))) } }
if ($left.Count) { throw ("Legacy references still present:`n" + ($left -join "`n")) }

[System.IO.File]::WriteAllLines($Template, $L, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("Template written: {0}  ({1} lines, {2:N0} bytes)" -f $Template, $L.Count, (Get-Item $Template).Length) -ForegroundColor Green
