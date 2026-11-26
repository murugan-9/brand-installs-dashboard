# =============================================================================
#  check-hub.ps1  -  automated validation of the built NPI HUB (.\hub\)
#  Exit 0 = all checks passed, 1 = at least one FAIL. Writes hub-check-report.json.
#  Never modifies the hub or the source dashboards, never touches git.
#
#  1. every tab file is byte-identical to its source dashboard (SHA-256), and the sources have not changed since the build
#  2. the hub is served over a local web server and opened in headless Edge; each tab is clicked and must load,
#     show its own header, have real content, a visible frame and ZERO JavaScript errors
#  3. clicking the dashboards' own inner tabs raises no errors
#  4. deep links (index.html#tab-id) select the right tab
# =============================================================================
$ErrorActionPreference = 'Stop'
$dir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$hub      = Join-Path $dir 'hub'
$manifest = Join-Path $hub 'hub-manifest.json'
$report   = Join-Path $dir 'hub-check-report.json'
$edge     = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

$fails = [System.Collections.Generic.List[string]]::new()
function Bad($m)  { $script:fails.Add($m); Write-Host "  FAIL: $m" -ForegroundColor Red }
function Good($m) { Write-Host "  OK:   $m" -ForegroundColor Green }

Write-Host "`n=== NPI HUB CHECKS ===" -ForegroundColor Cyan
if (-not (Test-Path $manifest)) { throw "hub\hub-manifest.json not found - run Build-Hub.ps1 first" }
$mf = @(Get-Content $manifest -Raw | ConvertFrom-Json | ForEach-Object { $_ })

# 1. byte identity
Write-Host "1. Tab files are byte-identical to their source dashboards" -ForegroundColor White
foreach ($t in $mf) {
    $copy = Join-Path $hub $t.file
    if (-not (Test-Path -LiteralPath $copy))   { Bad "$($t.label): missing $($t.file)"; continue }
    if (-not (Test-Path -LiteralPath $t.source)) { Bad "$($t.label): source not found $($t.source)"; continue }
    $hc = (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash
    $hs = (Get-FileHash -LiteralPath $t.source -Algorithm SHA256).Hash
    if ($hc -ne $t.sha256) { Bad "$($t.label): hub copy changed after build" }
    elseif ($hs -ne $t.sha256) { Bad "$($t.label): source dashboard changed since the hub was built - rebuild the hub" }
    else { Good "$($t.label): identical to source ($($hc.Substring(0,12))..., $([math]::Round($t.bytes/1MB,2)) MB)" }
}

# 2-4. headless browser test over http
Write-Host "2. Headless Edge test (served over http, every tab clicked)" -ForegroundColor White
if (-not (Test-Path $edge)) { Write-Host "  SKIP: Edge not found" -ForegroundColor Yellow }
else {
    $tmp = Join-Path $dir 'hub-test.tmp'
    if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
    Copy-Item $hub $tmp -Recurse
    $errHook = '<script>window.__errs=[];window.addEventListener("error",function(e){window.__errs.push(String(e.message))});window.addEventListener("unhandledrejection",function(e){window.__errs.push("promise: "+String(e.reason))});</script>'
    # instrument TEST COPIES only (error collector as the first script); the real hub files are untouched
    foreach ($t in $mf) {
        $p = Join-Path $tmp $t.file; $h = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
        $i = $h.IndexOf('<head>'); if ($i -lt 0) { throw "no <head> in $($t.file)" }
        [IO.File]::WriteAllText($p, $h.Insert($i + 6, $errHook), (New-Object Text.UTF8Encoding($false)))
    }
    $test = @'
<script>
(function(){
var out=[],perr=[];
window.addEventListener('error',function(e){perr.push(String(e.message))});
function log(s){out.push(s)}
function wait(cond,ms,cb){var t0=Date.now();(function p(){var ok=false;try{ok=cond()}catch(e){}if(ok||Date.now()-t0>ms)cb(ok);else setTimeout(p,100)})()}
window.addEventListener('load',function(){
 var btns=[].slice.call(document.querySelectorAll('.hub-tab')),i=0;
 log('HUBTITLE='+(document.querySelector('.hub-title')||{}).textContent+'  TABS='+btns.length);
 function next(){
  if(i>=btns.length){return deep();}
  var b=btns[i++];b.click();
  var f=document.querySelector('.hub-frame.active');
  wait(function(){var d=f.contentDocument;return d&&d.readyState==='complete'&&d.body&&d.body.innerText.length>200},40000,function(ok){
   var d=f.contentDocument,w=f.contentWindow,h1=d&&d.querySelector('h1');
   log('TAB['+b.textContent+'] loaded='+ok+' h1='+(h1?h1.textContent.trim():'')+' textLen='+(d?d.body.innerText.length:0)+' frameH='+f.offsetHeight+' frameW='+f.offsetWidth+' errs='+(w.__errs?w.__errs.length:-1));
   var inner=[].slice.call(d.querySelectorAll('#tabs button,#tabs [role=tab],#tabs .tab,.tab,.tabs button,nav button')).filter(function(x,ix,a){return a.indexOf(x)===ix}).slice(0,10),n=0,j=0;
   (function step(){
    if(j>=inner.length){log('TAB['+b.textContent+'] inner-clicks='+n+' errs-after='+(w.__errs?w.__errs.length:-1));if(w.__errs&&w.__errs.length)log('ERRTEXT '+w.__errs.slice(0,3).join(' || '));return setTimeout(next,200);}
    try{inner[j++].click();n++}catch(e){log('CLICKERR '+e.message)}
    setTimeout(step,150);
   })();
  });
 }
 function deep(){
  var res=[];
  var ids=[].slice.call(document.querySelectorAll('.hub-frame')).map(function(x){return x.id.replace('frame-','')});
  var k=0;(function s(){
   if(k>=ids.length){log('DEEPLINK '+res.join(','));return fin();}
   var id=ids[k++];location.hash='#'+id;
   setTimeout(function(){var a=document.querySelector('.hub-frame.active');res.push(id+'='+(a&&a.id==='frame-'+id));s();},300);
  })();
 }
 function fin(){
  log('HUBERRS='+perr.length);
  var pre=document.createElement('pre');pre.id='hubtest';pre.textContent=out.join('\n');document.body.appendChild(pre);
 }
 next();
});
})();
</script></body>
'@
    $ip = Join-Path $tmp 'index.html'; $ih = [IO.File]::ReadAllText($ip, [Text.Encoding]::UTF8)
    $k = $ih.LastIndexOf('</body>')
    [IO.File]::WriteAllText($ip, $ih.Substring(0, $k) + $test + $ih.Substring($k + 7), (New-Object Text.UTF8Encoding($false)))

    $port = Get-Random -Minimum 18000 -Maximum 18900
    $job = Start-Job -ArgumentList $tmp, $port -ScriptBlock {
        param($root, $port)
        $l = New-Object Net.HttpListener; $l.Prefixes.Add("http://localhost:$port/"); $l.Start()
        while ($l.IsListening) {
            $c = $l.GetContext()
            try {
                $p = [Uri]::UnescapeDataString($c.Request.Url.AbsolutePath.TrimStart('/')); if (-not $p) { $p = 'index.html' }
                $f = [IO.Path]::GetFullPath((Join-Path $root $p))
                if ($f.StartsWith($root) -and (Test-Path -LiteralPath $f -PathType Leaf)) {
                    $b = [IO.File]::ReadAllBytes($f)
                    $c.Response.ContentType = $(if ($f -like '*.json') { 'application/json' } else { 'text/html; charset=utf-8' })
                    $c.Response.ContentLength64 = $b.Length; $c.Response.OutputStream.Write($b, 0, $b.Length)
                } else { $c.Response.StatusCode = 404 }
            } catch {} finally { $c.Response.Close() }
        }
    }
    try {
        $up = $false; for ($i = 0; $i -lt 50 -and -not $up; $i++) { Start-Sleep -Milliseconds 200; try { $up = (Invoke-WebRequest "http://localhost:$port/hub-manifest.json" -UseBasicParsing -TimeoutSec 3).StatusCode -eq 200 } catch {} }
        if (-not $up) { throw "local test server did not start" }
        $ErrorActionPreference = 'Continue'
        $dom = & $edge --headless --disable-gpu --window-size=1400,900 --virtual-time-budget=120000 --dump-dom "http://localhost:$port/index.html" 2>$null | Out-String
        $ErrorActionPreference = 'Stop'
    } finally { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue; Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }

    $m = [regex]::Match($dom, '<pre id="hubtest">(.*?)</pre>', 'Singleline')
    if (-not $m.Success) { Bad "browser test produced no output (page did not finish loading)" }
    else {
        $txt = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
        $txt -split "`n" | ForEach-Object { Write-Host "      $_" -ForegroundColor DarkGray }
        if ($txt -notmatch 'HUBTITLE=NPI HUB') { Bad "hub header is not 'NPI HUB'" } else { Good "hub header is 'NPI HUB'" }
        if ($txt -notmatch "TABS=$($mf.Count)\b") { Bad "expected $($mf.Count) tabs" } else { Good "$($mf.Count) tabs: $((($mf | ForEach-Object { $_.label }) -join ' | '))" }
        foreach ($t in $mf) {
            $l = [regex]::Escape($t.label)
            $r = [regex]::Match($txt, "TAB\[$l\] loaded=(\w+) h1=(.*?) textLen=(\d+) frameH=(\d+) frameW=(\d+) errs=(-?\d+)")
            if (-not $r.Success) { Bad "$($t.label): no result"; continue }
            if ($r.Groups[1].Value -ne 'true') { Bad "$($t.label): did not load" }
            elseif ([int]$r.Groups[3].Value -lt 500) { Bad "$($t.label): almost no content ($($r.Groups[3].Value) chars)" }
            elseif ([int]$r.Groups[4].Value -lt 400 -or [int]$r.Groups[5].Value -lt 800) { Bad "$($t.label): frame too small ($($r.Groups[5].Value)x$($r.Groups[4].Value))" }
            elseif ($r.Groups[6].Value -ne '0') { Bad "$($t.label): $($r.Groups[6].Value) JavaScript error(s) on load" }
            else { Good "$($t.label): loaded, header '$($r.Groups[2].Value)', $($r.Groups[3].Value) chars, 0 JS errors" }
            $ic = [regex]::Match($txt, "TAB\[$l\] inner-clicks=(\d+) errs-after=(-?\d+)")
            if (-not $ic.Success) { Bad "$($t.label): inner tab clicks did not finish" }
            elseif ([int]$ic.Groups[1].Value -lt 1) { Bad "$($t.label): found no inner tabs to click" }
            elseif ($ic.Groups[2].Value -ne '0') { Bad "$($t.label): JavaScript errors after clicking its own tabs ($([regex]::Match($txt,'ERRTEXT (.*)').Groups[1].Value))" }
            else { Good "$($t.label): $($ic.Groups[1].Value) inner tabs clicked, 0 JS errors" }
        }
        if ($txt -match 'CLICKERR') { Bad "a click raised an exception: $([regex]::Match($txt,'CLICKERR (.*)').Groups[1].Value)" }
        $dl = [regex]::Match($txt, 'DEEPLINK (.*)').Groups[1].Value
        if (-not $dl -or $dl -match '=false') { Bad "deep links failed: $dl" } else { Good "deep links select the right tab ($dl)" }
        if ($txt -notmatch 'HUBERRS=0') { Bad "JavaScript errors in the hub shell" } else { Good "0 JavaScript errors in the hub shell" }
    }
}

$pass = ($fails.Count -eq 0)
[pscustomobject]@{
    pass = $pass; checkedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    manifestTicks = (Get-Item $manifest).LastWriteTimeUtc.Ticks
    tabs = @($mf | ForEach-Object { [pscustomobject]@{ id = $_.id; sha256 = $_.sha256 } })
    failures = @($fails)
} | ConvertTo-Json -Depth 4 | Set-Content $report -Encoding UTF8
Write-Host ""
if ($pass) { Write-Host "ALL HUB CHECKS PASSED" -ForegroundColor Green; exit 0 }
else       { Write-Host "$($fails.Count) HUB CHECK(S) FAILED - do not publish" -ForegroundColor Red; exit 1 }

