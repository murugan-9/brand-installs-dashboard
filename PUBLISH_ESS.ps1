# =============================================================================
#  PUBLISH_ESS.ps1  -  publish the reviewed dashboard-ess.html to GitHub Pages
#  Refuses to run without -Approved, and refuses unless check-ess.ps1 passed for the
#  CURRENT dashboard-ess.html (so a stale or failed build can never be published).
#  Copies dashboard-ess.html -> dashboard-final.html + index.html, commits, merges to main, pushes.
#  Revert: git revert <that commit>   (or tag pre-ess-18-programs = old 4-brand dashboard)
# =============================================================================
param([switch]$Approved)

$ErrorActionPreference = 'Stop'
$dir = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
Set-Location $dir

if (-not $Approved) { Write-Host "Not approved. Review dashboard-ess.html, then run:  .\PUBLISH_ESS.ps1 -Approved" -ForegroundColor Yellow; exit 2 }

function Invoke-Git {
    $ErrorActionPreference = 'Continue'      # git writes progress to stderr; judge by exit code only
    $out = & git.exe @args 2>&1
    $out | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') failed (exit $LASTEXITCODE)" }
}

# gate: the checks must have passed for this exact dashboard file
$rep = Get-Content "$dir\ess-check-report.json" -Raw | ConvertFrom-Json
$tick = (Get-Item "$dir\dashboard-ess.html").LastWriteTimeUtc.Ticks
if (-not $rep.pass -or [int64]$rep.dashboardTicks -ne $tick) { throw "check-ess.ps1 has not passed for the current dashboard-ess.html - run REFRESH_ESS.ps1 (or check-ess.ps1) first." }

$branch = (& git.exe branch --show-current).Trim()
Write-Host "Publishing from branch '$branch'..." -ForegroundColor Cyan
Copy-Item "$dir\dashboard-ess.html" "$dir\dashboard-final.html" -Force
Copy-Item "$dir\dashboard-ess.html" "$dir\index.html" -Force
Invoke-Git add dashboard-ess.html dashboard-final.html index.html ess-build-summary.json ess-query-summary.json UPDATE_DASHBOARD_ESS.ps1 UPDATE_DASHBOARD.ps1 REFRESH_ESS.ps1 PUBLISH_ESS.ps1 check-ess.ps1 build-ess-template.ps1 dashboard-ess.template.html dump-ess.ps1
$pending = (& git.exe status --porcelain --untracked-files=no) | Where-Object { $_ }
if ($pending) { Invoke-Git commit -m ("ESS dashboard refresh " + (Get-Date -Format 'yyyy-MM-dd HH:mm')) } else { Write-Host "  (nothing new to commit)" }
if ($branch -ne 'main') { Invoke-Git checkout main; Invoke-Git merge --no-edit $branch }
Invoke-Git push origin main

$a = (& git.exe rev-parse main).Trim(); $b = (& git.exe rev-parse origin/main).Trim()
if ($a -ne $b) { throw "main != origin/main after push" }
Write-Host "`nPublished. main == origin/main ($($a.Substring(0,7)))" -ForegroundColor Green
Write-Host "Live in 1-2 min: https://murugan-9.github.io/brand-installs-dashboard/  (Ctrl+Shift+R)" -ForegroundColor Cyan
