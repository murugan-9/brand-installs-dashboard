# ============================================================
#  Dashboard Publisher — Brand Installs Dashboard
#  Double-click this file to publish the latest dashboard
#  to GitHub Pages whenever you refresh with new data.
# ============================================================

$dashboardFolder = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data"
$sourceFile      = Join-Path $dashboardFolder "dashboard-final.html"
$targetFile      = Join-Path $dashboardFolder "index.html"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Brand Installs Dashboard Publisher" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Step 1 — Check source file exists
if (-Not (Test-Path $sourceFile)) {
    Write-Host "ERROR: dashboard-final.html not found at:" -ForegroundColor Red
    Write-Host "  $sourceFile" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please make sure you have refreshed the dashboard first." -ForegroundColor Yellow
    pause
    exit 1
}

# Step 2 — Copy dashboard-final.html to index.html
Write-Host "Step 1/4  Copying dashboard-final.html to index.html..." -ForegroundColor White
Copy-Item $sourceFile $targetFile -Force
Write-Host "          Done." -ForegroundColor Green

# Step 3 — Stage the file
Write-Host "Step 2/4  Staging changes..." -ForegroundColor White
Set-Location $dashboardFolder
git add index.html
Write-Host "          Done." -ForegroundColor Green

# Step 4 — Commit with today's date
$today = Get-Date -Format "yyyy-MM-dd HH:mm"
Write-Host "Step 3/4  Committing with message: 'Dashboard refresh $today'..." -ForegroundColor White
git commit -m "Dashboard refresh $today"
Write-Host "          Done." -ForegroundColor Green

# Step 5 — Push to GitHub
Write-Host "Step 4/4  Pushing to GitHub..." -ForegroundColor White
git push origin main
Write-Host "          Done." -ForegroundColor Green

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " SUCCESS! Your dashboard is now live at:" -ForegroundColor Green
Write-Host " https://murugan-9.github.io/brand-installs-dashboard/" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Users will see the latest data in 1-2 minutes." -ForegroundColor White
Write-Host ""
pause
