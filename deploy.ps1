# deploy.ps1
# Usage: ./deploy.ps1
# This script pushes your local code to multiple Google Apps Script projects.

# --- CONFIGURATION ---
# Add your target Script IDs here.
$targetScriptIds = @(
    "1td7dv1m_4vpsTreo7ASpMNx8SYtAul6NtkiqTQSiJuZqDIE-WqNNJl35",
    "1W6kF_6Y7_6v68uZqDIE-WqNNJl35"
)

# ----------------------
$claspFile = ".clasp.json"

if (-not (Test-Path $claspFile)) {
    Write-Error ".clasp.json not found! Please run this script in the root of your project."
    exit
}

# Backup original .clasp.json
$originalJson = Get-Content $claspFile -Raw | ConvertFrom-Json
$originalId = $originalJson.scriptId

Write-Host "--- Multi-Project Deployment Started ---" -ForegroundColor Cyan

foreach ($id in $targetScriptIds) {
    if ($id -match "REPLACE_WITH") {
        Write-Host "Skipping placeholder ID: $id" -ForegroundColor Yellow
        continue
    }

    Write-Host "`nTargeting Script ID: $id" -ForegroundColor Green
    
    # Update scriptId in .clasp.json
    $originalJson.scriptId = $id
    $originalJson | ConvertTo-Json | Set-Content $claspFile

    # Push code
    Write-Host "Pushing code via clasp..."
    clasp push
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to push to $id" -ForegroundColor Red
    } else {
        Write-Host "Successfully pushed to $id" -ForegroundColor Green
    }
}

# Restore original .clasp.json
Write-Host "`nRestoring original settings..." -ForegroundColor Cyan
$originalJson.scriptId = $originalId
$originalJson | ConvertTo-Json | Set-Content $claspFile

Write-Host "Done!" -ForegroundColor Green
