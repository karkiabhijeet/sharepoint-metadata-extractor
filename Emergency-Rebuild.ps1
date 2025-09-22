#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Emergency script rebuilder for corrupted GitHub downloads

.DESCRIPTION
    This script completely rebuilds the Master-SharePoint-Extractor.ps1 with clean syntax
    when the automatic fix doesn't work due to severe corruption.

.EXAMPLE
    .\Emergency-Rebuild.ps1

.NOTES
    Use this as a last resort if Fix-DownloadedScript.ps1 doesn't resolve the issues
#>

Write-Host "Emergency Script Rebuild Utility" -ForegroundColor Red
Write-Host "==================================" -ForegroundColor Gray

Write-Host "This will create a clean Master-SharePoint-Extractor.ps1 file" -ForegroundColor Yellow
$confirm = Read-Host "Continue? (Y/N)"

if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Cancelled." -ForegroundColor Gray
    exit
}

Write-Host "Creating clean script file..." -ForegroundColor Cyan

$cleanScript = @'
#Requires -Version 5.1
#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Master SharePoint Sensitivity Labels Extraction Orchestrator - Clean Version

.DESCRIPTION
    This is a rebuilt version with clean syntax to avoid GitHub download corruption issues.
    
.NOTES
    If you're seeing this, the Emergency-Rebuild.ps1 was used to create a clean script.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = ".\config.json"
)

Write-Host "SharePoint Metadata Extractor (Clean Build)" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Gray

# Check modules availability (fast)
Write-Host "Verifying PowerShell modules..." -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph", "ExchangeOnlineManagement")
$missingModules = @()

foreach ($moduleName in $requiredModules) {
    $module = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
    if ($module) {
        Write-Host "✓ $moduleName available (v$($module.Version))" -ForegroundColor Green
    } else {
        Write-Host "✗ $moduleName missing" -ForegroundColor Red
        $missingModules += $moduleName
    }
}

if ($missingModules.Count -gt 0) {
    Write-Host "`nInstalling missing modules..." -ForegroundColor Yellow
    foreach ($moduleName in $missingModules) {
        Install-Module -Name $moduleName -Scope CurrentUser -Force
        Write-Host "✓ Installed $moduleName" -ForegroundColor Green
    }
}

Write-Host "`n✓ All modules ready!" -ForegroundColor Green
Write-Host "`nTo run the full extractor, please download a fresh copy from GitHub" -ForegroundColor Yellow
Write-Host "or use the Fix-DownloadedScript.ps1 utility." -ForegroundColor Yellow

Write-Host "`nThis emergency version only validates module availability." -ForegroundColor Gray
Write-Host "For full functionality, ensure you have a clean download of the complete script." -ForegroundColor Gray
'@

# Write the clean script
$cleanScript | Out-File -FilePath ".\Master-SharePoint-Extractor-Clean.ps1" -Encoding UTF8 -Force

Write-Host "✓ Created: Master-SharePoint-Extractor-Clean.ps1" -ForegroundColor Green
Write-Host "`nThis emergency version only checks modules." -ForegroundColor Yellow
Write-Host "For the full extractor, try:" -ForegroundColor Cyan
Write-Host "1. Download fresh from GitHub" -ForegroundColor White  
Write-Host "2. Run: .\Fix-DownloadedScript.ps1" -ForegroundColor White
Write-Host "3. Or clone the repository instead of downloading ZIP" -ForegroundColor White