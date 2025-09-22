#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Fixes common PowerShell script download issues from GitHub

.DESCRIPTION
    This script fixes encoding and syntax issues that commonly occur when downloading
    PowerShell scripts from GitHub, especially the Master-SharePoint-Extractor.ps1

.PARAMETER ScriptPath
    Path to the downloaded script file (default: .\Master-SharePoint-Extractor.ps1)

.EXAMPLE
    .\Fix-DownloadedScript.ps1
    .\Fix-DownloadedScript.ps1 -ScriptPath "C:\Downloads\Master-SharePoint-Extractor.ps1"

.NOTES
    Run this script if you get parse errors after downloading from GitHub
#>

[CmdletBinding()]
param(
    [string]$ScriptPath = ".\Master-SharePoint-Extractor.ps1"
)

Write-Host "Fixing GitHub Download Issues for PowerShell Scripts" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Gray

# Check if script exists
if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script not found: $ScriptPath"
    Write-Host "Please make sure you've downloaded the script from GitHub" -ForegroundColor Red
    exit 1
}

Write-Host "Found script: $ScriptPath" -ForegroundColor Green

# Create backup
$backupPath = "$ScriptPath.backup"
Copy-Item $ScriptPath $backupPath -Force
Write-Host "Created backup: $backupPath" -ForegroundColor Yellow

try {
    # Read the file with proper encoding handling
    Write-Host "Reading and fixing encoding issues..." -ForegroundColor Cyan
    
    $content = Get-Content -Path $ScriptPath -Raw -Encoding UTF8
    
    # Fix common syntax issues from GitHub downloads
    Write-Host "Fixing common syntax issues..." -ForegroundColor Cyan
    
    # Fix backtick issues in URLs (common GitHub download corruption)
    $content = $content -replace '`\$select', '`$select'
    $content = $content -replace '`\$top', '`$top'
    $content = $content -replace '&`\$', '&`$'
    
    # Fix string interpolation issues
    $content = $content -replace '\$\(([^)]+)\)([^"]*)"', '$($1)$2"'
    
    # Fix percentage symbol issues
    $content = $content -replace '\$percentComplete%', '$($percentComplete)%'
    
    # Fix record count display
    $content = $content -replace '\$\(([^)]+)\.Count\)\s+records', '$($1.Count) records'
    
    # Remove any problematic Unicode characters
    $content = $content -replace '[^\x00-\x7F]', ''
    
    # Ensure proper line endings
    $content = $content -replace '\r\n', "`n" -replace '\r', "`n" -replace '\n', "`r`n"
    
    # Write back with UTF8 encoding
    $content | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    
    Write-Host "Fixed encoding and syntax issues" -ForegroundColor Green
    
    # Test syntax
    Write-Host "Testing PowerShell syntax..." -ForegroundColor Cyan
    
    $parseErrors = $null
    $tokens = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    
    if ($parseErrors.Count -gt 0) {
        Write-Warning "Found $($parseErrors.Count) syntax errors:"
        foreach ($parseError in $parseErrors) {
            Write-Host "  Line $($parseError.Extent.StartLineNumber): $($parseError.Message)" -ForegroundColor Red
        }
        Write-Host "`nScript may still have issues. Check the errors above." -ForegroundColor Yellow
    } else {
        Write-Host "✓ Script syntax is valid!" -ForegroundColor Green
        Write-Host "✓ Script is ready to run" -ForegroundColor Green
        
        # Remove backup since fix was successful
        Remove-Item $backupPath -Force
        Write-Host "Removed backup file (fix successful)" -ForegroundColor Gray
    }
    
    Write-Host "`nYou can now run: .\Master-SharePoint-Extractor.ps1" -ForegroundColor Cyan
    
} catch {
    Write-Error "Failed to fix script: $($_.Exception.Message)"
    Write-Host "Restoring from backup..." -ForegroundColor Yellow
    Copy-Item $backupPath $ScriptPath -Force
    Write-Host "Original file restored. Please try downloading again or report the issue." -ForegroundColor Red
}