#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Master SharePoint Sensitivity Labels Extraction Orchestrator

.DESCRIPTION
    This script provides a guided workflow to:
    1. Fetch sensitivity labels catalog for referen                Write-Host "  - $($label.DisplayName) [$($label.Guid)]" -ForegroundColor Graye
    2. Genera                Write-Host "  - $($label.displayName) [$($label.id)]" -ForegroundColor Graye inventory of all SharePoint sites with sizes
    3. Allow user to choose extraction scope (all sites vs specific sites)
    4. Generate individual CSV per site + consoli    Write-Host "PERFORMANCE OPTIMIZED MODE:" -ForegroundColor Greenated report

.PARAMETER ConfigPath
    Path to the configuration file (default: .\config.json)

.EXAMPLE
    .\Master-SharePoint-Extractor.ps1
    
.NOTES
    Author: SharePoint Sensitivity Labels Team
    Version: 1.0
    Requires: Microsoft.Graph, ExchangeOnlineManagement modules
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = ".\config.json"
)

# ============================================================================
# MODULE AVAILABILITY CHECK (Fast Startup)
# ============================================================================

Write-Host "Verifying required PowerShell modules are available..." -ForegroundColor Cyan
Write-Host "Modules will be imported only when needed for better performance" -ForegroundColor Gray

# Required modules for the script
$requiredModules = @(
    @{Name = "Microsoft.Graph"; MinVersion = "1.0.0"},
    @{Name = "ExchangeOnlineManagement"; MinVersion = "2.0.0"}
)

$missingModules = @()
foreach ($module in $requiredModules) {
    Write-Host "Checking availability: $($module.Name)..." -ForegroundColor Gray
    
    # Quick check if module is installed (no import)
    $installedModule = Get-Module -ListAvailable -Name $module.Name | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $installedModule) {
        Write-Host "✗ Module $($module.Name) not found" -ForegroundColor Red
        $missingModules += $module.Name
    } else {
        Write-Host "✓ Module $($module.Name) available (Version: $($installedModule.Version))" -ForegroundColor Green
    }
}

if ($missingModules.Count -gt 0) {
    Write-Host "`nMissing modules detected. Installing now..." -ForegroundColor Yellow
    foreach ($moduleName in $missingModules) {
        try {
            Write-Host "Installing $moduleName..." -ForegroundColor Yellow
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "✓ Successfully installed $moduleName" -ForegroundColor Green
        } catch {
            Write-Error "Failed to install $moduleName`: $($_.Exception.Message)"
            Write-Host "Please manually run: Install-Module -Name $moduleName -Scope CurrentUser -Force" -ForegroundColor Red
            exit 1
        }
    }
}

Write-Host "`n✓ All required modules are available!" -ForegroundColor Green
Write-Host "Modules will be imported on-demand for optimal performance" -ForegroundColor Gray
Write-Host ""

# Import configuration helpers
if (Test-Path ".\ConfigHelpers.ps1") {
    . .\ConfigHelpers.ps1
} else {
    Write-Host "ConfigHelpers.ps1 not found, using basic config loading" -ForegroundColor Yellow
    function Get-ConfigFromFile {
        param([string]$ConfigPath)
        return Get-Content $ConfigPath -Raw | ConvertFrom-Json
    }
}

function Show-Banner {
    Write-Host @"

===============================================================
          SharePoint Sensitivity Labels Master Extractor      
                                                               
  A comprehensive solution for enterprise-scale sensitivity    
  label extraction across SharePoint Online sites             
===============================================================

"@ -ForegroundColor Cyan
}

function Show-InitialMenu {
    Write-Host "`n SHAREPOINT INVENTORY OPTIONS:" -ForegroundColor Yellow
    Write-Host "1. Use existing SharePoint inventory list" -ForegroundColor Cyan
    Write-Host "2. Extract new SharePoint inventory (scan all sites)" -ForegroundColor Green
    Write-Host "3. Exit" -ForegroundColor Red
    Write-Host ""
}

function Show-MainMenu {
    Write-Host "`n EXTRACTION WORKFLOW OPTIONS:" -ForegroundColor Yellow
    Write-Host "1. Extract from ALL SharePoint sites (automated)" -ForegroundColor Green
    Write-Host "2. Extract from SPECIFIC sites (user-selected)" -ForegroundColor Green
    Write-Host "3. Exit" -ForegroundColor Red
    Write-Host ""
}

function Use-ExistingInventory {
    param([string]$RunFolder)
    
    Write-Host "`nUSING EXISTING SHAREPOINT INVENTORY:" -ForegroundColor Yellow
    Write-Host "
Please follow these steps:" -ForegroundColor Cyan
    Write-Host "1. Copy your existing SharePoint-Sites-Inventory_*.csv file to: $RunFolder" -ForegroundColor White
    Write-Host "2. Make sure the 'Extract' column is set to 'True' for sites you want to process" -ForegroundColor White
    Write-Host "3. Make sure the 'Extract' column is set to 'False' for sites you want to skip" -ForegroundColor White
    Write-Host "4. Press Enter when you've placed the file in the run folder" -ForegroundColor White
    Write-Host ""
    
    Read-Host "Press Enter when you've copied your inventory file"
    
    # Look for inventory file in run folder
    try {
        $inventoryFiles = Get-ChildItem -Path $RunFolder -Filter "SharePoint-Sites-Inventory_*.csv" -ErrorAction Stop
    } catch {
        Write-Error "Could not access run folder: $RunFolder"
        return $null
    }
    
    if ($inventoryFiles.Count -eq 0) {
        Write-Error "No SharePoint-Sites-Inventory_*.csv file found in $RunFolder"
        Write-Host "Please copy your inventory file and try again." -ForegroundColor Red
        return $null
    }
    
    if ($inventoryFiles.Count -gt 1) {
        Write-Warning "Multiple inventory files found. Using the first one: $($inventoryFiles[0].Name)"
    }
    
    # Get the full path of the first inventory file
    $inventoryFile = $inventoryFiles[0].FullName.Trim()
    
    if (-not (Test-Path $inventoryFile)) {
        Write-Error "Inventory file does not exist: $inventoryFile"
        return $null
    }
    
    Write-Host "Found inventory file: $inventoryFile" -ForegroundColor Green
    
    # Validate the file has required columns
    try {
        $testImport = Import-Csv -Path $inventoryFile | Select-Object -First 1
        $requiredColumns = @('SiteName', 'SiteUrl', 'Extract')
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $testImport.PSObject.Properties.Name }
        
        if ($missingColumns.Count -gt 0) {
            Write-Error "Missing required columns: $($missingColumns -join ', ')"
            return $null
        }
        
        Write-Host "Inventory file validated successfully!" -ForegroundColor Green
        return $inventoryFile
    } catch {
        Write-Error "Error reading inventory file: $($_.Exception.Message)"
        return $null
    }
}

function Get-AllSharePointSites {
    param([object]$Config)
    
    Write-Host "`nSTEP 1: Generating SharePoint Sites Inventory (Enterprise Mode)..." -ForegroundColor Yellow
    
    # Import Microsoft Graph modules only when needed
    Write-Host "Loading Microsoft Graph modules..." -ForegroundColor Cyan
    try {
        Import-Module Microsoft.Graph.Authentication -Force -WarningAction SilentlyContinue
        Import-Module Microsoft.Graph.Sites -Force -WarningAction SilentlyContinue
        Write-Host " Microsoft Graph modules loaded" -ForegroundColor Green
    } catch {
        Write-Warning "Could not load specific Graph modules, trying main module"
        Import-Module Microsoft.Graph -Force
    }
    
    # Connect to Graph API
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    try {
        if ($Config.authentication.preferredMethod -eq "appOnly") {
            $ClientSecretSecure = ConvertTo-SecureString $Config.authentication.appOnly.clientSecret -AsPlainText -Force
            $ClientSecretCredential = New-Object System.Management.Automation.PSCredential($Config.authentication.appOnly.clientId, $ClientSecretSecure)
            Connect-MgGraph -TenantId $Config.authentication.appOnly.tenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
        } else {
            Connect-MgGraph -Scopes $Config.authentication.interactive.scopes -NoWelcome
        }
        Write-Host " Connected to Microsoft Graph" -ForegroundColor Green
    } catch {
        Write-Error " Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        return $null
    }
    
    Write-Host " Using high-performance bulk site enumeration..." -ForegroundColor Cyan
    $allSites = @()
    $pageSize = 999  # Maximum page size for performance
    
    # Use SharePoint Admin API for fastest site collection
    try {
        Write-Host "Fetching all SharePoint sites (bulk operation)..." -ForegroundColor Yellow
        
        # Method 1: Use sites endpoint with expanded properties for maximum efficiency
        $sitesUri = "https://graph.microsoft.com/v1.0/sites?`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime,siteCollection&`$top=$pageSize"
        
        $totalSites = 0
        $batchCount = 0
        
        do {
            $batchCount++
            Write-Progress -Activity "Fast Site Collection" -Status "Processing batch $batchCount" -PercentComplete (($totalSites % 100))
            
            $sitesResponse = Invoke-MgGraphRequest -Uri $sitesUri -Method GET
            $totalSites += $sitesResponse.value.Count
            
            foreach ($site in $sitesResponse.value) {
                # Get storage and file estimates efficiently
                $storageUsedMB = 0
                $estimatedFiles = 0
                $docLibraryCount = 0
                
                try {
                    # Use site analytics for faster data gathering
                    if ($site.siteCollection) {
                        $storageUsedMB = [math]::Round($site.siteCollection.storageUsedInBytes / 1MB, 2)
                    }
                    
                    # Quick drive count (don't enumerate files yet - performance optimization)
                    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,driveType&`$top=50"
                    $drivesResponse = Invoke-MgGraphRequest -Uri $drivesUri -Method GET -ErrorAction SilentlyContinue
                    if ($drivesResponse.value) {
                        $docLibraryCount = ($drivesResponse.value | Where-Object { $_.driveType -eq 'documentLibrary' }).Count
                        # Rough file estimate based on drive count (avoid expensive file enumeration)
                        $estimatedFiles = $docLibraryCount * 50  # Conservative estimate
                    }
                    
                } catch {
                    # Skip individual site errors to maintain performance
                    Write-Debug "Could not get details for site: $($site.displayName)"
                }
                
                $siteInfo = [PSCustomObject]@{
                    SiteName = $site.displayName
                    SiteUrl = $site.webUrl
                    SiteId = $site.id
                    CreatedDateTime = $site.createdDateTime
                    LastModifiedDateTime = $site.lastModifiedDateTime
                    StorageUsedMB = $storageUsedMB
                    EstimatedFiles = $estimatedFiles
                    DocumentLibraries = $docLibraryCount
                    Extract = $true  # Default to extract
                    Status = "Pending"
                    LastScanned = ""
                }
                
                $allSites += $siteInfo
            }
            
            $sitesUri = $sitesResponse.'@odata.nextLink'
        } while ($sitesUri)
        
        Write-Progress -Activity "Fast Site Collection" -Completed
        Write-Host " High-speed enumeration complete: $($allSites.Count) SharePoint sites" -ForegroundColor Green
        
    } catch {
        Write-Error " Error in bulk site enumeration: $($_.Exception.Message)"
        Write-Warning "Falling back to slower individual site processing..."
        
        # Fallback to slower method if bulk fails
        $sitesUri = "https://graph.microsoft.com/v1.0/sites?`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime&`$top=50"
        $sitesResponse = Invoke-MgGraphRequest -Uri $sitesUri -Method GET
        
        foreach ($site in $sitesResponse.value) {
            $allSites += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                SiteId = $site.id
                CreatedDateTime = $site.createdDateTime
                LastModifiedDateTime = $site.lastModifiedDateTime
                StorageUsedMB = 0
                EstimatedFiles = 0
                DocumentLibraries = 0
                Extract = $true
                Status = "Pending"
                LastScanned = ""
            }
        }
    }
    
    return $allSites
}

function Export-SitesInventory {
    param(
        [array]$Sites,
        [string]$RunFolder,
        [string]$Timestamp
    )
    
    # Generate inventory file path in run folder
    $inventoryFile = "$RunFolder\SharePoint-Sites-Inventory_$Timestamp.csv"
    
    $Sites | Export-Csv -Path $inventoryFile -NoTypeInformation -Encoding UTF8
    
    Write-Host " Sites inventory exported to: $inventoryFile" -ForegroundColor Green
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "   Total Sites: $($Sites.Count)"
    Write-Host "   Total Storage: $([math]::Round(($Sites.StorageUsedMB | Measure-Object -Sum).Sum / 1024, 2)) GB"
    Write-Host "   Estimated Files: $([math]::Round(($Sites.EstimatedFiles | Measure-Object -Sum).Sum, 0))"
    
    return $inventoryFile
}

function Get-SensitivityLabelsCache {
    param([object]$Config, [string]$RunFolder)
    
    Write-Host "`n  STEP 0: Fetching Sensitivity Labels Catalog..." -ForegroundColor Yellow
    
    $labelMap = @{}
    
    # Try Purview first if not skipped
    if (-not $Config.labelResolution.skipPurview) {
        try {
            Write-Host "Loading Exchange Online Management module..." -ForegroundColor Cyan
            Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
            Write-Host "Connecting to Purview for label catalog..." -ForegroundColor Cyan
            Connect-IPPSSession -WarningAction SilentlyContinue
            
            $purviewLabels = Get-Label -WarningAction SilentlyContinue
            foreach ($label in $purviewLabels) {
                $labelMap[$label.Guid] = $label.DisplayName
                Write-Host "   $($label.DisplayName) [$($label.Guid)]" -ForegroundColor Gray
            }
            Write-Host " Loaded $($labelMap.Count) labels from Purview" -ForegroundColor Green
        } catch {
            Write-Warning "Could not connect to Purview: $($_.Exception.Message)"
        }
    }
    
    # Fallback to Graph API label catalog
    if ($labelMap.Count -eq 0 -and $Config.labelResolution.useGraphLabelCatalog) {
        try {
            Write-Host "Fetching labels from Graph API..." -ForegroundColor Cyan
            $graphLabelsUri = "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels"
            $graphLabelsResponse = Invoke-MgGraphRequest -Uri $graphLabelsUri -Method GET
            
            foreach ($label in $graphLabelsResponse.value) {
                $labelMap[$label.id] = $label.displayName
                Write-Host "   $($label.displayName) [$($label.id)]" -ForegroundColor Gray
            }
            Write-Host " Loaded $($labelMap.Count) labels from Graph API" -ForegroundColor Green
        } catch {
            Write-Warning "Could not fetch labels from Graph API: $($_.Exception.Message)"
        }
    }
    
    # Cache labels to file
    if ($labelMap.Count -gt 0) {
        $labelsFile = "$RunFolder\Sensitivity-Labels-Cache.json"
        # Convert hashtable to PSCustomObject for JSON serialization
        $labelsForJson = [PSCustomObject]@{}
        $labelMap.GetEnumerator() | ForEach-Object { 
            $labelsForJson | Add-Member -MemberType NoteProperty -Name $_.Key -Value $_.Value 
        }
        $labelsForJson | ConvertTo-Json -Depth 3 | Out-File -FilePath $labelsFile -Encoding UTF8
        Write-Host " Labels cached to: $labelsFile" -ForegroundColor Green
    }
    
    return $labelMap
}

function Start-ExtractionProcess {
    param(
        [string]$SitesInventoryFile,
        [hashtable]$LabelMap,
        [object]$Config,
        [string]$ExtractionType,
        [string]$RunFolder,
        [string]$Timestamp
    )
    
    # Trim any leading/trailing spaces from the file path
    $SitesInventoryFile = $SitesInventoryFile.Trim()
    
    Write-Host "`n STEP 2: Starting Sensitivity Labels Extraction..." -ForegroundColor Yellow
    Write-Host "Extraction Mode: $ExtractionType" -ForegroundColor Cyan
    
    # Load sites to process with validation
    $allSitesFromCsv = Import-Csv -Path $SitesInventoryFile | Where-Object { $_.Extract -eq "True" -or $_.Extract -eq $true }
    
    # Filter out sites with empty or null names
    $sitesToProcess = $allSitesFromCsv | Where-Object { 
        $_.SiteName -and 
        $_.SiteName.Trim() -ne "" -and 
        $_.SiteName -ne $null 
    }
    
    # Report any invalid sites that were filtered out
    $invalidSites = $allSitesFromCsv | Where-Object { 
        -not $_.SiteName -or 
        $_.SiteName.Trim() -eq "" -or 
        $_.SiteName -eq $null 
    }
    
    if ($invalidSites.Count -gt 0) {
        Write-Warning "Filtered out $($invalidSites.Count) sites with empty/invalid names"
        foreach ($invalidSite in $invalidSites) {
            Write-Warning "   - Row with SiteUrl: '$($invalidSite.SiteUrl)' (empty SiteName)"
        }
    }
    
    if ($sitesToProcess.Count -eq 0) {
        Write-Warning "No valid sites marked for extraction. Please check the inventory CSV."
        return
    }
    
    Write-Host " Processing $($sitesToProcess.Count) sites..." -ForegroundColor Green
    
    $processedCount = 0
    $skippedCount = 0
    $siteCsvFiles = @()
    foreach ($site in $sitesToProcess) {
        # Additional validation before processing
        if (-not $site.SiteName -or $site.SiteName.Trim() -eq "") {
            $skippedCount++
            Write-Warning "Skipping site with empty name: '$($site.SiteUrl)'"
            continue
        }
        
        $processedCount++
        $percentComplete = [math]::Round(($processedCount / $sitesToProcess.Count) * 100, 1)
        Write-Host "`n[$processedCount/$($sitesToProcess.Count)] Processing: $($site.SiteName) ($percentComplete%)" -ForegroundColor Yellow
        try {
            # Use the comprehensive scanner for each site (silent mode for performance)
            $siteResults = & .\Comprehensive-Scanner.ps1 -SiteName $site.SiteName -ReturnData -RunFolder $RunFolder -SkipModuleImport -SilentMode
            if ($siteResults -and $siteResults.Count -gt 0) {
                # Export individual site CSV
                $siteFileName = ($site.SiteName -replace '[^\w\s-]', '') -replace '\s+', '_'
                $siteOutputFile = "$RunFolder\Site_$($siteFileName)_$Timestamp.csv"
                $siteResults | Export-Csv -Path $siteOutputFile -NoTypeInformation -Encoding UTF8
                $siteCsvFiles += $siteOutputFile
                Write-Host " Site exported: $siteOutputFile ($($siteResults.Count) records)" -ForegroundColor Green
                $site.Status = "Completed"
                $site.LastScanned = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            } else {
                Write-Warning "No data extracted from site: $($site.SiteName)"
                $site.Status = "No Data"
                $site.LastScanned = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        } catch {
            Write-Error " Error processing site $($site.SiteName): $($_.Exception.Message)"
            $site.Status = "Error: $($_.Exception.Message)"
            $site.LastScanned = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    # Export consolidated results by merging all per-site CSVs
    $allResults = @()
    foreach ($csvFile in $siteCsvFiles) {
        if (Test-Path $csvFile) {
            $allResults += Import-Csv -Path $csvFile
        }
    }
    if ($allResults.Count -gt 0) {
        $consolidatedFile = "$RunFolder\Consolidated_SharePoint_SensitivityLabels_$Timestamp.csv"
        $allResults | Export-Csv -Path $consolidatedFile -NoTypeInformation -Encoding UTF8
        Write-Host " CONSOLIDATED REPORT: $consolidatedFile" -ForegroundColor Green
        Write-Host "Total Records: $($allResults.Count)" -ForegroundColor Cyan
    }
    
    # Update inventory with processing status
    $sitesToProcess | Export-Csv -Path $SitesInventoryFile -NoTypeInformation -Encoding UTF8
    
    # Generate summary report
    Write-Host "`n EXTRACTION SUMMARY:" -ForegroundColor Yellow
    Write-Host "Sites Processed: $processedCount"
    if ($skippedCount -gt 0) {
        Write-Host "Sites Skipped (empty names): $skippedCount" -ForegroundColor Yellow
    }
    Write-Host "Total Records: $($allResults.Count)"
    Write-Host "Files with Labels: $(($allResults | Where-Object { $_.LabelIds -and $_.LabelIds -ne '' }).Count)"
    Write-Host "Unique Sites: $(($allResults.SiteName | Sort-Object -Unique).Count)"
    Write-Host "Unique File Types: $(($allResults.FileExtension | Sort-Object -Unique | Where-Object { $_ }).Count)"
}

function Wait-ForUserReview {
    param([string]$InventoryFile)
    
    Write-Host "`n  PLEASE REVIEW AND EDIT THE SITES INVENTORY:" -ForegroundColor Yellow
    Write-Host " File: $InventoryFile" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor Green
    Write-Host "1. Open the CSV file in Excel or any text editor"
    Write-Host "2. Review the 'Extract' column (currently all set to 'True')"
    Write-Host "3. Change 'Extract' to 'False' for sites you DON'T want to process"
    Write-Host "4. Save the file and close it"
    Write-Host "5. Come back here and press Enter to continue"
    Write-Host ""
    
    # Open the file for user review
    try {
        Start-Process $InventoryFile
    } catch {
        Write-Host "Could not auto-open file. Please manually open: $InventoryFile" -ForegroundColor Yellow
    }
    
    Read-Host "Press Enter when you've finished editing the inventory file"
}

# Main execution
try {
    Show-Banner
    
    # Create unique run folder
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $runFolder = ".\Output\Run_$timestamp"
    New-Item -ItemType Directory -Path $runFolder -Force | Out-Null
    Write-Host " Created run folder: $runFolder" -ForegroundColor Green
    
    # Load configuration
    $config = Get-ConfigFromFile -ConfigPath $ConfigPath
    if (-not $config) {
        Write-Error " Could not load configuration from $ConfigPath"
        return
    }
    
    Write-Host " PERFORMANCE OPTIMIZED MODE:" -ForegroundColor Green
    Write-Host "  - Bulk site enumeration: ENABLED" -ForegroundColor Gray
    Write-Host "  - Silent file processing: ENABLED" -ForegroundColor Gray
    Write-Host "  - High-throughput Graph calls: ENABLED" -ForegroundColor Gray
    Write-Host "  - Enterprise scalability: READY" -ForegroundColor Gray
    Write-Host ""
    
    # Step 0: Get sensitivity labels cache
    $labelMap = Get-SensitivityLabelsCache -Config $config -RunFolder $runFolder
    
    # Initial Menu: Choose inventory method
    $inventoryFile = $null
    do {
        Show-InitialMenu
        $inventoryChoice = Read-Host "Please select an option (1-3)"
        
        switch ($inventoryChoice) {
            "1" {
                Write-Host "`n OPTION 1 SELECTED: Use existing SharePoint inventory" -ForegroundColor Cyan
                $inventoryFile = Use-ExistingInventory -RunFolder $runFolder
                if ($inventoryFile -and (Test-Path $inventoryFile)) {
                    Write-Host "Successfully loaded existing inventory: $inventoryFile" -ForegroundColor Green
                    break
                } else {
                    Write-Host "Failed to use existing inventory. Please try again." -ForegroundColor Red
                    $inventoryFile = $null
                }
            }
            "2" {
                Write-Host "`n OPTION 2 SELECTED: Extract new SharePoint inventory" -ForegroundColor Green
                
                # Step 1: Generate SharePoint sites inventory
                $allSites = Get-AllSharePointSites -Config $config
                if (-not $allSites -or $allSites.Count -eq 0) {
                    Write-Error " Could not retrieve SharePoint sites"
                    return
                }
                
                $inventoryFile = Export-SitesInventory -Sites $allSites -RunFolder $runFolder -Timestamp $timestamp
                break
            }
            "3" {
                Write-Host " Goodbye!" -ForegroundColor Yellow
                return
            }
            default {
                Write-Host " Invalid selection. Please choose 1, 2, or 3." -ForegroundColor Red
            }
        }
    } while ($inventoryChoice -notin @("1", "2", "3") -or -not $inventoryFile)
    
    # Final validation of inventory file
    if (-not $inventoryFile -or -not (Test-Path $inventoryFile)) {
        Write-Error "No valid inventory file available. Cannot proceed with extraction."
        return
    }
    
    # Step 2: Handle extraction based on initial choice
    if ($inventoryChoice -eq "1") {
        # For existing inventory, go directly to extraction (user already selected sites in CSV)
        Write-Host "`n PROCESSING EXISTING INVENTORY: Extracting from pre-selected sites" -ForegroundColor Green
        Start-ExtractionProcess -SitesInventoryFile $inventoryFile -LabelMap $labelMap -Config $config -ExtractionType "SELECTED_SITES" -RunFolder $runFolder -Timestamp $timestamp
    } else {
        # For new inventory, show extraction menu and get user choice
        do {
            Show-MainMenu
            $choice = Read-Host "Please select an option (1-3)"
            
            switch ($choice) {
                "1" {
                    Write-Host "`n OPTION 1 SELECTED: Extract from ALL sites" -ForegroundColor Green
                    Start-ExtractionProcess -SitesInventoryFile $inventoryFile -LabelMap $labelMap -Config $config -ExtractionType "ALL_SITES" -RunFolder $runFolder -Timestamp $timestamp
                    break
                }
                "2" {
                    Write-Host "`n OPTION 2 SELECTED: Extract from SPECIFIC sites" -ForegroundColor Green
                    Wait-ForUserReview -InventoryFile $inventoryFile
                    Start-ExtractionProcess -SitesInventoryFile $inventoryFile -LabelMap $labelMap -Config $config -ExtractionType "SELECTED_SITES" -RunFolder $runFolder -Timestamp $timestamp
                    break
                }
                "3" {
                    Write-Host " Goodbye!" -ForegroundColor Yellow
                    return
                }
                default {
                    Write-Host " Invalid selection. Please choose 1, 2, or 3." -ForegroundColor Red
                }
            }
        } while ($choice -notin @("1", "2", "3"))
    }
    
    Write-Host "`n EXTRACTION WORKFLOW COMPLETED!" -ForegroundColor Green
    Write-Host "Check the $runFolder directory for all generated files." -ForegroundColor Cyan
    
} catch {
    Write-Error " Fatal error in main execution: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
}
