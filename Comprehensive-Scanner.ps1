# Comprehensive File and Label Scanner
# Scans ALL files in ALL subfolders, regardless of file type or label status

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteName,
    [switch]$ReturnData,
    [string]$RunFolder = ".\Output",
    [switch]$SkipModuleImport,
    [switch]$SilentMode
)

if (-not $SilentMode) {
    Write-Host "=== Comprehensive File and Label Scanner: $SiteName ===" -ForegroundColor Cyan
}

try {
    # Import required modules only if not skipped (when called standalone)
    if (-not $SkipModuleImport) {
        if (-not (Get-Module Microsoft.Graph)) {
            Import-Module Microsoft.Graph -ErrorAction Stop
        }
        if (-not (Get-Module ExchangeOnlineManagement)) {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
        }
    }
    
    # Read configuration
    $config = Get-Content "config.json" -Raw | ConvertFrom-Json
    
    # Connect to Graph only if not already connected
    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $graphContext -or $graphContext.TenantId -ne $config.authentication.appOnly.tenantId) {
        if (-not $SilentMode) { Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow }
        $clientSecretSecure = ConvertTo-SecureString $config.authentication.appOnly.clientSecret -AsPlainText -Force
        $clientSecretCredential = New-Object System.Management.Automation.PSCredential($config.authentication.appOnly.clientId, $clientSecretSecure)
        Connect-MgGraph -TenantId $config.authentication.appOnly.tenantId -ClientSecretCredential $clientSecretCredential -NoWelcome
        if (-not $SilentMode) { Write-Host "Connected to Graph successfully!" -ForegroundColor Green }
    } else {
        if (-not $SilentMode) { Write-Host "Using existing Graph connection" -ForegroundColor Green }
    }
    
    # Build label map (only for known labels)
    if (-not $SilentMode) { Write-Host "Building sensitivity label catalog..." -ForegroundColor Yellow }
    $labelMap = @{}
    
    # First try to use cached labels from master script
    $cachedLabelsFile = "$RunFolder\Sensitivity-Labels-Cache.json"
    if (Test-Path $cachedLabelsFile) {
        try {
            if (-not $SilentMode) { Write-Host "Using cached labels from master script..." -ForegroundColor Green }
            $cachedLabels = Get-Content $cachedLabelsFile -Raw | ConvertFrom-Json
            $cachedLabels.PSObject.Properties | ForEach-Object {
                $labelMap[$_.Name] = $_.Value
            }
            if (-not $SilentMode) { Write-Host "Retrieved $($labelMap.Count) labels from cache" -ForegroundColor Green }
        } catch {
            Write-Warning "Could not load cached labels, falling back to direct fetch"
        }
    }
    
    # If no cached labels, fetch directly
    if ($labelMap.Count -eq 0) {
        # Check if we should skip Purview and use Graph instead
        if (-not $config.labelResolution.skipPurview) {
            try {
                Connect-IPPSSession | Out-Null
                $purviewLabels = Get-Label | Select-Object DisplayName, ImmutableId
                foreach ($label in $purviewLabels) {
                    if ($label.ImmutableId -and $label.DisplayName) {
                        $labelMap.Add([string]$label.ImmutableId, [string]$label.DisplayName)
                    }
                }
                Write-Host "Retrieved $($labelMap.Count) labels from Purview" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to connect to Purview: $($_.Exception.Message)"
            }
        }
    }
    
        # Try Graph API for labels if Purview is skipped or failed
        if ($labelMap.Count -eq 0 -or $config.labelResolution.useGraphLabelCatalog) {
            try {
                if (-not $SilentMode) { Write-Host "Getting labels from Graph API..." -ForegroundColor Yellow }
                $graphLabelsUri = "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels"
                $graphLabelsResponse = Invoke-MgGraphRequest -Uri $graphLabelsUri -Method GET
                
                foreach ($label in $graphLabelsResponse.value) {
                    if (-not $labelMap.ContainsKey([string]$label.id)) {
                        $labelMap.Add([string]$label.id, [string]$label.displayName)
                    }
                }
                if (-not $SilentMode) { Write-Host "Retrieved $($labelMap.Count) labels from Graph API" -ForegroundColor Green }
            } catch {
                if (-not $SilentMode) { Write-Warning "Failed to get labels from Graph API: $($_.Exception.Message)" }
            }
        }
        
        if (-not $SilentMode) { Write-Host "Total known sensitivity labels: $($labelMap.Count)" -ForegroundColor Green }    # Search for the target site
    if (-not $SilentMode) { Write-Host "Searching for site: $SiteName..." -ForegroundColor Yellow }
    $searchUri = "https://graph.microsoft.com/v1.0/sites?search=" + [uri]::EscapeDataString($SiteName)
    $sitesResponse = Invoke-MgGraphRequest -Uri $searchUri -Method GET
    
    if ($sitesResponse.value.Count -eq 0) {
        if (-not $SilentMode) { Write-Warning "No sites found matching '$SiteName'" }
        return
    }
    
    $targetSite = $sitesResponse.value[0]
    if (-not $SilentMode) { Write-Host "Found site: $($targetSite.displayName)" -ForegroundColor Green }
    
    # Get ALL document libraries (not just Documents)
    if (-not $SilentMode) { Write-Host "Getting ALL document libraries..." -ForegroundColor Yellow }
    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($targetSite.id)/drives"
    $drivesResponse = Invoke-MgGraphRequest -Uri $drivesUri -Method GET
    $docLibraries = $drivesResponse.value | Where-Object { $_.driveType -eq 'documentLibrary' }
    
    # Filter out system libraries
    $systemLibraries = @('Form Templates', 'Preservation Hold Library', 'Site Assets', 'Site Pages', 'Images', 'Pages', 'Settings', 'Style Library')
    $docLibraries = $docLibraries | Where-Object { $_.name -notin $systemLibraries }
    
    if (-not $SilentMode) {
        Write-Host "Found $($docLibraries.Count) document libraries to scan" -ForegroundColor Green
        foreach ($lib in $docLibraries) {
            Write-Host "  - $($lib.name)" -ForegroundColor Cyan
        }
    }
    
    $allResults = @()
    
    # Recursive function to get all files in folders
    function Get-AllFilesRecursive {
        param(
            [string]$DriveId,
            [string]$ParentId = "root",
            [string]$ParentPath = "/"
        )
        
        $files = @()
        $queue = New-Object System.Collections.Queue
        $queue.Enqueue(@{Id = $ParentId; Path = $ParentPath})
        
        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            
            try {
                $itemsUri = if ($current.Id -eq "root") {
                    "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children?`$select=id,name,file,folder,size,createdDateTime,lastModifiedDateTime,createdBy,webUrl,parentReference"
                } else {
                    "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($current.Id)/children?`$select=id,name,file,folder,size,createdDateTime,lastModifiedDateTime,createdBy,webUrl,parentReference"
                }
                
                $itemsResponse = Invoke-MgGraphRequest -Uri $itemsUri -Method GET
                
                foreach ($item in $itemsResponse.value) {
                    if ($item.file) {
                        # It's a file - add to results
                        $fileExtension = [System.IO.Path]::GetExtension($item.name).TrimStart('.')
                        $files += [PSCustomObject]@{
                            Id = $item.id
                            Name = $item.name
                            Extension = $fileExtension
                            Size = $item.size
                            FolderPath = $current.Path
                            WebUrl = $item.webUrl
                            CreatedBy = if ($item.createdBy.user.displayName) { $item.createdBy.user.displayName } else { "Unknown" }
                            CreatedDateTime = $item.createdDateTime
                            LastModifiedDateTime = $item.lastModifiedDateTime
                        }
                    } elseif ($item.folder) {
                        # It's a folder - add to queue for recursive processing
                        $subPath = if ($current.Path -eq "/") { "/$($item.name)" } else { "$($current.Path)/$($item.name)" }
                        $queue.Enqueue(@{Id = $item.id; Path = $subPath})
                    }
                }
            } catch {
                Write-Warning "Could not access folder $($current.Path): $($_.Exception.Message)"
            }
        }
        
        return $files
    }
    
    # Process each document library
    foreach ($drive in $docLibraries) {
        if (-not $SilentMode) { Write-Host "`n  Processing library: $($drive.name)" -ForegroundColor Yellow }
        
        # Get all files recursively from this library
        $libraryFiles = Get-AllFilesRecursive -DriveId $drive.id
        if (-not $SilentMode) { Write-Host "    Found $($libraryFiles.Count) total files" -ForegroundColor Cyan }
        
        $fileCount = 0
        foreach ($file in $libraryFiles) {
            $fileCount++
            
            # Show progress only in non-silent mode, and only occasionally for performance
            if (-not $SilentMode -and ($fileCount % 10 -eq 0 -or $fileCount -eq $libraryFiles.Count)) {
                Write-Progress -Activity "Processing $($drive.name)" -Status "File $fileCount of $($libraryFiles.Count)" -PercentComplete (($fileCount / $libraryFiles.Count) * 100)
            }
            
            # Initialize label-related variables
            $labelIds = ""
            $labelNames = ""
            $assignmentMethods = ""
            $labelExtractionStatus = "Not Attempted"
            
            # Try to extract sensitivity labels for supported file types
            $supportedExtensions = @('docx', 'xlsx', 'pptx', 'pdf', 'doc', 'xls', 'ppt')
            
            if ($file.Extension -in $supportedExtensions) {
                try {
                    $extractUri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($file.Id)/extractSensitivityLabels"
                    $labelResult = Invoke-MgGraphRequest -Uri $extractUri -Method POST
                    
                    if ($labelResult.labels -and $labelResult.labels.Count -gt 0) {
                        $labelIdsList = @()
                        $labelNamesList = @()
                        $assignmentMethodsList = @()
                        
                        foreach ($label in $labelResult.labels) {
                            $labelId = [string]$label.sensitivityLabelId
                            $labelIdsList += $labelId
                            $assignmentMethodsList += [string]$label.assignmentMethod
                            
                            # Try to resolve label name - if not found, leave empty
                            if ($labelMap.ContainsKey($labelId)) {
                                $labelNamesList += $labelMap[$labelId]
                            } else {
                                $labelNamesList += ""  # Empty string for unknown labels
                            }
                        }
                        
                        $labelIds = $labelIdsList -join ';'
                        $labelNames = $labelNamesList -join ';'
                        $assignmentMethods = $assignmentMethodsList -join ';'
                        $labelExtractionStatus = "Success"
                    } else {
                        $labelExtractionStatus = "No Labels Found"
                    }
                    
                } catch {
                    # Label extraction failed - this is normal for non-Office files
                    $labelExtractionStatus = "Failed: $($_.Exception.Message)"
                }
            }
            
            # Create result object for EVERY file
            $result = [PSCustomObject]@{
                TenantId = $config.authentication.appOnly.tenantId
                SiteId = $targetSite.id
                SiteName = $targetSite.displayName
                SiteUrl = $targetSite.webUrl
                DriveId = $drive.id
                LibraryName = $drive.name
                FolderPath = $file.FolderPath
                FileName = $file.Name
                FileExtension = $file.Extension
                FileWebUrl = $file.WebUrl
                FileSizeBytes = $file.Size
                CreatedBy = $file.CreatedBy
                CreatedDateTime = $file.CreatedDateTime
                LastModifiedDateTime = $file.LastModifiedDateTime
                LabelIds = $labelIds
                LabelNames = $labelNames
                AssignmentMethods = $assignmentMethods
                LabelExtractionStatus = $labelExtractionStatus
                ScanDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            $allResults += $result
        }
        
        # Clear progress bar
        if (-not $SilentMode) {
            Write-Progress -Activity "Processing $($drive.name)" -Completed
        }
    }
    
    # Export results or return data based on mode
    if ($allResults.Count -gt 0) {
        if (-not $ReturnData) {
            # Original behavior - export to CSV and show summary
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $siteNameClean = $targetSite.displayName -replace '[^\w\s-]', '' -replace '\s+', '_'
            $outputPath = ".\Output\$($siteNameClean)_Comprehensive_$timestamp.csv"
            
            $allResults | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
            
            Write-Host "`n=== COMPREHENSIVE SCAN RESULTS ===" -ForegroundColor Green
            Write-Host "Total files scanned: $($allResults.Count)" -ForegroundColor White
            Write-Host "Files with sensitivity labels: $(($allResults | Where-Object { $_.LabelIds -and $_.LabelIds -ne '' }).Count)" -ForegroundColor White
            Write-Host "Files without sensitivity labels: $(($allResults | Where-Object { -not $_.LabelIds -or $_.LabelIds -eq '' }).Count)" -ForegroundColor White
            Write-Host "Unique file extensions found: $(($allResults.FileExtension | Sort-Object -Unique | Where-Object { $_ }).Count)" -ForegroundColor White
            Write-Host "Unique folders scanned: $(($allResults.FolderPath | Sort-Object -Unique).Count)" -ForegroundColor White
            
            # File type breakdown
            Write-Host "`n=== FILE TYPE BREAKDOWN ===" -ForegroundColor Magenta
            $allResults | Group-Object FileExtension | Sort-Object Count -Descending | ForEach-Object {
                if ($_.Name) {
                    Write-Host "  $($_.Name) : $($_.Count) files" -ForegroundColor Yellow
                }
            }
            
            # Label statistics (show first 10)
            $filesWithLabels = $allResults | Where-Object { $_.LabelIds -and $_.LabelIds -ne '' }
            if ($filesWithLabels.Count -gt 0) {
                Write-Host "`n=== LABEL STATISTICS ===" -ForegroundColor Magenta
                $labelStats = @{}
                
                foreach ($file in $filesWithLabels) {
                    $ids = $file.LabelIds -split ';'
                    $names = $file.LabelNames -split ';'
                    
                    for ($i = 0; $i -lt $ids.Count; $i++) {
                        $id = $ids[$i].Trim()
                        $name = if ($i -lt $names.Count -and $names[$i] -ne "") { $names[$i] } else { "Unknown: $id" }
                        if ($labelStats.ContainsKey($name)) {
                            $labelStats[$name]++
                        } else {
                            $labelStats[$name] = 1
                        }
                    }
                }
                
                $labelStats.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
                    Write-Host "  $($_.Key) : $($_.Value) files" -ForegroundColor Green
                }
            }
            
            # Folder breakdown
            Write-Host "`n=== FOLDER BREAKDOWN ===" -ForegroundColor Magenta
            $allResults | Group-Object FolderPath | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
                Write-Host "  $($_.Name) : $($_.Count) files" -ForegroundColor Cyan
            }
            
            Write-Host "`nFull results exported to: $outputPath" -ForegroundColor Green
            Write-Host "Open this file in Excel or Power BI for detailed analysis!" -ForegroundColor Yellow
        } else {
            # Return data for master script consumption
            return $allResults
        }
        
    } else {
        if ($ReturnData) {
            return @()  # Return empty array
        } else {
            Write-Warning "No files found to process"
        }
    }
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
} finally {
    if (-not $SkipModuleImport) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}

if (-not $SilentMode) {
    Write-Host "`n=== Comprehensive Scan Complete ===" -ForegroundColor Green
}
