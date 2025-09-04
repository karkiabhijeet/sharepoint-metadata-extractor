# Enhanced Configuration Support
function Get-ConfigFromFile {
    <#
    .SYNOPSIS
        Loads configuration from JSON file and merges with script parameters.
    #>
    param(
        [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
        [hashtable]$ScriptParams = @{}
    )
    
    $config = @{}
    
    if (Test-Path $ConfigPath) {
        Write-Host "Loading configuration from: $ConfigPath" -ForegroundColor Green
        $jsonConfig = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        
        # Convert JSON to hashtable recursively
        $config = ConvertTo-Hashtable $jsonConfig
        Write-Host "Configuration loaded successfully." -ForegroundColor Green
    } else {
        Write-Warning "Config file not found: $ConfigPath. Using default values."
        $config = Get-DefaultConfig
    }
    
    # Merge script parameters (they take precedence)
    foreach ($key in $ScriptParams.Keys) {
        $config[$key] = $ScriptParams[$key]
    }
    
    return $config
}

function ConvertTo-Hashtable {
    <#
    .SYNOPSIS
        Recursively converts PSCustomObject to hashtable.
    #>
    param([Parameter(ValueFromPipeline)]$InputObject)
    
    process {
        if ($null -eq $InputObject) { return $null }
        
        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $collection = @(
                foreach ($object in $InputObject) { ConvertTo-Hashtable $object }
            )
            Write-Output -NoEnumerate $collection
        } elseif ($InputObject -is [psobject]) {
            $hash = @{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-Hashtable $property.Value
            }
            $hash
        } else {
            $InputObject
        }
    }
}

function Get-DefaultConfig {
    <#
    .SYNOPSIS
        Returns default configuration when config file is not available.
    #>
    return @{
        authentication = @{
            preferredMethod = "interactive"
            appOnly = @{
                tenantId = ""
                clientId = ""
                certThumbprint = ""
                certStore = "CurrentUser"
            }
        }
        siteFilters = @{
            includeSiteSearch = @("*")
            excludeSiteUrls = @()
        }
        libraryFilters = @{
            includeLibraries = @()
            excludeLibraries = @(
                'Form Templates','Preservation Hold Library','Site Assets','Site Pages',
                'Images','Pages','Settings','Style Library','AppPages','Apps for SharePoint',
                'Apps for Office','Site Collection Documents','Site Collection Images','Documents Shared with Everyone'
            )
        }
        fileProcessing = @{
            extensions = @('docx','xlsx','pptx','pdf')
        }
        performance = @{
            graphPageSize = 200
            maxRetries = 5
            baseDelaySeconds = 3
            batchSize = 1000
            enableProgressReporting = $true
        }
        output = @{
            exportFormat = "csv"
            fileNaming = @{
                useTimestamp = $true
                prefix = "SPO_SensitivityLabels"
            }
            splitting = @{
                enabled = $false
                maxRecordsPerFile = 100000
            }
        }
        largeTenant = @{
            enabled = $false
            streamingMode = $false
            checkpointInterval = 10000
        }
    }
}

# Export Management for Large Datasets
function Initialize-ExportStrategy {
    <#
    .SYNOPSIS
        Initializes the appropriate export strategy based on configuration.
    #>
    param(
        [hashtable]$Config,
        [string]$BasePath
    )
    
    $strategy = @{
        Format = $Config.output.exportFormat
        BasePath = $BasePath
        CurrentFileIndex = 1
        RecordsInCurrentFile = 0
        MaxRecordsPerFile = $Config.output.splitting.maxRecordsPerFile
        SplittingEnabled = $Config.output.splitting.enabled
        SplitBySite = $Config.output.splitting.splitBySite
        StreamingMode = $Config.largeTenant.streamingMode
        CurrentSite = ""
        FileHandles = @{}
    }
    
    # Initialize first output file
    if ($strategy.SplittingEnabled) {
        $strategy.CurrentFilePath = Get-NextOutputFilePath -Strategy $strategy
    } else {
        $strategy.CurrentFilePath = $BasePath
    }
    
    return $strategy
}

function Get-NextOutputFilePath {
    <#
    .SYNOPSIS
        Generates the next output file path for file splitting.
    #>
    param([hashtable]$Strategy)
    
    $directory = Split-Path $Strategy.BasePath -Parent
    $baseName = Split-Path $Strategy.BasePath -LeafBase
    $extension = Split-Path $Strategy.BasePath -Extension
    
    if ($Strategy.SplitBySite -and $Strategy.CurrentSite) {
        $safeSiteName = $Strategy.CurrentSite -replace '[^\w\-]', '_'
        $fileName = "{0}_{1}{2}" -f $baseName, $safeSiteName, $extension
    } else {
        $fileName = "{0}_Part{1:D3}{2}" -f $baseName, $Strategy.CurrentFileIndex, $extension
    }
    
    return Join-Path $directory $fileName
}

function Export-Records {
    <#
    .SYNOPSIS
        Exports records using the configured strategy (streaming, splitting, etc.).
    #>
    param(
        [System.Collections.Generic.List[object]]$Records,
        [hashtable]$Strategy,
        [hashtable]$Config
    )
    
    if ($Strategy.StreamingMode) {
        Export-RecordsStreaming -Records $Records -Strategy $Strategy -Config $Config
    } else {
        Export-RecordsStandard -Records $Records -Strategy $Strategy -Config $Config
    }
}

function Export-RecordsStreaming {
    <#
    .SYNOPSIS
        Exports records in streaming mode for large datasets.
    #>
    param(
        [System.Collections.Generic.List[object]]$Records,
        [hashtable]$Strategy,
        [hashtable]$Config
    )
    
    foreach ($record in $Records) {
        # Check if we need to start a new file
        if ($Strategy.SplittingEnabled) {
            $needNewFile = $false
            
            if ($Strategy.SplitBySite -and $record.SiteName -ne $Strategy.CurrentSite) {
                $Strategy.CurrentSite = $record.SiteName
                $needNewFile = $true
            } elseif ($Strategy.RecordsInCurrentFile -ge $Strategy.MaxRecordsPerFile) {
                $needNewFile = $true
            }
            
            if ($needNewFile) {
                # Close current file handle if exists
                if ($Strategy.FileHandles.ContainsKey($Strategy.CurrentFilePath)) {
                    $Strategy.FileHandles[$Strategy.CurrentFilePath].Close()
                    $Strategy.FileHandles.Remove($Strategy.CurrentFilePath)
                }
                
                $Strategy.CurrentFileIndex++
                $Strategy.CurrentFilePath = Get-NextOutputFilePath -Strategy $Strategy
                $Strategy.RecordsInCurrentFile = 0
            }
        }
        
        # Export single record
        Export-SingleRecord -Record $record -Strategy $Strategy -Config $Config
        $Strategy.RecordsInCurrentFile++
    }
}

function Export-RecordsStandard {
    <#
    .SYNOPSIS
        Standard export for smaller datasets.
    #>
    param(
        [System.Collections.Generic.List[object]]$Records,
        [hashtable]$Strategy,
        [hashtable]$Config
    )
    
    switch ($Strategy.Format.ToLower()) {
        "csv" {
            $Records | Export-Csv -Path $Strategy.CurrentFilePath -NoTypeInformation -Encoding UTF8 -Append
        }
        "json" {
            if (Test-Path $Strategy.CurrentFilePath) {
                $existingData = Get-Content $Strategy.CurrentFilePath -Raw | ConvertFrom-Json
                $combinedData = @($existingData) + @($Records)
            } else {
                $combinedData = $Records
            }
            $combinedData | ConvertTo-Json -Depth 10 | Set-Content $Strategy.CurrentFilePath -Encoding UTF8
        }
        "parquet" {
            Write-Warning "Parquet export requires additional modules. Falling back to CSV."
            $Records | Export-Csv -Path ($Strategy.CurrentFilePath -replace '\.parquet$', '.csv') -NoTypeInformation -Encoding UTF8 -Append
        }
    }
}

function Export-SingleRecord {
    <#
    .SYNOPSIS
        Exports a single record in streaming mode.
    #>
    param(
        [object]$Record,
        [hashtable]$Strategy,
        [hashtable]$Config
    )
    
    switch ($Strategy.Format.ToLower()) {
        "csv" {
            # For CSV, we need to handle headers on first record
            $fileExists = Test-Path $Strategy.CurrentFilePath
            if (-not $fileExists) {
                $Record | Export-Csv -Path $Strategy.CurrentFilePath -NoTypeInformation -Encoding UTF8
            } else {
                $Record | Export-Csv -Path $Strategy.CurrentFilePath -NoTypeInformation -Encoding UTF8 -Append
            }
        }
        "json" {
            # For JSON streaming, append to JSONL format (one JSON object per line)
            $jsonLine = ($Record | ConvertTo-Json -Compress)
            Add-Content -Path ($Strategy.CurrentFilePath -replace '\.json$', '.jsonl') -Value $jsonLine -Encoding UTF8
        }
    }
}

# Checkpoint Management for Large Tenants
function Save-Checkpoint {
    <#
    .SYNOPSIS
        Saves processing checkpoint for resumable operations.
    #>
    param(
        [hashtable]$Config,
        [string]$CurrentSiteId,
        [string]$CurrentDriveId,
        [int]$ProcessedSites,
        [int]$ProcessedFiles,
        [string]$LastProcessedItemId = ""
    )
    
    if (-not $Config.largeTenant.enabled) { return }
    
    $checkpoint = @{
        Timestamp = Get-Date
        CurrentSiteId = $CurrentSiteId
        CurrentDriveId = $CurrentDriveId
        ProcessedSites = $ProcessedSites
        ProcessedFiles = $ProcessedFiles
        LastProcessedItemId = $LastProcessedItemId
        Config = $Config
    }
    
    $checkpointPath = Join-Path $PSScriptRoot $Config.largeTenant.checkpointFile
    $checkpoint | ConvertTo-Json -Depth 10 | Set-Content $checkpointPath -Encoding UTF8
    Write-Verbose "Checkpoint saved: $ProcessedFiles files processed"
}

function Restore-Checkpoint {
    <#
    .SYNOPSIS
        Restores processing state from checkpoint file.
    #>
    param([hashtable]$Config)
    
    if (-not $Config.largeTenant.enabled -or -not $Config.largeTenant.resumeFromCheckpoint) {
        return $null
    }
    
    $checkpointPath = Join-Path $PSScriptRoot $Config.largeTenant.checkpointFile
    if (-not (Test-Path $checkpointPath)) {
        Write-Host "No checkpoint file found. Starting fresh." -ForegroundColor Yellow
        return $null
    }
    
    try {
        $checkpoint = Get-Content $checkpointPath -Raw | ConvertFrom-Json
        Write-Host "Checkpoint restored: Resume from $($checkpoint.ProcessedFiles) processed files" -ForegroundColor Green
        return ConvertTo-Hashtable $checkpoint
    } catch {
        Write-Warning "Failed to restore checkpoint: $($_.Exception.Message). Starting fresh."
        return $null
    }
}

# Memory Optimization for Large Datasets
function Optimize-Memory {
    <#
    .SYNOPSIS
        Performs garbage collection and memory optimization.
    #>
    param([hashtable]$Config)
    
    if ($Config.largeTenant.memoryOptimization) {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}

# Progress Reporting for Large Operations
function Write-LargeTenantProgress {
    <#
    .SYNOPSIS
        Specialized progress reporting for large tenant operations.
    #>
    param(
        [int]$ProcessedSites,
        [int]$TotalSites,
        [int]$ProcessedFiles,
        [string]$CurrentOperation,
        [hashtable]$Config
    )
    
    if (-not $Config.performance.enableProgressReporting) { return }
    
    $percent = if ($TotalSites -gt 0) { [math]::Round(($ProcessedSites / $TotalSites) * 100, 1) } else { 0 }
    
    Write-Progress -Activity "Large Tenant Processing" -Status $CurrentOperation -PercentComplete $percent
    
    if ($ProcessedFiles % 1000 -eq 0) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$timestamp] Processed: $ProcessedFiles files, $ProcessedSites/$TotalSites sites" -ForegroundColor Cyan
    }
}