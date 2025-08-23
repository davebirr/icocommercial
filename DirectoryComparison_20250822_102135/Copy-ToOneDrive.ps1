<#
.SYNOPSIS
    Copies files marked with Action 'C' to OneDrive with adjusted paths.

.DESCRIPTION
    Reads the CSV file and copies files with Action 'C' to the user's OneDrive directory,
    adjusting the relative paths to start with CDT Personal or CDT Business.

.PARAMETER ActionCSVPath
    Path to the CSV file containing the actions

.PARAMETER OneDriveBasePath
    Base path to the user's OneDrive directory (default: C:\Users\cameront\OneDrive)

.PARAMETER WhatIf
    Shows what actions would be performed without actually executing them

.PARAMETER LogFile
    Path for detailed operation log file

.EXAMPLE
    .\Copy-ToOneDrive.ps1 -ActionCSVPath ".\DirectoryDifferences_Actions.csv" -WhatIf

.EXAMPLE
    .\Copy-ToOneDrive.ps1 -ActionCSVPath ".\DirectoryDifferences_Actions.csv"

.NOTES
    Author: PowerShell Directory Comparison Toolkit
    Version: 1.0
    Requires: PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ActionCSVPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OneDriveBasePath = "C:\Users\cameront\OneDrive",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [string]$LogFile = "OneDriveCopy_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

# Logging functions
function Write-CopyLog {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogPath = $LogFile
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console
    Write-Host $logMessage -ForegroundColor $(switch($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "SKIP" { "Cyan" }
        default { "White" }
    })
    
    # Write to log file
    $logMessage | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Get-OneDrivePath {
    param(
        [string]$RelativePath,
        [string]$OneDriveBase
    )
    
    # Remove "Cameron Tapley\" from the beginning of the path
    if ($RelativePath -like "Cameron Tapley\*") {
        $adjustedPath = $RelativePath.Substring("Cameron Tapley\".Length)
    } else {
        $adjustedPath = $RelativePath
    }
    
    # Combine with OneDrive base path
    $oneDrivePath = Join-Path $OneDriveBase $adjustedPath
    
    return $oneDrivePath
}

function Copy-FileToOneDrive {
    param(
        [string]$SourcePath,
        [string]$OneDrivePath,
        [bool]$IsWhatIf = $false
    )
    
    try {
        if ($IsWhatIf) {
            Write-CopyLog "[WHATIF] Would copy: $SourcePath -> $OneDrivePath" "INFO"
            return $true
        }
        
        # Create destination directory if it doesn't exist
        $destDir = Split-Path $OneDrivePath -Parent
        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            Write-CopyLog "Created directory: $destDir"
        }
        
        # Copy the file
        Copy-Item -Path $SourcePath -Destination $OneDrivePath -Force
        Write-CopyLog "Copied to OneDrive: $SourcePath -> $OneDrivePath" "SUCCESS"
        return $true
    }
    catch {
        Write-CopyLog "Failed to copy to OneDrive $SourcePath : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-ReconstructedSourcePath {
    param(
        [string]$RelativePath,
        [array]$AllActions
    )
    
    # Try to find a valid source path to extract the base directory
    $validEntry = $AllActions | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SourcePath) } | Select-Object -First 1
    
    if ($validEntry) {
        # Extract base path from valid entry
        $basePath = $validEntry.SourcePath
        $relativePart = $validEntry.RelativePath
        if ($basePath.EndsWith($relativePart)) {
            $sourceBasePath = $basePath.Substring(0, $basePath.Length - $relativePart.Length)
            $reconstructedPath = Join-Path $sourceBasePath $RelativePath
            return $reconstructedPath
        }
    }
    
    return $null
}

# Main execution
Write-CopyLog "Starting OneDrive copy operation..." "SUCCESS"
Write-CopyLog "CSV File: $ActionCSVPath"
Write-CopyLog "OneDrive Base Path: $OneDriveBasePath"

if ($WhatIf) {
    Write-CopyLog "WHATIF MODE - No actual changes will be made" "WARNING"
}

try {
    # Verify OneDrive path exists
    if (-not $WhatIf -and -not (Test-Path $OneDriveBasePath)) {
        Write-CopyLog "OneDrive path does not exist: $OneDriveBasePath" "ERROR"
        Write-CopyLog "Please verify the OneDrive path or create the directory first" "ERROR"
        exit 1
    }
    
    # Read the CSV file
    Write-CopyLog "Reading action CSV file..."
    $actions = Import-Csv -Path $ActionCSVPath
    
    if (-not $actions -or $actions.Count -eq 0) {
        Write-CopyLog "No actions found in CSV file" "WARNING"
        exit 0
    }
    
    # Filter for Copy actions only
    $copyActions = $actions | Where-Object { $_.Action -eq "C" }
    
    if (-not $copyActions -or $copyActions.Count -eq 0) {
        Write-CopyLog "No copy actions found in CSV file" "WARNING"
        exit 0
    }
    
    Write-CopyLog "Found $($copyActions.Count) files to copy to OneDrive"
    
    # Statistics
    $stats = @{
        Total = $copyActions.Count
        Copied = 0
        Failed = 0
        Skipped = 0
    }
    
    # Process each copy action
    foreach ($action in $copyActions) {
        Write-CopyLog "Processing: $($action.RelativePath)"
        
        # Get or reconstruct source path
        $sourcePath = $action.SourcePath
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            $sourcePath = Get-ReconstructedSourcePath -RelativePath $action.RelativePath -AllActions $actions
            if ([string]::IsNullOrWhiteSpace($sourcePath)) {
                Write-CopyLog "Cannot determine source path for: $($action.RelativePath)" "ERROR"
                $stats.Failed++
                continue
            }
            Write-CopyLog "Reconstructed source path: $sourcePath" "INFO"
        }
        
        # Verify source file exists
        if (-not $WhatIf -and -not (Test-Path $sourcePath)) {
            Write-CopyLog "Source file not found: $sourcePath" "ERROR"
            $stats.Failed++
            continue
        }
        
        # Get OneDrive destination path
        $oneDrivePath = Get-OneDrivePath -RelativePath $action.RelativePath -OneDriveBase $OneDriveBasePath
        
        # Copy the file
        $success = Copy-FileToOneDrive -SourcePath $sourcePath -OneDrivePath $oneDrivePath -IsWhatIf $WhatIf
        if ($success) {
            $stats.Copied++
        } else {
            $stats.Failed++
        }
    }
    
    # Final summary
    Write-CopyLog "" 
    Write-CopyLog "ONEDRIVE COPY OPERATION COMPLETE!" "SUCCESS"
    Write-CopyLog "Results Summary:" "INFO"
    Write-CopyLog "  Total Files: $($stats.Total)" "INFO"
    Write-CopyLog "  Files Copied: $($stats.Copied)" "INFO"
    Write-CopyLog "  Failed Operations: $($stats.Failed)" "INFO"
    Write-CopyLog "  Items Skipped: $($stats.Skipped)" "INFO"
    Write-CopyLog "Detailed log saved to: $LogFile" "INFO"
    
    if ($WhatIf) {
        Write-CopyLog "This was a WHATIF run - no actual changes were made" "INFO"
        Write-CopyLog "Run without -WhatIf to execute the OneDrive copy operations" "INFO"
    }
    
} catch {
    Write-CopyLog "Script execution failed: $($_.Exception.Message)" "ERROR"
    exit 1
}
