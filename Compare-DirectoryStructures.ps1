<#
.SYNOPSIS
    Compares two directory structures and identifies differences for migration analysis.

.DESCRIPTION
    Analyzes two directories with deep folder nesting (e.g., Dropbox vs OneDrive) to identify:
    - Missing files and folders
    - Size differences
    - Directory structure variations
    - Generates actionable reports for user review and batch processing

.PARAMETER SourceDirectory
    The source directory path (e.g., Dropbox folder)

.PARAMETER TargetDirectory
    The target directory path (e.g., OneDrive folder)

.PARAMETER OutputDirectory
    Directory for output reports and CSV files

.PARAMETER IncludeHiddenFiles
    Include hidden files and system files in the analysis

.PARAMETER MaxDepth
    Maximum directory depth to scan (default: unlimited)

.PARAMETER ExcludePatterns
    Array of patterns to exclude (e.g., "*.tmp", "Thumbs.db")

.PARAMETER GenerateStructureReport
    Generate a detailed directory structure report

.PARAMETER GenerateCSVForActions
    Generate CSV file for user actions (Copy, Delete, Ignore)

.EXAMPLE
    .\Compare-DirectoryStructures.ps1 -SourceDirectory "C:\Users\John\Dropbox" -TargetDirectory "C:\Users\John\OneDrive" -GenerateStructureReport -GenerateCSVForActions

.EXAMPLE
    .\Compare-DirectoryStructures.ps1 -SourceDirectory "C:\OldData" -TargetDirectory "C:\NewData" -ExcludePatterns @("*.tmp", "*.log", "Thumbs.db") -MaxDepth 10

.NOTES
    Author: PowerShell Directory Comparison Toolkit
    Version: 1.0
    Requires: PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$SourceDirectory,
    
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$TargetDirectory,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = ".\DirectoryComparison_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeHiddenFiles,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 0,
    
    [Parameter(Mandatory = $false)]
    [string[]]$ExcludePatterns = @(),
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateStructureReport,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateCSVForActions
)

# Ensure output directory exists
if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

# Write progress and logging functions
function Write-Progress-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(switch($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    })
}

function Get-DirectoryAnalysis {
    param(
        [string]$Path,
        [string]$Label,
        [string[]]$ExcludePatterns,
        [bool]$IncludeHidden,
        [int]$MaxDepth
    )
    
    Write-Progress-Log "Analyzing $Label directory: $Path"
    
    $files = @()
    $directories = @()
    $totalSize = 0
    $fileCount = 0
    $dirCount = 0
    
    try {
        # Get all items recursively
        $getChildItemParams = @{
            Path = $Path
            Recurse = $true
            ErrorAction = 'SilentlyContinue'
        }
        
        if ($IncludeHidden) {
            $getChildItemParams.Force = $true
        }
        
        $allItems = Get-ChildItem @getChildItemParams
        
        foreach ($item in $allItems) {
            # Check if item matches exclude patterns
            $excluded = $false
            foreach ($pattern in $ExcludePatterns) {
                if ($item.Name -like $pattern) {
                    $excluded = $true
                    break
                }
            }
            
            if ($excluded) { continue }
            
            # Check depth if specified
            if ($MaxDepth -gt 0) {
                $relativePath = $item.FullName.Substring($Path.Length)
                $depth = ($relativePath.Split([System.IO.Path]::DirectorySeparator, [System.StringSplitOptions]::RemoveEmptyEntries)).Count
                if ($depth > $MaxDepth) { continue }
            }
            
            $relativePath = $item.FullName.Substring($Path.Length + 1)
            
            if ($item.PSIsContainer) {
                $directories += [PSCustomObject]@{
                    Type = "Directory"
                    RelativePath = $relativePath
                    FullPath = $item.FullName
                    Name = $item.Name
                    LastModified = $item.LastWriteTime
                    Created = $item.CreationTime
                    Size = 0
                    SizeFormatted = "0 B"
                }
                $dirCount++
            } else {
                $size = if ($item.Length) { $item.Length } else { 0 }
                $totalSize += $size
                
                $files += [PSCustomObject]@{
                    Type = "File"
                    RelativePath = $relativePath
                    FullPath = $item.FullPath
                    Name = $item.Name
                    Extension = $item.Extension
                    LastModified = $item.LastWriteTime
                    Created = $item.CreationTime
                    Size = $size
                    SizeFormatted = Format-FileSize $size
                    Directory = Split-Path $relativePath -Parent
                }
                $fileCount++
            }
        }
    }
    catch {
        Write-Progress-Log "Error analyzing $Label directory: $($_.Exception.Message)" "ERROR"
    }
    
    Write-Progress-Log "$Label analysis complete - Files: $fileCount, Directories: $dirCount, Total Size: $(Format-FileSize $totalSize)"
    
    return @{
        Files = $files
        Directories = $directories
        TotalSize = $totalSize
        FileCount = $fileCount
        DirectoryCount = $dirCount
        Path = $Path
        Label = $Label
    }
}

function Format-FileSize {
    param([long]$Size)
    
    if ($Size -eq 0) { return "0 B" }
    
    $units = @("B", "KB", "MB", "GB", "TB")
    $unitIndex = 0
    $sizeDouble = [double]$Size
    
    while ($sizeDouble -ge 1024 -and $unitIndex -lt ($units.Length - 1)) {
        $sizeDouble /= 1024
        $unitIndex++
    }
    
    return "{0:N2} {1}" -f $sizeDouble, $units[$unitIndex]
}

function Compare-DirectoryStructures {
    param(
        [object]$SourceAnalysis,
        [object]$TargetAnalysis
    )
    
    Write-Progress-Log "Comparing directory structures..."
    
    $differences = @()
    
    # Create lookup hashtables for faster comparison
    $sourceFiles = @{}
    $targetFiles = @{}
    $sourceDirs = @{}
    $targetDirs = @{}
    
    foreach ($file in $SourceAnalysis.Files) {
        $sourceFiles[$file.RelativePath] = $file
    }
    
    foreach ($file in $TargetAnalysis.Files) {
        $targetFiles[$file.RelativePath] = $file
    }
    
    foreach ($dir in $SourceAnalysis.Directories) {
        $sourceDirs[$dir.RelativePath] = $dir
    }
    
    foreach ($dir in $TargetAnalysis.Directories) {
        $targetDirs[$dir.RelativePath] = $dir
    }
    
    # Find files only in source
    foreach ($relativePath in $sourceFiles.Keys) {
        if (-not $targetFiles.ContainsKey($relativePath)) {
            $sourceFile = $sourceFiles[$relativePath]
            $differences += [PSCustomObject]@{
                Type = "File"
                Status = "OnlyInSource"
                RelativePath = $relativePath
                SourcePath = $sourceFile.FullPath
                TargetPath = Join-Path $TargetAnalysis.Path $relativePath
                Name = $sourceFile.Name
                Extension = $sourceFile.Extension
                SourceSize = $sourceFile.Size
                TargetSize = 0
                SizeDifference = $sourceFile.Size
                SourceModified = $sourceFile.LastModified
                TargetModified = $null
                Recommendation = "Copy to target"
                Action = ""
                Notes = "File missing from $($TargetAnalysis.Label)"
            }
        }
    }
    
    # Find files only in target
    foreach ($relativePath in $targetFiles.Keys) {
        if (-not $sourceFiles.ContainsKey($relativePath)) {
            $targetFile = $targetFiles[$relativePath]
            $differences += [PSCustomObject]@{
                Type = "File"
                Status = "OnlyInTarget"
                RelativePath = $relativePath
                SourcePath = Join-Path $SourceAnalysis.Path $relativePath
                TargetPath = $targetFile.FullPath
                Name = $targetFile.Name
                Extension = $targetFile.Extension
                SourceSize = 0
                TargetSize = $targetFile.Size
                SizeDifference = -$targetFile.Size
                SourceModified = $null
                TargetModified = $targetFile.LastModified
                Recommendation = "Review for deletion or keep"
                Action = ""
                Notes = "File not in $($SourceAnalysis.Label)"
            }
        }
    }
    
    # Find files with differences
    foreach ($relativePath in $sourceFiles.Keys) {
        if ($targetFiles.ContainsKey($relativePath)) {
            $sourceFile = $sourceFiles[$relativePath]
            $targetFile = $targetFiles[$relativePath]
            
            $sizeDiff = $targetFile.Size - $sourceFile.Size
            $timeDiff = $targetFile.LastModified - $sourceFile.LastModified
            
            if ($sizeDiff -ne 0 -or [Math]::Abs($timeDiff.TotalSeconds) -gt 2) {
                $status = if ($sizeDiff -ne 0) { "SizeDifference" } else { "TimeDifference" }
                $recommendation = if ($sourceFile.LastModified -gt $targetFile.LastModified) { 
                    "Update target (source is newer)" 
                } elseif ($targetFile.LastModified -gt $sourceFile.LastModified) { 
                    "Source is older - review" 
                } else { 
                    "Size difference - review content" 
                }
                
                $differences += [PSCustomObject]@{
                    Type = "File"
                    Status = $status
                    RelativePath = $relativePath
                    SourcePath = $sourceFile.FullPath
                    TargetPath = $targetFile.FullPath
                    Name = $sourceFile.Name
                    Extension = $sourceFile.Extension
                    SourceSize = $sourceFile.Size
                    TargetSize = $targetFile.Size
                    SizeDifference = $sizeDiff
                    SourceModified = $sourceFile.LastModified
                    TargetModified = $targetFile.LastModified
                    Recommendation = $recommendation
                    Action = ""
                    Notes = "Files exist in both locations but differ"
                }
            }
        }
    }
    
    # Find directories only in source
    foreach ($relativePath in $sourceDirs.Keys) {
        if (-not $targetDirs.ContainsKey($relativePath)) {
            $sourceDir = $sourceDirs[$relativePath]
            $differences += [PSCustomObject]@{
                Type = "Directory"
                Status = "OnlyInSource"
                RelativePath = $relativePath
                SourcePath = $sourceDir.FullPath
                TargetPath = Join-Path $TargetAnalysis.Path $relativePath
                Name = $sourceDir.Name
                Extension = ""
                SourceSize = 0
                TargetSize = 0
                SizeDifference = 0
                SourceModified = $sourceDir.LastModified
                TargetModified = $null
                Recommendation = "Create directory in target"
                Action = ""
                Notes = "Directory structure missing from $($TargetAnalysis.Label)"
            }
        }
    }
    
    # Find directories only in target
    foreach ($relativePath in $targetDirs.Keys) {
        if (-not $sourceDirs.ContainsKey($relativePath)) {
            $targetDir = $targetDirs[$relativePath]
            $differences += [PSCustomObject]@{
                Type = "Directory"
                Status = "OnlyInTarget"
                RelativePath = $relativePath
                SourcePath = Join-Path $SourceAnalysis.Path $relativePath
                TargetPath = $targetDir.FullPath
                Name = $targetDir.Name
                Extension = ""
                SourceSize = 0
                TargetSize = 0
                SizeDifference = 0
                SourceModified = $null
                TargetModified = $targetDir.LastModified
                Recommendation = "Review for removal or keep"
                Action = ""
                Notes = "Directory not in $($SourceAnalysis.Label)"
            }
        }
    }
    
    Write-Progress-Log "Comparison complete - Found $($differences.Count) differences"
    return $differences
}

function Generate-StructureReport {
    param(
        [object]$SourceAnalysis,
        [object]$TargetAnalysis,
        [array]$Differences,
        [string]$OutputPath
    )
    
    Write-Progress-Log "Generating directory structure report..."
    
    $report = @"
# Directory Structure Analysis Report
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

## Summary
- **Source Directory**: $($SourceAnalysis.Path) ($($SourceAnalysis.Label))
- **Target Directory**: $($TargetAnalysis.Path) ($($TargetAnalysis.Label))
- **Total Differences Found**: $($Differences.Count)

## Source Directory Statistics
- **Total Files**: $($SourceAnalysis.FileCount)
- **Total Directories**: $($SourceAnalysis.DirectoryCount)
- **Total Size**: $(Format-FileSize $SourceAnalysis.TotalSize)

## Target Directory Statistics
- **Total Files**: $($TargetAnalysis.FileCount)
- **Total Directories**: $($TargetAnalysis.DirectoryCount)
- **Total Size**: $(Format-FileSize $TargetAnalysis.TotalSize)

## Difference Categories

### Files Only in Source ($($SourceAnalysis.Label))
$(($Differences | Where-Object { $_.Status -eq "OnlyInSource" -and $_.Type -eq "File" }).Count) files

### Files Only in Target ($($TargetAnalysis.Label))
$(($Differences | Where-Object { $_.Status -eq "OnlyInTarget" -and $_.Type -eq "File" }).Count) files

### Files with Size Differences
$(($Differences | Where-Object { $_.Status -eq "SizeDifference" }).Count) files

### Files with Time Differences
$(($Differences | Where-Object { $_.Status -eq "TimeDifference" }).Count) files

### Directories Only in Source
$(($Differences | Where-Object { $_.Status -eq "OnlyInSource" -and $_.Type -eq "Directory" }).Count) directories

### Directories Only in Target
$(($Differences | Where-Object { $_.Status -eq "OnlyInTarget" -and $_.Type -eq "Directory" }).Count) directories

## Recommendations

### Immediate Actions Required
1. **Missing Files**: Review files that exist in source but not in target
2. **Size Differences**: Verify files with different sizes for data integrity
3. **Directory Structure**: Ensure proper folder hierarchy in target

### Directory Structure Optimization
1. **Deep Nesting**: Consider flattening overly nested directories
2. **Empty Directories**: Remove unnecessary empty folders
3. **Naming Conventions**: Standardize folder and file naming

### Migration Best Practices
1. **Verify Critical Files**: Ensure all important files are properly transferred
2. **Check File Integrity**: Compare checksums for critical documents
3. **Update Shortcuts**: Update any shortcuts or bookmarks pointing to old locations

## Next Steps
1. Review the detailed CSV report for specific actions
2. Use the action CSV to mark files for Copy (C), Delete (D), or Ignore (I)
3. Run the action processor script to execute the marked actions
4. Perform final verification after migration

---
*Report generated by Directory Comparison Toolkit*
"@

    $report | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Progress-Log "Structure report saved to: $OutputPath"
}

function Generate-ActionCSV {
    param(
        [array]$Differences,
        [string]$OutputPath
    )
    
    Write-Progress-Log "Generating action CSV file..."
    
    # Sort differences by type and status for better organization
    $sortedDifferences = $Differences | Sort-Object Type, Status, RelativePath
    
    # Export to CSV with action column
    $sortedDifferences | Select-Object @(
        @{Name="Action"; Expression={""}},
        "Type",
        "Status", 
        "RelativePath",
        "Name",
        "Extension",
        "SourcePath",
        "TargetPath",
        @{Name="SourceSize"; Expression={Format-FileSize $_.SourceSize}},
        @{Name="TargetSize"; Expression={Format-FileSize $_.TargetSize}},
        @{Name="SizeDifference"; Expression={Format-FileSize $_.SizeDifference}},
        "SourceModified",
        "TargetModified",
        "Recommendation",
        "Notes"
    ) | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Progress-Log "Action CSV saved to: $OutputPath"
    Write-Progress-Log "Instructions: Open in Excel and use Action column - C=Copy, D=Delete, I=Ignore" "SUCCESS"
}

# Main execution
Write-Progress-Log "Starting directory structure comparison..." "SUCCESS"
Write-Progress-Log "Source: $SourceDirectory"
Write-Progress-Log "Target: $TargetDirectory"
Write-Progress-Log "Output: $OutputDirectory"

try {
    # Analyze both directories
    $sourceAnalysis = Get-DirectoryAnalysis -Path $SourceDirectory -Label "Source" -ExcludePatterns $ExcludePatterns -IncludeHidden $IncludeHiddenFiles -MaxDepth $MaxDepth
    $targetAnalysis = Get-DirectoryAnalysis -Path $TargetDirectory -Label "Target" -ExcludePatterns $ExcludePatterns -IncludeHidden $IncludeHiddenFiles -MaxDepth $MaxDepth
    
    # Compare structures
    $differences = Compare-DirectoryStructures -SourceAnalysis $sourceAnalysis -TargetAnalysis $targetAnalysis
    
    # Generate reports
    if ($GenerateStructureReport) {
        $structureReportPath = Join-Path $OutputDirectory "DirectoryStructureReport.md"
        Generate-StructureReport -SourceAnalysis $sourceAnalysis -TargetAnalysis $targetAnalysis -Differences $differences -OutputPath $structureReportPath
    }
    
    if ($GenerateCSVForActions -or $differences.Count -gt 0) {
        $csvPath = Join-Path $OutputDirectory "DirectoryDifferences_Actions.csv"
        Generate-ActionCSV -Differences $differences -OutputPath $csvPath
    }
    
    # Export detailed analysis to JSON
    $analysisData = @{
        SourceAnalysis = $sourceAnalysis
        TargetAnalysis = $targetAnalysis
        Differences = $differences
        GeneratedDate = Get-Date
        Parameters = @{
            SourceDirectory = $SourceDirectory
            TargetDirectory = $TargetDirectory
            IncludeHiddenFiles = $IncludeHiddenFiles
            MaxDepth = $MaxDepth
            ExcludePatterns = $ExcludePatterns
        }
    }
    
    $jsonPath = Join-Path $OutputDirectory "DirectoryAnalysis.json"
    $analysisData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
    
    Write-Progress-Log "Analysis complete!" "SUCCESS"
    Write-Progress-Log "Results saved to: $OutputDirectory"
    
    # Summary
    Write-Host "`n" -NoNewline
    Write-Progress-Log "COMPARISON SUMMARY:" "SUCCESS"
    Write-Progress-Log "Source Files: $($sourceAnalysis.FileCount), Directories: $($sourceAnalysis.DirectoryCount)"
    Write-Progress-Log "Target Files: $($targetAnalysis.FileCount), Directories: $($targetAnalysis.DirectoryCount)"
    Write-Progress-Log "Differences Found: $($differences.Count)"
    Write-Progress-Log "Files Only in Source: $(($differences | Where-Object { $_.Status -eq 'OnlyInSource' -and $_.Type -eq 'File' }).Count)"
    Write-Progress-Log "Files Only in Target: $(($differences | Where-Object { $_.Status -eq 'OnlyInTarget' -and $_.Type -eq 'File' }).Count)"
    
    if ($GenerateCSVForActions -or $differences.Count -gt 0) {
        Write-Host "`n" -NoNewline
        Write-Progress-Log "NEXT STEPS:" "SUCCESS"
        Write-Progress-Log "1. Open DirectoryDifferences_Actions.csv in Excel"
        Write-Progress-Log "2. Fill Action column: C=Copy, D=Delete, I=Ignore"
        Write-Progress-Log "3. Run Process-DirectoryActions.ps1 to execute actions"
    }
}
catch {
    Write-Progress-Log "Error during analysis: $($_.Exception.Message)" "ERROR"
    exit 1
}
