<#
.SYNOPSIS
    Example usage of Directory Structure Comparison tools for Dropbox to OneDrive migration.

.DESCRIPTION
    Demonstrates the complete workflow for analyzing and synchronizing directories
    between Dropbox and OneDrive, helping users identify missing files and folders.

.NOTES
    Author: PowerShell Directory Comparison Toolkit
    Version: 1.0
    Example scenarios included for common migration tasks
#>

Write-Host "=== Directory Structure Comparison - Example Usage ===" -ForegroundColor Green
Write-Host ""

# Example paths (adjust these to your actual directories)
$dropboxPath = "C:\Users\$env:USERNAME\Dropbox"
$onedrivePath = "C:\Users\$env:USERNAME\OneDrive"

Write-Host "This script demonstrates the complete workflow for comparing directory structures." -ForegroundColor Yellow
Write-Host "Common use case: User migrated from Dropbox to OneDrive and wants to verify all files transferred." -ForegroundColor Yellow
Write-Host ""

Write-Host "Step 1: Basic Directory Comparison" -ForegroundColor Cyan
Write-Host "Command: .\Compare-DirectoryStructures.ps1 -SourceDirectory `"$dropboxPath`" -TargetDirectory `"$onedrivePath`" -GenerateStructureReport -GenerateCSVForActions"
Write-Host ""

Write-Host "Step 2: Advanced Comparison with Exclusions" -ForegroundColor Cyan
Write-Host "Command: .\Compare-DirectoryStructures.ps1 -SourceDirectory `"$dropboxPath`" -TargetDirectory `"$onedrivePath`" -ExcludePatterns @('*.tmp','*.log','Thumbs.db','.DS_Store') -IncludeHiddenFiles -GenerateStructureReport -GenerateCSVForActions"
Write-Host ""

Write-Host "Step 3: Limited Depth Scan (for very large directories)" -ForegroundColor Cyan
Write-Host "Command: .\Compare-DirectoryStructures.ps1 -SourceDirectory `"$dropboxPath`" -TargetDirectory `"$onedrivePath`" -MaxDepth 5 -GenerateStructureReport -GenerateCSVForActions"
Write-Host ""

Write-Host "Step 4: Review and Edit Actions in Excel" -ForegroundColor Cyan
Write-Host "1. Open DirectoryDifferences_Actions.csv in Excel"
Write-Host "2. Fill the Action column with:"
Write-Host "   C = Copy file from source to target"
Write-Host "   D = Delete file (will be backed up first)"
Write-Host "   I = Ignore (do nothing)"
Write-Host ""

Write-Host "Step 5: Test Actions (What-If Mode)" -ForegroundColor Cyan
Write-Host "Command: .\Process-DirectoryActions.ps1 -ActionCSVPath `".\DirectoryComparison_*\DirectoryDifferences_Actions.csv`" -WhatIf"
Write-Host ""

Write-Host "Step 6: Execute Actions with Backup" -ForegroundColor Cyan
Write-Host "Command: .\Process-DirectoryActions.ps1 -ActionCSVPath `".\DirectoryComparison_*\DirectoryDifferences_Actions.csv`" -BackupDirectory `"C:\MigrationBackup`""
Write-Host ""

Write-Host "=== Real-World Example Scenarios ===" -ForegroundColor Green
Write-Host ""

Write-Host "Scenario 1: Complete Dropbox to OneDrive Migration Analysis" -ForegroundColor Yellow
$example1 = @'
# Full comparison excluding temporary files
.\Compare-DirectoryStructures.ps1 `
    -SourceDirectory "C:\Users\John\Dropbox" `
    -TargetDirectory "C:\Users\John\OneDrive" `
    -ExcludePatterns @("*.tmp", "*.log", "Thumbs.db", ".DS_Store", "desktop.ini") `
    -IncludeHiddenFiles `
    -GenerateStructureReport `
    -GenerateCSVForActions

# Review results and mark actions in Excel
# Then execute with backup
.\Process-DirectoryActions.ps1 `
    -ActionCSVPath ".\DirectoryComparison_20250125_143022\DirectoryDifferences_Actions.csv" `
    -BackupDirectory "C:\MigrationBackup\DropboxToOneDrive" `
    -LogFile "Migration_Log.txt"
'@
Write-Host $example1 -ForegroundColor White
Write-Host ""

Write-Host "Scenario 2: Selective Document Folder Migration" -ForegroundColor Yellow
$example2 = @'
# Compare only Documents subfolder with limited depth
.\Compare-DirectoryStructures.ps1 `
    -SourceDirectory "C:\Users\John\Dropbox\Documents" `
    -TargetDirectory "C:\Users\John\OneDrive\Documents" `
    -MaxDepth 3 `
    -ExcludePatterns @("*.tmp", "~*", "*.bak") `
    -GenerateStructureReport `
    -GenerateCSVForActions

# Test actions first
.\Process-DirectoryActions.ps1 `
    -ActionCSVPath ".\DirectoryComparison_*\DirectoryDifferences_Actions.csv" `
    -WhatIf

# Execute after review
.\Process-DirectoryActions.ps1 `
    -ActionCSVPath ".\DirectoryComparison_*\DirectoryDifferences_Actions.csv" `
    -BackupDirectory "C:\Backup\Documents" `
    -Force
'@
Write-Host $example2 -ForegroundColor White
Write-Host ""

Write-Host "Scenario 3: Large Directory with Performance Optimization" -ForegroundColor Yellow
$example3 = @'
# For very large directories, limit scope and exclude unnecessary files
.\Compare-DirectoryStructures.ps1 `
    -SourceDirectory "D:\BigDropboxFolder" `
    -TargetDirectory "D:\OneDrive\BigFolder" `
    -MaxDepth 4 `
    -ExcludePatterns @("*.tmp", "*.log", "*.cache", "node_modules", ".git", "*.iso", "*.zip") `
    -GenerateStructureReport `
    -GenerateCSVForActions

# Process in smaller batches if needed
# Can filter CSV file to process only certain file types or directories
'@
Write-Host $example3 -ForegroundColor White
Write-Host ""

Write-Host "=== Interpreting Results ===" -ForegroundColor Green
Write-Host ""

Write-Host "CSV Action Guide:" -ForegroundColor Cyan
Write-Host "Status 'OnlyInSource' + Action 'C' = Copy missing file from Dropbox to OneDrive"
Write-Host "Status 'OnlyInTarget' + Action 'D' = Delete extra file from OneDrive (backed up first)"
Write-Host "Status 'SizeDifference' + Action 'C' = Replace OneDrive file with Dropbox version"
Write-Host "Status 'TimeDifference' + Action 'I' = Ignore minor timestamp differences"
Write-Host ""

Write-Host "Common Migration Patterns:" -ForegroundColor Cyan
Write-Host "1. Copy all 'OnlyInSource' files (these are missing from OneDrive)"
Write-Host "2. Review 'OnlyInTarget' files (these might be OneDrive-specific)"
Write-Host "3. For 'SizeDifference', check which version is newer/larger"
Write-Host "4. Ignore 'TimeDifference' unless file content actually differs"
Write-Host ""

Write-Host "=== Safety Features ===" -ForegroundColor Green
Write-Host ""

Write-Host "Backup Protection:" -ForegroundColor Cyan
Write-Host "- All deleted files are backed up before removal"
Write-Host "- Backup maintains original directory structure"
Write-Host "- Use -WhatIf to test before actual execution"
Write-Host ""

Write-Host "Logging:" -ForegroundColor Cyan
Write-Host "- All operations are logged with timestamps"
Write-Host "- Success and failure details recorded"
Write-Host "- JSON analysis data preserved for later review"
Write-Host ""

Write-Host "=== Quick Start for Dropbox Migration ===" -ForegroundColor Green
Write-Host ""

$quickStart = @'
# 1. Run comparison
.\Compare-DirectoryStructures.ps1 -SourceDirectory "C:\Users\[Username]\Dropbox" -TargetDirectory "C:\Users\[Username]\OneDrive" -GenerateStructureReport -GenerateCSVForActions

# 2. Open CSV in Excel, mark actions (C/D/I)

# 3. Test first
.\Process-DirectoryActions.ps1 -ActionCSVPath ".\DirectoryComparison_*\DirectoryDifferences_Actions.csv" -WhatIf

# 4. Execute with backup
.\Process-DirectoryActions.ps1 -ActionCSVPath ".\DirectoryComparison_*\DirectoryDifferences_Actions.csv" -BackupDirectory "C:\MigrationBackup"
'@

Write-Host $quickStart -ForegroundColor White
Write-Host ""

Write-Host "=== Ready to Start? ===" -ForegroundColor Green
Write-Host ""

$choice = Read-Host "Would you like to run a comparison now? Enter paths or 'N' to exit"

if ($choice -ne 'N' -and $choice -ne 'n' -and ![string]::IsNullOrWhiteSpace($choice)) {
    Write-Host "Please provide the directory paths:" -ForegroundColor Yellow
    $source = Read-Host "Source directory (e.g., Dropbox path)"
    $target = Read-Host "Target directory (e.g., OneDrive path)"
    
    if ((Test-Path $source) -and (Test-Path $target)) {
        Write-Host "Starting comparison..." -ForegroundColor Green
        & ".\Compare-DirectoryStructures.ps1" -SourceDirectory $source -TargetDirectory $target -GenerateStructureReport -GenerateCSVForActions
    } else {
        Write-Host "One or both paths do not exist. Please check the paths and try again." -ForegroundColor Red
    }
} else {
    Write-Host "Example completed. Run the individual scripts when ready!" -ForegroundColor Green
}
