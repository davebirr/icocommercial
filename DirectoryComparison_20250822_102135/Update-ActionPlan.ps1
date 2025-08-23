<#
.SYNOPSIS
    Updates the Action column in the directory comparison CSV file based on status.

.DESCRIPTION
    Automatically fills in the Action column with appropriate actions:
    - OnlyInSource: C (Copy from source to target)
    - OnlyInTarget: I (Ignore - extra files in target)
    - SizeDifference/TimeDifference: I (Ignore - manual review needed)

.PARAMETER CSVPath
    Path to the DirectoryDifferences_Actions.csv file

.EXAMPLE
    .\Update-ActionPlan.ps1 -CSVPath ".\DirectoryDifferences_Actions.csv"

.NOTES
    Author: PowerShell Directory Comparison Toolkit
    Version: 1.0
    Updates CSV file in place with backup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CSVPath = ".\DirectoryDifferences_Actions.csv"
)

# Validate CSV file exists
if (-not (Test-Path $CSVPath)) {
    Write-Error "CSV file not found: $CSVPath"
    exit 1
}

Write-Host "=== UPDATING ACTION PLAN ===" -ForegroundColor Green
Write-Host "CSV File: $CSVPath" -ForegroundColor Yellow

try {
    # Create backup of original file
    $backupPath = $CSVPath -replace '\.csv$', "_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Copy-Item -Path $CSVPath -Destination $backupPath
    Write-Host "Backup created: $backupPath" -ForegroundColor Cyan

    # Read the CSV file
    Write-Host "Reading CSV file..." -ForegroundColor Yellow
    $data = Import-Csv -Path $CSVPath

    Write-Host "Total entries: $($data.Count)" -ForegroundColor White

    # Track changes
    $updated = 0
    $copyActions = 0
    $ignoreActions = 0

    # Update Action column based on Status
    foreach ($row in $data) {
        $originalAction = $row.Action
        
        switch ($row.Status) {
            "OnlyInSource" {
                $row.Action = "C"
                $copyActions++
            }
            "OnlyInTarget" {
                $row.Action = "I"
                $ignoreActions++
            }
            { $_ -like "*Difference" } {
                $row.Action = "I"
                $ignoreActions++
            }
            default {
                # Keep existing action or leave blank
                continue
            }
        }
        
        # Count updates
        if ($originalAction -ne $row.Action) {
            $updated++
        }
    }

    # Export updated CSV
    Write-Host "Updating CSV file..." -ForegroundColor Yellow
    $data | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8

    # Summary
    Write-Host "`n=== ACTION PLAN SUMMARY ===" -ForegroundColor Green
    Write-Host "Total entries processed: $($data.Count)" -ForegroundColor White
    Write-Host "Entries updated: $updated" -ForegroundColor White
    Write-Host ""
    Write-Host "ACTION BREAKDOWN:" -ForegroundColor Cyan
    Write-Host "  Copy (C): $copyActions items" -ForegroundColor Green
    Write-Host "  Ignore (I): $ignoreActions items" -ForegroundColor Yellow
    Write-Host ""
    
    # Breakdown by status
    $statusBreakdown = $data | Group-Object Status | Sort-Object Count -Descending
    Write-Host "STATUS BREAKDOWN:" -ForegroundColor Cyan
    foreach ($group in $statusBreakdown) {
        $action = ($data | Where-Object { $_.Status -eq $group.Name } | Select-Object -First 1).Action
        Write-Host "  $($group.Name): $($group.Count) items → Action: $action" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Green
    Write-Host "1. Review the updated CSV file: $CSVPath" -ForegroundColor White
    Write-Host "2. For SizeDifference items, manually copy to review folder" -ForegroundColor White
    Write-Host "3. Run the migration script:" -ForegroundColor White
    Write-Host "   .\Process-DirectoryActions.ps1 -ActionCSVPath `"$CSVPath`" -WhatIf" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Files with differences (manual review needed):" -ForegroundColor Yellow
    $diffFiles = $data | Where-Object { $_.Status -like "*Difference" }
    if ($diffFiles.Count -gt 0) {
        foreach ($file in $diffFiles) {
            Write-Host "  $($file.Name) - $($file.Status)" -ForegroundColor White
        }
    } else {
        Write-Host "  No files with differences found" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "UPDATE COMPLETE!" -ForegroundColor Green
}
catch {
    Write-Error "Error updating CSV file: $($_.Exception.Message)"
    exit 1
}
