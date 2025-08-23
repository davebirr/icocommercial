# Focused Migration Analysis - Source vs Backup Drive
$csvPath = "DirectoryDifferences_Actions.csv"
$csv = Import-Csv $csvPath

Write-Host "=== FOCUSED MIGRATION ANALYSIS ===" -ForegroundColor Green
Write-Host "Source vs Backup Drive Comparison" -ForegroundColor Yellow
Write-Host "Total differences found: $($csv.Count)" -ForegroundColor Yellow
Write-Host ""

Write-Host "STATUS BREAKDOWN:" -ForegroundColor Cyan
$statusGroups = $csv | Group-Object Status | Sort-Object Count -Descending
foreach ($group in $statusGroups) {
    $percentage = [math]::Round(($group.Count / $csv.Count) * 100, 1)
    Write-Host "  $($group.Name): $($group.Count) ($percentage%)" -ForegroundColor White
}
Write-Host ""

Write-Host "TYPE BREAKDOWN:" -ForegroundColor Cyan
$typeGroups = $csv | Group-Object Type | Sort-Object Count -Descending
foreach ($group in $typeGroups) {
    $percentage = [math]::Round(($group.Count / $csv.Count) * 100, 1)
    Write-Host "  $($group.Name): $($group.Count) ($percentage%)" -ForegroundColor White
}
Write-Host ""

# Files only in source (missing from backup)
$onlyInSource = $csv | Where-Object { $_.Status -eq 'OnlyInSource' }
Write-Host "FILES MISSING FROM BACKUP (OnlyInSource): $($onlyInSource.Count)" -ForegroundColor Red
if ($onlyInSource.Count -gt 0) {
    Write-Host "File types missing from backup:" -ForegroundColor Yellow
    $sourceFileTypes = $onlyInSource | Where-Object { $_.Type -eq 'File' } | Group-Object Extension | Sort-Object Count -Descending | Select-Object -First 10
    foreach ($group in $sourceFileTypes) {
        $ext = if ([string]::IsNullOrEmpty($group.Name)) { "(no extension)" } else { $group.Name }
        Write-Host "  ${ext}: $($group.Count)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Top directory patterns missing from backup:" -ForegroundColor Yellow
    $sourceDirPatterns = $onlyInSource | Where-Object { $_.Type -eq 'Directory' } | ForEach-Object { 
        $parts = $_.RelativePath.Split('\')
        if ($parts.Length -ge 2) { $parts[0..1] -join '\' } else { $_.RelativePath }
    } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5
    
    foreach ($pattern in $sourceDirPatterns) {
        Write-Host "  $($pattern.Name): $($pattern.Count)" -ForegroundColor White
    }
}
Write-Host ""

# Files only in target (extra in backup)
$onlyInTarget = $csv | Where-Object { $_.Status -eq 'OnlyInTarget' }
Write-Host "EXTRA FILES IN BACKUP (OnlyInTarget): $($onlyInTarget.Count)" -ForegroundColor Blue
if ($onlyInTarget.Count -gt 0) {
    Write-Host "File types extra in backup:" -ForegroundColor Yellow
    $targetFileTypes = $onlyInTarget | Where-Object { $_.Type -eq 'File' } | Group-Object Extension | Sort-Object Count -Descending | Select-Object -First 10
    foreach ($group in $targetFileTypes) {
        $ext = if ([string]::IsNullOrEmpty($group.Name)) { "(no extension)" } else { $group.Name }
        Write-Host "  ${ext}: $($group.Count)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Top directory patterns extra in backup:" -ForegroundColor Yellow
    $targetDirPatterns = $onlyInTarget | Where-Object { $_.Type -eq 'Directory' } | ForEach-Object { 
        $parts = $_.RelativePath.Split('\')
        if ($parts.Length -ge 2) { $parts[0..1] -join '\' } else { $_.RelativePath }
    } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5
    
    foreach ($pattern in $targetDirPatterns) {
        Write-Host "  $($pattern.Name): $($pattern.Count)" -ForegroundColor White
    }
}
Write-Host ""

# Size and time differences
$differences = $csv | Where-Object { $_.Status -like '*Difference' }
Write-Host "VERSION CONFLICTS (Size/Time Differences): $($differences.Count)" -ForegroundColor Yellow
if ($differences.Count -gt 0) {
    Write-Host "Types of differences:" -ForegroundColor Yellow
    $diffTypes = $differences | Group-Object Status
    foreach ($group in $diffTypes) {
        Write-Host "  $($group.Name): $($group.Count)" -ForegroundColor White
    }
    
    if ($differences.Count -le 20) {
        Write-Host ""
        Write-Host "Specific files with differences:" -ForegroundColor Yellow
        foreach ($diff in $differences) {
            $source = if ($diff.SourceSize) { $diff.SourceSize } else { "N/A" }
            $target = if ($diff.TargetSize) { $diff.TargetSize } else { "N/A" }
            Write-Host "  $($diff.Name) - Source: $source, Target: $target" -ForegroundColor White
        }
    }
}
Write-Host ""

# Large files analysis
Write-Host "LARGE FILES MISSING FROM BACKUP (>10MB):" -ForegroundColor Cyan
$largeMissing = $onlyInSource | Where-Object { 
    $_.Type -eq 'File' -and $_.SourceSize -match '(\d+\.?\d*)\s*(MB|GB|TB)' 
} | Sort-Object { 
    $size = $_.SourceSize
    if ($size -match '(\d+\.?\d*)\s*GB') { [double]$matches[1] * 1000 }
    elseif ($size -match '(\d+\.?\d*)\s*TB') { [double]$matches[1] * 1000000 }
    elseif ($size -match '(\d+\.?\d*)\s*MB') { 
        $sizeValue = [double]$matches[1]
        if ($sizeValue -ge 10) { $sizeValue } else { 0 }
    }
    else { 0 }
} -Descending | Where-Object { $_ -gt 0 } | Select-Object -First 10

if ($largeMissing.Count -gt 0) {
    foreach ($file in $largeMissing) {
        Write-Host "  $($file.Name) ($($file.SourceSize))" -ForegroundColor White
    }
} else {
    Write-Host "  No large files missing from backup" -ForegroundColor Green
}

Write-Host ""
Write-Host "ANALYSIS COMPLETE" -ForegroundColor Green
