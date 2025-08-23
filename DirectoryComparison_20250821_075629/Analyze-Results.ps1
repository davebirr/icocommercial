# Quick analysis of directory comparison results
$csvPath = "DirectoryDifferences_Actions.csv"
$csv = Import-Csv $csvPath

Write-Host "=== DIRECTORY COMPARISON ANALYSIS ===" -ForegroundColor Green
Write-Host "Total entries: $($csv.Count)" -ForegroundColor Yellow
Write-Host ""

Write-Host "STATUS BREAKDOWN:" -ForegroundColor Cyan
$statusGroups = $csv | Group-Object Status | Sort-Object Count -Descending
foreach ($group in $statusGroups) {
    Write-Host "  $($group.Name): $($group.Count)" -ForegroundColor White
}
Write-Host ""

Write-Host "TYPE BREAKDOWN:" -ForegroundColor Cyan
$typeGroups = $csv | Group-Object Type | Sort-Object Count -Descending
foreach ($group in $typeGroups) {
    Write-Host "  $($group.Name): $($group.Count)" -ForegroundColor White
}
Write-Host ""

Write-Host "TOP 10 MISSING DIRECTORY PATTERNS:" -ForegroundColor Cyan
$missingDirs = $csv | Where-Object {$_.Status -eq 'OnlyInSource' -and $_.Type -eq 'Directory'}
$topPatterns = $missingDirs | ForEach-Object { 
    $parts = $_.RelativePath.Split('\')
    if ($parts.Length -ge 3) { $parts[0..2] -join '\' } else { $_.RelativePath }
} | Group-Object | Sort-Object Count -Descending | Select-Object -First 10

foreach ($pattern in $topPatterns) {
    Write-Host "  $($pattern.Name): $($pattern.Count)" -ForegroundColor White
}
Write-Host ""

Write-Host "FILE EXTENSION ANALYSIS (OnlyInSource):" -ForegroundColor Cyan
$missingFiles = $csv | Where-Object {$_.Status -eq 'OnlyInSource' -and $_.Type -eq 'File'}
$extGroups = $missingFiles | Group-Object Extension | Sort-Object Count -Descending | Select-Object -First 10
foreach ($group in $extGroups) {
    $ext = if ([string]::IsNullOrEmpty($group.Name)) { "(no extension)" } else { $group.Name }
    Write-Host "  ${ext}: $($group.Count)" -ForegroundColor White
}
Write-Host ""

Write-Host "LARGEST MISSING FILES (>10MB):" -ForegroundColor Cyan
$largeFiles = $missingFiles | Where-Object { 
    $_.SourceSize -match '(\d+\.?\d*)\s*(MB|GB|TB)' 
} | Sort-Object { 
    $size = $_.SourceSize
    if ($size -match '(\d+\.?\d*)\s*GB') { [double]$matches[1] * 1000 }
    elseif ($size -match '(\d+\.?\d*)\s*TB') { [double]$matches[1] * 1000000 }
    elseif ($size -match '(\d+\.?\d*)\s*MB') { [double]$matches[1] }
    else { 0 }
} -Descending | Select-Object -First 5

foreach ($file in $largeFiles) {
    Write-Host "  $($file.Name) ($($file.SourceSize))" -ForegroundColor White
}

Write-Host ""
Write-Host "ANALYSIS COMPLETE" -ForegroundColor Green
