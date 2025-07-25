# Compare-InstalledApps.ps1
# PowerShell script to compare installed applications between two computers
# Analyzes differences, version mismatches, and groups similar applications

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Computer1File,
    
    [Parameter(Mandatory = $true)]
    [string]$Computer2File,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\AppComparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToJson,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToMarkdown,
    
    [Parameter(Mandatory = $false)]
    [switch]$GroupSimilarApps,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowVersionDifferences,
    
    [Parameter(Mandatory = $false)]
    [switch]$DetailedReport
)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(
        switch ($Level) {
            "ERROR" { "Red" }
            "WARN" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
    )
}

function Import-AppData {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        throw "File not found: $FilePath"
    }
    
    try {
        $data = Get-Content $FilePath -Raw | ConvertFrom-Json
        return $data
    }
    catch {
        throw "Error reading file $FilePath`: $($_.Exception.Message)"
    }
}

function Get-SimilarityScore {
    param([string]$String1, [string]$String2)
    
    $str1 = $String1.ToLower() -replace '[^a-z0-9]', ''
    $str2 = $String2.ToLower() -replace '[^a-z0-9]', ''
    
    if ($str1 -eq $str2) { return 100 }
    if ($str1.Length -eq 0 -or $str2.Length -eq 0) { return 0 }
    
    # Check for substring matches
    if ($str1.Contains($str2) -or $str2.Contains($str1)) { return 80 }
    
    # Calculate Levenshtein distance
    $matrix = New-Object 'int[,]' ($str1.Length + 1), ($str2.Length + 1)
    
    for ($i = 0; $i -le $str1.Length; $i++) { $matrix[$i, 0] = $i }
    for ($j = 0; $j -le $str2.Length; $j++) { $matrix[0, $j] = $j }
    
    for ($i = 1; $i -le $str1.Length; $i++) {
        for ($j = 1; $j -le $str2.Length; $j++) {
            $cost = if ($str1[$i - 1] -eq $str2[$j - 1]) { 0 } else { 1 }
            $matrix[$i, $j] = [Math]::Min([Math]::Min($matrix[$i - 1, $j] + 1, $matrix[$i, $j - 1] + 1), $matrix[$i - 1, $j - 1] + $cost)
        }
    }
    
    $distance = $matrix[$str1.Length, $str2.Length]
    $maxLength = [Math]::Max($str1.Length, $str2.Length)
    $similarity = [Math]::Round((1 - ($distance / $maxLength)) * 100, 2)
    
    return $similarity
}

function Group-SimilarApplications {
    param([array]$Apps1, [array]$Apps2)
    
    Write-Log "Grouping similar applications..."
    
    $groups = @()
    $processed = @()
    
    foreach ($app1 in $Apps1) {
        if ($app1.Name -in $processed) { continue }
        
        $group = @{
            GroupName = $app1.Name
            Computer1Apps = @($app1)
            Computer2Apps = @()
            SimilarApps = @()
        }
        
        # Find similar apps in same computer
        foreach ($otherApp1 in $Apps1) {
            if ($otherApp1.Name -ne $app1.Name -and $otherApp1.Name -notin $processed) {
                $similarity = Get-SimilarityScore -String1 $app1.Name -String2 $otherApp1.Name
                if ($similarity -ge 70) {
                    $group.Computer1Apps += $otherApp1
                    $processed += $otherApp1.Name
                }
            }
        }
        
        # Find similar apps in other computer
        foreach ($app2 in $Apps2) {
            $similarity = Get-SimilarityScore -String1 $app1.Name -String2 $app2.Name
            if ($similarity -ge 70) {
                $group.Computer2Apps += $app2
            }
        }
        
        $processed += $app1.Name
        $groups += $group
    }
    
    # Find apps in Computer2 that weren't matched
    foreach ($app2 in $Apps2) {
        $found = $false
        foreach ($group in $groups) {
            if ($app2.Name -in $group.Computer2Apps.Name) {
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            $group = @{
                GroupName = $app2.Name
                Computer1Apps = @()
                Computer2Apps = @($app2)
                SimilarApps = @()
            }
            $groups += $group
        }
    }
    
    return $groups
}

function Compare-Applications {
    param([object]$Data1, [object]$Data2)
    
    Write-Log "Comparing applications between $($Data1.ComputerName) and $($Data2.ComputerName)..."
    
    $apps1 = $Data1.Applications
    $apps2 = $Data2.Applications
    
    # Find apps only in Computer 1
    $onlyInComputer1 = @()
    foreach ($app1 in $apps1) {
        $match = $apps2 | Where-Object { $_.Name -eq $app1.Name }
        if (-not $match) {
            $onlyInComputer1 += $app1
        }
    }
    
    # Find apps only in Computer 2
    $onlyInComputer2 = @()
    foreach ($app2 in $apps2) {
        $match = $apps1 | Where-Object { $_.Name -eq $app2.Name }
        if (-not $match) {
            $onlyInComputer2 += $app2
        }
    }
    
    # Find common apps with version differences
    $versionDifferences = @()
    foreach ($app1 in $apps1) {
        $match = $apps2 | Where-Object { $_.Name -eq $app1.Name }
        if ($match -and $app1.Version -ne $match.Version) {
            $versionDifferences += [PSCustomObject]@{
                Name = $app1.Name
                Computer1Version = $app1.Version
                Computer2Version = $match.Version
                Computer1Publisher = $app1.Publisher
                Computer2Publisher = $match.Publisher
            }
        }
    }
    
    # Find common apps (same name and version)
    $commonApps = @()
    foreach ($app1 in $apps1) {
        $match = $apps2 | Where-Object { $_.Name -eq $app1.Name -and $_.Version -eq $app1.Version }
        if ($match) {
            $commonApps += $app1
        }
    }
    
    return @{
        OnlyInComputer1 = $onlyInComputer1 | Sort-Object Name
        OnlyInComputer2 = $onlyInComputer2 | Sort-Object Name
        VersionDifferences = $versionDifferences | Sort-Object Name
        CommonApps = $commonApps | Sort-Object Name
        TotalAppsComputer1 = $apps1.Count
        TotalAppsComputer2 = $apps2.Count
    }
}

function Generate-HtmlReport {
    param([object]$ComparisonResult, [object]$Data1, [object]$Data2, [string]$OutputPath)
    
    Write-Log "Generating HTML report..."
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Application Comparison Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        h2 { color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 10px; margin-top: 30px; }
        h3 { color: #2980b9; margin-top: 25px; }
        .summary { background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-top: 10px; }
        .summary-item { background-color: white; padding: 10px; border-radius: 5px; text-align: center; }
        .summary-number { font-size: 24px; font-weight: bold; color: #2c3e50; }
        .summary-label { font-size: 12px; color: #7f8c8d; text-transform: uppercase; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid #bdc3c7; padding: 8px; text-align: left; }
        th { background-color: #3498db; color: white; font-weight: bold; }
        tr:nth-child(even) { background-color: #f8f9fa; }
        tr:hover { background-color: #e8f4f8; }
        .only-computer1 { background-color: #ffe6e6; }
        .only-computer2 { background-color: #e6f3ff; }
        .version-diff { background-color: #fff2e6; }
        .common { background-color: #e6ffe6; }
        .highlight { font-weight: bold; }
        .computer-name { color: #2980b9; font-weight: bold; }
        .stats { display: flex; justify-content: space-around; margin: 20px 0; }
        .stat-box { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; border-radius: 8px; text-align: center; min-width: 120px; }
        .legend { margin: 20px 0; }
        .legend-item { display: inline-block; padding: 5px 10px; margin: 5px; border-radius: 3px; font-size: 12px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Application Comparison Report</h1>
        
        <div class="summary">
            <h3>Comparison Summary</h3>
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="summary-number">$($Data1.ComputerName)</div>
                    <div class="summary-label">Computer 1</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($Data2.ComputerName)</div>
                    <div class="summary-label">Computer 2</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($ComparisonResult.TotalAppsComputer1)</div>
                    <div class="summary-label">Apps on Computer 1</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($ComparisonResult.TotalAppsComputer2)</div>
                    <div class="summary-label">Apps on Computer 2</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($ComparisonResult.OnlyInComputer1.Count)</div>
                    <div class="summary-label">Only on Computer 1</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($ComparisonResult.OnlyInComputer2.Count)</div>
                    <div class="summary-label">Only on Computer 2</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($ComparisonResult.VersionDifferences.Count)</div>
                    <div class="summary-label">Version Differences</div>
                </div>
                <div class="summary-item">
                    <div class="summary-number">$($ComparisonResult.CommonApps.Count)</div>
                    <div class="summary-label">Common Apps</div>
                </div>
            </div>
        </div>

        <div class="legend">
            <strong>Legend:</strong>
            <span class="legend-item only-computer1">Only on $($Data1.ComputerName)</span>
            <span class="legend-item only-computer2">Only on $($Data2.ComputerName)</span>
            <span class="legend-item version-diff">Version Differences</span>
            <span class="legend-item common">Common Applications</span>
        </div>
"@

    # Add section for apps only in Computer 1
    if ($ComparisonResult.OnlyInComputer1.Count -gt 0) {
        $html += @"
        <h2>Applications Only on <span class="computer-name">$($Data1.ComputerName)</span> ($($ComparisonResult.OnlyInComputer1.Count) apps)</h2>
        <table>
            <tr>
                <th>Application Name</th>
                <th>Version</th>
                <th>Publisher</th>
                <th>Source</th>
            </tr>
"@
        foreach ($app in $ComparisonResult.OnlyInComputer1) {
            $html += @"
            <tr class="only-computer1">
                <td>$($app.Name)</td>
                <td>$($app.Version)</td>
                <td>$($app.Publisher)</td>
                <td>$($app.Source)</td>
            </tr>
"@
        }
        $html += "</table>"
    }

    # Add section for apps only in Computer 2
    if ($ComparisonResult.OnlyInComputer2.Count -gt 0) {
        $html += @"
        <h2>Applications Only on <span class="computer-name">$($Data2.ComputerName)</span> ($($ComparisonResult.OnlyInComputer2.Count) apps)</h2>
        <table>
            <tr>
                <th>Application Name</th>
                <th>Version</th>
                <th>Publisher</th>
                <th>Source</th>
            </tr>
"@
        foreach ($app in $ComparisonResult.OnlyInComputer2) {
            $html += @"
            <tr class="only-computer2">
                <td>$($app.Name)</td>
                <td>$($app.Version)</td>
                <td>$($app.Publisher)</td>
                <td>$($app.Source)</td>
            </tr>
"@
        }
        $html += "</table>"
    }

    # Add section for version differences
    if ($ComparisonResult.VersionDifferences.Count -gt 0) {
        $html += @"
        <h2>Applications with Version Differences ($($ComparisonResult.VersionDifferences.Count) apps)</h2>
        <table>
            <tr>
                <th>Application Name</th>
                <th>$($Data1.ComputerName) Version</th>
                <th>$($Data2.ComputerName) Version</th>
                <th>Publisher</th>
            </tr>
"@
        foreach ($app in $ComparisonResult.VersionDifferences) {
            $html += @"
            <tr class="version-diff">
                <td>$($app.Name)</td>
                <td class="highlight">$($app.Computer1Version)</td>
                <td class="highlight">$($app.Computer2Version)</td>
                <td>$($app.Computer1Publisher)</td>
            </tr>
"@
        }
        $html += "</table>"
    }

    # Add section for common apps if detailed report is requested
    if ($DetailedReport -and $ComparisonResult.CommonApps.Count -gt 0) {
        $html += @"
        <h2>Common Applications ($($ComparisonResult.CommonApps.Count) apps)</h2>
        <table>
            <tr>
                <th>Application Name</th>
                <th>Version</th>
                <th>Publisher</th>
                <th>Source</th>
            </tr>
"@
        foreach ($app in $ComparisonResult.CommonApps) {
            $html += @"
            <tr class="common">
                <td>$($app.Name)</td>
                <td>$($app.Version)</td>
                <td>$($app.Publisher)</td>
                <td>$($app.Source)</td>
            </tr>
"@
        }
        $html += "</table>"
    }

    $html += @"
        <div style="margin-top: 30px; text-align: center; color: #7f8c8d; font-size: 12px;">
            <p>Report generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    return $OutputPath
}

# Main execution
try {
    Write-Log "Starting application comparison..." "SUCCESS"
    
    # Import data from both computers
    Write-Log "Loading data from $Computer1File..."
    $data1 = Import-AppData -FilePath $Computer1File
    
    Write-Log "Loading data from $Computer2File..."
    $data2 = Import-AppData -FilePath $Computer2File
    
    Write-Log "Computer 1: $($data1.ComputerName) - $($data1.Applications.Count) applications"
    Write-Log "Computer 2: $($data2.ComputerName) - $($data2.Applications.Count) applications"
    
    # Perform comparison
    $comparisonResult = Compare-Applications -Data1 $data1 -Data2 $data2
    
    # Group similar applications if requested
    if ($GroupSimilarApps) {
        Write-Log "Grouping similar applications..."
        $groupedApps = Group-SimilarApplications -Apps1 $data1.Applications -Apps2 $data2.Applications
        $comparisonResult.GroupedApps = $groupedApps
    }
    
    # Generate HTML report
    $htmlPath = Generate-HtmlReport -ComparisonResult $comparisonResult -Data1 $data1 -Data2 $data2 -OutputPath $OutputPath
    
    # Export to JSON if requested
    if ($ExportToJson) {
        $jsonPath = $OutputPath -replace '\.html$', '.json'
        $comparisonResult | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Log "JSON report exported to: $jsonPath"
    }
    
    # Export to Markdown if requested
    if ($ExportToMarkdown) {
        $markdownPath = $OutputPath -replace '\.html$', '.md'
        
        # Check if the Markdown report script exists
        if (Test-Path ".\New-MarkdownReport.ps1") {
            try {
                $markdownParams = @{
                    ComparisonData = $comparisonResult
                    Computer1Data = $data1
                    Computer2Data = $data2
                    OutputPath = $markdownPath
                    IncludeDetailedTables = $DetailedReport.IsPresent
                    IncludeStatistics = $true
                }
                
                & ".\New-MarkdownReport.ps1" @markdownParams
                Write-Log "Markdown report exported to: $markdownPath"
            }
            catch {
                Write-Log "Error generating Markdown report: $($_.Exception.Message)" "WARN"
            }
        } else {
            Write-Log "New-MarkdownReport.ps1 not found. Skipping Markdown export." "WARN"
        }
    }
    
    Write-Log "Comparison completed successfully!" "SUCCESS"
    Write-Log "HTML report generated: $htmlPath" "SUCCESS"
    
    # Display summary
    Write-Host "`nCOMPARISON SUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer 1: $($data1.ComputerName) ($($comparisonResult.TotalAppsComputer1) apps)" -ForegroundColor White
    Write-Host "Computer 2: $($data2.ComputerName) ($($comparisonResult.TotalAppsComputer2) apps)" -ForegroundColor White
    Write-Host ""
    Write-Host "Only on $($data1.ComputerName): $($comparisonResult.OnlyInComputer1.Count) apps" -ForegroundColor Red
    Write-Host "Only on $($data2.ComputerName): $($comparisonResult.OnlyInComputer2.Count) apps" -ForegroundColor Blue
    Write-Host "Version differences: $($comparisonResult.VersionDifferences.Count) apps" -ForegroundColor Yellow
    Write-Host "Common applications: $($comparisonResult.CommonApps.Count) apps" -ForegroundColor Green
    Write-Host ""
    Write-Host "Report saved to: $htmlPath" -ForegroundColor Cyan
    
    # Open the HTML report
    if (Test-Path $htmlPath) {
        Start-Process $htmlPath
    }
    
}
catch {
    Write-Log "Critical error during comparison: $($_.Exception.Message)" "ERROR"
    exit 1
}
