# New-MarkdownReport.ps1
# PowerShell script to generate Markdown reports from application comparison data

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [object]$ComparisonData,
    
    [Parameter(Mandatory = $true)]
    [object]$Computer1Data,
    
    [Parameter(Mandatory = $true)]
    [object]$Computer2Data,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDetailedTables,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeStatistics
)

function Format-ApplicationTable {
    param([array]$Applications, [string]$TableTitle)
    
    if ($Applications.Count -eq 0) {
        return "`n*No applications found in this category.*`n"
    }
    
    $markdown = "`n### $TableTitle`n`n"
    $markdown += "| Application Name | Version | Publisher | Source |`n"
    $markdown += "|------------------|---------|-----------|--------|`n"
    
    foreach ($app in $Applications | Sort-Object Name) {
        $name = ($app.Name -replace '\|', '\|') -replace '`', '\`'
        $version = if ($app.Version) { ($app.Version -replace '\|', '\|') } else { "N/A" }
        $publisher = if ($app.Publisher) { ($app.Publisher -replace '\|', '\|') } else { "N/A" }
        $source = if ($app.Source) { ($app.Source -replace '\|', '\|') } else { "N/A" }
        
        $markdown += "| $name | $version | $publisher | $source |`n"
    }
    
    $markdown += "`n**Total: $($Applications.Count) applications**`n"
    return $markdown
}

function Format-VersionDifferenceTable {
    param([array]$VersionDifferences)
    
    if ($VersionDifferences.Count -eq 0) {
        return "`n*No version differences found.*`n"
    }
    
    $markdown = "`n### Applications with Version Differences`n`n"
    $markdown += "| Application Name | $($Computer1Data.ComputerName) Version | $($Computer2Data.ComputerName) Version | Publisher |`n"
    $markdown += "|------------------|------------|------------|-----------|`n"
    
    foreach ($app in $VersionDifferences | Sort-Object Name) {
        $name = ($app.Name -replace '\|', '\|') -replace '`', '\`'
        $version1 = if ($app.Computer1Version) { ($app.Computer1Version -replace '\|', '\|') } else { "N/A" }
        $version2 = if ($app.Computer2Version) { ($app.Computer2Version -replace '\|', '\|') } else { "N/A" }
        $publisher = if ($app.Computer1Publisher) { ($app.Computer1Publisher -replace '\|', '\|') } else { "N/A" }
        
        $markdown += "| $name | **$version1** | **$version2** | $publisher |`n"
    }
    
    $markdown += "`n**Total: $($VersionDifferences.Count) applications with version differences**`n"
    return $markdown
}

function New-ExecutiveSummary {
    param([object]$ComparisonData, [object]$Computer1Data, [object]$Computer2Data)
    
    $summary = @"
## Executive Summary

This report compares the installed applications between **$($Computer1Data.ComputerName)** and **$($Computer2Data.ComputerName)**.

### Key Findings

- **$($Computer1Data.ComputerName)** has **$($ComparisonData.TotalAppsComputer1)** installed applications
- **$($Computer2Data.ComputerName)** has **$($ComparisonData.TotalAppsComputer2)** installed applications
- **$($ComparisonData.OnlyInComputer1.Count)** applications are only installed on $($Computer1Data.ComputerName)
- **$($ComparisonData.OnlyInComputer2.Count)** applications are only installed on $($Computer2Data.ComputerName)
- **$($ComparisonData.VersionDifferences.Count)** applications have different versions between computers
- **$($ComparisonData.CommonApps.Count)** applications are identical on both computers

### Risk Assessment

"@

    # Add risk assessment based on findings
    if ($ComparisonData.OnlyInComputer1.Count -gt 10 -or $ComparisonData.OnlyInComputer2.Count -gt 10) {
        $summary += "🔴 **HIGH**: Significant differences in installed applications detected.`n`n"
    } elseif ($ComparisonData.VersionDifferences.Count -gt 5) {
        $summary += "🟡 **MEDIUM**: Multiple version differences detected that may require attention.`n`n"
    } else {
        $summary += "🟢 **LOW**: Computers have similar software configurations.`n`n"
    }

    return $summary
}

function New-StatisticsSection {
    param([object]$ComparisonData, [object]$Computer1Data, [object]$Computer2Data)
    
    $stats = @"
## Detailed Statistics

### Application Sources Breakdown

#### $($Computer1Data.ComputerName)
"@

    # Count applications by source for Computer 1
    $computer1Sources = $Computer1Data.Applications | Group-Object Source | Sort-Object Count -Descending
    foreach ($source in $computer1Sources) {
        $stats += "- **$($source.Name)**: $($source.Count) applications`n"
    }

    $stats += "`n#### $($Computer2Data.ComputerName)`n"
    
    # Count applications by source for Computer 2
    $computer2Sources = $Computer2Data.Applications | Group-Object Source | Sort-Object Count -Descending
    foreach ($source in $computer2Sources) {
        $stats += "- **$($source.Name)**: $($source.Count) applications`n"
    }

    # Publisher analysis
    $stats += @"

### Top Publishers

#### $($Computer1Data.ComputerName)
"@

    $computer1Publishers = $Computer1Data.Applications | Where-Object { $_.Publisher } | Group-Object Publisher | Sort-Object Count -Descending | Select-Object -First 10
    foreach ($pub in $computer1Publishers) {
        $stats += "- **$($pub.Name)**: $($pub.Count) applications`n"
    }

    $stats += "`n#### $($Computer2Data.ComputerName)`n"
    
    $computer2Publishers = $Computer2Data.Applications | Where-Object { $_.Publisher } | Group-Object Publisher | Sort-Object Count -Descending | Select-Object -First 10
    foreach ($pub in $computer2Publishers) {
        $stats += "- **$($pub.Name)**: $($pub.Count) applications`n"
    }

    return $stats
}

function New-RecommendationsSection {
    param([object]$ComparisonData, [object]$Computer1Data, [object]$Computer2Data)
    
    $recommendations = @"
## Recommendations

### Immediate Actions Required

"@

    # Check for critical software differences
    $criticalApps = @("Microsoft Office", "Adobe", "Antivirus", "Windows Defender", "Visual Studio", "SQL Server")
    $criticalDifferences = @()
    
    foreach ($app in $ComparisonData.OnlyInComputer1) {
        foreach ($critical in $criticalApps) {
            if ($app.Name -like "*$critical*") {
                $criticalDifferences += "⚠️ **$($app.Name)** is installed on $($Computer1Data.ComputerName) but missing from $($Computer2Data.ComputerName)"
            }
        }
    }
    
    foreach ($app in $ComparisonData.OnlyInComputer2) {
        foreach ($critical in $criticalApps) {
            if ($app.Name -like "*$critical*") {
                $criticalDifferences += "⚠️ **$($app.Name)** is installed on $($Computer2Data.ComputerName) but missing from $($Computer1Data.ComputerName)"
            }
        }
    }

    if ($criticalDifferences.Count -gt 0) {
        foreach ($diff in $criticalDifferences) {
            $recommendations += "$diff`n`n"
        }
    } else {
        $recommendations += "✅ No critical software differences detected.`n`n"
    }

    # Version difference recommendations
    if ($ComparisonData.VersionDifferences.Count -gt 0) {
        $recommendations += @"
### Version Synchronization

Consider updating the following applications to maintain consistency:

"@
        $priorityVersions = $ComparisonData.VersionDifferences | Where-Object { 
            $_.Name -like "*Security*" -or 
            $_.Name -like "*Antivirus*" -or 
            $_.Name -like "*Microsoft*" -or
            $_.Name -like "*Adobe*"
        } | Select-Object -First 5

        foreach ($app in $priorityVersions) {
            $recommendations += "- **$($app.Name)**: Update to consistent version across both computers`n"
        }
        $recommendations += "`n"
    }

    # Standardization recommendations
    $recommendations += @"
### Standardization Opportunities

1. **Software Deployment**: Consider using a software deployment tool to maintain consistency
2. **Update Management**: Implement a centralized update management strategy
3. **Application Inventory**: Regular comparison reports can help maintain software compliance
4. **Documentation**: Keep an updated inventory of approved software for your organization

"@

    return $recommendations
}

# Main report generation
try {
    Write-Verbose "Generating Markdown report..."
    
    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $reportTitle = "Application Comparison Report: $($Computer1Data.ComputerName) vs $($Computer2Data.ComputerName)"
    
    $markdown = @"
# $reportTitle

**Generated:** $reportDate  
**Comparison Method:** PowerShell Application Comparison Toolkit  

---

"@

    # Add executive summary
    $markdown += New-ExecutiveSummary -ComparisonData $ComparisonData -Computer1Data $Computer1Data -Computer2Data $Computer2Data

    # Add comparison overview
    $markdown += @"
## Comparison Overview

| Metric | $($Computer1Data.ComputerName) | $($Computer2Data.ComputerName) |
|--------|------------|------------|
| Total Applications | $($ComparisonData.TotalAppsComputer1) | $($ComparisonData.TotalAppsComputer2) |
| Export Date | $($Computer1Data.ExportDate) | $($Computer2Data.ExportDate) |
| System Components Included | $($Computer1Data.IncludeSystemComponents) | $($Computer2Data.IncludeSystemComponents) |
| Updates Included | $($Computer1Data.IncludeUpdates) | $($Computer2Data.IncludeUpdates) |

### Difference Summary

| Category | Count | Description |
|----------|-------|-------------|
| Only on $($Computer1Data.ComputerName) | **$($ComparisonData.OnlyInComputer1.Count)** | Applications missing from $($Computer2Data.ComputerName) |
| Only on $($Computer2Data.ComputerName) | **$($ComparisonData.OnlyInComputer2.Count)** | Applications missing from $($Computer1Data.ComputerName) |
| Version Differences | **$($ComparisonData.VersionDifferences.Count)** | Same applications with different versions |
| Common Applications | **$($ComparisonData.CommonApps.Count)** | Identical applications on both computers |

"@

    # Add detailed tables
    if ($ComparisonData.OnlyInComputer1.Count -gt 0) {
        $markdown += "## Applications Only on $($Computer1Data.ComputerName)`n"
        $markdown += "*(Missing from $($Computer2Data.ComputerName))*`n"
        $markdown += Format-ApplicationTable -Applications $ComparisonData.OnlyInComputer1 -TableTitle "Missing Applications"
    }

    if ($ComparisonData.OnlyInComputer2.Count -gt 0) {
        $markdown += "## Applications Only on $($Computer2Data.ComputerName)`n"
        $markdown += "*(Missing from $($Computer1Data.ComputerName))*`n"
        $markdown += Format-ApplicationTable -Applications $ComparisonData.OnlyInComputer2 -TableTitle "Missing Applications"
    }

    if ($ComparisonData.VersionDifferences.Count -gt 0) {
        $markdown += "## Version Differences`n"
        $markdown += Format-VersionDifferenceTable -VersionDifferences $ComparisonData.VersionDifferences
    }

    if ($IncludeDetailedTables -and $ComparisonData.CommonApps.Count -gt 0) {
        $markdown += "## Common Applications`n"
        $markdown += "*(Identical on both computers)*`n"
        $markdown += Format-ApplicationTable -Applications $ComparisonData.CommonApps -TableTitle "Common Applications"
    }

    # Add statistics if requested
    if ($IncludeStatistics) {
        $markdown += New-StatisticsSection -ComparisonData $ComparisonData -Computer1Data $Computer1Data -Computer2Data $Computer2Data
    }

    # Add recommendations
    $markdown += New-RecommendationsSection -ComparisonData $ComparisonData -Computer1Data $Computer1Data -Computer2Data $Computer2Data

    # Add appendix
    $markdown += @"
---

## Appendix

### Report Metadata

- **Report Type:** Application Comparison Analysis
- **Generated By:** PowerShell Application Comparison Toolkit
- **Report Format:** Markdown
- **Total Applications Analyzed:** $($ComparisonData.TotalAppsComputer1 + $ComparisonData.TotalAppsComputer2)
- **Analysis Date:** $reportDate

### Data Sources

This report was generated by analyzing installed applications from multiple sources:

- **Windows Registry**: Programs and Features entries
- **Package Managers**: Get-Package, Windows Package Manager
- **AppX Packages**: Windows Store applications
- **WMI**: Windows Management Instrumentation (when available)

### Legend

- 🔴 **HIGH RISK**: Requires immediate attention
- 🟡 **MEDIUM RISK**: Should be reviewed and addressed
- 🟢 **LOW RISK**: Minimal differences detected
- ✅ **GOOD**: No issues found
- ⚠️ **WARNING**: Potential concern identified

---

*This report was automatically generated by the PowerShell Application Comparison Toolkit. For questions or support, please refer to the toolkit documentation.*

"@

    # Write the markdown file
    $markdown | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Verbose "Markdown report generated successfully: $OutputPath"
    return $true
}
catch {
    Write-Error "Error generating Markdown report: $($_.Exception.Message)"
    return $false
}
