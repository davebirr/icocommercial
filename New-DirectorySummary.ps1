# New-DirectorySummary.ps1
# Creates a comprehensive summary report for an output directory containing comparison reports

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DirectoryPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFileName = "SUMMARY_REPORT.md"
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

try {
    if (-not (Test-Path $DirectoryPath)) {
        throw "Directory not found: $DirectoryPath"
    }
    
    Write-Log "Generating directory summary for: $DirectoryPath"
    
    # Get all files in the directory
    $allFiles = Get-ChildItem -Path $DirectoryPath -File
    $jsonFiles = $allFiles | Where-Object { $_.Extension -eq ".json" }
    $htmlFiles = $allFiles | Where-Object { $_.Extension -eq ".html" }
    $mdFiles = $allFiles | Where-Object { $_.Extension -eq ".md" -and $_.Name -ne $OutputFileName }
    
    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $directoryName = Split-Path $DirectoryPath -Leaf
    
    $summary = @"
# Application Comparison Summary Report

**Directory:** `$DirectoryPath`  
**Generated:** $reportDate  
**Total Files:** $($allFiles.Count)  

---

## Overview

This directory contains the results of application comparison analysis performed using the PowerShell Application Comparison Toolkit.

### File Summary

| File Type | Count | Description |
|-----------|-------|-------------|
| JSON Exports | $($jsonFiles.Count) | Raw application data from computers |
| HTML Reports | $($htmlFiles.Count) | Interactive comparison reports |
| Markdown Reports | $($mdFiles.Count) | Text-based comparison reports |
| Other Files | $($allFiles.Count - $jsonFiles.Count - $htmlFiles.Count - $mdFiles.Count) | Additional files |

"@

    if ($jsonFiles.Count -gt 0) {
        $summary += @"
## Exported Computer Data

The following computers have been analyzed:

| Computer Name | Export Date | File Size | Applications Count |
|---------------|-------------|-----------|-------------------|
"@

        foreach ($jsonFile in $jsonFiles | Sort-Object Name) {
            try {
                $jsonContent = Get-Content $jsonFile.FullName -Raw | ConvertFrom-Json
                $computerName = $jsonContent.ComputerName
                $exportDate = $jsonContent.ExportDate
                $appCount = $jsonContent.TotalApplications
                $fileSize = [math]::Round($jsonFile.Length / 1KB, 2)
                
                $summary += "| $computerName | $exportDate | $fileSize KB | $appCount |`n"
            }
            catch {
                $summary += "| $($jsonFile.BaseName) | *Error reading file* | $([math]::Round($jsonFile.Length / 1KB, 2)) KB | *Unknown* |`n"
            }
        }
        $summary += "`n"
    }

    if ($htmlFiles.Count -gt 0) {
        $summary += @"
## HTML Comparison Reports

Interactive reports with visual formatting and charts:

"@

        foreach ($htmlFile in $htmlFiles | Sort-Object Name) {
            $fileSize = [math]::Round($htmlFile.Length / 1KB, 2)
            $summary += "- **[$($htmlFile.Name)](./$($htmlFile.Name))** ($fileSize KB) - *Created: $($htmlFile.CreationTime.ToString("yyyy-MM-dd HH:mm"))*`n"
        }
        $summary += "`n"
    }

    if ($mdFiles.Count -gt 0) {
        $summary += @"
## Markdown Comparison Reports

Text-based reports suitable for documentation and version control:

"@

        foreach ($mdFile in $mdFiles | Sort-Object Name) {
            $fileSize = [math]::Round($mdFile.Length / 1KB, 2)
            $summary += "- **[$($mdFile.Name)](./$($mdFile.Name))** ($fileSize KB) - *Created: $($mdFile.CreationTime.ToString("yyyy-MM-dd HH:mm"))*`n"
        }
        $summary += "`n"
    }

    # Analysis section
    if ($jsonFiles.Count -ge 2) {
        $summary += @"
## Analysis Summary

### Computers Analyzed

"@

        $computerNames = @()
        foreach ($jsonFile in $jsonFiles) {
            try {
                $jsonContent = Get-Content $jsonFile.FullName -Raw | ConvertFrom-Json
                $computerNames += $jsonContent.ComputerName
            }
            catch {
                $computerNames += $jsonFile.BaseName -replace "Apps_", "" -replace "_\d{8}_\d{6}", ""
            }
        }

        $uniqueComputers = $computerNames | Sort-Object -Unique
        foreach ($computer in $uniqueComputers) {
            $summary += "- **$computer**`n"
        }

        $summary += @"

### Comparison Matrix

The following comparisons have been performed:

"@

        if ($htmlFiles.Count -gt 0) {
            foreach ($htmlFile in $htmlFiles | Sort-Object Name) {
                $fileName = $htmlFile.BaseName
                if ($fileName -match "Comparison_(.+)_vs_(.+)") {
                    $comp1 = $matches[1] -replace "Apps_", "" -replace "_\d{8}_\d{6}", ""
                    $comp2 = $matches[2] -replace "Apps_", "" -replace "_\d{8}_\d{6}", ""
                    $summary += "- **$comp1** vs **$comp2** → [HTML Report](./$($htmlFile.Name))"
                    
                    # Check for corresponding Markdown file
                    $mdFileName = $fileName + ".md"
                    if ($mdFiles | Where-Object { $_.Name -eq $mdFileName }) {
                        $summary += " | [Markdown Report](./$mdFileName)"
                    }
                    $summary += "`n"
                }
            }
        }
        $summary += "`n"
    }

    # Recommendations section
    $summary += @"
## Recommendations

### For IT Administrators

1. **Review HTML Reports**: Open the HTML files in a web browser for the best visual experience
2. **Archive Documentation**: Use Markdown reports for long-term documentation and version control
3. **Regular Comparisons**: Schedule regular comparisons to maintain software consistency
4. **Action Items**: Review applications that appear only on one computer

### For Management

1. **Software Compliance**: Ensure all computers have required business applications
2. **License Management**: Identify unused software that may be consuming licenses
3. **Security Updates**: Pay attention to version differences for security-critical applications
4. **Cost Optimization**: Consider standardizing software across the organization

### Next Steps

- [ ] Review all comparison reports
- [ ] Identify critical software differences
- [ ] Plan software standardization activities
- [ ] Update software deployment procedures
- [ ] Schedule follow-up comparisons

"@

    # Technical details
    $summary += @"
---

## Technical Details

### File Locations

All files are stored in: ``$DirectoryPath``

### Toolkit Information

- **Generator**: PowerShell Application Comparison Toolkit
- **Report Version**: 1.0
- **Analysis Date**: $reportDate
- **Working Directory**: $directoryName

### File Descriptions

- **JSON Files**: Raw application data exported from each computer
- **HTML Files**: Interactive comparison reports with visual charts and color coding
- **Markdown Files**: Text-based reports suitable for documentation and sharing
- **This Summary**: Overview of all analysis performed in this directory

### Support

For questions about this analysis or the toolkit:

1. Review the README.md file in the toolkit directory
2. Check the Example-Usage.ps1 file for usage examples
3. Refer to the PowerShell Application Comparison Toolkit documentation

---

*This summary was automatically generated by the PowerShell Application Comparison Toolkit on $reportDate*

"@

    # Write the summary file
    $outputPath = Join-Path $DirectoryPath $OutputFileName
    $summary | Out-File -FilePath $outputPath -Encoding UTF8
    
    Write-Log "Directory summary created: $outputPath" "SUCCESS"
    
    # Display key statistics
    Write-Host "`nDIRECTORY SUMMARY:" -ForegroundColor Cyan
    Write-Host "Total Files: $($allFiles.Count)" -ForegroundColor White
    Write-Host "JSON Exports: $($jsonFiles.Count)" -ForegroundColor White
    Write-Host "HTML Reports: $($htmlFiles.Count)" -ForegroundColor White
    Write-Host "Markdown Reports: $($mdFiles.Count)" -ForegroundColor White
    Write-Host "Summary Report: $outputPath" -ForegroundColor Green
    
    return $outputPath
}
catch {
    Write-Log "Error generating directory summary: $($_.Exception.Message)" "ERROR"
    return $null
}
