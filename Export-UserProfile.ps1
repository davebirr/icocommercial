# Export-UserProfile.ps1
# Comprehensive script to export user profile settings for computer migration
# Includes applications, Office settings, printers, WiFi, and personalization

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = ".\UserProfile_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeOfficeSettings,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludePrinterSettings,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeWiFiProfiles,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludePersonalization,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludePasswords,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemComponents,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeUpdates,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateReport
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

function Test-ScriptExists {
    param([string]$ScriptName)
    
    $scriptPath = Join-Path $PSScriptRoot $ScriptName
    if (Test-Path $scriptPath) {
        return $scriptPath
    } else {
        Write-Log "Script not found: $ScriptName" "WARN"
        return $null
    }
}

function Invoke-ExportScript {
    param(
        [string]$ScriptPath,
        [string]$OutputPath,
        [hashtable]$Parameters = @{}
    )
    
    try {
        $scriptName = Split-Path $ScriptPath -Leaf
        Write-Log "Running $scriptName..."
        
        # Build parameter string
        $paramString = "-OutputPath `"$OutputPath`""
        foreach ($key in $Parameters.Keys) {
            if ($Parameters[$key] -is [bool] -and $Parameters[$key]) {
                $paramString += " -$key"
            } elseif ($Parameters[$key] -and $Parameters[$key] -isnot [bool]) {
                $paramString += " -$key `"$($Parameters[$key])`""
            }
        }
        
        # Execute the script
        $command = "& `"$ScriptPath`" $paramString"
        Invoke-Expression $command
        
        if (Test-Path $OutputPath) {
            Write-Log "$scriptName completed successfully" "SUCCESS"
            return $true
        } else {
            Write-Log "$scriptName failed to create output file" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error running $scriptName`: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function New-ProfileSummaryReport {
    param(
        [string]$OutputDirectory,
        [hashtable]$ExportResults
    )
    
    Write-Log "Generating user profile summary report..."
    
    $reportPath = Join-Path $OutputDirectory "UserProfile_Summary.md"
    
    $report = @"
# User Profile Export Summary

**Computer:** $ComputerName  
**Export Date:** $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")  
**Export Directory:** $OutputDirectory

## Export Results

"@

    foreach ($category in $ExportResults.Keys) {
        $result = $ExportResults[$category]
        $status = if ($result.Success) { "✅ Success" } else { "❌ Failed" }
        
        $report += @"

### $category
**Status:** $status  
**Output File:** $($result.OutputFile)  

"@
        
        if ($result.Success -and (Test-Path $result.OutputFile)) {
            try {
                $fileInfo = Get-Item $result.OutputFile
                $report += "**File Size:** $([math]::Round($fileInfo.Length / 1KB, 2)) KB  `n"
                $report += "**Created:** $($fileInfo.CreationTime)  `n"
                
                # Try to get summary info from JSON files
                if ($fileInfo.Extension -eq ".json") {
                    try {
                        $jsonContent = Get-Content $result.OutputFile -Raw | ConvertFrom-Json
                        
                        switch ($category) {
                            "Applications" {
                                if ($jsonContent.TotalApplications) {
                                    $report += "**Applications Found:** $($jsonContent.TotalApplications)  `n"
                                }
                            }
                            "Office Settings" {
                                if ($jsonContent.OfficeVersions) {
                                    $report += "**Office Versions:** $($jsonContent.OfficeVersions.Count)  `n"
                                }
                            }
                            "Printer Settings" {
                                if ($jsonContent.TotalPrinters) {
                                    $report += "**Printers:** $($jsonContent.TotalPrinters)  `n"
                                    if ($jsonContent.DefaultPrinter) {
                                        $report += "**Default Printer:** $($jsonContent.DefaultPrinter.Name)  `n"
                                    }
                                }
                            }
                            "WiFi Profiles" {
                                if ($jsonContent.TotalProfiles) {
                                    $report += "**WiFi Profiles:** $($jsonContent.TotalProfiles)  `n"
                                    if ($jsonContent.ProfilesWithPasswords) {
                                        $report += "**Profiles with Passwords:** $($jsonContent.ProfilesWithPasswords)  `n"
                                    }
                                }
                            }
                            "Windows Personalization" {
                                if ($jsonContent.WindowsVersion) {
                                    $report += "**Windows Version:** $($jsonContent.WindowsVersion)  `n"
                                }
                            }
                        }
                    }
                    catch {
                        Write-Log "Could not parse JSON for summary: $($_.Exception.Message)" "WARN"
                    }
                }
            }
            catch {
                Write-Log "Could not get file info for summary: $($_.Exception.Message)" "WARN"
            }
        }
    }
    
    $report += @"

## Next Steps

1. **Review Export Files:** Check each exported file to ensure all necessary data was captured
2. **Secure Storage:** Store exported files in a secure location, especially if passwords are included
3. **Test Import:** Verify that settings can be properly imported on the target computer
4. **Clean Up:** Remove temporary files and sensitive data when migration is complete

## Import Notes

- **Applications:** Use the comparison tools to identify which applications need to be installed
- **Office Settings:** May require manual configuration depending on Office version differences
- **Printers:** Network printers may need to be reconnected manually
- **WiFi Profiles:** Can be imported using `netsh wlan add profile` commands
- **Personalization:** Registry settings may need to be imported carefully to avoid conflicts

## Support

For assistance with this user profile migration toolkit, refer to the documentation in the repository.
"@

    try {
        $report | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Log "Summary report created: $reportPath" "SUCCESS"
        return $reportPath
    }
    catch {
        Write-Log "Failed to create summary report: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Main execution
try {
    Write-Log "Starting comprehensive user profile export for: $ComputerName" "SUCCESS"
    
    # Create output directory
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Log "Created output directory: $OutputDirectory"
    }
    
    $exportResults = @{}
    
    # 1. Export Applications (always included)
    Write-Log "Exporting installed applications..." "SUCCESS"
    $appsScript = Test-ScriptExists "Export-InstalledApps.ps1"
    if ($appsScript) {
        $appsOutput = Join-Path $OutputDirectory "InstalledApps.json"
        $appsParams = @{
            ComputerName = $ComputerName
        }
        if ($IncludeSystemComponents) { $appsParams["IncludeSystemComponents"] = $true }
        if ($IncludeUpdates) { $appsParams["IncludeUpdates"] = $true }
        $appsSuccess = Invoke-ExportScript -ScriptPath $appsScript -OutputPath $appsOutput -Parameters $appsParams
        $exportResults["Applications"] = @{ Success = $appsSuccess; OutputFile = $appsOutput }
    }
    
    # 2. Export Office Settings (if requested)
    if ($IncludeOfficeSettings) {
        Write-Log "Exporting Office settings..." "SUCCESS"
        $officeScript = Test-ScriptExists "Export-OfficeSettings.ps1"
        if ($officeScript) {
            $officeOutput = Join-Path $OutputDirectory "OfficeSettings.json"
        $officeParams = @{
            ComputerName = $ComputerName
        }
        if ($true) { $officeParams["IncludeTemplates"] = $true }
        if ($true) { $officeParams["IncludeSignatures"] = $true }
            $officeSuccess = Invoke-ExportScript -ScriptPath $officeScript -OutputPath $officeOutput -Parameters $officeParams
            $exportResults["Office Settings"] = @{ Success = $officeSuccess; OutputFile = $officeOutput }
        }
    }
    
    # 3. Export Printer Settings (if requested)
    if ($IncludePrinterSettings) {
        Write-Log "Exporting printer settings..." "SUCCESS"
        $printerScript = Test-ScriptExists "Export-PrinterSettings.ps1"
        if ($printerScript) {
            $printerOutput = Join-Path $OutputDirectory "PrinterSettings.json"
        $printerParams = @{
            ComputerName = $ComputerName
        }
        if ($true) { $printerParams["IncludeDriverDetails"] = $true }
            $printerSuccess = Invoke-ExportScript -ScriptPath $printerScript -OutputPath $printerOutput -Parameters $printerParams
            $exportResults["Printer Settings"] = @{ Success = $printerSuccess; OutputFile = $printerOutput }
        }
    }
    
    # 4. Export WiFi Profiles (if requested)
    if ($IncludeWiFiProfiles) {
        Write-Log "Exporting WiFi profiles..." "SUCCESS"
        $wifiScript = Test-ScriptExists "Export-WiFiProfiles.ps1"
        if ($wifiScript) {
            $wifiOutput = Join-Path $OutputDirectory "WiFiProfiles.json"
        $wifiParams = @{
            ComputerName = $ComputerName
        }
        if ($IncludePasswords) { $wifiParams["IncludePasswords"] = $true }
        if ($true) { $wifiParams["ExportProfiles"] = $true }
            $wifiSuccess = Invoke-ExportScript -ScriptPath $wifiScript -OutputPath $wifiOutput -Parameters $wifiParams
            $exportResults["WiFi Profiles"] = @{ Success = $wifiSuccess; OutputFile = $wifiOutput }
        }
    }
    
    # 5. Export Windows Personalization (if requested)
    if ($IncludePersonalization) {
        Write-Log "Exporting Windows personalization..." "SUCCESS"
        $personalizationScript = Test-ScriptExists "Export-WindowsPersonalization.ps1"
        if ($personalizationScript) {
            $personalizationOutput = Join-Path $OutputDirectory "WindowsPersonalization.json"
        $personalizationParams = @{
            ComputerName = $ComputerName
        }
        if ($true) { $personalizationParams["IncludeWallpaper"] = $true }
        if ($true) { $personalizationParams["IncludeStartLayout"] = $true }
            $personalizationSuccess = Invoke-ExportScript -ScriptPath $personalizationScript -OutputPath $personalizationOutput -Parameters $personalizationParams
            $exportResults["Windows Personalization"] = @{ Success = $personalizationSuccess; OutputFile = $personalizationOutput }
        }
    }
    
    # 6. Generate Summary Report (if requested)
    if ($GenerateReport) {
        $summaryReport = New-ProfileSummaryReport -OutputDirectory $OutputDirectory -ExportResults $exportResults
    }
    
    # Display final summary
    Write-Host "`n" -NoNewline
    Write-Log "User profile export completed!" "SUCCESS"
    
    Write-Host "`nFINAL SUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer: $ComputerName" -ForegroundColor White
    Write-Host "Output Directory: $OutputDirectory" -ForegroundColor White
    Write-Host "Exports Completed:" -ForegroundColor White
    
    $successCount = 0
    foreach ($category in $exportResults.Keys) {
        $result = $exportResults[$category]
        $status = if ($result.Success) { 
            $successCount++
            "✅" 
        } else { 
            "❌" 
        }
        Write-Host "  $status $category" -ForegroundColor White
    }
    
    Write-Host "`nSuccess Rate: $successCount/$($exportResults.Count)" -ForegroundColor $(if ($successCount -eq $exportResults.Count) { "Green" } else { "Yellow" })
    
    if ($GenerateReport -and $summaryReport) {
        Write-Host "Summary Report: $summaryReport" -ForegroundColor White
    }
    
}
catch {
    Write-Log "Critical error during user profile export: $($_.Exception.Message)" "ERROR"
    exit 1
}
