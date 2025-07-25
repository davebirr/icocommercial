# Export-OfficeSettings.ps1
# PowerShell script to export Microsoft Office settings and customizations
# Supports M365/Office 365, Office 2019, Office 2016

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\OfficeSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTemplates,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSignatures
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

function Get-OfficeVersions {
    Write-Log "Detecting installed Office versions..."
    
    $officeVersions = @()
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
    )
    
    foreach ($path in $registryPaths) {
        try {
            if (Test-Path $path) {
                $versions = Get-ChildItem $path -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\d+\.\d+' }
                foreach ($version in $versions) {
                    $versionNumber = Split-Path $version.Name -Leaf
                    $clickToRunPath = "$path\$versionNumber\ClickToRun"
                    $msiPath = "$path\$versionNumber\Common\InstallRoot"
                    
                    $installType = "Unknown"
                    $installPath = ""
                    
                    if (Test-Path $clickToRunPath) {
                        $installType = "Click-to-Run"
                        $installPath = (Get-ItemProperty $clickToRunPath -Name "InstallPath" -ErrorAction SilentlyContinue).InstallPath
                    } elseif (Test-Path $msiPath) {
                        $installType = "MSI"
                        $installPath = (Get-ItemProperty $msiPath -Name "Path" -ErrorAction SilentlyContinue).Path
                    }
                    
                    if ($installPath) {
                        $officeVersions += [PSCustomObject]@{
                            Version = $versionNumber
                            InstallType = $installType
                            InstallPath = $installPath
                            RegistryPath = "$path\$versionNumber"
                        }
                    }
                }
            }
        }
        catch {
            Write-Log "Error reading Office registry path $path`: $($_.Exception.Message)" "WARN"
        }
    }
    
    return $officeVersions
}

function Get-OfficeRegistrySettings {
    param([array]$OfficeVersions)
    
    Write-Log "Collecting Office registry settings..."
    
    $registrySettings = @{}
    
    foreach ($version in $OfficeVersions) {
        $versionSettings = @{}
        $userSettingsPath = "HKCU:\Software\Microsoft\Office\$($version.Version)"
        
        if (Test-Path $userSettingsPath) {
            # Common Office applications
            $applications = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "Project", "Visio")
            
            foreach ($app in $applications) {
                $appPath = "$userSettingsPath\$app"
                if (Test-Path $appPath) {
                    Write-Log "Collecting $app settings for Office $($version.Version)..."
                    
                    $appSettings = @{}
                    
                    # Get Options
                    $optionsPath = "$appPath\Options"
                    if (Test-Path $optionsPath) {
                        try {
                            $options = Get-ItemProperty $optionsPath -ErrorAction SilentlyContinue
                            $appSettings["Options"] = @{}
                            $options.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                                $appSettings["Options"][$_.Name] = $_.Value
                            }
                        }
                        catch {
                            Write-Log "Error reading options for $app`: $($_.Exception.Message)" "WARN"
                        }
                    }
                    
                    # Get User Templates (if requested)
                    if ($IncludeTemplates) {
                        $userTemplatePath = "$appPath\Options"
                        if (Test-Path $userTemplatePath) {
                            try {
                                $templateSettings = Get-ItemProperty $userTemplatePath -Name "*Template*" -ErrorAction SilentlyContinue
                                if ($templateSettings) {
                                    $appSettings["Templates"] = @{}
                                    $templateSettings.PSObject.Properties | Where-Object { $_.Name -match "Template" -and $_.Name -notmatch "^PS" } | ForEach-Object {
                                        $appSettings["Templates"][$_.Name] = $_.Value
                                    }
                                }
                            }
                            catch {
                                Write-Log "Error reading template settings for $app`: $($_.Exception.Message)" "WARN"
                            }
                        }
                    }
                    
                    # Get Preferences
                    $preferencesPath = "$appPath\Preferences"
                    if (Test-Path $preferencesPath) {
                        try {
                            $preferences = Get-ItemProperty $preferencesPath -ErrorAction SilentlyContinue
                            $appSettings["Preferences"] = @{}
                            $preferences.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                                $appSettings["Preferences"][$_.Name] = $_.Value
                            }
                        }
                        catch {
                            Write-Log "Error reading preferences for $app`: $($_.Exception.Message)" "WARN"
                        }
                    }
                    
                    if ($appSettings.Count -gt 0) {
                        $versionSettings[$app] = $appSettings
                    }
                }
            }
            
            # Get Common Office Settings
            $commonPath = "$userSettingsPath\Common"
            if (Test-Path $commonPath) {
                $commonSettings = @{}
                
                # General settings
                $generalPath = "$commonPath\General"
                if (Test-Path $generalPath) {
                    try {
                        $general = Get-ItemProperty $generalPath -ErrorAction SilentlyContinue
                        $commonSettings["General"] = @{}
                        $general.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                            $commonSettings["General"][$_.Name] = $_.Value
                        }
                    }
                    catch {
                        Write-Log "Error reading common general settings: $($_.Exception.Message)" "WARN"
                    }
                }
                
                if ($commonSettings.Count -gt 0) {
                    $versionSettings["Common"] = $commonSettings
                }
            }
        }
        
        if ($versionSettings.Count -gt 0) {
            $registrySettings["Office_$($version.Version)"] = $versionSettings
        }
    }
    
    return $registrySettings
}

function Get-OutlookSignatures {
    Write-Log "Collecting Outlook email signatures..."
    
    $signatures = @{}
    $signaturePath = "$env:APPDATA\Microsoft\Signatures"
    
    if (Test-Path $signaturePath) {
        try {
            $signatureFiles = Get-ChildItem $signaturePath -File -ErrorAction SilentlyContinue
            
            # Group signature files by base name
            $signatureGroups = $signatureFiles | Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -replace '_files$', '' }
            
            foreach ($group in $signatureGroups) {
                $signatureName = $group.Name
                $signatureData = @{
                    Name = $signatureName
                    Files = @()
                }
                
                foreach ($file in $group.Group) {
                    $fileInfo = @{
                        FileName = $file.Name
                        Extension = $file.Extension
                        Size = $file.Length
                        LastModified = $file.LastWriteTime
                    }
                    
                    # For text and HTML files, include content
                    if ($file.Extension -in @('.txt', '.htm', '.html')) {
                        try {
                            $content = Get-Content $file.FullName -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                            if ($content) {
                                $fileInfo["Content"] = $content
                            }
                        }
                        catch {
                            Write-Log "Could not read signature file $($file.Name): $($_.Exception.Message)" "WARN"
                        }
                    }
                    
                    $signatureData.Files += $fileInfo
                }
                
                $signatures[$signatureName] = $signatureData
            }
        }
        catch {
            Write-Log "Error reading signatures: $($_.Exception.Message)" "WARN"
        }
    } else {
        Write-Log "Outlook signatures folder not found" "WARN"
    }
    
    return $signatures
}

function Get-OfficeTemplates {
    Write-Log "Collecting Office templates..."
    
    $templates = @{}
    $templatePaths = @(
        "$env:APPDATA\Microsoft\Templates",
        "$env:USERPROFILE\Documents\Custom Office Templates",
        "$env:PROGRAMFILES\Microsoft Office\Templates",
        "$env:PROGRAMFILES(X86)\Microsoft Office\Templates"
    )
    
    foreach ($templatePath in $templatePaths) {
        if (Test-Path $templatePath) {
            try {
                $templateFiles = Get-ChildItem $templatePath -File -Recurse -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Extension -in @('.dotx', '.dotm', '.potx', '.potm', '.xltx', '.xltm') }
                
                $pathTemplates = @()
                foreach ($template in $templateFiles) {
                    $templateInfo = @{
                        Name = $template.Name
                        FullPath = $template.FullName
                        RelativePath = $template.FullName.Replace($templatePath, "").TrimStart('\')
                        Extension = $template.Extension
                        Size = $template.Length
                        LastModified = $template.LastWriteTime
                        Created = $template.CreationTime
                    }
                    $pathTemplates += $templateInfo
                }
                
                if ($pathTemplates.Count -gt 0) {
                    $templates[$templatePath] = $pathTemplates
                }
            }
            catch {
                Write-Log "Error reading templates from $templatePath`: $($_.Exception.Message)" "WARN"
            }
        }
    }
    
    return $templates
}

function Get-OfficeAddIns {
    Write-Log "Collecting Office Add-ins..."
    
    $addins = @{}
    $officeVersions = Get-OfficeVersions
    
    foreach ($version in $officeVersions) {
        $versionAddins = @{}
        $applications = @("Word", "Excel", "PowerPoint", "Outlook")
        
        foreach ($app in $applications) {
            $addinsPath = "HKCU:\Software\Microsoft\Office\$($version.Version)\$app\Addins"
            if (Test-Path $addinsPath) {
                try {
                    $appAddins = @()
                    $addinKeys = Get-ChildItem $addinsPath -ErrorAction SilentlyContinue
                    
                    foreach ($addin in $addinKeys) {
                        $addinProps = Get-ItemProperty $addin.PSPath -ErrorAction SilentlyContinue
                        if ($addinProps) {
                            $addinInfo = @{
                                Name = Split-Path $addin.Name -Leaf
                                LoadBehavior = $addinProps.LoadBehavior
                                Description = $addinProps.Description
                                FriendlyName = $addinProps.FriendlyName
                                Path = $addinProps.OPEN
                            }
                            $appAddins += $addinInfo
                        }
                    }
                    
                    if ($appAddins.Count -gt 0) {
                        $versionAddins[$app] = $appAddins
                    }
                }
                catch {
                    Write-Log "Error reading add-ins for $app`: $($_.Exception.Message)" "WARN"
                }
            }
        }
        
        if ($versionAddins.Count -gt 0) {
            $addins["Office_$($version.Version)"] = $versionAddins
        }
    }
    
    return $addins
}

# Main execution
try {
    Write-Log "Starting Office settings export for computer: $ComputerName" "SUCCESS"
    Write-Log "Output file: $OutputPath"
    
    # Detect Office versions
    $officeVersions = Get-OfficeVersions
    
    if ($officeVersions.Count -eq 0) {
        Write-Log "No Office installations found" "WARN"
        $exportData = [PSCustomObject]@{
            ComputerName = $ComputerName
            ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            OfficeVersions = @()
            RegistrySettings = @{}
            Signatures = @{}
            Templates = @{}
            AddIns = @{}
            IncludeTemplates = $IncludeTemplates.IsPresent
            IncludeSignatures = $IncludeSignatures.IsPresent
        }
    } else {
        Write-Log "Found $($officeVersions.Count) Office installation(s)"
        foreach ($version in $officeVersions) {
            Write-Log "  - Office $($version.Version) ($($version.InstallType))"
        }
        
        # Collect settings
        $registrySettings = Get-OfficeRegistrySettings -OfficeVersions $officeVersions
        $addins = Get-OfficeAddIns
        
        # Collect signatures if requested
        $signatures = @{}
        if ($IncludeSignatures) {
            $signatures = Get-OutlookSignatures
        }
        
        # Collect templates if requested
        $templates = @{}
        if ($IncludeTemplates) {
            $templates = Get-OfficeTemplates
        }
        
        # Create export object
        $exportData = [PSCustomObject]@{
            ComputerName = $ComputerName
            ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            OfficeVersions = $officeVersions
            RegistrySettings = $registrySettings
            Signatures = $signatures
            Templates = $templates
            AddIns = $addins
            IncludeTemplates = $IncludeTemplates.IsPresent
            IncludeSignatures = $IncludeSignatures.IsPresent
        }
    }
    
    # Export to JSON
    $exportData | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "Successfully exported Office settings to: $OutputPath" "SUCCESS"
    Write-Log "Export completed for computer: $ComputerName" "SUCCESS"
    
    # Display summary
    Write-Host "`nSUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer: $ComputerName" -ForegroundColor White
    Write-Host "Office Versions Found: $($officeVersions.Count)" -ForegroundColor White
    if ($signatures.Count -gt 0) {
        Write-Host "Email Signatures: $($signatures.Count)" -ForegroundColor White
    }
    if ($templates.Count -gt 0) {
        $templateCount = ($templates.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        Write-Host "Templates Found: $templateCount" -ForegroundColor White
    }
    Write-Host "Output File: $OutputPath" -ForegroundColor White
    
}
catch {
    Write-Log "Critical error during Office settings export: $($_.Exception.Message)" "ERROR"
    exit 1
}
