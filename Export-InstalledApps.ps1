# Export-InstalledApps.ps1
# PowerShell script to export installed applications from a Windows computer
# Supports multiple sources: Registry, Get-WmiObject, and Get-Package

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\InstalledApps_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemComponents,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeUpdates,
    
    [Parameter(Mandatory = $false)]
    [switch]$Verbose
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

function Get-InstalledAppsFromRegistry {
    Write-Log "Collecting applications from Registry..."
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $apps = @()
    
    foreach ($path in $registryPaths) {
        try {
            $items = Get-ItemProperty $path -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.DisplayName) {
                    # Skip system components unless explicitly requested
                    if (-not $IncludeSystemComponents -and (
                        $item.DisplayName -match "Microsoft Visual C\+\+" -or
                        $item.DisplayName -match "Microsoft .NET" -or
                        $item.DisplayName -match "Windows Software Development Kit" -or
                        $item.SystemComponent -eq 1
                    )) {
                        continue
                    }
                    
                    # Skip updates unless explicitly requested
                    if (-not $IncludeUpdates -and (
                        $item.DisplayName -match "Update for" -or
                        $item.DisplayName -match "Hotfix for" -or
                        $item.DisplayName -match "Security Update"
                    )) {
                        continue
                    }
                    
                    $app = [PSCustomObject]@{
                        Name = $item.DisplayName
                        Version = $item.DisplayVersion
                        Publisher = $item.Publisher
                        InstallDate = $item.InstallDate
                        UninstallString = $item.UninstallString
                        InstallLocation = $item.InstallLocation
                        Size = $item.EstimatedSize
                        Source = "Registry"
                        Architecture = if ($path -match "WOW6432Node") { "x86" } else { "x64" }
                    }
                    $apps += $app
                }
            }
        }
        catch {
            Write-Log "Error reading registry path $path`: $($_.Exception.Message)" "ERROR"
        }
    }
    
    return $apps
}

function Get-InstalledAppsFromWMI {
    Write-Log "Collecting applications from WMI..."
    
    try {
        $wmiApps = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue
        $apps = @()
        
        foreach ($app in $wmiApps) {
            $appObj = [PSCustomObject]@{
                Name = $app.Name
                Version = $app.Version
                Publisher = $app.Vendor
                InstallDate = $app.InstallDate
                UninstallString = $null
                InstallLocation = $app.InstallLocation
                Size = $null
                Source = "WMI"
                Architecture = "Unknown"
            }
            $apps += $appObj
        }
        
        return $apps
    }
    catch {
        Write-Log "Error reading from WMI: $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Get-InstalledAppsFromPackageManager {
    Write-Log "Collecting applications from Package Managers..."
    
    $apps = @()
    
    # Get-Package (includes MSI, Programs and Features, etc.)
    try {
        $packages = Get-Package -ErrorAction SilentlyContinue
        foreach ($package in $packages) {
            $app = [PSCustomObject]@{
                Name = $package.Name
                Version = $package.Version
                Publisher = $package.Source
                InstallDate = $null
                UninstallString = $null
                InstallLocation = $null
                Size = $null
                Source = "PackageManager"
                Architecture = "Unknown"
            }
            $apps += $app
        }
    }
    catch {
        Write-Log "Error reading from Package Manager: $($_.Exception.Message)" "WARN"
    }
    
    # Windows Store Apps (AppX packages)
    try {
        $appxPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        foreach ($package in $appxPackages) {
            $app = [PSCustomObject]@{
                Name = $package.Name
                Version = $package.Version
                Publisher = $package.Publisher
                InstallDate = $null
                UninstallString = $null
                InstallLocation = $package.InstallLocation
                Size = $null
                Source = "AppX"
                Architecture = $package.Architecture
            }
            $apps += $app
        }
    }
    catch {
        Write-Log "Error reading AppX packages: $($_.Exception.Message)" "WARN"
    }
    
    return $apps
}

function Merge-ApplicationLists {
    param([array]$Lists)
    
    Write-Log "Merging and deduplicating application lists..."
    
    $mergedApps = @()
    $uniqueApps = @{}
    
    foreach ($list in $Lists) {
        foreach ($app in $list) {
            # Create a unique key based on name and version
            $key = "$($app.Name)_$($app.Version)".ToLower()
            
            if (-not $uniqueApps.ContainsKey($key)) {
                $uniqueApps[$key] = $app
            } else {
                # If we have a duplicate, prefer the one with more information
                $existing = $uniqueApps[$key]
                if (($app.Publisher -and -not $existing.Publisher) -or
                    ($app.InstallLocation -and -not $existing.InstallLocation) -or
                    ($app.Size -and -not $existing.Size)) {
                    $uniqueApps[$key] = $app
                }
            }
        }
    }
    
    return $uniqueApps.Values | Sort-Object Name
}

# Main execution
try {
    Write-Log "Starting application export for computer: $ComputerName" "SUCCESS"
    Write-Log "Output file: $OutputPath"
    
    # Collect from multiple sources
    $allApps = @()
    
    $registryApps = Get-InstalledAppsFromRegistry
    $wmiApps = Get-InstalledAppsFromWMI
    $packageApps = Get-InstalledAppsFromPackageManager
    
    Write-Log "Found $($registryApps.Count) apps from Registry"
    Write-Log "Found $($wmiApps.Count) apps from WMI"
    Write-Log "Found $($packageApps.Count) apps from Package Managers"
    
    # Merge all lists
    $finalApps = Merge-ApplicationLists -Lists @($registryApps, $wmiApps, $packageApps)
    
    # Create export object
    $exportData = [PSCustomObject]@{
        ComputerName = $ComputerName
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalApplications = $finalApps.Count
        IncludeSystemComponents = $IncludeSystemComponents.IsPresent
        IncludeUpdates = $IncludeUpdates.IsPresent
        Applications = $finalApps
    }
    
    # Export to JSON
    $exportData | ConvertTo-Json -Depth 3 | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "Successfully exported $($finalApps.Count) applications to: $OutputPath" "SUCCESS"
    Write-Log "Export completed for computer: $ComputerName" "SUCCESS"
    
    # Display summary
    Write-Host "`nSUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer: $ComputerName" -ForegroundColor White
    Write-Host "Total Applications: $($finalApps.Count)" -ForegroundColor White
    Write-Host "Output File: $OutputPath" -ForegroundColor White
    
}
catch {
    Write-Log "Critical error during export: $($_.Exception.Message)" "ERROR"
    exit 1
}
