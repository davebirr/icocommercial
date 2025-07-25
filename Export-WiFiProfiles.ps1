# Export-WiFiProfiles.ps1
# PowerShell script to export WiFi network profiles and credentials
# Includes network names, security settings, and passwords

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\WiFiProfiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludePasswords,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportProfiles
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

function Test-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-WiFiProfiles {
    Write-Log "Collecting WiFi profiles..."
    
    $profiles = @()
    
    try {
        # Get list of WiFi profiles
        $netshOutput = netsh wlan show profiles
        
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Failed to retrieve WiFi profiles. WiFi adapter may not be available." "WARN"
            return @()
        }
        
        # Parse profile names from netsh output
        $profileLines = $netshOutput | Where-Object { $_ -match "All User Profile\s+:\s(.+)" }
        
        foreach ($line in $profileLines) {
            if ($line -match "All User Profile\s+:\s(.+)") {
                $profileName = $matches[1].Trim()
                
                Write-Log "Processing WiFi profile: $profileName"
                
                # Get detailed profile information
                $profileDetails = netsh wlan show profile name="$profileName"
                
                $profileInfo = [PSCustomObject]@{
                    ProfileName = $profileName
                    SSID = $null
                    ConnectionType = $null
                    NetworkType = $null
                    Authentication = $null
                    Encryption = $null
                    UseOneX = $null
                    AutoConnect = $null
                    AutoSwitch = $null
                    MACRandomization = $null
                    Password = $null
                    ProfileXML = $null
                }
                
                # Parse profile details
                foreach ($detailLine in $profileDetails) {
                    switch -Regex ($detailLine.Trim()) {
                        "SSID name\s+:\s(.+)" { $profileInfo.SSID = $matches[1].Trim('"') }
                        "Network type\s+:\s(.+)" { $profileInfo.NetworkType = $matches[1].Trim() }
                        "Radio type\s+:\s(.+)" { $profileInfo.ConnectionType = $matches[1].Trim() }
                        "Authentication\s+:\s(.+)" { $profileInfo.Authentication = $matches[1].Trim() }
                        "Cipher\s+:\s(.+)" { $profileInfo.Encryption = $matches[1].Trim() }
                        "Use 802.1X\s+:\s(.+)" { $profileInfo.UseOneX = $matches[1].Trim() }
                        "AutoConnect\s+:\s(.+)" { $profileInfo.AutoConnect = $matches[1].Trim() }
                        "AutoSwitch\s+:\s(.+)" { $profileInfo.AutoSwitch = $matches[1].Trim() }
                        "MAC Randomization\s+:\s(.+)" { $profileInfo.MACRandomization = $matches[1].Trim() }
                    }
                }
                
                # Get password if requested and user has admin rights
                if ($IncludePasswords) {
                    if (Test-AdminRights) {
                        try {
                            $keyOutput = netsh wlan show profile name="$profileName" key=clear
                            foreach ($keyLine in $keyOutput) {
                                if ($keyLine -match "Key Content\s+:\s(.+)") {
                                    $profileInfo.Password = $matches[1].Trim()
                                    break
                                }
                            }
                        }
                        catch {
                            Write-Log "Could not retrieve password for profile '$profileName': $($_.Exception.Message)" "WARN"
                        }
                    } else {
                        Write-Log "Administrator rights required to export passwords. Run as administrator or use -IncludePasswords switch." "WARN"
                    }
                }
                
                # Export profile XML if requested
                if ($ExportProfiles) {
                    try {
                        $tempPath = [System.IO.Path]::GetTempPath()
                        $xmlFileName = "WiFiProfile_$($profileName -replace '[^\w\-_\.]', '_').xml"
                        $xmlPath = Join-Path $tempPath $xmlFileName
                        
                        $null = netsh wlan export profile name="$profileName" folder="$tempPath" key=clear
                        
                        if (Test-Path $xmlPath) {
                            $xmlContent = Get-Content $xmlPath -Raw -Encoding UTF8
                            $profileInfo.ProfileXML = $xmlContent
                            Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                    catch {
                        Write-Log "Could not export XML for profile '$profileName': $($_.Exception.Message)" "WARN"
                    }
                }
                
                $profiles += $profileInfo
            }
        }
        
        Write-Log "Found $($profiles.Count) WiFi profiles"
        
    }
    catch {
        Write-Log "Error collecting WiFi profiles: $($_.Exception.Message)" "ERROR"
        return @()
    }
    
    return $profiles
}

function Get-WiFiAdapterInfo {
    Write-Log "Collecting WiFi adapter information..."
    
    $adapters = @()
    
    try {
        # Get WiFi adapters using netsh
        $interfaceOutput = netsh wlan show interfaces
        
        if ($LASTEXITCODE -eq 0) {
            $currentAdapter = @{}
            
            foreach ($line in $interfaceOutput) {
                $line = $line.Trim()
                
                if ($line -eq "" -and $currentAdapter.Count -gt 0) {
                    # End of current adapter info
                    $adapters += [PSCustomObject]$currentAdapter
                    $currentAdapter = @{}
                }
                elseif ($line -match "^(.+?)\s*:\s*(.+)$") {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    $currentAdapter[$key] = $value
                }
            }
            
            # Add the last adapter if exists
            if ($currentAdapter.Count -gt 0) {
                $adapters += [PSCustomObject]$currentAdapter
            }
        }
        
        # Get additional adapter info from WMI
        try {
            $wmiAdapters = Get-WmiObject -Class Win32_NetworkAdapter -Filter "NetConnectionStatus IS NOT NULL AND NetConnectionID LIKE '%Wi-Fi%' OR NetConnectionID LIKE '%Wireless%'" -ErrorAction SilentlyContinue
            
            foreach ($wmiAdapter in $wmiAdapters) {
                $adapterConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "Index = $($wmiAdapter.Index)" -ErrorAction SilentlyContinue
                
                if ($adapterConfig) {
                    $wmiAdapterInfo = [PSCustomObject]@{
                        Name = $wmiAdapter.Name
                        Description = $wmiAdapter.Description
                        MACAddress = $wmiAdapter.MACAddress
                        Manufacturer = $wmiAdapter.Manufacturer
                        NetConnectionID = $wmiAdapter.NetConnectionID
                        NetConnectionStatus = $wmiAdapter.NetConnectionStatus
                        Speed = $wmiAdapter.Speed
                        IPAddress = $adapterConfig.IPAddress
                        DefaultIPGateway = $adapterConfig.DefaultIPGateway
                        DNSServerSearchOrder = $adapterConfig.DNSServerSearchOrder
                        DHCPEnabled = $adapterConfig.DHCPEnabled
                    }
                    
                    # Add to adapters if not already present
                    $existingAdapter = $adapters | Where-Object { $_."Physical address" -eq $wmiAdapter.MACAddress -or $_.Name -eq $wmiAdapter.NetConnectionID }
                    if (-not $existingAdapter) {
                        $adapters += $wmiAdapterInfo
                    }
                }
            }
        }
        catch {
            Write-Log "Could not retrieve WMI adapter information: $($_.Exception.Message)" "WARN"
        }
        
        Write-Log "Found $($adapters.Count) WiFi adapters"
        
    }
    catch {
        Write-Log "Error collecting WiFi adapter information: $($_.Exception.Message)" "WARN"
    }
    
    return $adapters
}

function Get-WiFiDriverInfo {
    Write-Log "Collecting WiFi driver information..."
    
    $drivers = @()
    
    try {
        $netshDriverOutput = netsh wlan show drivers
        
        if ($LASTEXITCODE -eq 0) {
            $currentDriver = @{}
            
            foreach ($line in $netshDriverOutput) {
                $line = $line.Trim()
                
                if ($line -eq "" -and $currentDriver.Count -gt 0) {
                    $drivers += [PSCustomObject]$currentDriver
                    $currentDriver = @{}
                }
                elseif ($line -match "^(.+?)\s*:\s*(.+)$") {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    $currentDriver[$key] = $value
                }
            }
            
            if ($currentDriver.Count -gt 0) {
                $drivers += [PSCustomObject]$currentDriver
            }
        }
        
        Write-Log "Found $($drivers.Count) WiFi drivers"
        
    }
    catch {
        Write-Log "Error collecting WiFi driver information: $($_.Exception.Message)" "WARN"
    }
    
    return $drivers
}

# Main execution
try {
    Write-Log "Starting WiFi profiles export for computer: $ComputerName" "SUCCESS"
    Write-Log "Output file: $OutputPath"
    
    # Check if running as administrator for password export
    $isAdmin = Test-AdminRights
    if ($IncludePasswords -and -not $isAdmin) {
        Write-Log "Password export requested but not running as administrator. Passwords will not be included." "WARN"
    }
    
    # Collect WiFi information
    $profiles = Get-WiFiProfiles
    $adapters = Get-WiFiAdapterInfo
    $drivers = Get-WiFiDriverInfo
    
    # Create export object
    $exportData = [PSCustomObject]@{
        ComputerName = $ComputerName
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        IsAdministrator = $isAdmin
        WiFiProfiles = $profiles
        WiFiAdapters = $adapters
        WiFiDrivers = $drivers
        IncludePasswords = $IncludePasswords.IsPresent
        ExportProfiles = $ExportProfiles.IsPresent
        TotalProfiles = $profiles.Count
        ProfilesWithPasswords = ($profiles | Where-Object { $_.Password }).Count
    }
    
    # Export to JSON
    $exportData | ConvertTo-Json -Depth 4 | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "Successfully exported WiFi profiles to: $OutputPath" "SUCCESS"
    Write-Log "Export completed for computer: $ComputerName" "SUCCESS"
    
    # Display summary
    Write-Host "`nSUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer: $ComputerName" -ForegroundColor White
    Write-Host "WiFi Profiles: $($profiles.Count)" -ForegroundColor White
    Write-Host "WiFi Adapters: $($adapters.Count)" -ForegroundColor White
    Write-Host "WiFi Drivers: $($drivers.Count)" -ForegroundColor White
    if ($IncludePasswords) {
        $passwordCount = ($profiles | Where-Object { $_.Password }).Count
        Write-Host "Profiles with Passwords: $passwordCount" -ForegroundColor White
    }
    Write-Host "Output File: $OutputPath" -ForegroundColor White
    
    if ($IncludePasswords -and -not $isAdmin) {
        Write-Host "`nNOTE: Run as Administrator to export WiFi passwords" -ForegroundColor Yellow
    }
    
}
catch {
    Write-Log "Critical error during WiFi profiles export: $($_.Exception.Message)" "ERROR"
    exit 1
}
