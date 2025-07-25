# Export-PrinterSettings.ps1
# PowerShell script to export printer configurations and settings
# Includes installed printers, default printer, and printer preferences

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\PrinterSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDriverDetails
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

function Get-InstalledPrinters {
    Write-Log "Collecting installed printers..."
    
    $printers = @()
    
    try {
        # Get printers using WMI
        $wmiPrinters = Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue
        
        foreach ($printer in $wmiPrinters) {
            $printerInfo = [PSCustomObject]@{
                Name = $printer.Name
                ShareName = $printer.ShareName
                PortName = $printer.PortName
                DriverName = $printer.DriverName
                Location = $printer.Location
                Comment = $printer.Comment
                ServerName = $printer.ServerName
                PrinterStatus = $printer.PrinterStatus
                Default = $printer.Default
                Shared = $printer.Shared
                Local = $printer.Local
                Network = $printer.Network
                WorkOffline = $printer.WorkOffline
                PrinterPaperNames = $printer.PrinterPaperNames
                HorizontalResolution = $printer.HorizontalResolution
                VerticalResolution = $printer.VerticalResolution
                Attributes = $printer.Attributes
                Priority = $printer.Priority
                Capabilities = $printer.Capabilities
                CapabilityDescriptions = $printer.CapabilityDescriptions
            }
            
            $printers += $printerInfo
        }
        
        Write-Log "Found $($printers.Count) installed printers"
        
    }
    catch {
        Write-Log "Error collecting printers: $($_.Exception.Message)" "ERROR"
        return @()
    }
    
    return $printers
}

function Get-PrinterPorts {
    Write-Log "Collecting printer ports..."
    
    $ports = @()
    
    try {
        $wmiPorts = Get-WmiObject -Class Win32_PrinterPort -ErrorAction SilentlyContinue
        
        foreach ($port in $wmiPorts) {
            $portInfo = [PSCustomObject]@{
                Name = $port.Name
                Description = $port.Description
                Type = $port.Type
                HostAddress = $port.HostAddress
                PortNumber = $port.PortNumber
                Protocol = $port.Protocol
                SNMPEnabled = $port.SNMPEnabled
                SNMPCommunity = $port.SNMPCommunity
                QueueSize = $port.QueueSize
                ByteCount = $port.ByteCount
            }
            
            $ports += $portInfo
        }
        
        Write-Log "Found $($ports.Count) printer ports"
        
    }
    catch {
        Write-Log "Error collecting printer ports: $($_.Exception.Message)" "WARN"
        return @()
    }
    
    return $ports
}

function Get-PrinterDrivers {
    Write-Log "Collecting printer drivers..."
    
    $drivers = @()
    
    try {
        $wmiDrivers = Get-WmiObject -Class Win32_PrinterDriver -ErrorAction SilentlyContinue
        
        foreach ($driver in $wmiDrivers) {
            $driverInfo = [PSCustomObject]@{
                Name = $driver.Name
                Version = $driver.Version
                InfName = $driver.InfName
                OEMUrl = $driver.OEMUrl
                DriverPath = $driver.DriverPath
                DataFile = $driver.DataFile
                ConfigFile = $driver.ConfigFile
                HelpFile = $driver.HelpFile
                DriverDate = $driver.DriverDate
                FilePath = $driver.FilePath
                SupportedPlatform = $driver.SupportedPlatform
                MonitorName = $driver.MonitorName
                DefaultDataType = $driver.DefaultDataType
            }
            
            $drivers += $driverInfo
        }
        
        Write-Log "Found $($drivers.Count) printer drivers"
        
    }
    catch {
        Write-Log "Error collecting printer drivers: $($_.Exception.Message)" "WARN"
        return @()
    }
    
    return $drivers
}

function Get-DefaultPrinter {
    Write-Log "Getting default printer..."
    
    try {
        # Get default printer from registry
        $defaultPrinter = Get-ItemProperty "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows" -Name "Device" -ErrorAction SilentlyContinue
        
        if ($defaultPrinter -and $defaultPrinter.Device) {
            $deviceInfo = $defaultPrinter.Device.Split(',')
            return [PSCustomObject]@{
                Name = $deviceInfo[0]
                Driver = if ($deviceInfo.Length -gt 1) { $deviceInfo[1] } else { $null }
                Port = if ($deviceInfo.Length -gt 2) { $deviceInfo[2] } else { $null }
                FullString = $defaultPrinter.Device
            }
        }
        
        return $null
    }
    catch {
        Write-Log "Error getting default printer: $($_.Exception.Message)" "WARN"
        return $null
    }
}

function Get-PrinterPreferences {
    Write-Log "Collecting printer preferences from registry..."
    
    $preferences = @{}
    
    try {
        # Check both user and system printer preferences
        $registryPaths = @(
            "HKCU:\Printers\DevModePerUser",
            "HKCU:\Printers\Defaults"
        )
        
        foreach ($regPath in $registryPaths) {
            if (Test-Path $regPath) {
                $pathName = Split-Path $regPath -Leaf
                $preferences[$pathName] = @{}
                
                try {
                    $printerKeys = Get-ChildItem $regPath -ErrorAction SilentlyContinue
                    
                    foreach ($key in $printerKeys) {
                        $printerName = Split-Path $key.Name -Leaf
                        $printerPrefs = @{}
                        
                        $properties = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                        if ($properties) {
                            $properties.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                                # Convert binary data to base64 for JSON serialization
                                if ($_.Value -is [byte[]]) {
                                    $printerPrefs[$_.Name] = [Convert]::ToBase64String($_.Value)
                                } else {
                                    $printerPrefs[$_.Name] = $_.Value
                                }
                            }
                        }
                        
                        if ($printerPrefs.Count -gt 0) {
                            $preferences[$pathName][$printerName] = $printerPrefs
                        }
                    }
                }
                catch {
                    Write-Log "Error reading preferences from $regPath`: $($_.Exception.Message)" "WARN"
                }
            }
        }
        
    }
    catch {
        Write-Log "Error collecting printer preferences: $($_.Exception.Message)" "WARN"
    }
    
    return $preferences
}

function Get-PrintSpoolerSettings {
    Write-Log "Collecting print spooler settings..."
    
    $spoolerSettings = @{}
    
    try {
        $spoolerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print"
        if (Test-Path $spoolerPath) {
            $spoolerProps = Get-ItemProperty $spoolerPath -ErrorAction SilentlyContinue
            if ($spoolerProps) {
                $spoolerProps.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                    $spoolerSettings[$_.Name] = $_.Value
                }
            }
        }
        
        # Get spooler service info
        $spoolerService = Get-Service -Name "Spooler" -ErrorAction SilentlyContinue
        if ($spoolerService) {
            $spoolerSettings["ServiceStatus"] = $spoolerService.Status
            $spoolerSettings["ServiceStartType"] = $spoolerService.StartType
        }
        
    }
    catch {
        Write-Log "Error collecting print spooler settings: $($_.Exception.Message)" "WARN"
    }
    
    return $spoolerSettings
}

# Main execution
try {
    Write-Log "Starting printer settings export for computer: $ComputerName" "SUCCESS"
    Write-Log "Output file: $OutputPath"
    
    # Collect printer information
    $printers = Get-InstalledPrinters
    $ports = Get-PrinterPorts
    $defaultPrinter = Get-DefaultPrinter
    $preferences = Get-PrinterPreferences
    $spoolerSettings = Get-PrintSpoolerSettings
    
    # Collect drivers if requested
    $drivers = @()
    if ($IncludeDriverDetails) {
        $drivers = Get-PrinterDrivers
    }
    
    # Create export object
    $exportData = [PSCustomObject]@{
        ComputerName = $ComputerName
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        DefaultPrinter = $defaultPrinter
        Printers = $printers
        PrinterPorts = $ports
        PrinterDrivers = $drivers
        PrinterPreferences = $preferences
        SpoolerSettings = $spoolerSettings
        IncludeDriverDetails = $IncludeDriverDetails.IsPresent
        TotalPrinters = $printers.Count
    }
    
    # Export to JSON
    $exportData | ConvertTo-Json -Depth 4 | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "Successfully exported printer settings to: $OutputPath" "SUCCESS"
    Write-Log "Export completed for computer: $ComputerName" "SUCCESS"
    
    # Display summary
    Write-Host "`nSUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer: $ComputerName" -ForegroundColor White
    Write-Host "Total Printers: $($printers.Count)" -ForegroundColor White
    Write-Host "Printer Ports: $($ports.Count)" -ForegroundColor White
    if ($defaultPrinter) {
        Write-Host "Default Printer: $($defaultPrinter.Name)" -ForegroundColor White
    }
    if ($IncludeDriverDetails) {
        Write-Host "Printer Drivers: $($drivers.Count)" -ForegroundColor White
    }
    Write-Host "Output File: $OutputPath" -ForegroundColor White
    
}
catch {
    Write-Log "Critical error during printer settings export: $($_.Exception.Message)" "ERROR"
    exit 1
}
