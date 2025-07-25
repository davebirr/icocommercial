# Export-RemoteInstalledApps.ps1
# PowerShell script to export installed applications from remote computers using local credentials
# Supports computers not on a domain by using local username/password authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$ComputerNames,
    
    [Parameter(Mandatory = $false)]
    [string]$Username,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = ".\RemoteExports",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemComponents,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeUpdates,
    
    [Parameter(Mandatory = $false)]
    [switch]$UseWinRM,
    
    [Parameter(Mandatory = $false)]
    [int]$TimeoutSeconds = 300,
    
    [Parameter(Mandatory = $false)]
    [switch]$CopyScriptFirst
)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(
        switch ($Level) {
            "ERROR" { "Red" }
            "WARN" { "Yellow" }
            "SUCCESS" { "Green" }
            "INFO" { "Cyan" }
            default { "White" }
        }
    )
}

function Test-RemoteConnectivity {
    param([string]$ComputerName)
    
    Write-Log "Testing connectivity to $ComputerName..."
    
    # Test basic network connectivity
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        Write-Log "Cannot reach $ComputerName via ping" "ERROR"
        return $false
    }
    
    # Test WinRM if requested
    if ($UseWinRM) {
        try {
            Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
            Write-Log "WinRM connectivity to $ComputerName successful" "SUCCESS"
            return $true
        }
        catch {
            Write-Log "WinRM not available on $ComputerName`: $($_.Exception.Message)" "WARN"
            return $false
        }
    }
    
    # Test file share access (for script copying method)
    try {
        $adminShare = "\\$ComputerName\C$"
        if (Test-Path $adminShare -ErrorAction Stop) {
            Write-Log "Administrative share access to $ComputerName successful" "SUCCESS"
            return $true
        }
    }
    catch {
        Write-Log "Cannot access administrative share on $ComputerName`: $($_.Exception.Message)" "WARN"
        return $false
    }
    
    return $false
}

function Get-RemoteCredential {
    param([string]$ComputerName, [string]$Username)
    
    if ($Credential) {
        return $Credential
    }
    
    if ($Username) {
        $securePassword = Read-Host "Enter password for $Username on $ComputerName" -AsSecureString
        return New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    }
    
    # Prompt for credentials
    $message = "Enter local credentials for $ComputerName"
    return Get-Credential -Message $message
}

function Export-UsingWinRM {
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$RemoteCredential,
        [string]$OutputPath
    )
    
    Write-Log "Exporting applications from $ComputerName using WinRM..."
    
    try {
        # Create the script block to run remotely
        $scriptBlock = {
            param($IncludeSystemComponents, $IncludeUpdates, $ComputerName)
            
            # Import the functions we need (simplified version for remote execution)
            function Get-InstalledAppsFromRegistry {
                param($IncludeSystemComponents, $IncludeUpdates)
                
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
                        Write-Warning "Error reading registry path $path`: $($_.Exception.Message)"
                    }
                }
                
                return $apps
            }
            
            function Get-InstalledAppsFromPackageManager {
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
                    Write-Warning "Error reading from Package Manager: $($_.Exception.Message)"
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
                    Write-Warning "Error reading AppX packages: $($_.Exception.Message)"
                }
                
                return $apps
            }
            
            # Collect applications
            $registryApps = Get-InstalledAppsFromRegistry -IncludeSystemComponents $IncludeSystemComponents -IncludeUpdates $IncludeUpdates
            $packageApps = Get-InstalledAppsFromPackageManager
            
            # Merge and deduplicate
            $allApps = @()
            $uniqueApps = @{}
            
            foreach ($app in ($registryApps + $packageApps)) {
                $key = "$($app.Name)_$($app.Version)".ToLower()
                if (-not $uniqueApps.ContainsKey($key)) {
                    $uniqueApps[$key] = $app
                }
            }
            
            $finalApps = $uniqueApps.Values | Sort-Object Name
            
            # Create export object
            $exportData = [PSCustomObject]@{
                ComputerName = $ComputerName
                ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                TotalApplications = $finalApps.Count
                IncludeSystemComponents = $IncludeSystemComponents
                IncludeUpdates = $IncludeUpdates
                ExportMethod = "WinRM"
                Applications = $finalApps
            }
            
            return $exportData
        }
        
        # Execute the script block remotely
        $session = New-PSSession -ComputerName $ComputerName -Credential $RemoteCredential -ErrorAction Stop
        
        try {
            $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $IncludeSystemComponents.IsPresent, $IncludeUpdates.IsPresent, $ComputerName
            
            # Export to JSON
            $result | ConvertTo-Json -Depth 3 | Out-File -FilePath $OutputPath -Encoding UTF8
            
            Write-Log "Successfully exported $($result.TotalApplications) applications from $ComputerName" "SUCCESS"
            return $true
        }
        finally {
            Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log "Error exporting from $ComputerName via WinRM: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Export-UsingScriptCopy {
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$RemoteCredential,
        [string]$OutputPath
    )
    
    Write-Log "Exporting applications from $ComputerName using script copy method..."
    
    try {
        # Map network drive with credentials
        $tempDriveLetter = "Z:"
        $networkPath = "\\$ComputerName\C$"
        
        # Remove any existing mapping
        try { net use $tempDriveLetter /delete /yes 2>$null } catch { }
        
        # Map the drive with credentials
        $netUseCommand = "net use $tempDriveLetter $networkPath /user:$($RemoteCredential.UserName) $($RemoteCredential.GetNetworkCredential().Password)"
        $result = cmd /c $netUseCommand 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to map network drive: $result"
        }
        
        Write-Log "Successfully mapped network drive to $ComputerName"
        
        # Copy the export script to the remote computer
        $remoteScriptPath = "$tempDriveLetter\Temp\Export-InstalledApps-Remote.ps1"
        $remoteTempDir = "$tempDriveLetter\Temp"
        
        if (-not (Test-Path $remoteTempDir)) {
            New-Item -ItemType Directory -Path $remoteTempDir -Force | Out-Null
        }
        
        Copy-Item -Path ".\Export-InstalledApps.ps1" -Destination $remoteScriptPath -Force
        Write-Log "Copied script to remote computer"
        
        # Create a batch file to execute PowerShell script remotely
        $batchContent = @"
@echo off
cd /d C:\Temp
powershell.exe -ExecutionPolicy Bypass -File Export-InstalledApps-Remote.ps1 -OutputPath "C:\Temp\Apps_$ComputerName.json" -ComputerName "$ComputerName"$(if ($IncludeSystemComponents) { " -IncludeSystemComponents" })$(if ($IncludeUpdates) { " -IncludeUpdates" })
"@
        
        $batchPath = "$tempDriveLetter\Temp\RunExport.bat"
        $batchContent | Out-File -FilePath $batchPath -Encoding ASCII
        
        # Execute the batch file remotely using PsExec or scheduled task
        Write-Log "Executing export script on $ComputerName..."
        
        # Try using PsExec if available
        $psExecPath = Get-Command psexec.exe -ErrorAction SilentlyContinue
        if ($psExecPath) {
            $psExecArgs = @(
                "\\$ComputerName",
                "-u", $RemoteCredential.UserName,
                "-p", $RemoteCredential.GetNetworkCredential().Password,
                "-d",
                "C:\Temp\RunExport.bat"
            )
            
            & $psExecPath.Source @psExecArgs
            Start-Sleep -Seconds 30  # Wait for execution
        } else {
            Write-Log "PsExec not found. Please install PsExec or use WinRM method." "WARN"
            Write-Log "You can manually run: C:\Temp\RunExport.bat on $ComputerName" "INFO"
            Read-Host "Press Enter after manually executing the script on $ComputerName"
        }
        
        # Copy the result back
        $remoteOutputPath = "$tempDriveLetter\Temp\Apps_$ComputerName.json"
        if (Test-Path $remoteOutputPath) {
            Copy-Item -Path $remoteOutputPath -Destination $OutputPath -Force
            Write-Log "Successfully copied results from $ComputerName" "SUCCESS"
            
            # Cleanup remote files
            Remove-Item -Path $remoteScriptPath -Force -ErrorAction SilentlyContinue
            Remove-Item -Path $batchPath -Force -ErrorAction SilentlyContinue
            Remove-Item -Path $remoteOutputPath -Force -ErrorAction SilentlyContinue
            
            return $true
        } else {
            Write-Log "Output file not found on $ComputerName. Script may have failed." "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error with script copy method for $ComputerName`: $($_.Exception.Message)" "ERROR"
        return $false
    }
    finally {
        # Cleanup mapped drive
        try { net use $tempDriveLetter /delete /yes 2>$null } catch { }
    }
}

# Main execution
try {
    Write-Log "Starting remote application export for $($ComputerNames.Count) computer(s)..." "SUCCESS"
    
    # Create output directory
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Log "Created output directory: $OutputDirectory"
    }
    
    $successfulExports = 0
    $failedExports = 0
    
    foreach ($computerName in $ComputerNames) {
        Write-Log "Processing computer: $computerName" "INFO"
        
        # Test connectivity
        if (-not (Test-RemoteConnectivity -ComputerName $computerName)) {
            Write-Log "Skipping $computerName due to connectivity issues" "ERROR"
            $failedExports++
            continue
        }
        
        # Get credentials for this computer
        try {
            $remoteCredential = Get-RemoteCredential -ComputerName $computerName -Username $Username
        }
        catch {
            Write-Log "Failed to get credentials for $computerName`: $($_.Exception.Message)" "ERROR"
            $failedExports++
            continue
        }
        
        # Determine output path
        $outputPath = Join-Path $OutputDirectory "Apps_$computerName_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        
        # Try export methods
        $exportSuccess = $false
        
        if ($UseWinRM) {
            $exportSuccess = Export-UsingWinRM -ComputerName $computerName -RemoteCredential $remoteCredential -OutputPath $outputPath
        }
        
        if (-not $exportSuccess) {
            Write-Log "Attempting script copy method for $computerName..."
            $exportSuccess = Export-UsingScriptCopy -ComputerName $computerName -RemoteCredential $remoteCredential -OutputPath $outputPath
        }
        
        if ($exportSuccess) {
            $successfulExports++
            Write-Log "Export completed for $computerName" "SUCCESS"
        } else {
            $failedExports++
            Write-Log "Export failed for $computerName" "ERROR"
        }
        
        Write-Host "" # Add spacing between computers
    }
    
    # Summary
    Write-Host "`nREMOTE EXPORT SUMMARY:" -ForegroundColor Cyan
    Write-Host "Successful exports: $successfulExports" -ForegroundColor Green
    Write-Host "Failed exports: $failedExports" -ForegroundColor Red
    Write-Host "Output directory: $OutputDirectory" -ForegroundColor White
    
    if ($successfulExports -gt 0) {
        Write-Host "`nNext steps:" -ForegroundColor Yellow
        Write-Host "1. Review the exported JSON files in: $OutputDirectory" -ForegroundColor White
        Write-Host "2. Use Compare-InstalledApps.ps1 to compare the files" -ForegroundColor White
        Write-Host "3. Example: .\Compare-InstalledApps.ps1 -Computer1File 'file1.json' -Computer2File 'file2.json'" -ForegroundColor White
    }
    
}
catch {
    Write-Log "Critical error during remote export: $($_.Exception.Message)" "ERROR"
    exit 1
}
