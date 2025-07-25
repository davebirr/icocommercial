# Start-NetworkAppComparison.ps1
# Enhanced workflow script for comparing applications across computers on a local network
# Supports both local and remote computers with local authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$ComputerNames = @(),
    
    [Parameter(Mandatory = $false)]
    [string]$Username,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = ".\NetworkAppComparison",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemComponents,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeUpdates,
    
    [Parameter(Mandatory = $false)]
    [switch]$UseWinRM,
    
    [Parameter(Mandatory = $false)]
    [switch]$InteractiveMode
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

function Show-Menu {
    Write-Host @"

==========================================
  NETWORK APPLICATION COMPARISON
==========================================

This script helps you compare installed applications across multiple computers 
on your local network, including computers not joined to a domain.

AUTHENTICATION OPTIONS:
- Local user accounts (for workgroup computers)
- Administrator accounts on each computer
- Same credentials across multiple computers

METHODS SUPPORTED:
- WinRM (if enabled on remote computers)
- File share + script copy (using administrative shares)
- Manual export (copy scripts to each computer)

"@ -ForegroundColor Cyan
}

function Get-ComputerList {
    if ($ComputerNames.Count -gt 0) {
        return $ComputerNames
    }
    
    Write-Host "`nComputer Discovery Options:" -ForegroundColor Yellow
    Write-Host "1. Enter computer names manually"
    Write-Host "2. Scan local network subnet"
    Write-Host "3. Import from file"
    
    $choice = Read-Host "`nSelect option (1-3)"
    
    switch ($choice) {
        "1" {
            $computers = @()
            Write-Host "`nEnter computer names (press Enter on empty line to finish):"
            do {
                $computer = Read-Host "Computer name"
                if ($computer.Trim()) {
                    $computers += $computer.Trim()
                    Write-Host "Added: $computer" -ForegroundColor Green
                }
            } while ($computer.Trim())
            
            return $computers
        }
        
        "2" {
            Write-Host "`nScanning local network subnet for computers..." -ForegroundColor Yellow
            Write-Host "This may take a few minutes..." -ForegroundColor Gray
            
            # Get local IP range
            $localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -eq "Dhcp" -or $_.PrefixOrigin -eq "Manual" })[0].IPAddress
            $subnet = $localIP.Substring(0, $localIP.LastIndexOf('.'))
            
            $computers = @()
            $jobs = @()
            
            # Scan IP range
            for ($i = 1; $i -le 254; $i++) {
                $ip = "$subnet.$i"
                $jobs += Start-Job -ScriptBlock {
                    param($ip)
                    if (Test-Connection -ComputerName $ip -Count 1 -Quiet -TimeoutSeconds 1) {
                        try {
                            $hostname = [System.Net.Dns]::GetHostEntry($ip).HostName
                            return @{ IP = $ip; Hostname = $hostname }
                        }
                        catch {
                            return @{ IP = $ip; Hostname = $ip }
                        }
                    }
                } -ArgumentList $ip
            }
            
            # Wait for jobs and collect results
            $results = $jobs | Wait-Job | Receive-Job
            $jobs | Remove-Job
            
            if ($results) {
                Write-Host "`nFound computers:" -ForegroundColor Green
                for ($i = 0; $i -lt $results.Count; $i++) {
                    Write-Host "$($i + 1). $($results[$i].Hostname) ($($results[$i].IP))"
                }
                
                $selection = Read-Host "`nEnter numbers to include (e.g., 1,3,5 or 'all')"
                if ($selection -eq "all") {
                    $computers = $results | ForEach-Object { $_.Hostname }
                } else {
                    $indices = $selection.Split(',') | ForEach-Object { [int]$_.Trim() - 1 }
                    $computers = $indices | ForEach-Object { $results[$_].Hostname }
                }
            } else {
                Write-Host "No computers found on local network" -ForegroundColor Red
                return @()
            }
            
            return $computers
        }
        
        "3" {
            $filePath = Read-Host "Enter path to text file with computer names (one per line)"
            if (Test-Path $filePath) {
                return Get-Content $filePath | Where-Object { $_.Trim() }
            } else {
                Write-Host "File not found: $filePath" -ForegroundColor Red
                return @()
            }
        }
        
        default {
            Write-Host "Invalid selection" -ForegroundColor Red
            return @()
        }
    }
}

function Test-Prerequisites {
    Write-Log "Checking prerequisites..." "INFO"
    
    # Check if Export-RemoteInstalledApps.ps1 exists
    if (-not (Test-Path ".\Export-RemoteInstalledApps.ps1")) {
        Write-Log "Export-RemoteInstalledApps.ps1 not found in current directory" "ERROR"
        return $false
    }
    
    # Check if other required scripts exist
    $requiredScripts = @("Export-InstalledApps.ps1", "Compare-InstalledApps.ps1")
    foreach ($script in $requiredScripts) {
        if (-not (Test-Path ".\$script")) {
            Write-Log "$script not found in current directory" "ERROR"
            return $false
        }
    }
    
    # Check if running as administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Log "Running as administrator is recommended for best results" "WARN"
        $response = Read-Host "Continue anyway? (y/n)"
        if ($response -notmatch '^[Yy]') {
            return $false
        }
    }
    
    return $true
}

function Get-AuthenticationMethod {
    if ($Credential) {
        return @{ Method = "Credential"; Credential = $Credential }
    }
    
    Write-Host "`nAuthentication Options:" -ForegroundColor Yellow
    Write-Host "1. Use same credentials for all computers"
    Write-Host "2. Prompt for credentials per computer"
    Write-Host "3. Use current Windows credentials (domain only)"
    
    $choice = Read-Host "`nSelect authentication method (1-3)"
    
    switch ($choice) {
        "1" {
            $cred = Get-Credential -Message "Enter credentials to use for all remote computers"
            return @{ Method = "SharedCredential"; Credential = $cred }
        }
        "2" {
            return @{ Method = "PerComputer"; Credential = $null }
        }
        "3" {
            return @{ Method = "CurrentUser"; Credential = $null }
        }
        default {
            Write-Host "Invalid selection, using per-computer prompts" -ForegroundColor Yellow
            return @{ Method = "PerComputer"; Credential = $null }
        }
    }
}

# Main execution
try {
    Show-Menu
    
    if (-not (Test-Prerequisites)) {
        Write-Log "Prerequisites not met. Please ensure all required scripts are present." "ERROR"
        exit 1
    }
    
    # Create working directory
    if (-not (Test-Path $WorkingDirectory)) {
        New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
        Write-Log "Created working directory: $WorkingDirectory"
    }
    
    # Get list of computers to process
    Write-Log "Getting list of computers to process..." "INFO"
    $computers = Get-ComputerList
    
    if ($computers.Count -eq 0) {
        Write-Log "No computers specified. Exiting." "ERROR"
        exit 1
    }
    
    Write-Log "Will process $($computers.Count) computers: $($computers -join ', ')"
    
    # Get authentication method
    $authMethod = Get-AuthenticationMethod
    Write-Log "Authentication method: $($authMethod.Method)"
    
    # Export from current computer first
    Write-Log "Exporting applications from current computer ($env:COMPUTERNAME)..." "INFO"
    
    $localOutputPath = Join-Path $WorkingDirectory "Apps_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $exportParams = @{
        OutputPath = $localOutputPath
        ComputerName = $env:COMPUTERNAME
    }
    
    if ($IncludeSystemComponents) { $exportParams.IncludeSystemComponents = $true }
    if ($IncludeUpdates) { $exportParams.IncludeUpdates = $true }
    
    & ".\Export-InstalledApps.ps1" @exportParams
    
    # Export from remote computers
    if ($computers.Count -gt 0) {
        Write-Log "Starting remote exports..." "INFO"
        
        $remoteParams = @{
            ComputerNames = $computers
            OutputDirectory = $WorkingDirectory
        }
        
        if ($authMethod.Credential) { $remoteParams.Credential = $authMethod.Credential }
        if ($Username) { $remoteParams.Username = $Username }
        if ($IncludeSystemComponents) { $remoteParams.IncludeSystemComponents = $true }
        if ($IncludeUpdates) { $remoteParams.IncludeUpdates = $true }
        if ($UseWinRM) { $remoteParams.UseWinRM = $true }
        
        & ".\Export-RemoteInstalledApps.ps1" @remoteParams
    }
    
    # Find exported files and offer to compare them
    Write-Log "Scanning for exported files..." "INFO"
    $exportFiles = Get-ChildItem -Path $WorkingDirectory -Filter "Apps_*.json" | Sort-Object Name
    
    if ($exportFiles.Count -ge 2) {
        Write-Host "`nFound $($exportFiles.Count) exported files:" -ForegroundColor Green
        for ($i = 0; $i -lt $exportFiles.Count; $i++) {
            Write-Host "  $($i + 1). $($exportFiles[$i].Name)"
        }
        
        $response = Read-Host "`nWould you like to compare these files? (y/n)"
        if ($response -match '^[Yy]') {
            
            # For multiple files, offer comparison options
            if ($exportFiles.Count -eq 2) {
                # Simple comparison of two files
                $file1 = $exportFiles[0].FullName
                $file2 = $exportFiles[1].FullName
                
                Write-Log "Comparing $($exportFiles[0].Name) with $($exportFiles[1].Name)..."
                
                $compareParams = @{
                    Computer1File = $file1
                    Computer2File = $file2
                    OutputPath = Join-Path $WorkingDirectory "Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
                    DetailedReport = $true
                    ExportToJson = $true
                }
                
                & ".\Compare-InstalledApps.ps1" @compareParams
                
            } else {
                # Multiple files - offer comparison matrix
                Write-Host "`nComparison Options:" -ForegroundColor Yellow
                Write-Host "1. Compare all files with the first file"
                Write-Host "2. Compare specific pair of files"
                Write-Host "3. Create comparison matrix (all vs all)"
                
                $compChoice = Read-Host "Select option (1-3)"
                
                switch ($compChoice) {
                    "1" {
                        $baseFile = $exportFiles[0].FullName
                        for ($i = 1; $i -lt $exportFiles.Count; $i++) {
                            $compareFile = $exportFiles[$i].FullName
                            $outputName = "Comparison_$($exportFiles[0].BaseName)_vs_$($exportFiles[$i].BaseName).html"
                            $outputPath = Join-Path $WorkingDirectory $outputName
                            
                            Write-Log "Comparing $($exportFiles[0].Name) with $($exportFiles[$i].Name)..."
                            
                            & ".\Compare-InstalledApps.ps1" -Computer1File $baseFile -Computer2File $compareFile -OutputPath $outputPath -DetailedReport -ExportToJson -ExportToMarkdown
                        }
                    }
                    
                    "2" {
                        Write-Host "Select first file:"
                        for ($i = 0; $i -lt $exportFiles.Count; $i++) {
                            Write-Host "  $($i + 1). $($exportFiles[$i].Name)"
                        }
                        $file1Index = [int](Read-Host "Enter number") - 1
                        
                        Write-Host "Select second file:"
                        for ($i = 0; $i -lt $exportFiles.Count; $i++) {
                            if ($i -ne $file1Index) {
                                Write-Host "  $($i + 1). $($exportFiles[$i].Name)"
                            }
                        }
                        $file2Index = [int](Read-Host "Enter number") - 1
                        
                        $file1 = $exportFiles[$file1Index].FullName
                        $file2 = $exportFiles[$file2Index].FullName
                        $outputName = "Comparison_$($exportFiles[$file1Index].BaseName)_vs_$($exportFiles[$file2Index].BaseName).html"
                        $outputPath = Join-Path $WorkingDirectory $outputName
                        
                        & ".\Compare-InstalledApps.ps1" -Computer1File $file1 -Computer2File $file2 -OutputPath $outputPath -DetailedReport -ExportToJson -ExportToMarkdown
                    }
                    
                    "3" {
                        Write-Log "Creating comparison matrix..."
                        for ($i = 0; $i -lt $exportFiles.Count; $i++) {
                            for ($j = $i + 1; $j -lt $exportFiles.Count; $j++) {
                                $file1 = $exportFiles[$i].FullName
                                $file2 = $exportFiles[$j].FullName
                                $outputName = "Comparison_$($exportFiles[$i].BaseName)_vs_$($exportFiles[$j].BaseName).html"
                                $outputPath = Join-Path $WorkingDirectory $outputName
                                
                                Write-Log "Comparing $($exportFiles[$i].Name) with $($exportFiles[$j].Name)..."
                                
                                & ".\Compare-InstalledApps.ps1" -Computer1File $file1 -Computer2File $file2 -OutputPath $outputPath -DetailedReport -ExportToJson -ExportToMarkdown
                            }
                        }
                    }
                }
            }
        }
    } else {
        Write-Log "Need at least 2 exported files to perform comparison" "WARN"
    }
    
    # Generate directory summary report
    Write-Log "Generating directory summary report..." "INFO"
    if (Test-Path ".\New-DirectorySummary.ps1") {
        try {
            & ".\New-DirectorySummary.ps1" -DirectoryPath $WorkingDirectory
        }
        catch {
            Write-Log "Error generating directory summary: $($_.Exception.Message)" "WARN"
        }
    }

    Write-Host @"==========================================
  NETWORK COMPARISON COMPLETED
==========================================

Working Directory: $WorkingDirectory

Next Steps:
1. Review the exported JSON files
2. Open the generated HTML comparison reports
3. Check the detailed comparison results

All files are saved in: $WorkingDirectory

"@ -ForegroundColor Green

}
catch {
    Write-Log "Critical error during network comparison: $($_.Exception.Message)" "ERROR"
    exit 1
}
