# AppComparisonToolkit.psm1
# PowerShell module for Application Comparison Toolkit

# Import the main scripts as functions
. "$PSScriptRoot\Export-InstalledApps.ps1"
. "$PSScriptRoot\Compare-InstalledApps.ps1"
. "$PSScriptRoot\Start-AppComparison.ps1"

# Create function aliases that follow PowerShell naming conventions
function Export-InstalledApplications {
    <#
    .SYNOPSIS
    Exports installed applications from a Windows computer.
    
    .DESCRIPTION
    This function collects installed application information from multiple sources including
    Windows Registry, WMI, Package Managers, and AppX packages. The results are exported
    to a JSON file for later comparison.
    
    .PARAMETER OutputPath
    Path for the output JSON file. If not specified, auto-generates a timestamped filename.
    
    .PARAMETER ComputerName
    Name to identify the computer in reports. Defaults to current computer name.
    
    .PARAMETER IncludeSystemComponents
    Include system components like Visual C++ redistributables and .NET frameworks.
    
    .PARAMETER IncludeUpdates
    Include Windows updates and hotfixes in the export.
    
    .EXAMPLE
    Export-InstalledApplications -ComputerName "WORKSTATION01"
    
    .EXAMPLE
    Export-InstalledApplications -OutputPath "C:\Reports\MyComputer.json" -IncludeSystemComponents
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = ".\InstalledApps_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
        
        [Parameter(Mandatory = $false)]
        [string]$ComputerName = $env:COMPUTERNAME,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeSystemComponents,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeUpdates
    )
    
    # Call the original script with parameters
    $scriptParams = @{
        OutputPath = $OutputPath
        ComputerName = $ComputerName
    }
    
    if ($IncludeSystemComponents) { $scriptParams.IncludeSystemComponents = $true }
    if ($IncludeUpdates) { $scriptParams.IncludeUpdates = $true }
    
    & "$PSScriptRoot\Export-InstalledApps.ps1" @scriptParams
}

function Compare-ApplicationLists {
    <#
    .SYNOPSIS
    Compares installed applications between two computers.
    
    .DESCRIPTION
    This function analyzes application differences between two computers by comparing
    their exported application lists. It generates detailed reports showing applications
    that exist only on each computer, version differences, and common applications.
    
    .PARAMETER Computer1File
    Path to the first computer's JSON export file.
    
    .PARAMETER Computer2File
    Path to the second computer's JSON export file.
    
    .PARAMETER OutputPath
    Path for the HTML report output. Auto-generates if not specified.
    
    .PARAMETER ExportToJson
    Also export comparison results to JSON format.
    
    .PARAMETER GroupSimilarApps
    Group applications with similar names together.
    
    .PARAMETER DetailedReport
    Include common applications in the report.
    
    .EXAMPLE
    Compare-ApplicationLists -Computer1File "PC1.json" -Computer2File "PC2.json"
    
    .EXAMPLE
    Compare-ApplicationLists -Computer1File "PC1.json" -Computer2File "PC2.json" -DetailedReport -ExportToJson
    #>
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
        [switch]$GroupSimilarApps,
        
        [Parameter(Mandatory = $false)]
        [switch]$DetailedReport
    )
    
    # Call the original script with parameters
    $scriptParams = @{
        Computer1File = $Computer1File
        Computer2File = $Computer2File
        OutputPath = $OutputPath
    }
    
    if ($ExportToJson) { $scriptParams.ExportToJson = $true }
    if ($GroupSimilarApps) { $scriptParams.GroupSimilarApps = $true }
    if ($DetailedReport) { $scriptParams.DetailedReport = $true }
    
    & "$PSScriptRoot\Compare-InstalledApps.ps1" @scriptParams
}

function Start-ApplicationComparison {
    <#
    .SYNOPSIS
    Starts the guided application comparison workflow.
    
    .DESCRIPTION
    This function provides a simplified, guided workflow for comparing applications
    between computers. It handles the export process and helps users through the
    comparison steps with interactive prompts.
    
    .PARAMETER Computer1Name
    Name for the first computer (current computer).
    
    .PARAMETER Computer2Name
    Name for the second computer.
    
    .PARAMETER WorkingDirectory
    Directory for storing export files and reports.
    
    .PARAMETER IncludeSystemComponents
    Include system components in the export.
    
    .PARAMETER IncludeUpdates
    Include Windows updates in the export.
    
    .PARAMETER GroupSimilarApps
    Group similar applications in the comparison.
    
    .PARAMETER DetailedReport
    Generate detailed reports including common applications.
    
    .PARAMETER ExportToJson
    Also export comparison results to JSON format.
    
    .EXAMPLE
    Start-ApplicationComparison
    
    .EXAMPLE
    Start-ApplicationComparison -WorkingDirectory "C:\AppReports" -DetailedReport
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Computer1Name = "Computer1",
        
        [Parameter(Mandatory = $false)]
        [string]$Computer2Name = "Computer2",
        
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory = ".\AppComparison",
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeSystemComponents,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeUpdates,
        
        [Parameter(Mandatory = $false)]
        [switch]$GroupSimilarApps,
        
        [Parameter(Mandatory = $false)]
        [switch]$DetailedReport,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportToJson
    )
    
    # Call the original script with parameters
    $scriptParams = @{
        Computer1Name = $Computer1Name
        Computer2Name = $Computer2Name
        WorkingDirectory = $WorkingDirectory
    }
    
    if ($IncludeSystemComponents) { $scriptParams.IncludeSystemComponents = $true }
    if ($IncludeUpdates) { $scriptParams.IncludeUpdates = $true }
    if ($GroupSimilarApps) { $scriptParams.GroupSimilarApps = $true }
    if ($DetailedReport) { $scriptParams.DetailedReport = $true }
    if ($ExportToJson) { $scriptParams.ExportToJson = $true }
    
    & "$PSScriptRoot\Start-AppComparison.ps1" @scriptParams
}

# Export the functions
Export-ModuleMember -Function Export-InstalledApplications, Compare-ApplicationLists, Start-ApplicationComparison

# Display module information when imported
Write-Host @"
==========================================
  APPLICATION COMPARISON TOOLKIT v1.0.0
==========================================

Module loaded successfully!

Available functions:
- Export-InstalledApplications
- Compare-ApplicationLists  
- Start-ApplicationComparison

For help with any function, use: Get-Help <FunctionName> -Detailed

Quick start: Start-ApplicationComparison

"@ -ForegroundColor Green
