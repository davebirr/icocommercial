# Start-AppComparison.ps1
# Simplified script to run the complete application comparison workflow

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

# Create working directory if it doesn't exist
if (-not (Test-Path $WorkingDirectory)) {
    New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
    Write-Log "Created working directory: $WorkingDirectory"
}

Write-Host @"

==========================================
  APPLICATION COMPARISON WORKFLOW
==========================================

This script will help you compare installed applications between two computers.

WORKFLOW:
1. Export applications from current computer (Computer 1)
2. Provide instructions for Computer 2
3. Compare the application lists
4. Generate detailed reports

"@ -ForegroundColor Cyan

# Step 1: Export from current computer
Write-Log "Step 1: Exporting applications from current computer ($env:COMPUTERNAME)..." "INFO"

$computer1File = Join-Path $WorkingDirectory "Apps_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"

$exportParams = @{
    OutputPath = $computer1File
    ComputerName = $env:COMPUTERNAME
}

if ($IncludeSystemComponents) { $exportParams.IncludeSystemComponents = $true }
if ($IncludeUpdates) { $exportParams.IncludeUpdates = $true }

try {
    & ".\Export-InstalledApps.ps1" @exportParams
    Write-Log "Successfully exported applications from $env:COMPUTERNAME" "SUCCESS"
}
catch {
    Write-Log "Error exporting applications: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Step 2: Instructions for Computer 2
Write-Host @"

==========================================
  INSTRUCTIONS FOR COMPUTER 2
==========================================

To complete the comparison, you need to run the export script on the second computer.

STEPS FOR COMPUTER 2:
1. Copy these files to the second computer:
   - Export-InstalledApps.ps1
   
2. Run this command on Computer 2:
   .\Export-InstalledApps.ps1 -OutputPath "Apps_$Computer2Name.json" -ComputerName "$Computer2Name"
   
3. Copy the generated JSON file back to this computer in the folder:
   $WorkingDirectory

4. Run this script again with the -Computer2File parameter:
   .\Start-AppComparison.ps1 -Computer2File "$WorkingDirectory\Apps_$Computer2Name.json"

"@ -ForegroundColor Yellow

# Check if Computer 2 file is already available
$computer2Pattern = Join-Path $WorkingDirectory "Apps_*.json"
$computer2Files = Get-ChildItem $computer2Pattern -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne (Split-Path $computer1File -Leaf) }

if ($computer2Files) {
    Write-Host "`nFound existing export files from other computers:" -ForegroundColor Green
    $computer2Files | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
    
    $response = Read-Host "`nWould you like to compare with one of these files? (y/n)"
    if ($response -match '^[Yy]') {
        if ($computer2Files.Count -eq 1) {
            $computer2File = $computer2Files[0].FullName
        } else {
            Write-Host "`nSelect a file to compare with:"
            for ($i = 0; $i -lt $computer2Files.Count; $i++) {
                Write-Host "  $($i + 1). $($computer2Files[$i].Name)"
            }
            do {
                $selection = Read-Host "Enter number (1-$($computer2Files.Count))"
            } while (-not ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $computer2Files.Count))
            
            $computer2File = $computer2Files[[int]$selection - 1].FullName
        }
        
        # Perform comparison
        Write-Log "Step 3: Comparing applications between computers..." "INFO"
        
        $compareParams = @{
            Computer1File = $computer1File
            Computer2File = $computer2File
            OutputPath = Join-Path $WorkingDirectory "Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        }
        
        if ($GroupSimilarApps) { $compareParams.GroupSimilarApps = $true }
        if ($DetailedReport) { $compareParams.DetailedReport = $true }
        if ($ExportToJson) { $compareParams.ExportToJson = $true }
        $compareParams.ExportToMarkdown = $true  # Always generate Markdown reports
        
        try {
            & ".\Compare-InstalledApps.ps1" @compareParams
            Write-Log "Comparison completed successfully!" "SUCCESS"
        }
        catch {
            Write-Log "Error during comparison: $($_.Exception.Message)" "ERROR"
        }
    }
}

Write-Host @"

==========================================
  NEXT STEPS
==========================================

Current computer export completed: $computer1File

To complete the comparison:
1. Export apps from the second computer
2. Place the second computer's JSON file in: $WorkingDirectory
3. Run the comparison script

Or use the individual scripts:
- .\Export-InstalledApps.ps1 (for exporting)
- .\Compare-InstalledApps.ps1 (for comparing)

"@ -ForegroundColor Cyan
