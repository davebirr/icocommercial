# Application Comparison Toolkit

A PowerShell toolkit for comparing installed applications between two Windows computers. This solution helps identify software differences, version mismatches, and provides detailed reports for system administrators and IT professionals.

## Features

- **Comprehensive Application Detection**: Scans multiple sources (Registry, WMI, Package Managers, AppX packages)
- **Multiple Report Formats**: HTML (interactive), Markdown (documentation), and JSON (programmatic) reports
- **Directory Summaries**: Automatically generated overview reports for output directories
- **Remote Computer Support**: Connect to non-domain computers using local credentials
- **Version Difference Detection**: Identifies applications with different versions
- **Similar Application Grouping**: Groups related applications together
- **Flexible Filtering**: Options to include/exclude system components and updates
- **User-Friendly Interface**: Color-coded output and progress logging

## Files Included

| File | Description |
|------|-------------|
| `Export-InstalledApps.ps1` | Exports installed applications from a computer |
| `Compare-InstalledApps.ps1` | Compares application lists from two computers |
| `Start-AppComparison.ps1` | Simplified workflow script for complete comparison |
| `Export-RemoteInstalledApps.ps1` | **NEW**: Exports apps from remote computers using local credentials |
| `Start-NetworkAppComparison.ps1` | **NEW**: Network-wide comparison workflow with remote support |
| `New-MarkdownReport.ps1` | **NEW**: Generates professional Markdown reports |
| `New-DirectorySummary.ps1` | **NEW**: Creates comprehensive directory overview reports |
| `README.md` | This documentation file |

## Quick Start

### Method 1: Network-Wide Comparison (NEW - For Multiple Computers)
```powershell
.\Start-NetworkAppComparison.ps1
```
- **Perfect for non-domain computers**: Uses local username/password authentication
- **Multiple connection methods**: WinRM, file shares, or manual script copying
- **Network discovery**: Automatically scan your local network for computers
- **Batch processing**: Handle multiple computers in one workflow

### Method 2: Using the Workflow Script (Single Computer Pairs)
```powershell
.\Start-AppComparison.ps1
```

### Method 3: Remote Export (For Specific Remote Computers)
```powershell
# Export from multiple remote computers
.\Export-RemoteInstalledApps.ps1 -ComputerNames "PC1","PC2","PC3"
```

### Method 4: Manual Process (Original Method)

1. **Export applications from Computer 1:**
   ```powershell
   .\Export-InstalledApps.ps1 -OutputPath "Computer1_Apps.json" -ComputerName "Computer1"
   ```

2. **Export applications from Computer 2:**
   ```powershell
   .\Export-InstalledApps.ps1 -OutputPath "Computer2_Apps.json" -ComputerName "Computer2"
   ```

3. **Compare the exported data:**
   ```powershell
   .\Compare-InstalledApps.ps1 -Computer1File "Computer1_Apps.json" -Computer2File "Computer2_Apps.json"
   ```

## Script Parameters

### Export-RemoteInstalledApps.ps1 (NEW)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ComputerNames` | String[] | Required | Array of computer names to connect to |
| `Username` | String | Optional | Username for authentication (will prompt if not provided) |
| `Credential` | PSCredential | Optional | Pre-created credential object |
| `OutputDirectory` | String | .\RemoteExports | Directory for output files |
| `IncludeSystemComponents` | Switch | False | Include system components |
| `IncludeUpdates` | Switch | False | Include Windows updates |
| `UseWinRM` | Switch | False | Prefer WinRM over file share method |
| `TimeoutSeconds` | Int | 300 | Timeout for remote operations |

**Example:**
```powershell
.\Export-RemoteInstalledApps.ps1 -ComputerNames "WORKSTATION01","WORKSTATION02" -Username "Administrator" -UseWinRM
```

### Start-NetworkAppComparison.ps1 (NEW)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ComputerNames` | String[] | Empty | Pre-specified computer names (optional) |
| `Username` | String | Optional | Default username for all computers |
| `Credential` | PSCredential | Optional | Shared credential for all computers |
| `WorkingDirectory` | String | .\NetworkAppComparison | Working directory for files |
| `UseWinRM` | Switch | False | Prefer WinRM connections |

**Example:**
```powershell
.\Start-NetworkAppComparison.ps1 -ComputerNames "PC1","PC2","PC3" -UseWinRM
```

### Export-InstalledApps.ps1

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `OutputPath` | String | Auto-generated | Path for the output JSON file |
| `ComputerName` | String | Current computer | Name to identify the computer |
| `IncludeSystemComponents` | Switch | False | Include system components and runtime libraries |
| `IncludeUpdates` | Switch | False | Include Windows updates and hotfixes |
| `Verbose` | Switch | False | Show detailed progress information |

**Example:**
```powershell
.\Export-InstalledApps.ps1 -OutputPath "MyComputer.json" -ComputerName "WORKSTATION01" -IncludeSystemComponents
```

### Compare-InstalledApps.ps1

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Computer1File` | String | Required | Path to first computer's JSON export |
| `Computer2File` | String | Required | Path to second computer's JSON export |
| `OutputPath` | String | Auto-generated | Path for the HTML report |
| `ExportToJson` | Switch | False | Also export comparison results to JSON |
| `GroupSimilarApps` | Switch | False | Group applications with similar names |
| `ShowVersionDifferences` | Switch | False | Highlight version differences |
| `DetailedReport` | Switch | False | Include common applications in report |

**Example:**
```powershell
.\Compare-InstalledApps.ps1 -Computer1File "PC1.json" -Computer2File "PC2.json" -DetailedReport -ExportToJson
```

## Remote Computer Authentication

### For Non-Domain Computers (Workgroup)

The toolkit now supports connecting to computers that are **not joined to a domain** using local user accounts:

#### Authentication Methods:
1. **Local Administrator Account**: Use the built-in Administrator account
2. **Local User Account**: Any local user with appropriate permissions
3. **Same Credentials**: Use identical local accounts across multiple computers
4. **Per-Computer Credentials**: Prompt for different credentials for each computer

#### Connection Methods:
1. **WinRM** (Recommended if available):
   ```powershell
   # Enable WinRM on target computers (run as Administrator):
   Enable-PSRemoting -Force
   Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
   ```

2. **File Share + Script Copy** (Fallback method):
   - Uses administrative shares (C$, ADMIN$)
   - Copies and executes scripts remotely
   - Works when WinRM is not available

3. **Manual Export** (Most compatible):
   - Copy scripts to each computer manually
   - Run locally on each computer
   - Copy results back manually

#### Example Usage:

**Single Computer with Local Admin:**
```powershell
.\Export-RemoteInstalledApps.ps1 -ComputerNames "WORKSTATION01" -Username "Administrator"
```

**Multiple Computers with Shared Credentials:**
```powershell
$cred = Get-Credential -Message "Enter local admin credentials"
.\Export-RemoteInstalledApps.ps1 -ComputerNames "PC1","PC2","PC3" -Credential $cred -UseWinRM
```

**Network Discovery and Interactive Setup:**
```powershell
.\Start-NetworkAppComparison.ps1
# Follow prompts to:
# 1. Discover computers on local network
# 2. Choose authentication method
# 3. Select connection method
# 4. Automatically compare results
```

### Prerequisites for Remote Access:

1. **Network Connectivity**: Computers must be reachable via network
2. **Local Admin Rights**: User account must have local administrator privileges
3. **File Sharing**: For script copy method, administrative shares must be accessible
4. **WinRM** (Optional): For best performance, enable PowerShell remoting

### Security Considerations:

- Credentials are only used for the duration of the script execution
- No credentials are stored in files or logs
- Administrative shares are accessed temporarily
- Remote scripts are cleaned up after execution

## Understanding the Reports

### HTML Report Sections

1. **Summary Dashboard**: Overview statistics and computer information
2. **Applications Only on Computer 1**: Software missing from Computer 2
3. **Applications Only on Computer 2**: Software missing from Computer 1
4. **Version Differences**: Same applications with different versions
5. **Common Applications**: Identical software on both computers (if detailed report enabled)

### Color Coding

- 🔴 **Red**: Applications only on Computer 1
- 🔵 **Blue**: Applications only on Computer 2
- 🟡 **Yellow**: Applications with version differences
- 🟢 **Green**: Common applications (identical)

## Report Formats

### HTML Reports
- **Interactive visual reports** with charts and color coding
- **Best for**: Presentations, management reviews, immediate analysis
- **Features**: Responsive design, professional styling, summary dashboard

### Markdown Reports
- **Professional text-based reports** suitable for documentation
- **Best for**: Version control, documentation, sharing via text-based systems
- **Features**: Executive summaries, risk assessments, actionable recommendations
- **Includes**: Statistics, top publishers, standardization recommendations

### JSON Reports
- **Machine-readable data** for programmatic use
- **Best for**: Integration with other tools, automated processing
- **Features**: Complete raw data, metadata, structured format

### Directory Summary Reports
- **Comprehensive overview** of all analysis in a directory
- **Auto-generated** after network comparisons
- **Features**: File inventory, comparison matrix, next steps checklist

## Advanced Usage

### Including System Components

```powershell
.\Export-InstalledApps.ps1 -IncludeSystemComponents -IncludeUpdates
```

This includes:
- Microsoft Visual C++ Redistributables
- .NET Framework versions
- Windows SDK components
- System updates and hotfixes

### Grouping Similar Applications

```powershell
.\Compare-InstalledApps.ps1 -Computer1File "PC1.json" -Computer2File "PC2.json" -GroupSimilarApps
```

This groups applications like:
- "Adobe Acrobat Reader DC" and "Adobe Reader"
- "Microsoft Office Professional 2019" and "Microsoft Office Standard 2019"
- Different versions of the same software

### Batch Processing Multiple Computers

Create a batch script to process multiple computers:

```powershell
# Process multiple computers
$computers = @("PC1", "PC2", "PC3", "PC4")
$exportDir = "C:\AppExports"

foreach ($computer in $computers) {
    # Run on each computer (remotely or manually)
    Invoke-Command -ComputerName $computer -ScriptBlock {
        .\Export-InstalledApps.ps1 -OutputPath "\\SharePath\$using:computer.json" -ComputerName $using:computer
    }
}

# Compare all combinations
for ($i = 0; $i -lt $computers.Count; $i++) {
    for ($j = $i + 1; $j -lt $computers.Count; $j++) {
        $comp1 = $computers[$i]
        $comp2 = $computers[$j]
        .\Compare-InstalledApps.ps1 -Computer1File "$exportDir\$comp1.json" -Computer2File "$exportDir\$comp2.json" -OutputPath "Comparison_$comp1_vs_$comp2.html"
    }
}
```

## Troubleshooting

### Common Issues

1. **Execution Policy Error**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

2. **Access Denied Errors**
   - Run PowerShell as Administrator
   - Ensure proper permissions for registry access

3. **WMI Errors**
   - WMI service issues may prevent some applications from being detected
   - Registry-based detection will still work

4. **Large Output Files**
   - Use `-IncludeSystemComponents:$false` to reduce file size
   - Filter out updates with `-IncludeUpdates:$false`

### Performance Optimization

- **WMI queries can be slow**: The script uses registry as primary source for better performance
- **Large installations**: Consider filtering system components for faster processing
- **Network shares**: Use local paths when possible for better performance

## Data Sources

The scripts collect application data from multiple sources:

1. **Windows Registry**
   - `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*`
   - `HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*`

2. **WMI (Windows Management Instrumentation)**
   - `Win32_Product` class

3. **Package Managers**
   - `Get-Package` cmdlet
   - Windows Package Manager

4. **AppX Packages**
   - Windows Store applications
   - Universal Windows Platform (UWP) apps

## Output Format

### JSON Export Structure

```json
{
  "ComputerName": "WORKSTATION01",
  "ExportDate": "2025-01-15 14:30:25",
  "TotalApplications": 127,
  "IncludeSystemComponents": false,
  "IncludeUpdates": false,
  "Applications": [
    {
      "Name": "Mozilla Firefox",
      "Version": "121.0.1",
      "Publisher": "Mozilla Corporation",
      "InstallDate": "20240115",
      "UninstallString": "...",
      "InstallLocation": "C:\\Program Files\\Mozilla Firefox",
      "Size": 102400,
      "Source": "Registry",
      "Architecture": "x64"
    }
  ]
}
```

## Security Considerations

- Scripts only read system information, no modifications are made
- No sensitive data is collected (passwords, user data, etc.)
- Registry access requires appropriate permissions
- Consider data privacy when sharing export files

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Windows 10/11 or Windows Server 2016+
- Administrator privileges recommended for complete application detection
- Network access if comparing remote computers

## Support

For issues or feature requests:
1. Check the troubleshooting section above
2. Verify PowerShell execution policy
3. Ensure proper permissions and administrator access
4. Review the verbose output for specific error messages

## Version History

- **v1.0**: Initial release with basic comparison functionality
- **v1.1**: Added similar application grouping
- **v1.2**: Enhanced HTML reports and multiple data sources
- **v1.3**: Added workflow script and batch processing support

---

*This toolkit is designed for IT professionals and system administrators to efficiently manage and compare software installations across multiple Windows computers.*
