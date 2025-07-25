# Example-Usage.ps1
# Example script demonstrating how to use the Application Comparison Toolkit

Write-Host @"
==========================================
  APPLICATION COMPARISON TOOLKIT
  USAGE EXAMPLES
==========================================
"@ -ForegroundColor Cyan

Write-Host "`n1. BASIC USAGE - Export apps from current computer:" -ForegroundColor Yellow
Write-Host "   .\Export-InstalledApps.ps1" -ForegroundColor White

Write-Host "`n2. EXPORT WITH CUSTOM NAME:" -ForegroundColor Yellow
Write-Host "   .\Export-InstalledApps.ps1 -OutputPath 'MyComputer.json' -ComputerName 'WORKSTATION01'" -ForegroundColor White

Write-Host "`n3. INCLUDE SYSTEM COMPONENTS AND UPDATES:" -ForegroundColor Yellow
Write-Host "   .\Export-InstalledApps.ps1 -IncludeSystemComponents -IncludeUpdates" -ForegroundColor White

Write-Host "`n4. COMPARE TWO COMPUTERS:" -ForegroundColor Yellow
Write-Host "   .\Compare-InstalledApps.ps1 -Computer1File 'PC1.json' -Computer2File 'PC2.json'" -ForegroundColor White

Write-Host "`n5. DETAILED COMPARISON WITH MARKDOWN AND GROUPING:" -ForegroundColor Yellow
Write-Host "   .\Compare-InstalledApps.ps1 -Computer1File 'PC1.json' -Computer2File 'PC2.json' -DetailedReport -GroupSimilarApps -ExportToJson -ExportToMarkdown" -ForegroundColor White

Write-Host "`n6. SIMPLIFIED WORKFLOW (RECOMMENDED FOR BEGINNERS):" -ForegroundColor Yellow
Write-Host "   .\Start-AppComparison.ps1" -ForegroundColor White

Write-Host "`n7. NETWORK-WIDE COMPARISON (NEW - FOR MULTIPLE COMPUTERS):" -ForegroundColor Yellow
Write-Host "   .\Start-NetworkAppComparison.ps1" -ForegroundColor White

Write-Host "`n8. EXPORT FROM REMOTE COMPUTERS:" -ForegroundColor Yellow
Write-Host "   .\Export-RemoteInstalledApps.ps1 -ComputerNames 'PC1','PC2','PC3' -UseWinRM" -ForegroundColor White

Write-Host "`n9. GENERATE DIRECTORY SUMMARY:" -ForegroundColor Yellow
Write-Host "   .\New-DirectorySummary.ps1 -DirectoryPath '.\AppComparison'" -ForegroundColor White

Write-Host "`n10. BATCH PROCESSING EXAMPLE:" -ForegroundColor Yellow
Write-Host @"
   # Export from multiple computers
   foreach (`$computer in @('PC1', 'PC2', 'PC3')) {
       .\Export-InstalledApps.ps1 -OutputPath "`$computer.json" -ComputerName `$computer
   }
   
   # Compare PC1 with all others
   .\Compare-InstalledApps.ps1 -Computer1File 'PC1.json' -Computer2File 'PC2.json'
   .\Compare-InstalledApps.ps1 -Computer1File 'PC1.json' -Computer2File 'PC3.json'
"@ -ForegroundColor White

Write-Host "`n11. TROUBLESHOOTING - Set execution policy if needed:" -ForegroundColor Yellow
Write-Host "   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor White

Write-Host @"

==========================================
  TYPICAL WORKFLOW
==========================================

Step 1: Export from Computer 1
   .\Export-InstalledApps.ps1 -ComputerName "Computer1"

Step 2: Copy script to Computer 2 and run
   .\Export-InstalledApps.ps1 -ComputerName "Computer2"

Step 3: Copy both JSON files to one location

Step 4: Compare the files
   .\Compare-InstalledApps.ps1 -Computer1File "Computer1_Apps.json" -Computer2File "Computer2_Apps.json"

Step 5: Open the generated HTML report in your browser

==========================================
"@ -ForegroundColor Cyan

# Prompt to run a quick demo
$response = Read-Host "`nWould you like to run a quick export of this computer's applications? (y/n)"
if ($response -match '^[Yy]') {
    Write-Host "`nRunning export demo..." -ForegroundColor Green
    if (Test-Path ".\Export-InstalledApps.ps1") {
        .\Export-InstalledApps.ps1 -ComputerName "Demo_$env:COMPUTERNAME"
        Write-Host "Demo export completed! Check the generated JSON file." -ForegroundColor Green
    } else {
        Write-Host "Export-InstalledApps.ps1 not found in current directory." -ForegroundColor Red
    }
} else {
    Write-Host "`nTo get started, run: .\Start-AppComparison.ps1" -ForegroundColor Cyan
}
