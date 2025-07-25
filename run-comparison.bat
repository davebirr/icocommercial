@echo off
REM run-comparison.bat
REM Simple batch file to run the Application Comparison Toolkit

echo ==========================================
echo   APPLICATION COMPARISON TOOLKIT
echo ==========================================
echo.

echo Choose an option:
echo 1. Export applications from this computer
echo 2. Start guided comparison workflow (local)
echo 3. Start network comparison workflow (remote computers)
echo 4. Export from remote computers
echo 5. Show usage examples
echo 6. Export with system components included
echo 7. Exit
echo.

set /p choice="Enter your choice (1-7): "

if "%choice%"=="1" (
    echo.
    echo Exporting applications from this computer...
    powershell.exe -ExecutionPolicy Bypass -File Export-InstalledApps.ps1
    pause
) else if "%choice%"=="2" (
    echo.
    echo Starting guided comparison workflow...
    powershell.exe -ExecutionPolicy Bypass -File Start-AppComparison.ps1
    pause
) else if "%choice%"=="3" (
    echo.
    echo Starting network comparison workflow...
    powershell.exe -ExecutionPolicy Bypass -File Start-NetworkAppComparison.ps1
    pause
) else if "%choice%"=="4" (
    echo.
    echo Starting remote computer export...
    powershell.exe -ExecutionPolicy Bypass -File Export-RemoteInstalledApps.ps1
    pause
) else if "%choice%"=="5" (
    echo.
    echo Showing usage examples...
    powershell.exe -ExecutionPolicy Bypass -File Example-Usage.ps1
    pause
) else if "%choice%"=="6" (
    echo.
    echo Exporting applications with system components...
    powershell.exe -ExecutionPolicy Bypass -File Export-InstalledApps.ps1 -IncludeSystemComponents -IncludeUpdates
    pause
) else if "%choice%"=="7" (
    echo.
    echo Goodbye!
    exit /b 0
) else (
    echo.
    echo Invalid choice. Please run the script again and choose 1-7.
    pause
)

echo.
echo Script completed. Check the generated files in this directory.
pause
