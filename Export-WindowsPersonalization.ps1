# Export-WindowsPersonalization.ps1
# PowerShell script to export Windows personalization settings
# Includes desktop, taskbar, start menu, and visual preferences

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\WindowsPersonalization_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeWallpaper,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeStartLayout
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

function Get-DesktopSettings {
    Write-Log "Collecting desktop settings..."
    
    $desktopSettings = @{}
    
    try {
        # Desktop registry settings
        $desktopPath = "HKCU:\Control Panel\Desktop"
        if (Test-Path $desktopPath) {
            $desktop = Get-ItemProperty $desktopPath -ErrorAction SilentlyContinue
            if ($desktop) {
                $desktopSettings["Wallpaper"] = $desktop.Wallpaper
                $desktopSettings["WallpaperStyle"] = $desktop.WallpaperStyle
                $desktopSettings["TileWallpaper"] = $desktop.TileWallpaper
                $desktopSettings["Pattern"] = $desktop.Pattern
                $desktopSettings["ScreenSaveActive"] = $desktop.ScreenSaveActive
                $desktopSettings["ScreenSaveTimeOut"] = $desktop.ScreenSaveTimeOut
                $desktopSettings["SCRNSAVE.EXE"] = $desktop."SCRNSAVE.EXE"
                $desktopSettings["FontSmoothing"] = $desktop.FontSmoothing
                $desktopSettings["FontSmoothingType"] = $desktop.FontSmoothingType
                $desktopSettings["DragFullWindows"] = $desktop.DragFullWindows
                $desktopSettings["CursorBlinkRate"] = $desktop.CursorBlinkRate
                $desktopSettings["MenuShowDelay"] = $desktop.MenuShowDelay
            }
        }
        
        # Desktop icon settings
        $desktopIconsPath = "HKCU:\Software\Microsoft\Windows\Shell\Bags\1\Desktop"
        if (Test-Path $desktopIconsPath) {
            $desktopIcons = Get-ItemProperty $desktopIconsPath -ErrorAction SilentlyContinue
            if ($desktopIcons) {
                $desktopSettings["IconArrange"] = $desktopIcons.IconArrange
                $desktopSettings["Sort"] = $desktopIcons.Sort
            }
        }
        
        # Get current wallpaper info if requested
        if ($IncludeWallpaper) {
            try {
                Add-Type -TypeDefinition @"
                using System;
                using System.Runtime.InteropServices;
                public class Wallpaper {
                    [DllImport("user32.dll", CharSet = CharSet.Auto)]
                    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
                }
"@
                $wallpaperPath = " " * 260
                $null = [Wallpaper]::SystemParametersInfo(0x0073, $wallpaperPath.Length, $wallpaperPath, 0)
                $currentWallpaper = $wallpaperPath.Trim("`0")
                
                if ($currentWallpaper -and (Test-Path $currentWallpaper)) {
                    $wallpaperInfo = Get-ItemProperty $currentWallpaper -ErrorAction SilentlyContinue
                    $desktopSettings["CurrentWallpaperPath"] = $currentWallpaper
                    $desktopSettings["CurrentWallpaperSize"] = $wallpaperInfo.Length
                    $desktopSettings["CurrentWallpaperModified"] = $wallpaperInfo.LastWriteTime
                }
            }
            catch {
                Write-Log "Could not retrieve current wallpaper information: $($_.Exception.Message)" "WARN"
            }
        }
        
    }
    catch {
        Write-Log "Error collecting desktop settings: $($_.Exception.Message)" "WARN"
    }
    
    return $desktopSettings
}

function Get-TaskbarSettings {
    Write-Log "Collecting taskbar settings..."
    
    $taskbarSettings = @{}
    
    try {
        # Main taskbar settings
        $taskbarPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        if (Test-Path $taskbarPath) {
            $taskbar = Get-ItemProperty $taskbarPath -ErrorAction SilentlyContinue
            if ($taskbar) {
                $taskbarSettings["TaskbarGlomLevel"] = $taskbar.TaskbarGlomLevel
                $taskbarSettings["TaskbarSizeMove"] = $taskbar.TaskbarSizeMove
                $taskbarSettings["TaskbarSmallIcons"] = $taskbar.TaskbarSmallIcons
                $taskbarSettings["ShowTaskViewButton"] = $taskbar.ShowTaskViewButton
                $taskbarSettings["ShowCortanaButton"] = $taskbar.ShowCortanaButton
                $taskbarSettings["SearchboxTaskbarMode"] = $taskbar.SearchboxTaskbarMode
                $taskbarSettings["TaskbarMn"] = $taskbar.TaskbarMn
                $taskbarSettings["TaskbarDa"] = $taskbar.TaskbarDa
                $taskbarSettings["TaskbarAl"] = $taskbar.TaskbarAl
            }
        }
        
        # Taskbar position and auto-hide
        $stuckRectsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3"
        if (Test-Path $stuckRectsPath) {
            $stuckRects = Get-ItemProperty $stuckRectsPath -Name "Settings" -ErrorAction SilentlyContinue
            if ($stuckRects -and $stuckRects.Settings) {
                $taskbarSettings["StuckRectsSettings"] = [Convert]::ToBase64String($stuckRects.Settings)
            }
        }
        
        # System tray settings
        $trayPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
        if (Test-Path $trayPath) {
            $tray = Get-ItemProperty $trayPath -ErrorAction SilentlyContinue
            if ($tray) {
                $taskbarSettings["EnableAutoTray"] = $tray.EnableAutoTray
            }
        }
        
        # Notification area icons
        $notifyIconsPath = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify"
        if (Test-Path $notifyIconsPath) {
            $notifyIcons = Get-ItemProperty $notifyIconsPath -ErrorAction SilentlyContinue
            if ($notifyIcons) {
                if ($notifyIcons.IconStreams) {
                    $taskbarSettings["IconStreams"] = [Convert]::ToBase64String($notifyIcons.IconStreams)
                }
                if ($notifyIcons.PastIconsStream) {
                    $taskbarSettings["PastIconsStream"] = [Convert]::ToBase64String($notifyIcons.PastIconsStream)
                }
            }
        }
        
    }
    catch {
        Write-Log "Error collecting taskbar settings: $($_.Exception.Message)" "WARN"
    }
    
    return $taskbarSettings
}

function Get-StartMenuSettings {
    Write-Log "Collecting Start Menu settings..."
    
    $startMenuSettings = @{}
    
    try {
        # Start menu personalization
        $startPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        if (Test-Path $startPath) {
            $start = Get-ItemProperty $startPath -ErrorAction SilentlyContinue
            if ($start) {
                $startMenuSettings["Start_ShowOnDesktop"] = $start.Start_ShowOnDesktop
                $startMenuSettings["StartMenuInit"] = $start.StartMenuInit
                $startMenuSettings["Start_PowerButtonAction"] = $start.Start_PowerButtonAction
                $startMenuSettings["Start_ShowRun"] = $start.Start_ShowRun
                $startMenuSettings["Start_ShowMyComputer"] = $start.Start_ShowMyComputer
                $startMenuSettings["Start_ShowMyDocs"] = $start.Start_ShowMyDocs
                $startMenuSettings["Start_ShowControlPanel"] = $start.Start_ShowControlPanel
                $startMenuSettings["Start_ShowHelp"] = $start.Start_ShowHelp
                $startMenuSettings["Start_ShowSearch"] = $start.Start_ShowSearch
            }
        }
        
        # Start layout (if requested)
        if ($IncludeStartLayout) {
            try {
                $startLayoutPath = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml"
                if (Test-Path $startLayoutPath) {
                    $startLayout = Get-Content $startLayoutPath -Raw -Encoding UTF8
                    $startMenuSettings["CustomStartLayout"] = $startLayout
                }
                
                # Also check for default layout
                $defaultLayoutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
                if (Test-Path $defaultLayoutPath) {
                    $shortcuts = Get-ChildItem $defaultLayoutPath -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue
                    $startMenuShortcuts = @()
                    
                    foreach ($shortcut in $shortcuts) {
                        $shortcutInfo = @{
                            Name = $shortcut.Name
                            Path = $shortcut.FullName
                            RelativePath = $shortcut.FullName.Replace($defaultLayoutPath, "").TrimStart('\')
                            LastModified = $shortcut.LastWriteTime
                        }
                        $startMenuShortcuts += $shortcutInfo
                    }
                    
                    $startMenuSettings["StartMenuShortcuts"] = $startMenuShortcuts
                }
            }
            catch {
                Write-Log "Could not retrieve Start Menu layout: $($_.Exception.Message)" "WARN"
            }
        }
        
    }
    catch {
        Write-Log "Error collecting Start Menu settings: $($_.Exception.Message)" "WARN"
    }
    
    return $startMenuSettings
}

function Get-WindowsThemeSettings {
    Write-Log "Collecting Windows theme settings..."
    
    $themeSettings = @{}
    
    try {
        # Personalization settings
        $personalizePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        if (Test-Path $personalizePath) {
            $personalize = Get-ItemProperty $personalizePath -ErrorAction SilentlyContinue
            if ($personalize) {
                $themeSettings["AppsUseLightTheme"] = $personalize.AppsUseLightTheme
                $themeSettings["SystemUsesLightTheme"] = $personalize.SystemUsesLightTheme
                $themeSettings["EnableTransparency"] = $personalize.EnableTransparency
                $themeSettings["ColorPrevalence"] = $personalize.ColorPrevalence
            }
        }
        
        # Color settings
        $dwmPath = "HKCU:\Software\Microsoft\Windows\DWM"
        if (Test-Path $dwmPath) {
            $dwm = Get-ItemProperty $dwmPath -ErrorAction SilentlyContinue
            if ($dwm) {
                $themeSettings["AccentColor"] = $dwm.AccentColor
                $themeSettings["AccentColorInactive"] = $dwm.AccentColorInactive
                $themeSettings["ColorizationColor"] = $dwm.ColorizationColor
                $themeSettings["ColorizationColorBalance"] = $dwm.ColorizationColorBalance
                $themeSettings["ColorizationAfterglowBalance"] = $dwm.ColorizationAfterglowBalance
                $themeSettings["ColorizationBlurBalance"] = $dwm.ColorizationBlurBalance
                $themeSettings["ColorizationGlassAttribute"] = $dwm.ColorizationGlassAttribute
                $themeSettings["EnableAeroPeek"] = $dwm.EnableAeroPeek
            }
        }
        
        # Current theme
        $currentThemePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes"
        if (Test-Path $currentThemePath) {
            $currentTheme = Get-ItemProperty $currentThemePath -ErrorAction SilentlyContinue
            if ($currentTheme) {
                $themeSettings["CurrentTheme"] = $currentTheme.CurrentTheme
                $themeSettings["LastTheme"] = $currentTheme.LastTheme
            }
        }
        
    }
    catch {
        Write-Log "Error collecting theme settings: $($_.Exception.Message)" "WARN"
    }
    
    return $themeSettings
}

function Get-FileExplorerSettings {
    Write-Log "Collecting File Explorer settings..."
    
    $explorerSettings = @{}
    
    try {
        # Explorer advanced settings
        $advancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        if (Test-Path $advancedPath) {
            $advanced = Get-ItemProperty $advancedPath -ErrorAction SilentlyContinue
            if ($advanced) {
                $explorerSettings["Hidden"] = $advanced.Hidden
                $explorerSettings["HideFileExt"] = $advanced.HideFileExt
                $explorerSettings["ShowSuperHidden"] = $advanced.ShowSuperHidden
                $explorerSettings["LaunchTo"] = $advanced.LaunchTo
                $explorerSettings["SeparateProcess"] = $advanced.SeparateProcess
                $explorerSettings["NavPaneExpandToCurrentFolder"] = $advanced.NavPaneExpandToCurrentFolder
                $explorerSettings["NavPaneShowAllFolders"] = $advanced.NavPaneShowAllFolders
                $explorerSettings["ShowStatusBar"] = $advanced.ShowStatusBar
                $explorerSettings["ShowPreviewHandlers"] = $advanced.ShowPreviewHandlers
                $explorerSettings["AutoCheckSelect"] = $advanced.AutoCheckSelect
                $explorerSettings["FullRowSelect"] = $advanced.FullRowSelect
            }
        }
        
        # Folder view settings
        $viewStatePath = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU"
        if (Test-Path $viewStatePath) {
            $bagMRU = Get-ItemProperty $viewStatePath -ErrorAction SilentlyContinue
            if ($bagMRU) {
                $bagMRUSettings = @{}
                $bagMRU.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                    if ($_.Value -is [byte[]]) {
                        $bagMRUSettings[$_.Name] = [Convert]::ToBase64String($_.Value)
                    } else {
                        $bagMRUSettings[$_.Name] = $_.Value
                    }
                }
                $explorerSettings["BagMRU"] = $bagMRUSettings
            }
        }
        
    }
    catch {
        Write-Log "Error collecting File Explorer settings: $($_.Exception.Message)" "WARN"
    }
    
    return $explorerSettings
}

function Get-SoundSettings {
    Write-Log "Collecting sound settings..."
    
    $soundSettings = @{}
    
    try {
        # Sound scheme
        $soundPath = "HKCU:\AppEvents\Schemes"
        if (Test-Path $soundPath) {
            $sound = Get-ItemProperty $soundPath -ErrorAction SilentlyContinue
            if ($sound) {
                $soundSettings["CurrentSoundScheme"] = $sound."(Default)"
            }
        }
        
        # Individual sound events
        $appsPath = "HKCU:\AppEvents\Schemes\Apps"
        if (Test-Path $appsPath) {
            $soundEvents = @{}
            
            try {
                $apps = Get-ChildItem $appsPath -ErrorAction SilentlyContinue
                foreach ($app in $apps) {
                    $appName = Split-Path $app.Name -Leaf
                    $appSounds = @{}
                    
                    $events = Get-ChildItem $app.PSPath -ErrorAction SilentlyContinue
                    foreach ($soundEvent in $events) {
                        $eventName = Split-Path $soundEvent.Name -Leaf
                        $currentPath = "$($soundEvent.PSPath)\.Current"
                        
                        if (Test-Path $currentPath) {
                            $currentSound = Get-ItemProperty $currentPath -ErrorAction SilentlyContinue
                            if ($currentSound -and $currentSound."(Default)") {
                                $appSounds[$eventName] = $currentSound."(Default)"
                            }
                        }
                    }
                    
                    if ($appSounds.Count -gt 0) {
                        $soundEvents[$appName] = $appSounds
                    }
                }
                
                if ($soundEvents.Count -gt 0) {
                    $soundSettings["SoundEvents"] = $soundEvents
                }
            }
            catch {
                Write-Log "Error collecting sound events: $($_.Exception.Message)" "WARN"
            }
        }
        
    }
    catch {
        Write-Log "Error collecting sound settings: $($_.Exception.Message)" "WARN"
    }
    
    return $soundSettings
}

# Main execution
try {
    Write-Log "Starting Windows personalization export for computer: $ComputerName" "SUCCESS"
    Write-Log "Output file: $OutputPath"
    
    # Collect personalization settings
    $desktopSettings = Get-DesktopSettings
    $taskbarSettings = Get-TaskbarSettings
    $startMenuSettings = Get-StartMenuSettings
    $themeSettings = Get-WindowsThemeSettings
    $explorerSettings = Get-FileExplorerSettings
    $soundSettings = Get-SoundSettings
    
    # Create export object
    $exportData = [PSCustomObject]@{
        ComputerName = $ComputerName
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        DesktopSettings = $desktopSettings
        TaskbarSettings = $taskbarSettings
        StartMenuSettings = $startMenuSettings
        ThemeSettings = $themeSettings
        FileExplorerSettings = $explorerSettings
        SoundSettings = $soundSettings
        IncludeWallpaper = $IncludeWallpaper.IsPresent
        IncludeStartLayout = $IncludeStartLayout.IsPresent
        WindowsVersion = [System.Environment]::OSVersion.Version.ToString()
        WindowsBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
    }
    
    # Export to JSON
    $exportData | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "Successfully exported Windows personalization settings to: $OutputPath" "SUCCESS"
    Write-Log "Export completed for computer: $ComputerName" "SUCCESS"
    
    # Display summary
    Write-Host "`nSUMMARY:" -ForegroundColor Cyan
    Write-Host "Computer: $ComputerName" -ForegroundColor White
    Write-Host "Windows Version: $([System.Environment]::OSVersion.Version)" -ForegroundColor White
    Write-Host "Desktop Settings: $($desktopSettings.Count) properties" -ForegroundColor White
    Write-Host "Taskbar Settings: $($taskbarSettings.Count) properties" -ForegroundColor White
    Write-Host "Start Menu Settings: $($startMenuSettings.Count) properties" -ForegroundColor White
    Write-Host "Theme Settings: $($themeSettings.Count) properties" -ForegroundColor White
    Write-Host "File Explorer Settings: $($explorerSettings.Count) properties" -ForegroundColor White
    Write-Host "Sound Settings: $($soundSettings.Count) properties" -ForegroundColor White
    Write-Host "Output File: $OutputPath" -ForegroundColor White
    
}
catch {
    Write-Log "Critical error during Windows personalization export: $($_.Exception.Message)" "ERROR"
    exit 1
}
