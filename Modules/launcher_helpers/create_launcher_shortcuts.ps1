# Create ninjaDEV Launcher shortcuts
$currentScriptPath = $MyInvocation.MyCommand.Path
$helperFolder = Split-Path -Parent $currentScriptPath
$modulesFolder = Split-Path -Parent $helperFolder
$mainFolder = Split-Path -Parent $modulesFolder
$scriptPath = Join-Path -Path $mainFolder -ChildPath "launcher.ps1"
$scriptName = "ninjaDEV Launcher"

# Define custom icon path
$iconFolder = Join-Path -Path $helperFolder -ChildPath "icons"
$iconPath = Join-Path -Path $iconFolder -ChildPath "icon.ico"

# Prompt user for shortcut locations
$createDesktop = Read-Host "Create desktop shortcut? (Y/N)"
$createStartMenu = Read-Host "Create Start Menu shortcut? (Y/N)"

# Create WshShell COM object for shortcut creation
$WshShell = New-Object -ComObject WScript.Shell

# Function to create the shortcut
function Create-Shortcut {
    param (
        [string]$shortcutPath
    )
    
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = "powershell.exe"
    $Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    
    # Use custom icon if it exists, otherwise fall back to PowerShell icon
    if (Test-Path $iconPath) {
        $Shortcut.IconLocation = $iconPath
        Write-Host "Using custom icon from: $iconPath" -ForegroundColor Cyan
    } else {
        $Shortcut.IconLocation = "powershell.exe,0"
        Write-Host "Custom icon not found at: $iconPath" -ForegroundColor Yellow
        Write-Host "Using default PowerShell icon instead" -ForegroundColor Yellow
    }
    
    # Optional properties
    $Shortcut.Description = "Shortcut to $scriptName"
    $Shortcut.WindowStyle = 1 # 1 = Normal, 3 = Maximized, 7 = Minimized
    $Shortcut.Save()
    
    Write-Host "Created shortcut at: $shortcutPath" -ForegroundColor Green
}

# Create desktop shortcut if requested
if ($createDesktop -eq "Y" -or $createDesktop -eq "y") {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $desktopShortcut = "$desktopPath\$scriptName.lnk"
    Create-Shortcut -shortcutPath $desktopShortcut
}

# Create Start Menu shortcut if requested
if ($createStartMenu -eq "Y" -or $createStartMenu -eq "y") {
    # Create folder in Start Menu Programs
    $startMenuFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\ninjaDEV Launcher"
    
    # Create the folder if it doesn't exist
    if (!(Test-Path $startMenuFolder)) {
        New-Item -Path $startMenuFolder -ItemType Directory | Out-Null
        Write-Host "Created folder: $startMenuFolder" -ForegroundColor Green
    }
    
    # Create the shortcut in the Start Menu folder
    $startMenuShortcut = "$startMenuFolder\$scriptName.lnk"
    Create-Shortcut -shortcutPath $startMenuShortcut
}

if ($createDesktop -ne "Y" -and $createDesktop -ne "y" -and $createStartMenu -ne "Y" -and $createStartMenu -ne "y") {
    Write-Host "No shortcuts were created." -ForegroundColor Yellow
}