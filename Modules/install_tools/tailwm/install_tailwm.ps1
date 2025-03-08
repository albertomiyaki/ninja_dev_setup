# TailWm setup

# This script creates a tailwm.bat file for tailing log files with a configurable tool
# and adds the batch file to the system PATH

# Configuration section - Edit these variables as needed
$tailToolName = "baretail"  # Options: baretail, glogg, or any other tool
$batchFilePath = "$env:USERPROFILE\Tools"  # Where to save the batch file
$logMappings = @{
    "myapp1" = "c:\program files\myapp1\log\server.log"
    "myapp2" = "c:\program files\myapp2\log\server.log"
    "myapp3" = "c:\program files\myapp3\log\server.log"
    # Add more mappings as needed
}

# Create directory if it doesn't exist
if (-not (Test-Path -Path $batchFilePath)) {
    New-Item -ItemType Directory -Path $batchFilePath -Force | Out-Null
    Write-Host "Created directory: $batchFilePath" -ForegroundColor Green
}

# Build batch file content
$batchContent = "@echo off`r`n"

# Add each app mapping
$sortedKeys = $logMappings.Keys | Sort-Object
$firstApp = $true

foreach ($app in $sortedKeys) {
    $logPath = $logMappings[$app]
    if ($firstApp) {
        $batchContent += "if ""%1""==""$app"" (`r`n"
        $firstApp = $false
    } else {
        $batchContent += ") else if ""%1""==""$app"" (`r`n"
    }
    $batchContent += "    $tailToolName ""$logPath""`r`n"
}

# Add help section
$batchContent += ") else (`r`n"
$batchContent += "    echo Unknown app: %1`r`n"
$batchContent += "    echo Available options: $($logMappings.Keys -join ', ')`r`n"
$batchContent += ")`r`n"

# Write batch file
$batchFile = Join-Path -Path $batchFilePath -ChildPath "tailmylog.bat"
$batchContent | Out-File -FilePath $batchFile -Encoding ASCII -Force

Write-Host "Created batch file: $batchFile" -ForegroundColor Green
Write-Host "Using tail tool: $tailToolName" -ForegroundColor Green

# Add to PATH if not already there
$currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($currentPath -notlike "*$batchFilePath*") {
    [Environment]::SetEnvironmentVariable("PATH", "$currentPath;$batchFilePath", "User")
    Write-Host "Added $batchFilePath to user PATH" -ForegroundColor Green
    Write-Host "Please restart your command prompt or PowerShell for PATH changes to take effect" -ForegroundColor Yellow
} else {
    Write-Host "$batchFilePath is already in PATH" -ForegroundColor Cyan
}

# Show usage information
Write-Host "`nSetup complete! You can now use:" -ForegroundColor Green
Write-Host "tailmylog <appname>" -ForegroundColor White
Write-Host "Where <appname> is one of: $($logMappings.Keys -join ', ')" -ForegroundColor White
Write-Host "`nTo modify log mappings or change the tail tool, edit this script and run it again." -ForegroundColor Cyan