# Edge Backup & Restore Tool

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.EdgeProfilePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default"
$sync.SelectedBackupOptions = @()
$sync.SelectedRestoreOptions = @()
$sync.BackupItems = @{
    "Favorites" = @{
        Path = "Bookmarks"
        Description = "Edge bookmarks and favorites"
    }
    "SearchEngines" = @{
        Path = "Web Data"
        Description = "Custom search engines"
    }
    "Settings" = @{
        Path = "Preferences"
        Description = "Browser settings and preferences"
    }
    "Extensions" = @{
        Path = "Extensions"
        Description = "Installed browser extensions"
        IsFolder = $true
    }
    "AutofillData" = @{
        Path = "Web Data"
        Description = "Saved form data (addresses, etc.)"
    }
    "Passwords" = @{
        Path = "Login Data"
        Description = "Saved passwords"
    }
}

# Create log function for centralized logging
function Write-ActivityLog {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Information", "Warning", "Error", "Success")]
        [string]$Type = "Information"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # Add to the log text box if UI is loaded
    if ($sync.LogTextBox -ne $null) {
        $sync.LogTextBox.Dispatcher.Invoke([action]{
            $sync.LogTextBox.AppendText("$logMessage`r`n")
            $sync.LogTextBox.ScrollToEnd()
        })
    }
    
    # Also write to console for debugging purposes
    switch ($Type) {
        "Information" { Write-Host $logMessage -ForegroundColor Gray }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
    }
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Edge Settings Manager" 
    Height="900" 
    Width="900"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button" x:Key="DefaultButton">
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#003C6A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#CCCCCC"/>
                                <Setter Property="Foreground" Value="#666666"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="15,5"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,8,0,8"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="Edge Settings Manager" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Backup and restore your Microsoft Edge browser settings" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="15" x:Name="MainTabControl">
            <!-- Backup Tab -->
            <TabItem Header="Backup Settings">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Select items to backup:" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    
                    <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3" Padding="15" Margin="0,0,0,15">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="BackupOptionsPanel">
                                <!-- Checkboxes will be added here programmatically -->
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="SelectAllBackupButton" Content="Select All" 
                                Style="{StaticResource DefaultButton}" 
                                Width="100"/>
                        <Button x:Name="ClearAllBackupButton" Content="Clear All" 
                                Style="{StaticResource DefaultButton}" 
                                Width="100"/>
                        <Button x:Name="BackupButton" Content="Create Backup" 
                                Style="{StaticResource DefaultButton}" 
                                Width="150"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <!-- Restore Tab -->
            <TabItem Header="Restore Settings">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Select backup file to restore from:" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    
                    <Grid Grid.Row="1" Margin="0,0,0,15">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="RestoreFilePathTextBox" IsReadOnly="True" Padding="8" BorderThickness="1" BorderBrush="#CCCCCC"/>
                        <Button Grid.Column="1" x:Name="BrowseRestoreButton" Content="Browse..." 
                                Style="{StaticResource DefaultButton}" 
                                Width="100"/>
                    </Grid>
                    
                    <Border Grid.Row="2" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3" Padding="15" Margin="0,0,0,15">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="RestoreOptionsPanel">
                                <TextBlock Text="Select backup file to view available restore options" 
                                           Foreground="#666666" 
                                           HorizontalAlignment="Center" 
                                           VerticalAlignment="Center" 
                                           FontStyle="Italic"
                                           Margin="0,20,0,0"/>
                                <!-- Checkboxes will be added here programmatically -->
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="SelectAllRestoreButton" Content="Select All" 
                                Style="{StaticResource DefaultButton}" 
                                Width="100" IsEnabled="False"/>
                        <Button x:Name="ClearAllRestoreButton" Content="Clear All" 
                                Style="{StaticResource DefaultButton}" 
                                Width="100" IsEnabled="False"/>
                        <Button x:Name="RestoreButton" Content="Restore Selected" 
                                Style="{StaticResource DefaultButton}" 
                                Width="150" IsEnabled="False"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Log Panel -->
        <Grid Grid.Row="2" Margin="15,0,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="120"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="Activity Log" FontWeight="Bold" FontSize="14" Margin="0,0,0,5"/>
            <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBox x:Name="LogTextBox" IsReadOnly="True" TextWrapping="Wrap" 
                             FontFamily="Consolas" Background="#F8F8F8" BorderThickness="0"
                             Padding="10"/>
                </ScrollViewer>
            </Border>
        </Grid>
        
        <!-- Status Bar -->
        <Border Grid.Row="3" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
            <TextBlock x:Name="StatusBar" Text="Ready" FontSize="12"/>
        </Border>
    </Grid>
</Window>
"@

# Create a form object from the XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Store references in the sync hashtable
$sync.Window = $window
$sync.MainTabControl = $window.FindName("MainTabControl")
$sync.BackupOptionsPanel = $window.FindName("BackupOptionsPanel")
$sync.RestoreOptionsPanel = $window.FindName("RestoreOptionsPanel")
$sync.SelectAllBackupButton = $window.FindName("SelectAllBackupButton")
$sync.ClearAllBackupButton = $window.FindName("ClearAllBackupButton")
$sync.BackupButton = $window.FindName("BackupButton")
$sync.RestoreFilePathTextBox = $window.FindName("RestoreFilePathTextBox")
$sync.BrowseRestoreButton = $window.FindName("BrowseRestoreButton")
$sync.SelectAllRestoreButton = $window.FindName("SelectAllRestoreButton")
$sync.ClearAllRestoreButton = $window.FindName("ClearAllRestoreButton")
$sync.RestoreButton = $window.FindName("RestoreButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")

# Function to populate backup options
function Initialize-BackupOptions {
    $sync.BackupOptionsPanel.Children.Clear()
    $sync.SelectedBackupOptions = @()
    
    foreach ($key in $sync.BackupItems.Keys) {
        $item = $sync.BackupItems[$key]
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = "$key - $($item.Description)"
        $checkBox.Tag = $key
        $checkBox.IsChecked = $false
        
        # Add event handler for checkbox
        $checkBox.Add_Checked({
            $option = $this.Tag
            if ($sync.SelectedBackupOptions -notcontains $option) {
                $sync.SelectedBackupOptions += $option
            }
            Write-ActivityLog "$option selected for backup" -Type Information
        })
        
        $checkBox.Add_Unchecked({
            $option = $this.Tag
            $sync.SelectedBackupOptions = $sync.SelectedBackupOptions | Where-Object { $_ -ne $option }
            Write-ActivityLog "$option deselected from backup" -Type Information
        })
        
        $sync.BackupOptionsPanel.Children.Add($checkBox)
    }
}

# Function to create backup
function New-EdgeBackup {
    param (
        [string]$BackupPath,
        [array]$Options
    )
    
    if ($Options.Count -eq 0) {
        Write-ActivityLog "No backup options selected" -Type Warning
        [System.Windows.MessageBox]::Show(
            "Please select at least one item to backup",
            "Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    try {
        # Create a temp directory for backup files
        $tempDir = Join-Path $env:TEMP "EdgeBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        Write-ActivityLog "Created temporary directory for backup: $tempDir" -Type Information
        
        # Copy selected items to temp directory
        foreach ($option in $Options) {
            $item = $sync.BackupItems[$option]
            $sourcePath = Join-Path $sync.EdgeProfilePath $item.Path
            $destPath = Join-Path $tempDir $option
            
            if (!(Test-Path -Path $sourcePath)) {
                Write-ActivityLog "Source path not found: $sourcePath" -Type Warning
                continue
            }
            
            if ($item.IsFolder -eq $true) {
                # Copy entire folder
                New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                Copy-Item -Path "$sourcePath\*" -Destination $destPath -Recurse -Force
            } else {
                # Copy single file
                Copy-Item -Path $sourcePath -Destination $destPath -Force
            }
            
            Write-ActivityLog "Backed up $option" -Type Success
        }
        
        # Create a metadata file
        $metadata = @{
            CreatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Options = $Options
            EdgeVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" -ErrorAction SilentlyContinue).Version
        }
        $metadata | ConvertTo-Json | Out-File (Join-Path $tempDir "metadata.json") -Force
        
        # Create the ZIP file
        if (Test-Path $BackupPath) {
            Remove-Item -Path $BackupPath -Force
        }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $BackupPath)
        Write-ActivityLog "Backup saved to: $BackupPath" -Type Success
        
        # Clean up temp directory
        Remove-Item -Path $tempDir -Recurse -Force
        Write-ActivityLog "Temporary files cleaned up" -Type Information
        
        $sync.StatusBar.Text = "Backup completed successfully"
        
        [System.Windows.MessageBox]::Show(
            "Backup completed successfully!`nFile saved to: $BackupPath",
            "Backup Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    } catch {
        Write-ActivityLog "Error creating backup: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "An error occurred while creating the backup:`n$_",
            "Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Function to initialize restore options from backup file
function Initialize-RestoreOptions {
    param (
        [string]$BackupPath
    )
    
    try {
        $sync.RestoreOptionsPanel.Children.Clear()
        $sync.SelectedRestoreOptions = @()
        
        # Extract metadata from the zip file
        $tempDir = Join-Path $env:TEMP "EdgeRestore_Temp"
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Extract just the metadata file
        $zip = [System.IO.Compression.ZipFile]::OpenRead($BackupPath)
        $metadataEntry = $zip.Entries | Where-Object { $_.Name -eq "metadata.json" }
        
        if ($null -eq $metadataEntry) {
            Write-ActivityLog "Invalid backup file: Metadata not found" -Type Error
            $zip.Dispose()
            return
        }
        
        $metadataPath = Join-Path $tempDir "metadata.json"
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($metadataEntry, $metadataPath, $true)
        $zip.Dispose()
        
        # Read metadata
        $metadata = Get-Content $metadataPath -Raw | ConvertFrom-Json
        
        # Create text block with backup info
        $infoBlock = New-Object System.Windows.Controls.TextBlock
        $infoBlock.Margin = "0,0,0,15"
        $infoBlock.TextWrapping = "Wrap"
        $infoBlock.Text = "Backup created on: $($metadata.CreatedOn)`nEdge version: $($metadata.EdgeVersion)"
        $sync.RestoreOptionsPanel.Children.Add($infoBlock)
        
        # Add separator
        $separator = New-Object System.Windows.Controls.Separator
        $separator.Margin = "0,0,0,15"
        $sync.RestoreOptionsPanel.Children.Add($separator)
        
        # Add options label
        $optionsLabel = New-Object System.Windows.Controls.TextBlock
        $optionsLabel.Text = "Select items to restore:"
        $optionsLabel.FontWeight = "SemiBold"
        $optionsLabel.FontSize = 14
        $optionsLabel.Margin = "0,0,0,10"
        $sync.RestoreOptionsPanel.Children.Add($optionsLabel)
        
        # Add checkboxes for each available item
        foreach ($option in $metadata.Options) {
            $item = $sync.BackupItems[$option]
            $checkBox = New-Object System.Windows.Controls.CheckBox
            $checkBox.Content = "$option - $($item.Description)"
            $checkBox.Tag = $option
            $checkBox.IsChecked = $false
            
            # Add event handler for checkbox
            $checkBox.Add_Checked({
                $option = $this.Tag
                if ($sync.SelectedRestoreOptions -notcontains $option) {
                    $sync.SelectedRestoreOptions += $option
                }
                Write-ActivityLog "$option selected for restore" -Type Information
            })
            
            $checkBox.Add_Unchecked({
                $option = $this.Tag
                $sync.SelectedRestoreOptions = $sync.SelectedRestoreOptions | Where-Object { $_ -ne $option }
                Write-ActivityLog "$option deselected from restore" -Type Information
            })
            
            $sync.RestoreOptionsPanel.Children.Add($checkBox)
        }
        
        # Enable restore buttons
        $sync.SelectAllRestoreButton.IsEnabled = $true
        $sync.ClearAllRestoreButton.IsEnabled = $true
        $sync.RestoreButton.IsEnabled = $true
        
        # Clean up
        Remove-Item -Path $tempDir -Recurse -Force
        
        Write-ActivityLog "Loaded restore options from backup file" -Type Success
    } catch {
        Write-ActivityLog "Error loading restore options: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "An error occurred while loading the backup file:`n$_",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Function to restore from backup
function Restore-EdgeSettings {
    param (
        [string]$BackupPath,
        [array]$Options
    )
    
    if ($Options.Count -eq 0) {
        Write-ActivityLog "No restore options selected" -Type Warning
        [System.Windows.MessageBox]::Show(
            "Please select at least one item to restore",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    try {
        # Check if Edge is running
        $edgeProcess = Get-Process -Name "msedge" -ErrorAction SilentlyContinue
        if ($edgeProcess) {
            $result = [System.Windows.MessageBox]::Show(
                "Microsoft Edge is currently running. It must be closed before restoring settings.`n`nWould you like to close Edge now?",
                "Close Edge",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-ActivityLog "Closing Microsoft Edge" -Type Information
                $edgeProcess | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep -Seconds 2
                
                $edgeProcess = Get-Process -Name "msedge" -ErrorAction SilentlyContinue
                if ($edgeProcess) {
                    Write-ActivityLog "Forcing Edge to close" -Type Warning
                    $edgeProcess | Stop-Process -Force
                    # Allow time for system to release file locks
                    Start-Sleep -Seconds 5
                }
            } else {
                Write-ActivityLog "Restore cancelled - Edge must be closed first" -Type Warning
                return
            }
        }
        
        # Add additional check for msedge.exe processes that might be running in background
        $edgeProcesses = Get-Process | Where-Object { $_.Name -like "*edge*" } | Where-Object { $_.Name -ne "EdgeSettingsManager" }
        if ($edgeProcesses) {
            Write-ActivityLog "Found additional Edge-related processes, attempting to close them" -Type Warning
            $edgeProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
        }
        
        # Extract files from the backup
        $tempDir = Join-Path $env:TEMP "EdgeRestore_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($BackupPath, $tempDir)
        Write-ActivityLog "Extracted backup to temporary location" -Type Information
        
                    # Restore each selected item
        foreach ($option in $Options) {
            $item = $sync.BackupItems[$option]
            $sourcePath = Join-Path $tempDir $option
            $destPath = Join-Path $sync.EdgeProfilePath $item.Path
            
            if (!(Test-Path -Path $sourcePath)) {
                Write-ActivityLog "Source path not found in backup: $sourcePath" -Type Warning
                continue
            }
            
            # Create a backup of the current file/folder
            $backupPath = "$destPath.bak"
            if (Test-Path $backupPath) {
                Remove-Item -Path $backupPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            # Handle file locks by trying multiple approaches
            try {
                if (Test-Path $destPath) {
                    if ($item.IsFolder -eq $true) {
                        Copy-Item -Path $destPath -Destination $backupPath -Recurse -Force -ErrorAction SilentlyContinue
                        # Use robocopy for folders to handle locked files
                        Write-ActivityLog "Using robocopy to handle folder restore for $option" -Type Information
                        $null = robocopy "$sourcePath" "$destPath" /E /IS /IT /IM /R:1 /W:1 /NP /NFL /NDL
                    } else {
                        # Try to handle locked files
                        try {
                            Copy-Item -Path $destPath -Destination $backupPath -Force -ErrorAction Stop
                            Copy-Item -Path $sourcePath -Destination $destPath -Force -ErrorAction Stop
                        } catch {
                            Write-ActivityLog "Using alternative method for locked file: $destPath" -Type Warning
                            # Create a temporary copy
                            $tempFile = "$destPath.new"
                            Copy-Item -Path $sourcePath -Destination $tempFile -Force
                            
                            # Register a scheduled task to replace the file after a short delay
                            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"
                                Start-Sleep -Seconds 3;
                                Remove-Item -Path '$destPath' -Force;
                                Rename-Item -Path '$tempFile' -NewName '$(Split-Path $destPath -Leaf)' -Force;
                                Unregister-ScheduledTask -TaskName 'EdgeRestore_FileCopy' -Confirm:`$false
                            `""
                            
                            $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(1)
                            $principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -RunLevel Highest
                            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
                            
                            # Register and start the task
                            $task = Register-ScheduledTask -TaskName "EdgeRestore_FileCopy" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
                            
                            Write-ActivityLog "Scheduled file replacement for $option" -Type Information
                        }
                    }
                } else {
                    if ($item.IsFolder -eq $true) {
                        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                        Copy-Item -Path "$sourcePath\*" -Destination $destPath -Recurse -Force
                    } else {
                        Copy-Item -Path $sourcePath -Destination $destPath -Force
                    }
                }
            } catch {
                Write-ActivityLog "Error restoring $option`: $($_ -replace ':', '')" -Type Error
                continue
            }
            
            Write-ActivityLog "Restored $option" -Type Success
        }
        
        # Clean up temp directory
        Remove-Item -Path $tempDir -Recurse -Force
        Write-ActivityLog "Temporary files cleaned up" -Type Information
        
        $sync.StatusBar.Text = "Restore completed successfully"
        
        [System.Windows.MessageBox]::Show(
            "Restore completed successfully!`n`nYou may need to restart Microsoft Edge for all changes to take effect.",
            "Restore Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    } catch {
        Write-ActivityLog "Error restoring settings: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "An error occurred while restoring settings:`n$_",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Set up event handlers
$sync.SelectAllBackupButton.Add_Click({
    foreach ($child in $sync.BackupOptionsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox]) {
            $child.IsChecked = $true
        }
    }
})

$sync.ClearAllBackupButton.Add_Click({
    foreach ($child in $sync.BackupOptionsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox]) {
            $child.IsChecked = $false
        }
    }
})

$sync.BackupButton.Add_Click({
    if ($sync.SelectedBackupOptions.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select at least one item to backup",
            "Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    # Show save file dialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Zip Files (*.zip)|*.zip"
    $saveFileDialog.Title = "Save Edge Settings Backup"
    $saveFileDialog.FileName = "EdgeBackup_$(Get-Date -Format 'yyyyMMdd').zip"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.StatusBar.Text = "Creating backup..."
        Write-ActivityLog "Creating backup to $($saveFileDialog.FileName)" -Type Information
        New-EdgeBackup -BackupPath $saveFileDialog.FileName -Options $sync.SelectedBackupOptions
    }
})

$sync.BrowseRestoreButton.Add_Click({
    # Show open file dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Zip Files (*.zip)|*.zip"
    $openFileDialog.Title = "Select Edge Settings Backup"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.RestoreFilePathTextBox.Text = $openFileDialog.FileName
        $sync.StatusBar.Text = "Loading backup file..."
        Write-ActivityLog "Loading backup file: $($openFileDialog.FileName)" -Type Information
        Initialize-RestoreOptions -BackupPath $openFileDialog.FileName
    }
})

$sync.SelectAllRestoreButton.Add_Click({
    foreach ($child in $sync.RestoreOptionsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox]) {
            $child.IsChecked = $true
        }
    }
})

$sync.ClearAllRestoreButton.Add_Click({
    foreach ($child in $sync.RestoreOptionsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox]) {
            $child.IsChecked = $false
        }
    }
})

$sync.RestoreButton.Add_Click({
    if ($sync.SelectedRestoreOptions.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select at least one item to restore",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "This will replace your current Edge settings with the ones from the backup.`n`nDo you want to continue?",
        "Confirm Restore",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $sync.StatusBar.Text = "Restoring settings..."
        Write-ActivityLog "Restoring settings from $($sync.RestoreFilePathTextBox.Text)" -Type Information
        Restore-EdgeSettings -BackupPath $sync.RestoreFilePathTextBox.Text -Options $sync.SelectedRestoreOptions
    }
})

# Initialize the application
Initialize-BackupOptions

# Start with welcome message
Write-ActivityLog "Edge Settings Manager started." -Type Information
$sync.StatusBar.Text = "Ready"

# Show the window
$sync.Window.ShowDialog() | Out-Null