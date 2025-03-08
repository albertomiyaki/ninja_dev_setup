# Install Extra Tools

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Check if the script is running with administrator privileges
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Self-elevate the script if required
if (-not (Test-Administrator)) {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    # Build arguments to pass to the elevated process
    $argString = ""
    if ($MyInvocation.BoundParameters.Count -gt 0) {
        $argString = "-ExecutionPolicy Bypass -File `"$scriptPath`""
        foreach ($key in $MyInvocation.BoundParameters.Keys) {
            $value = $MyInvocation.BoundParameters[$key]
            if ($value -is [System.String]) {
                $argString += " -$key `"$value`""
            } else {
                $argString += " -$key $value"
            }
        }
    } else {
        $argString = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    }
    
    # Start PowerShell as administrator with this script
    Start-Process -FilePath PowerShell.exe -ArgumentList $argString -Verb RunAs
    
    # Exit the non-elevated script
    exit
}

# Get the script directory
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$toolsFile = Join-Path $scriptRoot "tools.json"

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.ToolsInstalling = $false
$sync.CurrentTool = ""
$sync.TotalTools = 0
$sync.CompletedTools = 0

# Create log function for centralized logging
function Write-ToolLog {
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
    if ($null -ne $sync.LogTextBox) {
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

# Function to resolve UNC or local paths
function Resolve-Path-Safe {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$false)]
        [string]$BasePath = $scriptRoot
    )
    
    # Check if path is already absolute
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }
    
    # Check if it's a UNC path
    if ($Path -match "^\\\\") {
        return $Path
    }
    
    # Otherwise, combine with base path
    return [System.IO.Path]::Combine($BasePath, $Path)
}

# Function to load tools from JSON
function Get-ToolsList {
    try {
        if (-not (Test-Path $toolsFile)) {
            Write-ToolLog "Error: tools.json not found at $toolsFile" -Type Error
            [System.Windows.MessageBox]::Show(
                "tools.json file not found at $toolsFile. Please make sure the file exists.",
                "File Not Found",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return @()
        }
        
        $toolsJson = Get-Content -Path $toolsFile -Raw | ConvertFrom-Json
        
        if (-not $toolsJson) {
            Write-ToolLog "No tools found in tools.json!" -Type Warning
            return @()
        }
        
        # Process each tool to ensure paths are absolute
        foreach ($tool in $toolsJson) {
            if ($tool.PSObject.Properties.Name -contains "Path") {
                $tool.Path = Resolve-Path-Safe -Path $tool.Path
            }
            if ($tool.PSObject.Properties.Name -contains "ScriptPath") {
                $tool.ScriptPath = Resolve-Path-Safe -Path $tool.ScriptPath
            }
        }
        
        Write-ToolLog "Successfully loaded $($toolsJson.Count) tools from tools.json" -Type Success
        return $toolsJson
    }
    catch {
        Write-ToolLog "Error loading tools.json: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "Error loading tools.json: $_",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return @()
    }
}

# Function to add tool path to environment variables
function Add-ToolToEnvPath {
    param ($tool)
    
    try {
        # Check if tool has EnvPath property
        if ($tool.PSObject.Properties.Name -contains "EnvPath") {
            $envPath = $tool.EnvPath
            
            # Expand environment variables if present in the path
            if ($envPath -match '\$env:') {
                Write-ToolLog "Expanding environment variables in path: $envPath" -Type Information
                $envPath = $ExecutionContext.InvokeCommand.ExpandString($envPath)
                Write-ToolLog "Expanded path: $envPath" -Type Information
            }
            
            # Resolve the path if it's relative
            if (-not [System.IO.Path]::IsPathRooted($envPath)) {
                $envPath = Resolve-Path-Safe -Path $envPath
            }
            
            # Verify path exists
            if ($envPath -and (Test-Path $envPath)) {
                Write-ToolLog "Processing EnvPath for $($tool.Name): $envPath" -Type Information
                
                # Get current system PATH
                $currentPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
                
                # Check if path is already in the PATH variable
                if ($currentPath -notlike "*$envPath*") {
                    # Append and update system PATH
                    $newPath = "$currentPath;$envPath"
                    [System.Environment]::SetEnvironmentVariable("Path", $newPath, [System.EnvironmentVariableTarget]::Machine)
                    Write-ToolLog "Added $envPath to system PATH" -Type Success
                    return $true
                } else {
                    Write-ToolLog "$envPath is already in system PATH" -Type Information
                    return $true
                }
            } else {
                Write-ToolLog "EnvPath for $($tool.Name) not found or invalid: $envPath" -Type Warning
                return $false
            }
        }
        
        # No EnvPath property - nothing to do
        return $true
    }
    catch {
        Write-ToolLog "Error adding $($tool.Name) to environment PATH: $_" -Type Error
        return $false
    }
}

# Function to install EXE installers
function Install-EXE {
    param ($tool)
    try {
        Write-ToolLog "Installing $($tool.Name) (EXE Installer)" -Type Information
        
        # Check if local path exists, if not copy from fallback
        if (-not (Test-Path $tool.Path)) {
            Write-ToolLog "Installer not found at $($tool.Path), attempting to copy from fallback location" -Type Warning
            
            # Check if fallback path exists
            if ($tool.PSObject.Properties.Name -contains "FallbackPath" -and (Test-Path $tool.FallbackPath)) {
                # Create directory if it doesn't exist
                $directory = Split-Path -Path $tool.Path -Parent
                if (-not (Test-Path $directory)) {
                    New-Item -ItemType Directory -Path $directory -Force | Out-Null
                    Write-ToolLog "Created directory: $directory" -Type Information
                }
                
                # Copy the file from fallback
                Copy-Item -Path $tool.FallbackPath -Destination $tool.Path -Force
                Write-ToolLog "Copied installer from $($tool.FallbackPath) to $($tool.Path)" -Type Information
                
                # Verify copy was successful
                if (-not (Test-Path $tool.Path)) {
                    Write-ToolLog "Error: Failed to copy installer from fallback location" -Type Error
                    return $false
                }
            }
            else {
                Write-ToolLog "Error: Installer not found at primary path and no valid fallback path available" -Type Error
                return $false
            }
        }
        
        # Get arguments or use default
        $arguments = if ($tool.PSObject.Properties.Name -contains "Arguments") { $tool.Arguments } else { "/quiet" }
        
        # Update status and run installer
        $sync.StatusBar.Text = "Installing $($tool.Name)..."
        Start-Process -FilePath $tool.Path -ArgumentList $arguments -Wait
        
        Write-ToolLog "$($tool.Name) installed successfully" -Type Success
        return $true
    }
    catch {
        Write-ToolLog "Error installing $($tool.Name): $_" -Type Error
        return $false
    }
}

# Function to install Portable apps (ZIP extraction)
function Install-Portable {
    param ($tool)
    
    try {
        Write-ToolLog "Extracting $($tool.Name) to $($tool.InstallLocation)" -Type Information
        
        if (-not (Test-Path $tool.Path)) {
            Write-ToolLog "Error: ZIP file not found at $($tool.Path)" -Type Error
            return $false
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $tool.InstallLocation)) {
            New-Item -ItemType Directory -Path $tool.InstallLocation -Force | Out-Null
            Write-ToolLog "Created directory: $($tool.InstallLocation)" -Type Information
        }
        
        $sync.StatusBar.Text = "Extracting $($tool.Name)..."
        Expand-Archive -Path $tool.Path -DestinationPath $tool.InstallLocation -Force
        
        Write-ToolLog "$($tool.Name) extracted successfully to $($tool.InstallLocation)" -Type Success
        return $true
    }
    catch {
        Write-ToolLog "Error extracting $($tool.Name): $_" -Type Error
        return $false
    }
}

# Function to install ZIP-based apps
function Install-Zip {
    param ($tool)
    
    try {
        Write-ToolLog "Extracting $($tool.Name) to $($tool.InstallLocation)" -Type Information
        
        if (-not (Test-Path $tool.Path)) {
            Write-ToolLog "Error: ZIP file not found at $($tool.Path)" -Type Error
            return $false
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $tool.InstallLocation)) {
            New-Item -ItemType Directory -Path $tool.InstallLocation -Force | Out-Null
            Write-ToolLog "Created directory: $($tool.InstallLocation)" -Type Information
        }
        
        $sync.StatusBar.Text = "Extracting $($tool.Name)..."
        Expand-Archive -Path $tool.Path -DestinationPath $tool.InstallLocation -Force
        
        Write-ToolLog "$($tool.Name) extracted successfully to $($tool.InstallLocation)" -Type Success
        return $true
    }
    catch {
        Write-ToolLog "Error extracting $($tool.Name): $_" -Type Error
        return $false
    }
}

# Function to install via Chocolatey
function Install-Choco {
    param ($tool)
    try {
        $packageName = $tool.Package
        $extraArgs = if ($tool.PSObject.Properties.Name -contains "Arguments") { $tool.Arguments } else { "" }
        Write-ToolLog "Installing $($tool.Name) via Chocolatey (package: $packageName)" -Type Information
        
        # Check if Chocolatey is installed
        if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
            Write-ToolLog "Chocolatey is not installed." -Type Warning
            
            # Create a prompt window to ask user if they want to install Chocolatey
            $chocoPrompt = New-Object System.Windows.Window
            $chocoPrompt.Title = "Chocolatey Installation Required"
            $chocoPrompt.SizeToContent = "WidthAndHeight"
            $chocoPrompt.WindowStartupLocation = "CenterScreen"
            $chocoPrompt.ResizeMode = "NoResize"
            $chocoPrompt.Topmost = $true
            
            $promptGrid = New-Object System.Windows.Controls.Grid
            $promptGrid.Margin = "15"
            
            # Define grid rows
            $row1 = New-Object System.Windows.Controls.RowDefinition
            $row2 = New-Object System.Windows.Controls.RowDefinition
            $promptGrid.RowDefinitions.Add($row1)
            $promptGrid.RowDefinitions.Add($row2)
            
            # Message text
            $messageText = New-Object System.Windows.Controls.TextBlock
            $messageText.Text = "Chocolatey package manager is required to install $($tool.Name) but is not installed on this system. Would you like to install Chocolatey now?"
            $messageText.TextWrapping = "Wrap"
            $messageText.Margin = "0,0,0,15"
            $messageText.MaxWidth = "400"
            [System.Windows.Controls.Grid]::SetRow($messageText, 0)
            $promptGrid.Children.Add($messageText)
            
            # Buttons panel
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.HorizontalAlignment = "Right"
            [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
            
            # Yes button
            $yesButton = New-Object System.Windows.Controls.Button
            $yesButton.Content = "Yes, Install Chocolatey"
            $yesButton.Padding = "10,5"
            $yesButton.Margin = "0,0,10,0"
            $yesButton.Background = "#0078D7"
            $yesButton.Foreground = "White"
            
            # No button
            $noButton = New-Object System.Windows.Controls.Button
            $noButton.Content = "No, Skip Installation"
            $noButton.Padding = "10,5"
            
            $userConsent = $false
            
            # Button click handlers
            $yesButton.Add_Click({
                $script:userConsent = $true
                $chocoPrompt.Close()
            })
            
            $noButton.Add_Click({
                $script:userConsent = $false
                $chocoPrompt.Close()
            })
            
            # Add buttons to panel
            $buttonPanel.Children.Add($yesButton)
            $buttonPanel.Children.Add($noButton)
            $promptGrid.Children.Add($buttonPanel)
            
            # Set window content and show dialog
            $chocoPrompt.Content = $promptGrid
            $chocoPrompt.ShowDialog() | Out-Null
            
            # Check user's response
            if ($script:userConsent -eq $true) {
                Write-ToolLog "User consented to Chocolatey installation. Installing now..." -Type Information
                $sync.StatusBar.Text = "Installing Chocolatey..."
                
                Set-ExecutionPolicy Bypass -Scope Process -Force
                Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
                
                if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                    Write-ToolLog "Failed to install Chocolatey" -Type Error
                    return $false
                }
                Write-ToolLog "Chocolatey installed successfully" -Type Success
            } else {
                Write-ToolLog "User declined Chocolatey installation. Aborting installation of $($tool.Name)." -Type Warning
                return $false
            }
        }
        
        # Construct Chocolatey install command with optional arguments
        $sync.StatusBar.Text = "Installing $($tool.Name) via Chocolatey..."
        $command = "choco install $packageName -y $extraArgs"
        
        # Use PowerShell to execute and capture output
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$command`"" -NoNewWindow -PassThru -Wait
        
        if ($process.ExitCode -eq 0) {
            Write-ToolLog "$($tool.Name) installed successfully via Chocolatey" -Type Success
            return $true
        } else {
            Write-ToolLog "Chocolatey installation failed with exit code $($process.ExitCode)" -Type Error
            return $false
        }
    } catch {
        Write-ToolLog "Error installing $($tool.Name) via Chocolatey: $_" -Type Error
        return $false
    }
}

# Function to run custom installation script
function Install-Custom {
    param ($tool)
    
    try {
        Write-ToolLog "Running custom installation for $($tool.Name)" -Type Information
        
        if (-not (Test-Path $tool.ScriptPath)) {
            Write-ToolLog "Error: Script not found at $($tool.ScriptPath)" -Type Error
            return $false
        }
        
        $arguments = if ($tool.PSObject.Properties.Name -contains "ScriptArguments") { $tool.ScriptArguments } else { "" }
        
        $sync.StatusBar.Text = "Running custom installation for $($tool.Name)..."
        
        # Execute the PowerShell script with parameters if provided
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($tool.ScriptPath)`" $arguments" -NoNewWindow -PassThru -Wait
        
        if ($process.ExitCode -eq 0) {
            Write-ToolLog "Custom installation for $($tool.Name) completed successfully" -Type Success
            return $true
        }
        else {
            Write-ToolLog "Custom installation failed with exit code $($process.ExitCode)" -Type Error
            return $false
        }
    }
    catch {
        Write-ToolLog "Error during custom installation for $($tool.Name): $_" -Type Error
        return $false
    }
}

# Function to execute installation based on type
function Install-Tool {
    param ($tool)
    
    $sync.CurrentTool = $tool.Name
    $sync.StatusBar.Text = "Installing $($tool.Name)..."
    
    try {
        $result = switch ($tool.Type) {
            "exe_installer" { Install-EXE -tool $tool }
            "portable" { Install-Portable -tool $tool }
            "zip" { Install-Zip -tool $tool }
            "choco" { Install-Choco -tool $tool }
            "custom" { Install-Custom -tool $tool }
            default { 
                Write-ToolLog "Unknown install type: $($tool.Type) for $($tool.Name)" -Type Warning
                $false
            }
        }
        
        # If installation was successful, add to environment path if specified
        if ($result) {
            Add-ToolToEnvPath -tool $tool
        }
        
        $sync.CompletedTools++
        $progress = [int](($sync.CompletedTools / $sync.TotalTools) * 100)
        $sync.ProgressBar.Value = $progress
        
        if ($result) {
            $sync.StatusBar.Text = "$($tool.Name) installed successfully!"
            return $true
        } else {
            $sync.StatusBar.Text = "Failed to install $($tool.Name)"
            return $false
        }
    }
    catch {
        Write-ToolLog "Error in Install-Tool for $($tool.Name): $_" -Type Error
        $sync.StatusBar.Text = "Error installing $($tool.Name)"
        return $false
    }
}

# Function to install selected tools
function Install-SelectedTools {
    param (
        [array]$SelectedTools
    )
    
    if ($SelectedTools.Count -eq 0) {
        Write-ToolLog "No tools selected for installation" -Type Warning
        $sync.StatusBar.Text = "No tools selected"
        return
    }
    
    try {
        $sync.ToolsInstalling = $true
        $sync.InstallButton.IsEnabled = $false
        $sync.TotalTools = $SelectedTools.Count
        $sync.CompletedTools = 0
        $sync.ProgressBar.Value = 0
        
        Write-ToolLog "Starting installation of $($SelectedTools.Count) tools" -Type Information
        
        $successful = 0
        $failed = 0
        
        foreach ($tool in $SelectedTools) {
            $result = Install-Tool -tool $tool
            if ($result) { $successful++ } else { $failed++ }
        }
        
        $sync.ProgressBar.Value = 100
        $sync.StatusBar.Text = "Installation completed: $successful succeeded, $failed failed"
        Write-ToolLog "Installation completed: $successful tools installed successfully, $failed failed" -Type Information
    }
    catch {
        Write-ToolLog "Error during installation process: $_" -Type Error
    }
    finally {
        $sync.ToolsInstalling = $false
        $sync.InstallButton.IsEnabled = $true
    }
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Tool Installer" 
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
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Tool Installer" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Install software tools" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="RefreshButton" Grid.Column="1" 
                        Style="{StaticResource DefaultButton}"
                        Content="⟳ Refresh" 
                        ToolTip="Refresh tool list"
                        VerticalAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="15">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Search Box -->
            <Grid Grid.Row="0" Margin="0,0,0,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="🔍 Search: " VerticalAlignment="Center" Margin="0,0,10,0" FontSize="14"/>
                <TextBox x:Name="SearchBox" Grid.Column="1" Padding="8" FontSize="14" BorderThickness="1" BorderBrush="#CCCCCC">
                    <TextBox.Resources>
                        <Style TargetType="{x:Type Border}">
                            <Setter Property="CornerRadius" Value="3"/>
                        </Style>
                    </TextBox.Resources>
                </TextBox>
            </Grid>
            
            <!-- Tools List View -->
            <ListView x:Name="ToolListView" Grid.Row="1" 
                      BorderThickness="1" 
                      BorderBrush="#CCCCCC"
                      Padding="0"
                      FontSize="14"
                      SelectionMode="Multiple">
                <ListView.Resources>
                    <Style TargetType="{x:Type ListViewItem}">
                        <Setter Property="Padding" Value="10,8"/>
                        <Setter Property="BorderThickness" Value="0,0,0,1"/>
                        <Setter Property="BorderBrush" Value="#EEEEEE"/>
                        <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#E3F2FD"/>
                                <Setter Property="BorderBrush" Value="#BBDEFB"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#F5F5F5"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </ListView.Resources>
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Name" Width="250">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Name}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Type" Width="100">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Type}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Details" Width="480">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock>
                                        <TextBlock.Text>
                                            <MultiBinding StringFormat="{}{0}{1}{2}{3}">
                                                <Binding Path="Path" FallbackValue="" />
                                                <Binding Path="Package" FallbackValue="" />
                                                <Binding Path="InstallLocation" FallbackValue="" />
                                                <Binding Path="ScriptPath" FallbackValue="" />
                                            </MultiBinding>
                                        </TextBlock.Text>
                                    </TextBlock>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                    </GridView>
                </ListView.View>
            </ListView>
        </Grid>
        
        <!-- Progress Bar -->
        <Grid Grid.Row="2" Margin="15,0,15,15">
            <ProgressBar x:Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Value="0" />
        </Grid>
        
        <!-- Tool Details -->
        <Border Grid.Row="3" Background="#F5F5F5" Padding="15" BorderThickness="0,1,0,1" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="SelectedToolCount" Text="No tools selected" FontWeight="Bold" FontSize="14"/>
                    <TextBlock x:Name="SelectedToolDetails" Text="" Foreground="#666666" FontSize="12" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="InstallButton" 
                        Grid.Column="1" 
                        Content="Install Selected Tools" 
                        Style="{StaticResource DefaultButton}"
                        FontSize="14"/>
            </Grid>
        </Border>
        
        <!-- Log Panel -->
        <Grid Grid.Row="4" Margin="15,15,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="150"/>
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
        <Border Grid.Row="5" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
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
$sync.ToolListView = $window.FindName("ToolListView")
$sync.SearchBox = $window.FindName("SearchBox")
$sync.RefreshButton = $window.FindName("RefreshButton")
$sync.ProgressBar = $window.FindName("ProgressBar")
$sync.InstallButton = $window.FindName("InstallButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.SelectedToolCount = $window.FindName("SelectedToolCount")
$sync.SelectedToolDetails = $window.FindName("SelectedToolDetails")

# Initial state - Install button disabled
$sync.InstallButton.IsEnabled = $false

# Populate the ListView with tools
function Update-ToolList {
    $tools = Get-ToolsList
    $sync.ToolListView.Items.Clear()
    
    $searchText = $sync.SearchBox.Text.ToLower()
    
    if ($tools.Count -eq 0) {
        $sync.StatusBar.Text = "No tools found in tools.json"
        return
    }
    
    foreach ($tool in $tools) {
        # Apply search filter if search text is not empty
        if ([string]::IsNullOrWhiteSpace($searchText) -or 
            $tool.Name.ToLower().Contains($searchText) -or 
            $tool.Type.ToLower().Contains($searchText)) {
            $sync.ToolListView.Items.Add($tool)
        }
    }
    
    $sync.StatusBar.Text = "Found $($sync.ToolListView.Items.Count) tools"
}

# Set up event handlers
$sync.SearchBox.Add_TextChanged({
    Update-ToolList
})

$sync.RefreshButton.Add_Click({
    Update-ToolList
    $sync.StatusBar.Text = "Tool list refreshed"
})

$sync.InstallButton.Add_Click({
    if ($sync.ToolsInstalling) {
        [System.Windows.MessageBox]::Show("Installation is already in progress", "Operation in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $selectedItems = $sync.ToolListView.SelectedItems
    
    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one tool to install", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Confirm installation
    $confirmResult = [System.Windows.MessageBox]::Show(
        "Are you sure you want to install $($selectedItems.Count) selected tools?",
        "Confirm Installation",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
        Install-SelectedTools -SelectedTools $selectedItems
    }
    else {
        Write-ToolLog "Installation cancelled by user" -Type Information
    }
})

# Selection changed event
$sync.ToolListView.Add_SelectionChanged({
    $selectedItems = $sync.ToolListView.SelectedItems
    
    if ($selectedItems.Count -gt 0) {
        $sync.SelectedToolCount.Text = "$($selectedItems.Count) tool(s) selected"
        
        if ($selectedItems.Count -eq 1) {
            $tool = $selectedItems[0]
            $details = "Type: $($tool.Type)"
            
            switch ($tool.Type) {
                "exe_installer" { $details += " | Path: $($tool.Path)" }
                "portable" { $details += " | Path: $($tool.Path) | Install Location: $($tool.InstallLocation)" }
                "zip" { $details += " | Path: $($tool.Path) | Install Location: $($tool.InstallLocation)" }
                "choco" { $details += " | Package: $($tool.Package)" }
                "custom" { $details += " | Script Path: $($tool.ScriptPath)" }
            }
            
            # Add EnvPath to details if available
            if ($tool.PSObject.Properties.Name -contains "EnvPath") {
                $details += " | EnvPath: $($tool.EnvPath)"
            }
            
            $sync.SelectedToolDetails.Text = $details
        }
        else {
            $sync.SelectedToolDetails.Text = "Multiple tools selected"
        }
        
        $sync.InstallButton.IsEnabled = -not $sync.ToolsInstalling
    }
    else {
        $sync.SelectedToolCount.Text = "No tools selected"
        $sync.SelectedToolDetails.Text = ""
        $sync.InstallButton.IsEnabled = $false
    }
})

# Window closing event
$sync.Window.Add_Closing({
    if ($sync.ToolsInstalling) {
        $result = [System.Windows.MessageBox]::Show(
            "An installation is currently in progress. Are you sure you want to exit?", 
            "Confirm Exit", 
            [System.Windows.MessageBoxButton]::YesNo, 
            [System.Windows.MessageBoxImage]::Question)
            
        if ($result -eq [System.Windows.MessageBoxResult]::No) {
            $_.Cancel = $true
        }
    }
})

# Initial population of the tool list
Update-ToolList

# Start with welcome message
Write-ToolLog "Tool Installer started. Found $($sync.ToolListView.Items.Count) tools in tools.json" -Type Information

# Show the window
$sync.Window.ShowDialog() | Out-Null