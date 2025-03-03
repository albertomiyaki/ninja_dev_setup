# ninjaDEV Setup Launcher
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Get the root directory
$scriptRoot = $PSScriptRoot
$modulesFolder = Join-Path $scriptRoot "Modules"

# Create a synchronized hashtable for sharing data between runspaces
$sync = [hashtable]::Synchronized(@{})
$sync.ScriptRunning = $false
$sync.CurrentScriptName = ""
$sync.ActiveRunspaces = @()
$sync.RunningScripts = @{}

# Create log function for centralized logging
function Write-LauncherLog {
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

# Ensure the Modules folder exists
if (-not (Test-Path $modulesFolder)) {
    Write-LauncherLog "Error: The 'Modules' folder does not exist!" -Type Error
    [System.Windows.MessageBox]::Show(
        "The 'Modules' folder does not exist in $scriptRoot. Please create it and add your PowerShell scripts.",
        "Modules Folder Missing",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    exit
}

# Function to extract the first comment from a script as its description
function Get-ScriptDescription {
    param (
        [string]$ScriptPath
    )

    $lines = Get-Content -Path $ScriptPath -TotalCount 5 -ErrorAction SilentlyContinue

    if ($null -eq $lines) {
        return "Could not read script file"
    }

    foreach ($line in $lines) {
        if ($line -match "^\s*#\s*(.+)") {
            return $matches[1].Trim()  # Extract and return the comment text
        }
    }

    return "No description available"
}

# Function to get all scripts inside the Modules folder
function Get-AvailableScripts {
    try {
        Write-LauncherLog "Scanning for scripts in: $modulesFolder" -Type Information
       
        # Get immediate subdirectories under the Modules folder
        $subfolders = Get-ChildItem -Path $modulesFolder -Directory -ErrorAction Stop
        
        $scripts = @()
        
        # For each subfolder, get only the PS1 files directly in that folder (not recursive)
        foreach ($folder in $subfolders) {
            $folderScripts = Get-ChildItem -Path $folder.FullName -Filter "*.ps1" -File -ErrorAction SilentlyContinue
            
            foreach ($script in $folderScripts) {
                $scripts += [PSCustomObject]@{
                    Name = "$($folder.Name)\$($script.Name)"  # Show folder and script name
                    Description = Get-ScriptDescription -ScriptPath $script.FullName
                    FullPath = $script.FullName
                    ModuleName = $folder.Name
                    ScriptName = $script.Name
                    Status = "Ready"
                    LastRun = $null
                }
            }
        }
        
        # If no scripts are found, return empty array
        if ($null -eq $scripts -or $scripts.Count -eq 0) {
            Write-LauncherLog "No scripts found in the Modules folder" -Type Warning
            return @()
        }
       
        Write-LauncherLog "Found $($scripts.Count) scripts" -Type Success
        return $scripts
    }
    catch {
        Write-LauncherLog "Error loading scripts: $_" -Type Error
        return @()
    }
}

# Function to run a script in a new PowerShell instance
function Invoke-SelectedScript {
    param (
        [string]$ScriptPath,
        [string]$ScriptName
    )
    
    try {
        $scriptID = [Guid]::NewGuid().ToString()
        $sync.RunningScripts[$scriptID] = @{
            Name = $ScriptName
            Path = $ScriptPath
            StartTime = Get-Date
        }
        
        Write-LauncherLog "Launching script: $ScriptName" -Type Information
        $sync.StatusBar.Text = "Launching script: $ScriptName..."
        
        # Create a script block that will run the selected script
        $scriptBlock = {
            param($scriptPath, $scriptID)
            
            try {
                # Run the script and capture its exit code
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -PassThru
                return @{
                    ID = $scriptID
                    Success = $true
                    ExitCode = $process.ExitCode
                }
            }
            catch {
                return @{
                    ID = $scriptID
                    Success = $false
                    Error = $_.Exception.Message
                }
            }
        }
        
        # Run the script in a new runspace
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = "STA"
        $runspace.ThreadOptions = "ReuseThread"
        $runspace.Open()
        
        # Create PowerShell instance and supply the script block
        $powershell = [powershell]::Create()
        $powershell.Runspace = $runspace
        
        # Add the script block and parameters
        $powershell.AddScript($scriptBlock).AddArgument($ScriptPath).AddArgument($scriptID) | Out-Null
        
        # Begin asynchronous execution
        $asyncResult = $powershell.BeginInvoke()
        
        # Store the runspace information
        $sync.ActiveRunspaces += [PSCustomObject]@{
            Runspace = $runspace
            PowerShell = $powershell
            AsyncResult = $asyncResult
            ScriptID = $scriptID
        }
        
        # Update the UI to show the script is running
        $script = $sync.ScriptListView.Items | Where-Object { $_.FullPath -eq $ScriptPath }
        if ($script) {
            $script.Status = "Running"
            $script.LastRun = Get-Date
            $sync.ScriptListView.Items.Refresh()
        }
        
        Write-LauncherLog "Script launched: $ScriptName (ID: $scriptID)" -Type Success
        $sync.StatusBar.Text = "Script launched: $ScriptName"
        
        return $true
    }
    catch {
        Write-LauncherLog "Error launching script: $ScriptName - $_" -Type Error
        $sync.StatusBar.Text = "Error launching script: $ScriptName"
        return $false
    }
}

# Function to check and clean up completed runspaces
function Update-RunspaceStatus {
    $runspacesToRemove = @()
    
    foreach ($runspaceInfo in $sync.ActiveRunspaces) {
        if ($runspaceInfo.AsyncResult.IsCompleted) {
            try {
                # Get the result
                $result = $runspaceInfo.PowerShell.EndInvoke($runspaceInfo.AsyncResult)
                
                # Find the script in the ListView
                $scriptID = $runspaceInfo.ScriptID
                $scriptInfo = $sync.RunningScripts[$scriptID]
                
                if ($scriptInfo) {
                    $script = $sync.ScriptListView.Items | Where-Object { $_.FullPath -eq $scriptInfo.Path }
                    
                    if ($script) {
                        if ($result.Success) {
                            $script.Status = "Completed"
                            Write-LauncherLog "Script completed: $($scriptInfo.Name)" -Type Success
                        }
                        else {
                            $script.Status = "Failed"
                            Write-LauncherLog "Script failed: $($scriptInfo.Name) - $($result.Error)" -Type Error
                        }
                        
                        $sync.ScriptListView.Items.Refresh()
                    }
                    
                    # Remove from running scripts
                    $sync.RunningScripts.Remove($scriptID)
                }
                
                # Close and clean up the runspace
                $runspaceInfo.Runspace.Close()
                $runspaceInfo.Runspace.Dispose()
                $runspaceInfo.PowerShell.Dispose()
                
                # Mark for removal
                $runspacesToRemove += $runspaceInfo
            }
            catch {
                Write-LauncherLog "Error processing completed script: $_" -Type Error
            }
        }
    }
    
    # Remove completed runspaces from the active list
    foreach ($runspaceInfo in $runspacesToRemove) {
        $sync.ActiveRunspaces = $sync.ActiveRunspaces | Where-Object { $_ -ne $runspaceInfo }
    }
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ninjaDEV Script Launcher" 
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
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock Text="ninjaDEV Script Launcher" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="ninjaDEV: Launch and manage PowerShell scripts" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="RefreshButton" Grid.Column="1" 
                        Style="{StaticResource DefaultButton}"
                        Content="⟳ Refresh" 
                        ToolTip="Refresh script list"
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
            
            <!-- Scripts List View -->
            <ListView x:Name="ScriptListView" Grid.Row="1" 
                      BorderThickness="1" 
                      BorderBrush="#CCCCCC"
                      Padding="0"
                      FontSize="14">
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
                        <GridViewColumn Header="Module" Width="140">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ModuleName}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Description" Width="400">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Description}" TextWrapping="Wrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Script" Width="200">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ScriptName}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Status" Width="100">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Status}" TextWrapping="NoWrap">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Status}" Value="Running">
                                                        <Setter Property="Foreground" Value="Blue"/>
                                                        <Setter Property="FontWeight" Value="Bold"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Completed">
                                                        <Setter Property="Foreground" Value="Green"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Failed">
                                                        <Setter Property="Foreground" Value="Red"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                    </GridView>
                </ListView.View>
            </ListView>
        </Grid>
        
        <!-- Script Details -->
        <Border Grid.Row="2" Background="#F5F5F5" Padding="15" BorderThickness="0,1,0,1" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="SelectedScriptName" Text="No script selected" FontWeight="Bold" FontSize="14"/>
                    <TextBlock x:Name="SelectedScriptPath" Text="" Foreground="#666666" FontSize="12" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="RunButton" 
                        Grid.Column="1" 
                        Content="▶ Run Script" 
                        Style="{StaticResource DefaultButton}"
                        FontSize="14"/>
            </Grid>
        </Border>
        
        <!-- Log Panel -->
        <Grid Grid.Row="3" Margin="15,15,15,0">
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
        <Border Grid.Row="4" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
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
    Write-LauncherLog "Error loading XAML: $_" -Type Error
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Store references in the sync hashtable
$sync.Window = $window
$sync.ScriptListView = $window.FindName("ScriptListView")
$sync.SearchBox = $window.FindName("SearchBox")
$sync.RefreshButton = $window.FindName("RefreshButton")
$sync.RunButton = $window.FindName("RunButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.SelectedScriptName = $window.FindName("SelectedScriptName")
$sync.SelectedScriptPath = $window.FindName("SelectedScriptPath")

# Initial state - Run button disabled
$sync.RunButton.IsEnabled = $false

# Populate the ListView with scripts
function Update-ScriptList {
    $scripts = Get-AvailableScripts
    $sync.ScriptListView.Items.Clear()
    
    $searchText = $sync.SearchBox.Text.ToLower()
    
    if ($scripts.Count -eq 0) {
        $sync.StatusBar.Text = "No scripts found in the Modules folder"
        return
    }
    
    foreach ($script in $scripts) {
        # Apply search filter if search text is not empty
        if ([string]::IsNullOrWhiteSpace($searchText) -or 
            $script.Name.ToLower().Contains($searchText) -or 
            $script.Description.ToLower().Contains($searchText) -or
            $script.ModuleName.ToLower().Contains($searchText)) {
            $sync.ScriptListView.Items.Add($script)
        }
    }
    
    $sync.StatusBar.Text = "Found $($sync.ScriptListView.Items.Count) scripts"
}

# Create a timer to check runspace status
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(1000)
$timer.Add_Tick({ Update-RunspaceStatus })
$timer.Start()

# Set up event handlers
$sync.SearchBox.Add_TextChanged({
    Update-ScriptList
})

$sync.RefreshButton.Add_Click({
    Update-ScriptList
    $sync.StatusBar.Text = "Script list refreshed"
})

$sync.RunButton.Add_Click({
    $selectedItem = $sync.ScriptListView.SelectedItem
    
    if ($selectedItem -eq $null) {
        [System.Windows.MessageBox]::Show("Please select a script to run", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    Invoke-SelectedScript -ScriptPath $selectedItem.FullPath -ScriptName $selectedItem.Name
})

# Selection changed event
$sync.ScriptListView.Add_SelectionChanged({
    $selectedItem = $sync.ScriptListView.SelectedItem
    
    if ($selectedItem -ne $null) {
        $sync.SelectedScriptName.Text = "$($selectedItem.ModuleName)\$($selectedItem.ScriptName)"
        $sync.SelectedScriptPath.Text = $selectedItem.FullPath
        $sync.RunButton.IsEnabled = $true
    }
    else {
        $sync.SelectedScriptName.Text = "No script selected"
        $sync.SelectedScriptPath.Text = ""
        $sync.RunButton.IsEnabled = $false
    }
})

# Double-click to run a script
$sync.ScriptListView.Add_MouseDoubleClick({
    $selectedItem = $sync.ScriptListView.SelectedItem
    
    if ($selectedItem -ne $null) {
        Invoke-SelectedScript -ScriptPath $selectedItem.FullPath -ScriptName $selectedItem.Name
    }
})

# Window closing event
$sync.Window.Add_Closing({
    # Clean up any remaining runspaces
    foreach ($runspaceInfo in $sync.ActiveRunspaces) {
        try {
            $runspaceInfo.PowerShell.Stop()
            $runspaceInfo.Runspace.Close()
            $runspaceInfo.Runspace.Dispose()
            $runspaceInfo.PowerShell.Dispose()
        }
        catch {
            # Just continue with cleanup
        }
    }
    
    # Stop the timer
    $timer.Stop()
})

# Initial population of the script list
Update-ScriptList

# Start with welcome message
Write-LauncherLog "Script Launcher started. Found $($sync.ScriptListView.Items.Count) scripts in the Modules folder" -Type Information

# Show the window
$sync.Window.ShowDialog() | Out-Null