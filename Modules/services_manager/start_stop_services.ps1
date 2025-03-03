# Service Manager - start/stop your services and SQL Server instances
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Check if the script is running with administrator privileges
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Self-elevate the script if required
if (-not (Test-Administrator)) {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    
    $scriptPath = $MyInvocation.MyCommand.Definition
    $scriptDir = Split-Path -Parent $scriptPath
    $scriptFile = Split-Path -Leaf $scriptPath
    
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

# Define multiple regex patterns to filter services (comma separated)
$regexPatterns = @("Xbox.*", "SQL Server \(.*")

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.ServiceToggling = $false
$sync.CurrentService = ""

# Create log function for centralized logging
function Write-ServiceLog {
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

# Function to get services matching the regex patterns
function Get-FilteredServices {
    Write-ServiceLog "Refreshing service list..." -Type Information
    
    # Get services matching any pattern
    $filteredServices = Get-Service | Where-Object { 
        foreach ($pattern in $regexPatterns) {
            if ($_.DisplayName -match $pattern) { return $true }
        }
        return $false
    } | ForEach-Object {
        $startupType = (Get-WmiObject -Class Win32_Service -Filter "Name='$($_.Name)'").StartMode
        # Add properties to the service object
        $_ | Select-Object Name, DisplayName, Status, @{
            Name="StartupType"; 
            Expression={ $startupType }
        }, @{
            Name="Action"; 
            Expression={ if ($_.Status -eq "Running") { "Stop" } else { "Start" } }
        }
    }
    
    # Return the filtered services
    return $filteredServices
}

# Function to toggle service state
function Set-ServiceState {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Service
    )
    
    try {
        $sync.ServiceToggling = $true
        $sync.CurrentService = $Service.DisplayName
        $sync.ToggleButton.IsEnabled = $false
        $sync.StatusBar.Text = "Toggling service: $($Service.DisplayName)..."
        
        if ($Service.Status -eq "Running") {
            Write-ServiceLog "Stopping service: $($Service.DisplayName)..." -Type Warning
            Stop-Service -Name $Service.Name -Force -ErrorAction Stop
            Write-ServiceLog "Service '$($Service.DisplayName)' stopped successfully." -Type Success
        } else {
            Write-ServiceLog "Starting service: $($Service.DisplayName)..." -Type Warning
            Start-Service -Name $Service.Name -ErrorAction Stop
            Write-ServiceLog "Service '$($Service.DisplayName)' started successfully." -Type Success
        }
        
        $sync.ServiceToggling = $false
        $sync.ToggleButton.IsEnabled = $true
        $sync.StatusBar.Text = "Service '$($Service.DisplayName)' toggled successfully."
        return $true
    }
    catch {
        Write-ServiceLog "Error toggling service '$($Service.DisplayName)': $_" -Type Error
        $sync.ServiceToggling = $false
        $sync.ToggleButton.IsEnabled = $true
        $sync.StatusBar.Text = "Error toggling service: $_"
        return $false
    }
}

# Function to change service startup type
function Set-ServiceStartupType {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Automatic", "Manual", "Disabled")]
        [string]$StartupType
    )
    
    try {
        $sync.ServiceChanging = $true
        $sync.StatusBar.Text = "Changing startup type for service: $ServiceName..."
        
        Write-ServiceLog "Changing startup type for service '$ServiceName' to $StartupType..." -Type Warning
        
        # Convert UI startup type to sc.exe parameter
        $startupArg = switch ($StartupType) {
            "Automatic" { "auto" }
            "Manual" { "demand" }
            "Disabled" { "disabled" }
            default { "demand" }
        }
        
        # Use sc.exe command which has better compatibility than WMI
        $process = Start-Process -FilePath "sc.exe" -ArgumentList "config `"$ServiceName`" start= $startupArg" -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-ServiceLog "Service '$ServiceName' startup type changed to $StartupType successfully." -Type Success
            $sync.StatusBar.Text = "Service '$ServiceName' startup type changed successfully."
            $sync.ServiceChanging = $false
            return $true
        } else {
            throw "SC.exe returned error code: $($process.ExitCode)"
        }
    }
    catch {
        Write-ServiceLog "Error changing startup type for service '$ServiceName': $_" -Type Error
        $sync.StatusBar.Text = "Error changing startup type: $_"
        $sync.ServiceChanging = $false
        return $false
    }
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Service Manager" 
    Height="900" 
    Width="1000"
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
                    <TextBlock Text="Service Manager" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Start/stop services and control startup mode" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="RefreshButton" Grid.Column="1" 
                        Style="{StaticResource DefaultButton}"
                        Content="⟳ Refresh" 
                        ToolTip="Refresh service list"
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
            
            <!-- Services List View -->
            <ListView x:Name="ServiceListView" Grid.Row="1" 
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
                        <GridViewColumn Header="Service Name" Width="200">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Name}" ToolTip="{Binding Name}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Display Name" Width="370">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding DisplayName}" TextWrapping="Wrap"/>
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
                                                        <Setter Property="Foreground" Value="Green"/>
                                                        <Setter Property="FontWeight" Value="Bold"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Stopped">
                                                        <Setter Property="Foreground" Value="Red"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Startup Type" Width="150">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding StartupType}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                    </GridView>
                </ListView.View>
            </ListView>
        </Grid>
        
        <!-- Service Details -->
        <Border Grid.Row="2" Background="#F5F5F5" Padding="15" BorderThickness="0,1,0,1" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="SelectedServiceName" Text="No service selected" FontWeight="Bold" FontSize="14"/>
                    <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                        <TextBlock x:Name="SelectedServiceStatus" Text="" Foreground="#666666" FontSize="12"/>
                        <TextBlock Text=" | " Foreground="#666666" FontSize="12"/>
                        <TextBlock Text="Startup Type: " Foreground="#666666" FontSize="12"/>
                        <TextBlock x:Name="SelectedServiceStartupType" Text="" Foreground="#666666" FontSize="12"/>
                    </StackPanel>
                </StackPanel>
                
                <ComboBox x:Name="StartupTypeComboBox" Grid.Column="1" Margin="5,0" Width="150" VerticalAlignment="Center">
                    <ComboBoxItem Content="Automatic"/>
                    <ComboBoxItem Content="Manual"/>
                    <ComboBoxItem Content="Disabled"/>
                </ComboBox>
                
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button x:Name="SetStartupTypeButton" 
                            Content="Set Startup Type" 
                            Style="{StaticResource DefaultButton}"
                            FontSize="14"/>
                    
                    <Button x:Name="ToggleButton" 
                            Content="Toggle Service" 
                            Style="{StaticResource DefaultButton}"
                            FontSize="14"/>
                </StackPanel>
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
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Store references in the sync hashtable
$sync.Window = $window
$sync.ServiceListView = $window.FindName("ServiceListView")
$sync.SearchBox = $window.FindName("SearchBox")
$sync.RefreshButton = $window.FindName("RefreshButton")
$sync.ToggleButton = $window.FindName("ToggleButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.SelectedServiceName = $window.FindName("SelectedServiceName")
$sync.SelectedServiceStatus = $window.FindName("SelectedServiceStatus")
$sync.SelectedServiceStartupType = $window.FindName("SelectedServiceStartupType")
$sync.StartupTypeComboBox = $window.FindName("StartupTypeComboBox")
$sync.SetStartupTypeButton = $window.FindName("SetStartupTypeButton")

# Initial state - Buttons disabled
$sync.ToggleButton.IsEnabled = $false
$sync.SetStartupTypeButton.IsEnabled = $false

# Populate the ListView with filtered services
function Update-ServiceList {
    $services = Get-FilteredServices
    $sync.ServiceListView.Items.Clear()
    
    $searchText = $sync.SearchBox.Text.ToLower()
    
    if ($services.Count -eq 0) {
        $sync.StatusBar.Text = "No services found matching the filter patterns"
        return
    }
    
    foreach ($service in $services) {
        # Apply search filter if search text is not empty
        if ([string]::IsNullOrWhiteSpace($searchText) -or 
            $service.Name.ToLower().Contains($searchText) -or 
            $service.DisplayName.ToLower().Contains($searchText)) {
            $sync.ServiceListView.Items.Add($service)
        }
    }
    
    $sync.StatusBar.Text = "Found $($sync.ServiceListView.Items.Count) services"
}

# Set up event handlers
$sync.SearchBox.Add_TextChanged({
    Update-ServiceList
})

$sync.RefreshButton.Add_Click({
    Update-ServiceList
    $sync.StatusBar.Text = "Service list refreshed"
})

$sync.ToggleButton.Add_Click({
    if ($sync.ServiceToggling) {
        [System.Windows.MessageBox]::Show("A service operation is already in progress", "Operation in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $selectedItems = $sync.ServiceListView.SelectedItems
    
    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one service to toggle", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # If multiple services are selected, confirm action
    if ($selectedItems.Count -gt 1) {
        $confirmResult = [System.Windows.MessageBox]::Show(
            "Are you sure you want to toggle $($selectedItems.Count) services?",
            "Confirm Multiple Actions",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($confirmResult -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-ServiceLog "Multiple service toggle action cancelled by user." -Type Information
            return
        }
    }
    
    # Process each selected service
    $changedCount = 0
    foreach ($service in $selectedItems) {
        $result = Set-ServiceState -Service $service
        if ($result) { $changedCount++ }
    }
    
    # Refresh the list after changes
    Update-ServiceList
    
    Write-ServiceLog "Successfully toggled $changedCount out of $($selectedItems.Count) services." -Type Success
})

# Set startup type button event
$sync.SetStartupTypeButton.Add_Click({
    $selectedItem = $sync.ServiceListView.SelectedItem
    
    if ($selectedItem -eq $null) {
        [System.Windows.MessageBox]::Show("Please select a service to change its startup type", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $startupType = $sync.StartupTypeComboBox.SelectedItem.Content
    
    # If multiple services are selected, confirm action
    $selectedItems = $sync.ServiceListView.SelectedItems
    if ($selectedItems.Count -gt 1) {
        $confirmResult = [System.Windows.MessageBox]::Show(
            "Are you sure you want to set $($selectedItems.Count) services to $startupType startup?",
            "Confirm Multiple Actions",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($confirmResult -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-ServiceLog "Multiple startup type change action cancelled by user." -Type Information
            return
        }
        
        # Process each selected service
        $changedCount = 0
        foreach ($service in $selectedItems) {
            $result = Set-ServiceStartupType -ServiceName $service.Name -StartupType $startupType
            if ($result) { $changedCount++ }
        }
        
        # Refresh the list after changes
        Update-ServiceList
        
        Write-ServiceLog "Successfully changed startup type to $startupType for $changedCount out of $($selectedItems.Count) services." -Type Success
    } else {
        # Change startup type for a single service
        Set-ServiceStartupType -ServiceName $selectedItem.Name -StartupType $startupType
        
        # Refresh the list after changes
        Update-ServiceList
    }
})

# Selection changed event
$sync.ServiceListView.Add_SelectionChanged({
    $selectedItem = $sync.ServiceListView.SelectedItem
    
    if ($selectedItem -ne $null) {
        $sync.SelectedServiceName.Text = $selectedItem.DisplayName
        $sync.SelectedServiceStatus.Text = "Status: $($selectedItem.Status)"
        $sync.SelectedServiceStartupType.Text = $selectedItem.StartupType
        $sync.ToggleButton.Content = "$($selectedItem.Action) Service"
        $sync.ToggleButton.IsEnabled = -not $sync.ServiceToggling
        $sync.SetStartupTypeButton.IsEnabled = $true
        
        # Set the combo box to the current startup type of the service
        switch ($selectedItem.StartupType) {
            "Auto" { $sync.StartupTypeComboBox.SelectedIndex = 0 } # Automatic
            "Manual" { $sync.StartupTypeComboBox.SelectedIndex = 1 } # Manual
            "Disabled" { $sync.StartupTypeComboBox.SelectedIndex = 2 } # Disabled
            default { $sync.StartupTypeComboBox.SelectedIndex = 1 } # Default to Manual
        }
    }
    else {
        $sync.SelectedServiceName.Text = "No service selected"
        $sync.SelectedServiceStatus.Text = ""
        $sync.SelectedServiceStartupType.Text = ""
        $sync.ToggleButton.Content = "Toggle Service"
        $sync.ToggleButton.IsEnabled = $false
        $sync.SetStartupTypeButton.IsEnabled = $false
    }
})

# Double-click to toggle a service
$sync.ServiceListView.Add_MouseDoubleClick({
    if ($sync.ServiceToggling) {
        [System.Windows.MessageBox]::Show("A service operation is already in progress", "Operation in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $selectedItem = $sync.ServiceListView.SelectedItem
    
    if ($selectedItem -ne $null) {
        Set-ServiceState -Service $selectedItem
        Update-ServiceList
    }
})

# Window closing event
$sync.Window.Add_Closing({
    if ($sync.ServiceToggling) {
        $result = [System.Windows.MessageBox]::Show(
            "A service operation is currently in progress. Are you sure you want to exit?", 
            "Confirm Exit", 
            [System.Windows.MessageBoxButton]::YesNo, 
            [System.Windows.MessageBoxImage]::Question)
            
        if ($result -eq [System.Windows.MessageBoxResult]::No) {
            $_.Cancel = $true
        }
    }
})

# Initial population of the service list
Update-ServiceList

# Default startup type selection
$sync.StartupTypeComboBox.SelectedIndex = 0

# Start with welcome message
Write-ServiceLog "Service Manager started. Found $($sync.ServiceListView.Items.Count) matching services." -Type Information

# Show the window
$sync.Window.ShowDialog() | Out-Null