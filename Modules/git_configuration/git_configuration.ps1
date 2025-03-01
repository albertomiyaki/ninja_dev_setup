# Modern Git Configuration Tool
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.ConfigUpdating = $false

# Check if Git is installed
function Test-GitInstalled {
    try {
        $gitVersion = git --version
        return $true
    }
    catch {
        return $false
    }
}

# Create log function for centralized logging
function Write-GitLog {
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

# Function to get current git configuration
function Get-GitConfig {
    $config = @{}
    
    try {
        # Get existing git config if available
        $userName = git config --global user.name 2>$null
        $userEmail = git config --global user.email 2>$null
        $editor = git config --global core.editor 2>$null
        $defaultBranch = git config --global init.defaultBranch 2>$null
        $credentialHelper = git config --global credential.helper 2>$null
        $autoCRLF = git config --global core.autocrlf 2>$null
        $eol = git config --global core.eol 2>$null
        $colorUI = git config --global color.ui 2>$null
        $pushDefault = git config --global push.default 2>$null
        
        # Get aliases
        $aliases = @{}
        $aliasOutput = git config --global --get-regexp "^alias\." 2>$null
        
        if ($aliasOutput) {
            foreach ($line in $aliasOutput) {
                if ($line -match "^alias\.(.+)\s(.+)$") {
                    $aliasName = $matches[1]
                    $aliasCommand = $matches[2]
                    $aliases[$aliasName] = $aliasCommand
                }
            }
        }
        
        # Create config object
        $config = @{
            UserName = if ([string]::IsNullOrWhiteSpace($userName)) { "Your Name" } else { $userName }
            UserEmail = if ([string]::IsNullOrWhiteSpace($userEmail)) { "your-email@example.com" } else { $userEmail }
            Editor = if ([string]::IsNullOrWhiteSpace($editor)) { "code --wait" } else { $editor }
            DefaultBranch = if ([string]::IsNullOrWhiteSpace($defaultBranch)) { "main" } else { $defaultBranch }
            CredentialHelper = if ([string]::IsNullOrWhiteSpace($credentialHelper)) { "manager-core" } else { $credentialHelper }
            AutoCRLF = if ([string]::IsNullOrWhiteSpace($autoCRLF)) { "true" } else { $autoCRLF }
            EOL = if ([string]::IsNullOrWhiteSpace($eol)) { "native" } else { $eol }
            ColorUI = if ([string]::IsNullOrWhiteSpace($colorUI)) { "auto" } else { $colorUI }
            PushDefault = if ([string]::IsNullOrWhiteSpace($pushDefault)) { "simple" } else { $pushDefault }
            Aliases = $aliases
        }
        
        Write-GitLog "Loaded existing Git configuration" -Type Success
    }
    catch {
        Write-GitLog "Error loading Git configuration: $_" -Type Error
        
        # Set defaults if we couldn't load the config
        $config = @{
            UserName = "Your Name"
            UserEmail = "your-email@example.com"
            Editor = "code --wait"
            DefaultBranch = "main"
            CredentialHelper = "manager-core"
            AutoCRLF = "true"
            ColorUI = "auto"
            PushDefault = "simple"
            Aliases = @{
                "st" = "status"
                "ci" = "commit -m"
                "co" = "checkout"
                "br" = "branch"
                "lg" = "log --oneline --graph --decorate"
            }
        }
    }
    
    return $config
}

# Function to set Git configuration
function Set-GitConfig {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        $sync.ConfigUpdating = $true
        $sync.SaveButton.IsEnabled = $false
        $sync.StatusBar.Text = "Updating Git configuration..."
        
        # Core settings
        Write-GitLog "Setting user.name to $($Config.UserName)" -Type Information
        git config --global user.name "$($Config.UserName)"
        
        Write-GitLog "Setting user.email to $($Config.UserEmail)" -Type Information
        git config --global user.email "$($Config.UserEmail)"
        
        Write-GitLog "Setting core.editor to $($Config.Editor)" -Type Information
        git config --global core.editor "$($Config.Editor)"
        
        Write-GitLog "Setting init.defaultBranch to $($Config.DefaultBranch)" -Type Information
        git config --global init.defaultBranch "$($Config.DefaultBranch)"
        
        Write-GitLog "Setting credential.helper to $($Config.CredentialHelper)" -Type Information
        git config --global credential.helper "$($Config.CredentialHelper)"
        
        Write-GitLog "Setting core.autocrlf to $($Config.AutoCRLF)" -Type Information
        git config --global core.autocrlf $($Config.AutoCRLF)
        
        Write-GitLog "Setting core.eol to $($Config.EOL)" -Type Information
        git config --global core.eol $($Config.EOL)
        
        Write-GitLog "Setting color.ui to $($Config.ColorUI)" -Type Information
        git config --global color.ui $($Config.ColorUI)
        
        Write-GitLog "Setting push.default to $($Config.PushDefault)" -Type Information
        git config --global push.default $($Config.PushDefault)
        
        # Aliases
        if ($Config.Aliases -and $Config.Aliases.Count -gt 0) {
            Write-GitLog "Configuring Git aliases" -Type Information
            
            # Remove all existing aliases first
            $existingAliases = git config --global --get-regexp "^alias\." 2>$null
            if ($existingAliases) {
                foreach ($line in $existingAliases) {
                    if ($line -match "^alias\.(.+)\s") {
                        $aliasName = $matches[1]
                        git config --global --unset "alias.$aliasName" 2>$null
                    }
                }
            }
            
            # Add aliases from our config
            foreach ($alias in $Config.Aliases.GetEnumerator()) {
                $aliasName = $alias.Key
                $aliasCommand = $alias.Value
                
                if (-not [string]::IsNullOrWhiteSpace($aliasName) -and -not [string]::IsNullOrWhiteSpace($aliasCommand)) {
                    Write-GitLog "Setting alias.$aliasName to '$aliasCommand'" -Type Information
                    git config --global alias.$aliasName "$aliasCommand"
                }
            }
        }
        
        # Show final configuration
        $finalConfig = git config --global --list
        Write-GitLog "Git configuration completed successfully!" -Type Success
        Write-GitLog "Final configuration:" -Type Information
        
        foreach ($line in $finalConfig) {
            Write-GitLog $line -Type Information
        }
        
        $sync.StatusBar.Text = "Git configuration updated successfully!"
        return $true
    }
    catch {
        Write-GitLog "Error setting Git configuration: $_" -Type Error
        $sync.StatusBar.Text = "Error updating Git configuration."
        return $false
    }
    finally {
        $sync.ConfigUpdating = $false
        $sync.SaveButton.IsEnabled = $true
    }
}

# Function to create a custom dialog for editing aliases
function Edit-GitAliases {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Aliases
    )
    
    # Create a copy of the aliases to work with
    $workingAliases = @{}
    foreach ($key in $Aliases.Keys) {
        $workingAliases[$key] = $Aliases[$key]
    }
    
    # Create the dialog
    $aliasDialog = New-Object System.Windows.Window
    $aliasDialog.Title = "Edit Git Aliases"
    $aliasDialog.Width = 500
    $aliasDialog.Height = 500
    $aliasDialog.WindowStartupLocation = "CenterOwner"
    $aliasDialog.Owner = $sync.Window
    
    # Create the main grid
    $mainGrid = New-Object System.Windows.Controls.Grid
    $aliasDialog.Content = $mainGrid
    
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto"}))
    
    # Create the ScrollViewer and StackPanel for aliases
    $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = "10,10,10,10"
    $scrollViewer.Content = $stackPanel
    [System.Windows.Controls.Grid]::SetRow($scrollViewer, 0)
    $mainGrid.Children.Add($scrollViewer)
    
    # Create the button panel
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Right"
    $buttonPanel.Margin = "10,0,10,10"
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
    $mainGrid.Children.Add($buttonPanel)
    
    # Helper function to create an alias row
    function Add-AliasRow {
        param (
            [string]$AliasName,
            [string]$AliasCommand,
            [bool]$IsNew = $false
        )
        
        $rowGrid = New-Object System.Windows.Controls.Grid
        $rowGrid.Margin = "0,5,0,5"
        
        $rowGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "Auto"}))
        $rowGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
        $rowGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*"}))
        $rowGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "Auto"}))
        
        # Label
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = "alias."
        $label.VerticalAlignment = "Center"
        $label.Margin = "0,0,5,0"
        [System.Windows.Controls.Grid]::SetColumn($label, 0)
        $rowGrid.Children.Add($label)
        
        # Alias name textbox
        $nameTextBox = New-Object System.Windows.Controls.TextBox
        $nameTextBox.Text = $AliasName
        $nameTextBox.Height = 30
        $nameTextBox.Padding = "5,5,5,5"
        $nameTextBox.Margin = "0,0,5,0"
        [System.Windows.Controls.Grid]::SetColumn($nameTextBox, 1)
        $rowGrid.Children.Add($nameTextBox)
        
        # Alias command textbox
        $commandTextBox = New-Object System.Windows.Controls.TextBox
        $commandTextBox.Text = $AliasCommand
        $commandTextBox.Height = 30
        $commandTextBox.Padding = "5,5,5,5"
        $commandTextBox.Margin = "0,0,5,0"
        [System.Windows.Controls.Grid]::SetColumn($commandTextBox, 2)
        $rowGrid.Children.Add($commandTextBox)
        
        # Delete button
        $deleteButton = New-Object System.Windows.Controls.Button
        $deleteButton.Content = "🗑️"
        $deleteButton.Width = 30
        $deleteButton.Height = 30
        $deleteButton.ToolTip = "Remove this alias"
        [System.Windows.Controls.Grid]::SetColumn($deleteButton, 3)
        $rowGrid.Children.Add($deleteButton)
        
        # Add event handler for delete button
        $deleteButton.Add_Click({
            $stackPanel.Children.Remove($rowGrid)
        })
        
        # Add the row to the stackpanel
        $stackPanel.Children.Add($rowGrid)
        
        # Return the controls for later reference
        return @{
            Row = $rowGrid
            NameTextBox = $nameTextBox
            CommandTextBox = $commandTextBox
            DeleteButton = $deleteButton
        }
    }
    
    # Add all existing aliases
    foreach ($alias in $workingAliases.GetEnumerator() | Sort-Object Key) {
        Add-AliasRow -AliasName $alias.Key -AliasCommand $alias.Value
    }
    
    # Add the "Add New Alias" button
    $addButton = New-Object System.Windows.Controls.Button
    $addButton.Content = "Add New Alias"
    $addButton.Padding = "10,5,10,5"
    $addButton.Margin = "0,0,5,0"
    $buttonPanel.Children.Add($addButton)
    
    # Add event handler for the add button
    $addButton.Add_Click({
        Add-AliasRow -AliasName "" -AliasCommand "" -IsNew $true
    })
    
    # Add Cancel button
    $cancelButton = New-Object System.Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Padding = "10,5,10,5"
    $cancelButton.Margin = "0,0,5,0"
    $buttonPanel.Children.Add($cancelButton)
    
    # Add event handler for cancel button
    $cancelButton.Add_Click({
        $aliasDialog.DialogResult = $false
        $aliasDialog.Close()
    })
    
    # Add Save button
    $saveButton = New-Object System.Windows.Controls.Button
    $saveButton.Content = "Save"
    $saveButton.Padding = "10,5,10,5"
    $buttonPanel.Children.Add($saveButton)
    
    # Add event handler for save button
    $saveButton.Add_Click({
        $newAliases = @{}
        
        foreach ($child in $stackPanel.Children) {
            $nameTextBox = $child.Children | Where-Object { [System.Windows.Controls.Grid]::GetColumn($_) -eq 1 }
            $commandTextBox = $child.Children | Where-Object { [System.Windows.Controls.Grid]::GetColumn($_) -eq 2 }
            
            if ($nameTextBox -and $commandTextBox) {
                $name = $nameTextBox.Text.Trim()
                $command = $commandTextBox.Text.Trim()
                
                if (-not [string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($command)) {
                    $newAliases[$name] = $command
                }
            }
        }
        
        # Update the original hashtable with our new values
        $Aliases.Clear()
        foreach ($key in $newAliases.Keys) {
            $Aliases[$key] = $newAliases[$key]
        }
        
        $aliasDialog.DialogResult = $true
        $aliasDialog.Close()
    })
    
    # Show the dialog
    return $aliasDialog.ShowDialog()
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Git Configuration Tool" 
    Height="700" 
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
        
        <Style TargetType="TextBox" x:Key="ConfigTextBox">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="#CCCCCC"
                                BorderThickness="1"
                                CornerRadius="3">
                            <ScrollViewer x:Name="PART_ContentHost" Focusable="false" />
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style TargetType="ComboBox" x:Key="ConfigComboBox">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="FontSize" Value="14"/>
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
                    <TextBlock Text="Git Configuration Tool" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Configure your Git settings with a modern UI" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="ReloadButton" Grid.Column="1" 
                        Style="{StaticResource DefaultButton}"
                        Content="⟳ Reload Settings" 
                        ToolTip="Reload current Git configuration"
                        VerticalAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <ScrollViewer Grid.Row="1" Margin="15" VerticalScrollBarVisibility="Auto">
            <StackPanel>
                <!-- User Info Section -->
                <Border Background="#F5F5F5" Padding="15" CornerRadius="3" Margin="0,0,0,15">
                    <StackPanel>
                        <TextBlock Text="User Information" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>

                        <TextBlock Text="Name" FontWeight="Medium" Margin="0,0,0,5"/>
                        <TextBox x:Name="UserNameTextBox" Style="{StaticResource ConfigTextBox}"
                                 Text="Your Name"/>
                        
                        <TextBlock Text="Email" FontWeight="Medium" Margin="0,0,0,5"/>
                        <TextBox x:Name="UserEmailTextBox" Style="{StaticResource ConfigTextBox}"
                                 Text="your-email@example.com"/>
                    </StackPanel>
                </Border>
                
                <!-- Editor Section -->
                <Border Background="#F5F5F5" Padding="15" CornerRadius="3" Margin="0,0,0,15">
                    <StackPanel>
                        <TextBlock Text="Editor Settings" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                        
                        <TextBlock Text="Default Text Editor" FontWeight="Medium" Margin="0,0,0,5"/>
                        <TextBox x:Name="EditorTextBox" Style="{StaticResource ConfigTextBox}"
                                 Text="code --wait"/>

                        <TextBlock FontStyle="Italic" TextWrapping="Wrap" Margin="0,0,0,10" Foreground="#666666">
                            Common editors: "code --wait" (VS Code), "notepad++.exe -multiInst -notabbar -nosession" (Notepad++), "notepad.exe" (Windows Notepad)
                        </TextBlock>
                    </StackPanel>
                </Border>
                
                <!-- Branch Settings -->
                <Border Background="#F5F5F5" Padding="15" CornerRadius="3" Margin="0,0,0,15">
                    <StackPanel>
                        <TextBlock Text="Branch Settings" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                        
                        <TextBlock Text="Default Branch Name" FontWeight="Medium" Margin="0,0,0,5"/>
                        <ComboBox x:Name="DefaultBranchComboBox" Style="{StaticResource ConfigComboBox}">
                            <ComboBoxItem Content="main"/>
                            <ComboBoxItem Content="master"/>
                            <ComboBoxItem Content="development"/>
                            <ComboBoxItem Content="trunk"/>
                        </ComboBox>
                    </StackPanel>
                </Border>
                
                <!-- System Settings -->
                <Border Background="#F5F5F5" Padding="15" CornerRadius="3" Margin="0,0,0,15">
                    <StackPanel>
                        <TextBlock Text="System Settings" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                        
                        <TextBlock Text="Credential Helper" FontWeight="Medium" Margin="0,0,0,5"/>
                        <ComboBox x:Name="CredentialHelperComboBox" Style="{StaticResource ConfigComboBox}">
                            <ComboBoxItem Content="manager-core"/>
                            <ComboBoxItem Content="manager"/>
                            <ComboBoxItem Content="wincred"/>
                            <ComboBoxItem Content="store"/>
                            <ComboBoxItem Content="cache"/>
                        </ComboBox>
                        
                        <TextBlock Text="Line Endings (autocrlf)" FontWeight="Medium" Margin="0,0,0,5"/>
                        <ComboBox x:Name="AutoCRLFComboBox" Style="{StaticResource ConfigComboBox}">
                            <ComboBoxItem Content="true"/>
                            <ComboBoxItem Content="false"/>
                            <ComboBoxItem Content="input"/>
                        </ComboBox>
                        
                        <TextBlock Text="End of Line (eol)" FontWeight="Medium" Margin="0,0,0,5"/>
                        <ComboBox x:Name="EOLComboBox" Style="{StaticResource ConfigComboBox}">
                            <ComboBoxItem Content="lf"/>
                            <ComboBoxItem Content="crlf"/>
                            <ComboBoxItem Content="native"/>
                        </ComboBox>
                        
                        <TextBlock Text="Color UI" FontWeight="Medium" Margin="0,0,0,5"/>
                        <ComboBox x:Name="ColorUIComboBox" Style="{StaticResource ConfigComboBox}">
                            <ComboBoxItem Content="auto"/>
                            <ComboBoxItem Content="always"/>
                            <ComboBoxItem Content="false"/>
                        </ComboBox>
                        
                        <TextBlock Text="Push Default" FontWeight="Medium" Margin="0,0,0,5"/>
                        <ComboBox x:Name="PushDefaultComboBox" Style="{StaticResource ConfigComboBox}">
                            <ComboBoxItem Content="simple"/>
                            <ComboBoxItem Content="current"/>
                            <ComboBoxItem Content="upstream"/>
                            <ComboBoxItem Content="matching"/>
                            <ComboBoxItem Content="nothing"/>
                        </ComboBox>
                    </StackPanel>
                </Border>
                
                <!-- Aliases Section -->
                <Border Background="#F5F5F5" Padding="15" CornerRadius="3" Margin="0,0,0,15">
                    <StackPanel>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Git Aliases" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                            <Button Grid.Column="1" x:Name="EditAliasesButton" Content="Edit Aliases" Style="{StaticResource DefaultButton}"/>
                        </Grid>
                        
                        <TextBlock x:Name="AliasesTextBlock" TextWrapping="Wrap" Margin="0,10,0,0"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>
        
        <!-- Save Button -->
        <Border Grid.Row="2" Background="#F5F5F5" Padding="15" BorderThickness="0,1,0,1" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="GitStatusText" Text="Ready to update Git configuration" FontWeight="Bold" FontSize="14"/>
                    <TextBlock x:Name="GitVersionText" Text="" Foreground="#666666" FontSize="12" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="SaveButton" 
                        Grid.Column="1" 
                        Content="Save Configuration" 
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
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Store references in the sync hashtable
$sync.Window = $window
$sync.ReloadButton = $window.FindName("ReloadButton")
$sync.SaveButton = $window.FindName("SaveButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.GitStatusText = $window.FindName("GitStatusText")
$sync.GitVersionText = $window.FindName("GitVersionText")

# Form fields
$sync.UserNameTextBox = $window.FindName("UserNameTextBox")
$sync.UserEmailTextBox = $window.FindName("UserEmailTextBox")
$sync.EditorTextBox = $window.FindName("EditorTextBox")
$sync.DefaultBranchComboBox = $window.FindName("DefaultBranchComboBox")
$sync.CredentialHelperComboBox = $window.FindName("CredentialHelperComboBox")
$sync.AutoCRLFComboBox = $window.FindName("AutoCRLFComboBox")
$sync.ColorUIComboBox = $window.FindName("ColorUIComboBox")
$sync.PushDefaultComboBox = $window.FindName("PushDefaultComboBox")
$sync.AliasesTextBlock = $window.FindName("AliasesTextBlock")
$sync.EditAliasesButton = $window.FindName("EditAliasesButton")

# Check if Git is installed
if (-not (Test-GitInstalled)) {
    Write-GitLog "Git is not installed! This tool requires Git to be installed." -Type Error
    [System.Windows.MessageBox]::Show(
        "Git is not installed or not in your PATH. Please install Git and try again.",
        "Git Not Found",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    exit
}

# Get Git version
$gitVersion = git --version
$sync.GitVersionText.Text = $gitVersion

# Load current Git configuration
$gitConfig = Get-GitConfig

# Initialize the form with current values
function Initialize-Form {
    # User Info
    $sync.UserNameTextBox.Text = $gitConfig.UserName
    $sync.UserEmailTextBox.Text = $gitConfig.UserEmail
    
    # Editor
    $sync.EditorTextBox.Text = $gitConfig.Editor
    
    # Branch settings
    $defaultBranchItem = $sync.DefaultBranchComboBox.Items | Where-Object { $_.Content -eq $gitConfig.DefaultBranch } | Select-Object -First 1
    if ($defaultBranchItem) {
        $sync.DefaultBranchComboBox.SelectedItem = $defaultBranchItem
    } else {
        $sync.DefaultBranchComboBox.Text = $gitConfig.DefaultBranch
    }
    
    # System settings
    $credentialHelperItem = $sync.CredentialHelperComboBox.Items | Where-Object { $_.Content -eq $gitConfig.CredentialHelper } | Select-Object -First 1
    if ($credentialHelperItem) {
        $sync.CredentialHelperComboBox.SelectedItem = $credentialHelperItem
    } else {
        $sync.CredentialHelperComboBox.Text = $gitConfig.CredentialHelper
    }
    
    $autoCRLFItem = $sync.AutoCRLFComboBox.Items | Where-Object { $_.Content -eq $gitConfig.AutoCRLF } | Select-Object -First 1
    if ($autoCRLFItem) {
        $sync.AutoCRLFComboBox.SelectedItem = $autoCRLFItem
    } else {
        $sync.AutoCRLFComboBox.Text = $gitConfig.AutoCRLF
    }
    
    $eolItem = $sync.EOLComboBox.Items | Where-Object { $_.Content -eq $gitConfig.EOL } | Select-Object -First 1
    if ($eolItem) {
        $sync.EOLComboBox.SelectedItem = $eolItem
    } else {
        # Add a new item if the value is not in the predefined list
        $newItem = New-Object System.Windows.Controls.ComboBoxItem
        $newItem.Content = $gitConfig.EOL
        $sync.EOLComboBox.Items.Add($newItem)
        $sync.EOLComboBox.SelectedItem = $newItem
    }
    
    $colorUIItem = $sync.ColorUIComboBox.Items | Where-Object { $_.Content -eq $gitConfig.ColorUI } | Select-Object -First 1
    if ($colorUIItem) {
        $sync.ColorUIComboBox.SelectedItem = $colorUIItem
    } else {
        $sync.ColorUIComboBox.Text = $gitConfig.ColorUI
    }
    
    $pushDefaultItem = $sync.PushDefaultComboBox.Items | Where-Object { $_.Content -eq $gitConfig.PushDefault } | Select-Object -First 1
    if ($pushDefaultItem) {
        $sync.PushDefaultComboBox.SelectedItem = $pushDefaultItem
    } else {
        $sync.PushDefaultComboBox.Text = $gitConfig.PushDefault
    }
    
    # Update aliases display
    Update-AliasesDisplay
    
    Write-GitLog "Form initialized with current Git configuration" -Type Information
}

# Update aliases display
function Update-AliasesDisplay {
    $aliasList = @()
    
    foreach ($alias in $gitConfig.Aliases.GetEnumerator() | Sort-Object Key) {
        $aliasList += "alias.$($alias.Key) = $($alias.Value)"
    }
    
    if ($aliasList.Count -eq 0) {
        $sync.AliasesTextBlock.Text = "No aliases configured."
    } else {
        $sync.AliasesTextBlock.Text = $aliasList -join "`r`n"
    }
}

# Set up event handlers
$sync.ReloadButton.Add_Click({
    $global:gitConfig = Get-GitConfig
    Initialize-Form
    $sync.StatusBar.Text = "Git configuration reloaded"
})

$sync.EditAliasesButton.Add_Click({
    $result = Edit-GitAliases -Aliases $gitConfig.Aliases
    if ($result) {
        Update-AliasesDisplay
        $sync.StatusBar.Text = "Aliases updated"
    }
})

$sync.SaveButton.Add_Click({
    if ($sync.ConfigUpdating) {
        [System.Windows.MessageBox]::Show("Configuration is already being updated", "Operation in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Confirm changes
    $confirmResult = [System.Windows.MessageBox]::Show(
        "Are you sure you want to apply these Git configuration changes?",
        "Confirm Configuration",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($confirmResult -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-GitLog "Configuration update cancelled by user" -Type Information
        return
    }
    
    # Get values from form
    $updatedConfig = @{
        UserName = $sync.UserNameTextBox.Text
        UserEmail = $sync.UserEmailTextBox.Text
        Editor = $sync.EditorTextBox.Text
        DefaultBranch = if ($sync.DefaultBranchComboBox.SelectedItem) { $sync.DefaultBranchComboBox.SelectedItem.Content } else { $sync.DefaultBranchComboBox.Text }
        CredentialHelper = if ($sync.CredentialHelperComboBox.SelectedItem) { $sync.CredentialHelperComboBox.SelectedItem.Content } else { $sync.CredentialHelperComboBox.Text }
        AutoCRLF = if ($sync.AutoCRLFComboBox.SelectedItem) { $sync.AutoCRLFComboBox.SelectedItem.Content } else { $sync.AutoCRLFComboBox.Text }
        EOL = if ($sync.EOLComboBox.SelectedItem) { $sync.EOLComboBox.SelectedItem.Content } else { $sync.EOLComboBox.Text }
        ColorUI = if ($sync.ColorUIComboBox.SelectedItem) { $sync.ColorUIComboBox.SelectedItem.Content } else { $sync.ColorUIComboBox.Text }
        PushDefault = if ($sync.PushDefaultComboBox.SelectedItem) { $sync.PushDefaultComboBox.SelectedItem.Content } else { $sync.PushDefaultComboBox.Text }
        Aliases = $gitConfig.Aliases
    }
    
    # Update the configuration
    $result = Set-GitConfig -Config $updatedConfig
    
    if ($result) {
        # Update our gitConfig object with the new values
        $global:gitConfig = $updatedConfig
        
        # Update form
        Initialize-Form
    }
})

# Initialize the form
Initialize-Form

# Show the window
$sync.Window.ShowDialog() | Out-Null