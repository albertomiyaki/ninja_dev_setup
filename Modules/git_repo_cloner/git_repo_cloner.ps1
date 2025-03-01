# GitHub Enterprise Repository Cloner
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Security

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.CloningInProgress = $false
$sync.JobQueue = @()
$sync.ClonedRepos = @()
$sync.SkippedRepos = @()
$sync.FailedRepos = @()
$sync.CancellationRequested = $false

# Check if Git is installed
function Test-GitInstalled {
    try {
        $gitVersion = git --version
        return $true
    } catch {
        return $false
    }
}

# Create log function for centralized logging
function Write-CloneLog {
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
    
    # Write to log file if configured
    if ($sync.LogFile) {
        Add-Content -Path $sync.LogFile -Value $logMessage
    }
}

# Function to convert SecureString to plain text (for API calls)
function ConvertFrom-SecureToPlain {
    param(
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$SecureString
    )
    
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    
    return $plainText
}

# Function to get repositories from GitHub Enterprise
function Get-GitHubRepositories {
    param(
        [string]$ApiUrl,
        [System.Security.SecureString]$Token
    )
    
    $allRepos = @()
    $page = 1
    $perPage = 100  # GitHub API returns max 100 per request
    
    try {
        $plainToken = ConvertFrom-SecureToPlain -SecureString $Token
        $headers = @{ 
            Authorization = "token $plainToken"
            Accept = "application/vnd.github+json"
        }
        
        Write-CloneLog "Fetching repositories from $ApiUrl" -Type Information
        $sync.StatusBar.Text = "Fetching repositories..."
        
        while (-not $sync.CancellationRequested) {
            $url = "$ApiUrl`?page=$page&per_page=$perPage"
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
            
            if ($response.Count -eq 0) { 
                break  # Stop if no more repos
            }
            
            $allRepos += $response
            Write-CloneLog "Fetched page $page ($($response.Count) repositories)" -Type Information
            $page++
        }
        
        return $allRepos | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.name
                CloneUrl = $_.clone_url
                Description = $_.description
                Language = $_.language
                IsPrivate = $_.private
                UpdatedAt = [DateTime]$_.updated_at
                Size = $_.size
            }
        }
    } 
    catch {
        Write-CloneLog "Error fetching repositories: $_" -Type Error
        $sync.StatusBar.Text = "Error fetching repositories"
        throw $_
    }
}

# Function to update the UI with jobs progress
function Update-JobStatus {
    # Count active jobs
    $runningJobs = $sync.JobQueue | Where-Object { $_.Job.State -eq "Running" }
    $completedJobs = $sync.JobQueue | Where-Object { $_.Job.State -eq "Completed" } 
    $failedJobs = $sync.JobQueue | Where-Object { $_.Job.State -eq "Failed" }
    
    # Update progress
    $totalJobs = $sync.TotalReposToClone
    $processedJobs = ($sync.ClonedRepos.Count + $sync.FailedRepos.Count)
    
    if ($totalJobs -gt 0) {
        $progressPercentage = [Math]::Min(100, [Math]::Round(($processedJobs / $totalJobs) * 100))
        $sync.ProgressBar.Value = $progressPercentage
    } else {
        $sync.ProgressBar.Value = 0
    }
    
    # Update status text
    $sync.CloneStatusTextBlock.Text = "Cloning: $($runningJobs.Count) active, " + 
                                      "$($sync.ClonedRepos.Count) completed, " + 
                                      "$($sync.SkippedRepos.Count) skipped, " + 
                                      "$($sync.FailedRepos.Count) failed"
    
    # Update status bar
    if ($sync.CloningInProgress) {
        $sync.StatusBar.Text = "Cloning repositories ($progressPercentage% complete)..."
    }
}

# Process completed jobs
function Process-CompletedJobs {
    $jobsToRemove = @()
    
    foreach ($jobItem in $sync.JobQueue) {
        if ($jobItem.Job.State -eq "Completed") {
            $result = Receive-Job -Job $jobItem.Job -Wait -ErrorAction SilentlyContinue
            $repoName = $jobItem.Name
            
            if ($result -match "SUCCESS") {
                Write-CloneLog "Successfully cloned '$repoName'" -Type Success
                $sync.ClonedRepos += $repoName
                
                # Update repository list status
                $repoItem = $sync.RepoListView.Items | Where-Object { $_.Name -eq $repoName }
                if ($repoItem) { 
                    $repoItem.Status = "Cloned"
                    $sync.RepoListView.Items.Refresh()
                }
            } else {
                Write-CloneLog "Failed to clone '$repoName': $result" -Type Error
                $sync.FailedRepos += $repoName
                
                # Update repository list status
                $repoItem = $sync.RepoListView.Items | Where-Object { $_.Name -eq $repoName }
                if ($repoItem) { 
                    $repoItem.Status = "Failed"
                    $sync.RepoListView.Items.Refresh()
                }
            }
            
            # Add job to remove list
            $jobsToRemove += $jobItem
        }
    }
    
    # Remove completed jobs from queue
    foreach ($item in $jobsToRemove) {
        Remove-Job -Job $item.Job -Force
        $sync.JobQueue = $sync.JobQueue | Where-Object { $_.Job.Id -ne $item.Job.Id }
    }
    
    # Update UI
    Update-JobStatus
    
    # Check if we're done
    if ($sync.JobQueue.Count -eq 0 -and $sync.CloningInProgress) {
        Complete-CloneProcess
    }
}

# Start cloning repositories
function Start-RepoCloning {
    param(
        [array]$Repositories,
        [string]$ClonePath,
        [System.Security.SecureString]$Token,
        [int]$MaxParallelJobs = 5
    )
    
    if ($sync.CloningInProgress) {
        Write-CloneLog "Cloning operation already in progress" -Type Warning
        return
    }
    
    if ($Repositories.Count -eq 0) {
        Write-CloneLog "No repositories selected for cloning" -Type Warning
        return
    }
    
    # Create log file
    $sync.LogFile = Join-Path $ClonePath "clone_log.txt"
    New-Item -ItemType File -Path $sync.LogFile -Force | Out-Null
    Add-Content -Path $sync.LogFile -Value "`n--- GitHub Repository Cloning Log ---`n$(Get-Date) `n"
    
    $plainToken = ConvertFrom-SecureToPlain -SecureString $Token
    
    # Initialize the process
    $sync.CloningInProgress = $true
    $sync.CancellationRequested = $false
    $sync.ClonedRepos = @()
    $sync.SkippedRepos = @()
    $sync.FailedRepos = @()
    $sync.JobQueue = @()
    $sync.TotalReposToClone = $Repositories.Count
    
    # Update UI
    $sync.CloneButton.Content = "Cancel Cloning"
    $sync.ProgressBar.Value = 0
    $sync.StatusBar.Text = "Preparing to clone repositories..."
    
    # Make sure the destination directory exists
    if (-not (Test-Path $ClonePath)) {
        New-Item -ItemType Directory -Path $ClonePath -Force | Out-Null
        Write-CloneLog "Created directory: $ClonePath" -Type Information
    }
    
    Write-CloneLog "Starting clone process for $($Repositories.Count) repositories" -Type Information
    
    # Clone each repository
    foreach ($repo in $Repositories) {
        if ($sync.CancellationRequested) {
            Write-CloneLog "Cloning process cancelled by user" -Type Warning
            break
        }
        
        $repoName = $repo.Name
        $repoPath = Join-Path $ClonePath $repoName
        
        # Check if repo already exists
        if (Test-Path $repoPath) {
            Write-CloneLog "Repository '$repoName' already exists. Skipping..." -Type Warning
            $sync.SkippedRepos += $repoName
            
            # Update repository list status
            $repoItem = $sync.RepoListView.Items | Where-Object { $_.Name -eq $repoName }
            if ($repoItem) { 
                $repoItem.Status = "Skipped"
                $sync.RepoListView.Items.Refresh()
            }
            
            continue
        }
        
        # Wait until we have less than max parallel jobs
        while (($sync.JobQueue | Where-Object { $_.Job.State -eq "Running" }).Count -ge $MaxParallelJobs) {
            Start-Sleep -Milliseconds 500
            Process-CompletedJobs
            
            if ($sync.CancellationRequested) {
                break
            }
        }
        
        if ($sync.CancellationRequested) {
            break
        }
        
        # Mark repository as in progress
        $repoItem = $sync.RepoListView.Items | Where-Object { $_.Name -eq $repoName }
        if ($repoItem) { 
            $repoItem.Status = "Cloning"
            $sync.RepoListView.Items.Refresh()
        }
        
        # Start clone job
        Write-CloneLog "Starting clone job for '$repoName'" -Type Information
        
        $job = Start-Job -ScriptBlock {
            param ($repoUrl, $repoPath, $token)
            
            # Set git credentials in the process
            $env:GIT_ASKPASS = "echo"
            $env:GIT_USERNAME = "x-access-token"
            $env:GIT_PASSWORD = $token
            
            # Create directory
            New-Item -ItemType Directory -Path $repoPath -Force | Out-Null
            
            # Clone repository
            $output = git clone $repoUrl $repoPath 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                return "SUCCESS"
            } else {
                return "FAILED: $output"
            }
        } -ArgumentList $repo.CloneUrl, $repoPath, $plainToken
        
        # Add to job queue
        $sync.JobQueue += [PSCustomObject]@{
            Name = $repoName
            Job = $job
        }
        
        # Update UI
        Update-JobStatus
    }
    
    # Set up a timer to update the UI and check for completed jobs
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(1)
    
    $timer.Add_Tick({
        Process-CompletedJobs
        
        # If all jobs are done or cancellation requested, clean up
        if (($sync.JobQueue.Count -eq 0 -or $sync.CancellationRequested) -and $sync.CloningInProgress) {
            $timer.Stop()
            Complete-CloneProcess
        }
    })
    
    $timer.Start()
    $sync.Timer = $timer
}

# Complete the cloning process
function Complete-CloneProcess {
    $sync.CloningInProgress = $false
    
    # Stop and clean up any remaining jobs
    foreach ($jobItem in $sync.JobQueue) {
        if ($jobItem.Job.State -eq "Running") {
            Stop-Job -Job $jobItem.Job
        }
        Remove-Job -Job $jobItem.Job -Force
    }
    
    $sync.JobQueue = @()
    
    # Update UI
    $sync.CloneButton.Content = "Clone Selected"
    $sync.ProgressBar.Value = 100
    
    # Log summary
    Write-CloneLog "Clone process completed" -Type Success
    Write-CloneLog "Cloned: $($sync.ClonedRepos.Count) repositories" -Type Success
    Write-CloneLog "Skipped: $($sync.SkippedRepos.Count) repositories" -Type Information
    Write-CloneLog "Failed: $($sync.FailedRepos.Count) repositories" -Type Error
    
    $sync.StatusBar.Text = "Clone process completed: $($sync.ClonedRepos.Count) cloned, $($sync.SkippedRepos.Count) skipped, $($sync.FailedRepos.Count) failed"
    
    # Update the clone status text
    $sync.CloneStatusTextBlock.Text = "Completed: " + 
                                      "$($sync.ClonedRepos.Count) cloned, " + 
                                      "$($sync.SkippedRepos.Count) skipped, " + 
                                      "$($sync.FailedRepos.Count) failed"
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="GitHub Enterprise Repository Cloner" 
    Height="750" 
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
        
        <Style TargetType="PasswordBox" x:Key="ConfigPasswordBox">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type PasswordBox}">
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
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
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
                    <TextBlock Text="GitHub Enterprise Repository Cloner" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Clone repositories from your GitHub Enterprise instance" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="FetchButton" Grid.Column="1" 
                        Style="{StaticResource DefaultButton}"
                        Content="Fetch Repositories" 
                        ToolTip="Fetch repositories from GitHub Enterprise"
                        VerticalAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Configuration Panel -->
        <Border Grid.Row="1" Background="#F5F5F5" Padding="15" BorderThickness="0,0,0,1" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <!-- GitHub Enterprise Settings -->
                <StackPanel Grid.Column="0" Margin="0,0,10,0">
                    <TextBlock Text="GitHub Enterprise Settings" FontWeight="Bold" Margin="0,0,0,10"/>
                    
                    <TextBlock Text="Host URL" FontWeight="Medium" Margin="0,0,0,5"/>
                    <TextBox x:Name="HostUrlTextBox" Style="{StaticResource ConfigTextBox}" 
                             Text="https://github.example.com"/>
                    
                    <TextBlock Text="API Version" FontWeight="Medium" Margin="0,0,0,5"/>
                    <TextBox x:Name="ApiVersionTextBox" Style="{StaticResource ConfigTextBox}" 
                             Text="api/v3"/>
                    
                    <TextBlock Text="Organization" FontWeight="Medium" Margin="0,0,0,5"/>
                    <TextBox x:Name="OrgNameTextBox" Style="{StaticResource ConfigTextBox}" 
                             Text="your-organization"/>
                    
                    <TextBlock Text="Access Token" FontWeight="Medium" Margin="0,0,0,5"/>
                    <PasswordBox x:Name="TokenPasswordBox" Style="{StaticResource ConfigPasswordBox}"/>
                </StackPanel>
                
                <!-- Clone Settings -->
                <StackPanel Grid.Column="1" Margin="10,0,10,0">
                    <TextBlock Text="Clone Settings" FontWeight="Bold" Margin="0,0,0,10"/>
                    
                    <TextBlock Text="Clone Path" FontWeight="Medium" Margin="0,0,0,5"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="ClonePathTextBox" Style="{StaticResource ConfigTextBox}" 
                                 Text="C:\GitHubRepos" Grid.Column="0"/>
                        <Button x:Name="BrowseButton" Content="Browse" Grid.Column="1" Padding="8,5" Margin="5,0,0,10"/>
                    </Grid>
                    
                    <TextBlock Text="Max Parallel Jobs" FontWeight="Medium" Margin="0,0,0,5"/>
                    <TextBox x:Name="MaxJobsTextBox" Style="{StaticResource ConfigTextBox}" 
                             Text="5"/>
                    
                    <TextBlock Text="Repository Filter" FontWeight="Medium" Margin="0,0,0,5"/>
                    <TextBox x:Name="RepoFilterTextBox" Style="{StaticResource ConfigTextBox}" 
                             Text="" PlaceholderText="Filter repositories by name"/>
                </StackPanel>
                
                <!-- Advanced Settings -->
                <StackPanel Grid.Column="2" Margin="10,0,0,0">
                    <TextBlock Text="Advanced Settings" FontWeight="Bold" Margin="0,0,0,10"/>
                    
                    <CheckBox x:Name="IncludePrivateCheckBox" Content="Include Private Repositories" IsChecked="True" Margin="0,0,0,10"/>
                    <CheckBox x:Name="IncludePublicCheckBox" Content="Include Public Repositories" IsChecked="True" Margin="0,0,0,10"/>
                    <CheckBox x:Name="SkipExistingCheckBox" Content="Skip Existing Repositories" IsChecked="True" Margin="0,0,0,10"/>
                    <CheckBox x:Name="CreateLogFileCheckBox" Content="Create Log File" IsChecked="True" Margin="0,0,0,10"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Search and Filter -->
        <Grid Grid.Row="2" Margin="15,15,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Search Box -->
            <Grid Grid.Row="0" Margin="0,0,0,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="🔍 Search: " VerticalAlignment="Center" Margin="0,0,10,0" FontSize="14"/>
                <TextBox x:Name="SearchBox" Grid.Column="1" Padding="8" FontSize="14" BorderThickness="1" BorderBrush="#CCCCCC">
                    <TextBox.Resources>
                        <Style TargetType="{x:Type Border}">
                            <Setter Property="CornerRadius" Value="3"/>
                        </Style>
                    </TextBox.Resources>
                </TextBox>
                <Button x:Name="SelectAllButton" Grid.Column="2" Content="Select All" Margin="10,0,0,0" Padding="8,5"/>
            </Grid>
            
            <!-- Repository List -->
            <ListView x:Name="RepoListView" Grid.Row="1" 
                      BorderThickness="1" 
                      BorderBrush="#CCCCCC"
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
                        <GridViewColumn Header="Name" Width="200">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Name}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Description" Width="400">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Description}" TextWrapping="Wrap" MaxWidth="390"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Language" Width="100">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Language}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Private" Width="70">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock>
                                        <TextBlock.Text>
                                            <Binding Path="IsPrivate" Converter="{StaticResource BooleanToStringConverter}"/>
                                        </TextBlock.Text>
                                    </TextBlock>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Updated" Width="120">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding UpdatedAt, StringFormat=\{0:yyyy-MM-dd\}}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Status" Width="80">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Status}" TextWrapping="NoWrap">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Status}" Value="Cloned">
                                                        <Setter Property="Foreground" Value="Green"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Cloning">
                                                        <Setter Property="Foreground" Value="Blue"/>
                                                        <Setter Property="FontWeight" Value="Bold"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Failed">
                                                        <Setter Property="Foreground" Value="Red"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Skipped">
                                                        <Setter Property="Foreground" Value="Orange"/>
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
        
        <!-- Progress Bar -->
        <Grid Grid.Row="3" Margin="15,15,15,0">
            <ProgressBar x:Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Value="0" />
        </Grid>
        
        <!-- Clone Status -->
        <Border Grid.Row="4" Background="#F5F5F5" Padding="15" BorderThickness="0,1,0,1" BorderBrush="#DDDDDD" Margin="0,15,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="SelectedRepoCount" Text="No repositories selected" FontWeight="Bold" FontSize="14"/>
                    <TextBlock x:Name="CloneStatusTextBlock" Text="Ready to clone" Foreground="#666666" FontSize="12" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="CloneButton" 
                        Grid.Column="1" 
                        Content="Clone Selected" 
                        Style="{StaticResource DefaultButton}"
                        FontSize="14"/>
            </Grid>
        </Border>
        
        <!-- Log Panel -->
        <Grid Grid.Row="5" Margin="15,15,15,0">
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
        <Border Grid.Row="6" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
            <TextBlock x:Name="StatusBar" Text="Ready" FontSize="12"/>
        </Border>
    </Grid>
</Window>
"@

# Create a boolean to string converter for the UI
$convertToString = [ScriptBlock]{ 
    param($value) 
    if ($value -eq $true) { return "Yes" } else { return "No" }
}

$resources = New-Object System.Windows.ResourceDictionary
$resources.Add("BooleanToStringConverter", 
    [ValueConverter]::Create($convertToString)
)

# Create a form object from the XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.Resources = $resources
}
catch {
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Check if Git is installed
if (-not (Test-GitInstalled)) {
    [System.Windows.MessageBox]::Show(
        "Git is not installed or not in your PATH. Please install Git and try again.",
        "Git Not Found",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    exit
}

# Store references in the sync hashtable
$sync.Window = $window
$sync.FetchButton = $window.FindName("FetchButton")
$sync.CloneButton = $window.FindName("CloneButton")
$sync.ProgressBar = $window.FindName("ProgressBar")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.RepoListView = $window.FindName("RepoListView")
$sync.SearchBox = $window.FindName("SearchBox")
$sync.SelectAllButton = $window.FindName("SelectAllButton")
$sync.BrowseButton = $window.FindName("BrowseButton")
$sync.SelectedRepoCount = $window.FindName("SelectedRepoCount")
$sync.CloneStatusTextBlock = $window.FindName("CloneStatusTextBlock")

# Form fields
$sync.HostUrlTextBox = $window.FindName("HostUrlTextBox")
$sync.ApiVersionTextBox = $window.FindName("ApiVersionTextBox")
$sync.OrgNameTextBox = $window.FindName("OrgNameTextBox")
$sync.TokenPasswordBox = $window.FindName("TokenPasswordBox")
$sync.ClonePathTextBox = $window.FindName("ClonePathTextBox")
$sync.MaxJobsTextBox = $window.FindName("MaxJobsTextBox")
$sync.RepoFilterTextBox = $window.FindName("RepoFilterTextBox")
$sync.IncludePrivateCheckBox = $window.FindName("IncludePrivateCheckBox")
$sync.IncludePublicCheckBox = $window.FindName("IncludePublicCheckBox")
$sync.SkipExistingCheckBox = $window.FindName("SkipExistingCheckBox")
$sync.CreateLogFileCheckBox = $window.FindName("CreateLogFileCheckBox")

# Class to help with ValueConverter
Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Data;
    using System.Globalization;

    public static class ValueConverter
    {
        public static IValueConverter Create(Func<object, object> convert)
        {
            return new DelegateConverter(convert);
        }

        private class DelegateConverter : IValueConverter
        {
            private readonly Func<object, object> _convert;

            public DelegateConverter(Func<object, object> convert)
            {
                _convert = convert;
            }

            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                return _convert(value);
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                throw new NotImplementedException();
            }
        }
    }
"@

# Function to update the selected repo count
function Update-SelectedRepoCount {
    $selectedCount = $sync.RepoListView.SelectedItems.Count
    
    if ($selectedCount -eq 0) {
        $sync.SelectedRepoCount.Text = "No repositories selected"
        $sync.CloneButton.IsEnabled = $false
    } else {
        $sync.SelectedRepoCount.Text = "$selectedCount repositories selected"
        $sync.CloneButton.IsEnabled = -not $sync.CloningInProgress
    }
}

# Function to filter repositories based on search text
function Update-RepoSearch {
    $searchText = $sync.SearchBox.Text.ToLower()
    $allRepos = $sync.AllRepositories
    
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        $filteredRepos = $allRepos
    } else {
        $filteredRepos = $allRepos | Where-Object { 
            $_.Name.ToLower().Contains($searchText) -or 
            ($_.Description -ne $null -and $_.Description.ToLower().Contains($searchText)) -or
            ($_.Language -ne $null -and $_.Language.ToLower().Contains($searchText))
        }
    }
    
    # Apply privacy filter
    if ($sync.IncludePrivateCheckBox.IsChecked -eq $false) {
        $filteredRepos = $filteredRepos | Where-Object { $_.IsPrivate -eq $false }
    }
    
    if ($sync.IncludePublicCheckBox.IsChecked -eq $false) {
        $filteredRepos = $filteredRepos | Where-Object { $_.IsPrivate -eq $true }
    }
    
    # Update the ListView
    $sync.RepoListView.Items.Clear()
    
    foreach ($repo in $filteredRepos) {
        # Add status property
        $repo | Add-Member -NotePropertyName Status -NotePropertyValue "Ready" -Force
        $sync.RepoListView.Items.Add($repo)
    }
    
    $sync.StatusBar.Text = "Showing $($filteredRepos.Count) of $($allRepos.Count) repositories"
}

# Set up event handlers
$sync.FetchButton.Add_Click({
    if ($sync.CloningInProgress) {
        [System.Windows.MessageBox]::Show("Please wait for the current cloning operation to complete.", "Operation in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Get values from form
    $hostUrl = $sync.HostUrlTextBox.Text.TrimEnd('/')
    $apiVersion = $sync.ApiVersionTextBox.Text.Trim('/')
    $orgName = $sync.OrgNameTextBox.Text.Trim()
    $token = $sync.TokenPasswordBox.SecurePassword
    
    if ([string]::IsNullOrWhiteSpace($hostUrl) -or [string]::IsNullOrWhiteSpace($orgName)) {
        [System.Windows.MessageBox]::Show("Please enter the GitHub Host URL and Organization name.", "Missing Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if ($token.Length -eq 0) {
        [System.Windows.MessageBox]::Show("Please enter your GitHub Personal Access Token.", "Missing Token", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Construct API URL
    $apiUrl = "$hostUrl/$apiVersion/orgs/$orgName/repos"
    
    $sync.RepoListView.Items.Clear()
    $sync.FetchButton.IsEnabled = $false
    $sync.StatusBar.Text = "Fetching repositories from $apiUrl..."
    
    # Use a background job to fetch repositories
    $job = Start-Job -ScriptBlock {
        param($apiUrl, $tokenStr)
        
        # Convert secure string back to secure string (needed for job)
        $secureToken = ConvertTo-SecureString $tokenStr -AsPlainText -Force
        
        try {
            # Add type definition for secure string conversion
            Add-Type -AssemblyName System.Security
            
            # Create helper function in the job context
            function ConvertFrom-JobSecureString {
                param([System.Security.SecureString]$secureString)
                $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
                $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                return $plainText
            }
            
            $plainToken = ConvertFrom-JobSecureString -secureString $secureToken
            $headers = @{ 
                Authorization = "token $plainToken"
                Accept = "application/vnd.github+json"
            }
            
            $allRepos = @()
            $page = 1
            $perPage = 100
            
            while ($true) {
                $url = "$apiUrl`?page=$page&per_page=$perPage"
                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
                
                if ($response.Count -eq 0) { 
                    break
                }
                
                $allRepos += $response
                $page++
            }
            
            return $allRepos | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.name
                    CloneUrl = $_.clone_url
                    Description = $_.description
                    Language = $_.language
                    IsPrivate = $_.private
                    UpdatedAt = [DateTime]$_.updated_at
                    Size = $_.size
                }
            }
        } 
        catch {
            throw $_
        }
    } -ArgumentList $apiUrl, (ConvertFrom-SecureToPlain -SecureString $token)
    
    # Set up a timer to check job status
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(1)
    
    $timer.Add_Tick({
        if ($job.State -eq "Completed") {
            $timer.Stop()
            
            try {
                $repos = Receive-Job -Job $job -ErrorAction Stop
                
                # Store all repositories
                $sync.AllRepositories = $repos
                
                # Update UI
                Update-RepoSearch
                
                Write-CloneLog "Retrieved $($repos.Count) repositories from GitHub Enterprise" -Type Success
                $sync.StatusBar.Text = "Retrieved $($repos.Count) repositories"
            } 
            catch {
                Write-CloneLog "Error retrieving repositories: $_" -Type Error
                $sync.StatusBar.Text = "Error retrieving repositories"
                
                [System.Windows.MessageBox]::Show(
                    "Failed to retrieve repositories: $_",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            }
            
            # Clean up
            Remove-Job -Job $job -Force
            $sync.FetchButton.IsEnabled = $true
        }
    })
    
    $timer.Start()
})

$sync.SearchBox.Add_TextChanged({
    Update-RepoSearch
})

$sync.SelectAllButton.Add_Click({
    if ($sync.RepoListView.Items.Count -gt 0) {
        $sync.RepoListView.SelectAll()
        Update-SelectedRepoCount
    }
})

$sync.CloneButton.Add_Click({
    if ($sync.CloningInProgress) {
        # Cancel operation
        $sync.CancellationRequested = $true
        $sync.CloneButton.IsEnabled = $false
        $sync.StatusBar.Text = "Cancelling cloning operation..."
        Write-CloneLog "Cancelling cloning operation..." -Type Warning
        return
    }
    
    $selectedRepos = $sync.RepoListView.SelectedItems
    
    if ($selectedRepos.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one repository to clone.", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Get values from form
    $clonePath = $sync.ClonePathTextBox.Text.TrimEnd('\')
    $maxJobs = [int]$sync.MaxJobsTextBox.Text
    $token = $sync.TokenPasswordBox.SecurePassword
    
    # Validate
    if ([string]::IsNullOrWhiteSpace($clonePath)) {
        [System.Windows.MessageBox]::Show("Please enter a valid clone path.", "Missing Path", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if ($maxJobs -lt 1) {
        $maxJobs = 5
    }
    
    # Confirm
    $confirmResult = [System.Windows.MessageBox]::Show(
        "Are you sure you want to clone $($selectedRepos.Count) repositories to $clonePath?",
        "Confirm Clone",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($confirmResult -ne [System.Windows.MessageBoxResult]::Yes) {
        return
    }
    
    # Start cloning
    Start-RepoCloning -Repositories $selectedRepos -ClonePath $clonePath -Token $token -MaxParallelJobs $maxJobs
})

$sync.BrowseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Clone Destination Folder"
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    
    if ($sync.ClonePathTextBox.Text -and (Test-Path $sync.ClonePathTextBox.Text)) {
        $folderBrowser.SelectedPath = $sync.ClonePathTextBox.Text
    }
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.ClonePathTextBox.Text = $folderBrowser.SelectedPath
    }
})

$sync.RepoListView.Add_SelectionChanged({
    Update-SelectedRepoCount
})

$sync.IncludePrivateCheckBox.Add_Checked({ Update-RepoSearch })
$sync.IncludePrivateCheckBox.Add_Unchecked({ Update-RepoSearch })
$sync.IncludePublicCheckBox.Add_Checked({ Update-RepoSearch })
$sync.IncludePublicCheckBox.Add_Unchecked({ Update-RepoSearch })

# Initial UI setup
$sync.CloneButton.IsEnabled = $false
$sync.ProgressBar.Value = 0
Write-CloneLog "Application started" -Type Information
$sync.StatusBar.Text = "Ready"

# Show the window
$sync.Window.ShowDialog() | Out-Null