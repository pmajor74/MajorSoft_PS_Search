<#
.SYNOPSIS
MajorSoft PS Search is a powerful file search tool developed by Patrick Major. It allows users to search for files based on 
various criteria such as file path, file pattern, text content, date range, and file size. Additionally, it offers the ability 
to generate a CSV report of the search results. 

.DESCRIPTION
MajorSoft PS Search is designed to simplify the process of finding files within a specified directory and its subfolders. 
It supports advanced search options, including the ability to search for files based on their content using regular expressions. 
The tool also provides options to filter files based on their last modified date and size range. The search results are displayed
in two tabs: "Results" and "CSV Report", allowing users to easily navigate and analyze the findings.

.AUTHOR
Author: Patrick Major
#>


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# XAML for the UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MajorSoft PS Search" Height="650" Width="1000" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="Search Path:" Margin="0,0,10,5"/>
        <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtPath" Margin="0,0,0,5"/>
        <Button Grid.Row="0" Grid.Column="2" Content="Browse" x:Name="btnBrowse" Margin="5,0,0,5" Padding="5,0"/>

        <Label Grid.Row="1" Grid.Column="0" Content="File Pattern:" Margin="0,0,10,5"/>
        <TextBox Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtPattern" Margin="0,0,0,5"/>

        <Label Grid.Row="2" Grid.Column="0" Content="Contains Text:" Margin="0,0,10,5"/>
        <TextBox Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtContains" Margin="0,0,0,5"/>

        <StackPanel Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3" Orientation="Horizontal" Margin="0,0,0,5">
            <CheckBox x:Name="chkSubfolders" Content="Include Subfolders" Margin="0,0,20,0" IsChecked="True"/>
            <CheckBox x:Name="chkCountFiles" Content="Count Files (Slower)" Margin="0,0,20,0" IsChecked="True"/>
            <CheckBox x:Name="chkRegex" Content="Use Regex" Margin="0,0,20,0"/>
            <Label Content="Date Range:"/>
            <DatePicker x:Name="dpStartDate" Margin="5,0,5,0"/>
            <Label Content="to"/>
            <DatePicker x:Name="dpEndDate" Margin="5,0,20,0"/>
            <Label Content="Size Range (KB):"/>
            <TextBox x:Name="txtMinSize" Width="50" Margin="5,0,5,0"/>
            <Label Content="to"/>
            <TextBox x:Name="txtMaxSize" Width="50" Margin="5,0,0,0"/>
        </StackPanel>

        <Button Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Content="Search" x:Name="btnSearch" Margin="0,0,5,10"/>
        <Button Grid.Row="4" Grid.Column="2" Content="Cancel" x:Name="btnCancel" Margin="0,0,0,10" IsEnabled="False"/>

        <ProgressBar Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" x:Name="progressBar" Height="20" Margin="0,0,0,5"/>
        <TextBlock Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" x:Name="txtStatus" Margin="0,5,0,0"/>

        <TabControl Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,5,0,0">
            <TabItem Header="Results">
                <ListView x:Name="lstResults">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Header="File Path" Width="400" DisplayMemberBinding="{Binding FilePath}"/>
                            <GridViewColumn Header="Matches" Width="300" DisplayMemberBinding="{Binding Matches}"/>
                        </GridView>
                    </ListView.View>
                </ListView>
            </TabItem>
            <TabItem Header="CSV Report">
                <DataGrid x:Name="dgCSVReport" AutoGenerateColumns="False" IsReadOnly="True">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="File Path" Binding="{Binding FilePath}" Width="*"/>
                        <DataGridTextColumn Header="File Date" Binding="{Binding FileDate}" Width="150"/>
                        <DataGridTextColumn Header="File Size (KB)" Binding="{Binding FileSizeKB}" Width="100"/>
                        <DataGridTextColumn Header="Detections" Binding="{Binding Detections}" Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$txtPath = $window.FindName("txtPath")
$btnBrowse = $window.FindName("btnBrowse")
$txtPattern = $window.FindName("txtPattern")
$txtContains = $window.FindName("txtContains")
$chkSubfolders = $window.FindName("chkSubfolders")
$chkCountFiles = $window.FindName("chkCountFiles")
$dpStartDate = $window.FindName("dpStartDate")
$dpEndDate = $window.FindName("dpEndDate")
$txtMinSize = $window.FindName("txtMinSize")
$txtMaxSize = $window.FindName("txtMaxSize")
$btnSearch = $window.FindName("btnSearch")
$btnCancel = $window.FindName("btnCancel")
$progressBar = $window.FindName("progressBar")
$txtStatus = $window.FindName("txtStatus")
$lstResults = $window.FindName("lstResults")
$dgCSVReport = $window.FindName("dgCSVReport")
$chkRegex = $window.FindName("chkRegex")

$settingsFile = Join-Path $env:APPDATA "MajorSoftPSSearch_Settings.json"

function SaveSettings {
    $settings = @{
        Path = $txtPath.Text
        Pattern = $txtPattern.Text
        Contains = $txtContains.Text
        IncludeSubfolders = $chkSubfolders.IsChecked
        CountFiles = $chkCountFiles.IsChecked
        UseRegex = $chkRegex.IsChecked
        StartDate = $dpStartDate.SelectedDate
        EndDate = $dpEndDate.SelectedDate
        MinSize = $txtMinSize.Text
        MaxSize = $txtMaxSize.Text
    }
    $settings | ConvertTo-Json | Set-Content -Path $settingsFile
}

function LoadSettings {
    if (Test-Path $settingsFile) {
        $settings = Get-Content -Path $settingsFile | ConvertFrom-Json
        $txtPath.Text = $settings.Path
        $txtPattern.Text = $settings.Pattern
        $txtContains.Text = $settings.Contains
        $chkSubfolders.IsChecked = $settings.IncludeSubfolders
        $chkCountFiles.IsChecked = $settings.CountFiles
        $chkRegex.IsChecked = $settings.UseRegex
        $dpStartDate.SelectedDate = $settings.StartDate
        $dpEndDate.SelectedDate = $settings.EndDate
        $txtMinSize.Text = $settings.MinSize
        $txtMaxSize.Text = $settings.MaxSize
    }
}

# Browse function
function BrowseFolder {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.ValidateNames = $false
    $openFileDialog.CheckFileExists = $false
    $openFileDialog.CheckPathExists = $true
    $openFileDialog.FileName = "Folder Selection"
    $openFileDialog.Filter = "Folders|no_files"
    $openFileDialog.Title = "Select a folder"

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedPath = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
        $txtPath.Text = $selectedPath
    }
}

# Search function
$script:cancelSearch = $false

function PerformSearch {
    # Save current settings
    SaveSettings

    $script:cancelSearch = $false
    $btnSearch.IsEnabled = $false
    $btnCancel.IsEnabled = $true
    $progressBar.Value = 0
    $lstResults.Items.Clear()
    $dgCSVReport.Items.Clear()
    $txtStatus.Text = "Preparing search..."

    $path = $txtPath.Text
    $pattern = $txtPattern.Text
    if ([string]::IsNullOrWhiteSpace($pattern)) { $pattern = "*" }
    $contains = $txtContains.Text
    $useRegex = $chkRegex.IsChecked
    $includeSubfolders = $chkSubfolders.IsChecked
    $countFiles = $chkCountFiles.IsChecked
    $startDate = $dpStartDate.SelectedDate
    $endDate = $dpEndDate.SelectedDate
    $minSize = if ($txtMinSize.Text) { [int]$txtMinSize.Text * 1KB } else { $null }
    $maxSize = if ($txtMaxSize.Text) { [int]$txtMaxSize.Text * 1KB } else { $null }

    $searchParams = @{
        Path = $path
        Filter = $pattern
        Recurse = $includeSubfolders
        File = $true
    }

    if ($countFiles) {
        $txtStatus.Text = "Counting files... This may take a while."
        [System.Windows.Forms.Application]::DoEvents()
        $allFiles = @(Get-ChildItem @searchParams)
        $totalFiles = $allFiles.Count
        $txtStatus.Text = "Searching $totalFiles files..."
    } else {
        $allFiles = Get-ChildItem @searchParams
        $totalFiles = 0
        $progressBar.Visibility = "Collapsed"
        $txtStatus.Text = "Searching files..."
    }

    $processedFiles = 0

    foreach ($file in $allFiles) {
        if ($script:cancelSearch) {
            $txtStatus.Text = "Search cancelled."
            break
        }

        $include = $true

        if ($startDate -and $file.LastWriteTime -lt $startDate) { $include = $false }
        if ($endDate -and $file.LastWriteTime -gt $endDate) { $include = $false }
        if ($minSize -and $file.Length -lt $minSize) { $include = $false }
        if ($maxSize -and $file.Length -gt $maxSize) { $include = $false }

        if ($include) {
            $matchLines = @()
            $detections = @()
            if (![string]::IsNullOrWhiteSpace($contains)) {
                try {
                    $content = Get-Content $file.FullName -Raw
                    $pattern = if ($useRegex) { $contains } else { [regex]::Escape($contains) }
                    if ($content -match $pattern) {
                        $lines = $content -split "`r?`n"
                        for ($i = 0; $i -lt $lines.Count; $i++) {
                            if ($lines[$i] -match $pattern) {
                                $matchLines += "Line $($i + 1)"
                                $detections += "Line $($i + 1): $($lines[$i])"
                            }
                        }
                    }
                    else {
                        $include = $false
                    }
                }
                catch {
                    Write-Host "Error processing file $($file.FullName): $_"
                    $include = $false
                }
            }

            if ($include) {
                $lstResults.Items.Add([PSCustomObject]@{
                    FilePath = $file.FullName
                    Matches = if ($matchLines.Count -gt 0) { $matchLines -join ", " } else { "File match" }
                })

                $dgCSVReport.Items.Add([PSCustomObject]@{
                    FilePath = $file.FullName
                    FileDate = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    FileSizeKB = [math]::Round($file.Length / 1KB, 2)
                    Detections = if ($detections.Count -gt 0) { $detections -join "`n" } else { "File match" }
                })
            }
        }

        if ($countFiles) {
            $processedFiles++
            $progress = [math]::Min(100, ($processedFiles / $totalFiles) * 100)
            $progressBar.Value = $progress
            $txtStatus.Text = "Searching... ($processedFiles / $totalFiles)"
        } else {
            $processedFiles++
            if ($processedFiles % 100 -eq 0) {
                $txtStatus.Text = "Searching... (Processed $processedFiles files)"
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $txtStatus.Text = "Search completed. Found $($lstResults.Items.Count) results."
    $btnSearch.IsEnabled = $true
    $btnCancel.IsEnabled = $false
    $progressBar.Value = 100
    $progressBar.Visibility = "Visible"
}

# Cancel search function
function CancelSearch {
    $script:cancelSearch = $true
    $btnCancel.IsEnabled = $false
}

# Open file function
function OpenFile($filePath) {
    Start-Process $filePath
}

# Navigate to file location function
function NavigateToFileLocation($filePath) {
    $folderPath = [System.IO.Path]::GetDirectoryName($filePath)
    Start-Process "explorer.exe" -ArgumentList "/select,`"$filePath`""
}

# Save CSV Report function
function SaveCSVReport {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save CSV Report"
    $saveFileDialog.FileName = "SearchResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = $saveFileDialog.FileName
        $dgCSVReport.Items | Select-Object FilePath, FileDate, FileSizeKB, Detections | Export-Csv -Path $csvPath -NoTypeInformation
        [System.Windows.MessageBox]::Show("CSV report saved successfully.", "Save Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
}

# Attach functions to button click events
$btnBrowse.Add_Click({ BrowseFolder })
$btnSearch.Add_Click({ PerformSearch })
$btnCancel.Add_Click({ CancelSearch })

# Handle double-click on result item
$lstResults.Add_MouseDoubleClick({
    $selectedItem = $lstResults.SelectedItem
    if ($selectedItem) {
        OpenFile($selectedItem.FilePath)
    }
})

# Handle double-click on CSV report item
$dgCSVReport.Add_MouseDoubleClick({
    $selectedItem = $dgCSVReport.SelectedItem
    if ($selectedItem) {
        OpenFile($selectedItem.FilePath)
    }
})

# Handle right-click on result item
$contextMenu = New-Object System.Windows.Controls.ContextMenu
$menuItem = New-Object System.Windows.Controls.MenuItem
$menuItem.Header = "Navigate here"
$menuItem.Add_Click({
    $selectedItem = $lstResults.SelectedItem
    if ($selectedItem) {
        NavigateToFileLocation($selectedItem.FilePath)
    }
})
$contextMenu.Items.Add($menuItem)

$lstResults.ContextMenu = $contextMenu

# Handle right-click on CSV report item
$csvContextMenu = New-Object System.Windows.Controls.ContextMenu
$csvMenuItem = New-Object System.Windows.Controls.MenuItem
$csvMenuItem.Header = "Navigate here"
$csvMenuItem.Add_Click({
    $selectedItem = $dgCSVReport.SelectedItem
    if ($selectedItem) {
        NavigateToFileLocation($selectedItem.FilePath)
    }
})
$csvContextMenu.Items.Add($csvMenuItem)

# Add "Save CSV Report" menu item
$saveCSVMenuItem = New-Object System.Windows.Controls.MenuItem
$saveCSVMenuItem.Header = "Save CSV Report"
$saveCSVMenuItem.Add_Click({ SaveCSVReport })
$csvContextMenu.Items.Add($saveCSVMenuItem)

$dgCSVReport.ContextMenu = $csvContextMenu

# Handle Enter key press to trigger search
$window.Add_KeyDown({
    if ($_.Key -eq 'Enter' -and $btnSearch.IsEnabled) {
        $btnSearch.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
})

$window.Add_Closing({
    SaveSettings
})

LoadSettings

# Bring window to foreground
$window.Topmost = $true
$window.Activate()
$window.Topmost = $false

# Show the window
$window.ShowDialog() | Out-Null