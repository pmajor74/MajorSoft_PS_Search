Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# XAML for the UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MajorSoft PS Search" Height="650" Width="800" WindowStartupLocation="CenterScreen">
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
        <TextBlock Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" x:Name="txtStatus" Margin="0,0,0,5"/>

        <ListView Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="3" x:Name="lstResults" Margin="0,5,0,0">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="File Path" Width="400" DisplayMemberBinding="{Binding FilePath}"/>
                    <GridViewColumn Header="Matches" Width="300" DisplayMemberBinding="{Binding Matches}"/>
                </GridView>
            </ListView.View>
        </ListView>
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
$dpStartDate = $window.FindName("dpStartDate")
$dpEndDate = $window.FindName("dpEndDate")
$txtMinSize = $window.FindName("txtMinSize")
$txtMaxSize = $window.FindName("txtMaxSize")
$btnSearch = $window.FindName("btnSearch")
$btnCancel = $window.FindName("btnCancel")
$progressBar = $window.FindName("progressBar")
$txtStatus = $window.FindName("txtStatus")
$lstResults = $window.FindName("lstResults")

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
    $script:cancelSearch = $false
    $btnSearch.IsEnabled = $false
    $btnCancel.IsEnabled = $true
    $progressBar.Value = 0
    $lstResults.Items.Clear()
    $txtStatus.Text = "Preparing search..."

    $path = $txtPath.Text
    $pattern = $txtPattern.Text
    if ([string]::IsNullOrWhiteSpace($pattern)) { $pattern = "*" }
    $contains = $txtContains.Text
    $includeSubfolders = $chkSubfolders.IsChecked
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

    $allFiles = @(Get-ChildItem @searchParams)
    $totalFiles = $allFiles.Count
    $processedFiles = 0

    $txtStatus.Text = "Searching $totalFiles files..."

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

        if ($include -and $contains) {
            $matches = @()
            $lineNumber = 0
            foreach ($line in Get-Content $file.FullName) {
                $lineNumber++
                if ($line -match [regex]::Escape($contains)) {
                    $matches += "Line $lineNumber"
                }
            }
            if ($matches.Count -gt 0) {
                $lstResults.Items.Add([PSCustomObject]@{
                    FilePath = $file.FullName
                    Matches = $matches -join ", "
                })
            }
        }

        $processedFiles++
        $progress = [math]::Min(100, ($processedFiles / $totalFiles) * 100)
        $progressBar.Value = $progress
        $txtStatus.Text = "Searching... ($processedFiles / $totalFiles)"
        [System.Windows.Forms.Application]::DoEvents()
    }

    $txtStatus.Text = "Search completed. Found $($lstResults.Items.Count) results."
    $btnSearch.IsEnabled = $true
    $btnCancel.IsEnabled = $false
    $progressBar.Value = 100
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

# Handle Enter key press
$window.Add_KeyDown({
    if ($_.Key -eq 'Enter' -and $btnSearch.IsEnabled) {
        $btnSearch.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
})

# Bring window to foreground
$window.Topmost = $true
$window.Activate()
$window.Topmost = $false

# Show the window
$window.ShowDialog() | Out-Null