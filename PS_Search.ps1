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
            <CheckBox x:Name="chkSubfolders" Content="Include Subfolders" Margin="0,0,20,0"/>
            <Label Content="Date Range:"/>
            <DatePicker x:Name="dpStartDate" Margin="5,0,5,0"/>
            <Label Content="to"/>
            <DatePicker x:Name="dpEndDate" Margin="5,0,20,0"/>
            <Label Content="Size Range (KB):"/>
            <TextBox x:Name="txtMinSize" Width="50" Margin="5,0,5,0"/>
            <Label Content="to"/>
            <TextBox x:Name="txtMaxSize" Width="50" Margin="5,0,0,0"/>
        </StackPanel>

        <Button Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3" Content="Search" x:Name="btnSearch" Margin="0,0,0,10"/>

        <ProgressBar Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" x:Name="progressBar" Height="20" Margin="0,0,0,5"/>
        <TextBlock Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="2" x:Name="txtStatus" Margin="0,0,0,5"/>
        <Button Grid.Row="6" Grid.Column="2" Content="Cancel" x:Name="btnCancel" Margin="0,0,0,5" Visibility="Collapsed"/>

        <TextBox Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="3" x:Name="txtResults" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
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
$txtResults = $window.FindName("txtResults")
$progressBar = $window.FindName("progressBar")
$txtStatus = $window.FindName("txtStatus")
$btnCancel = $window.FindName("btnCancel")

# Browse function
function BrowseFolder {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder to search"
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = $folderBrowser.SelectedPath
    }
}

# Search function
$script:cancelSearch = $false

function PerformSearch {
    $script:cancelSearch = $false
    $btnSearch.IsEnabled = $false
    $btnCancel.Visibility = "Visible"
    $progressBar.Value = 0
    $txtResults.Text = ""
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

    $results = @()

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
            $content = Get-Content $file.FullName -Raw
            if ($content -notmatch [regex]::Escape($contains)) { $include = $false }
        }

        if ($include) {
            $results += "$($file.FullName) ($('{0:N2}' -f ($file.Length / 1KB)) KB, $($file.LastWriteTime))"
        }

        $processedFiles++
        $progressBar.Value = ($processedFiles / $totalFiles) * 100
        $txtStatus.Text = "Searching... ($processedFiles / $totalFiles)"
    }

    $txtResults.Text = $results -join "`r`n"
    $txtStatus.Text = "Search completed. Found $($results.Count) results."
    $btnSearch.IsEnabled = $true
    $btnCancel.Visibility = "Collapsed"
}

# Cancel search function
function CancelSearch {
    $script:cancelSearch = $true
}

# Attach functions to button click events
$btnBrowse.Add_Click({ BrowseFolder })
$btnSearch.Add_Click({ PerformSearch })
$btnCancel.Add_Click({ CancelSearch })

# Bring window to foreground
$window.Topmost = $true
$window.Activate()
$window.Topmost = $false

# Show the window
$window.ShowDialog() | Out-Null