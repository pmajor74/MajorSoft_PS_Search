Add-Type -AssemblyName PresentationFramework

# XAML for the UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MajorSoft PS Search" Height="600" Width="800">
    <Grid Margin="10">
        <Grid.RowDefinitions>
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
        </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="Search Path:" Margin="0,0,10,5"/>
        <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtPath" Margin="0,0,0,5"/>

        <Label Grid.Row="1" Grid.Column="0" Content="File Pattern:" Margin="0,0,10,5"/>
        <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtPattern" Margin="0,0,0,5"/>

        <Label Grid.Row="2" Grid.Column="0" Content="Contains Text:" Margin="0,0,10,5"/>
        <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtContains" Margin="0,0,0,5"/>

        <StackPanel Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,0,0,5">
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

        <Button Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Content="Search" x:Name="btnSearch" Margin="0,0,0,10"/>

        <TextBox Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="2" x:Name="txtResults" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$txtPath = $window.FindName("txtPath")
$txtPattern = $window.FindName("txtPattern")
$txtContains = $window.FindName("txtContains")
$chkSubfolders = $window.FindName("chkSubfolders")
$dpStartDate = $window.FindName("dpStartDate")
$dpEndDate = $window.FindName("dpEndDate")
$txtMinSize = $window.FindName("txtMinSize")
$txtMaxSize = $window.FindName("txtMaxSize")
$btnSearch = $window.FindName("btnSearch")
$txtResults = $window.FindName("txtResults")

# Search function
function PerformSearch {
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

    Get-ChildItem @searchParams | ForEach-Object {
        $include = $true

        if ($startDate -and $_.LastWriteTime -lt $startDate) { $include = $false }
        if ($endDate -and $_.LastWriteTime -gt $endDate) { $include = $false }
        if ($minSize -and $_.Length -lt $minSize) { $include = $false }
        if ($maxSize -and $_.Length -gt $maxSize) { $include = $false }

        if ($include -and $contains) {
            $content = Get-Content $_.FullName -Raw
            if ($content -notmatch [regex]::Escape($contains)) { $include = $false }
        }

        if ($include) {
            $results += "$($_.FullName) ($('{0:N2}' -f ($_.Length / 1KB)) KB, $($_.LastWriteTime))"
        }
    }

    $txtResults.Text = $results -join "`r`n"
}

# Attach search function to button click event
$btnSearch.Add_Click({ PerformSearch })

# Show the window
$window.ShowDialog() | Out-Null