Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- GUI Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sales Viewer"
$form.Size = New-Object System.Drawing.Size(650,600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"

# --- Folder Selection ---
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = "Folder:"
$lblFolder.Location = New-Object System.Drawing.Point(20,20)
$lblFolder.AutoSize = $true
$form.Controls.Add($lblFolder)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(80,18)
$txtFolder.Size = New-Object System.Drawing.Size(400,20)
# Set default folder path
$txtFolder.Text = Join-Path $env:USERPROFILE "AppData\LocalLow\Elder Game\Project Gorgon\Books"
$form.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(500,15)
$btnBrowse.Size = New-Object System.Drawing.Size(80,25)
$form.Controls.Add($btnBrowse)

$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $txtFolder.Text = $folderBrowser.SelectedPath
    }
})

# --- GroupBy Dropdown ---
$lblGroup = New-Object System.Windows.Forms.Label
$lblGroup.Text = "Group By:"
$lblGroup.Location = New-Object System.Drawing.Point(20,60)
$lblGroup.AutoSize = $true
$form.Controls.Add($lblGroup)

$cmbGroup = New-Object System.Windows.Forms.ComboBox
$cmbGroup.Location = New-Object System.Drawing.Point(100,58)
$cmbGroup.Size = New-Object System.Drawing.Size(150,20)
# Updated the items to include month and day names
$cmbGroup.Items.AddRange(@("Buyer", "Item", "Year", "Month", "Month-Name", "Week", "Day", "Day-Name"))
$cmbGroup.SelectedIndex = 1
$form.Controls.Add($cmbGroup)

# --- Buyer Filter ---
$lblBuyer = New-Object System.Windows.Forms.Label
$lblBuyer.Text = "Buyer Filter:"
$lblBuyer.Location = New-Object System.Drawing.Point(20,100)
$lblBuyer.AutoSize = $true
$form.Controls.Add($lblBuyer)

$txtBuyer = New-Object System.Windows.Forms.TextBox
$txtBuyer.Location = New-Object System.Drawing.Point(100,98)
$txtBuyer.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($txtBuyer)

# --- Item Filter ---
$lblItem = New-Object System.Windows.Forms.Label
$lblItem.Text = "Item Filter:"
$lblItem.Location = New-Object System.Drawing.Point(20,140)
$lblItem.AutoSize = $true
$form.Controls.Add($lblItem)

$txtItem = New-Object System.Windows.Forms.TextBox
$txtItem.Location = New-Object System.Drawing.Point(100,138)
$txtItem.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($txtItem)

# --- Date Filters ---
$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = "Start Date (MM/DD/YYYY):"
$lblStart.Location = New-Object System.Drawing.Point(20,180)
$lblStart.AutoSize = $true
$form.Controls.Add($lblStart)

$txtStart = New-Object System.Windows.Forms.TextBox
$txtStart.Location = New-Object System.Drawing.Point(180,178)
$txtStart.Size = New-Object System.Drawing.Size(100,20)
$txtStart.Text = (Get-Date -Year (Get-Date).Year -Month 1 -Day 1).ToString("MM/dd/yyyy")
$form.Controls.Add($txtStart)

$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = "End Date (MM/DD/YYYY):"
$lblEnd.Location = New-Object System.Drawing.Point(20,220)
$lblEnd.AutoSize = $true
$form.Controls.Add($lblEnd)

$txtEnd = New-Object System.Windows.Forms.TextBox
$txtEnd.Location = New-Object System.Drawing.Point(180,218)
$txtEnd.Size = New-Object System.Drawing.Size(100,20)
$txtEnd.Text = (Get-Date).ToString("MM/dd/yyyy")
$form.Controls.Add($txtEnd)

# --- Top N ---
$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = "Top N Results:"
$lblTop.Location = New-Object System.Drawing.Point(20,260)
$lblTop.AutoSize = $true
$form.Controls.Add($lblTop)

$txtTop = New-Object System.Windows.Forms.TextBox
$txtTop.Location = New-Object System.Drawing.Point(120,258)
$txtTop.Size = New-Object System.Drawing.Size(50,20)
$txtTop.Text = "20"
$form.Controls.Add($txtTop)

# --- Sort By Dropdown ---
$lblSort = New-Object System.Windows.Forms.Label
$lblSort.Text = "Sort By:"
$lblSort.Location = New-Object System.Drawing.Point(20,300)
$lblSort.AutoSize = $true
$form.Controls.Add($lblSort)

$cmbSort = New-Object System.Windows.Forms.ComboBox
$cmbSort.Location = New-Object System.Drawing.Point(100,298)
$cmbSort.Size = New-Object System.Drawing.Size(150,20)
$cmbSort.Items.AddRange(@("Group","TotalSold","TotalEarned","AvgPrice"))
$cmbSort.SelectedIndex = 2   # Default = TotalEarned
$form.Controls.Add($cmbSort)

# --- Run Button ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(200,340)
$btnRun.Size = New-Object System.Drawing.Size(80,30)
$form.Controls.Add($btnRun)

# --- Output Box ---
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20,390)
$txtOutput.Size = New-Object System.Drawing.Size(600,160)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = "Vertical"
$txtOutput.Font = New-Object System.Drawing.Font("Consolas",10)
$txtOutput.Anchor = "Top, Bottom, Left, Right"
$form.Controls.Add($txtOutput)

# --- Helper function to parse sales date from log content ---
function Get-SalesDateFromLog {
    param(
        [string]$FilePath,
        [datetime]$ExtractDate
    )
    
    $firstLine = Get-Content $FilePath -First 1
    if ($firstLine -match "^(\w{3}) (\w{3}) (\d{1,2}) \d{2}:\d{2}") {
        $dayOfWeek = $matches[1]
        $month = $matches[2]
        $day = [int]$matches[3]
        
        # Convert month abbreviation to number
        $monthNumber = switch ($month) {
            "Jan" { 1 }  "Feb" { 2 }  "Mar" { 3 }  "Apr" { 4 }
            "May" { 5 }  "Jun" { 6 }  "Jul" { 7 }  "Aug" { 8 }
            "Sep" { 9 }  "Oct" { 10 } "Nov" { 11 } "Dec" { 12 }
            default { 1 }
        }
        
        # Try the extract year first
        $year = $ExtractDate.Year
        $candidateDate = Get-Date -Year $year -Month $monthNumber -Day $day
        
        # If the candidate date is more than 10 days after extract date, try previous year
        if ($candidateDate -gt $ExtractDate.AddDays(10)) {
            $year = $ExtractDate.Year - 1
            $candidateDate = Get-Date -Year $year -Month $monthNumber -Day $day
        }
        
        # If still more than 10 days in the future, try next year
        if ($candidateDate -lt $ExtractDate.AddDays(-10)) {
            $year = $ExtractDate.Year + 1
            $candidateDate = Get-Date -Year $year -Month $monthNumber -Day $day
        }
        
        return $candidateDate
    }
    
    # Fallback to extract date if we can't parse the log
    return $ExtractDate
}

# --- Button click action ---
$btnRun.Add_Click({
    $folder = $txtFolder.Text
    if (-not (Test-Path $folder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid folder.","Error","OK","Error")
        return
    }

    $groupBy = $cmbGroup.SelectedItem
    $buyerFilter = $txtBuyer.Text
    $itemFilter  = $txtItem.Text
    $topObsToShow = if ([int]::TryParse($txtTop.Text,[ref]$null)) { [int]$txtTop.Text } else { 0 }
    $startDate = [datetime]::Parse($txtStart.Text)
    $endDate   = [datetime]::Parse($txtEnd.Text)
    $sortBy = $cmbSort.SelectedItem

    # --- Get files and parse extract dates ---
    $files = Get-ChildItem -Path $folder -Filter "PlayerShopLog_*.txt" |
             Where-Object { $_.BaseName -match '^PlayerShopLog_(\d{6})_(\d+)$' }

    # Create file objects with extract dates and sales dates
    $fileInfo = foreach ($file in $files) {
        if ($file.BaseName -match '^PlayerShopLog_(\d{6})_(\d+)$') {
            $extractDateStr = $matches[1]
            $sequence = [int]$matches[2]
            
            # Parse extract date from filename
            $yy = [int]$extractDateStr.Substring(0,2)
            $mm = [int]$extractDateStr.Substring(2,2)
            $dd = [int]$extractDateStr.Substring(4,2)
            $extractDate = Get-Date -Year (2000 + $yy) -Month $mm -Day $dd
            
            # Get sales date from log content
            $salesDate = Get-SalesDateFromLog -FilePath $file.FullName -ExtractDate $extractDate
            
            [PSCustomObject]@{
                File = $file
                ExtractDate = $extractDate
                SalesDate = $salesDate
                Sequence = $sequence
            }
        }
    }

    # Filter by date range based on sales date
    $fileInfo = $fileInfo | Where-Object {
        ($_.SalesDate -ge $startDate) -and ($_.SalesDate -le $endDate)
    }

    # Handle duplicates: Group by sales date and take the latest extract date
    $latestFiles = $fileInfo |
        Group-Object { $_.SalesDate.ToString("yyyy-MM-dd") } |
        ForEach-Object {
            # For same sales date, take the latest extract date (then highest sequence if same extract date)
            $_.Group | Sort-Object ExtractDate, Sequence -Descending | Select-Object -First 1
        }

    # Process sales data
    $salesWithDate = foreach ($fileData in $latestFiles) {
        foreach ($line in Get-Content $fileData.File.FullName | Where-Object { $_ -match 'bought' }) {
            if ($line -match "- (?<buyer>\S+) bought\s+(?<item>.+?)(?:\s*x(?<qty>\d+))?\s+at a cost.*=\s*(?<earned>\d+)$") {
                [PSCustomObject]@{
                    Buyer    = $matches['buyer']
                    Item     = $matches['item'].Trim()
                    Quantity = if ($matches['qty']) { [int]$matches['qty'] } else { 1 }
                    Earned   = [int]$matches['earned']
                    SaleDate = $fileData.SalesDate
                }
            }
        }
    }

    if ($buyerFilter) { $salesWithDate = $salesWithDate | Where-Object { $_.Buyer -eq $buyerFilter } }
    if ($itemFilter)  { $salesWithDate = $salesWithDate | Where-Object { $_.Item -eq $itemFilter } }

    if (-not $salesWithDate) {
        $txtOutput.Text = "No sales found for the applied filters."
        return
    }

    switch ($groupBy) {
        "Buyer" { $groupProperty = "Buyer" }
        "Item"  { $groupProperty = "Item" }
        "Year"  { $groupProperty = @{Expression={$_.SaleDate.Year}} }
        "Month" { $groupProperty = @{Expression={$_.SaleDate.ToString("yyyy-MM")}} }
        "Month-Name" { $groupProperty = @{Expression={$_.SaleDate.ToString("MMMM")}} }
        "Week"  { $groupProperty = @{Expression={(Get-Date $_.SaleDate -UFormat "%Y-%U")}} }
        "Day"   { $groupProperty = @{Expression={$_.SaleDate.ToString("yyyy-MM-dd")}} }
        "Day-Name" { $groupProperty = @{Expression={$_.SaleDate.ToString("dddd")}} }
        default { $groupProperty = "Item" }
    }

    $result = $salesWithDate | Group-Object -Property $groupProperty | ForEach-Object {
        $totalQty    = ($_.Group | Measure-Object Quantity -Sum).Sum
        $totalEarned = ($_.Group | Measure-Object Earned -Sum).Sum
        $avgPrice    = if ($totalQty -gt 0) { [math]::Round($totalEarned / $totalQty,0) } else { 0 }
        [PSCustomObject]@{
            Group       = $_.Name
            TotalSold   = $totalQty
            TotalEarned = $totalEarned
            AvgPrice    = $avgPrice
        }
    }

    # Sorting: Group ascending, others descending
    $sortDescending = $true
    if ($sortBy -eq "Group") { $sortDescending = $false }
    $result = $result | Sort-Object -Property $sortBy -Descending:$sortDescending

    if ($topObsToShow -and $topObsToShow -gt 0) {
        $result = $result | Select-Object -First $topObsToShow
    }

    $outputText = "Showing totals grouped by $groupBy (sorted by $sortBy)`r`n"
    $outputText += ($result | Format-Table @{Label="Group";Expression={$_.Group}},
                                         @{Label="TotalSold";Expression={"{0:N0}" -f $_.TotalSold}},
                                         @{Label="TotalEarned";Expression={"{0:N0}" -f $_.TotalEarned}},
                                         @{Label="AvgPrice";Expression={"{0:N0}" -f $_.AvgPrice}} -AutoSize | Out-String)
    $txtOutput.Text = $outputText
})

# --- Show the form ---
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
