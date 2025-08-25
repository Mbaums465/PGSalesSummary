Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- GUI Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sales Viewer"
$form.Size = New-Object System.Drawing.Size(650,550)
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
$cmbGroup.Items.AddRange(@("Buyer","Item","Year","Month","Week","Day"))
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
$txtTop.Text = "10"
$form.Controls.Add($txtTop)

# --- Run Button ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(200,300)
$btnRun.Size = New-Object System.Drawing.Size(80,30)
$form.Controls.Add($btnRun)

# --- Output Box ---
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20,350)
$txtOutput.Size = New-Object System.Drawing.Size(600,160)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = "Vertical"
$txtOutput.Font = New-Object System.Drawing.Font("Consolas",10)
$txtOutput.Anchor = "Top, Bottom, Left, Right"
$form.Controls.Add($txtOutput)

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

    # --- Get files & latest logic ---
    $files = Get-ChildItem -Path $folder -Filter "PlayerShopLog_*.txt" |
             Where-Object { $_.BaseName -match '^PlayerShopLog_(\d{6})_(\d+)$' }

    $files = $files | Where-Object {
        $fileDate = $_.BaseName -replace '^PlayerShopLog_(\d{6})_\d+$','$1'
        $yy = [int]$fileDate.Substring(0,2)
        $mm = [int]$fileDate.Substring(2,2)
        $dd = [int]$fileDate.Substring(4,2)
        $dateObj = Get-Date -Year (2000 + $yy) -Month $mm -Day $dd
        ($dateObj -ge $startDate) -and ($dateObj -le $endDate)
    }

    $latestFiles = $files |
        Group-Object { $_.BaseName -replace '^PlayerShopLog_(\d{6})_\d+$','$1' } |
        ForEach-Object {
            $_.Group | Sort-Object { [int]($_.BaseName -replace '^PlayerShopLog_\d{6}_(\d+)$','$1') } -Descending |
            Select-Object -First 1
        }

    $salesWithDate = foreach ($file in $latestFiles) {
        $fileDate = $file.BaseName -replace '^PlayerShopLog_(\d{6})_\d+$','$1'
        $yy = [int]$fileDate.Substring(0,2)
        $mm = [int]$fileDate.Substring(2,2)
        $dd = [int]$fileDate.Substring(4,2)
        $dateObj = Get-Date -Year (2000 + $yy) -Month $mm -Day $dd

        foreach ($line in Get-Content $file.FullName | Where-Object { $_ -match 'bought' }) {
            if ($line -match "- (?<buyer>\S+) bought\s+(?<item>.+?)(?:\s*x(?<qty>\d+))?\s+at a cost.*=\s*(?<earned>\d+)$") {
                [PSCustomObject]@{
                    Buyer    = $matches['buyer']
                    Item     = $matches['item'].Trim()
                    Quantity = if ($matches['qty']) { [int]$matches['qty'] } else { 1 }
                    Earned   = [int]$matches['earned']
                    SaleDate = $dateObj
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
        "Buyer" { $groupProperty = "Buyer"; $sortBy = "TotalEarned"; $sortDescending = $true }
        "Item" { $groupProperty = "Item"; $sortBy = "TotalEarned"; $sortDescending = $true }
        "Year" { $groupProperty = @{Expression={$_.SaleDate.Year}}; $sortBy = "Group"; $sortDescending = $false }
        "Month" { $groupProperty = @{Expression={$_.SaleDate.ToString("yyyy-MM")}}; $sortBy = "Group"; $sortDescending = $false }
        "Week" { $groupProperty = @{Expression={(Get-Date $_.SaleDate -UFormat "%Y-%U")}}; $sortBy = "Group"; $sortDescending = $false }
        "Day" { $groupProperty = @{Expression={$_.SaleDate.ToString("yyyy-MM-dd")}}; $sortBy = "Group"; $sortDescending = $false }
        default { $groupProperty = "Item"; $sortBy = "TotalEarned"; $sortDescending = $true }
    }

    $result = $salesWithDate | Group-Object -Property $groupProperty | ForEach-Object {
        $totalQty    = ($_.Group | Measure-Object Quantity -Sum).Sum
        $totalEarned = ($_.Group | Measure-Object Earned -Sum).Sum
        $avgPrice    = if ($totalQty -gt 0) { [math]::Round($totalEarned / $totalQty,0) } else { 0 }
        [PSCustomObject]@{
            Group       = $_.Name
            TotalSold   = "{0:N0}" -f $totalQty
            TotalEarned = "{0:N0}" -f $totalEarned
            AvgPrice    = "{0:N0}" -f $avgPrice
        }
    } | Sort-Object -Property $sortBy -Descending:$sortDescending

    if ($topObsToShow -and $topObsToShow -gt 0) {
        $result = $result | Select-Object -First $topObsToShow
    }

    $outputText = "Showing totals grouped by $groupBy`r`n"
    $outputText += ($result | Format-Table -AutoSize | Out-String)
    $txtOutput.Text = $outputText
})

# --- Show the form ---
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
