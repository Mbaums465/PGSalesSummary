# Project Gorgon Sales Viewer
#Sorry, this is no longer supported 

A PowerShell GUI application for analyzing Project Gorgon shop sales data with flexible filtering, grouping, and sorting capabilities.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Filtering Options](#filtering-options)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Overview

Sales Viewer parses Project Gorgon shop log files and provides an intuitive GUI for analyzing your sales history. Track your best-selling items, most frequent buyers, revenue trends, and more with customizable filters and grouping options.

## Features

- 📊 **Flexible Grouping** - Organize sales by Buyer, Item, Year, Month, Week, or Day
- 🔍 **Advanced Filtering** - Filter by buyer name, item name (exact or partial match), and date range
- 📈 **Multiple Sort Options** - Sort by group name, quantity sold, total earnings, or average price
- 💰 **Revenue Analysis** - View total earnings, quantities sold, and average prices
- 🎯 **Top N Results** - Limit results to see your top performers
- 🖥️ **User-Friendly GUI** - Easy-to-use Windows Forms interface

## Requirements

- Windows PowerShell 5.0 or later (or PowerShell 7+)
- Project Gorgon installed with shop logs enabled
- Shop log files in format: `PlayerShopLog_YYMMDD_#.txt`

## Installation

1. Clone or download this repository:
   ```bash
   git clone https://github.com/Mbaums465/PGSalesSummary.git
   ```

2. Navigate to the project directory:
   ```bash
   cd PGSalesSummary
   ```

3. No additional dependencies required - uses built-in PowerShell modules

## Usage

### Quick Start

**Option 1:** Double-click `RunSalesSummary.bat`

**Option 2:** Run via PowerShell:
```powershell
.\sales_viewer.ps1
```

### Default Log Location

The application automatically looks for shop logs in:
```
%AppData%\LocalLow\Elder Game\Project Gorgon\Books
```

Use the **Browse** button to select a different folder if needed.

## Filtering Options

### Group By
Organize your results by:
- **Buyer** - See sales per customer
- **Item** - Track item performance
- **Year/Month/Week/Day** - Analyze time-based trends

### Filters

| Filter | Description |
|--------|-------------|
| **Buyer Filter** | Enter a player name to view their purchases only |
| **Item Filter** | Search for specific items (partial or exact match) |
| **Exact Match** | ☑ Checked = exact name only<br>☐ Unchecked = partial match |
| **Date Range** | Start and end dates (MM/DD/YYYY format) |
| **Top N Results** | Limit to top X results (leave blank for all) |

### Sort Options

- **Group** - Alphabetical by group name (A-Z)
- **TotalSold** - By quantity sold (high to low)
- **TotalEarned** - By total revenue (high to low) ⭐ *Default*
- **AvgPrice** - By average price per unit (high to low)

## Examples

### Example 1: Top 5 Most Profitable Items
```
Group By: Item
Sort By: TotalEarned
Top N Results: 5
Item Filter: (blank)
```
*Shows your 5 highest-grossing items*

### Example 2: Specific Item Sales by Buyer
```
Group By: Buyer
Item Filter: Iron Ore
Exact: ☑ CHECKED
Sort By: TotalSold
```
*Shows who bought "Iron Ore" (exact match) and how much*

### Example 3: Find All Sword-Related Items
```
Group By: Item
Item Filter: sword
Exact: ☐ UNCHECKED
Sort By: TotalEarned
```
*Returns all items containing "sword" in the name*

### Example 4: Daily Sales Report for a Customer
```
Group By: Day
Buyer Filter: PlayerName
Sort By: TotalEarned
Date Range: 01/01/2025 - 03/31/2025
```
*Shows daily purchase totals for a specific player*

### Example 5: This Month's Top 10 Buyers
```
Group By: Buyer
Date Range: [First of month] - [Today]
Sort By: TotalEarned
Top N Results: 10
```
*See your best customers this month*

## Troubleshooting

### "Please select a valid folder"
- Verify the folder exists and contains `PlayerShopLog_*.txt` files
- Use the **Browse** button to manually select the correct folder
- Check that Project Gorgon has created shop log files

### "No sales found for the applied filters"
- Adjust your date range - logs may not exist for all dates
- Check filter spelling (buyer and item names are case-insensitive)
- Verify the selected folder contains log files

### Script Won't Run
- Check PowerShell execution policy:
  ```powershell
  Get-ExecutionPolicy
  ```
- If restricted, you may need to run:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
- Ensure the `.ps1` file hasn't been corrupted during download

### Missing or Incomplete Data
- Only the **latest version** of each day's log file is processed
- Ensure shop logs are enabled in Project Gorgon settings
- Check that log files aren't corrupted or empty

## Technical Notes

- Item name matching is case-insensitive
- Quantities and prices are formatted with commas for readability
- All price calculations use integer math (rounded values)
- Duplicate log files (same day) - only the highest numbered file is used

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## License

What? I have no License.

---

**Project Gorgon** is a trademark of Elder Game, LLC. This tool is a fan-made project and is not officially associated with or endorsed by Elder Game.
