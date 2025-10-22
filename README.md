================================================================================
                        SALES VIEWER - README
================================================================================

DESCRIPTION:
Sales Viewer is a PowerShell GUI application that reads Project Gorgon shop logs 
and displays sales data with flexible filtering, grouping, and sorting options.

================================================================================
REQUIREMENTS:
================================================================================
- Windows PowerShell 5.0 or later (or PowerShell 7+)
- Project Gorgon shop log files in the default or custom folder location
- Shop log files must be named: PlayerShopLog_YYMMDD_#.txt

================================================================================
HOW TO RUN:
================================================================================
1. Double-click RunSalesSummary.bat to launch the application

OR

2. Open PowerShell
3. Run the script:
   .\sales_viewer.ps1

The GUI window will open immediately.

================================================================================
FEATURES & CONTROLS:
================================================================================

FOLDER SELECTION:
  - Default folder: AppData\LocalLow\Elder Game\Project Gorgon\Books
  - Browse: Click to manually select a different folder

GROUP BY:
  - Select how to organize results: Buyer, Item, Year, Month, Week, or Day

BUYER FILTER:
  - Enter a buyer name to filter sales by specific players
  - Leave blank to include all buyers

ITEM FILTER:
  - Enter an item name to filter sales
  - "Exact" checkbox: 
    * CHECKED = exact item name match only
    * UNCHECKED = partial match (finds "sword", "wooden sword", etc.)
  - Leave blank to include all items

DATE FILTERS:
  - Start Date: Beginning of the date range (defaults to Jan 1 of current year)
  - End Date: End of the date range (defaults to today)
  - Format: MM/DD/YYYY

TOP N RESULTS:
  - Enter a number to limit results (e.g., 10 for top 10)
  - Leave blank or 0 to show all results

SORT BY:
  - Group: Sort alphabetically by group name (ascending)
  - TotalSold: Sort by quantity sold (descending)
  - TotalEarned: Sort by total earnings (descending, default)
  - AvgPrice: Sort by average price per item (descending)

RUN:
  - Click to execute the query and display results

OUTPUT:
  - Results appear in the text box below
  - Shows: Group, TotalSold, TotalEarned, AvgPrice

================================================================================
EXAMPLES:
================================================================================

EXAMPLE 1: Top 5 most profitable items
  - Item Filter: (blank)
  - Group By: Item
  - Sort By: TotalEarned
  - Top N Results: 5

EXAMPLE 2: Exact match for a specific item
  - Item Filter: Iron Ore
  - Exact: CHECKED
  - Group By: Buyer
  - Shows only sales of exactly "Iron Ore" by each buyer

EXAMPLE 3: Search for items containing "sword"
  - Item Filter: sword
  - Exact: UNCHECKED
  - Group By: Item
  - Shows all items with "sword" in the name

EXAMPLE 4: Daily sales report for a specific buyer
  - Buyer Filter: PlayerName
  - Group By: Day
  - Sort By: TotalEarned
  - Shows daily totals for that player

================================================================================
TROUBLESHOOTING:
================================================================================

ERROR: "Please select a valid folder"
  - Verify the folder path exists and contains PlayerShopLog_*.txt files
  - Click Browse to manually select the folder

ERROR: "No sales found for the applied filters"
  - Adjust your filters (dates, buyer name, item name)
  - Ensure files exist in the selected folder for the date range

SCRIPT WON'T RUN:
  - Check PowerShell execution policy
  - Ensure .ps1 file is saved in correct location

================================================================================
NOTES:
================================================================================
- Only the latest version of each day's log file is used
- Item matching is case-insensitive
- Quantities and prices are formatted with commas for readability
- All calculations use integer math (prices are rounded)

================================================================================
