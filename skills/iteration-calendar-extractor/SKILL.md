---
name: "iteration-calendar-extractor"
description: "Extracts iteration calendar data from Excel first sheet starting from row 9, formats as '迭代I42, 2025/12/31 是迭代启动日期, 2026/1/14 是迭代发布日期'. Invoke when user needs to extract structured iteration data from Excel calendars."
---

# Iteration Calendar Extractor

This skill extracts iteration calendar data from the first sheet of an Excel file, starting from row 9, and formats it into readable Markdown entries.

## When to Use

- Extracting iteration calendar data from Excel files
- Converting Excel-based sprint/iteration schedules to Markdown
- Processing project timelines with iteration numbers and dates
- Generating formatted iteration summaries

## How to Use

1. Run the PowerShell script with the Excel file path
2. The script reads the **first sheet only**
3. Starts processing from **row 9** (skipping headers)
4. Extracts iteration number, start date, and end date
5. Formats output as: `迭代I42，2025/12/31 是迭代启动日期，2026/1/14 是迭代发布日期`
6. Generates a Markdown document

## Expected Excel Structure

The Excel first sheet should have data starting from row 9 with columns:
- Column A: Iteration number (e.g., I42, I43)
- Column B: Start date (e.g., 2025/12/31)
- Column C: End/Release date (e.g., 2026/1/14)

## Output Format

```markdown
# 迭代日历

**Source:** 2026年迭代日历V2.xlsx

**Generated:** 2026-04-21 13:14:00

---

## 迭代周期列表

- 迭代I42，2025/12/31 是迭代启动日期，2026/1/14 是迭代发布日期
- 迭代I43，2026/1/15 是迭代启动日期，2026/1/28 是迭代发布日期
- 迭代I44，2026/1/30 是迭代启动日期，2026/2/26 是迭代发布日期
...
```

## Implementation

```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ""
)

# Validate input file
if (-not (Test-Path $ExcelFilePath)) {
    Write-Error "Excel file not found: $ExcelFilePath"
    exit 1
}

# Set default output path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFilePath)
    $OutputPath = Join-Path (Split-Path $ExcelFilePath -Parent) "$fileName.md"
}

# Create Excel application
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open workbook
$workbook = $excel.Workbooks.Open($ExcelFilePath)

# Get first sheet
$sheet = $workbook.Sheets(1)
Write-Host "Processing first sheet: $($sheet.Name)" -ForegroundColor Green

$results = @()
$row = 9  # Start from row 9

# Process data until empty row
while ($true) {
    $iteration = $sheet.Cells($row, 1).Text  # Column A: Iteration
    $startDate = $sheet.Cells($row, 2).Text   # Column B: Start date
    $endDate = $sheet.Cells($row, 3).Text     # Column C: End date
    
    # Stop if iteration column is empty
    if ([string]::IsNullOrWhiteSpace($iteration)) { break }
    
    # Clean up values
    $iteration = $iteration.Trim()
    $startDate = $startDate.Trim()
    $endDate = $endDate.Trim()
    
    # Format the entry
    $entry = "迭代$iteration，$startDate 是迭代启动日期，$endDate 是迭代发布日期"
    
    Write-Host "  [Row $row] $entry" -ForegroundColor Yellow
    
    $results += [PSCustomObject]@{
        Row = $row
        Entry = $entry
    }
    
    $row++
}

# Close Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "`nExtracted $($results.Count) iterations" -ForegroundColor Cyan

# Generate Markdown
$mdContent = "# 迭代日历`n`n"
$mdContent += "**Source:** $(Split-Path $ExcelFilePath -Leaf)`n`n"
$mdContent += "**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n`n"
$mdContent += "---`n`n"
$mdContent += "## 迭代周期列表`n`n"

foreach ($item in $results) {
    $mdContent += "- $($item.Entry)`n"
}

# Save with UTF-8 encoding
$utf8Encoding = New-Object System.Text.UTF8Encoding $true
$streamWriter = New-Object System.IO.StreamWriter($OutputPath, $false, $utf8Encoding)
$streamWriter.Write($mdContent)
$streamWriter.Close()

Write-Host "`nMarkdown document saved: $OutputPath" -ForegroundColor Green
```

## Command Line Usage

```powershell
powershell -ExecutionPolicy Bypass -File Extract-IterationCalendar.ps1 -ExcelFilePath "c:\code\2026年迭代日历V2.xlsx"
```

## Notes

- Requires Microsoft Excel to be installed
- Only processes the first sheet
- Starts from row 9 (change `$row = 9` if needed)
- Expects iteration data in columns A, B, C
- Stops when column A is empty
- Outputs UTF-8 encoded Markdown file
