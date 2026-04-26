# Iteration Calendar Extractor
# Extracts iteration data from Excel first sheet starting from row 9

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

# Set default output path if not provided
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFilePath)
    $OutputPath = Join-Path (Split-Path $ExcelFilePath -Parent) "$fileName.md"
}

Write-Host "Processing Excel file: $ExcelFilePath"
Write-Host "Output will be saved to: $OutputPath"

# Create Excel application object
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
}
catch {
    Write-Error "Failed to create Excel application. Make sure Microsoft Excel is installed."
    exit 1
}

# Open workbook
try {
    $workbook = $excel.Workbooks.Open($ExcelFilePath)
}
catch {
    Write-Error "Failed to open Excel file: $_"
    $excel.Quit()
    exit 1
}

# Get the first sheet
$sheet = $workbook.Sheets(1)
Write-Host ""
Write-Host "Processing first sheet: $($sheet.Name)"

$results = @()
$row = 9

# Process data row by row until empty iteration cell
while ($true) {
    $iteration = $sheet.Cells($row, 1).Text
    $startDate = $sheet.Cells($row, 2).Text
    $endDate = $sheet.Cells($row, 3).Text
    
    # Stop if iteration column is empty
    if ([string]::IsNullOrWhiteSpace($iteration)) { break }
    
    # Clean up values
    $iteration = $iteration.Trim()
    $startDate = $startDate.Trim()
    $endDate = $endDate.Trim()
    
    # Format the entry
    $entry = "Iteration $iteration, $startDate is iteration start date, $endDate is iteration release date"
    
    Write-Host "  Row $row : $entry"
    
    $results += [PSCustomObject]@{
        Row = $row
        Iteration = $iteration
        StartDate = $startDate
        EndDate = $endDate
        Entry = $entry
    }
    
    $row++
}

# Close Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "Extracted $($results.Count) iterations"

# Generate Markdown content
$mdContent = "# Iteration Calendar`n`n"
$mdContent += "**Source:** $(Split-Path $ExcelFilePath -Leaf)`n`n"
$mdContent += "**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n`n"
$mdContent += "---`n`n"
$mdContent += "## Iteration List`n`n"

foreach ($item in $results) {
    $mdContent += "- $($item.Entry)`n"
}

# Save Markdown file
try {
    $utf8Encoding = New-Object System.Text.UTF8Encoding $true
    $streamWriter = New-Object System.IO.StreamWriter($OutputPath, $false, $utf8Encoding)
    $streamWriter.Write($mdContent)
    $streamWriter.Close()
    Write-Host ""
    Write-Host "Markdown document saved: $OutputPath"
}
catch {
    Write-Error "Failed to save Markdown file: $_"
    exit 1
}

Write-Host "Done!"
