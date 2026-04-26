# 迭代日历抽取器
# 从 Excel 第一个工作表第 9 行开始抽取迭代数据

param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ""
)

# 校验输入文件
if (-not (Test-Path $ExcelFilePath)) {
    Write-Error "未找到 Excel 文件：$ExcelFilePath"
    exit 1
}

# 未指定输出路径时，默认输出到 Excel 同目录的同名 Markdown 文件
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFilePath)
    $OutputPath = Join-Path (Split-Path $ExcelFilePath -Parent) "$fileName.md"
}

Write-Host "正在处理 Excel 文件：$ExcelFilePath"
Write-Host "输出文件将保存到：$OutputPath"

# 创建 Excel 应用对象
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
}
catch {
    Write-Error "无法创建 Excel 应用对象。请确认本机已安装 Microsoft Excel。"
    exit 1
}

# 打开工作簿
try {
    $workbook = $excel.Workbooks.Open($ExcelFilePath)
}
catch {
    Write-Error "无法打开 Excel 文件：$_"
    $excel.Quit()
    exit 1
}

# 读取第一个工作表
$sheet = $workbook.Sheets(1)
Write-Host ""
Write-Host "正在处理第一个工作表：$($sheet.Name)"

$results = @()
$row = 9

# 逐行处理数据，直到迭代编号单元格为空
while ($true) {
    $iteration = $sheet.Cells($row, 1).Text
    $startDate = $sheet.Cells($row, 2).Text
    $endDate = $sheet.Cells($row, 3).Text
    
    # A 列为空时停止
    if ([string]::IsNullOrWhiteSpace($iteration)) { break }
    
    # 清理单元格文本
    $iteration = $iteration.Trim()
    $startDate = $startDate.Trim()
    $endDate = $endDate.Trim()
    
    # 格式化输出条目
    $entry = "迭代$iteration，$startDate 是迭代启动日期，$endDate 是迭代发布日期"
    
    Write-Host "  第 $row 行：$entry"
    
    $results += [PSCustomObject]@{
        Row = $row
        Iteration = $iteration
        StartDate = $startDate
        EndDate = $endDate
        Entry = $entry
    }
    
    $row++
}

# 关闭 Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "已抽取 $($results.Count) 个迭代周期"

# 生成 Markdown 内容
$mdContent = "# 迭代日历`n`n"
$mdContent += "**来源：** $(Split-Path $ExcelFilePath -Leaf)`n`n"
$mdContent += "**生成时间：** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n`n"
$mdContent += "---`n`n"
$mdContent += "## 迭代周期列表`n`n"

foreach ($item in $results) {
    $mdContent += "- $($item.Entry)`n"
}

# 保存 Markdown 文件
try {
    $utf8Encoding = New-Object System.Text.UTF8Encoding $true
    $streamWriter = New-Object System.IO.StreamWriter($OutputPath, $false, $utf8Encoding)
    $streamWriter.Write($mdContent)
    $streamWriter.Close()
    Write-Host ""
    Write-Host "Markdown 文档已保存：$OutputPath"
}
catch {
    Write-Error "无法保存 Markdown 文件：$_"
    exit 1
}

Write-Host "处理完成。"
