param(
  [string]$DataPath = "E:\Codex Project\英语学习\daily-english\data\vocabulary.json",
  [string]$OutputPath = "E:\Codex Project\英语学习\daily-english\data\advanced_english_vocabulary.xlsx"
)

$rows = Get-Content -LiteralPath $DataPath -Raw -Encoding UTF8 | ConvertFrom-Json
$outputDir = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Add()
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = 'Vocabulary'

  $headers = @('日期', '单词', '音标', '解释', '例句')
  for ($i = 0; $i -lt $headers.Count; $i++) {
    $cell = $sheet.Cells.Item(1, $i + 1)
    $cell.Value2 = $headers[$i]
    $cell.Font.Bold = $true
    $cell.Font.Name = 'Microsoft YaHei'
    $cell.Font.Color = 0xFFFFFF
    $cell.Interior.Color = 0x0F766E
  }

  $rowIndex = 2
  foreach ($item in $rows) {
    $sheet.Cells.Item($rowIndex, 1).Value2 = $item.date
    $sheet.Cells.Item($rowIndex, 2).Value2 = $item.word
    $sheet.Cells.Item($rowIndex, 3).Value2 = $item.phonetic
    $sheet.Cells.Item($rowIndex, 4).Value2 = $item.definition
    $sheet.Cells.Item($rowIndex, 5).Value2 = $item.example
    $rowIndex++
  }

  $lastRow = [Math]::Max(2, $rows.Count + 1)
  $range = $sheet.Range("A1:E$lastRow")
  $range.WrapText = $true
  $range.VerticalAlignment = -4160
  $range.Borders.LineStyle = 1
  $range.Borders.Color = 0xE5E7EB
  $sheet.Range("A2:E$lastRow").Font.Name = 'Microsoft YaHei'
  $sheet.Range("A2:C$lastRow").HorizontalAlignment = -4108

  $sheet.Columns.Item('A').ColumnWidth = 13
  $sheet.Columns.Item('B').ColumnWidth = 16
  $sheet.Columns.Item('C').ColumnWidth = 20
  $sheet.Columns.Item('D').ColumnWidth = 38
  $sheet.Columns.Item('E').ColumnWidth = 58

  $sheet.Rows.Item(1).AutoFilter()
  $sheet.Application.ActiveWindow.SplitRow = 1
  $sheet.Application.ActiveWindow.FreezePanes = $true

  $xlOpenXMLWorkbook = 51
  $workbook.SaveAs($OutputPath, $xlOpenXMLWorkbook)
}
finally {
  if ($workbook) {
    $workbook.Close($true)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }
  if ($excel) {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
