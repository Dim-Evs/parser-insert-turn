$ErrorActionPreference = 'Stop'

$workbookPath = Join-Path $PSScriptRoot 'parser_insert_turn.xlsx'
$mapPath = Join-Path $PSScriptRoot 'zadnii_ugol_map.json'
$formula = '=IF(A2="","",IFERROR(VLOOKUP(MID(A2,2,1),Map!$U$2:$V$12,2,FALSE),""))'

$angleMap = Get-Content -LiteralPath $mapPath -Raw -Encoding UTF8 | ConvertFrom-Json

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $workbook = $excel.Workbooks.Open($workbookPath)
  $parserSheet = $workbook.Worksheets.Item('Parser')
  $mapSheet = $workbook.Worksheets.Item('Map')

  $mapSheet.Columns.Item(21).NumberFormat = '@'

  for ($i = 0; $i -lt $angleMap.Count; $i++) {
    $row = $i + 1
    $mapSheet.Cells.Item($row, 21).Value2 = [string]$angleMap[$i].code
    $mapSheet.Cells.Item($row, 22).Value2 = [string]$angleMap[$i].value
  }

  $parserSheet.Cells.Item(1, 18).Value2 = 'Zadnii_ugol'
  $target = $parserSheet.Range('R2:R1001')
  $parserSheet.Range('R2').Formula = $formula
  $parserSheet.Range('R2').AutoFill($target)

  $workbook.Save()
  $workbook.Close($true)
  Write-Host "Updated Zadnii_ugol column in parser_insert_turn.xlsx"
}
finally {
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
