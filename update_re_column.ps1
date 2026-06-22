$ErrorActionPreference = 'Stop'

$workbookPath = Join-Path $PSScriptRoot 'parser_insert_turn.xlsx'
$formula = '=IF(A2="","",IFERROR(VALUE(MID(A2,9,2))/10,""))'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $workbook = $excel.Workbooks.Open($workbookPath)
  $worksheet = $workbook.Worksheets.Item('Parser')
  $target = $worksheet.Range('F2:F1001')

  $worksheet.Range('F2').Formula = $formula
  $worksheet.Range('F2').AutoFill($target)

  $workbook.Save()
  $workbook.Close($true)
  Write-Host "Updated re column in parser_insert_turn.xlsx"
}
finally {
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
