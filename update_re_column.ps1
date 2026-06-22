$ErrorActionPreference = 'Stop'

$workbookPath = Join-Path $PSScriptRoot 'parser_insert_turn.xlsx'
$mapPath = Join-Path $PSScriptRoot 're_map.json'
$formula = '=IF(A2="","",IFERROR(VLOOKUP(MID(A2,9,2),Map!$R$2:$S$14,2,FALSE),""))'

$reMap = Get-Content -LiteralPath $mapPath -Raw -Encoding UTF8 | ConvertFrom-Json

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $workbook = $excel.Workbooks.Open($workbookPath)
  $parserSheet = $workbook.Worksheets.Item('Parser')
  $mapSheet = $workbook.Worksheets.Item('Map')

  $mapSheet.Columns.Item(18).NumberFormat = '@'

  for ($i = 0; $i -lt $reMap.Count; $i++) {
    $row = $i + 1
    $codeCell = $mapSheet.Cells.Item($row, 18)
    $valueCell = $mapSheet.Cells.Item($row, 19)
    $codeCell.Value2 = [string]$reMap[$i].code
    $value = $reMap[$i].value
    if ($value -is [string]) {
      $valueCell.Value2 = $value
    } else {
      $valueCell.Value2 = [double]$value
    }
  }

  $target = $parserSheet.Range('F2:F1001')
  $parserSheet.Range('F2').Formula = $formula
  $parserSheet.Range('F2').AutoFill($target)

  $workbook.Save()
  $workbook.Close($true)
  Write-Host "Updated re lookup table and formulas in parser_insert_turn.xlsx"
}
finally {
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
