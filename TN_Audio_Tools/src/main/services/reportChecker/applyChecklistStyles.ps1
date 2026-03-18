param(
  [Parameter(Mandatory = $true)]
  [string]$WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string]$SheetName,

  [string]$DecimalCells = '',
  [string]$PercentCells = ''
)

$ErrorActionPreference = 'Stop'

$progIds = @(
  'Excel.Application',
  'ket.Application',
  'et.Application'
)

$xlContinuous = 1
$xlThin = 2
$xlThick = 4
$xlHAlignCenter = -4108
$xlHAlignLeft = -4131
$xlVAlignCenter = -4108
$xlEdgeLeft = 7
$xlEdgeTop = 8
$xlEdgeBottom = 9
$xlEdgeRight = 10
$xlInsideVertical = 11
$xlInsideHorizontal = 12

function Apply-RangeAlignment {
  param(
    $Range,
    [int]$HorizontalAlignment,
    [int]$VerticalAlignment
  )

  $Range.HorizontalAlignment = $HorizontalAlignment
  $Range.VerticalAlignment = $VerticalAlignment
}

function Set-CellBorderEdge {
  param(
    $Cell,
    [int]$Edge,
    [int]$Weight
  )

  $border = $Cell.Borders.Item($Edge)
  $border.LineStyle = $xlContinuous
  $border.Weight = $Weight
}

function Apply-ChecklistStyles {
  param(
    $Worksheet,
    [string[]]$DecimalCellList,
    [string[]]$PercentCellList
  )

  $region = $Worksheet.Range('A3:K75')

  # 在存在 merge 的模板里，直接对整块 Range 设边框会在局部产生异常粗线；
  # 这里改为按物理单元格逐格设置四条边，保证外围粗、内部细稳定落地。
  for ($rowNumber = 3; $rowNumber -le 75; $rowNumber++) {
    for ($columnNumber = 1; $columnNumber -le 11; $columnNumber++) {
      $cell = $Worksheet.Cells.Item($rowNumber, $columnNumber)

      Set-CellBorderEdge -Cell $cell -Edge $xlEdgeLeft -Weight $(if ($columnNumber -eq 1) { $xlThick } else { $xlThin })
      Set-CellBorderEdge -Cell $cell -Edge $xlEdgeRight -Weight $(if ($columnNumber -eq 11) { $xlThick } else { $xlThin })
      Set-CellBorderEdge -Cell $cell -Edge $xlEdgeTop -Weight $(if ($rowNumber -eq 3) { $xlThick } else { $xlThin })
      Set-CellBorderEdge -Cell $cell -Edge $xlEdgeBottom -Weight $(if ($rowNumber -eq 75) { $xlThick } else { $xlThin })
    }
  }

  Apply-RangeAlignment -Range $region -HorizontalAlignment $xlHAlignCenter -VerticalAlignment $xlVAlignCenter

  foreach ($rowNumber in @(5, 66, 71)) {
    $rowRange = $Worksheet.Range("A${rowNumber}:K${rowNumber}")
    Apply-RangeAlignment -Range $rowRange -HorizontalAlignment $xlHAlignLeft -VerticalAlignment $xlVAlignCenter
  }

  foreach ($cell in $DecimalCellList) {
    if (-not [string]::IsNullOrWhiteSpace($cell)) {
      $Worksheet.Range($cell).NumberFormat = '0.00'
    }
  }

  foreach ($cell in $PercentCellList) {
    if (-not [string]::IsNullOrWhiteSpace($cell)) {
      $Worksheet.Range($cell).NumberFormat = '0.00%'
    }
  }
}

$application = $null
$workbook = $null
$worksheet = $null
$selectedProgId = $null
$lastError = $null

foreach ($progId in $progIds) {
  try {
    $application = New-Object -ComObject $progId
    $selectedProgId = $progId
    break
  } catch {
    $lastError = $_
  }
}

if (-not $application) {
  throw "未找到可用的 Excel/WPS COM 组件。最后一次错误: $($lastError.Exception.Message)"
}

try {
  if ($application.PSObject.Properties.Name -contains 'Visible') {
    $application.Visible = $false
  }
  if ($application.PSObject.Properties.Name -contains 'DisplayAlerts') {
    $application.DisplayAlerts = $false
  }

  $workbook = $application.Workbooks.Open($WorkbookPath)
  $worksheet = $workbook.Worksheets.Item($SheetName)
  if (-not $worksheet) {
    $worksheet = $workbook.Worksheets.Item(1)
  }

  $decimalCellList = @()
  if (-not [string]::IsNullOrWhiteSpace($DecimalCells)) {
    $decimalCellList = $DecimalCells.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
  }

  $percentCellList = @()
  if (-not [string]::IsNullOrWhiteSpace($PercentCells)) {
    $percentCellList = $PercentCells.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
  }

  Apply-ChecklistStyles -Worksheet $worksheet -DecimalCellList $decimalCellList -PercentCellList $percentCellList

  $workbook.Save()
  Write-Output "STYLE_ENGINE=$selectedProgId"
} finally {
  if ($worksheet) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
  }

  if ($workbook) {
    try {
      $workbook.Close($true)
    } catch {
      # Ignore close failures during cleanup.
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }

  if ($application) {
    try {
      $application.Quit()
    } catch {
      # Ignore quit failures during cleanup.
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($application)
  }

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
