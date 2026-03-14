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

function Apply-ChecklistStyles {
  param(
    $Worksheet,
    [string[]]$DecimalCellList,
    [string[]]$PercentCellList
  )

  $region = $Worksheet.Range('A3:K75')

  # 先铺整块细线，再把外边框提升成粗线，这个顺序最接近手工在 Excel 里的操作。
  $region.Borders.LineStyle = $xlContinuous
  $region.Borders.Weight = $xlThin
  $region.Borders.Item($xlInsideVertical).LineStyle = $xlContinuous
  $region.Borders.Item($xlInsideVertical).Weight = $xlThin
  $region.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
  $region.Borders.Item($xlInsideHorizontal).Weight = $xlThin

  foreach ($edge in @($xlEdgeLeft, $xlEdgeTop, $xlEdgeBottom, $xlEdgeRight)) {
    $region.Borders.Item($edge).LineStyle = $xlContinuous
    $region.Borders.Item($edge).Weight = $xlThick
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
