param(
  [Parameter(Mandatory = $true)]
  [string]$WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string]$SheetName,

  [string]$DecimalCells = '',
  [string]$PercentCells = '',
  [string]$SkippedCells = ''
)

$ErrorActionPreference = 'Stop'

$progIds = @(
  'Excel.Application',
  'ket.Application',
  'et.Application'
)

$xlPatternSolid = 1

function Convert-RgbToOleColor {
  param(
    [int]$Red,
    [int]$Green,
    [int]$Blue
  )

  return ($Blue -shl 16) -bor ($Green -shl 8) -bor $Red
}

function Parse-CellList {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return @()
  }

  return $Value.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Set-RangeFill {
  param(
    $Range,
    [int]$ColorValue
  )

  $Range.Interior.Pattern = $xlPatternSolid
  $Range.Interior.Color = $ColorValue
}

function Apply-ChecklistStyles {
  param(
    $Worksheet,
    [string[]]$DecimalCellList,
    [string[]]$PercentCellList,
    [string[]]$SkippedCellList
  )

  $skippedFillColor = Convert-RgbToOleColor -Red 217 -Green 217 -Blue 217

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

  foreach ($cell in $SkippedCellList) {
    if (-not [string]::IsNullOrWhiteSpace($cell) -and $cell -match '^[Ii]\d+$') {
      Set-RangeFill -Range $Worksheet.Range($cell) -ColorValue $skippedFillColor
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

  $decimalCellList = Parse-CellList -Value $DecimalCells
  $percentCellList = Parse-CellList -Value $PercentCells
  $skippedCellList = Parse-CellList -Value $SkippedCells

  Apply-ChecklistStyles -Worksheet $worksheet -DecimalCellList $decimalCellList -PercentCellList $percentCellList -SkippedCellList $skippedCellList

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
