param(
  [Parameter(Mandatory = $true)]
  [string]$WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string]$SheetName,

  [string]$DecimalCells = '',
  [string]$PercentCells = '',
  [string]$SkippedCells = '',
  [string]$UpdatePayloadPath = ''
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

function Parse-JsonTextList {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return @()
  }

  try {
    $parsed = $Value | ConvertFrom-Json
    if ($parsed -is [System.Array]) {
      return $parsed
    }

    if ($null -ne $parsed) {
      return @($parsed)
    }
  } catch {
    return @()
  }

  return @()
}

function Read-UpdatePayload {
  param([string]$PayloadPath)

  if ([string]::IsNullOrWhiteSpace($PayloadPath) -or -not (Test-Path -LiteralPath $PayloadPath)) {
    return @{ valueUpdates = @(); reportUpdates = @() }
  }

  try {
    $payloadText = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8
    $payload = $payloadText | ConvertFrom-Json
    return @{
      valueUpdates = Parse-JsonTextList -Value ($payload.valueUpdates | ConvertTo-Json -Compress)
      reportUpdates = Parse-JsonTextList -Value ($payload.reportUpdates | ConvertTo-Json -Compress)
    }
  } catch {
    return @{ valueUpdates = @(); reportUpdates = @() }
  }
}

function Get-WorksheetSafely {
  param(
    $Workbook,
    [string]$WorksheetName
  )

  if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
    return $null
  }

  try {
    return $Workbook.Worksheets.Item($WorksheetName)
  } catch {
    return $null
  }
}

function Set-CellValue {
  param(
    $Worksheet,
    [string]$CellAddress,
    $CellValue
  )

  if (-not $Worksheet -or [string]::IsNullOrWhiteSpace($CellAddress)) {
    return
  }

  $range = $null
  try {
    $range = $Worksheet.Range($CellAddress)
    if ($CellValue -is [double] -or $CellValue -is [float] -or $CellValue -is [decimal] -or $CellValue -is [int] -or $CellValue -is [long]) {
      $range.Value2 = $CellValue
      return
    }

    if ($null -eq $CellValue) {
      $range.Value2 = ''
      return
    }

    $range.Value2 = [string]$CellValue
  } finally {
    if ($range) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range)
    }
  }
}

function Apply-CellUpdates {
  param(
    $Workbook,
    [string]$DefaultSheetName,
    $Updates
  )

  foreach ($entry in $Updates) {
    if (-not $entry) {
      continue
    }

    $targetSheetName = [string]($entry.sheetName)
    if ([string]::IsNullOrWhiteSpace($targetSheetName)) {
      $targetSheetName = $DefaultSheetName
    }

    $targetSheet = $null
    try {
      $targetSheet = Get-WorksheetSafely -Workbook $Workbook -WorksheetName $targetSheetName
      if (-not $targetSheet) {
        continue
      }

      Set-CellValue -Worksheet $targetSheet -CellAddress ([string]$entry.cellAddress) -CellValue $entry.value
    } finally {
      if ($targetSheet -and $targetSheet.Name -ne $DefaultSheetName) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetSheet)
      }
    }
  }
}

function Set-RangeFill {
  param(
    $Range,
    [int]$ColorValue
  )

  $Range.Interior.Pattern = $xlPatternSolid
  $Range.Interior.Color = $ColorValue
}

function Set-CellNumberFormat {
  param(
    $Worksheet,
    [string]$CellAddress,
    [string]$NumberFormat
  )

  $range = $null
  try {
    $range = $Worksheet.Range($CellAddress)
    $range.NumberFormat = $NumberFormat
  } finally {
    if ($range) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range)
    }
  }
}

function Set-CellFillColor {
  param(
    $Worksheet,
    [string]$CellAddress,
    [int]$ColorValue
  )

  $range = $null
  try {
    $range = $Worksheet.Range($CellAddress)
    Set-RangeFill -Range $range -ColorValue $ColorValue
  } finally {
    if ($range) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range)
    }
  }
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
      Set-CellNumberFormat -Worksheet $Worksheet -CellAddress $cell -NumberFormat '0.00'
    }
  }

  foreach ($cell in $PercentCellList) {
    if (-not [string]::IsNullOrWhiteSpace($cell)) {
      Set-CellNumberFormat -Worksheet $Worksheet -CellAddress $cell -NumberFormat '0.00%'
    }
  }

  foreach ($cell in $SkippedCellList) {
    if (-not [string]::IsNullOrWhiteSpace($cell) -and $cell -match '^[Ii]\d+$') {
      Set-CellFillColor -Worksheet $Worksheet -CellAddress $cell -ColorValue $skippedFillColor
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
  if ($application.PSObject.Properties.Name -contains 'ScreenUpdating') {
    $application.ScreenUpdating = $false
  }

  $workbook = $application.Workbooks.Open($WorkbookPath)
  $worksheet = $workbook.Worksheets.Item($SheetName)
  if (-not $worksheet) {
    $worksheet = $workbook.Worksheets.Item(1)
  }

  if ($application.PSObject.Properties.Name -contains 'Calculation') {
    try {
      $application.Calculation = -4135
    } catch {
      # Some Excel/WPS COM engines expose Calculation but reject assignment.
    }
  }

  $decimalCellList = Parse-CellList -Value $DecimalCells
  $percentCellList = Parse-CellList -Value $PercentCells
  $skippedCellList = Parse-CellList -Value $SkippedCells
  $payload = Read-UpdatePayload -PayloadPath $UpdatePayloadPath
  $valueUpdates = $payload.valueUpdates
  $reportUpdates = $payload.reportUpdates

  Apply-CellUpdates -Workbook $workbook -DefaultSheetName $SheetName -Updates $reportUpdates
  Apply-CellUpdates -Workbook $workbook -DefaultSheetName $SheetName -Updates $valueUpdates

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
