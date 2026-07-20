param(
  [string]$WorkbookPath = (Join-Path $env:USERPROFILE "Desktop\151515.xlsx"),
  [string]$Password = $env:PASSWORD_VAULT_EXCEL_PASSWORD,
  [int]$MaxRows = 1000,
  [int]$MaxColumns = 40
)

$ErrorActionPreference = "Stop"

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

if (-not $Password) {
  throw "PASSWORD_VAULT_EXCEL_PASSWORD is required."
}

$resolvedPath = Resolve-Path -LiteralPath $WorkbookPath -ErrorAction Stop
$excel = $null
$workbook = $null
$oldDisplayAlerts = $null
$oldScreenUpdating = $null
$oldEnableEvents = $null

function Get-CellText {
  param($Cell)

  $text = ""

  try {
    $text = [string]$Cell.Text
  } catch {
    $text = ""
  }

  if ([string]::IsNullOrWhiteSpace($text) -or $text -match "^#+$") {
    try {
      $value = $Cell.Value2
      if ($null -ne $value) {
        $text = [string]$value
      }
    } catch {
      $text = ""
    }
  }

  return (($text -replace "\r?\n", " ").Trim())
}

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $oldDisplayAlerts = $excel.DisplayAlerts
  $oldScreenUpdating = $excel.ScreenUpdating
  $oldEnableEvents = $excel.EnableEvents
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.EnableEvents = $false
  $excel.AutomationSecurity = 3

  $missing = [Type]::Missing
  $workbook = $excel.Workbooks.Open($resolvedPath.Path, 0, $true, $missing, $Password)
  $sheets = @()

  foreach ($sheet in $workbook.Worksheets) {
    $usedRange = $sheet.UsedRange
    $rowCount = [Math]::Min([int]$usedRange.Rows.Count, $MaxRows)
    $columnCount = [Math]::Min([int]$usedRange.Columns.Count, $MaxColumns)
    $rows = @()

    for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
      $rowValues = New-Object 'System.Collections.Generic.List[string]'
      $hasValue = $false

      for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
        $text = Get-CellText $usedRange.Cells.Item($rowIndex, $columnIndex)

        if (-not [string]::IsNullOrWhiteSpace($text)) {
          $hasValue = $true
        }

        [void]$rowValues.Add($text)
      }

      while ($rowValues.Count -gt 0 -and [string]::IsNullOrWhiteSpace($rowValues[$rowValues.Count - 1])) {
        $rowValues.RemoveAt($rowValues.Count - 1)
      }

      if ($hasValue) {
        $rows += ,@($rowValues.ToArray())
      }
    }

    $sheets += [pscustomobject]@{
      name = [string]$sheet.Name
      rows = $rows
    }

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
  }

  $result = [pscustomobject]@{
    ok = $true
    file = [System.IO.Path]::GetFileName($resolvedPath.Path)
    updatedAt = (Get-Item -LiteralPath $resolvedPath.Path).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    sheets = $sheets
  }

  $result | ConvertTo-Json -Depth 12 -Compress
} finally {
  if ($workbook) {
    try { $workbook.Close($false) | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
  }

  if ($excel) {
    try { if ($null -ne $oldEnableEvents) { $excel.EnableEvents = $oldEnableEvents } } catch {}
    try { if ($null -ne $oldScreenUpdating) { $excel.ScreenUpdating = $oldScreenUpdating } } catch {}
    try { if ($null -ne $oldDisplayAlerts) { $excel.DisplayAlerts = $oldDisplayAlerts } } catch {}
    try { $excel.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
