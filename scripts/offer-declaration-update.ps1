param(
  [Parameter(Mandatory=$true)][string]$PoNo,
  [Parameter(Mandatory=$true)][string]$DeclarationNo,
  [string]$WorkbookPath = "",
  [switch]$Commit
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

trap {
  Write-Error ("LINE {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
  exit 1
}

$xlUp = -4162

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class MouseClicker {
  [DllImport("user32.dll")]
  public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")]
  public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
  public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
  public const int MOUSEEVENTF_LEFTUP = 0x0004;
}
"@

function Invoke-WithComRetry {
  param(
    [Parameter(Mandatory=$true)][scriptblock]$Operation,
    [string]$Label = "Excel operation",
    [int]$Attempts = 90,
    [int]$DelayMs = 500
  )

  for ($attempt = 1; $attempt -le $Attempts; $attempt += 1) {
    try { return & $Operation } catch {
      $message = $_.Exception.Message
      $hresult = try { [uint32]$_.Exception.HResult } catch { 0 }
      $isBusy =
        $message -like "*Call was rejected by callee*" -or
        $message -like "*RPC_E_CALL_REJECTED*" -or
        $message -like "*0x80010001*" -or
        $message -like "*0x800AC472*" -or
        $hresult -eq 0x80010001 -or
        $hresult -eq 0x800AC472

      if (-not $isBusy -or $attempt -eq $Attempts) { throw }
      Start-Sleep -Milliseconds $DelayMs
    }
  }
}

function Normalize-PoNo {
  param($Value)
  return ([string]$Value).
    Normalize([Text.NormalizationForm]::FormKC).
    Replace([string][char]0x2010, "-").
    Replace([string][char]0x2011, "-").
    Replace([string][char]0x2012, "-").
    Replace([string][char]0x2013, "-").
    Replace([string][char]0x2014, "-").
    Replace([string][char]0x2015, "-").
    Replace([string][char]0x2212, "-").
    Replace(" ", "").
    Trim().
    ToUpperInvariant()
}

function U {
  param([int[]]$Codes)
  return -join ($Codes | ForEach-Object { [char]$_ })
}

function Resolve-WorkbookPath {
  param([string]$RequestedPath)
  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) {
    return (Resolve-Path -LiteralPath $RequestedPath).Path
  }
  if ($env:OFFER_LIST_PATH -and (Test-Path -LiteralPath $env:OFFER_LIST_PATH)) {
    return (Resolve-Path -LiteralPath $env:OFFER_LIST_PATH).Path
  }

  $offerListText = U 0xC624,0xD37C,0xBC1C,0xD589,0xB0B4,0xC5ED
  $formatText = U 0xC591,0xC2DD,0xC218,0xC815
  $desktopName = "{0}C8-2({1}).xlsm" -f $offerListText, $formatText
  $desktopPath = Join-Path ([Environment]::GetFolderPath("Desktop")) $desktopName
  if (Test-Path -LiteralPath $desktopPath) {
    return (Resolve-Path -LiteralPath $desktopPath).Path
  }

  $serverDir = "Z:\{0}\{1}(Main)\A60-{2}" -f `
    (U 0xB3D9,0xC9C4,0xD30C,0xB9C8), `
    (U 0xACF5,0xC720,0xBB38,0xC11C), `
    (U 0xC624,0xD37C,0xB9AC,0xC2A4,0xD2B8)
  $defaultPath = Join-Path $serverDir $desktopName
  if (Test-Path -LiteralPath $defaultPath) {
    return (Resolve-Path -LiteralPath $defaultPath).Path
  }

  if (Test-Path -LiteralPath $serverDir) {
    $candidate = Get-ChildItem -LiteralPath $serverDir -File -Filter "*C8-2*.xlsm" -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -like "*$offerListText*" } |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    if ($candidate) {
      return $candidate.FullName
    }
  }

  throw "Offer list workbook was not found. Set OFFER_LIST_PATH."
}

function Get-RunningExcel {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { return $null }
}

function Find-Workbook {
  param($Excel, [string]$Path)
  foreach ($workbook in @($Excel.Workbooks)) {
    try { if ($workbook.FullName -eq $Path) { return $workbook } } catch {}
  }
  return $null
}

function Run-WorkbookMacro {
  param($Excel, $Workbook, [string]$MacroName)
  $macro = "'{0}'!{1}" -f $Workbook.Name, $MacroName
  Invoke-WithComRetry { $Excel.Run($macro) | Out-Null } "macro: $MacroName"
}

function Invoke-FirstMacro {
  param($Excel, $Workbook, $Sheet, [string[]]$MacroNames, [string]$Label)
  $errors = @()

  foreach ($macroName in $MacroNames) {
    if (-not $macroName) { continue }
    try {
      Run-WorkbookMacro $Excel $Workbook $macroName
      return $macroName
    } catch {
      $errors += ("{0}: {1}" -f $macroName, $_.Exception.Message)
    }
  }

  if ($Sheet) {
    foreach ($shape in @($Sheet.Shapes)) {
      $shapeText = ""
      $shapeAction = ""
      try { $shapeText = [string]$shape.TextFrame2.TextRange.Text } catch {}
      if (-not $shapeText) {
        try { $shapeText = [string]$shape.TextFrame.Characters().Text } catch {}
      }
      if (-not $shapeText) {
        try { $shapeText = [string]$shape.AlternativeText } catch {}
      }
      if (-not $shapeText) {
        try { $shapeText = [string]$shape.Title } catch {}
      }
      try { $shapeAction = [string]$shape.OnAction } catch {}
      if (-not $shapeAction) {
        try { $shapeAction = [string]$shape.DrawingObject.OnAction } catch {}
      }
      if (-not $shapeAction) {
        try { $shapeAction = [string]$shape.OLEFormat.Object.OnAction } catch {}
      }

      $normalizedShapeText = $shapeText.Trim() -replace "\s+", ""
      $matchesText = @($MacroNames | Where-Object {
        if (-not $_) { return $false }
        $normalizedMacroName = $_ -replace "\s+", ""
        return $normalizedShapeText -eq $normalizedMacroName -or $normalizedShapeText -like ("*{0}*" -f $normalizedMacroName)
      }).Count -gt 0
      $matchesAction = @($MacroNames | Where-Object { $_ -and $shapeAction -like "*$_*" }).Count -gt 0

      if (($matchesText -or $matchesAction) -and $shapeAction) {
        try {
          Invoke-WithComRetry { $Excel.Run($shapeAction) | Out-Null } "shape action: $shapeAction"
          return $shapeAction
        } catch {
          $errors += ("shape {0}: {1}" -f $normalizedShapeText, $_.Exception.Message)
        }
      }
    }
  }

  throw "$Label macro failed. $($errors -join ' / ')"
}

function Click-WorksheetCellCenter {
  param($Excel, $Sheet, [string]$Address, [string]$Label)
  Invoke-WithComRetry { $Sheet.Activate() | Out-Null } "activate sheet for $Label"
  $range = Invoke-WithComRetry { $Sheet.Range($Address) } "range for $Label"
  $window = Invoke-WithComRetry { $Excel.ActiveWindow } "active window for $Label"
  $x = Invoke-WithComRetry { $window.PointsToScreenPixelsX($range.Left + ($range.Width / 2)) } "screen x for $Label"
  $y = Invoke-WithComRetry { $window.PointsToScreenPixelsY($range.Top + ($range.Height / 2)) } "screen y for $Label"
  [MouseClicker]::SetCursorPos([int]$x, [int]$y) | Out-Null
  Start-Sleep -Milliseconds 150
  [MouseClicker]::mouse_event([MouseClicker]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
  Start-Sleep -Milliseconds 80
  [MouseClicker]::mouse_event([MouseClicker]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
  Start-Sleep -Milliseconds 700
  return "clicked:$Address"
}

function Select-WorksheetCell {
  param($Excel, $Sheet, [string]$Address, [string]$Label)

  Invoke-WithComRetry { $Sheet.Parent.Activate() | Out-Null } "activate workbook for $Label"
  Invoke-WithComRetry { $Sheet.Activate() | Out-Null } "activate sheet for $Label"
  $range = Invoke-WithComRetry { $Sheet.Range($Address) } "range for $Label"

  try {
    $window = Invoke-WithComRetry { $Excel.ActiveWindow } "active window for $Label"
    Invoke-WithComRetry { $window.ScrollRow = [Math]::Max(1, [int]$range.Row - 2) } "scroll target row for $Label"
    Invoke-WithComRetry { $window.ScrollColumn = [Math]::Max(1, [int]$range.Column - 2) } "scroll target column for $Label"
  } catch {}

  try {
    Invoke-WithComRetry { $Excel.Goto($range, $true) | Out-Null } "goto $Address for $Label"
    Start-Sleep -Milliseconds 250
    return "goto:$Address"
  } catch {}

  Invoke-WithComRetry { $range.Select() | Out-Null } "select $Address for $Label"
  Start-Sleep -Milliseconds 250
  return "select:$Address"
}

function Invoke-MacroOrColumnButton {
  param($Excel, $Workbook, $Sheet, [string[]]$MacroNames, [string]$ColumnLetter, [string]$Label)
  try {
    return Invoke-FirstMacro $Excel $Workbook $Sheet $MacroNames $Label
  } catch {
    $macroError = $_.Exception.Message
    try {
      return Click-WorksheetCellCenter $Excel $Sheet ("{0}1" -f $ColumnLetter) $Label
    } catch {
      throw "$Label failed. Macro error: $macroError / Click error: $($_.Exception.Message)"
    }
  }
}

$resolvedPath = Resolve-WorkbookPath $WorkbookPath
$normalizedPoNo = Normalize-PoNo $PoNo
$declarationText = $DeclarationNo.Trim()
if (-not $declarationText) { throw "Declaration number is empty." }
$sheetName = "PO{0}" -f (U 0xBA54,0xC778)
$copyMacroName = "PO{0}" -f (U 0xBCF5,0xC0AC)
$orderSelectMacroName = "{0}{1}" -f (U 0xC624,0xB354), (U 0xC120,0xD0DD)
$poOrderSelectMacroName = "PO{0}{1}" -f (U 0xC218,0xC815), (U 0xC120,0xD0DD)
$updateMacroName = U 0xC218,0xC815
$poUpdateMacroName = "PO{0}" -f (U 0xC218,0xC815)

$excel = Get-RunningExcel
$createdExcel = $false
$openedWorkbook = $false
if (-not $excel) {
  $excel = New-Object -ComObject Excel.Application
  $createdExcel = $true
}

$workbook = $null
$ranOrderSelectMacro = ""
$ranUpdateMacro = ""

try {
  Invoke-WithComRetry { $excel.Visible = $true } "show Excel"
  Invoke-WithComRetry { $excel.DisplayAlerts = $false } "DisplayAlerts"

  $workbook = Invoke-WithComRetry { Find-Workbook $excel $resolvedPath } "find offer workbook"
  if (-not $workbook) {
    $workbook = Invoke-WithComRetry { $excel.Workbooks.Open($resolvedPath) } "open offer workbook"
    $openedWorkbook = $true
  }

  $sheet = Invoke-WithComRetry { $workbook.Worksheets.Item($sheetName) } "open PO sheet"
  Invoke-WithComRetry { $sheet.Activate() | Out-Null } "activate PO sheet"
  Invoke-WithComRetry { $sheet.Range("B1").Value2 = $normalizedPoNo } "enter PO search"
  try { Run-WorkbookMacro $excel $workbook $copyMacroName } catch {}

  $lastRow = Invoke-WithComRetry { $sheet.Range("A500000").End($xlUp).Row } "last row"
  $targetRow = 0
  for ($row = 3; $row -le $lastRow; $row += 1) {
    $candidates = @($sheet.Cells.Item($row, 1).Text, $sheet.Cells.Item($row, 2).Text, $sheet.Cells.Item($row, 3).Text)
    foreach ($candidate in $candidates) {
      if ((Normalize-PoNo $candidate) -eq $normalizedPoNo) {
        $targetRow = $row
        break
      }
    }
    if ($targetRow -gt 0) { break }
  }
  if ($targetRow -le 0) { throw "PO '$PoNo' was not found in offer list." }

  $beforeDeclarationNo = [string]$sheet.Cells.Item($targetRow, 34).Text

  if ($Commit) {
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 34).Value2 = $declarationText } "set declaration no"
    Select-WorksheetCell $excel $sheet "A$targetRow" "select offer row" | Out-Null
    $ranOrderSelectMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($orderSelectMacroName, $poOrderSelectMacroName, "OrderSelect", "SelectOrder") "A" "offer order select"
    Select-WorksheetCell $excel $sheet "K$targetRow" "select offer update cell" | Out-Null
    $ranUpdateMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($updateMacroName, $poUpdateMacroName, "Update", "Modify") "K" "offer update"
    Invoke-WithComRetry { $workbook.Save() } "save offer workbook"
  }

  [ordered]@{
    ok = $true
    committed = [bool]$Commit
    workbookPath = $resolvedPath
    sheetName = $sheetName
    row = $targetRow
    poNo = [string]$sheet.Cells.Item($targetRow, 1).Text
    declarationNo = $declarationText
    beforeDeclarationNo = $beforeDeclarationNo
    orderSelectMacro = $ranOrderSelectMacro
    updateMacro = $ranUpdateMacro
  } | ConvertTo-Json -Depth 5 -Compress
} finally {
  if ($workbook -and $openedWorkbook) {
    try { $workbook.Close([bool]$Commit) } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
  }
  if ($excel -and $createdExcel) {
    try { $excel.Quit() } catch {}
  }
  if ($excel) {
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  try {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  } catch {}
}
