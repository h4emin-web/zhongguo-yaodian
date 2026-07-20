param(
  [Parameter(Mandatory=$true)][string]$PoNo,
  [Parameter(Mandatory=$true)][string]$InstockDate,
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
  [DllImport("user32.dll")]
  public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
  public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
  public const int MOUSEEVENTF_LEFTUP = 0x0004;
  public const int SW_RESTORE = 9;
}
"@

function U {
  param([int[]]$Codes)
  return -join ($Codes | ForEach-Object { [char]$_ })
}

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

function Find-PoRow {
  param($Sheet, [string]$NormalizedPoNo)

  $lastRow = Invoke-WithComRetry { $Sheet.Range("A500000").End($xlUp).Row } "last row"

  for ($row = 3; $row -le $lastRow; $row += 1) {
    $candidates = @($Sheet.Cells.Item($row, 1).Text, $Sheet.Cells.Item($row, 2).Text, $Sheet.Cells.Item($row, 3).Text)
    foreach ($candidate in $candidates) {
      if ((Normalize-PoNo $candidate) -eq $NormalizedPoNo) {
        return $row
      }
    }
  }

  return 0
}

function Convert-ToNumber {
  param($Value)
  $match = [regex]::Match(([string]$Value), "-?\d+(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?")
  if ($match.Success) {
    return [double]::Parse($match.Value.Replace(",", ""), [Globalization.CultureInfo]::InvariantCulture)
  }
  return 0
}

function Convert-ToDateValue {
  param($Value)
  if ($Value -is [datetime]) { return $Value.Date }
  if ($Value -is [double] -or $Value -is [int]) {
    try { return [datetime]::FromOADate([double]$Value).Date } catch {}
  }
  $text = ([string]$Value).Trim().Replace("/", "-").Replace(".", "-")
  $formats = @("yyyy-MM-dd", "yyyyMMdd", "yy-MM-dd")
  $parsed = [datetime]::MinValue
  foreach ($format in $formats) {
    if ([datetime]::TryParseExact($text, $format, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
      return $parsed.Date
    }
  }
  if ([datetime]::TryParse($text, [ref]$parsed)) { return $parsed.Date }
  return $null
}

function Resolve-WorkbookPath {
  param([string]$RequestedPath)
  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) {
    return (Resolve-Path -LiteralPath $RequestedPath).Path
  }
  if ($env:OFFER_LIST_PATH -and (Test-Path -LiteralPath $env:OFFER_LIST_PATH)) {
    return (Resolve-Path -LiteralPath $env:OFFER_LIST_PATH).Path
  }

  $desktopName = "{0}C8-2({1}).xlsm" -f (U 0xC624,0xD37C,0xBC1C,0xD589,0xB0B4,0xC5ED), (U 0xC591,0xC2DD,0xC218,0xC815)
  $desktopPath = Join-Path ([Environment]::GetFolderPath("Desktop")) $desktopName
  if (Test-Path -LiteralPath $desktopPath) {
    return (Resolve-Path -LiteralPath $desktopPath).Path
  }

  $defaultPath = "Z:\{0}\{1}(Main)\A60-{2}\{3}C8-2({4}).xlsm" -f `
    (U 0xB3D9,0xC9C4,0xD30C,0xB9C8), `
    (U 0xACF5,0xC720,0xBB38,0xC11C), `
    (U 0xC624,0xD37C,0xB9AC,0xC2A4,0xD2B8), `
    (U 0xC624,0xD37C,0xBC1C,0xD589,0xB0B4,0xC5ED), `
    (U 0xC591,0xC2DD,0xC218,0xC815)
  if (Test-Path -LiteralPath $defaultPath) {
    return (Resolve-Path -LiteralPath $defaultPath).Path
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
      try { $shapeAction = [string]$shape.OnAction } catch {}

      $normalizedShapeText = $shapeText.Trim()
      $matchesText = @($MacroNames | Where-Object { $_ -and ($normalizedShapeText -eq $_ -or $normalizedShapeText -like "*$_*") }).Count -gt 0
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
  try {
    [MouseClicker]::ShowWindow([IntPtr]$Excel.Hwnd, [MouseClicker]::SW_RESTORE) | Out-Null
    [MouseClicker]::SetForegroundWindow([IntPtr]$Excel.Hwnd) | Out-Null
    Start-Sleep -Milliseconds 250
  } catch {}
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

function Invoke-TopButtonInColumn {
  param($Excel, $Sheet, [string]$ColumnLetter, [string[]]$Needles, [string]$Label)

  Invoke-WithComRetry { $Sheet.Activate() | Out-Null } "activate sheet for $Label"
  try {
    [MouseClicker]::ShowWindow([IntPtr]$Excel.Hwnd, [MouseClicker]::SW_RESTORE) | Out-Null
    [MouseClicker]::SetForegroundWindow([IntPtr]$Excel.Hwnd) | Out-Null
    Start-Sleep -Milliseconds 250
  } catch {}
  $window = Invoke-WithComRetry { $Excel.ActiveWindow } "active window for $Label"
  Invoke-WithComRetry { $window.ScrollRow = 1 } "scroll row for $Label"
  Invoke-WithComRetry { $window.ScrollColumn = 1 } "scroll column for $Label"

  $targetColumn = Invoke-WithComRetry { $Sheet.Range(("{0}1" -f $ColumnLetter)).Column } "button column for $Label"
  $candidates = @()

  foreach ($shape in @($Sheet.Shapes)) {
    $text = ""
    try { $text = [string]$shape.TextFrame2.TextRange.Text } catch {}
    if (-not $text) {
      try { $text = [string]$shape.TextFrame.Characters().Text } catch {}
    }

    $normalizedText = $text.Trim() -replace "\s+", ""
    $matchesNeedle = $false
    foreach ($needle in $Needles) {
      if (-not $needle) { continue }
      $normalizedNeedle = $needle -replace "\s+", ""
      if ($normalizedText -eq $normalizedNeedle -or $normalizedText -like ("*{0}*" -f $normalizedNeedle)) {
        $matchesNeedle = $true
        break
      }
    }

    $topRow = 0
    $leftColumn = 0
    $rightColumn = 0
    try { $topRow = [int]$shape.TopLeftCell.Row } catch {}
    try { $leftColumn = [int]$shape.TopLeftCell.Column } catch {}
    try { $rightColumn = [int]$shape.BottomRightCell.Column } catch { $rightColumn = $leftColumn }

    if (($matchesNeedle -or ($leftColumn -le $targetColumn -and $rightColumn -ge $targetColumn)) -and $topRow -le 3) {
      $candidates += $shape
    }
  }

  $targetShape = $candidates | Sort-Object @{ Expression = { try { [double]$_.Top } catch { 999999 } } }, @{ Expression = { try { [double]$_.Left } catch { 999999 } } } | Select-Object -First 1
  if ($targetShape) {
    $shapeName = try { [string]$targetShape.Name } catch { "" }
    $shapeAction = try { [string]$targetShape.OnAction } catch { "" }

    if ($shapeAction) {
      Invoke-WithComRetry { $Excel.Run($shapeAction) | Out-Null } "shape action for $Label"
      Start-Sleep -Milliseconds 1200
      return ("shape-action:{0}:{1}" -f $shapeName, $shapeAction)
    }

    $x = Invoke-WithComRetry { $window.PointsToScreenPixelsX($targetShape.Left + ($targetShape.Width / 2)) } "shape screen x for $Label"
    $y = Invoke-WithComRetry { $window.PointsToScreenPixelsY($targetShape.Top + ($targetShape.Height / 2)) } "shape screen y for $Label"

    [MouseClicker]::SetCursorPos([int]$x, [int]$y) | Out-Null
    Start-Sleep -Milliseconds 150
    [MouseClicker]::mouse_event([MouseClicker]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 80
    [MouseClicker]::mouse_event([MouseClicker]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 1200
    return "clicked-shape:$shapeName"
  }

  throw "$Label top button was not found in column $ColumnLetter."
}

function Invoke-MacroOrColumnButton {
  param($Excel, $Workbook, $Sheet, [string[]]$MacroNames, [string]$ColumnLetter, [string]$Label)

  try {
    return Invoke-TopButtonInColumn $Excel $Sheet $ColumnLetter $MacroNames $Label
  } catch {
    $buttonError = $_.Exception.Message
    try {
      return Invoke-FirstMacro $Excel $Workbook $Sheet $MacroNames $Label
    } catch {
      throw "$Label failed. Button error: $buttonError / Macro error: $($_.Exception.Message)"
    }
  }
}

$resolvedPath = Resolve-WorkbookPath $WorkbookPath
$instockDateValue = Convert-ToDateValue $InstockDate
if (-not $instockDateValue) { throw "Invalid instock date: $InstockDate" }

$normalizedPoNo = Normalize-PoNo $PoNo
$poMainSheetName = "PO{0}" -f (U 0xBA54,0xC778)
$statusInStock = U 0xC785,0xACE0
$copyMacroName = "PO{0}" -f (U 0xBCF5,0xC0AC)
$orderSelectMacroName = if ($env:PENDING_RECEIPT_OFFER_SELECT_MACRO) { $env:PENDING_RECEIPT_OFFER_SELECT_MACRO } else { "{0}{1}" -f (U 0xC624,0xB354), (U 0xC120,0xD0DD) }
$updateMacroName = if ($env:PENDING_RECEIPT_OFFER_UPDATE_MACRO) { $env:PENDING_RECEIPT_OFFER_UPDATE_MACRO } else { U 0xC218,0xC815 }
$offerMacroWarning = ""
$ranOrderSelectMacro = ""
$ranUpdateMacro = ""

$excel = Get-RunningExcel
$createdExcel = $false
$openedWorkbook = $false

if (-not $excel) {
  $excel = New-Object -ComObject Excel.Application
  $createdExcel = $true
}

$workbook = $null

try {
  Invoke-WithComRetry { $excel.Visible = $true } "show Excel"
  Invoke-WithComRetry { $excel.DisplayAlerts = $false } "DisplayAlerts"

  $workbook = Invoke-WithComRetry { Find-Workbook $excel $resolvedPath } "find offer workbook"
  if (-not $workbook) {
    $workbook = Invoke-WithComRetry { $excel.Workbooks.Open($resolvedPath) } "open offer workbook"
    $openedWorkbook = $true
  }

  $sheet = Invoke-WithComRetry { $workbook.Worksheets.Item($poMainSheetName) } "open PO sheet"
  Invoke-WithComRetry { $sheet.Activate() | Out-Null } "activate PO sheet"
  Invoke-WithComRetry { $sheet.Range("B1").Value2 = $normalizedPoNo } "enter PO search"
  try { Run-WorkbookMacro $excel $workbook $copyMacroName } catch {}

  $targetRow = Find-PoRow $sheet $normalizedPoNo

  if ($targetRow -le 0) {
    throw "PO '$PoNo' was not found in offer list."
  }

  $beforeStatus = [string]$sheet.Cells.Item($targetRow, 3).Text
  $beforeInstock = [string]$sheet.Cells.Item($targetRow, 25).Text
  $verifiedRow = $targetRow
  $verifiedStatus = ""
  $verifiedInstock = ""

  if ($Commit) {
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 3).Value2 = $statusInStock } "set status"
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 25).Value2 = $instockDateValue.ToString("yyyy-MM-dd") } "set instock"
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 25).NumberFormat = "yyyy-mm-dd" } "instock format"

    try {
      Invoke-WithComRetry { $sheet.Range("A$targetRow").Select() | Out-Null } "select offer row"
      $ranOrderSelectMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($orderSelectMacroName, "PO수정선택", "OrderSelect", "SelectOrder") "A" "offer order select"
      Invoke-WithComRetry { $sheet.Range("K$targetRow").Select() | Out-Null } "select offer update cell"
      $ranUpdateMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($updateMacroName, "PO수정", "Update", "Modify") "K" "offer update"
    } catch {
      $offerMacroWarning = "Offer list server update macro failed: $($_.Exception.Message)"
      throw $offerMacroWarning
    }

    Invoke-WithComRetry { $workbook.Save() } "save offer workbook"

    Invoke-WithComRetry { $sheet.Range("B1").Value2 = $normalizedPoNo } "re-enter PO search"
    Run-WorkbookMacro $excel $workbook $copyMacroName
    $verifiedRow = Find-PoRow $sheet $normalizedPoNo
    if ($verifiedRow -le 0) {
      throw "Offer list save verification failed: PO '$PoNo' was not found after update."
    }

    $verifiedStatus = ([string]$sheet.Cells.Item($verifiedRow, 3).Text).Trim()
    $verifiedInstockDate = Convert-ToDateValue $sheet.Cells.Item($verifiedRow, 25).Value2
    $verifiedInstock = if ($verifiedInstockDate) { $verifiedInstockDate.ToString("yyyy-MM-dd") } else { ([string]$sheet.Cells.Item($verifiedRow, 25).Text).Trim() }

    if ($verifiedStatus -ne $statusInStock -or $verifiedInstock -ne $instockDateValue.ToString("yyyy-MM-dd")) {
      throw ("Offer list save verification failed. status={0}, instock={1}" -f $verifiedStatus, $verifiedInstock)
    }

    $targetRow = $verifiedRow
  }

  $result = [ordered]@{
    ok = $true
    committed = [bool]$Commit
    workbookPath = $resolvedPath
    sheetName = $poMainSheetName
    row = $targetRow
    poNo = [string]$sheet.Cells.Item($targetRow, 1).Text
    messrs = [string]$sheet.Cells.Item($targetRow, 2).Text
    status = [string]$sheet.Cells.Item($targetRow, 3).Text
    poDate = (Convert-ToDateValue $sheet.Cells.Item($targetRow, 4).Value2).ToString("yyyy-MM-dd")
    maker = [string]$sheet.Cells.Item($targetRow, 7).Text
    supplier = [string]$sheet.Cells.Item($targetRow, 8).Text
    productCode = [string]$sheet.Cells.Item($targetRow, 9).Text
    productName = [string]$sheet.Cells.Item($targetRow, 10).Text
    quantity = Convert-ToNumber $sheet.Cells.Item($targetRow, 11).Text
    unit = [string]$sheet.Cells.Item($targetRow, 12).Text
    instockDate = $instockDateValue.ToString("yyyy-MM-dd")
    scNo = [string]$sheet.Cells.Item($targetRow, 14).Text
    orderSelectMacro = $ranOrderSelectMacro
    updateMacro = $ranUpdateMacro
    verifiedRow = $verifiedRow
    verifiedStatus = $verifiedStatus
    verifiedInstockDate = $verifiedInstock
    warning = $offerMacroWarning
    before = @{
      status = $beforeStatus
      instockDate = $beforeInstock
    }
  }

  $result | ConvertTo-Json -Depth 5 -Compress
} finally {
  if ($workbook -and $openedWorkbook) {
    try { $workbook.Close($true) } catch {}
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
