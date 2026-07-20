param(
  [Parameter(Mandatory=$true)][string]$ProductCode,
  [Parameter(Mandatory=$true)][double]$Quantity,
  [Parameter(Mandatory=$true)][string]$PoDate,
  [Parameter(Mandatory=$true)][string]$InstockDate,
  [string]$PoNo = "",
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
$xlShiftDown = -4121
$xlCalculationAutomatic = -4105

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
    [int]$Attempts = 120,
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

function Normalize-Code {
  param($Value)
  return ([string]$Value).Trim().ToUpperInvariant()
}

function Convert-ToNumber {
  param($Value)
  $match = [regex]::Match(([string]$Value), "-?\d+(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?")
  if ($match.Success) {
    return [double]::Parse($match.Value.Replace(",", ""), [Globalization.CultureInfo]::InvariantCulture)
  }
  return 0
}

function Format-NumberText {
  param([double]$Value)
  return $Value.ToString("0.##########", [Globalization.CultureInfo]::InvariantCulture)
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
  if ($env:INOUT_ORDER_PATH -and (Test-Path -LiteralPath $env:INOUT_ORDER_PATH)) {
    return (Resolve-Path -LiteralPath $env:INOUT_ORDER_PATH).Path
  }

  $desktopName = "{0}C7-v04(z{1}).xlsm" -f (U 0xC785,0xCD9C,0xACE0,0xC9C0,0xC2DC,0xC11C), (U 0xAD6C,0xBD84)
  $desktopPath = Join-Path ([Environment]::GetFolderPath("Desktop")) $desktopName
  if (Test-Path -LiteralPath $desktopPath) {
    return (Resolve-Path -LiteralPath $desktopPath).Path
  }

  throw "In/out order workbook was not found. Set INOUT_ORDER_PATH."
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

function Find-HeaderColumn {
  param($Sheet, [string[]]$Needles, [int]$MaxRows = 8, [int]$MaxColumns = 80)

  for ($row = 1; $row -le $MaxRows; $row += 1) {
    for ($column = 1; $column -le $MaxColumns; $column += 1) {
      $text = ([string]$Sheet.Cells.Item($row, $column).Text).Trim() -replace "\s+", ""
      if (-not $text) { continue }

      foreach ($needle in $Needles) {
        $normalizedNeedle = $needle -replace "\s+", ""
        if ($text -eq $normalizedNeedle -or $text -like "*$normalizedNeedle*") {
          return $column
        }
      }
    }
  }

  return 0
}

function Find-InoutOrderRow {
  param($Sheet, [string]$OrderNo, [string]$ProductCode, [string]$PoNo, [int]$RemarkColumn)

  $lastRow = Invoke-WithComRetry { $Sheet.Range("A500000").End($xlUp).Row } "last row for in/out verification"
  $normalizedCode = Normalize-Code $ProductCode
  $normalizedOrderNo = ([string]$OrderNo).Trim()
  $normalizedPoNo = ([string]$PoNo).Trim()

  for ($row = 3; $row -le $lastRow; $row += 1) {
    $rowOrderNo = ([string]$Sheet.Cells.Item($row, 1).Text).Trim()
    $rowCode = Normalize-Code $Sheet.Cells.Item($row, 5).Text

    if ($normalizedOrderNo -and $rowOrderNo -eq $normalizedOrderNo) {
      return $row
    }

    if ($normalizedCode -and $rowCode -eq $normalizedCode -and $RemarkColumn -gt 0) {
      $rowRemark = ([string]$Sheet.Cells.Item($row, $RemarkColumn).Text).Trim()
      if ($normalizedPoNo -and $rowRemark -eq $normalizedPoNo) {
        return $row
      }
    }
  }

  return 0
}

$resolvedPath = Resolve-WorkbookPath $WorkbookPath
$poDateValue = Convert-ToDateValue $PoDate
$instockDateValue = Convert-ToDateValue $InstockDate

if (-not $poDateValue) { throw "Invalid PO date: $PoDate" }
if (-not $instockDateValue) { throw "Invalid instock date: $InstockDate" }
if (-not $ProductCode.Trim()) { throw "Product code is empty." }
if ($Quantity -le 0) { throw "Quantity must be greater than 0." }

$sheetName = U 0xCD9C,0xACE0,0xC9C0,0xC2DC
$inText = U 0xC785
$searchMacroName = "{0}_{1}" -f (U 0xC870,0xD68C), (U 0xBA54,0xC774,0xCEE4)
$orderSelectMacroName = if ($env:PENDING_RECEIPT_INOUT_ORDER_SELECT_MACRO) { $env:PENDING_RECEIPT_INOUT_ORDER_SELECT_MACRO } else { "{0}{1}" -f (U 0xC624,0xB354), (U 0xC120,0xD0DD) }
$orderNoMacroName = if ($env:PENDING_RECEIPT_INOUT_ORDERNO_MACRO) { $env:PENDING_RECEIPT_INOUT_ORDERNO_MACRO } else { "{0}{1}" -f (U 0xC624,0xB354,0xBC88,0xD638), (U 0xB123,0xAE30) }
$addMacroName = if ($env:PENDING_RECEIPT_INOUT_ADD_MACRO) { $env:PENDING_RECEIPT_INOUT_ADD_MACRO } else { U 0xCD94,0xAC00 }
$macroWarning = ""
$quantityText = Format-NumberText $Quantity
$remarkText = $PoNo.Trim()
$ranOrderSelectMacro = ""
$ranOrderNoMacro = ""
$ranAddMacro = ""
$clearedDeliveryColumns = @()
$remarkColumn = 0
$savedOrderNo = ""
$verifiedRow = 0

$excel = Get-RunningExcel
$createdExcel = $false
$openedWorkbook = $false

if (-not $excel) {
  $excel = New-Object -ComObject Excel.Application
  $createdExcel = $true
}

$workbook = $null
$oldDisplayAlerts = $null
$oldScreenUpdating = $null
$oldEnableEvents = $null
$oldCalculation = $null

try {
  Invoke-WithComRetry { $excel.Visible = $true } "show Excel"
  $oldDisplayAlerts = Invoke-WithComRetry { $excel.DisplayAlerts } "read DisplayAlerts"
  $oldScreenUpdating = Invoke-WithComRetry { $excel.ScreenUpdating } "read ScreenUpdating"
  $oldEnableEvents = Invoke-WithComRetry { $excel.EnableEvents } "read EnableEvents"
  $oldCalculation = Invoke-WithComRetry { $excel.Calculation } "read Calculation"
  Invoke-WithComRetry { $excel.DisplayAlerts = $false } "DisplayAlerts false"
  Invoke-WithComRetry { $excel.ScreenUpdating = $true } "ScreenUpdating true"
  Invoke-WithComRetry { $excel.EnableEvents = $true } "EnableEvents true"
  Invoke-WithComRetry { $excel.Calculation = $xlCalculationAutomatic } "Calculation automatic"

  $workbook = Invoke-WithComRetry { Find-Workbook $excel $resolvedPath } "find in/out workbook"
  if (-not $workbook) {
    $workbook = Invoke-WithComRetry { $excel.Workbooks.Open($resolvedPath) } "open in/out workbook"
    $openedWorkbook = $true
  }

  $sheet = Invoke-WithComRetry { $workbook.Worksheets.Item($sheetName) } "open in/out sheet"
  Invoke-WithComRetry { $sheet.Activate() | Out-Null } "activate in/out sheet"
  Invoke-WithComRetry { $sheet.Range("B1").Value2 = $ProductCode } "enter product search"
  try { Run-WorkbookMacro $excel $workbook $searchMacroName } catch {}

  $lastRow = Invoke-WithComRetry { $sheet.Range("A500000").End($xlUp).Row } "last row"
  $normalizedCode = Normalize-Code $ProductCode
  $matches = @()

  for ($row = 3; $row -le $lastRow; $row += 1) {
    $rowCode = Normalize-Code $sheet.Cells.Item($row, 5).Text
    if ($rowCode -ne $normalizedCode) { continue }

    $rowDirection = ([string]$sheet.Cells.Item($row, 3).Text).Trim()
    if ($rowDirection -ne $inText) { continue }

    $rowQuantity = Convert-ToNumber $sheet.Cells.Item($row, 9).Text
    $rowDate = Convert-ToDateValue $sheet.Cells.Item($row, 2).Value2

    $matches += [pscustomobject]@{
      Row = $row
      OrderNo = [string]$sheet.Cells.Item($row, 1).Text
      ProductName = [string]$sheet.Cells.Item($row, 6).Text
      Quantity = $rowQuantity
      DateScore = if ($rowDate) { $rowDate.Ticks } else { 0 }
    }
  }

  if (-not $matches -or $matches.Count -eq 0) {
    throw "Existing incoming row was not found for product code $ProductCode."
  }

  $source = $matches | Sort-Object @{ Expression = "DateScore"; Descending = $true }, @{ Expression = "Row"; Descending = $true } | Select-Object -First 1
  $sourceRow = [int]$source.Row
  $newRow = $sourceRow + 1

  if ($Commit) {
    Invoke-WithComRetry { $sheet.Rows.Item($newRow).Insert($xlShiftDown) | Out-Null } "insert new row"
    Invoke-WithComRetry { $sheet.Rows.Item($sourceRow).Copy($sheet.Rows.Item($newRow)) | Out-Null } "copy source row"

    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 1).ClearContents() } "clear order no"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 2).Value2 = $poDateValue.ToString("yyyy-MM-dd") } "set order date"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 2).NumberFormat = "yyyy-mm-dd" } "format order date"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 3).Value2 = $inText } "set direction"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 5).Value2 = $ProductCode } "set product code"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 9).Value2 = $quantityText } "set quantity"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 10).ClearContents() } "clear unit price"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 11).ClearContents() } "clear amount"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 12).Value2 = $instockDateValue.ToString("yyyy-MM-dd") } "set due date"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 12).NumberFormat = "yyyy-mm-dd" } "format due date"

    $remarkColumn = Find-HeaderColumn $sheet @((U 0xBE44,0xACE0))
    if ($remarkColumn -gt 0) {
      Invoke-WithComRetry { $sheet.Cells.Item($newRow, $remarkColumn).ClearContents() } "clear remarks"
      if ($remarkText) {
        Invoke-WithComRetry { $sheet.Cells.Item($newRow, $remarkColumn).Value2 = $remarkText } "set remarks"
      }
    }

    $deliveryStatusColumn = Find-HeaderColumn $sheet @((U 0xBC30,0xC1A1,0xC5EC,0xBD80), (U 0xBC30,0xC1A1,0xC5EC,0xBD80,0x004F))
    $deliveryDoneColumn = Find-HeaderColumn $sheet @((U 0xBC30,0xC1A1,0xC644,0xB8CC,0xC77C), (U 0xBC30,0xC1A1,0xC644,0xB8CC))

    if ($deliveryStatusColumn -gt 0) {
      Invoke-WithComRetry { $sheet.Cells.Item($newRow, $deliveryStatusColumn).ClearContents() } "clear delivery status"
      $clearedDeliveryColumns += $deliveryStatusColumn
    }
    if ($deliveryDoneColumn -gt 0) {
      Invoke-WithComRetry { $sheet.Cells.Item($newRow, $deliveryDoneColumn).ClearContents() } "clear delivery complete date"
      $clearedDeliveryColumns += $deliveryDoneColumn
    }

    try {
      Invoke-WithComRetry { $sheet.Range("A$newRow").Select() | Out-Null } "select new row"
      $ranOrderSelectMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($orderSelectMacroName, "출고수정선택", "OrderSelect", "SelectOrder") "A" "in/out order select"
      Invoke-WithComRetry { $sheet.Range("A$newRow").Select() | Out-Null } "select new row after order select"
      $ranOrderNoMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($orderNoMacroName, "오더번호", "OrderNo", "OrderNoInsert", "InsertOrderNo") "B" "in/out order no"
      $savedOrderNo = ([string]$sheet.Cells.Item($newRow, 1).Text).Trim()
      if (-not $savedOrderNo) {
        throw "in/out order no button did not create an order number in column A."
      }
      Invoke-WithComRetry { $sheet.Range("G$newRow").Select() | Out-Null } "select add cell"
      $ranAddMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @($addMacroName, "출고삽입", "추가선택", "Add", "Insert", "Append") "G" "in/out add"
    } catch {
      $macroWarning = "In/out order select/order no/add macro failed: $($_.Exception.Message)"
      throw $macroWarning
    }

    Invoke-WithComRetry { $workbook.Save() } "save in/out workbook"

    Invoke-WithComRetry { $sheet.Range("B1").Value2 = $ProductCode } "re-enter product search"
    Run-WorkbookMacro $excel $workbook $searchMacroName
    Start-Sleep -Milliseconds 1200
    $verifiedRow = Find-InoutOrderRow $sheet $savedOrderNo $ProductCode $remarkText $remarkColumn
    if ($verifiedRow -le 0) {
      throw "In/out order save verification failed: order no '$savedOrderNo' was not found after add."
    }
  }

  $result = [ordered]@{
    ok = $true
    committed = [bool]$Commit
    workbookPath = $resolvedPath
    sheetName = $sheetName
    sourceRow = $sourceRow
    newRow = $newRow
    sourceOrderNo = $source.OrderNo
    productCode = $ProductCode
    productName = $source.ProductName
    quantity = $quantityText
    poNo = $remarkText
    savedOrderNo = $savedOrderNo
    searchedProductCode = $ProductCode
    searchedDirection = $inText
    poDate = $poDateValue.ToString("yyyy-MM-dd")
    dueDate = $instockDateValue.ToString("yyyy-MM-dd")
    unitPriceCleared = $true
    amountCleared = $true
    remarkColumn = $remarkColumn
    deliveryColumnsCleared = $clearedDeliveryColumns
    orderSelectMacro = $ranOrderSelectMacro
    orderNoMacro = $ranOrderNoMacro
    addMacro = $ranAddMacro
    verifiedRow = $verifiedRow
    warning = $macroWarning
  }

  $result | ConvertTo-Json -Depth 5 -Compress
} finally {
  if ($excel) {
    try { if ($null -ne $oldCalculation) { $excel.Calculation = $oldCalculation } } catch {}
    try { if ($null -ne $oldEnableEvents) { $excel.EnableEvents = $oldEnableEvents } } catch {}
    try { if ($null -ne $oldScreenUpdating) { $excel.ScreenUpdating = $oldScreenUpdating } } catch {}
    try { if ($null -ne $oldDisplayAlerts) { $excel.DisplayAlerts = $oldDisplayAlerts } } catch {}
  }

  if ($workbook -and $openedWorkbook) {
    try { $workbook.Close($false) } catch {}
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
