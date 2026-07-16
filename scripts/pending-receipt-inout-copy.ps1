param(
  [Parameter(Mandatory=$true)][string]$ProductCode,
  [Parameter(Mandatory=$true)][double]$Quantity,
  [Parameter(Mandatory=$true)][string]$PoDate,
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
$xlShiftDown = -4121
$xlCalculationAutomatic = -4105

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
$addMacroName = if ($env:PENDING_RECEIPT_INOUT_ADD_MACRO) { $env:PENDING_RECEIPT_INOUT_ADD_MACRO } else { U 0xCD9C,0xACE0,0xCD94,0xAC00 }
$macroWarning = ""

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
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 9).Value2 = $Quantity } "set quantity"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 10).ClearContents() } "clear unit price"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 11).ClearContents() } "clear amount"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 12).Value2 = $instockDateValue.ToString("yyyy-MM-dd") } "set due date"
    Invoke-WithComRetry { $sheet.Cells.Item($newRow, 12).NumberFormat = "yyyy-mm-dd" } "format due date"

    try {
      Invoke-WithComRetry { $sheet.Range("A$newRow").Select() | Out-Null } "select new row"
      Run-WorkbookMacro $excel $workbook $addMacroName
    } catch {
      $macroWarning = "In/out add macro '$addMacroName' failed: $($_.Exception.Message)"
    }

    Invoke-WithComRetry { $workbook.Save() } "save in/out workbook"
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
    quantity = $Quantity
    poDate = $poDateValue.ToString("yyyy-MM-dd")
    dueDate = $instockDateValue.ToString("yyyy-MM-dd")
    unitPriceCleared = $true
    amountCleared = $true
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
