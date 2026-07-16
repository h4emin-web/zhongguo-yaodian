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

  if ($targetRow -le 0) {
    throw "PO '$PoNo' was not found in offer list."
  }

  $beforeStatus = [string]$sheet.Cells.Item($targetRow, 3).Text
  $beforeInstock = [string]$sheet.Cells.Item($targetRow, 25).Text

  if ($Commit) {
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 3).Value2 = $statusInStock } "set status"
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 25).Value2 = $instockDateValue.ToString("yyyy-MM-dd") } "set instock"
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 25).NumberFormat = "yyyy-mm-dd" } "instock format"

    try {
      Invoke-WithComRetry { $sheet.Range("A$targetRow").Select() | Out-Null } "select offer row"
      $ranOrderSelectMacro = Invoke-FirstMacro $excel $workbook $sheet @($orderSelectMacroName, "OrderSelect", "SelectOrder") "offer order select"
      $ranUpdateMacro = Invoke-FirstMacro $excel $workbook $sheet @($updateMacroName, "Update", "Modify") "offer update"
    } catch {
      $offerMacroWarning = "Offer list server update macro failed: $($_.Exception.Message)"
      throw $offerMacroWarning
    }

    Invoke-WithComRetry { $workbook.Save() } "save offer workbook"
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
