param(
  [Parameter(Mandatory=$true)][string]$ProductCode,
  [Parameter(Mandatory=$true)][double]$Quantity,
  [Parameter(Mandatory=$true)][string]$DueDate,
  [Parameter(Mandatory=$true)][double]$UnitPrice,
  [Parameter(Mandatory=$true)][double]$Amount,
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
$xlCalculationAutomatic = -4105

function Invoke-WithComRetry {
  param(
    [Parameter(Mandatory=$true)][scriptblock]$Operation,
    [string]$Label = "Excel 작업",
    [int]$Attempts = 90,
    [int]$DelayMs = 500
  )

  $lastError = $null

  for ($attempt = 1; $attempt -le $Attempts; $attempt += 1) {
    try {
      return & $Operation
    } catch {
      $lastError = $_
      $message = $_.Exception.Message
      $hresult = try { [uint32]$_.Exception.HResult } catch { 0 }
      $isBusy =
        $message -like "*Call was rejected by callee*" -or
        $message -like "*RPC_E_CALL_REJECTED*" -or
        $message -like "*0x80010001*" -or
        $message -like "*0x800AC472*" -or
        $message -like "*0x800A9C68*" -or
        $hresult -eq 0x80010001 -or
        $hresult -eq 0x800AC472 -or
        $hresult -eq 0x800A9C68

      if (-not $isBusy -or $attempt -eq $Attempts) {
        throw
      }

      Start-Sleep -Milliseconds $DelayMs
    }
  }

  if ($lastError) {
    throw "$Label 실패: $($lastError.Exception.Message)"
  }
}

function Convert-ToNumber {
  param($Value)

  if ($null -eq $Value) {
    return 0
  }

  $text = [string]$Value
  $match = [regex]::Match($text, "-?\d+(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?")

  if (-not $match.Success) {
    return 0
  }

  return [double]::Parse($match.Value.Replace(",", ""), [Globalization.CultureInfo]::InvariantCulture)
}

function Convert-ToDateValue {
  param($Value)

  if ($Value -is [datetime]) {
    return $Value.Date
  }

  if ($Value -is [double] -or $Value -is [int]) {
    try {
      return [datetime]::FromOADate([double]$Value).Date
    } catch {
      return $null
    }
  }

  $text = ([string]$Value).Trim()

  if (-not $text) {
    return $null
  }

  $formats = @("yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyyMMdd", "yy-MM-dd", "yy/MM/dd")
  $parsed = [datetime]::MinValue

  foreach ($format in $formats) {
    if ([datetime]::TryParseExact($text, $format, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
      return $parsed.Date
    }
  }

  if ([datetime]::TryParse($text, [ref]$parsed)) {
    return $parsed.Date
  }

  return $null
}

function Normalize-Code {
  param($Value)
  return ([string]$Value).Trim().ToUpperInvariant()
}

function Get-DesktopPath {
  return [Environment]::GetFolderPath("Desktop")
}

function Resolve-WorkbookPath {
  param([string]$RequestedPath)

  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) {
    return (Resolve-Path -LiteralPath $RequestedPath).Path
  }

  if ($env:INOUT_ORDER_PATH -and (Test-Path -LiteralPath $env:INOUT_ORDER_PATH)) {
    return (Resolve-Path -LiteralPath $env:INOUT_ORDER_PATH).Path
  }

  $desktopPath = Join-Path (Get-DesktopPath) "입출고지시서C7-v04(z구분).xlsm"

  if (Test-Path -LiteralPath $desktopPath) {
    return (Resolve-Path -LiteralPath $desktopPath).Path
  }

  throw "입출고 지시서 파일을 찾을 수 없습니다. INOUT_ORDER_PATH를 설정하거나 바탕화면에 입출고지시서C7-v04(z구분).xlsm 파일을 두세요."
}

function Get-RunningExcel {
  try {
    return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  } catch {
    return $null
  }
}

function Find-Workbook {
  param($Excel, [string]$Path)

  foreach ($workbook in @($Excel.Workbooks)) {
    try {
      if ($workbook.FullName -eq $Path) {
        return $workbook
      }
    } catch {
    }
  }

  return $null
}

function Run-WorkbookMacro {
  param($Excel, $Workbook, [string]$MacroName)
  $macro = "'{0}'!{1}" -f $Workbook.Name, $MacroName
  Invoke-WithComRetry { $Excel.Run($macro) | Out-Null } "매크로 실행: $MacroName"
}

function Find-OrderRow {
  param($Sheet, [string]$OrderNo)

  $lastRow = $Sheet.Range("A500000").End($xlUp).Row

  for ($row = 3; $row -le $lastRow; $row++) {
    if (([string]$Sheet.Cells.Item($row, 1).Text).Trim() -eq $OrderNo) {
      return $row
    }
  }

  return 0
}

function Start-EnterKeyHelper {
  param([string]$WorkbookName = "")

  $command = @"
Start-Sleep -Milliseconds 1200
`$shell = New-Object -ComObject WScript.Shell
`$titles = @("Microsoft Excel", "Excel", "$WorkbookName") | Where-Object { `$_ }
for (`$i = 0; `$i -lt 75; `$i++) {
  foreach (`$title in `$titles) {
    if (`$shell.AppActivate(`$title)) {
      Start-Sleep -Milliseconds 150
      `$shell.SendKeys("{ENTER}")
      break
    }
  }
  if (`$i -lt 10) {
    Start-Sleep -Milliseconds 250
  } else {
    Start-Sleep -Milliseconds 900
  }
}
"@
  $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
  Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList @("-NoProfile", "-EncodedCommand", $encoded) | Out-Null
}

$resolvedPath = Resolve-WorkbookPath $WorkbookPath
$dueDateValue = Convert-ToDateValue $DueDate

if (-not $dueDateValue) {
  throw "납기일 형식을 확인할 수 없습니다: $DueDate"
}

if (-not $ProductCode.Trim()) {
  throw "품목코드가 비어 있습니다."
}

if ($Quantity -le 0) {
  throw "수량은 0보다 커야 합니다."
}

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
  Invoke-WithComRetry { $excel.Visible = $true } "Excel 표시"
  $oldDisplayAlerts = Invoke-WithComRetry { $excel.DisplayAlerts } "DisplayAlerts 읽기"
  $oldScreenUpdating = Invoke-WithComRetry { $excel.ScreenUpdating } "ScreenUpdating 읽기"
  $oldEnableEvents = Invoke-WithComRetry { $excel.EnableEvents } "EnableEvents 읽기"
  $oldCalculation = Invoke-WithComRetry { $excel.Calculation } "Calculation 읽기"
  Invoke-WithComRetry { $excel.DisplayAlerts = $false } "DisplayAlerts 설정"
  # This workbook's save macros depend on the active selection/event flow.
  Invoke-WithComRetry { $excel.ScreenUpdating = $true } "ScreenUpdating 설정"
  Invoke-WithComRetry { $excel.EnableEvents = $true } "EnableEvents 설정"
  Invoke-WithComRetry { $excel.Calculation = $xlCalculationAutomatic } "Calculation 설정"

  $workbook = Invoke-WithComRetry { Find-Workbook $excel $resolvedPath } "입출고 통합문서 확인"

  if (-not $workbook) {
    $workbook = Invoke-WithComRetry { $excel.Workbooks.Open($resolvedPath) } "입출고 통합문서 열기"
    $openedWorkbook = $true
  }

  $sheet = Invoke-WithComRetry { $workbook.Worksheets.Item("출고지시") } "출고지시 시트 열기"
  Invoke-WithComRetry { $sheet.Activate() | Out-Null } "출고지시 시트 활성화"
  Invoke-WithComRetry { $sheet.Range("B1").Value2 = $ProductCode } "품목코드 검색어 입력"

  Run-WorkbookMacro $excel $workbook "조회_메이커"

  $lastRow = $sheet.Range("A500000").End($xlUp).Row
  $normalizedCode = Normalize-Code $ProductCode
  $matches = @()

  for ($row = 3; $row -le $lastRow; $row++) {
    $rowCode = Normalize-Code $sheet.Cells.Item($row, 5).Text

    if ($rowCode -ne $normalizedCode) {
      continue
    }

    $rowQuantity = Convert-ToNumber $sheet.Cells.Item($row, 9).Text

    if ([math]::Abs($rowQuantity - $Quantity) -gt 0.0001) {
      continue
    }

    $rowDirection = ([string]$sheet.Cells.Item($row, 3).Text).Trim()
    $rowDate = Convert-ToDateValue $sheet.Cells.Item($row, 2).Value2
    $dateScore = if ($rowDate) { $rowDate.Ticks } else { 0 }

    $matches += [pscustomobject]@{
      Row = $row
      OrderNo = [string]$sheet.Cells.Item($row, 1).Text
      Direction = $rowDirection
      ProductCode = [string]$sheet.Cells.Item($row, 5).Text
      ProductName = [string]$sheet.Cells.Item($row, 6).Text
      Quantity = $rowQuantity
      OrderDate = if ($rowDate) { $rowDate.ToString("yyyy-MM-dd") } else { "" }
      DateScore = $dateScore
      IsInbound = if ($rowDirection -eq "입") { 1 } else { 0 }
    }
  }

  if (-not $matches -or $matches.Count -eq 0) {
    throw "출고지시서에서 품목코드 $ProductCode / 수량 $Quantity 행을 찾지 못했습니다."
  }

  $target = $matches |
    Sort-Object @{ Expression = "IsInbound"; Descending = $true }, @{ Expression = "DateScore"; Descending = $true }, @{ Expression = "Row"; Descending = $true } |
    Select-Object -First 1

  $rowIndex = [int]$target.Row
  $orderNo = [string]$target.OrderNo
  $before = [ordered]@{
    dueDate = [string]$sheet.Cells.Item($rowIndex, 12).Text
    unitPrice = [string]$sheet.Cells.Item($rowIndex, 10).Text
    amount = [string]$sheet.Cells.Item($rowIndex, 11).Text
  }

  if ($Commit) {
    Invoke-WithComRetry {
      $sheet.Cells.Item($rowIndex, 12).Value2 = $dueDateValue.ToString("yyyy-MM-dd")
      $sheet.Cells.Item($rowIndex, 12).NumberFormat = "yyyy-mm-dd"
      $sheet.Cells.Item($rowIndex, 10).Formula = ([math]::Round($UnitPrice, 0)).ToString([Globalization.CultureInfo]::InvariantCulture)
      $sheet.Cells.Item($rowIndex, 10).NumberFormat = "#,##0"
      $sheet.Cells.Item($rowIndex, 11).Formula = ([math]::Round($Amount, 0)).ToString([Globalization.CultureInfo]::InvariantCulture)
      $sheet.Cells.Item($rowIndex, 11).NumberFormat = "#,##0"
    } "입출고 지시서 값 입력"

    Invoke-WithComRetry { $sheet.Range("A$rowIndex").Select() | Out-Null } "입출고 오더선택 셀 선택"
    Run-WorkbookMacro $excel $workbook "출고수정선택"
    Start-EnterKeyHelper $workbook.Name
    Run-WorkbookMacro $excel $workbook "출고수정"

    Start-Sleep -Milliseconds 1800
    Invoke-WithComRetry {
      $sheet.Activate() | Out-Null
      $sheet.Range("B1").Value2 = $ProductCode
    } "입출고 저장 후 품목코드 재검색어 입력" 120 700
    Run-WorkbookMacro $excel $workbook "조회_메이커"

    $verifiedRow = Invoke-WithComRetry { Find-OrderRow $sheet $orderNo } "입출고 저장 후 오더번호 재조회" 90 600

    if ($verifiedRow -eq 0) {
      throw "입출고 지시서 저장 후 오더번호 $orderNo 재조회에 실패했습니다."
    }

    $verifiedDueDate = Invoke-WithComRetry { Convert-ToDateValue $sheet.Cells.Item($verifiedRow, 12).Text } "입출고 저장 후 납기일 확인"
    $verifiedUnitPrice = Invoke-WithComRetry { [math]::Round((Convert-ToNumber $sheet.Cells.Item($verifiedRow, 10).Text), 0) } "입출고 저장 후 단가 확인"
    $verifiedAmount = Invoke-WithComRetry { [math]::Round((Convert-ToNumber $sheet.Cells.Item($verifiedRow, 11).Text), 0) } "입출고 저장 후 금액 확인"
    $expectedUnitPrice = [math]::Round($UnitPrice, 0)
    $expectedAmount = [math]::Round($Amount, 0)

    if (-not $verifiedDueDate -or $verifiedDueDate.ToString("yyyy-MM-dd") -ne $dueDateValue.ToString("yyyy-MM-dd") -or $verifiedUnitPrice -ne $expectedUnitPrice -or $verifiedAmount -ne $expectedAmount) {
      throw ("입출고 지시서 저장 후 값 확인에 실패했습니다. 오더번호={0}, 확인값=납기일 {1}, 단가 {2}, 금액 {3}" -f $orderNo, $(if ($verifiedDueDate) { $verifiedDueDate.ToString("yyyy-MM-dd") } else { "" }), $verifiedUnitPrice, $verifiedAmount)
    }

    $rowIndex = $verifiedRow
  }

  $result = [ordered]@{
    workbookPath = $resolvedPath
    sheetName = "출고지시"
    row = $rowIndex
    committed = [bool]$Commit
    orderNo = $orderNo
    productCode = $ProductCode
    productName = $target.ProductName
    quantity = $Quantity
    dueDate = $dueDateValue.ToString("yyyy-MM-dd")
    unitPrice = [math]::Round($UnitPrice, 0)
    amount = [math]::Round($Amount, 0)
    before = $before
  }

  $resultJson = $result | ConvertTo-Json -Depth 5 -Compress
  $debugPath = Join-Path (Join-Path (Get-Location) "tmp") "inout-order-save-result.json"
  [System.IO.Directory]::CreateDirectory((Split-Path -Parent $debugPath)) | Out-Null
  [System.IO.File]::WriteAllText($debugPath, $resultJson, [System.Text.Encoding]::UTF8)
  $resultJson
} finally {
  if ($workbook -and $openedWorkbook) {
    try { $workbook.Close($false) } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
  }

  if ($excel -and $createdExcel) {
    try { if ($null -ne $oldCalculation) { $excel.Calculation = $oldCalculation } } catch {}
    try { $excel.EnableEvents = $true } catch {}
    try { $excel.ScreenUpdating = $true } catch {}
    try { if ($null -ne $oldDisplayAlerts) { $excel.DisplayAlerts = $oldDisplayAlerts } } catch {}
    try { $excel.Quit() } catch {}
  }

  if ($excel) {
    if (-not $createdExcel) {
      try { if ($null -ne $oldCalculation) { $excel.Calculation = $oldCalculation } } catch {}
      try { $excel.EnableEvents = $true } catch {}
      try { $excel.ScreenUpdating = $true } catch {}
      try { if ($null -ne $oldDisplayAlerts) { $excel.DisplayAlerts = $oldDisplayAlerts } } catch {}
    }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }

  try {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  } catch {}
}
