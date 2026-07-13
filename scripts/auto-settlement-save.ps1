param(
  [Parameter(Mandatory = $true)][string]$SettlementPath,
  [Parameter(Mandatory = $true)][string]$ManagerName,
  [Parameter(Mandatory = $true)][string]$PoNo,
  [Parameter(Mandatory = $true)][string]$BoardingDate,
  [Parameter(Mandatory = $true)][string]$InstockDate,
  [Parameter(Mandatory = $true)][string]$ExchangeRate,
  [string]$Quantity = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
trap {
  Write-Error ("LINE {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
  exit 1
}

$DefaultTestTargetPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "3.수입정산서C3-1(26년)-박성윤,송하형,강해민 - 복사본.xlsm"
$TargetPath = if ($env:AUTO_SETTLEMENT_TARGET_PATH) {
  $env:AUTO_SETTLEMENT_TARGET_PATH
} elseif ($env:AUTO_SETTLEMENT_REAL_TARGET -eq "1") {
  ""
} else {
  $DefaultTestTargetPath
}
$TargetDir = if ($env:AUTO_SETTLEMENT_TARGET_DIR) {
  $env:AUTO_SETTLEMENT_TARGET_DIR
} elseif ($env:AUTO_SETTLEMENT_REAL_TARGET -eq "1") {
  "Z:\동진파마\공유문서(Main)\A61-OFFER자료(강동연)\수입정산서\수입정산서-마케팅지원팀\2026년\1. 수입정산서 신규파일"
} else {
  ""
}

function Convert-ToNumber($value) {
  if ($null -eq $value -or [string]$value -eq "") {
    return 0
  }

  $text = [string]$value
  $text = $text -replace "[^\d\.\-]", ""

  if (-not $text) {
    return 0
  }

  return [double]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-CellNumber($sheet, [string]$address) {
  return (Convert-ToNumber ($sheet.Range($address).Value2))
}

function Get-CellText($sheet, [string]$address) {
  $value = $sheet.Range($address).Text

  if ($null -eq $value -or [string]$value -eq "") {
    $value = $sheet.Range($address).Value2
  }

  return [string]$value
}

function Get-LastRegexNumber([string]$text, [string]$label) {
  $matches = [regex]::Matches($text, "$label\s*([\d,]+(?:\.\d+)?)")

  if ($matches.Count -eq 0) {
    return 0
  }

  return (Convert-ToNumber ($matches[$matches.Count - 1].Groups[1].Value))
}

function Sum-LeftLabels($sheet, [string[]]$labels) {
  $amount = 0
  $vat = 0

  for ($row = 1; $row -le 60; $row++) {
    $label = ([string]$sheet.Cells.Item($row, 2).Text).Trim() -replace "\s+", " "

    if ($labels -contains $label) {
      $amount += Get-CellNumber $sheet "D$row"
      $vat += Get-CellNumber $sheet "G$row"
    }
  }

  return @{ Amount = $amount; Vat = $vat }
}

function Sum-RightLabels($sheet, [string[]]$labels) {
  $amount = 0
  $vat = 0

  for ($row = 1; $row -le 60; $row++) {
    $label = ([string]$sheet.Cells.Item($row, 12).Text).Trim() -replace "\s+", " "

    if ($labels -contains $label) {
      $amount += Get-CellNumber $sheet "M$row"

      $nextLabel = ([string]$sheet.Cells.Item($row + 1, 12).Text).Trim() -replace "\s+", " "
      if ($nextLabel -eq "부가세") {
        $vat += Get-CellNumber $sheet "M$($row + 1)"
      }
    }
  }

  return @{ Amount = $amount; Vat = $vat }
}

function Find-LastBlockStart($sheet) {
  $used = $sheet.UsedRange
  $lastRow = $used.Row + $used.Rows.Count - 1
  $startRow = 0

  for ($row = 1; $row -le $lastRow; $row++) {
    $label = ([string]$sheet.Cells.Item($row, 1).Text).Trim()

    if ($label -eq "오퍼번호") {
      $startRow = $row
    }
  }

  if ($startRow -le 0) {
    throw "새 양식 시작 행을 찾지 못했습니다."
  }

  return $startRow
}

function Set-DateCell($sheet, [string]$address, [datetime]$date) {
  $cell = $sheet.Range($address)
  $cell.Value = $date
  $cell.NumberFormat = "yyyy-mm-dd"
}

function Set-NumberCell($sheet, [int]$row, [int]$column, $value) {
  $number = Convert-ToNumber $value
  $sheet.Cells.Item($row, $column).Formula = $number.ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

if (-not (Test-Path -LiteralPath $SettlementPath)) {
  throw "정산서 파일을 찾을 수 없습니다: $SettlementPath"
}

$targetFile = if ($TargetPath) {
  if (-not (Test-Path -LiteralPath $TargetPath)) {
    throw "테스트용 수입정산서 복사본을 찾지 못했습니다: $TargetPath"
  }

  Get-Item -LiteralPath $TargetPath
} else {
  Get-ChildItem -LiteralPath $TargetDir -File |
    Where-Object { $_.Extension -in ".xls", ".xlsx", ".xlsm" } |
    Where-Object { $_.Name -like "*$ManagerName*" -and $_.Name -notlike "※*" } |
    Select-Object -First 1
}

if (-not $targetFile) {
  throw "담당자 '$ManagerName' 이 포함된 수입정산서 파일을 찾지 못했습니다."
}

$boarding = [datetime]::Parse($BoardingDate, [System.Globalization.CultureInfo]::InvariantCulture)
$instock = [datetime]::Parse($InstockDate, [System.Globalization.CultureInfo]::InvariantCulture)
$sheetName = "$($instock.Month)월"
$exchange = Convert-ToNumber $ExchangeRate

if ($exchange -le 0) {
  throw "ERP 환율이 올바르지 않습니다."
}

$excel = $null
$sourceWorkbook = $null
$targetWorkbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 1

  $sourceWorkbook = $excel.Workbooks.Open($SettlementPath, $false, $true)
  $sourceSheet = $sourceWorkbook.Worksheets.Item(1)

  $declarationNo = Get-CellText $sourceSheet "C7"
  $taxText = Get-CellText $sourceSheet "J39"
  $duty = Get-LastRegexNumber $taxText "관세"
  $importVat = Get-LastRegexNumber $taxText "부가세"

  $port = Sum-LeftLabels $sourceSheet @("입항료", "콘테이너세", "터미날 취급료", "화물적임료", "선박/항공 운임", "셔틀료", "하역료")
  $transport = Sum-LeftLabels $sourceSheet @("시외 운송료", "시내 운송료", "철 도 운 송 료", "취급수수료 H/C")
  $warehouse = Sum-LeftLabels $sourceSheet @("보관료", "출고상차료")
  $insurance = Sum-LeftLabels $sourceSheet @("화 재 보 험 료")
  $quarantine = Sum-RightLabels $sourceSheet @("검역수수료", "검역교통비")
  $clearance = Sum-RightLabels $sourceSheet @("통관수수료", "부대수수료", "임개수수료", "타장수수료", "인지대", "복사대", "팩시밀리 대", "시외 전화료")

  $targetWorkbook = $excel.Workbooks.Open($targetFile.FullName)

  if ($targetWorkbook.ReadOnly) {
    throw "수입정산서 파일이 읽기 전용으로 열렸습니다. 다른 사용자가 열고 있는지 확인해주세요."
  }

  $targetSheet = $targetWorkbook.Worksheets.Item($sheetName)
  $targetSheet.Activate() | Out-Null
  $excel.Run("'$($targetWorkbook.Name)'!양식") | Out-Null
  $startRow = Find-LastBlockStart $targetSheet

  $targetSheet.Cells.Item($startRow, 2).Value2 = $PoNo
  $targetSheet.Cells.Item($startRow, 2).Select() | Out-Null
  $excel.Run("'$($targetWorkbook.Name)'!오퍼내용") | Out-Null

  $totalForeign = Convert-ToNumber ($targetSheet.Cells.Item($startRow + 4, 2).Value2)
  $quantityValue = Convert-ToNumber $Quantity

  if ($quantityValue -le 0) {
    $quantityValue = Convert-ToNumber ($targetSheet.Cells.Item($startRow + 2, 2).Text)
  }

  Set-DateCell $targetSheet "F$($startRow + 2)" $boarding
  Set-DateCell $targetSheet "O$($startRow + 5)" $instock
  Set-NumberCell $targetSheet ($startRow + 5) 6 $exchange
  $targetSheet.Cells.Item($startRow + 1, 8).Formula = "=F$($startRow + 5)*B$($startRow + 4)"
  $targetSheet.Cells.Item($startRow + 1, 11).Formula = "=H$($startRow + 1)"

  Set-NumberCell $targetSheet ($startRow + 2) 12 $duty
  Set-NumberCell $targetSheet ($startRow + 2) 13 $importVat
  Set-NumberCell $targetSheet ($startRow + 3) 12 $port.Amount
  Set-NumberCell $targetSheet ($startRow + 3) 13 $port.Vat
  Set-NumberCell $targetSheet ($startRow + 4) 12 $transport.Amount
  Set-NumberCell $targetSheet ($startRow + 4) 13 $transport.Vat
  Set-NumberCell $targetSheet ($startRow + 6) 12 $insurance.Amount
  Set-NumberCell $targetSheet ($startRow + 6) 13 $insurance.Vat
  Set-NumberCell $targetSheet ($startRow + 7) 12 $warehouse.Amount
  Set-NumberCell $targetSheet ($startRow + 7) 13 $warehouse.Vat
  Set-NumberCell $targetSheet ($startRow + 9) 12 $quarantine.Amount
  Set-NumberCell $targetSheet ($startRow + 9) 13 $quarantine.Vat
  Set-NumberCell $targetSheet ($startRow + 10) 12 $clearance.Amount
  Set-NumberCell $targetSheet ($startRow + 10) 13 $clearance.Vat

  if ($declarationNo) {
    $targetSheet.Cells.Item($startRow + 1, 14).Value2 = $declarationNo
  }

  $targetSheet.Cells.Item($startRow + 11, 14).Value2 = $ManagerName
  $targetSheet.Cells.Item($startRow + 13, 14).Value2 = "정산완료"
  $targetSheet.Cells.Item($startRow + 13, 15).Value2 = "O"

  $excel.CalculateFull() | Out-Null

  $unitPriceFormula = '=ROUND(H{0}/VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B{1})),"kg",""),",","")),0)' -f ($startRow + 13), ($startRow + 2)
  $targetSheet.Cells.Item($startRow + 4, 6).Formula = $unitPriceFormula
  $targetSheet.Cells.Item($startRow + 4, 6).NumberFormat = "#,##0"
  $excel.CalculateFull() | Out-Null

  $targetWorkbook.Save()

  $unitPrice = [math]::Round((Convert-ToNumber ($targetSheet.Cells.Item($startRow + 4, 6).Value2)), 0)
  $purchaseUnitPrice = if ($quantityValue -gt 0) {
    [math]::Round(($totalForeign / $quantityValue), 4)
  } else {
    0
  }
  $krwAmount = [math]::Round((Convert-ToNumber ($targetSheet.Cells.Item($startRow + 13, 8).Value2)), 0)

  $result = [ordered]@{
    targetFile = $targetFile.Name
    targetPath = $targetFile.FullName
    sheetName = $sheetName
    startRow = $startRow
    poNo = $PoNo
    productCode = [string]$targetSheet.Cells.Item($startRow + 1, 6).Text
    quantity = $quantityValue
    exchangeRate = $exchange
    unitPrice = $unitPrice
    purchaseUnitPrice = $purchaseUnitPrice
    foreignAmount = $totalForeign
    krwAmount = $krwAmount
    remittanceAmount = [math]::Round($exchange * $totalForeign, 2)
  }

  $result | ConvertTo-Json -Compress
} finally {
  if ($sourceWorkbook) {
    $sourceWorkbook.Close($false)
  }
  if ($targetWorkbook) {
    $targetWorkbook.Close($true)
  }
  if ($excel) {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
}
