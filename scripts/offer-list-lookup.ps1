param(
  [Parameter(Mandatory=$true)][string]$PoNo,
  [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

trap {
  Write-Error ("LINE {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
  exit 1
}

$xlUp = -4162

function Convert-ToDateText {
  param($Value)

  if ($Value -is [datetime]) {
    return $Value.ToString("yyyy-MM-dd")
  }

  if ($Value -is [double] -or $Value -is [int]) {
    try {
      return ([datetime]::FromOADate([double]$Value)).ToString("yyyy-MM-dd")
    } catch {
    }
  }

  $text = ([string]$Value).Trim()

  if (-not $text) {
    return ""
  }

  $match = $text.Replace("/", "-").Replace(".", "-") -match "(\d{4})-(\d{1,2})-(\d{1,2})"

  if ($match) {
    return "{0}-{1:D2}-{2:D2}" -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3]
  }

  $monthText = $text -match "(\d{4})[./-]?(\d{1,2})\s*월\s*(초|중순|말)?"

  if ($monthText) {
    $year = [int]$Matches[1]
    $month = [int]$Matches[2]
    $day = switch ($Matches[3]) {
      "초" { 5 }
      "중순" { 15 }
      "말" { [datetime]::DaysInMonth($year, $month) }
      default { 1 }
    }

    return "{0}-{1:D2}-{2:D2}" -f $year, $month, $day
  }

  return $text
}

function Normalize-PoNo {
  param($Value)

  return ([string]$Value).Normalize([Text.NormalizationForm]::FormKC).Replace("‐", "-").Replace("‑", "-").Replace("‒", "-").Replace("–", "-").Replace("—", "-").Replace("−", "-").Replace(" ", "").Replace("`t", "").Trim().ToUpperInvariant()
}

function Get-DesktopPath {
  return [Environment]::GetFolderPath("Desktop")
}

function Resolve-OfferWorkbookPath {
  param([string]$RequestedPath)

  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) {
    return (Resolve-Path -LiteralPath $RequestedPath).Path
  }

  if ($env:OFFER_LIST_PATH -and (Test-Path -LiteralPath $env:OFFER_LIST_PATH)) {
    return (Resolve-Path -LiteralPath $env:OFFER_LIST_PATH).Path
  }

  $desktopPath = Join-Path (Get-DesktopPath) "오퍼발행내역C8-2(양식수정).xlsm"

  if (Test-Path -LiteralPath $desktopPath) {
    return (Resolve-Path -LiteralPath $desktopPath).Path
  }

  $serverPath = "Z:\동진파마\공유문서(Main)\A60-오퍼리스트\오퍼발행내역C8-2(양식수정).xlsm"

  if (Test-Path -LiteralPath $serverPath) {
    return (Resolve-Path -LiteralPath $serverPath).Path
  }

  throw "오퍼발행내역 파일을 찾지 못했습니다."
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
  $Excel.Run($macro) | Out-Null
}

$resolvedPath = Resolve-OfferWorkbookPath $WorkbookPath
$normalizedPoNo = Normalize-PoNo $PoNo
$excel = Get-RunningExcel
$createdExcel = $false
$openedWorkbook = $false

if (-not $excel) {
  $excel = New-Object -ComObject Excel.Application
  $createdExcel = $true
}

$workbook = $null
$oldDisplayAlerts = $null

try {
  $excel.Visible = $true
  $oldDisplayAlerts = $excel.DisplayAlerts
  $excel.DisplayAlerts = $false

  $workbook = Find-Workbook $excel $resolvedPath

  if (-not $workbook) {
    $workbook = $excel.Workbooks.Open($resolvedPath)
    $openedWorkbook = $true
  }

  $sheet = $workbook.Worksheets.Item("PO메인")
  $sheet.Activate() | Out-Null
  $sheet.Range("B1").Value2 = $normalizedPoNo

  Run-WorkbookMacro $excel $workbook "PO복사"

  $lastRow = $sheet.Range("A500000").End($xlUp).Row
  $targetRow = 0

  for ($row = 3; $row -le $lastRow; $row++) {
    $candidates = @(
      $sheet.Cells.Item($row, 1).Text,
      $sheet.Cells.Item($row, 2).Text,
      $sheet.Cells.Item($row, 3).Text
    )

    if (@($candidates | Where-Object { (Normalize-PoNo $_) -eq $normalizedPoNo }).Count -gt 0) {
      $targetRow = $row
      break
    }
  }

  if ($targetRow -eq 0) {
    throw "오퍼발행내역에서 PO '$PoNo' 를 찾지 못했습니다."
  }

  $boarding = Convert-ToDateText $sheet.Cells.Item($targetRow, 24).Text
  $instock = Convert-ToDateText $sheet.Cells.Item($targetRow, 25).Text

  if (-not $boarding) {
    throw "PO '$PoNo' 의 Boarding 값이 비어 있습니다."
  }

  if (-not $instock) {
    throw "PO '$PoNo' 의 Instock 값이 비어 있습니다."
  }

  $result = [ordered]@{
    poNo = $PoNo
    boardingDate = $boarding
    instockDate = $instock
    offerListPath = $resolvedPath
    offerListRow = $targetRow
    productCode = [string]$sheet.Cells.Item($targetRow, 9).Text
    productName = [string]$sheet.Cells.Item($targetRow, 10).Text
    quantity = [string]$sheet.Cells.Item($targetRow, 11).Text
  }

  $result | ConvertTo-Json -Compress
} finally {
  if ($workbook -and $openedWorkbook) {
    try { $workbook.Close($false) } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
  }

  if ($excel -and $createdExcel) {
    try { if ($null -ne $oldDisplayAlerts) { $excel.DisplayAlerts = $oldDisplayAlerts } } catch {}
    try { $excel.Quit() } catch {}
  }

  if ($excel) {
    if (-not $createdExcel) {
      try { if ($null -ne $oldDisplayAlerts) { $excel.DisplayAlerts = $oldDisplayAlerts } } catch {}
    }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }

  try {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  } catch {}
}
