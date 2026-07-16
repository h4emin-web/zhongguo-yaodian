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

function Resolve-WorkbookPath {
  param([string]$RequestedPath)
  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) {
    return (Resolve-Path -LiteralPath $RequestedPath).Path
  }
  if ($env:OFFER_LIST_PATH -and (Test-Path -LiteralPath $env:OFFER_LIST_PATH)) {
    return (Resolve-Path -LiteralPath $env:OFFER_LIST_PATH).Path
  }

  $desktopPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "오퍼발행내역C8-2(양식수정).xlsm"
  if (Test-Path -LiteralPath $desktopPath) {
    return (Resolve-Path -LiteralPath $desktopPath).Path
  }

  $defaultPath = "Z:\동진파마\공유문서(Main)\A60-오퍼리스트\오퍼발행내역C8-2(양식수정).xlsm"
  if (Test-Path -LiteralPath $defaultPath) {
    return (Resolve-Path -LiteralPath $defaultPath).Path
  }

  throw "오퍼발행내역 파일을 찾지 못했습니다. OFFER_LIST_PATH를 설정해주세요."
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
if (-not $declarationText) { throw "수입신고번호가 비어 있습니다." }

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

  $sheet = Invoke-WithComRetry { $workbook.Worksheets.Item("PO메인") } "open PO sheet"
  Invoke-WithComRetry { $sheet.Activate() | Out-Null } "activate PO sheet"
  Invoke-WithComRetry { $sheet.Range("B1").Value2 = $normalizedPoNo } "enter PO search"
  try { Run-WorkbookMacro $excel $workbook "PO복사" } catch {}

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
  if ($targetRow -le 0) { throw "오퍼발행내역에서 PO '$PoNo' 를 찾지 못했습니다." }

  $beforeDeclarationNo = [string]$sheet.Cells.Item($targetRow, 34).Text

  if ($Commit) {
    Invoke-WithComRetry { $sheet.Cells.Item($targetRow, 34).Value2 = $declarationText } "set declaration no"
    Invoke-WithComRetry { $sheet.Range("A$targetRow").Select() | Out-Null } "select offer row"
    $ranOrderSelectMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @("오더선택", "OrderSelect", "SelectOrder") "A" "offer order select"
    Invoke-WithComRetry { $sheet.Range("K$targetRow").Select() | Out-Null } "select offer update cell"
    $ranUpdateMacro = Invoke-MacroOrColumnButton $excel $workbook $sheet @("수정", "Update", "Modify") "K" "offer update"
    Invoke-WithComRetry { $workbook.Save() } "save offer workbook"
  }

  [ordered]@{
    ok = $true
    committed = [bool]$Commit
    workbookPath = $resolvedPath
    sheetName = "PO메인"
    row = $targetRow
    poNo = [string]$sheet.Cells.Item($targetRow, 1).Text
    declarationNo = $declarationText
    beforeDeclarationNo = $beforeDeclarationNo
    orderSelectMacro = $ranOrderSelectMacro
    updateMacro = $ranUpdateMacro
  } | ConvertTo-Json -Depth 5 -Compress
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
