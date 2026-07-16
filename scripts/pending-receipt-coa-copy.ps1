param(
  [Parameter(Mandatory=$true)][string]$PoNo,
  [Parameter(Mandatory=$true)][string]$ProductName,
  [Parameter(Mandatory=$true)][string]$InstockDate,
  [string]$Quantity = "",
  [string]$Messrs = "",
  [string]$SourceRoot = "",
  [string]$TargetRoot = "",
  [switch]$Commit
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

trap {
  Write-Error ("LINE {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
  exit 1
}

function U {
  param([int[]]$Codes)
  return -join ($Codes | ForEach-Object { [char]$_ })
}

function Normalize-Text {
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
    Trim().
    ToUpperInvariant()
}

function Convert-ToDateValue {
  param($Value)
  if ($Value -is [datetime]) { return $Value.Date }
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

function Resolve-SourceRoot {
  param([string]$RequestedPath)
  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) { return (Resolve-Path -LiteralPath $RequestedPath).Path }
  if ($env:PENDING_RECEIPT_COA_SOURCE_ROOT -and (Test-Path -LiteralPath $env:PENDING_RECEIPT_COA_SOURCE_ROOT)) {
    return (Resolve-Path -LiteralPath $env:PENDING_RECEIPT_COA_SOURCE_ROOT).Path
  }

  $offerText = U 0xC624,0xD30C
  $yearOfferText = "2026{0} {1}" -f (U 0xB144), $offerText
  $defaultPath = "Z:\{0}\{1}(Main)\A61-OFFER{2}({3})\OFFER-JG\DJP_{4}\DJP_{5}\{6} {7} {8}" -f `
    (U 0xB3D9,0xC9C4,0xD30C,0xB9C8), `
    (U 0xACF5,0xC720,0xBB38,0xC11C), `
    (U 0xC790,0xB8CC), `
    (U 0xAC15,0xB3D9,0xC5F0), `
    (U 0xC624,0xD30C,0xBC1C,0xD589,0xB0B4,0xC5ED), `
    $yearOfferText, `
    (U 0xC5D0,0xC774,0xC2A4,0xBC14,0xC774,0xC624,0xD31C), `
    $offerText, `
    (U 0xC120,0xC801,0xC11C,0xB958)
  if (Test-Path -LiteralPath $defaultPath) { return (Resolve-Path -LiteralPath $defaultPath).Path }

  throw "COA source root was not found. Set PENDING_RECEIPT_COA_SOURCE_ROOT."
}

function Resolve-TargetRoot {
  param([string]$RequestedPath)
  if ($RequestedPath -and (Test-Path -LiteralPath $RequestedPath)) { return (Resolve-Path -LiteralPath $RequestedPath).Path }
  if ($env:PENDING_RECEIPT_COA_TARGET_ROOT -and (Test-Path -LiteralPath $env:PENDING_RECEIPT_COA_TARGET_ROOT)) {
    return (Resolve-Path -LiteralPath $env:PENDING_RECEIPT_COA_TARGET_ROOT).Path
  }

  $defaultPath = "Y:\{0}\A40-{1}({2})\2026{3}\2026 {4}" -f `
    (U 0xC790,0xB8CC,0xACF5,0xC720), `
    (U 0xC5F0,0xB3C4,0xBCC4,0xC790,0xB8CC), `
    (U 0xC870,0xC740,0xBE44), `
    (U 0xC790,0xB8CC), `
    (U 0xC131,0xC801,0xC11C)
  if (Test-Path -LiteralPath $defaultPath) { return (Resolve-Path -LiteralPath $defaultPath).Path }

  throw "COA target root was not found. Set PENDING_RECEIPT_COA_TARGET_ROOT."
}

function Get-SafeFileName {
  param([string]$Value)
  $invalid = [IO.Path]::GetInvalidFileNameChars()
  $clean = $Value
  foreach ($char in $invalid) {
    $clean = $clean.Replace([string]$char, " ")
  }
  return ($clean -replace "\s+", " ").Trim()
}

function Find-PoFolder {
  param([string]$Root, [string]$Po)
  $normalizedPo = Normalize-Text $Po
  $folders = Get-ChildItem -LiteralPath $Root -Directory -Recurse -ErrorAction SilentlyContinue |
    Where-Object { (Normalize-Text $_.Name) -like "*$normalizedPo*" } |
    Sort-Object FullName
  return $folders | Select-Object -First 1
}

function Find-CoaFile {
  param([string]$FolderPath)
  $coaPatterns = @("COA", "CERTIFICATE", "ANALYSIS", (U 0xC131,0xC801), (U 0xBD84,0xC11D))
  $extensions = @(".pdf", ".xlsx", ".xlsm", ".xls", ".jpg", ".jpeg", ".png")
  $files = Get-ChildItem -LiteralPath $FolderPath -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() }

  $named = $files | Where-Object {
    $name = Normalize-Text $_.Name
    $coaPatterns | Where-Object { $name -like "*$(Normalize-Text $_)*" }
  } | Sort-Object FullName

  if ($named) { return $named | Select-Object -First 1 }
  return $files | Sort-Object FullName | Select-Object -First 1
}

function Find-ProductFolder {
  param([string]$Root, [string]$Product)
  $normalizedProduct = Normalize-Text $Product
  $folders = Get-ChildItem -LiteralPath $Root -Directory -Recurse -ErrorAction SilentlyContinue

  $exact = $folders | Where-Object { (Normalize-Text $_.Name) -like "*$normalizedProduct*" } | Sort-Object FullName | Select-Object -First 1
  if ($exact) { return $exact }

  $keywords = ($Product -split "[\s,/()_-]+") | Where-Object { $_.Length -ge 4 } | Select-Object -First 4
  foreach ($keyword in $keywords) {
    $normalizedKeyword = Normalize-Text $keyword
    $match = $folders | Where-Object { (Normalize-Text $_.Name) -like "*$normalizedKeyword*" } | Sort-Object FullName | Select-Object -First 1
    if ($match) { return $match }
  }

  throw "Product folder was not found for '$Product'."
}

function Get-BatchNo {
  param([string[]]$Candidates)
  $patterns = @(
    "(?i)\b(?:batch|lot)\s*[:#_\- ]*\s*([A-Za-z0-9][A-Za-z0-9._\-]{2,})",
    "\b([A-Z]{1,4}\d{4,}[A-Z0-9._\-]*)\b",
    "\b(\d{6,}[A-Z0-9._\-]*)\b"
  )

  foreach ($candidate in $Candidates) {
    foreach ($pattern in $patterns) {
      $match = [regex]::Match($candidate, $pattern)
      if ($match.Success) {
        return $match.Groups[1].Value.Trim("._- ")
      }
    }
  }

  return "BatchCheckNeeded"
}

$sourceRootPath = Resolve-SourceRoot $SourceRoot
$targetRootPath = Resolve-TargetRoot $TargetRoot
$instockDateValue = Convert-ToDateValue $InstockDate
if (-not $instockDateValue) { throw "Invalid instock date: $InstockDate" }

$poFolder = Find-PoFolder $sourceRootPath $PoNo
if (-not $poFolder) { throw "PO folder was not found for '$PoNo'." }

$coaFile = Find-CoaFile $poFolder.FullName
if (-not $coaFile) { throw "COA file was not found under '$($poFolder.FullName)'." }

$productFolder = Find-ProductFolder $targetRootPath $ProductName
$batchNo = Get-BatchNo @($coaFile.BaseName, $poFolder.Name)
$dateToken = $instockDateValue.ToString("yyMMdd")
$customer = if ($Messrs.Trim()) { $Messrs.Trim() } else { "CustomerCheckNeeded" }
$quantityToken = if ($Quantity.Trim()) { $Quantity.Trim() } else { "QuantityCheckNeeded" }
$inStockText = "ABP{0}" -f (U 0xC785,0xACE0)
$targetBaseName = Get-SafeFileName ("{0}, {1}-{2} ({3}), {4}" -f $batchNo, $customer, $inStockText, $dateToken, $quantityToken)
$targetPath = Join-Path $productFolder.FullName ($targetBaseName + $coaFile.Extension)

if ($Commit) {
  Copy-Item -LiteralPath $coaFile.FullName -Destination $targetPath -Force
}

$result = [ordered]@{
  ok = $true
  committed = [bool]$Commit
  poNo = $PoNo
  sourceRoot = $sourceRootPath
  poFolder = $poFolder.FullName
  sourceFile = $coaFile.FullName
  targetRoot = $targetRootPath
  productFolder = $productFolder.FullName
  targetFile = $targetPath
  batchNo = $batchNo
  instockDate = $instockDateValue.ToString("yyyy-MM-dd")
}

$result | ConvertTo-Json -Depth 5 -Compress
