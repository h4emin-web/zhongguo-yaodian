$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$envPath = Join-Path $root ".env.local"
$logDir = Join-Path $root "logs"
$port = if ($env:PORT) { [int]$env:PORT } else { 4173 }

New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir "workspace-local-server.log"

function Write-Log {
  param([string]$Message)
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Add-Content -Path $logPath -Value "[$timestamp] $Message" -Encoding UTF8
}

function Import-EnvFile {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    throw ".env.local file was not found: $Path"
  }

  Get-Content -LiteralPath $Path -Encoding UTF8 | ForEach-Object {
    $line = $_.Trim()

    if (-not $line -or $line.StartsWith("#")) {
      return
    }

    $separator = $line.IndexOf("=")

    if ($separator -le 0) {
      return
    }

    $name = $line.Substring(0, $separator).Trim()
    $value = $line.Substring($separator + 1).Trim()

    if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
      $value = $value.Substring(1, $value.Length - 2)
    }

    [Environment]::SetEnvironmentVariable($name, $value, "Process")
  }
}

function Test-PortListening {
  param([int]$Port)

  $connection = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    Where-Object { $_.OwningProcess -ne $PID } |
    Select-Object -First 1

  return [bool]$connection
}

Import-EnvFile $envPath

if (-not $env:PORT) {
  $env:PORT = [string]$port
}

Set-Location -LiteralPath $root
Write-Log "Haemin Workspace local server launcher started. Root=$root Port=$env:PORT"

if (Test-PortListening -Port ([int]$env:PORT)) {
  Write-Log "Port $env:PORT is already listening. Launcher exits without starting a duplicate server."
  exit 0
}

while ($true) {
  Write-Log "Starting workspace-local-server.mjs"
  & node "scripts\workspace-local-server.mjs" *> $logPath
  $exitCode = $LASTEXITCODE
  Write-Log "workspace-local-server.mjs exited with code $exitCode. Restarting in 5 seconds."
  Start-Sleep -Seconds 5
}
