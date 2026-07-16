$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$launcherPath = Join-Path $root "scripts\start-workspace-local-server.ps1"

if (-not (Test-Path -LiteralPath $launcherPath)) {
  throw "Local launcher script was not found: $launcherPath"
}

$powershellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$protocolRoot = "HKCU:\Software\Classes\haemin-workspace"
$commandKey = Join-Path $protocolRoot "shell\open\command"
$command = "`"$powershellPath`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$launcherPath`" `"%1`""

New-Item -Path $protocolRoot -Force | Out-Null
Set-Item -Path $protocolRoot -Value "URL:Haemin WorkSpace Local Launcher"
New-ItemProperty -Path $protocolRoot -Name "URL Protocol" -Value "" -PropertyType String -Force | Out-Null
New-Item -Path $commandKey -Force | Out-Null
Set-Item -Path $commandKey -Value $command

Write-Output "haemin-workspace://start protocol registered."
