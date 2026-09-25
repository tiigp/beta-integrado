$ErrorActionPreference = "Stop"

$startupDir = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupDir "Control de combustible - Tailscale.lnk"
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) "Control de combustible\run_tailscale.ps1"
$powerShellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path $scriptPath)) {
    throw "No se encontró el script de arranque: $scriptPath"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $powerShellPath
$shortcut.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
$shortcut.WorkingDirectory = Split-Path $scriptPath
$shortcut.WindowStyle = 7
$shortcut.Description = "Inicia y supervisa Control de combustible mediante Tailscale"
$shortcut.Save()

Write-Host "Inicio automático instalado para el usuario actual: $shortcutPath"
