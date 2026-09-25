$ErrorActionPreference = "Stop"

$taskName = "Control de combustible - Tailscale"
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) "Control de combustible\run_tailscale.ps1"
$powerShellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path $scriptPath)) {
    throw "No se encontró el script de arranque: $scriptPath"
}

$action = New-ScheduledTaskAction -Execute $powerShellPath -Argument "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -WorkingDirectory (Split-Path $scriptPath)
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 10 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Start-ScheduledTask -TaskName $taskName
Write-Host "Tarea instalada e iniciada: $taskName"
Write-Host "Script supervisado: $scriptPath"
