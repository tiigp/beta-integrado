$ErrorActionPreference = "Stop"
$taskName = "Control de combustible - Tailscale"
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($task) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Tarea eliminada: $taskName"
} else {
    Write-Host "La tarea no estaba instalada: $taskName"
}
