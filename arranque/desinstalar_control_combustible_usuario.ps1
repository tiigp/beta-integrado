$shortcutPath = Join-Path ([Environment]::GetFolderPath("Startup")) "Control de combustible - Tailscale.lnk"
if (Test-Path $shortcutPath) {
    Remove-Item $shortcutPath -Force
    Write-Host "Inicio automático del usuario eliminado."
} else {
    Write-Host "No se encontró el acceso automático del usuario."
}
