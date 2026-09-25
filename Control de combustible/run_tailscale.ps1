$ErrorActionPreference = "Continue"

$projectDir = $PSScriptRoot
$pythonPath = "C:\Users\PC\AppData\Local\Python\pythoncore-3.14-64\python.exe"
$logPath = Join-Path $projectDir "logs\control-combustible-service.log"
$port = 5000

New-Item -ItemType Directory -Force -Path (Split-Path $logPath) | Out-Null

$mutex = New-Object System.Threading.Mutex($false, "Global\ControlCombustibleStartup")
if (-not $mutex.WaitOne(0)) {
    exit 0
}

try {
    while ($true) {
        $startedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logPath -Value "[$startedAt] Iniciando Control de combustible para proxy HTTPS."

        Push-Location $projectDir
        try {
            $env:PORT = "$port"
            $env:APP_SSL_CERT = "__https_disabled_for_tailscale_proxy__"
            $env:APP_SSL_KEY = "__https_disabled_for_tailscale_proxy__"
            & $pythonPath "app.py" *>> $logPath
        }
        catch {
            Add-Content -Path $logPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error: $($_.Exception.Message)"
        }
        finally {
            Pop-Location
        }

        Add-Content -Path $logPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] La aplicacion se detuvo; reintentando en 5 segundos."
        Start-Sleep -Seconds 5
    }
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
