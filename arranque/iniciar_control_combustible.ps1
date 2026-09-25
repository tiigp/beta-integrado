$ErrorActionPreference = "Continue"

$projectDir = Join-Path $PSScriptRoot "..\Control de combustible"
$pythonPath = "C:\Users\PC\AppData\Local\Python\pythoncore-3.14-64\python.exe"
$logPath = Join-Path $projectDir "logs\control-combustible-service.log"
$port = 5000

New-Item -ItemType Directory -Force -Path (Split-Path $logPath) | Out-Null

$mutex = New-Object System.Threading.Mutex($false, "Global\ControlCombustibleStartup")
if (-not $mutex.WaitOne(0)) {
    exit 0
}

# Evita que una instancia manual o antigua compita por el puerto del sistema.
$owners = @(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique)
foreach ($owner in $owners) {
    Stop-Process -Id $owner -Force -ErrorAction SilentlyContinue
}

try {
    while ($true) {
        $startedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logPath -Value "[$startedAt] Iniciando Control de combustible."

        Push-Location $projectDir
        try {
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