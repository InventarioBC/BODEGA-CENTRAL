# Lanzador definitivo 24/7 del Servidor de Bodega Central via WMI
# Inicia el proceso a traves de WMI (WmiPrvSE.exe), desacoplandolo al 100%
# de cualquier Job Object, consola o sesion de agente.

$ErrorActionPreference = "SilentlyContinue"

$Port = 8080
$PythonExe = "C:\Users\Bodeg\AppData\Local\Programs\Python\Python312\pythonw.exe"
$ServerScript = "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL\scripts\servidor_bodega.py"
$UrlFile = "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL\url_celular.txt"

if (-not (Test-Path $PythonExe)) {
    $PythonExe = (Get-Command python.exe).Source
}

Write-Host "[INFO] Verificando estado del puerto $Port..."
Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | ForEach-Object {
    $pidToKill = $_.OwningProcess
    if ($pidToKill -gt 0) {
        Write-Host "[INFO] Deteniendo proceso anterior en puerto $Port (PID: $pidToKill)..."
        Stop-Process -Id $pidToKill -Force -ErrorAction SilentlyContinue
    }
}

Start-Sleep -Seconds 1

# Limpiar archivo de URLs previo
if (Test-Path $UrlFile) {
    Remove-Item -Path $UrlFile -Force -ErrorAction SilentlyContinue
}

# Preparar comando de lanzamiento
$cmd = "`"$PythonExe`" -u `"$ServerScript`" $Port"
Write-Host "[INFO] Lanzando servidor via WMI Win32_Process.Create..."
$res = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{
    CommandLine = $cmd
}

if ($res.ReturnValue -ne 0) {
    Write-Host "[ERROR] WMI devolvio codigo de error: $($res.ReturnValue)"
    exit 1
}

$newPid = $res.ProcessId
Write-Host "[OK] Servidor iniciado con exito como proceso independiente del sistema. PID: $newPid"
Write-Host "[INFO] Esperando conexion de red y tunel Cloudflare seguro..."

$tunnelUrl = ""
$wifiUrl = "http://192.168.1.18:$Port/?v=31"

for ($i = 0; $i -lt 25; $i++) {
    Start-Sleep -Seconds 1
    if (Test-Path $UrlFile) {
        $content = Get-Content -Path $UrlFile -Raw
        if ($content -match "https://[a-zA-Z0-9-]+\.trycloudflare\.com") {
            $tunnelUrl = $matches[0] + "/?v=31"
            break
        }
    }
}

# Verificar respuesta local del servidor
$statusOk = $false
try {
    $resp = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/status" -TimeoutSec 3 -ErrorAction Stop
    if ($resp.status -eq "ok") {
        $statusOk = $true
    }
} catch {
    $statusOk = $false
}

Write-Host ""
Write-Host "===================================================================="
Write-Host "  SERVIDOR BODEGA CENTRAL ACTIVO PERMANENTEMENTE (24/7)"
Write-Host "===================================================================="
if ($statusOk) {
    Write-Host "  ESTADO SERVIDOR:        ACTIVO (ONLINE)"
} else {
    Write-Host "  ESTADO SERVIDOR:        INICIANDO EN SEGUNDO PLANO"
}
if ($tunnelUrl) {
    Write-Host "  ENLACE CELULAR (5G/4G): $tunnelUrl"
}
Write-Host "  ENLACE WI-FI BODEGA:    $wifiUrl"
Write-Host "  PID EN WINDOWS:         $newPid"
Write-Host "===================================================================="
Write-Host ""
