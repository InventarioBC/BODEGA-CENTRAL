# Inicia el servidor de Bodega Central de manera permanente a traves del Shell de Windows
$ErrorActionPreference = "SilentlyContinue"

$Port = 8080
$BaseDir = "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL"
$ScriptsDir = Join-Path $BaseDir "scripts"
$ServerScript = Join-Path $ScriptsDir "servidor_bodega.py"
$PythonExe = "C:\Users\Bodeg\AppData\Local\Programs\Python\Python312\pythonw.exe"

# 1. Liberar puerto 8080 si esta ocupado
Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | ForEach-Object {
    $procId = $_.OwningProcess
    if ($procId -gt 0) {
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    }
}
Start-Sleep -Seconds 1

# 2. Ejecutar a traves de Shell.Application usando variable no reservada
$shell = New-Object -ComObject Shell.Application
$paramString = "`"$ServerScript`" $Port"
$shell.ShellExecute($PythonExe, $paramString, $BaseDir, "open", 0)

Write-Host "[OK] Solicitud de ejecucion enviada a Windows Shell con argumentos correctos."
