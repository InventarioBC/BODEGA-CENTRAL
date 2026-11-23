$ErrorActionPreference = "SilentlyContinue"

$Port = 8080
$PythonExe = "C:\Users\Bodeg\AppData\Local\Programs\Python\Python312\pythonw.exe"
$ServerScript = "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL\scripts\servidor_bodega.py"
$BaseDir = "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL"

# Liberar puerto anterior
Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.OwningProcess -gt 0) {
        Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
    }
}
Start-Sleep -Seconds 1

# Iniciar proceso con Start-Process sin dobles comillas redundantes
Start-Process -FilePath $PythonExe -ArgumentList @($ServerScript, "8080") -WorkingDirectory $BaseDir

Write-Host "[OK] Proceso pythonw lanzado."
