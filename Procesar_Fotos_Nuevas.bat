@echo off
chcp 65001 > nul
title Procesador Automático de Fotos — Bodega Central Guzman y Larrain
echo ====================================================================
echo   DEPTO. ELÉCTRICO GUZMÁN Y LARRAÍN — PROCESADOR DE FOTOS BODEGA
echo ====================================================================
echo.
echo [1/2] Analizando fotos nuevas en la carpeta FOTOS...
echo.

"C:\Users\Bodeg\AppData\Local\Programs\Python\Python312\python.exe" "G:\Unidades compartidas\Departamento Eléctrico\COMPRAS DEPTO ELECTRICO\MATERIAL BODEGA CENTRAL\scripts\procesar_fotos_bodega.py"

echo.
echo [2/2] Proceso finalizado.
echo.
echo Presiona cualquier tecla para cerrar esta ventana...
pause > nul
