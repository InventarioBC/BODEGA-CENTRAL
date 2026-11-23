@echo off
title Servidor Bodega Central - Depto Electrico
color 1F
echo ======================================================================
echo   INICIANDO SERVIDOR WEB BODEGA CENTRAL - CAMARA TABLET Y CELULAR
echo ======================================================================
echo.

set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

if not exist %PYTHON_CMD% (
    set PYTHON_CMD=python
)

%PYTHON_CMD% "%~dp0scripts\servidor_bodega.py" 8080

echo.
pause
