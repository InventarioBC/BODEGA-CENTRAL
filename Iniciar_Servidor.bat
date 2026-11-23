@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
"%PYTHON_EXE%" "%~dp0scripts\servidor_bodega.py" 8080 > "%~dp0scripts\servidor_service.log" 2>&1
