#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lanzador permanente y desacoplado del Servidor de Bodega Central.

Inicia el servidor en segundo plano como proceso independiente de Windows
(DETACHED_PROCESS + CREATE_NO_WINDOW), garantizando que se mantenga en ejecución
24/7 sin cerrarse al cerrar consolas o sesiones de agentes.
"""

import os
import subprocess
import sys
import time
import urllib.request
import json
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPTS_DIR.parent
SERVER_SCRIPT = SCRIPTS_DIR / "servidor_bodega.py"
LOG_FILE = SCRIPTS_DIR / "servidor_service.log"
URL_FILE = BASE_DIR / "url_celular.txt"

# Usar pythonw.exe si está disponible para ejecución totalmente desacoplada sin ventana
sys_py = Path(sys.executable)
pythonw_candidate = sys_py.parent / "pythonw.exe"
PYTHON_EXE = pythonw_candidate if pythonw_candidate.exists() else sys_py

DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
CREATE_NO_WINDOW = 0x08000000


def is_server_running(port: int = 8080) -> bool:
    try:
        url = f"http://127.0.0.1:{port}/api/status"
        req = urllib.request.Request(url, headers={"User-Agent": "HealthCheck/1.0"})
        with urllib.request.urlopen(req, timeout=1.5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data.get("status") == "ok"
    except Exception:
        return False


def stop_previous_server(port: int = 8080):
    """Detiene cualquier servidor previo en el puerto especificado."""
    print(f"[INFO] Verificando y liberando puerto {port}...")
    try:
        subprocess.run(
            ["powershell", "-Command", f"Get-NetTCPConnection -LocalPort {port} -ErrorAction SilentlyContinue | ForEach-Object {{ Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }}"],
            timeout=8,
            creationflags=CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
        time.sleep(1.5)
    except Exception as e:
        print(f"[WARN] No se pudo liberar puerto {port}: {e}")


def start_service(port: int = 8080, force_restart: bool = False):
    if force_restart:
        stop_previous_server(port)
    elif is_server_running(port):
        print(f"[OK] El servidor ya está activo y respondiendo en el puerto {port}.")
        if URL_FILE.exists():
            print("\n" + URL_FILE.read_text(encoding="utf-8"))
        return

    # Si el puerto está ocupado por algo que no responde bien, liberarlo
    stop_previous_server(port)

    # Limpiar archivo de URL anterior para esperar la nueva
    if URL_FILE.exists():
        try:
            URL_FILE.unlink()
        except Exception:
            pass

    print(f"[INFO] Iniciando servidor desacoplado en segundo plano con {PYTHON_EXE.name} (puerto {port})...")
    log_fp = open(LOG_FILE, "a", encoding="utf-8", buffering=1)
    log_fp.write(f"\n\n=== INICIO DE SERVICIO {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
    log_fp.flush()

    cmd = [str(PYTHON_EXE), "-u", str(SERVER_SCRIPT), str(port)]
    flags = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW

    p = subprocess.Popen(
        cmd,
        stdin=subprocess.DEVNULL,
        stdout=log_fp,
        stderr=log_fp,
        cwd=str(BASE_DIR),
        creationflags=flags,
    )
    print(f"[OK] Proceso independiente de Windows iniciado con PID: {p.pid}")

    # Esperar hasta 20 segundos a que el túnel Cloudflare y el servidor estén listos
    print("[INFO] Conectando servidor y túnel seguro HTTPS Cloudflare...")
    tunnel_url = None
    local_url = None
    for _ in range(20):
        time.sleep(1)
        if URL_FILE.exists():
            content = URL_FILE.read_text(encoding="utf-8")
            if "trycloudflare.com" in content:
                for line in content.splitlines():
                    if "trycloudflare.com" in line:
                        tunnel_url = line.strip()
                    elif "http://" in line and "192.168" in line:
                        local_url = line.strip()
                if tunnel_url:
                    break

    if is_server_running(port):
        print("\n" + "=" * 68)
        print("  ¡SERVIDOR BODEGA CENTRAL ACTIVO PERMANENTEMENTE!")
        print("=" * 68)
        if tunnel_url:
            print(f"  ENLACE CELULAR (4G / 5G): {tunnel_url}")
        if local_url:
            print(f"  ENLACE WI-FI BODEGA:      {local_url}")
        else:
            print(f"  ENLACE WI-FI BODEGA:      http://192.168.1.18:{port}/?v=30")
        print("=" * 68 + "\n")
    else:
        print("[WARN] El servidor tardó más de lo esperado en responder. Revisa scripts/servidor_service.log.")


if __name__ == "__main__":
    force = "--restart" in sys.argv or "-r" in sys.argv
    start_service(8080, force_restart=force)
