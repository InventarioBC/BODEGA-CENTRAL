#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Servidor Web Local + Tunnel Seguro para Bodega Central y Cámara Móvil (Tablet / Celular).

Permite acceder al Buscador de Materiales desde cualquier tablet, teléfono o PC:
1. Conectado a la red Wi-Fi de bodega (IP local: 192.168.1.18:8080).
2. Conectado mediante datos móviles (5G / 4G) a través de un túnel seguro HTTPS Cloudflare.
3. Tomar fotografías de pallets y subirlas inmediatamente, actualizando el inventario
   y la planilla Excel en tiempo real.
"""

from __future__ import annotations

import atexit
import base64
import io
import json
import mimetypes
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

SCRIPTS_DIR = Path(__file__).resolve().parent

# Safe stdout/stderr for pythonw.exe (windowless background mode)
if sys.stdout is None:
    try:
        sys.stdout = open(SCRIPTS_DIR / "servidor_stdout.log", "a", encoding="utf-8", buffering=1)
    except Exception:
        sys.stdout = io.StringIO()
if sys.stderr is None:
    try:
        sys.stderr = open(SCRIPTS_DIR / "servidor_stderr.log", "a", encoding="utf-8", buffering=1)
    except Exception:
        sys.stderr = io.StringIO()

try:
    import qrcode
except ImportError:
    qrcode = None
BASE_DIR = SCRIPTS_DIR.parent
FOTOS_DIR = BASE_DIR / "FOTOS"
MAPEO_FOTOS_FILE = FOTOS_DIR / "mapeo_fotos.json"
HTML_FILE = BASE_DIR / "Buscador_Bodega_Central.html"
LAUNCHER_HTML_FILE = BASE_DIR / "Abrir_en_Celular.html"
QR_CARD_FILE = BASE_DIR / "QR_Acceso_Celular.png"
URL_TXT_FILE = BASE_DIR / "url_celular.txt"
READY_TO_SHIP_FILE = FOTOS_DIR / "listos_para_enviar.json"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generar_buscador_bodega as gen

DEFAULT_PORT = 8080
PUBLIC_TUNNEL_URL: str | None = None
TUNNEL_PROCESS: subprocess.Popen | None = None


def get_local_ip() -> str:
    """Detect the machine's local IPv4 address."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def find_cloudflared() -> str | None:
    """Find cloudflared executable in standard paths or PATH."""
    standard_paths = [
        Path(r"C:\Program Files (x86)\cloudflared\cloudflared.exe"),
        Path(r"C:\Program Files\cloudflared\cloudflared.exe"),
        Path(r"C:\ProgramData\chocolatey\bin\cloudflared.exe"),
    ]
    for p in standard_paths:
        if p.exists():
            return str(p)
    which_path = shutil.which("cloudflared")
    if which_path:
        return which_path
    return None


def generate_qr_base64(target_url: str) -> str:
    """Generate a base64 encoded PNG data URI for a QR code."""
    if not qrcode:
        return ""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(target_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#0b2239", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/png;base64,{b64}"


def update_launcher_files(public_url: str, local_ip: str, port: int) -> None:
    """Update Abrir_en_Celular.html, QR_Acceso_Celular.png, and url_celular.txt with active URLs."""
    local_url = f"http://{local_ip}:{port}/?v=35"
    primary_url = f"{public_url.rstrip('/')}/?v=35" if public_url else local_url
    apk_url = f"{public_url.rstrip('/')}/BodegaCentral.apk" if public_url else f"http://{local_ip}:{port}/BodegaCentral.apk"
    qr_b64 = generate_qr_base64(primary_url)

    # 1. Save plain text file with URL for easy copy/paste
    URL_TXT_FILE.write_text(
        f"ENLACE MOVIL (DATOS 5G / 4G):\n{primary_url}\n\n"
        f"ENLACE WI-FI BODEGA (PERMANENTE / NUNCA VENCE):\n{local_url}\n\n"
        f"DESCARGAR APP ANDROID (.APK):\n{apk_url}\n",
        encoding="utf-8",
    )

    # 2. Save standalone high-res QR card
    if qrcode:
        try:
            qr = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=12,
                border=3,
            )
            qr.add_data(primary_url)
            qr.make(fit=True)
            img = qr.make_image(fill_color="#0b2239", back_color="white")
            img.save(str(QR_CARD_FILE))
        except Exception as e_qr:
            print(f"[WARN] Error guardando QR_Acceso_Celular.png: {e_qr}")

    # 3. Save Abrir_en_Celular.html with embedded base64 QR and Drive workaround instructions
    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="Bodega Central">
    <title>Bodega Central — Acceso Celular</title>
    <link rel="icon" href="icon.png">
    <link rel="apple-touch-icon" href="apple-touch-icon.png">
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
        body {{ background: #EAEFF4; color: #2D3748; min-height: 100vh; padding: 16px; display: flex; align-items: center; justify-content: center; }}
        .card {{ background: white; border-radius: 24px; padding: 26px 20px; max-width: 440px; width: 100%; box-shadow: 0 16px 40px rgba(88, 111, 134, 0.14); border: 1px solid #E2E8F0; text-align: center; }}
        .badge-status {{ display: inline-flex; align-items: center; gap: 6px; background: #EBF5F0; color: #2D7A56; font-weight: 700; font-size: 0.8rem; padding: 6px 14px; border-radius: 9999px; border: 1px solid #86C1A4; margin-bottom: 12px; }}
        .badge-status::before {{ content: ""; width: 8px; height: 8px; background: #57A773; border-radius: 50%; display: inline-block; }}
        h1 {{ font-size: 1.45rem; font-weight: 800; color: #4A657E; margin-bottom: 4px; }}
        p.subtitle {{ color: #64748b; font-size: 0.88rem; margin-bottom: 18px; }}
        .btn-launch {{ display: block; width: 100%; background: linear-gradient(135deg, #65859F 0%, #4D6E87 100%); color: white !important; text-decoration: none; font-weight: 700; font-size: 1.12rem; padding: 15px 16px; border-radius: 14px; box-shadow: 0 4px 16px rgba(77, 110, 135, 0.35); margin-bottom: 12px; transition: transform 0.15s ease; border: none; cursor: pointer; text-align: center; }}
        .btn-launch:active {{ transform: scale(0.98); }}
        .alert-drive {{ background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 12px; padding: 12px; font-size: 0.84rem; color: #92400E; margin-bottom: 14px; text-align: left; line-height: 1.45; }}
        .url-box {{ background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 10px 12px; font-family: monospace; font-size: 0.82rem; color: #2D3748; word-break: break-all; margin-bottom: 8px; user-select: all; }}
        .btn-copy {{ display: flex; align-items: center; justify-content: center; gap: 6px; width: 100%; background: #EDF3F8; color: #4A657E; font-weight: 600; font-size: 0.92rem; padding: 11px 16px; border-radius: 12px; border: 1px solid #CFDCE8; cursor: pointer; margin-bottom: 16px; }}
        .btn-copy:active {{ background: #DEE9F3; }}
        .qr-section {{ background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 16px; padding: 14px; margin-bottom: 16px; }}
        .qr-img {{ width: 190px; height: 190px; margin: 0 auto 8px auto; display: block; border-radius: 12px; border: 1px solid #CBD5E1; background: white; padding: 6px; }}
        .instructions {{ background: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 14px; padding: 14px 16px; text-align: left; font-size: 0.84rem; color: #1E3A8A; line-height: 1.5; }}
        .instructions ol {{ margin-left: 18px; margin-top: 6px; }}
    </style>
</head>
<body>
    <div class="card">
        <div class="badge-status">Servidor Activo (5G / 4G / Wi-Fi)</div>
        <h1>Bodega Central</h1>
        <p class="subtitle">Departamento Eléctrico Guzmán y Larraín</p>

        <!-- Botón 1: Datos Móviles / Acceso Global -->
        <a id="btnLink" href="{primary_url}" target="_blank" rel="noopener noreferrer" class="btn-launch">
            🌐 ENTRAR CON DATOS MÓVILES (5G / 4G)
        </a>

        <!-- Botón 2: Wi-Fi Bodega (Enlace Permanente que nunca vence) -->
        <a id="btnWifi" href="{local_url}" target="_blank" rel="noopener noreferrer" style="display: block; width: 100%; background: linear-gradient(135deg, #57A773 0%, #449260 100%); color: white !important; text-decoration: none; font-weight: 700; font-size: 1.05rem; padding: 14px 16px; border-radius: 14px; box-shadow: 0 4px 14px rgba(68, 146, 96, 0.3); margin-bottom: 14px; text-align: center;">
            📶 ENTRAR POR WI-FI BODEGA (Enlace Fijo Permanente)
        </a>

        <!-- Botón 3: Descargar App Android (.APK) -->
        <a id="btnApk" href="{apk_url}" download="BodegaCentral.apk" style="display: block; width: 100%; background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); color: white !important; text-decoration: none; font-weight: 700; font-size: 1.05rem; padding: 14px 16px; border-radius: 14px; box-shadow: 0 4px 14px rgba(2, 132, 199, 0.3); margin-bottom: 14px; text-align: center;">
            🤖 DESCARGAR APP ANDROID (.APK)
        </a>

        <div style="background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 12px; padding: 12px; font-size: 0.84rem; color: #166534; margin-bottom: 14px; text-align: left; line-height: 1.45;">
            <strong>💡 Recomendación para tu Pantalla de Inicio:</strong><br>
            Si trabajas conectado al Wi-Fi de bodega, abre el botón verde de <strong>Wi-Fi Bodega</strong> y agrégalo a tu pantalla de inicio. ¡Ese enlace es fijo y no cambia!
        </div>

        <!-- Mensaje de ayuda si está dentro de la app Google Drive en iPhone -->
        <div class="alert-drive">
            <strong>⚠️ ¿Estás viendo esto dentro de Google Drive?</strong><br>
            Google Drive en iPhone bloquea los enlaces. Para abrir:<br>
            1. Toca los <strong>tres puntos (•••)</strong> arriba a la derecha y elige <strong>"Abrir en Safari"</strong>.<br>
            2. O presiona <strong>Copiar Enlace</strong> abajo y pégalo en Safari.
        </div>

        <div class="url-box" id="urlText">{primary_url}</div>
        <button type="button" class="btn-copy" onclick="copyLink()">
            📋 Copiar Enlace para Safari
        </button>

        <!-- Código QR en Base64 (Nunca se rompe en Drive) -->
        <div class="qr-section">
            <p style="font-size: 0.8rem; font-weight: 700; color: #475569; margin-bottom: 8px;">O ESCANEA CON LA CÁMARA DE TU IPHONE:</p>
            <img src="{qr_b64}" alt="Código QR Bodega Central" class="qr-img">
            <small style="color: #64748b; font-size: 0.74rem;">Funciona conectado a Wi-Fi o datos móviles 5G</small>
        </div>

        <!-- Instrucciones PWA -->
        <div class="instructions">
            <strong>📲 Cómo dejarlo instalado como App en tu celular:</strong>
            <ol>
                <li>Abre el enlace en <strong>Safari</strong> (iPhone) o <strong>Chrome</strong> (Android).</li>
                <li>Toca el botón <strong>Compartir</strong> (cuadrado con flecha hacia arriba ⎋).</li>
                <li>Selecciona <strong>"Agregar a pantalla de inicio"</strong>.</li>
            </ol>
        </div>
    </div>

    <script>
        function copyLink() {{
            const url = document.getElementById('urlText').innerText.trim();
            if (navigator.clipboard && navigator.clipboard.writeText) {{
                navigator.clipboard.writeText(url).then(() => {{
                    alert('¡Enlace copiado! Ábrelo en Safari para usar la app.');
                }}).catch(() => {{
                    prompt('Copia este enlace y pégalo en Safari:', url);
                }});
            }} else {{
                prompt('Copia este enlace y pégalo en Safari:', url);
            }}
        }}

        // Redirección suave si se abre fuera de Google Drive
        if (!navigator.userAgent.includes('GoogleDrive') && window.location.protocol.startsWith('file')) {{
            // No auto-redirect on local files to allow viewing QR
        }}
    </script>
</body>
</html>
"""
    LAUNCHER_HTML_FILE.write_text(html_content, encoding="utf-8")
    print(f"[OK] Archivos lanzadores móviles actualizados: {LAUNCHER_HTML_FILE}")


LAST_SERVER_UPDATE = time.time()


def rebuild_inventory_async():
    global LAST_SERVER_UPDATE

    def _worker():
        global LAST_SERVER_UPDATE
        try:
            import importlib
            importlib.reload(gen)
            gen.build_all()
            LAST_SERVER_UPDATE = time.time()
        except Exception as e_reb:
            print(f"[WARN] Error al regenerar buscador en background: {e_reb}")
    threading.Thread(target=_worker, daemon=True).start()


FALLBACK_TUNNEL_PROCESS = None
FALLBACK_TUNNEL_STARTED = False


def start_ssh_tunnel(port: int) -> None:
    """Fallback secure HTTPS tunnel using native OpenSSH and localhost.run."""
    global PUBLIC_TUNNEL_URL, FALLBACK_TUNNEL_PROCESS, FALLBACK_TUNNEL_STARTED

    if FALLBACK_TUNNEL_STARTED:
        return
    FALLBACK_TUNNEL_STARTED = True

    ssh_bin = shutil.which("ssh") or r"C:\WINDOWS\System32\OpenSSH\ssh.exe"
    if not Path(ssh_bin).exists():
        print("[WARN] ssh.exe no encontrado para túnel fallback.")
        return

    def ssh_worker():
        global PUBLIC_TUNNEL_URL, FALLBACK_TUNNEL_PROCESS
        while True:
            try:
                cmd = [
                    ssh_bin,
                    "-o", "StrictHostKeyChecking=no",
                    "-o", "ServerAliveInterval=30",
                    "-R", f"80:127.0.0.1:{port}",
                    "nokey@localhost.run"
                ]
                creationflags = 0x08000000 if sys.platform == "win32" else 0
                FALLBACK_TUNNEL_PROCESS = subprocess.Popen(
                    cmd,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    creationflags=creationflags,
                )
                print(f"[OK] Túnel SSH fallback iniciado con PID: {FALLBACK_TUNNEL_PROCESS.pid}")

                for line in iter(FALLBACK_TUNNEL_PROCESS.stdout.readline, ""):
                    m = re.search(r"https://[a-zA-Z0-9-]+\.lhr\.life", line)
                    if m:
                        new_url = m.group(0)
                        # NEVER overwrite an active Cloudflare tunnel with slow SSH fallback
                        if not PUBLIC_TUNNEL_URL or "lhr.life" in PUBLIC_TUNNEL_URL:
                            PUBLIC_TUNNEL_URL = new_url
                            local_ip = get_local_ip()
                            update_launcher_files(PUBLIC_TUNNEL_URL, local_ip, port)
                            print("\n" + "=" * 68)
                            print("  ¡TÚNEL SEGURO HTTPS CONECTADO (SSH FALLBACK)!")
                            print("=" * 68)
                            print(f"  Acceso Celular (5G / 4G):  {PUBLIC_TUNNEL_URL}")
                            print(f"  Acceso Wi-Fi Bodega:       http://{local_ip}:{port}/")
                            print("=" * 68 + "\n")
                FALLBACK_TUNNEL_PROCESS.wait()
            except Exception as e_ssh:
                print(f"[WARN] Error en túnel SSH fallback: {e_ssh}")
            time.sleep(5)

    t = threading.Thread(target=ssh_worker, daemon=True)
    t.start()


def start_cloudflare_tunnel(port: int) -> None:
    """Start cloudflared quick tunnel in background and extract public HTTPS URL, with SSH fallback."""
    global PUBLIC_TUNNEL_URL, TUNNEL_PROCESS

    cf_bin = find_cloudflared()
    if not cf_bin:
        print("[WARN] cloudflared no encontrado. Iniciando túnel SSH...")
        start_ssh_tunnel(port)
        return

    cf_log = Path(r"C:\Users\Bodeg\cloudflared.log")

    def tunnel_worker():
        global PUBLIC_TUNNEL_URL, TUNNEL_PROCESS
        backoff = 5
        while True:
            try:
                # Truncate log file safely
                cmd = [
                    cf_bin,
                    "tunnel",
                    "--url", f"http://127.0.0.1:{port}",
                ]
                creationflags = 0x08000000 if sys.platform == "win32" else 0
                t_start = time.time()
                TUNNEL_PROCESS = subprocess.Popen(
                    cmd,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    creationflags=creationflags,
                )
                print(f"[OK] cloudflared iniciado con PID: {TUNNEL_PROCESS.pid}")

                found_url = False
                for line in iter(TUNNEL_PROCESS.stdout.readline, ""):
                    m = re.search(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com", line)
                    if m:
                        new_url = m.group(0)
                        found_url = True
                        if new_url != PUBLIC_TUNNEL_URL:
                            PUBLIC_TUNNEL_URL = new_url
                            local_ip = get_local_ip()
                            update_launcher_files(PUBLIC_TUNNEL_URL, local_ip, port)
                            print("\n" + "=" * 68)
                            print("  ¡TÚNEL SEGURO HTTPS CLOUDFLARE CONECTADO EXITOSAMENTE!")
                            print("=" * 68)
                            print(f"  Acceso Celular (5G / 4G):  {PUBLIC_TUNNEL_URL}")
                            print(f"  Acceso Wi-Fi Bodega:       http://{local_ip}:{port}/")
                            print("=" * 68 + "\n")
                            if FALLBACK_TUNNEL_PROCESS:
                                try:
                                    FALLBACK_TUNNEL_PROCESS.terminate()
                                except Exception:
                                    pass
                    if "429" in line:
                        print("[WARN] Cloudflare respondió 429. Activando túnel fallback...")
                        start_ssh_tunnel(port)

                # Esperar a que el proceso realmente termine si aún sigue vivo
                if TUNNEL_PROCESS.poll() is None:
                    TUNNEL_PROCESS.wait()

                duration = time.time() - t_start
                if duration > 60:
                    backoff = 5
                else:
                    backoff = min(backoff * 2, 120)

                print(f"[WARN] cloudflared terminó con código {TUNNEL_PROCESS.returncode}. Reintentando en {backoff}s...")
                # Asegurar que el túnel fallback esté activo si cloudflared cae
                if not PUBLIC_TUNNEL_URL or "trycloudflare" in PUBLIC_TUNNEL_URL:
                    start_ssh_tunnel(port)
            except Exception as e_tun:
                print(f"[ERROR] Error en túnel Cloudflare: {e_tun}")
                start_ssh_tunnel(port)
                backoff = min(backoff * 2, 120)
            time.sleep(backoff)

    t = threading.Thread(target=tunnel_worker, daemon=True)
    t.start()


def cleanup():
    """Cleanup background tunnel processes on exit."""
    global TUNNEL_PROCESS, FALLBACK_TUNNEL_PROCESS
    for proc in (TUNNEL_PROCESS, FALLBACK_TUNNEL_PROCESS):
        if proc:
            try:
                proc.terminate()
                proc.wait(timeout=2)
            except Exception:
                pass


atexit.register(cleanup)


def get_next_pallet_info() -> dict:
    """Inspect mapeo_fotos.json and calculate next available pallet number."""
    mapping = gen.load_photo_mapping()
    max_num = 0
    total_pallets = 0
    seen_pallets = set()
    for info in mapping.values():
        p_list = info.get("pallets") if isinstance(info.get("pallets"), list) else [info]
        for p_item in p_list:
            p_name = str(p_item.get("pallet", "")).strip()
            if p_name and p_name.lower() not in seen_pallets:
                seen_pallets.add(p_name.lower())
                total_pallets += 1
            m = re.search(r"(?:bulto|pallet)\s*(\d+)", p_name.lower())
            if m:
                max_num = max(max_num, int(m.group(1)))
    next_num = max_num + 1 if max_num > 0 else 1
    return {
        "next_pallet_num": next_num,
        "next_pallet_name": f"Bulto {next_num}",
        "total_pallets_registered": total_pallets,
    }


class BodegaHTTPRequestHandler(SimpleHTTPRequestHandler):
    """Custom HTTP handler serving the dashboard, API, and file uploads."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html", "/Buscador_Bodega_Central.html"):
            if HTML_FILE.exists():
                data = HTML_FILE.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                accept_encoding = self.headers.get("Accept-Encoding", "")
                if "gzip" in accept_encoding:
                    import gzip
                    compressed = gzip.compress(data, compresslevel=6)
                    self.send_header("Content-Encoding", "gzip")
                    self.send_header("Content-Length", str(len(compressed)))
                    self.end_headers()
                    self.wfile.write(compressed)
                else:
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                return

        if path == "/sw.js":
            sw_path = BASE_DIR / "sw.js"
            if sw_path.exists():
                data = sw_path.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/javascript; charset=utf-8")
                self.send_header("Service-Worker-Allowed", "/")
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

        if path == "/manifest.json":
            m_path = BASE_DIR / "manifest.json"
            if m_path.exists():
                data = m_path.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/manifest+json; charset=utf-8")
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return

        if path == "/api/status":
            local_ip = get_local_ip()
            self.send_json({
                "status": "ok",
                "ip": local_ip,
                "port": self.server.server_port,
                "public_url": PUBLIC_TUNNEL_URL,
                "server_time": datetime.now().isoformat(),
                "server_name": "Servidor Bodega Central - Depto. Eléctrico",
                "last_updated": LAST_SERVER_UPDATE,
            })
            return

        if path == "/api/next_pallet":
            self.send_json(get_next_pallet_info())
            return

        if path == "/api/qrcode.png":
            local_ip = get_local_ip()
            target_url = PUBLIC_TUNNEL_URL or f"http://{local_ip}:{self.server.server_port}/"
            if qrcode:
                img = qrcode.make(target_url)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                png_bytes = buf.getvalue()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(png_bytes)))
                self.end_headers()
                self.wfile.write(png_bytes)
            else:
                self.send_error(HTTPStatus.NOT_FOUND, "Librería qrcode no disponible")
            return

        if path in ("/api/sync_inventory", "/api/inventario"):
            try:
                import importlib
                importlib.reload(gen)
                consolidated, sync_ts, src_info, photo_map = gen.get_consolidated_inventory()
                rebuild_inventory_async()
                self.send_json({
                    "success": True,
                    "sync_ts": sync_ts,
                    "src_info": src_info,
                    "total_items": len(consolidated),
                    "items": consolidated,
                    "pallets": photo_map,
                    "server_time": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                })
                print(f"[OK] Sincronización ultrarrápida enviada al cliente: {len(consolidated)} materiales, {len(photo_map)} pallets")
            except Exception as e:
                print(f"[ERROR] Error al sincronizar inventario: {e}")
                self.send_json({"success": False, "error": str(e)}, status=500)
            return

        if path == "/api/pallets":
            mapping = gen.load_photo_mapping()
            self.send_json({"success": True, "pallets": mapping})
            return

        if path == "/api/pallet_detail":
            query = parse_qs(parsed.query)
            pallet_name = query.get("pallet", [""])[0].strip()
            self.handle_get_pallet_detail(pallet_name)
            return

        if path == "/api/gemini_status":
            import gemini_vision_scanner as gemini_scan
            has_key = gemini_scan.is_gemini_available()
            self.send_json({"success": True, "gemini_active": has_key})
            return

        if path == "/api/ready_to_ship":
            self.handle_get_ready_to_ship()
            return

        super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/api/sync_inventory", "/api/inventario"):
            self.path = parsed.path
            self.do_GET()
            return
        elif parsed.path == "/api/ready_to_ship":
            self.handle_save_ready_to_ship()
            return
        elif parsed.path == "/api/upload_pallet":
            self.handle_upload_pallet()
            return
        elif parsed.path == "/api/link_pallet_photo":
            self.handle_link_pallet_photo()
            return
        elif parsed.path == "/api/scan_pallet":
            self.handle_scan_pallet()
            return
        elif parsed.path == "/api/set_gemini_key":
            self.handle_set_gemini_key()
            return
        elif parsed.path == "/api/edit_pallet":
            self.handle_edit_pallet()
            return
        elif parsed.path == "/api/delete_pallet":
            self.handle_delete_pallet()
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Endpoint desconocido")

    def handle_get_ready_to_ship(self):
        try:
            if READY_TO_SHIP_FILE.exists():
                data = json.loads(READY_TO_SHIP_FILE.read_text(encoding="utf-8"))
            else:
                data = {"codes": [], "pallets": [], "pallet_notes": {}, "updated_at": ""}
            if "pallet_notes" not in data or not isinstance(data.get("pallet_notes"), dict):
                data["pallet_notes"] = {}
            self.send_json({"success": True, **data})
        except Exception as e:
            print(f"[WARN] Error leyendo listos_para_enviar.json: {e}")
            self.send_json({"success": False, "codes": [], "pallets": [], "pallet_notes": {}, "error": str(e)})

    def handle_save_ready_to_ship(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo vacío"}, status=400)
                return
            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))

            current = {"codes": [], "pallets": [], "pallet_notes": {}}
            if READY_TO_SHIP_FILE.exists():
                try:
                    current = json.loads(READY_TO_SHIP_FILE.read_text(encoding="utf-8"))
                except Exception:
                    pass

            codes_set = set(current.get("codes", []))
            pallets_set = set(current.get("pallets", []))
            pallet_notes = dict(current.get("pallet_notes", {})) if isinstance(current.get("pallet_notes"), dict) else {}

            # Reemplazo de lista completa si se envía lista
            has_pallets_list = "pallets" in data and isinstance(data["pallets"], list)
            if has_pallets_list:
                pallets_set = set(data["pallets"])
            if "pallet_notes" in data and isinstance(data["pallet_notes"], dict):
                pallet_notes = dict(data["pallet_notes"])

            has_codes_list = "codes" in data and isinstance(data["codes"], list)
            if has_codes_list:
                codes_set = set(data["codes"])
            else:
                # Alternar código individual solo si no se envió lista explícita
                toggle_code = data.get("toggle_code") or (data.get("code") if data.get("action") == "toggle_code" else None)
                if toggle_code:
                    toggle_code = str(toggle_code).strip().upper()
                    if toggle_code in codes_set:
                        codes_set.remove(toggle_code)
                    else:
                        codes_set.add(toggle_code)

            # Alternar pallet individual SOLO si no se envió lista explícita
            if not has_pallets_list:
                toggle_pallet = data.get("toggle_pallet") or (data.get("pallet") if data.get("action") == "toggle_pallet" else None)
                if toggle_pallet:
                    toggle_pallet = str(toggle_pallet).strip()
                    # Búsqueda insensible a mayúsculas/minúsculas
                    existing = None
                    for p in pallets_set:
                        if p.strip().lower() == toggle_pallet.lower():
                            existing = p
                            break
                    if existing:
                        pallets_set.remove(existing)
                    else:
                        pallets_set.add(toggle_pallet)

            # Guardar o actualizar nota de un pallet individual
            p_note = data.get("pallet_note")
            if p_note and isinstance(p_note, dict):
                p_name = str(p_note.get("pallet", "")).strip()
                p_txt = str(p_note.get("note", "")).strip()
                if p_name:
                    if p_txt:
                        pallet_notes[p_name] = p_txt
                    else:
                        pallet_notes.pop(p_name, None)

            if data.get("action") == "save_note" or ("pallet" in data and "note" in data):
                p_name = str(data.get("pallet", "")).strip()
                p_txt = str(data.get("note", "")).strip()
                if p_name:
                    if p_txt:
                        pallet_notes[p_name] = p_txt
                    else:
                        pallet_notes.pop(p_name, None)

            # Limpiar todo si se solicita
            if data.get("clear_all") or data.get("action") == "clear_all":
                codes_set.clear()
                pallets_set.clear()
                pallet_notes.clear()

            out_data = {
                "codes": sorted(list(codes_set)),
                "pallets": sorted(list(pallets_set)),
                "pallet_notes": pallet_notes,
                "updated_at": datetime.now().isoformat()
            }
            READY_TO_SHIP_FILE.parent.mkdir(parents=True, exist_ok=True)
            READY_TO_SHIP_FILE.write_text(json.dumps(out_data, ensure_ascii=False, indent=2), encoding="utf-8")
            global LAST_SERVER_UPDATE
            LAST_SERVER_UPDATE = time.time()
            print(f"[DISPATCH] Listos para enviar guardado: {len(pallets_set)} bultos: {list(pallets_set)}, {len(pallet_notes)} notas")
            self.send_json({"success": True, **out_data})
        except Exception as e:
            print(f"[ERROR] Error guardando listos_para_enviar.json: {e}")
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_set_gemini_key(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo de solicitud vacío"}, status=400)
                return
            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))
            api_key = str(data.get("api_key", "")).strip()
            if not api_key:
                self.send_json({"success": False, "error": "Clave API vacía"}, status=400)
                return
            import gemini_vision_scanner as gemini_scan
            ok = gemini_scan.save_gemini_api_key(api_key)
            if ok:
                self.send_json({"success": True, "message": "Clave API de Google Gemini guardada exitosamente"})
            else:
                self.send_json({"success": False, "error": "No se pudo guardar la clave API"}, status=500)
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_scan_pallet(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo de solicitud vacío"}, status=400)
                return

            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))
            foto_base64 = data.get("foto_base64", "")
            if not foto_base64:
                self.send_json({"success": False, "error": "No se recibió imagen para escanear"}, status=400)
                return

            if "," in foto_base64:
                _, b64_clean = foto_base64.split(",", 1)
            else:
                b64_clean = foto_base64
            img_bytes = base64.b64decode(b64_clean)
            try:
                Path(r"C:\Users\Bodeg\last_scanned_image.jpg").write_bytes(img_bytes)
            except Exception:
                pass

            next_info = get_next_pallet_info()
            next_name = next_info.get("next_pallet_name", "Pallet 1")

            import importlib
            import scanner_pallet_ai
            importlib.reload(scanner_pallet_ai)
            res = scanner_pallet_ai.scan_pallet_data(img_bytes, next_pallet_name=next_name)
            self.send_json(res)
        except Exception as e:
            print(f"[ERROR] Error procesando escaneo de pallet: {e}")
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_upload_pallet(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo de solicitud vacío"}, status=400)
                return

            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))

            materiales_list = data.get("materiales", [])
            codigo = str(data.get("codigo", "")).strip().upper()
            if not codigo and materiales_list:
                codigo = str(materiales_list[0].get("codigo", "")).strip().upper()

            pallet = str(data.get("pallet", "")).strip()
            ubicacion = str(data.get("ubicacion", "")).strip()
            cantidad = data.get("cantidad_pallet", data.get("cantidad", 0))
            guia = str(data.get("guia", "")).strip()
            proveedor = str(data.get("proveedor", "")).strip()
            descripcion = str(data.get("descripcion", "")).strip()
            foto_base64 = data.get("foto_base64", "")

            if not codigo and not materiales_list:
                self.send_json({"success": False, "error": "El código del material es obligatorio"}, status=400)
                return
            if not pallet:
                self.send_json({"success": False, "error": "El nombre o número del pallet es obligatorio"}, status=400)
                return
            clean_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", pallet).strip().upper()
            if not clean_name:
                clean_name = f"PALLET_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

            filename = ""
            if foto_base64:
                try:
                    if "," in foto_base64:
                        _, b64_clean = foto_base64.split(",", 1)
                    else:
                        b64_clean = foto_base64
                    img_bytes = base64.b64decode(b64_clean)
                    filename = f"{clean_name}.jpeg"

                    FOTOS_DIR.mkdir(parents=True, exist_ok=True)
                    photo_dest = FOTOS_DIR / filename
                    photo_dest.write_bytes(img_bytes)
                except Exception as img_err:
                    print(f"[WARN] Error procesando imagen base64 para {pallet}: {img_err}")
                    filename = ""

            mapping = gen.load_photo_mapping()
            for k, v in mapping.items():
                if v.get("foto", "").startswith("FOTOS/"):
                    v["foto"] = v["foto"][6:]

            items_to_save = []
            if materiales_list and isinstance(materiales_list, list):
                items_to_save = materiales_list
            else:
                items_to_save = [{
                    "codigo": codigo,
                    "cantidad_pallet": float(cantidad) if cantidad else 0,
                    "descripcion": descripcion,
                }]

            saved_codes = []
            for itm in items_to_save:
                c_code = str(itm.get("codigo", "")).strip().upper()
                if not c_code:
                    continue
                c_qty = itm.get("cantidad_pallet", itm.get("cantidad", cantidad))
                c_desc = itm.get("descripcion", descripcion)

                existing_entry = mapping.get(c_code, {})
                existing_pallets = existing_entry.get("pallets")
                if not isinstance(existing_pallets, list):
                    if existing_entry.get("foto") or existing_entry.get("pallet"):
                        existing_pallets = [{
                            "pallet": existing_entry.get("pallet", ""),
                            "ubicacion": existing_entry.get("ubicacion", ""),
                            "foto": existing_entry.get("foto", ""),
                            "descripcion": existing_entry.get("descripcion", ""),
                            "guia": existing_entry.get("guia", ""),
                            "cantidad_pallet": existing_entry.get("cantidad_pallet", 0),
                            "proveedor": existing_entry.get("proveedor", ""),
                            "fecha_registro": existing_entry.get("fecha_registro", ""),
                        }]
                    else:
                        existing_pallets = []

                for p in existing_pallets:
                    if p.get("foto", "").startswith("FOTOS/"):
                        p["foto"] = p["foto"][6:]

                # Si no se subió foto en esta petición, conservar foto existente si ya tenía
                item_photo = filename
                if not item_photo:
                    for ep in existing_pallets:
                        if ep.get("pallet", "").strip().lower() == pallet.strip().lower() and ep.get("foto"):
                            item_photo = ep.get("foto")
                            break
                    if not item_photo and existing_entry.get("foto") and existing_entry.get("pallet", "").strip().lower() == pallet.strip().lower():
                        item_photo = existing_entry.get("foto")

                new_p = {
                    "pallet": pallet,
                    "ubicacion": ubicacion,
                    "foto": item_photo or "",
                    "descripcion": c_desc,
                    "guia": guia,
                    "cantidad_pallet": float(c_qty) if c_qty else 0,
                    "proveedor": proveedor,
                    "fecha_registro": datetime.now().strftime("%d/%m/%Y %H:%M"),
                }

                # Actualizar si ya existía este pallet en el material, o agregar
                found = False
                for idx_p, ep in enumerate(existing_pallets):
                    if ep.get("pallet", "").strip().lower() == pallet.strip().lower():
                        existing_pallets[idx_p] = new_p
                        found = True
                        break
                if not found:
                    existing_pallets.append(new_p)

                mapping[c_code] = {
                    "pallets": existing_pallets,
                    "pallet": pallet,
                    "ubicacion": ubicacion,
                    "foto": item_photo or "",
                    "descripcion": c_desc,
                    "guia": guia,
                    "cantidad_pallet": float(c_qty) if c_qty else 0,
                    "proveedor": proveedor,
                    "fecha_registro": datetime.now().strftime("%d/%m/%Y %H:%M"),
                }
                saved_codes.append(c_code)

            MAPEO_FOTOS_FILE.write_text(
                json.dumps(mapping, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

            rebuild_inventory_async()

            msg_cods = ", ".join(saved_codes)
            has_photo = bool(filename)
            resp_msg = (
                f"¡Foto guardada exitosamente como {filename} y asociada a {msg_cods} ({pallet})!"
                if has_photo
                else f"¡Contenido de {pallet} guardado exitosamente con {len(saved_codes)} material(es)!"
            )
            self.send_json({
                "success": True,
                "message": resp_msg,
                "has_photo": has_photo,
                "codigo": codigo,
                "codigos": saved_codes,
                "pallet": pallet,
                "ubicacion": ubicacion,
                "filename": filename,
                "foto_url": f"FOTOS/{filename}" if filename else "",
            })
            print(f"[OK] {pallet} ({msg_cods}) registrado exitosamente (Foto: {'Sí' if has_photo else 'No/Pendiente'}).")

        except Exception as e:
            print(f"[ERROR] Error procesando upload de pallet: {e}")
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_link_pallet_photo(self):
        """Asocia / vincula una fotografía tomada a un Pallet QR y a todos sus materiales."""
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo de solicitud vacío"}, status=400)
                return

            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))
            pallet = str(data.get("pallet", "")).strip()
            foto_base64 = data.get("foto_base64", "")

            if not pallet:
                self.send_json({"success": False, "error": "El nombre o código de pallet es obligatorio"}, status=400)
                return
            if not foto_base64:
                self.send_json({"success": False, "error": "No se recibió la fotografía"}, status=400)
                return

            if "," in foto_base64:
                _, b64_clean = foto_base64.split(",", 1)
            else:
                b64_clean = foto_base64
            img_bytes = base64.b64decode(b64_clean)

            clean_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", pallet).strip().upper()
            if not clean_name:
                clean_name = f"PALLET_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            filename = f"{clean_name}.jpeg"

            FOTOS_DIR.mkdir(parents=True, exist_ok=True)
            photo_dest = FOTOS_DIR / filename
            photo_dest.write_bytes(img_bytes)

            mapping = gen.load_photo_mapping()
            updated_codes = []
            pallet_norm = pallet.strip().lower()

            for code, entry in mapping.items():
                pallets_list = entry.get("pallets", [])
                if not isinstance(pallets_list, list):
                    pallets_list = [entry] if entry.get("pallet") else []
                changed = False
                for p in pallets_list:
                    if p.get("pallet", "").strip().lower() == pallet_norm:
                        p["foto"] = filename
                        changed = True
                if entry.get("pallet", "").strip().lower() == pallet_norm:
                    entry["foto"] = filename
                    changed = True
                if changed:
                    entry["pallets"] = pallets_list
                    updated_codes.append(code)

            for k, v in mapping.items():
                if v.get("foto", "").startswith("FOTOS/"):
                    v["foto"] = v["foto"][6:]
                if isinstance(v.get("pallets"), list):
                    for p in v["pallets"]:
                        if p.get("foto", "").startswith("FOTOS/"):
                            p["foto"] = p["foto"][6:]

            MAPEO_FOTOS_FILE.write_text(
                json.dumps(mapping, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

            rebuild_inventory_async()

            self.send_json({
                "success": True,
                "message": f"¡Foto vinculada exitosamente al {pallet} ({len(updated_codes)} materiales actualizados)!",
                "pallet": pallet,
                "filename": filename,
                "foto_url": f"FOTOS/{filename}",
                "updated_codes": updated_codes,
            })
            print(f"[OK] Foto {filename} vinculada exitosamente al {pallet} ({len(updated_codes)} códigos).")

        except Exception as e:
            print(f"[ERROR] Error vinculando foto a pallet: {e}")
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_get_pallet_detail(self, pallet_name: str):
        """Devuelve todos los materiales, rack y foto registrados para un pallet o QR."""
        try:
            if not pallet_name:
                self.send_json({"success": False, "error": "Nombre de pallet requerido"}, status=400)
                return
            mapping = gen.load_photo_mapping()
            pallet_norm = pallet_name.strip().lower()
            matching_materials = []
            rack = ""
            foto = ""
            guia = ""
            prov = ""

            for code, entry in mapping.items():
                pallets_list = entry.get("pallets", [])
                if not isinstance(pallets_list, list):
                    pallets_list = [entry] if entry.get("pallet") else []
                for p in pallets_list:
                    if p.get("pallet", "").strip().lower() == pallet_norm:
                        p_foto = p.get("foto", "")
                        if p_foto and not p_foto.startswith("FOTOS/"):
                            p_foto = f"FOTOS/{p_foto}"
                        matching_materials.append({
                            "codigo": code,
                            "descripcion": p.get("descripcion", entry.get("descripcion", "")),
                            "cantidad": p.get("cantidad_pallet", entry.get("cantidad_pallet", 0)),
                            "guia": p.get("guia", ""),
                            "proveedor": p.get("proveedor", ""),
                            "foto": p_foto,
                        })
                        if not rack and p.get("ubicacion"):
                            rack = p.get("ubicacion")
                        if not foto and p_foto:
                            foto = p_foto
                        if not guia and p.get("guia"):
                            guia = p.get("guia")
                        if not prov and p.get("proveedor"):
                            prov = p.get("proveedor")

            self.send_json({
                "success": True,
                "pallet": pallet_name,
                "found": len(matching_materials) > 0,
                "ubicacion": rack,
                "foto": foto,
                "guia": guia,
                "proveedor": prov,
                "materiales": matching_materials,
            })
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_edit_pallet(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo de solicitud vacío"}, status=400)
                return

            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))

            codigo = str(data.get("codigo", "")).strip().upper()
            old_codigo = str(data.get("old_codigo", "")).strip().upper()
            pallet = str(data.get("pallet", "")).strip()
            ubicacion = str(data.get("ubicacion", "")).strip()
            cantidad = data.get("cantidad_pallet", 0)
            guia = str(data.get("guia", "")).strip()
            proveedor = str(data.get("proveedor", "")).strip()
            descripcion = str(data.get("descripcion", "")).strip()

            if not codigo:
                self.send_json({"success": False, "error": "El código del material es obligatorio"}, status=400)
                return

            mapping = gen.load_photo_mapping()
            existing = {}
            if old_codigo and old_codigo in mapping and old_codigo != codigo:
                existing = mapping.pop(old_codigo)
            elif codigo in mapping:
                existing = mapping.get(codigo, {})

            foto = data.get("foto") or existing.get("foto", "")
            if foto.startswith("FOTOS/"):
                foto = foto[6:]

            try:
                qty_num = float(cantidad) if (cantidad is not None and str(cantidad).strip() != "") else existing.get("cantidad_pallet", 0)
            except Exception:
                qty_num = 0

            target_pallet_name = str(data.get("old_pallet") or pallet).strip().lower()
            existing_pallets = existing.get("pallets")
            if not isinstance(existing_pallets, list):
                if existing.get("foto") or existing.get("pallet"):
                    existing_pallets = [{
                        "pallet": existing.get("pallet", ""),
                        "ubicacion": existing.get("ubicacion", ""),
                        "foto": existing.get("foto", ""),
                        "descripcion": existing.get("descripcion", ""),
                        "guia": existing.get("guia", ""),
                        "cantidad_pallet": existing.get("cantidad_pallet", 0),
                        "proveedor": existing.get("proveedor", ""),
                        "fecha_registro": existing.get("fecha_registro", ""),
                    }]
                else:
                    existing_pallets = []

            for p in existing_pallets:
                if p.get("foto", "").startswith("FOTOS/"):
                    p["foto"] = p["foto"][6:]

            updated_in_list = False
            for idx_p, ep in enumerate(existing_pallets):
                if ep.get("pallet", "").strip().lower() == target_pallet_name:
                    ep.update({
                        "pallet": pallet or ep.get("pallet", "Pallet"),
                        "ubicacion": ubicacion or ep.get("ubicacion", ""),
                        "foto": foto or ep.get("foto", ""),
                        "descripcion": descripcion or ep.get("descripcion", ""),
                        "guia": guia or ep.get("guia", ""),
                        "cantidad_pallet": qty_num,
                        "proveedor": proveedor or ep.get("proveedor", ""),
                        "fecha_edicion": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    })
                    updated_in_list = True
                    break

            if not updated_in_list:
                existing_pallets.append({
                    "pallet": pallet or "Pallet",
                    "ubicacion": ubicacion or "",
                    "foto": foto or "",
                    "descripcion": descripcion or "",
                    "guia": guia or "",
                    "cantidad_pallet": qty_num,
                    "proveedor": proveedor or "",
                    "fecha_edicion": datetime.now().strftime("%d/%m/%Y %H:%M"),
                })

            mapping[codigo] = {
                "pallets": existing_pallets,
                "pallet": pallet or existing.get("pallet", "Pallet"),
                "ubicacion": ubicacion or existing.get("ubicacion", ""),
                "foto": foto,
                "descripcion": descripcion or existing.get("descripcion", ""),
                "guia": guia or existing.get("guia", ""),
                "cantidad_pallet": qty_num,
                "proveedor": proveedor or existing.get("proveedor", ""),
                "fecha_registro": existing.get("fecha_registro", datetime.now().strftime("%d/%m/%Y %H:%M")),
                "fecha_edicion": datetime.now().strftime("%d/%m/%Y %H:%M"),
            }

            MAPEO_FOTOS_FILE.write_text(
                json.dumps(mapping, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

            rebuild_inventory_async()

            self.send_json({
                "success": True,
                "message": f"¡Pallet {pallet} ({codigo}) actualizado exitosamente!",
                "pallet_data": mapping[codigo]
            })
            print(f"[OK] Pallet {pallet} ({codigo}) editado y guardado exitosamente.")

        except Exception as e:
            print(f"[ERROR] Error procesando edición de pallet: {e}")
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_delete_pallet(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length <= 0:
                self.send_json({"success": False, "error": "Cuerpo de solicitud vacío"}, status=400)
                return

            body_bytes = self.rfile.read(content_length)
            data = json.loads(body_bytes.decode("utf-8"))
            codigo = str(data.get("codigo", "")).strip().upper()
            pallet = str(data.get("pallet", "")).strip()

            if not codigo:
                self.send_json({"success": False, "error": "El código es obligatorio"}, status=400)
                return

            mapping = gen.load_photo_mapping()
            target_key = None
            if codigo and codigo in mapping:
                target_key = codigo
            elif pallet:
                for k, v in mapping.items():
                    if str(v.get("pallet", "")).strip().lower() == pallet.strip().lower():
                        target_key = k
                        break

            # Limpiar archivo físico si existe
            deleted_file = False
            for cand_name in [f"{pallet}.jpeg", f"{pallet}.jpg", f"{pallet.upper()}.jpeg", f"{pallet.lower()}.jpeg"]:
                p_cand = FOTOS_DIR / cand_name
                if p_cand.exists():
                    try:
                        p_cand.unlink()
                        deleted_file = True
                        print(f"[OK] Archivo de foto eliminado directamente: {p_cand}")
                    except Exception as e_del:
                        print(f"[WARN] Error eliminando {p_cand}: {e_del}")

            if target_key and target_key in mapping:
                info = mapping.get(target_key, {})
                existing_pallets = info.get("pallets")
                if isinstance(existing_pallets, list) and pallet:
                    # Remove only the specified pallet from list
                    remaining = [p for p in existing_pallets if p.get("pallet", "").strip().lower() != pallet.strip().lower()]
                    if remaining:
                        mapping[target_key]["pallets"] = remaining
                        mapping[target_key]["pallet"] = remaining[-1].get("pallet", "")
                        mapping[target_key]["ubicacion"] = remaining[-1].get("ubicacion", "")
                        mapping[target_key]["foto"] = remaining[-1].get("foto", "")
                        mapping[target_key]["cantidad_pallet"] = remaining[-1].get("cantidad_pallet", 0)
                    else:
                        mapping.pop(target_key, None)
                else:
                    mapping.pop(target_key, None)

                foto_filename = info.get("foto", "")
                if foto_filename and not mapping.get(target_key):
                    clean_fname = Path(foto_filename).name
                    foto_path = FOTOS_DIR / clean_fname
                    try:
                        if foto_path.exists():
                            foto_path.unlink()
                            deleted_file = True
                            print(f"[OK] Archivo de foto eliminado: {foto_path}")
                    except Exception as e_f:
                        print(f"[WARN] Error eliminando archivo de foto {foto_path}: {e_f}")

                MAPEO_FOTOS_FILE.write_text(
                    json.dumps(mapping, indent=2, ensure_ascii=False),
                    encoding="utf-8",
                )

            rebuild_inventory_async()

            self.send_json({
                "success": True,
                "message": f"¡Pallet {pallet or codigo} eliminado exitosamente y foto borrada!",
                "codigo": codigo or target_key,
            })
            print(f"[OK] Pallet {pallet} ({codigo or target_key}) y su fotografía eliminados exitosamente.")

        except Exception as e:
            print(f"[ERROR] Error procesando eliminación de pallet: {e}")
            self.send_json({"success": False, "error": str(e)}, status=500)

    def send_json(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        if "gzip" in self.headers.get("Accept-Encoding", "") and len(body) > 1024:
            import gzip
            body = gzip.compress(body, compresslevel=6)
            self.send_header("Content-Encoding", "gzip")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def start_fotos_folder_watcher() -> None:
    """Continuously monitors FOTOS directory and processes new photos automatically."""
    def _watcher():
        import importlib
        import procesar_fotos_bodega
        while True:
            time.sleep(15)
            try:
                importlib.reload(procesar_fotos_bodega)
                mapping = gen.load_photo_mapping()
                processed = procesar_fotos_bodega.get_existing_processed_files(mapping)
                all_images = [p for p in FOTOS_DIR.glob("*.*") if p.suffix.lower() in (".jpg", ".jpeg", ".png")]
                pending = [p for p in all_images if p.name.lower() not in processed]
                if pending:
                    print(f"[WATCHER] ¡Detectadas {len(pending)} fotos nuevas en FOTOS! Procesando automáticamente...")
                    procesar_fotos_bodega.process_all_photos(force_reprocess=False)
                    global LAST_SERVER_UPDATE
                    LAST_SERVER_UPDATE = time.time()
                    print("[WATCHER] ¡Fotos procesadas y catálogo actualizado exitosamente!")
            except Exception as e_watch:
                print(f"[WATCHER] Error en monitor de fotos: {e_watch}")

    t = threading.Thread(target=_watcher, daemon=True)
    t.start()


def run_server(port: int = DEFAULT_PORT) -> None:
    local_ip = get_local_ip()
    update_launcher_files("", local_ip, port)
    start_cloudflare_tunnel(port)
    # start_fotos_folder_watcher()  # Deshabilitado: el usuario solicitó no subir/procesar fotos automáticamente de carpetas

    if port != 80:
        def _port80_thread():
            try:
                ThreadingHTTPServer.allow_reuse_address = True
                httpd80 = ThreadingHTTPServer(("0.0.0.0", 80), BodegaHTTPRequestHandler)
                print("[OK] Servidor escuchando también en puerto 80 estándar (http://192.168.1.18/)")
                httpd80.serve_forever()
            except Exception as e_80:
                print(f"[INFO] Puerto 80 no disponible o en uso: {e_80}")

        threading.Thread(target=_port80_thread, daemon=True).start()

    ThreadingHTTPServer.allow_reuse_address = True
    server_address = ("0.0.0.0", port)
    try:
        httpd = ThreadingHTTPServer(server_address, BodegaHTTPRequestHandler)
    except OSError as e_sock:
        if "10048" in str(e_sock) or "Address already in use" in str(e_sock):
            print(f"[WARN] Puerto {port} ocupado, intentando liberar proceso anterior...")
            try:
                creationflags = 0x08000000 if sys.platform == "win32" else 0
                my_pid = os.getpid()
                subprocess.run(
                    ["powershell", "-Command", f"Get-NetTCPConnection -LocalPort {port} -State Listen -ErrorAction SilentlyContinue | Where-Object {{ $_.OwningProcess -ne {my_pid} }} | ForEach-Object {{ Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }}"],
                    timeout=5,
                    creationflags=creationflags,
                )
                time.sleep(1.5)
                httpd = ThreadingHTTPServer(server_address, BodegaHTTPRequestHandler)
            except Exception:
                raise e_sock
        else:
            raise e_sock
    print("=" * 68)
    print("  SERVIDOR BODEGA CENTRAL — DEPTO. ELÉCTRICO GUZMÁN Y LARRAÍN")
    print("=" * 68)
    print(f"  Acceso Local PC:       http://localhost:{port}/")
    print(f"  Acceso Wi-Fi Bodega:   http://{local_ip}:{port}/")
    print("  Conectando túnel HTTPS para acceso celular...")
    print("=" * 68)
    print("  Presiona Ctrl+C para detener el servidor.")
    print("=" * 68)
    while True:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n[INFO] KeyboardInterrupt ignored in server loop. Still running.")
            time.sleep(0.5)
        except Exception as e:
            print(f"[WARN] Error en bucle del servidor HTTP: {e}")
            time.sleep(1)
    httpd.server_close()
    cleanup()


if __name__ == "__main__":
    p = DEFAULT_PORT
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        p = int(sys.argv[1])
    run_server(p)
