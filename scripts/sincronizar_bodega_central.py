import subprocess
import sys
import webbrowser
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generar_buscador_bodega as gen

ICONSTRUYE_EXTRACTOR = Path(r"C:\Bodega\scripts\iconstruye_bodega_central.py")


def main():
    print("===============================================================")
    print("  SINCRONIZACION BODEGA CENTRAL (56 & 121) DESDE ICONSTRUYE    ")
    print("===============================================================\n")

    if ICONSTRUYE_EXTRACTOR.exists():
        print("[1/2] Descargando inventario actualizado desde iConstruye...")
        res = subprocess.run([sys.executable, str(ICONSTRUYE_EXTRACTOR)], check=False)
        if res.returncode != 0:
            print("\n[AVISO] No se pudo completar la descarga en linea desde iConstruye.")
            print("        Se regenerara el buscador con los ultimos archivos locales disponibles.")
    else:
        print(f"[AVISO] No se encontro {ICONSTRUYE_EXTRACTOR}. Usando archivos locales.")

    print("\n[2/2] Consolidando Bodega 56 + Bodega 121 y generando Buscador Web + Excel...")
    out_html = gen.build_all()

    print("\nListo. Abriendo Buscador de Bodega Central en tu navegador...")
    if out_html.exists():
        webbrowser.open(out_html.as_uri())


if __name__ == "__main__":
    main()
