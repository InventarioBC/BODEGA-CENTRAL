import sys
import webbrowser
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generar_buscador_bodega as gen


def main():
    out_html = gen.OUT_HTML
    try:
        inv_sources, mov_sources = gen.collect_latest_sources()
        all_paths = list(inv_sources.values()) + list(mov_sources.values())
        needs_rebuild = not out_html.exists()
        if not needs_rebuild and all_paths:
            html_mtime = out_html.stat().st_mtime
            latest_src = max(p.stat().st_mtime for p in all_paths)
            if latest_src > html_mtime:
                needs_rebuild = True
        if needs_rebuild and inv_sources:
            gen.build_all()
    except Exception:
        pass

    if out_html.exists():
        webbrowser.open(out_html.as_uri())


if __name__ == "__main__":
    main()
