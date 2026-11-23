#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aislador de proceso para Windows Native OCR (WinRT).

Ejecuta el reconocimiento OCR en un subproceso aislado e independiente, evitando que
posibles errores de C++/destructores de WinRT en Python 3.12 afecten o boten el servidor HTTP.
"""

import asyncio
import json
import os
import sys
from pathlib import Path
from PIL import Image

try:
    import winocr
except ImportError:
    winocr = None


async def run_ocr(image_path: str) -> list[str]:
    if not winocr:
        return []
    try:
        img = Image.open(image_path)
        res = await winocr.recognize_pil(img, lang="es")
        return [l.text.strip() for l in res.lines if l.text.strip()]
    except Exception as e:
        sys.stderr.write(f"[OCR_WORKER_ERROR] {e}\n")
        return []


def main():
    if len(sys.argv) < 2:
        print("[]")
        return

    image_path = sys.argv[1]
    if not os.path.exists(image_path):
        print("[]")
        return

    try:
        lines = asyncio.run(run_ocr(image_path))
        print(json.dumps(lines, ensure_ascii=True))
    except Exception as e:
        sys.stderr.write(f"[OCR_WORKER_FATAL] {e}\n")
        print("[]")


if __name__ == "__main__":
    main()
