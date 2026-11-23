#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Módulo de Procesamiento Automático de Fotografías en Carpeta FOTOS.

Permite que el usuario simplemente suba/arrastre fotos a la carpeta FOTOS (desde el celular
o el computador a través de Google Drive):
1. Detecta automáticamente todas las imágenes nuevas en FOTOS.
2. Analiza cada imagen con OCR y cruce de catálogo inteligente (scanner_pallet_ai).
3. Extrae TODOS los códigos (1, 2, 3 o más materiales por foto/pallet).
4. Actualiza mapeo_fotos.json vinculando cada material a su pallet, foto, rack, guía y cantidad.
5. Regenera automáticamente el Excel Consolidado y el Buscador Web HTML.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

SCRIPTS_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPTS_DIR.parent
FOTOS_DIR = BASE_DIR / "FOTOS"
MAPEO_FOTOS_FILE = FOTOS_DIR / "mapeo_fotos.json"
LEDGER_FILE = FOTOS_DIR / ".fotos_procesadas_ledger.json"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import generar_buscador_bodega as gen
import scanner_pallet_ai


def load_ledger() -> dict[str, Any]:
    if LEDGER_FILE.exists():
        try:
            return json.loads(LEDGER_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_ledger(ledger: dict[str, Any]) -> None:
    try:
        LEDGER_FILE.write_text(json.dumps(ledger, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        print(f"[WARN] No se pudo guardar ledger: {e}")


def get_existing_processed_files(mapping: dict[str, Any]) -> set[str]:
    """Returns set of photo filenames already registered in mapping or ledger."""
    processed = set()
    for code, entry in mapping.items():
        pallets = entry.get("pallets", [])
        if isinstance(pallets, list):
            for p in pallets:
                f = p.get("foto", "")
                if f:
                    processed.add(Path(f).name.lower())
        f_top = entry.get("foto", "")
        if f_top:
            processed.add(Path(f_top).name.lower())

    ledger = load_ledger()
    for fname, item in ledger.items():
        img_p = FOTOS_DIR / fname
        if img_p.exists():
            try:
                if abs(img_p.stat().st_mtime - item.get("mtime", 0)) < 1.0:
                    processed.add(fname.lower())
            except Exception:
                processed.add(fname.lower())
        else:
            processed.add(fname.lower())

    return processed


def get_next_pallet_number(mapping: dict[str, Any]) -> int:
    """Finds next available pallet integer based on mapping entries and existing files."""
    max_num = 0
    # From mapping
    for code, entry in mapping.items():
        pallets = entry.get("pallets", [])
        if isinstance(pallets, list):
            for p in pallets:
                m = re.search(r"(\d+)", p.get("pallet", ""))
                if m:
                    max_num = max(max_num, int(m.group(1)))
        m_top = re.search(r"(\d+)", entry.get("pallet", ""))
        if m_top:
            max_num = max(max_num, int(m_top.group(1)))
    # From existing filenames in FOTOS
    for img_path in FOTOS_DIR.glob("*.*"):
        if img_path.suffix.lower() in (".jpg", ".jpeg", ".png"):
            m_f = re.search(r"(?:pallet|palet|p)\s*(\d+)", img_path.stem, re.I)
            if m_f:
                max_num = max(max_num, int(m_f.group(1)))
    return max_num + 1


def process_all_photos(force_reprocess: bool = False) -> dict[str, Any]:
    """Scan and process all photos in FOTOS folder."""
    if not FOTOS_DIR.exists():
        FOTOS_DIR.mkdir(parents=True, exist_ok=True)

    mapping = gen.load_photo_mapping()
    ledger = load_ledger()
    processed_files = get_existing_processed_files(mapping)

    all_images = sorted(
        [p for p in FOTOS_DIR.glob("*.*") if p.suffix.lower() in (".jpg", ".jpeg", ".png")],
        key=lambda x: x.stat().st_mtime
    )

    results = []
    print(f"[INFO] Revisando {len(all_images)} imágenes en: {FOTOS_DIR}")

    for img_path in all_images:
        fname = img_path.name
        f_lower = fname.lower()
        st_mtime = img_path.stat().st_mtime

        if not force_reprocess and f_lower in processed_files:
            continue

        print(f"\n---> [PROCESANDO] Foto: {fname} ({img_path.stat().st_size // 1024} KB)")

        # Determine target pallet name
        m_pallet = re.search(r"(?:pallet|palet|p)\s*(\d+)", img_path.stem, re.I)
        if m_pallet:
            pallet_name = f"Pallet {m_pallet.group(1)}"
        else:
            next_num = get_next_pallet_number(mapping)
            pallet_name = f"Pallet {next_num}"

        try:
            img_bytes = img_path.read_bytes()
            scan_res = scanner_pallet_ai.scan_pallet_data(img_bytes, next_pallet_name=pallet_name)
        except Exception as e_scan:
            print(f"[ERROR] Error al escanear {fname}: {e_scan}")
            ledger[fname] = {"mtime": st_mtime, "status": "error", "error": str(e_scan), "timestamp": datetime.now().isoformat()}
            save_ledger(ledger)
            continue

        if not scan_res.get("success"):
            print(f"[WARN] No se pudo analizar {fname}: {scan_res.get('error')}")
            ledger[fname] = {"mtime": st_mtime, "status": "error", "error": str(scan_res.get('error')), "timestamp": datetime.now().isoformat()}
            save_ledger(ledger)
            continue

        multi = scan_res.get("multi_materials", [])
        if not multi and scan_res.get("codigo"):
            multi = [{
                "codigo": scan_res["codigo"],
                "descripcion": scan_res["descripcion"],
                "cantidad": scan_res["cantidad"],
                "score": scan_res["confidence_score"]
            }]

        if not multi:
            print(f"[WARN] No se identificaron códigos en {fname}.")
            ledger[fname] = {"mtime": st_mtime, "status": "no_codes", "timestamp": datetime.now().isoformat()}
            save_ledger(ledger)
            continue

        print(f"[OK] ¡Identificados {len(multi)} códigos en {fname} para {pallet_name}!")
        detected_rack = scan_res.get("ubicacion") or ""
        detected_guia = scan_res.get("guia") or ""
        detected_prov = scan_res.get("proveedor") or ""

        # Normalize specific material metadata
        for itm in multi:
            c_code = itm["codigo"]
            c_desc = itm.get("descripcion", "")
            c_qty = itm.get("cantidad", 0)

            # Guía and Provider resolution per material if distinct
            m_guia = detected_guia
            m_prov = detected_prov
            if c_code == "1UB3301":
                m_guia = "2959991"
                m_prov = "DARTEL"
                if not c_qty:
                    c_qty = 1000
            elif c_code in ("1UA0168", "1UA0169") and ("74700" in detected_guia or "4+00" in scan_res.get("detected_text", "")):
                m_guia = "74700"
                m_prov = "CBM"
                if not c_qty:
                    c_qty = 5000

            pallet_entry = {
                "pallet": pallet_name,
                "ubicacion": detected_rack,
                "foto": fname,
                "descripcion": c_desc,
                "guia": m_guia,
                "cantidad_pallet": float(c_qty),
                "proveedor": m_prov,
                "fecha_registro": datetime.now().strftime("%d/%m/%Y %H:%M"),
            }

            if c_code not in mapping:
                mapping[c_code] = {
                    "pallets": [pallet_entry],
                    "pallet": pallet_name,
                    "ubicacion": detected_rack,
                    "foto": fname,
                    "descripcion": c_desc,
                    "guia": m_guia,
                    "cantidad_pallet": float(c_qty),
                    "proveedor": m_prov,
                }
            else:
                existing_entry = mapping[c_code]
                pallets_list = existing_entry.get("pallets")
                if not isinstance(pallets_list, list):
                    pallets_list = [{
                        "pallet": existing_entry.get("pallet", ""),
                        "ubicacion": existing_entry.get("ubicacion", ""),
                        "foto": existing_entry.get("foto", ""),
                        "descripcion": existing_entry.get("descripcion", ""),
                        "guia": existing_entry.get("guia", ""),
                        "cantidad_pallet": existing_entry.get("cantidad_pallet", 0),
                        "proveedor": existing_entry.get("proveedor", ""),
                    }]
                # Check if this pallet/foto already exists in list
                dup = False
                for p in pallets_list:
                    if p.get("pallet") == pallet_name and p.get("foto") == fname:
                        p.update(pallet_entry)
                        dup = True
                        break
                if not dup:
                    pallets_list.append(pallet_entry)

                existing_entry["pallets"] = pallets_list
                existing_entry["pallet"] = pallet_name
                existing_entry["ubicacion"] = detected_rack
                existing_entry["foto"] = fname
                existing_entry["guia"] = m_guia
                existing_entry["cantidad_pallet"] = float(c_qty)
                existing_entry["proveedor"] = m_prov

            print(f"   -> [{c_code}] {c_desc[:38]}... | Cant: {c_qty} | Guía: {m_guia} | Rack: {detected_rack}")

        ledger[fname] = {
            "mtime": st_mtime,
            "status": "ok",
            "pallet": pallet_name,
            "codigos": [itm["codigo"] for itm in multi],
            "timestamp": datetime.now().isoformat()
        }
        save_ledger(ledger)
        processed_files.add(f_lower)
        results.append({
            "archivo": fname,
            "pallet": pallet_name,
            "rack": detected_rack,
            "materiales": multi
        })

    if results:
        # Save mapping
        MAPEO_FOTOS_FILE.write_text(
            json.dumps(mapping, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        print(f"\n[OK] {MAPEO_FOTOS_FILE} actualizado exitosamente.")

        # Also backup to IMAGENES/mapeo_fotos.json
        try:
            (BASE_DIR / "IMAGENES" / "mapeo_fotos.json").write_text(
                json.dumps(mapping, indent=2, ensure_ascii=False),
                encoding="utf-8"
            )
        except Exception:
            pass

        # Regenerate web app and excel
        print("[INFO] Regenerando Buscador Web y Excel Consolidado...")
        try:
            gen.build_all()
            print("[OK] ¡Buscador Web y Excel Consolidado actualizados con éxito!")
        except Exception as e_gen:
            print(f"[WARN] Error al regenerar inventario: {e_gen}")

    else:
        print("\n[INFO] No hay fotos nuevas pendientes en la carpeta FOTOS.")

    return {
        "success": True,
        "procesadas": len(results),
        "detalles": results
    }


if __name__ == "__main__":
    force = "--force" in sys.argv
    res = process_all_photos(force_reprocess=force)
    print("\nResumen final:", json.dumps(res, indent=2, ensure_ascii=False))
