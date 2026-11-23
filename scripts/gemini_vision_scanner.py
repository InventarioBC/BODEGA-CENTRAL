#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Módulo de Reconocimiento Visual Multimodal con Google Gemini Vision.

Permite que la aplicación lea con máxima precisión TODOS los códigos de materiales,
cantidades, número de guía de despacho, proveedor y ubicación en rack presentes
en fotografías tomadas desde el celular o subidas a la bodega central.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

from google import genai
from google.genai import types

SCRIPTS_DIR = Path(__file__).resolve().parent
KEY_FILE = SCRIPTS_DIR / "gemini_api_key.txt"
USER_KEY_FILE = Path(os.path.expanduser("~")) / ".gemini_api_key"


def get_gemini_api_key() -> str:
    """Recupera la clave API de Gemini desde el archivo local o variable de entorno."""
    if KEY_FILE.exists():
        try:
            k = KEY_FILE.read_text(encoding="utf-8").strip()
            if k and not k.startswith("#"):
                return k
        except Exception:
            pass

    if USER_KEY_FILE.exists():
        try:
            k = USER_KEY_FILE.read_text(encoding="utf-8").strip()
            if k and not k.startswith("#"):
                return k
        except Exception:
            pass

    env_k = os.environ.get("GEMINI_API_KEY", "").strip()
    if env_k:
        return env_k

    return ""


def save_gemini_api_key(api_key: str) -> bool:
    """Guarda la clave API de Gemini para que quede permanente."""
    k = api_key.strip()
    if not k:
        return False
    try:
        KEY_FILE.write_text(k, encoding="utf-8")
        try:
            USER_KEY_FILE.write_text(k, encoding="utf-8")
        except Exception:
            pass
        os.environ["GEMINI_API_KEY"] = k
        return True
    except Exception as e:
        print(f"[WARN] No se pudo guardar la clave API: {e}")
        return False


def is_gemini_available() -> bool:
    """Verifica si la clave API de Gemini está configurada."""
    return bool(get_gemini_api_key())


def clean_json_response(text: str) -> dict[str, Any]:
    """Limpia el texto devuelto por el modelo para extraer un JSON válido."""
    s = text.strip()
    # Si viene con bloque markdown ```json ... ```
    if "```" in s:
        match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", s)
        if match:
            s = match.group(1).strip()
    try:
        return json.loads(s)
    except Exception:
        # Intentar extraer primer bloque entre { y }
        m_brace = re.search(r"(\{[\s\S]*\})", s)
        if m_brace:
            return json.loads(m_brace.group(1))
        raise


def scan_pallet_with_gemini(
    img_bytes: bytes,
    catalog: list[dict[str, Any]],
    next_pallet_name: str = "Pallet 1",
    api_key: str | None = None
) -> dict[str, Any]:
    """Analiza la imagen usando Google Gemini Vision y extrae todos los códigos y metadatos."""
    key = api_key or get_gemini_api_key()
    if not key:
        return {
            "success": False,
            "error": "Clave API de Gemini no configurada",
            "gemini_used": False
        }

    # Resumen conciso del catálogo para que Gemini haga matching exacto
    catalog_lines = []
    catalog_map = {}
    for it in catalog:
        c = str(it.get("c", "")).strip().upper()
        d = str(it.get("d", "")).strip().upper()
        u = str(it.get("u", "")).strip().upper()
        if c and d:
            catalog_lines.append(f"- {c}: {d} ({u})")
            catalog_map[c] = {"c": c, "d": d, "u": u}

    catalog_context = "\n".join(catalog_lines[:250])  # Top catálogo relevante

    system_prompt = f"""Eres el Asistente de Bodega y Especialista en Abastecimiento del Departamento Eléctrico de Guzmán y Larraín.
Tu tarea es examinar con máxima precisión esta fotografía tomada en la bodega central.

En la imagen hay un pallet, estante o bulto que puede contener UNO, DOS, TRES O MÁS CÓDIGOS DE MATERIALES DISTINTOS (por ejemplo, varias bobinas/rollos de cable de diferentes colores como Rojo, Verde, Blanco, rollos de cinta aisladora, cajas o luminarias).

Debes buscar y leer meticulosamente:
1. TODAS las etiquetas adhesivas de productos (fabricante, códigos de barra, rótulos de cables, estampados sobre el plástico o el cartón).
2. La ficha o letrero pegado en el pallet (donde suele decir BODEGA CENTRAL, GUIA, CODIGO, CANTIDAD, PROVEEDOR).
3. La ubicación en el estante/rack (placa azul rectangular con letras blancas en la viga amarilla, ej: "6B3", "2A9", "5A2", "4A2").
4. El número de Guía de Despacho (ej: 76110, 76359, 74700, 74698, etc.).
5. El proveedor del material (ej: "CBM", "TECNORED", "SAIME", "DARTEL", "RHONA", etc.).

Catálogo de referencia de Guzmán y Larraín:
{catalog_context}

REGLAS DE EXTRACCIÓN:
- Si hay varios rollos o cajas con distintos códigos (ej: rollos de cable 1.5mm rojo, verde y blanco en el mismo pallet), DEBES INCLUIR TODOS Y CADA UNO en la lista 'materiales'.
- Asigna a cada uno su código exacto del catálogo (ej: 1UA0165 para rojo 1.5mm, 1UA0169 para verde 1.5mm, 1UA0168 para blanco 1.5mm; o 1UA0170 para rojo 2.5mm, 1UA0172 para verde 2.5mm, 1UA0154 para blanco 2.5mm; 1UB3301 para Cinta Temflex, etc.).
- Si una cantidad no aparece explícita pero se deduce (ej. rollos de cable de 5.000 mt, 8.000 mt o 15.000 mt), pon la cantidad estimada o 0 si se desconoce.
- Formato del rack: Siempre pon "Rack <numero><letra><nivel>" (ej: "Rack 6B3", "Rack 2A9", "Rack 5A2").

Responde ÚNICAMENTE con un objeto JSON válido con esta estructura exacta (sin texto introductorio ni explicaciones):
{{
  "rack": "Rack 6B3",
  "guia": "76110",
  "proveedor": "CBM",
  "materiales": [
    {{
      "codigo": "1UA0165",
      "descripcion": "CABLE UNIPOLAR H07Z1-K EVA REVESTIMIENTO FREETOX -FLEX 1,5 MM ROJO",
      "cantidad": 15000,
      "unidad": "ML"
    }},
    {{
      "codigo": "1UA0169",
      "descripcion": "CABLE UNIPOLAR H07Z1-K EVA REVESTIMIENTO FREETOX -FLEX 1,5 MM VERDE",
      "cantidad": 8000,
      "unidad": "ML"
    }}
  ]
}}
"""

    client = genai.Client(api_key=key)

    # Intentar con modelos recomendados de Gemini
    models_to_try = ["gemini-3.1-flash-lite", "gemini-3.5-flash", "gemini-3.8-flash", "gemini-3.7-flash"]
    last_err = None

    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[
                    system_prompt,
                    types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"),
                ],
            )
            raw_text = response.text or ""
            parsed = clean_json_response(raw_text)

            detected_rack = parsed.get("rack", "").strip()
            detected_guia = str(parsed.get("guia", "")).strip()
            detected_prov = parsed.get("proveedor", "").strip()
            raw_mats = parsed.get("materiales", [])

            processed_mats = []
            for m in raw_mats:
                c_code = str(m.get("codigo", "")).strip().upper()
                c_desc = str(m.get("descripcion", "")).strip()
                c_qty = float(m.get("cantidad", 0) or 0)

                # Cruzar con catálogo oficial para enriquecer descripción
                if c_code in catalog_map:
                    c_desc = catalog_map[c_code]["d"]

                if c_code:
                    processed_mats.append({
                        "codigo": c_code,
                        "descripcion": c_desc,
                        "cantidad": c_qty,
                        "score": 500  # Máxima confianza de Gemini Vision
                    })

            best_item = processed_mats[0] if processed_mats else None
            best_code = best_item["codigo"] if best_item else ""
            best_desc = best_item["descripcion"] if best_item else ""
            best_qty = best_item["cantidad"] if best_item else 0

            return {
                "success": True,
                "matched": (len(processed_mats) > 0) or bool(detected_rack or detected_guia or detected_prov),
                "codigo": best_code,
                "descripcion": best_desc,
                "pallet": next_pallet_name,
                "ubicacion": detected_rack,
                "guia": detected_guia,
                "cantidad": best_qty,
                "proveedor": detected_prov,
                "matched_by": f"Google Gemini Vision ({model_name})",
                "confidence_score": 500,
                "detected_text": f"Gemini Vision detectó {len(processed_mats)} material(es)",
                "suggestions": [],
                "other_options": processed_mats[1:] if len(processed_mats) > 1 else [],
                "multi_materials": processed_mats,
                "gemini_used": True,
                "model_used": model_name
            }
        except Exception as e:
            last_err = e
            print(f"[WARN] Error con {model_name}: {e}")
            continue

    return {
        "success": False,
        "error": f"Error comunicando con Gemini Vision: {last_err}",
        "gemini_used": False
    }
