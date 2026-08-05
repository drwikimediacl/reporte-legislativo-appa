#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un Excel con el detalle de cada proyecto de la watchlist,
combinando datos oficiales del webservice del Senado con las
columnas manuales (tema/importancia) de la planilla de Google.
Añade columnas de último trámite usando los datos de monitor-legislativo/db.json
"""
import re
import time
import os
import json
import pandas as pd
import requests
import xml.etree.ElementTree as ET

WATCHLIST_URL = "https://docs.google.com/spreadsheets/d/1RBS_VB7d3jyJvJL87gBZg5hx8I2PKslBfd3kWAxVdgA/export?format=csv"
SENADO_WS = "https://tramitacion.senado.cl/wspublico/tramitacion.php?boletin={numero}"
OUTPUT_XLSX = "proyectos_legislativos.xlsx"

# Fallback raw URL (public) al db.json en el repo monitor-legislativo
MONITOR_DB_RAW_URL = "https://raw.githubusercontent.com/drwikimediacl/monitor-legislativo/main/db.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


def cargar_watchlist() -> pd.DataFrame:
    df = pd.read_csv(WATCHLIST_URL)
    df.columns = [c.strip().lower() for c in df.columns]
    df = df.dropna(subset=["boletin"])
    df["boletin"] = df["boletin"].astype(str).str.strip()
    return df


def consultar_senado(boletin: str) -> dict:
    numero = boletin.split("-")[0]
    url = SENADO_WS.format(numero=numero)
    try:
        r = requests.get(url, timeout=15, headers=HEADERS)
        r.raise_for_status()
        root = ET.fromstring(r.content)
        proyecto = root.find(".//proyecto")
        if proyecto is None:
            return {}

        desc = proyecto.find("descripcion")
        titulo = desc.findtext("titulo", "").strip() if desc is not None else ""
        fecha_ingreso = desc.findtext("fecha_ingreso", "").strip() if desc is not None else ""
        iniciativa = desc.findtext("iniciativa", "").strip() if desc is not None else ""
        etapa = desc.findtext("etapa", "").strip() if desc is not None else ""
        subetapa = desc.findtext("subetapa", "").strip() if desc is not None else ""
        estado = desc.findtext("estado", "").strip() if desc is not None else ""

        autores = [
            a.findtext("PARLAMENTARIO", "").strip()
            for a in proyecto.findall(".//autores/autor")
        ]
        autores = [a for a in autores if a]

        anio = ""
        m = re.search(r"\d{2}/\d{2}/(\d{4})", fecha_ingreso)
        if m:
            anio = m.group(1)

        comision = subetapa if "omisi" in subetapa.lower() else etapa

        return {
            "titulo": titulo,
            "anio": anio,
            "iniciativa": iniciativa,
            "autores": ", ".join(autores),
            "comision": comision,
            "estado": f"{estado} - {etapa}".strip(" -"),
        }
    except Exception as e:
        print(f"  ⚠️ Error consultando boletín {numero}: {e}")
        return {}


def load_monitor_db() -> dict:
    local_path = "db.json"
    if os.path.exists(local_path):
        try:
            with open(local_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                print(f"DEBUG: cargado db.json local con {len(data)} entradas")
                return data
        except Exception as e:
            print(f"  ⚠️ Error leyendo {local_path}: {e}")

    try:
        r = requests.get(MONITOR_DB_RAW_URL, timeout=15, headers=HEADERS)
        r.raise_for_status()
        data = r.json()
        print(f"DEBUG: descargado db.json remoto con {len(data)} entradas")
        return data
    except Exception as e:
        print(f"  ⚠️ No pude cargar db.json remoto ({MONITOR_DB_RAW_URL}): {e}")
        return {}


def find_monitor_entry(monitor_db: dict, boletin: str) -> dict:
    if not monitor_db:
        return {}
    if boletin in monitor_db:
        return monitor_db[boletin]
    for k, v in monitor_db.items():
        if k.startswith(boletin):
            return v
    for k, v in monitor_db.items():
        if boletin in k:
            return v
    return {}


def format_ultimo_tramite(entry: dict) -> (str, str):
    if not entry:
        return "", ""
    fecha = entry.get("ultimo_tramite_fecha", "") or ""
    descripcion = entry.get("ultimo_tramite_descripcion", "") or ""
    etapa = entry.get("ultimo_tramite_etapa", "") or ""
    parts = []
    if fecha:
        parts.append(fecha)
    if etapa:
        parts.append(etapa)
    if descripcion:
        parts.append(descripcion)
    descripcion_form = " — ".join(parts) if parts else ""
    if len(descripcion_form) > 800:
        descripcion_form = descripcion_form[:800].rsplit(" ", 1)[0] + "…"
    return fecha, descripcion_form


def main():
    print("Descargando watchlist desde Google Sheets...")
    watchlist = cargar_watchlist()

    print("Cargando base de datos de monitor-legislativo (db.json)...")
    monitor_db = load_monitor_db()

    print("DEBUG: filas en watchlist =", len(watchlist))

    filas = []
    for _, row in watchlist.iterrows():
        boletin = row["boletin"]
        print(f"Consultando boletín {boletin}...")
        datos = consultar_senado(boletin)
        time.sleep(0.5)

        monitor_entry = find_monitor_entry(monitor_db, boletin)
        fecha_ultimo, desc_ultimo = format_ultimo_tramite(monitor_entry)

        filas.append({
            "Nombre proyecto": datos.get("titulo") or row.get("nombre", ""),
            "Boletín": boletin,
            "Año presentación": datos.get("anio", ""),
            "Moción/Mensaje": datos.get("iniciativa", ""),
            "Autores": datos.get("autores", ""),
            "Comisión": datos.get("comision", ""),
            "Estado actual": datos.get("estado", ""),
            "Comentarios": row.get("tema", ""),
            "Link": row.get("url", ""),
            "Importancia": row.get("importancia", ""),
            "Fecha de último trámite": fecha_ultimo,
            "Descripción último trámite": desc_ultimo,
        })

    columnas = [
        "Nombre proyecto", "Boletín", "Año presentación", "Moción/Mensaje",
        "Autores", "Comisión", "Estado actual", "Comentarios", "Link", "Importancia",
        "Fecha de último trámite", "Descripción último trámite",
    ]
    df_out = pd.DataFrame(filas, columns=columnas)
    df_out.to_excel(OUTPUT_XLSX, index=False)
    print(f"✅ Listo: {OUTPUT_XLSX} ({len(df_out)} proyectos)")


if __name__ == "__main__":
    main()
