#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un Excel con el detalle de cada proyecto de la watchlist,
combinando datos oficiales del webservice del Senado con las
columnas manuales (tema/importancia) de la planilla de Google.
"""

import re
import time
import pandas as pd
import requests
import xml.etree.ElementTree as ET

WATCHLIST_URL = "https://docs.google.com/spreadsheets/d/1RBS_VB7d3jyJvJL87gBZg5hx8I2PKslBfd3kWAxVdgA/export?format=csv"
SENADO_WS = "https://tramitacion.senado.cl/wspublico/tramitacion.php?boletin={numero}"
OUTPUT_XLSX = "proyectos_legislativos.xlsx"

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
    """Pega al webservice del Senado y extrae los campos que necesitamos
    del proyecto (título, año, autores, comisión, estado)."""
    numero = boletin.split("-")[0]  # el webservice solo acepta el correlativo
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

        # la subetapa suele traer el nombre de la comisión; si no, cae a la etapa
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


def main():
    print("Descargando watchlist desde Google Sheets...")
    watchlist = cargar_watchlist()

    filas = []
    for _, row in watchlist.iterrows():
        boletin = row["boletin"]
        print(f"Consultando boletín {boletin}...")
        datos = consultar_senado(boletin)
        time.sleep(0.5)  # evitar saturar el webservice del Senado

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
        })

    columnas = [
        "Nombre proyecto", "Boletín", "Año presentación", "Moción/Mensaje",
        "Autores", "Comisión", "Estado actual", "Comentarios", "Link", "Importancia",
    ]
    df_out = pd.DataFrame(filas, columns=columnas)
    df_out.to_excel(OUTPUT_XLSX, index=False)
    print(f"✅ Listo: {OUTPUT_XLSX} ({len(df_out)} proyectos)")


if __name__ == "__main__":
    main()
