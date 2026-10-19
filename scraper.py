"""
Scraper Empleos Públicos Chile — API JSON directa (sin Selenium).
Extrae convocatorias abiertas -> CSV/Excel -> Supabase (upsert + limpieza de cerradas).
"""
import csv
import hashlib
import logging
import os
import re
import sys
from datetime import datetime, timezone
from html import unescape
from urllib.parse import parse_qs, urlparse

import pandas as pd
import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

API_URL = "https://www.empleospublicos.cl/apiConvocatorias.ashx"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Referer": "https://www.empleospublicos.cl/",
    "X-Requested-With": "XMLHttpRequest",
}
TABLE = "empleos_publicos"
BATCH_SIZE = 200


def s(item: dict, *keys: str) -> str:
    """Primer campo presente (la API usa 'Nº' con distintos caracteres)."""
    for k in keys:
        v = item.get(k)
        if v is not None:
            return unescape(str(v)).strip()
    return ""


def parse_renta(val: str) -> int:
    digits = re.sub(r"[^\d]", "", val.split(",")[0]) if val else ""
    return int(digits) if digits else 0


def parse_fecha(val: str):
    """'19/10/2026 23:59:00' -> datetime | None"""
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(val.strip(), fmt)
        except ValueError:
            continue
    return None


def id_convocatoria(url: str, cargo: str, inst: str) -> str:
    """'ep-<i>' para ofertas del portal; 'ext-<hash>' para las que apuntan a sitios externos."""
    q = parse_qs(urlparse(url).query)
    i = q.get("i", [""])[0]
    if i.isdigit():
        return f"ep-{i}"
    return "ext-" + hashlib.sha1(f"{url}|{cargo}|{inst}".encode()).hexdigest()[:12]


def fetch(status: str = "Abierta") -> list[dict]:
    resp = requests.get(
        API_URL, headers=HEADERS, timeout=30,
        params={"page": 1, "pageSize": 1000, "status": status},
    )
    resp.raise_for_status()
    body = resp.json()
    items = body.get("data", [])
    total = body.get("total", len(items))
    logging.info("API total=%s obtenidos=%s", total, len(items))
    if total > len(items):  # paginar si hubiera más de 1000
        page = 2
        while len(items) < total:
            r = requests.get(API_URL, headers=HEADERS, timeout=30,
                             params={"page": page, "pageSize": 1000, "status": status})
            r.raise_for_status()
            chunk = r.json().get("data", [])
            if not chunk:
                break
            items += chunk
            page += 1
    return items


def transformar(items: list[dict]) -> pd.DataFrame:
    ahora = datetime.now(timezone.utc)
    hoy = datetime.now().date()
    rows = {}
    for it in items:
        url = s(it, "url")
        idc = id_convocatoria(url, s(it, "Cargo"), s(it, "Institución / Entidad"))
        renta = s(it, "Renta Bruta")
        cierre = parse_fecha(s(it, "Fecha Cierre Convocatoria"))
        inicio = parse_fecha(s(it, "Fecha Inicio"))
        try:
            vac = int(s(it, "Nº de Vacantes", "N° de Vacantes", "Nº de Vacantes") or 1)
        except ValueError:
            vac = 1
        rows[idc] = {
            "id_convocatoria": idc,
            "cargo": s(it, "Cargo"),
            "ministerio": s(it, "Ministerio"),
            "institucion": s(it, "Institución / Entidad"),
            "numero_vacantes": vac,
            "area_trabajo": s(it, "Área de Trabajo"),
            "region": s(it, "Región"),
            "ciudad": s(it, "Ciudad"),
            "tipo_vacante": s(it, "Tipo de Vacante"),
            "renta_bruta": renta,
            "renta_numerica": parse_renta(renta),
            "fecha_inicio": inicio.date().isoformat() if inicio else None,
            "fecha_cierre": cierre.date().isoformat() if cierre else None,
            "dias_restantes": max((cierre.date() - hoy).days, 0) if cierre else None,
            "tipo_postulacion": s(it, "Tipo postulacion"),
            "cargo_profesional": s(it, "Cargo Profesional"),
            "primer_empleo": bool(it.get("esPrimerEmpleo", False)),
            "url_oferta": url,
            "estado": s(it, "EstadoInferencia") or "Abierta",
            "extraido_en": ahora.isoformat(),
        }
    return pd.DataFrame(list(rows.values()))


def exportar(df: pd.DataFrame):
    df.to_excel("empleos_publicos.xlsx", index=False)
    df.to_csv("empleos_publicos.csv", index=False, encoding="utf-8-sig", sep=";", quoting=csv.QUOTE_ALL)
    logging.info("Generados empleos_publicos.xlsx y .csv")


def subir_supabase(df: pd.DataFrame):
    url, key = os.environ.get("SUPABASE_URL"), os.environ.get("SUPABASE_KEY")
    if not url or not key:
        raise RuntimeError("Faltan SUPABASE_URL / SUPABASE_KEY")
    from supabase import create_client
    sb = create_client(url, key)

    records = df.astype(object).where(pd.notnull(df), None).to_dict(orient="records")
    run_ts = records[0]["extraido_en"]
    for i in range(0, len(records), BATCH_SIZE):
        sb.table(TABLE).upsert(records[i:i + BATCH_SIZE], on_conflict="id_convocatoria").execute()
        logging.info("Upsert %s-%s", i + 1, min(i + BATCH_SIZE, len(records)))
    # Borra solo lo que ya no está en la lista (sólo si la carga completa fue exitosa)
    sb.table(TABLE).delete().lt("extraido_en", run_ts).execute()
    logging.info("Eliminadas convocatorias que ya no están abiertas.")


def main():
    try:
        items = fetch("Abierta")
        if not items:
            logging.error("La API no devolvió datos; no se toca la base.")
            sys.exit(1)
        df = transformar(items)
        logging.info("Procesadas %s ofertas", len(df))
        exportar(df)
        subir_supabase(df)
    except Exception:
        logging.exception("Error en ejecución")
        sys.exit(1)


if __name__ == "__main__":
    main()
