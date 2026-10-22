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
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
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


NUMEROS = {"un": 1, "uno": 1, "una": 1, "dos": 2, "tres": 3, "cuatro": 4, "cinco": 5, "seis": 6,
           "siete": 7, "ocho": 8, "nueve": 9, "diez": 10, "doce": 12, "quince": 15}
CARRERAS = [
    "administracion publica", "administrador publico", "ingenieria comercial", "ingenieria civil industrial",
    "ingenieria industrial", "contador auditor", "contador", "auditor", "ingenieria en administracion",
    "administracion de empresas", "economia", "ingenieria civil", "ciencia politica", "ciencias politicas",
    "sociologia", "trabajo social", "psicologia", "derecho", "abogado", "arquitectura", "medicina",
    "enfermeria", "estadistica", "control de gestion", "ingenieria informatica", "ingenieria en computacion",
    "ciencia de datos", "geografia", "periodismo", "pedagogia", "educacion parvularia", "kinesiologia", "nutricion", "odontologia",
]
DETALLE_COLS = ["formacion", "especializacion", "experiencia", "anios_exp_min", "anios_exp_max", "carreras",
                "grado", "tipo_vacante_det", "postulantes", "ultima_postulacion"]


def _sin_tildes(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s.lower()) if unicodedata.category(c) != "Mn")


def _texto_html(fragmento: str) -> str:
    t = re.sub(r"<[^>]+>", " ", fragmento)
    return re.sub(r"\s+", " ", unescape(t)).strip()


def _bloque(html_txt: str, art_id: str) -> str:
    m = re.search(r'id="%s"[^>]*>(.*?)</div>' % art_id, html_txt, re.S)
    if not m:
        return ""
    t = _texto_html(m.group(1))
    return re.sub(r"^(Formaci.n Educacional|Especializaci.n y/o Capacitaci.n|Experiencia sector p.blico / sector privado)\s*", "", t)


def anios_experiencia(texto: str):
    """(min, max) de años mencionados en el bloque de experiencia; (None, None) si no hay."""
    vals = []
    t = _sin_tildes(texto)
    for m in re.finditer(r"(\d{1,2}|" + "|".join(NUMEROS) + r")\s*(?:\(\d+\)\s*)?anos?\b", t):
        g = m.group(1)
        vals.append(int(g) if g.isdigit() else NUMEROS[g])
    vals = [v for v in vals if 0 < v <= 40]
    return (min(vals), max(vals)) if vals else (None, None)


def normalizar_estamento(tipo: str):
    """'Ingreso a Planta' -> Planta, 'Contrata Ley 19.664' -> Contrata, etc. None si viene vacío."""
    t = _sin_tildes(tipo if isinstance(tipo, str) else "")
    if not t.strip():
        return None
    if "planta" in t:
        return "Planta"
    if "contrata" in t:
        return "Contrata"
    if "honorario" in t:
        return "Honorarios"
    if "codigo del trabajo" in t:
        return "Código del Trabajo"
    if "suplencia" in t or "reemplazo" in t:
        return "Suplencia/Reemplazo"
    return "Otro"


def extraer_grado(texto: str):
    """Grado de la escala (EUS) mencionado en el texto de la oferta, como entero 1..31; None si no hay."""
    t = _sin_tildes(texto)
    for m in re.finditer(r"(?<![a-z])grado\s*(?:n[o°º]?\.?\s*)?(\d{1,2})", t):
        g = int(m.group(1))
        if 1 <= g <= 31:
            return g
    return None


def fetch_detalle(url: str) -> dict:
    vacio = {c: None for c in DETALLE_COLS}
    if "convpostularavisoTrabajo" not in url and "convFicha" not in url:
        return vacio
    try:
        for intento in range(3):
            try:
                r = requests.get(url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=30)
                r.raise_for_status()
                break
            except requests.RequestException:
                if intento == 2:
                    raise
                time.sleep(1.5)
        h = r.text
        grado = extraer_grado(_texto_html(h))
        mt = re.search(r"<dt>Tipo de Vacante</dt><dd>(.*?)</dd>", h, re.S)
        tipo_det = _texto_html(mt.group(1)) if mt else None
        formacion = _bloque(h, "articleFormacion")
        espec = _bloque(h, "articleEspecializaciones")
        exp = _bloque(h, "articleExperiencias")
        mp = re.search(r'detail-template__metric-row.*?<strong>\s*(\d+)\s*</strong>', h, re.S)
        mf = re.search(r'id=['"]lblFechaPost['"][^>]*>.*?(\d{1,2}:\d{2}:\d{2}).*?(\d{2}-\d{2}-\d{4})', h, re.S)
        ultima_post = f"{mf.group(2).split('-')[2]}-{mf.group(2).split('-')[1]}-{mf.group(2).split('-')[0]} {mf.group(1)}" if mf else None
        base = {"grado": grado, "tipo_vacante_det": tipo_det, "postulantes": int(mp.group(1)) if mp else None, "ultima_postulacion": ultima_post}
        if not (formacion or espec or exp):
            return {**vacio, **base}  # ficha externa (PDF / iframe): sin bloques
        mn, mx = anios_experiencia(exp)
        nf = _sin_tildes(formacion)
        carreras = ", ".join(c for c in CARRERAS if c in nf) or None
        return {"formacion": formacion or None, "especializacion": espec or None,
                "experiencia": exp or None, "anios_exp_min": mn, "anios_exp_max": mx, "carreras": carreras, **base}
    except Exception as e:
        logging.warning("Detalle falló (%s): %s", url[-60:], e)
        return vacio


def agregar_detalle(df: pd.DataFrame) -> pd.DataFrame:
    """Entra a cada oferta (en paralelo, sin Selenium) y rescata requisitos académicos y laborales."""
    urls = df["url_oferta"].fillna("").tolist()
    with ThreadPoolExecutor(max_workers=8) as ex:
        detalles = list(ex.map(fetch_detalle, urls))
    det = pd.DataFrame(detalles, columns=DETALLE_COLS, index=df.index)
    for c in ("anios_exp_min", "anios_exp_max", "grado", "postulantes"):
        det[c] = pd.to_numeric(det[c]).astype("Int64")  # enteros (no 5.0) para la columna int de Postgres
    df = df.copy()
    df["estamento"] = [normalizar_estamento(a) or normalizar_estamento(b)
                       for a, b in zip(df["tipo_vacante"], det.pop("tipo_vacante_det"))]
    for c in ("grado",):  # respaldo: grado mencionado en el cargo
        det[c] = det[c].fillna(pd.Series([extraer_grado(x) for x in df["cargo"]], index=df.index, dtype="Int64"))
    con = det["formacion"].notna().sum()
    logging.info("Detalle rescatado en %s de %s ofertas (grado %s, estamento %s)", con, len(df),
                 det["grado"].notna().sum(), df["estamento"].notna().sum())
    return pd.concat([df, det], axis=1)


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
        batch = records[i:i + BATCH_SIZE]
        try:
            sb.table(TABLE).upsert(batch, on_conflict="id_convocatoria").execute()
        except Exception as e:
            if "ultima_postulacion" in str(e).lower():
                logging.warning("Columna 'ultima_postulacion' pendiente en Supabase. Insertando temporalmente sin ella...")
                for r in batch:
                    r.pop("ultima_postulacion", None)
                sb.table(TABLE).upsert(batch, on_conflict="id_convocatoria").execute()
            else:
                raise
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
        df = agregar_detalle(df)
        exportar(df)
        subir_supabase(df)
    except Exception:
        logging.exception("Error en ejecución")
        sys.exit(1)


if __name__ == "__main__":
    main()
