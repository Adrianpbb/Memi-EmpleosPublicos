# =========================
# SCRAPER EMPLEOS PÚBLICOS
# GitHub Actions + Supabase
# =========================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urljoin
from supabase import create_client
import pandas as pd
import time
import re
import csv
import os
import sys


# =========================
# CONFIGURACIÓN GENERAL
# =========================

BASE_URL = "https://www.empleospublicos.cl"
TABLE_NAME = "empleos_publicos"
BATCH_SIZE = 100

EXPORT_EXCEL = True
EXPORT_CSV = True

EXCEL_FILE = "empleos_publicos.xlsx"
CSV_FILE = "empleos_publicos.csv"


# =========================
# FUNCIONES DE LIMPIEZA
# =========================

def clean_text(text):
    """
    Limpia textos largos sin destruir información útil.
    Elimina saltos de línea, tabs, espacios repetidos y caracteres invisibles.
    """
    if text is None:
        return ""

    text = str(text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def safe_text_from_element(parent, by, selector, default="No especificado"):
    try:
        return clean_text(parent.find_element(by, selector).text)
    except Exception:
        return default


def get_block_text_by_id(driver, element_id, remove_headings=True, default="No especificado"):
    """
    Extrae texto completo de un bloque por ID.
    Es más robusto que depender de p/span/li específicos.
    """
    try:
        element = driver.find_element(By.ID, element_id)
        text = clean_text(element.text)

        if not text:
            return default

        if remove_headings:
            headings_to_remove = [
                "Formación Educacional",
                "Especialización y/o Capacitación",
                "Experiencia sector público / sector privado",
                "Condiciones",
                "Renta Bruta",
                "Objetivo del cargo",
                "Funciones del Cargo",
                "Perfil de la Función",
                "Competencias",
                "Requisitos Generales",
                "Requisitos Específicos",
                "Etapas de Selección"
            ]

            for heading in headings_to_remove:
                text = re.sub(rf"^{re.escape(heading)}\s*", "", text, flags=re.IGNORECASE)
                text = re.sub(rf"\b{re.escape(heading)}\b", "", text, flags=re.IGNORECASE)

            text = clean_text(text)

        return text if text else default

    except Exception:
        return default


# =========================
# EXTRACCIÓN DE DATOS
# =========================

def parse_datos_aviso(driver):
    """
    Extrae pares h3 -> p del bloque lblAvisoTrabajoDatos.
    Ejemplo:
    Institución, Convocatoria, Nº de Vacantes, Área de Trabajo, Región, Ciudad, Tipo de Vacante.
    """
    datos = {}

    try:
        contenedor = driver.find_element(By.ID, "lblAvisoTrabajoDatos")
        headings = contenedor.find_elements(By.TAG_NAME, "h3")

        for h in headings:
            label = clean_text(h.text).replace(":", "")
            value = ""

            try:
                p = h.find_element(By.XPATH, "following-sibling::p[1]")
                value = clean_text(p.text)
            except Exception:
                value = ""

            if label:
                datos[label] = value

    except Exception:
        pass

    return datos


def extract_postulantes(driver):
    try:
        return clean_text(driver.find_element(By.CSS_SELECTOR, "#paneInfo h1 strong").text)
    except Exception:
        return "No especificado"


def extract_renta(driver):
    """
    Extrae renta de forma robusta:
    1. Busca en lblRenta.
    2. Busca en lblCondiciones.
    3. Busca en panelBase.
    """
    textos_posibles = []

    for element_id in ["lblRenta", "lblCondiciones", "panelBase"]:
        try:
            texto = clean_text(driver.find_element(By.ID, element_id).text)
            if texto:
                textos_posibles.append(texto)
        except Exception:
            pass

    texto_total = " ".join(textos_posibles)

    if not texto_total:
        return "No especificado"

    patrones = [
        r"\$\s?(\d{1,3}(?:\.\d{3})+)",
        r"Renta Bruta[^0-9$]*(?:\$)?\s?(\d{1,3}(?:\.\d{3})+)",
        r"Renta[^0-9$]*(?:\$)?\s?(\d{1,3}(?:\.\d{3})+)",
        r"(\d{1,3}(?:\.\d{3})+)"
    ]

    for patron in patrones:
        match = re.search(patron, texto_total, flags=re.IGNORECASE)
        if match:
            return clean_text(match.group(1))

    return "No especificado"


def renta_a_numero(renta_texto):
    """
    Convierte '2.300.000' a 2300000.
    """
    if not renta_texto or renta_texto == "No especificado":
        return None

    solo_numero = re.sub(r"[^\d]", "", str(renta_texto))

    if solo_numero:
        try:
            return int(solo_numero)
        except Exception:
            return None

    return None


def extract_grado(driver):
    """
    Extrae grado desde distintos textos posibles.
    Ejemplos:
    - Grado: 12° E.U.S
    - La renta corresponde a Grado 13° E.U.S.
    - Profesional Grado 9° de la E.U.S.
    """
    textos_posibles = []

    for element_id in ["lblCondiciones", "lblAvisoTrabajoDatos", "panelBase"]:
        try:
            texto = clean_text(driver.find_element(By.ID, element_id).text)
            if texto:
                textos_posibles.append(texto)
        except Exception:
            pass

    texto_total = " ".join(textos_posibles)

    if not texto_total:
        return "No especificado"

    patrones = [
        r"Grado\s*:\s*(\d{1,2}\s*°?\s*(?:E\.?\s*U\.?\s*S\.?|EUS)?)",
        r"Grado\s*(\d{1,2}\s*°?\s*(?:E\.?\s*U\.?\s*S\.?|EUS)?)",
        r"Grado\s*(\d{1,2})\s*°?\s*de\s*la\s*E\.?\s*U\.?\s*S\.?",
        r"grado\s*(\d{1,2})"
    ]

    for patron in patrones:
        match = re.search(patron, texto_total, flags=re.IGNORECASE)
        if match:
            return clean_text(match.group(1))

    return "No especificado"


# =========================
# SCROLL Y TARJETAS
# =========================

def scroll_to_load_all_jobs(driver, wait, pause=1.0, max_scrolls=100):
    """
    La página carga ofertas hacia abajo.
    Hace scroll hasta que deja de aumentar la cantidad de div.items.
    """
    last_count = 0
    stable_rounds = 0

    for i in range(max_scrolls):
        job_elements = driver.find_elements(By.CSS_SELECTOR, "div.items")
        current_count = len(job_elements)

        print(f"Scroll {i + 1}: {current_count} ofertas detectadas")

        if current_count == last_count:
            stable_rounds += 1
        else:
            stable_rounds = 0

        if stable_rounds >= 4:
            print("No aparecen más ofertas. Se detiene el scroll.")
            break

        last_count = current_count
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.items")))
    return driver.find_elements(By.CSS_SELECTOR, "div.items")


def collect_job_cards(driver):
    """
    Extrae datos básicos y URL de cada tarjeta.
    Se hace antes de entrar al detalle para evitar stale elements.
    """
    cards = []
    job_elements = driver.find_elements(By.CSS_SELECTOR, "div.items")

    for index, job_element in enumerate(job_elements):
        try:
            cargo_element = job_element.find_element(By.CSS_SELECTOR, "h3 a")

            cargo = clean_text(cargo_element.text)
            href = cargo_element.get_attribute("href")
            url_oferta = urljoin(BASE_URL, href) if href else ""

            servicio_publico = safe_text_from_element(job_element, By.CSS_SELECTOR, "div.top p")
            ministerio = safe_text_from_element(job_element, By.CSS_SELECTOR, "div.cnt p")
            fecha_postulacion = safe_text_from_element(job_element, By.CSS_SELECTOR, "div.label-estado")

            cards.append({
                "index": index,
                "Cargo": cargo,
                "Servicio Público": servicio_publico,
                "Ministerio": ministerio,
                "Fecha Postulación Hasta": fecha_postulacion,
                "URL Oferta": url_oferta
            })

        except Exception as e:
            print(f"No se pudo leer tarjeta {index}: {e}")

    return cards


# =========================
# SELENIUM
# =========================

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    return webdriver.Chrome(options=options)


# =========================
# SCRAPER PRINCIPAL
# =========================

def scrape_empleos_publicos():
    driver = create_driver()
    wait = WebDriverWait(driver, 25)

    jobs = []

    try:
        print("Entrando a Empleos Públicos...")
        driver.get(BASE_URL)

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.items")))

        scroll_to_load_all_jobs(driver, wait, pause=1.0, max_scrolls=100)

        job_cards = collect_job_cards(driver)

        print(f"\nTotal de ofertas recolectadas: {len(job_cards)}\n")

        for i, card in enumerate(job_cards, start=1):
            cargo = card.get("Cargo", "Sin cargo")
            url_oferta = card.get("URL Oferta", "")

            print(f"[{i}/{len(job_cards)}] Procesando: {cargo}")

            if not url_oferta:
                print(f"Sin URL para: {cargo}")
                continue

            try:
                driver.get(url_oferta)

                try:
                    wait.until(EC.presence_of_element_located((By.ID, "panelBase")))
                except Exception:
                    print(f"No cargó panelBase en: {cargo}")

                datos_aviso = parse_datos_aviso(driver)

                institucion = datos_aviso.get("Institución", "No especificado")
                convocatoria = datos_aviso.get("Convocatoria", "No especificado")

                numero_vacantes = (
                    datos_aviso.get("Nº de Vacantes", "") or
                    datos_aviso.get("N° de Vacantes", "") or
                    datos_aviso.get("Nro de Vacantes", "") or
                    "No especificado"
                )

                area_trabajo = datos_aviso.get("Área de Trabajo", "No especificado")
                region = datos_aviso.get("Región", "No especificado")
                ciudad = datos_aviso.get("Ciudad", "No especificado")
                tipo_vacante = datos_aviso.get("Tipo de Vacante", "No especificado")

                postulantes = extract_postulantes(driver)

                grado = extract_grado(driver)
                renta = extract_renta(driver)
                renta_numero = renta_a_numero(renta)

                condiciones = get_block_text_by_id(driver, "lblCondiciones", remove_headings=True)
                formacion_academica = get_block_text_by_id(driver, "lblFormacion", remove_headings=True)
                especializacion = get_block_text_by_id(driver, "lblEspecializaciones", remove_headings=True)
                experiencia = get_block_text_by_id(driver, "lblExperiencias", remove_headings=True)

                funciones = get_block_text_by_id(driver, "lblFunciones", remove_headings=True)
                competencias = get_block_text_by_id(driver, "lblCompetencias", remove_headings=True)
                requisitos = get_block_text_by_id(driver, "lblRequisitos", remove_headings=True)
                etapas_seleccion = get_block_text_by_id(driver, "lblCriterio", remove_headings=True)

                jobs.append({
                    "Cargo": card.get("Cargo", "No especificado"),
                    "Servicio Público": card.get("Servicio Público", "No especificado"),
                    "Ministerio": card.get("Ministerio", "No especificado"),
                    "Fecha Postulación Hasta": card.get("Fecha Postulación Hasta", "No especificado"),
                    "URL Oferta": url_oferta,

                    "Institución": institucion,
                    "Convocatoria": convocatoria,
                    "Nº de Vacantes": numero_vacantes,
                    "Área de Trabajo": area_trabajo,
                    "Región": region,
                    "Ciudad": ciudad,
                    "Tipo Vacante": tipo_vacante,

                    "Postulantes": postulantes,
                    "Grado": grado,
                    "Renta Bruta": renta,
                    "Renta Numérica": renta_numero,

                    "Condiciones": condiciones,
                    "Formación Académica": formacion_academica,
                    "Especialización": especializacion,
                    "Experiencia": experiencia,
                    "Funciones": funciones,
                    "Competencias": competencias,
                    "Requisitos": requisitos,
                    "Etapas de Selección": etapas_seleccion
                })

            except Exception as e:
                print(f"Error procesando [{cargo}]: {e}")

    finally:
        driver.quit()

    df = pd.DataFrame(jobs)
    print(f"\nTotal de registros extraídos: {len(df)}")

    return df


# =========================
# EXPORTACIÓN LOCAL / ARTIFACTS
# =========================

def export_files(df):
    if EXPORT_EXCEL:
        df.to_excel(EXCEL_FILE, index=False)
        print(f"Excel generado: {EXCEL_FILE}")

    if EXPORT_CSV:
        df.to_csv(
            CSV_FILE,
            index=False,
            encoding="utf-8-sig",
            sep=";",
            quoting=csv.QUOTE_ALL
        )
        print(f"CSV generado: {CSV_FILE}")


# =========================
# SUPABASE
# =========================

def prepare_dataframe_for_supabase(df):
    """
    Renombra columnas para que coincidan con la tabla Supabase.
    """
    df_supabase = df.rename(columns={
        "Cargo": "cargo",
        "Servicio Público": "servicio_publico",
        "Ministerio": "ministerio",
        "Fecha Postulación Hasta": "fecha_postulacion_hasta",
        "URL Oferta": "url_oferta",

        "Institución": "institucion",
        "Convocatoria": "convocatoria",
        "Nº de Vacantes": "numero_vacantes",
        "Área de Trabajo": "area_trabajo",
        "Región": "region",
        "Ciudad": "ciudad",
        "Tipo Vacante": "tipo_vacante",

        "Postulantes": "postulantes",
        "Grado": "grado",
        "Renta Bruta": "renta_bruta",
        "Renta Numérica": "renta_numerica",

        "Condiciones": "condiciones",
        "Formación Académica": "formacion_academica",
        "Especialización": "especializacion",
        "Experiencia": "experiencia",
        "Funciones": "funciones",
        "Competencias": "competencias",
        "Requisitos": "requisitos",
        "Etapas de Selección": "etapas_seleccion"
    })

    expected_columns = [
        "cargo",
        "servicio_publico",
        "ministerio",
        "fecha_postulacion_hasta",
        "url_oferta",

        "institucion",
        "convocatoria",
        "numero_vacantes",
        "area_trabajo",
        "region",
        "ciudad",
        "tipo_vacante",

        "postulantes",
        "grado",
        "renta_bruta",
        "renta_numerica",

        "condiciones",
        "formacion_academica",
        "especializacion",
        "experiencia",
        "funciones",
        "competencias",
        "requisitos",
        "etapas_seleccion"
    ]

    for col in expected_columns:
        if col not in df_supabase.columns:
            df_supabase[col] = None

    df_supabase = df_supabase[expected_columns]

    # Convertir NaN a None para JSON/Postgres
    df_supabase = df_supabase.astype(object).where(pd.notnull(df_supabase), None)

    return df_supabase


def upload_to_supabase(df):
    SUPABASE_URL = os.environ.get("SUPABASE_URL")
    SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

    if not SUPABASE_URL or not SUPABASE_KEY:
        raise Exception("Faltan variables de entorno SUPABASE_URL o SUPABASE_KEY")

    print("Conectando a Supabase...")
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

    df_supabase = prepare_dataframe_for_supabase(df)
    records = df_supabase.to_dict(orient="records")

    print(f"Registros preparados para Supabase: {len(records)}")

    print("Vaciando tabla empleos_publicos...")
    supabase.rpc("reset_empleos_publicos").execute()

    if not records:
        print("No hay registros para insertar.")
        return

    print("Insertando registros en Supabase...")

    for i in range(0, len(records), BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        supabase.table(TABLE_NAME).insert(batch).execute()
        print(f"Insertados registros {i + 1} a {i + len(batch)}")

    print("Carga a Supabase completada.")


# =========================
# MAIN
# =========================

def main():
    try:
        df = scrape_empleos_publicos()

        if df.empty:
            print("No se extrajeron registros. Se detiene el proceso.")
            sys.exit(1)

        print("\nVista previa:")
        print(df.head())

        export_files(df)
        upload_to_supabase(df)

        print("\nProceso completado correctamente.")

    except Exception as e:
        print(f"\nERROR GENERAL: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
