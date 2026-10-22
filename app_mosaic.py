"""
Programa de Abasto Rural 2026 · Tablero ejecutivo (diseño estilo Mosaic)
Subgerencia de Abarrotes · Alimentación para el Bienestar

Ejecutar:   streamlit run app_mosaic.py

Origen de los datos (en este orden):
  1. OneDrive / SharePoint: si en los secretos (.streamlit/secrets.toml o "Secrets" en
     Streamlit Community Cloud) existe [onedrive] link = "...", el Excel se descarga de ahí
     y se vuelve a descargar solo cada [onedrive] minutos (10 por omisión).
  2. El archivo "PROGRAMA DE ABASTO RURAL 2026.xlsx" junto a este archivo.

Se lee la hoja VIGENTES (o la primera hoja si no existe); la fila de encabezados se busca sola.
El mapa usa mexico_estados.json (contornos de los 32 estados, simplificados a partir de
github.com/angelnmara/geojson), junto a este archivo.

Diseño inspirado en Mosaic Lite (cruip/tailwind-dashboard-template): sidebar con
navegación, header superior, tarjetas rounded-xl con encabezado, KPIs con etiqueta
en mayúsculas y modo oscuro. El acento violeta de Mosaic se sustituye por el
guinda institucional.
"""
import base64
import io
import json
import math
import re
from contextlib import contextmanager
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# ──────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Abasto Rural 2026 · Alimentación para el Bienestar",
    page_icon=":material/agriculture:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CARPETA = Path(__file__).parent
# Nombre actual del Excel y el anterior, por si todavía se usa la copia vieja
ARCHIVOS = [CARPETA / "PROGRAMA DE ABASTO RURAL 2026.xlsx", CARPETA / "PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx"]
ARCHIVO = next((a for a in ARCHIVOS if a.exists()), ARCHIVOS[0])
MAPA = CARPETA / "mexico_estados.json"

# En la nube el servidor corre en UTC: fechas y horas se muestran en hora del centro de México
try:
    ZONA = ZoneInfo("America/Mexico_City")
except Exception:  # noqa: BLE001  (Windows sin el paquete tzdata)
    ZONA = None


def ahora():
    return datetime.now(ZONA) if ZONA else datetime.now()


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Estructura operativa vigente desde el 1 de enero de 2020 (Gerencias Regionales).
# sigla de COBERTURA -> (nombre, clave, unidades). Cada unidad: (nombre, estado en el mapa, sede).
# Sin sede, la unidad lleva el nombre del estado y lo cubre completo; con sede, es una zona del
# estado nombrada por el municipio o la región donde está.
REGIONES = {
    "NOE": ("Noroeste", 41, [("U.O. Sinaloa", "Sinaloa", None),
                             ("U.O. Baja California", "Baja California", None),
                             ("U.O. Baja California Sur", "Baja California Sur", None),
                             ("U.O. Sonora", "Sonora", None)]),
    "NOR": ("Norte", 42, [("U.O. Chihuahua", "Chihuahua", None),
                          ("U.O. Durango", "Durango", None),
                          ("U.O. Coahuila", "Coahuila", None)]),
    "NC": ("Norte Centro", 43, [("U.O. Zacatecas", "Zacatecas", None),
                                ("U.O. San Luis Potosí", "San Luis Potosí", None),
                                ("U.O. Aguascalientes", "Aguascalientes", None),
                                ("U.O. Nuevo León", "Nuevo León", None)]),
    "GLF": ("Golfo", 44, [("Oficina Central Xalapa", "Veracruz", "municipio de Xalapa"),
                          ("U.O. Acayucán", "Veracruz", "municipio de Acayucan"),
                          ("U.O. Orizaba", "Veracruz", "municipio de Orizaba"),
                          ("U.O. Poza Rica", "Veracruz", "municipio de Poza Rica"),
                          ("U.O. Tamaulipas", "Tamaulipas", None)]),
    "OCC": ("Occidente", 45, [("U.O. Jalisco", "Jalisco", None),
                              ("U.O. Nayarit", "Nayarit", None),
                              ("U.O. Colima", "Colima", None)]),
    "CEN": ("Centro", 46, [("U.O. Hidalgo", "Hidalgo", None),
                           ("U.O. Querétaro", "Querétaro", None),
                           ("U.O. Guanajuato", "Guanajuato", None)]),
    "OAX": ("Oaxaca", 47, [("U.O. Oaxaca", "Oaxaca", None),
                           ("U.O. Istmo", "Oaxaca", "región del Istmo"),
                           ("U.O. Mixteca", "Oaxaca", "región Mixteca")]),
    "MET": ("Metropolitana", 48, [("U.O. Estado de México", "México", None),
                                  ("U.O. Morelos", "Morelos", None)]),
    "PEN": ("Peninsular", 49, [("U.O. Yucatán", "Yucatán", None),
                               ("U.O. Campeche", "Campeche", None),
                               ("U.O. Quintana Roo", "Quintana Roo", None)]),
    "STE": ("Sureste", 50, [("U.O. Tabasco", "Tabasco", None),
                            ("U.O. Tuxtla Gutiérrez", "Chiapas", "municipio de Tuxtla Gutiérrez"),
                            ("U.O. Tapachula", "Chiapas", "municipio de Tapachula")]),
    "PSR": ("Pacífico Sur", 51, [("U.O. Michoacán", "Michoacán", None),
                                 ("U.O. Guerrero", "Guerrero", None)]),
    "PUE": ("Puebla", 52, [("U.O. Puebla", "Puebla", None),
                           ("U.O. Tlaxcala", "Tlaxcala", None)]),
}
ESTADO_REGION = {estado: sig for sig, (_, _, unidades) in REGIONES.items() for _, estado, _ in unidades}
N_UNIDADES = sum(len(u) for _, _, u in REGIONES.values())

# Paleta institucional gob.mx
GUINDA, GUINDA_MED = "#611232", "#9D2449"
VERDE_MED, VERDE_CLARO = "#235B4E", "#3E8E7E"
ORO, ORO_CLARO = "#A57F2C", "#BC955C"
AZUL = "#2E6B9E"

# Íconos de línea (trazos de Lucide, licencia ISC): heredan el color del texto con currentColor
ICONOS = {
    "proveedor": '<path d="M2 20a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8l-7 5V8l-7 5V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/>'
                 '<path d="M17 18h1"/><path d="M12 18h1"/><path d="M7 18h1"/>',
    "producto": '<circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 '
                '2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/>',
    "articulo": '<path d="m7.5 4.27 9 5.15"/><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 '
                '0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>',
    "marca": '<path d="M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 '
             '2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r=".5" fill="currentColor"/>',
    "precio": '<line x1="12" x2="12" y1="2" y2="22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
    "idea": '<path d="M15 14c.2-1 .7-1.7 1.5-2.5 1-.9 1.5-2.2 1.5-3.5A6 6 0 0 0 6 8c0 1 .2 2.2 1.5 3.5.7.7 1.3 1.5 1.5 2.5"/>'
            '<path d="M9 18h6"/><path d="M10 22h4"/>',
    "tendencia": '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
    "estrella": '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
    "pastel": '<path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
    "enlace": '<path d="M9 17H7A5 5 0 0 1 7 7h2"/><path d="M15 7h2a5 5 0 1 1 0 10h-2"/><line x1="8" x2="16" y1="12" y2="12"/>',
    "mapa": '<path d="M14.106 5.553a2 2 0 0 0 1.788 0l3.659-1.83A1 1 0 0 1 21 4.619v12.764a1 1 0 0 1-.553.894l-4.553 '
            '2.277a2 2 0 0 1-1.788 0l-4.212-2.106a2 2 0 0 0-1.788 0l-3.659 1.83A1 1 0 0 1 3 19.381V6.618a1 1 0 0 1 '
            '.553-.894l4.553-2.277a2 2 0 0 1 1.788 0z"/><path d="M15 5.764v15"/><path d="M9 3.236v15"/>',
    "alerta": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/>'
              '<path d="M12 17h.01"/>',
    "ok": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
    "flecha": '<path d="m6 9 6 6 6-6"/>',
    "buscar": '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    "nacional": '<circle cx="12" cy="12" r="10"/><path d="M12 2a14.5 14.5 0 0 0 0 20 14.5 14.5 0 0 0 0-20"/>'
                '<path d="M2 12h20"/>',
    "regional": '<path d="M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 '
                '8 0 0 1 16 0"/><circle cx="12" cy="10" r="3"/>',
}


def svg(nombre, tam=18):
    return (f'<svg width="{tam}" height="{tam}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICONOS[nombre]}</svg>')


def icono_mask(nombre, tam=18):
    """Ícono para st.html: su sanitizador quita los <svg>, así que el ícono va como máscara CSS (data URI)
    pintada con el color del texto, igual que los <svg> con stroke="currentColor"."""
    dato = base64.b64encode(svg(nombre, tam).replace('stroke="currentColor"', 'stroke="#000"').encode()).decode()
    return (f'<span class="svgi" aria-hidden="true" style="width:{tam}px;height:{tam}px;'
            f'--i:url(data:image/svg+xml;base64,{dato})"></span>')


# Tokens de tema (grises de Mosaic; en oscuro, colores de datos más claros para que contrasten)
TEMAS = {
    "claro": dict(
        fondo="#f3f4f6", card="#ffffff", sutil="#f9fafb", linea="#f3f4f6", borde="#e5e7eb",
        texto="#1f2937", texto2="#4b5563", texto3="#6b7280", tenue="#9ca3af",
        acento=GUINDA_MED, acento_rgb="157 36 73", ok="#239f52", alerta="#e63939", otros="#d1d5db",
        paleta=[GUINDA_MED, VERDE_MED, ORO, AZUL, VERDE_CLARO, GUINDA, "#7A5C1E", "#B5547A", "#5F6B6D"],
        regiones=[GUINDA_MED, VERDE_MED, ORO, AZUL, VERDE_CLARO, GUINDA, "#7A5C1E", "#B5547A", "#5F6B6D",
                  "#4C7A3A", "#7B4F9E", "#C06A2B"],
        verde=VERDE_MED, oro=ORO, azul=AZUL, verde_claro=VERDE_CLARO,
        sombra="0 1px 1px 0 rgb(0 0 0 / .05), 0 1px 2px 0 rgb(0 0 0 / .02)",
    ),
    "oscuro": dict(
        fondo="#111827", card="#1f2937", sutil="#263142", linea="#2d3748", borde="#374151",
        texto="#f3f4f6", texto2="#d1d5db", texto3="#9ca3af", tenue="#6b7280",
        acento="#E0698C", acento_rgb="224 105 140", ok="#4bd37d", alerta="#ff7474", otros="#4b5563",
        paleta=["#E0698C", "#4FB39A", "#D9AE55", "#6AA6DE", "#8FD3C2", "#C77D93", "#B8955A", "#E59AB3", "#94A3B8"],
        regiones=["#E0698C", "#4FB39A", "#D9AE55", "#6AA6DE", "#8FD3C2", "#C77D93", "#B8955A", "#E59AB3", "#94A3B8",
                  "#8CC26E", "#B58BD9", "#E8955A"],
        verde="#4FB39A", oro="#D9AE55", azul="#6AA6DE", verde_claro="#8FD3C2",
        sombra="0 1px 2px 0 rgb(0 0 0 / .35)",
    ),
}

# ──────────────────────────────────────────────────────────────
# CARGA Y LIMPIEZA
# ──────────────────────────────────────────────────────────────
RENOMBRAR = {
    "NEGOCIACION": "NEGOCIACION", "CVE": "CVE", "PROVEEDOR": "PROVEEDOR",
    "PROD.": "PROD", "PRODUCTOS": "PRODUCTO", "ART.": "ART_ID", "ARTICULOS": "ARTICULO",
    "MARCA": "MARCA", "CAP": "CAP", "GRAM": "GRAM", "U/M": "UM",
    "COSTO DE ADQUISICIÓN 2026": "COSTO_CAJA",
    "COSTO UNITARIO DE ADQUISICIÓN 2026": "COSTO_UNIT",
    "IMPUESTO": "IMPUESTO", "CODIGO BARRAS": "CODIGO_BARRAS",
    "TIPO DE INVENTARIO": "INVENTARIO", "COBERTURA": "COBERTURA", "VIGENCIA": "VIGENCIA",
}
NUMERICAS = ["CAP", "GRAM", "COSTO_CAJA", "COSTO_UNIT"]
INVENTARIO_LABEL = {"PAR": "PAR", "OFERTA": "Oferta"}
OTROS = "Otros"


def nombre_corto(nombre: str) -> str:
    """Quita la razón social (S.A. de C.V., SA DE CV, S.A.P.I., S. de R.L., S.C.L.…) y pasa a formato título."""
    corto = re.sub(r"[,.\s]+(S\.?\s?A\.?|S\.?\s?DE\s?R\.?\s?L\.?|S\.?\s?C\.?|S\.?\s?P\.?\s?R\.?)(?![A-Z]).*$", "",
                   str(nombre).strip(), flags=re.I)
    return corto.strip(" ,.").title() or str(nombre)


def clave_texto(s: pd.Series) -> pd.Series:
    """Claves numéricas como texto, sin el ".0" que agrega Excel (653.0 -> "653")."""
    s = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return s.where(~s.isin(["", "nan", "NaN", "<NA>", "None"]), pd.NA)


def presentacion(cap, gram, um) -> str:
    if pd.isna(cap) or pd.isna(gram):
        return "–"
    return f"{cap:g} × {gram:g} {um}".strip()


@st.cache_data(show_spinner="Preparando el tablero…")
def cargar(contenido: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(contenido))
    hoja = "VIGENTES" if "VIGENTES" in xls.sheet_names else xls.sheet_names[0]
    # El encabezado no siempre está en la misma fila: se usa la primera que diga NEGOCIACION o CVE
    vista = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=30)
    fila = next((i for i, r in vista.iterrows()
                 if r.astype(str).str.strip().str.upper().isin(["NEGOCIACION", "CVE"]).any()), None)
    if fila is None:
        raise ValueError(f"no encontré la fila de encabezados en la hoja {hoja}")
    raw = pd.read_excel(xls, sheet_name=hoja, header=fila)
    raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]
    raw = raw.rename(columns=RENOMBRAR)
    raw = raw.loc[:, ~raw.columns.str.startswith("Unnamed")]
    raw = raw.loc[:, ~raw.columns.duplicated()]
    faltan = [c for c in ["CVE", "PROVEEDOR", "PRODUCTO", "ART_ID", "ARTICULO"] if c not in raw]
    if faltan:
        raise ValueError("faltan las columnas " + ", ".join(faltan))

    df = raw[raw["ART_ID"].notna()].copy()
    for c in ["NEGOCIACION", "PROD", "MARCA", "UM", "IMPUESTO", "CODIGO_BARRAS", "INVENTARIO", "COBERTURA",
              "VIGENCIA", *NUMERICAS]:
        if c not in df:
            df[c] = pd.NA
    for c in NUMERICAS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["NEGOCIACION", "CVE", "PROD", "ART_ID", "CODIGO_BARRAS"]:
        df[c] = clave_texto(df[c])

    for c in ["PROVEEDOR", "PRODUCTO", "ARTICULO", "MARCA", "UM", "IMPUESTO"]:
        df[c] = df[c].astype(str).str.strip().str.upper().replace({"NAN": pd.NA, "": pd.NA})
    df["MARCA"] = df["MARCA"].fillna("SIN MARCA").str.title()
    df["PRODUCTO"] = df["PRODUCTO"].fillna("SIN CLASIFICAR").str.capitalize()
    df["ARTICULO"] = df["ARTICULO"].fillna("SIN NOMBRE")
    df["IMPUESTO"] = df["IMPUESTO"].fillna("Sin dato")
    df["UM"] = df["UM"].fillna("")
    df["PROV_CORTO"] = df["PROVEEDOR"].map(nombre_corto)
    # Los productos se identifican por PROD.; si una clave trae varios nombres se usa el más frecuente
    df["PROD"] = df["PROD"].fillna(df["PRODUCTO"])
    df["PRODUCTO"] = df.groupby("PROD")["PRODUCTO"].transform(lambda s: s.mode().iat[0])

    inv = clave_texto(df["INVENTARIO"]).str.upper()
    df["INVENTARIO"] = inv.map(lambda x: INVENTARIO_LABEL.get(x, x), na_action="ignore").fillna("Sin dato")

    # COBERTURA: NAL = todo el país; si no, la lista de siglas de región (NOE NOR NC …)
    cob = df["COBERTURA"].astype(str).str.upper().str.replace(r"[^A-Z]+", " ", regex=True).str.strip()
    cob = cob.where(~cob.isin(["", "NAN", "NA"]), pd.NA)
    fichas = cob.fillna("").str.split()
    df["COBERTURA"] = cob.fillna("Sin dato")
    df["ALCANCE"] = ["Nacional" if "NAL" in f else ("Regional" if f else "Sin dato") for f in fichas]
    df["REGIONES"] = [tuple(REGIONES) if "NAL" in f else tuple(r for r in REGIONES if r in f) for f in fichas]

    df["VIG_FECHA"] = pd.to_datetime(df["VIGENCIA"], errors="coerce")
    df["PRESENTACION"] = [presentacion(c, g, u) for c, g, u in zip(df["CAP"], df["GRAM"], df["UM"])]
    return df.reset_index(drop=True)


def n_art(d) -> int:
    """Artículos distintos: la clave ART. se repite cuando el artículo está en varias negociaciones."""
    return d["ART_ID"].nunique()


def conteo(d, col) -> pd.Series:
    return d.groupby(col)["ART_ID"].nunique().sort_values(ascending=False)


def por_region(d) -> pd.Series:
    """Artículos distintos que llegan a cada región (los nacionales cuentan en todas)."""
    e = d[["ART_ID", "REGIONES"]].explode("REGIONES").dropna()
    return e.groupby("REGIONES")["ART_ID"].nunique().reindex(list(REGIONES), fill_value=0)


# ──────────────────────────────────────────────────────────────
# MAPA Y SILUETAS DE LOS ESTADOS
# ──────────────────────────────────────────────────────────────
@st.cache_data
def geo_estados():
    return json.loads(MAPA.read_text(encoding="utf-8")) if MAPA.exists() else None


def _anillos(geometria):
    """Contorno exterior de cada polígono (las islas cuentan como polígonos aparte)."""
    poligonos = geometria["coordinates"] if geometria["type"] == "MultiPolygon" else [geometria["coordinates"]]
    return [p[0] for p in poligonos]


@st.cache_data
def centros_region():
    """Punto medio aproximado de cada región (centro de los recuadros de sus estados), para la sigla."""
    gj = geo_estados()
    if not gj:
        return {}
    cajas = {}
    for f in gj["features"]:
        pts = [p for a in _anillos(f["geometry"]) for p in a]
        lons, lats = [p[0] for p in pts], [p[1] for p in pts]
        cajas[f["properties"]["name"]] = ((min(lons) + max(lons)) / 2, (min(lats) + max(lats)) / 2)
    salida = {}
    for sig, (_, _, unidades) in REGIONES.items():
        c = [cajas[e] for e in {e for _, e, _ in unidades} if e in cajas]
        if c:
            salida[sig] = (sum(x for x, _ in c) / len(c), sum(y for _, y in c) / len(c))
    return salida


# ──────────────────────────────────────────────────────────────
# TEMA (va primero: define el CSS). El interruptor se dibuja más abajo, en el encabezado; su valor
# ya está en session_state al empezar cada ejecución.
# ──────────────────────────────────────────────────────────────
def md(texto, destino=st):
    destino.markdown(texto, unsafe_allow_html=True)


# El tema se guarda en la URL (?tema=oscuro): sobrevive a recargas y se puede compartir
if "oscuro" not in st.session_state:
    st.session_state["oscuro"] = st.query_params.get("tema") == "oscuro"
oscuro = st.session_state["oscuro"]
st.query_params["tema"] = "oscuro" if oscuro else "claro"

T = TEMAS["oscuro" if oscuro else "claro"]
COLOR_REGION = dict(zip(REGIONES, T["regiones"]))
NAL, NEGRO = "NAL", "#111111"   # botón «Nacional» de la sección 4: todo el mapa en negro
COLOR_ALCANCE = {"Nacional": T["verde"], "Regional": T["oro"], "Sin dato": T["tenue"]}
COLOR_INV = {"PAR": T["acento"], "Oferta": T["azul"], "301": T["verde"], "75": T["oro"], "Sin dato": T["tenue"]}

# ──────────────────────────────────────────────────────────────
# ESTILO (Mosaic). Los colores viven en variables CSS para cambiar de tema.
# ──────────────────────────────────────────────────────────────
VARIABLES = ":root{" + "".join(f"--{k.replace('_', '-')}:{v};" for k, v in T.items() if isinstance(v, str)) + "}"
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, .stApp, [class*="css"], button, input, textarea, select {
  font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif; }
.stApp { background: var(--fondo); color: var(--texto2); }
.block-container { padding-top: .5rem !important; padding-bottom: 3rem; max-width: 1440px; }
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent; }
h1, h2, h3, h4 { font-family: 'Inter', sans-serif !important; color: var(--texto); letter-spacing: -.01em; }
.stApp p, .stApp label { color: var(--texto2); }
[data-testid="stIconMaterial"] { color: var(--texto3); }
[data-testid="stCaptionContainer"] { color: var(--texto3); }

/* Sin barra lateral: todo el tablero va en una sola página */
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="stExpandSidebarButton"] {
  display: none !important; }

/* Controles de Streamlit 1.66 (react-aria). Se apuntan por data-testid, role y estados
   (aria-checked / data-selected), no por clases internas, para que sigan el tema claro/oscuro. */

/* Interruptor (modo oscuro) */
[data-testid="stCheckbox"] label > div:not([data-testid]) { background: var(--borde) !important; }
[data-testid="stCheckbox"][data-selected="true"] label > div:not([data-testid]) { background: var(--acento) !important; }
[data-testid="stCheckbox"] label > div:not([data-testid]) > div { background: #fff !important; }
[data-testid="stCheckbox"] p, [data-testid="stCheckbox"] span[role="img"] { color: var(--texto2) !important; }

/* Entradas: selectbox y texto */
[data-testid="stSelectbox"] div[role="group"][data-rac], [data-testid="stTextInputRootElement"] {
  background: var(--sutil) !important; border: 1px solid var(--borde) !important; border-radius: 8px !important; }
[data-testid="stSelectbox"] div[role="group"][data-focus-within="true"],
[data-testid="stTextInputRootElement"]:focus-within { border-color: var(--acento) !important; }
[data-testid="stSelectbox"] input, [data-testid="stTextInputField"] {
  color: var(--texto) !important; background: transparent !important; caret-color: var(--acento); }
[data-testid="stSelectbox"] input::placeholder, [data-testid="stTextInputField"]::placeholder {
  color: var(--tenue) !important; opacity: 1; }
[data-testid="stSelectbox"] button { color: var(--texto3) !important; background: transparent !important; }

/* Listas desplegables (se abren fuera de la app, en un portal) */
[data-testid$="Dropdown"] { background: var(--card) !important; border: 1px solid var(--borde) !important;
  border-radius: 8px !important; box-shadow: 0 8px 24px rgb(0 0 0 / .18) !important; }
[data-testid$="Dropdown"] [role="listbox"] { background: var(--card) !important; }
[data-testid$="Dropdown"] [role="option"] { color: var(--texto2) !important; background: transparent !important; }
[data-testid$="Dropdown"] [role="option"][data-focused="true"],
[data-testid$="Dropdown"] [role="option"]:hover { background: var(--sutil) !important; color: var(--texto) !important; }
[data-testid$="Dropdown"] [role="option"][aria-selected="true"] { color: var(--acento) !important; font-weight: 600; }

/* Banner: cada recuadro es un <details>; al hacer clic se despliega su listado de nombres */
.banner { display: grid; grid-template-columns: repeat(4, 1fr); gap: 24px; align-items: start; margin-bottom: 24px; }
details.kpi-det > summary { list-style: none; cursor: pointer; border-radius: 12px; transition: box-shadow .15s; }
details.kpi-det > summary::-webkit-details-marker { display: none; }
details.kpi-det > summary:hover { box-shadow: inset 0 0 0 1px rgb(var(--acento-rgb) / .45); }
details.kpi-det .kpi > div:nth-child(2) { flex: 1; min-width: 0; }
details.kpi-det .flecha { color: var(--tenue); transition: transform .2s; margin-top: 2px; }
details.kpi-det[open] .flecha { transform: rotate(180deg); color: var(--acento); }
details.kpi-det .l { font-size: 1.2rem; }
/* Título y pista en una sola línea: los cuatro recuadros quedan de la misma altura */
details.kpi-det .l, details.kpi-det .pista { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
details.kpi-det .v { font-size: 3.2rem; }
details.kpi-det .ico { width: 58px; height: 58px; }
details.kpi-det .pista { font-size: 1.1rem; color: var(--acento); font-weight: 500; margin-top: 4px; }
details.kpi-det[open] > summary { box-shadow: inset 0 0 0 1px rgb(var(--acento-rgb) / .45); }
details.kpi-det .nombres { margin: 0 20px 16px 20px; border-top: 1px solid var(--linea); }
details.kpi-det .nombres li { font-size: 1.05rem; }
details.kpi-det .s { font-size: 1.05rem; }
.svgi { display: inline-block; flex-shrink: 0; background: currentColor;
  -webkit-mask: var(--i) center / contain no-repeat; mask: var(--i) center / contain no-repeat; }
.banner.dos { grid-template-columns: repeat(2, 1fr); }
/* Buscador de cada listado: filtra solo los nombres de su recuadro */
.buscador { display: flex; align-items: center; gap: 8px; margin: 12px 20px 4px 20px; padding: 6px 10px;
  background: var(--sutil); border: 1px solid var(--borde); border-radius: 8px; color: var(--tenue); }
.buscador:focus-within { border-color: var(--acento); }
.buscador input { flex: 1; min-width: 0; border: none; outline: none; background: transparent; font-size: .95rem;
  color: var(--texto); caret-color: var(--acento); }
.buscador input::placeholder { color: var(--tenue); opacity: 1; }
.buscador .cuenta { font-size: .78rem; white-space: nowrap; }
details.kpi-det .nombres li.vacio { color: var(--tenue); font-style: italic; border-bottom: none; }

/* Encabezado de cada sección de la página */
.seccion { display: flex; align-items: center; gap: 12px; margin: 40px 0 16px 0; }
.seccion .ico { width: 36px; height: 36px; border-radius: 10px; flex-shrink: 0; display: grid; place-items: center;
  background: rgb(var(--acento-rgb) / .1); color: var(--acento); }
.seccion h2 { font-size: 1.3rem !important; font-weight: 700 !important; margin: 0; padding: 0 !important; }
.seccion p { margin: 0; font-size: .84rem; color: var(--texto3); }

/* Lista estilo "Recent Activity" (Lectura rápida) */
.lista { list-style: none; margin: 0; padding: 0; }
.lista li { display: flex; gap: 12px; padding: 12px 0; border-bottom: 1px solid var(--linea);
  font-size: .875rem; color: var(--texto2); align-items: center; }
.lista li:last-child { border-bottom: none; }
.lista .ico { width: 36px; height: 36px; border-radius: 50%; flex-shrink: 0; display: grid; place-items: center; color: #fff; }
.lista b { color: var(--texto); font-weight: 600; }

/* ── Header superior (ancho completo) ─────────────────── */
.topbar { background: var(--card); display: flex; align-items: center; justify-content: space-between;
  gap: 16px; height: 64px; margin-bottom: 28px;
  box-shadow: 0 0 0 100vmax var(--card), 0 1px 0 100vmax var(--borde); clip-path: inset(0 -100vmax -1px -100vmax); }
.topbar .crumb { font-size: .875rem; color: var(--texto3); }
.topbar .crumb b { color: var(--texto); font-weight: 600; }
.topbar .fecha { display: inline-flex; align-items: center; gap: 8px; font-size: .8rem; color: var(--texto2);
  font-weight: 500; border: 1px solid var(--borde); border-radius: 8px; padding: 6px 12px; white-space: nowrap; }
.topbar .fecha .origen { color: var(--tenue); }
/* Controles del encabezado (modo oscuro y actualizar datos) */
.st-key-actualizar button { background: var(--card) !important; border: 1px solid var(--borde) !important;
  color: var(--texto2) !important; border-radius: 8px !important; font-size: .8rem; }
.st-key-actualizar button:hover { border-color: var(--acento) !important; color: var(--acento) !important; }
.st-key-actualizar button * { color: inherit !important; }
.st-key-controles { align-items: flex-end; }
.topbar .fecha i { width: 8px; height: 8px; border-radius: 50%; background: var(--ok);
  box-shadow: 0 0 0 3px rgb(62 201 114 / .2); }

/* ── Título de página ───────────────────────────────── */
.titulo { display: flex; justify-content: space-between; align-items: flex-end; gap: 16px;
  flex-wrap: wrap; margin-bottom: 24px; }
.titulo h1 { font-size: 2.6rem; font-weight: 700; margin: 0; padding: 0; line-height: 1.33; }
.titulo p { margin: 4px 0 0 0; font-size: 1.1rem; color: var(--texto3); }
.etiq { display: inline-flex; align-items: center; gap: 6px; font-size: .75rem; font-weight: 500; color: var(--texto2);
  background: var(--card); border: 1px solid var(--borde); border-radius: 999px; padding: 3px 10px; margin: 2px 0 2px 6px; }
.etiq i { width: 7px; height: 7px; border-radius: 50%; }

/* ── Tarjetas ───────────────────────────────────────── */
.m-card, div[class*="st-key-card_"] { background: var(--card); border-radius: 12px; border: none !important;
  box-shadow: var(--sombra); }
div[class*="st-key-card_"] { padding: 0 20px 16px 20px !important; gap: .5rem; }
.card-h { margin: 0 -20px 8px -20px; padding: 16px 20px; border-bottom: 1px solid var(--linea);
  display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
.card-h h2 { font-size: 1rem !important; font-weight: 600 !important; margin: 0; padding: 0; color: var(--texto); }
.card-h p { font-size: .8rem; color: var(--tenue); margin: 2px 0 0 0; }
.card-h .extras { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 6px; }
.card-h .extra { font-size: .75rem; font-weight: 600; color: var(--acento); white-space: nowrap;
  background: rgb(var(--acento-rgb) / .1); border-radius: 999px; padding: 2px 10px; }
[data-testid="stPlotlyChart"] { background: transparent; }

/* KPIs */
.kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 24px; margin-bottom: 24px; }
.kpi { padding: 20px; display: flex; gap: 14px; align-items: flex-start; }
.kpi .ico { width: 40px; height: 40px; border-radius: 10px; flex-shrink: 0; display: grid; place-items: center;
  background: rgb(var(--acento-rgb) / .1); color: var(--acento); }
.kpi .l { font-size: .72rem; font-weight: 600; color: var(--tenue); text-transform: uppercase; letter-spacing: .04em; }
.kpi .row { display: flex; align-items: flex-start; flex-wrap: wrap; gap: 4px 8px; margin-top: 2px; }
.kpi .v { font-size: 1.75rem; font-weight: 700; color: var(--texto); line-height: 1.2; }
.kpi .s { font-size: .78rem; color: var(--texto3); margin-top: 2px; }
.badge { font-size: .75rem; font-weight: 600; border-radius: 999px; padding: 1px 8px; white-space: nowrap; margin-top: 6px; }
.badge.ok { color: var(--ok); background: rgb(62 201 114 / .18); }
.badge.neutro { color: var(--texto3); background: var(--sutil); }

/* Listado de nombres del banner */
.nombres { list-style: none; margin: 0; padding: 0 4px 0 0; max-height: 320px; overflow-y: auto; }
.nombres li { padding: 6px 2px; border-bottom: 1px solid var(--borde); font-size: .82rem; color: var(--texto); }
.nombres li:last-child { border-bottom: none; }
.nombres::-webkit-scrollbar { width: 6px; }
.nombres::-webkit-scrollbar-thumb { background: var(--borde); border-radius: 999px; }

/* Texto para listas vacías */
.lb-mas { font-size: .78rem; color: var(--tenue); padding-top: 8px; margin-top: 4px; border-top: 1px dashed var(--borde); }

/* Tabla estilo "Top Channels" (con desplazamiento y encabezado fijo cuando es larga) */
.tabla-wrap { overflow-x: auto; }
.tabla-wrap.scroll { overflow-y: auto; }
.tabla { width: 100%; border-collapse: collapse; font-size: .875rem; }
.tabla thead th { font-size: .7rem; text-transform: uppercase; color: var(--tenue); font-weight: 600;
  background: var(--sutil); padding: 8px; text-align: left; border: none; white-space: nowrap; }
.tabla-wrap.scroll thead th { position: sticky; top: 0; z-index: 1; }
.tabla thead th:first-child { border-radius: 6px 0 0 6px; }
.tabla thead th:last-child { border-radius: 0 6px 6px 0; }
.tabla td { padding: 10px 8px; border: none; border-bottom: 1px solid var(--linea); color: var(--texto2); font-weight: 500; }
.tabla tr:last-child td { border-bottom: none; }
.tabla td.fuerte { color: var(--texto); }
.tabla .num { text-align: right; white-space: nowrap; }
.tabla small { display: block; color: var(--tenue); font-size: .74rem; font-weight: 500; }

/* Fichas (tipos de inventario de un proveedor) */
.chip { display: inline-flex; align-items: center; gap: 5px; font-size: .74rem; font-weight: 500; color: var(--texto2);
  border: 1px solid var(--borde); border-radius: 999px; padding: 1px 8px; margin: 1px 4px 1px 0; white-space: nowrap; }
.chip i { width: 7px; height: 7px; border-radius: 50%; }
.chip b { color: var(--texto); font-weight: 600; }
.chip.sel { border-color: var(--acento); background: rgb(var(--acento-rgb) / .08); }

/* Guía de lectura */
.guia { display: flex; gap: 10px; align-items: flex-start; background: rgb(188 149 92 / .12);
  border-radius: 8px; padding: 10px 14px; font-size: .82rem; color: var(--texto2); margin: 0 0 6px 0; }
.guia svg { color: var(--oro); flex-shrink: 0; margin-top: 1px; }
.guia b { color: var(--texto); }

/* Leyenda de las donas: cada renglón se elige con un clic, igual que su color en la gráfica */
div[class*="st-key-leyenda_"] { gap: 0 !important; margin-top: 4px; }
div[class*="st-key-leyitem_"] { position: relative; gap: 0 !important; border-bottom: 1px solid var(--linea);
  border-radius: 6px; transition: background .15s; }
div[class*="st-key-leyitem_"]:last-child { border-bottom: none; }
div[class*="st-key-leyitem_"]:hover { background: var(--sutil); }
div[class*="st-key-leyitem_"]:hover .n { color: var(--acento); }
div[class*="st-key-leyitem_"][class*="__sel"] { background: rgb(var(--acento-rgb) / .08); }
div[class*="st-key-leyitem_"][class*="__sel"] .n { color: var(--acento); font-weight: 700; }
div[class*="st-key-leyitem_"] [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
.ley { display: flex; align-items: center; gap: 8px; padding: 7px 6px; font-size: .84rem; }
.ley i { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
.ley .n { flex: 1; min-width: 0; color: var(--texto); font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ley .v { color: var(--texto2); font-weight: 600; white-space: nowrap; }
.ley small { color: var(--tenue); font-weight: 500; margin-left: 6px; }
.ley .ovalo { font-size: .74rem; font-weight: 600; color: var(--texto); white-space: nowrap; border: 1px solid;
  border-radius: 999px; padding: 1px 9px; }
.ley em { font-style: normal; color: var(--tenue); font-weight: 500; display: inline-block; min-width: 44px; text-align: right; }
.pista-clic { font-size: .75rem; color: var(--tenue); text-align: center; margin-top: -6px; }

/* Regiones como botones: todo el recuadro (región y sus U.O.) se elige con un clic.
   Un st.button transparente cubre el recuadro completo. */
div[class*="st-key-regbtn_"] { position: relative; border: 1px solid var(--borde); border-radius: 10px;
  padding: 8px 12px !important; background: var(--card); gap: 0 !important; transition: all .15s; }
div[class*="st-key-regbtn_"]:hover { border-color: var(--acento); background: rgb(var(--acento-rgb) / .04); }
div[class*="st-key-regbtn_"][class*="_on"] { border-color: var(--acento); background: rgb(var(--acento-rgb) / .1);
  box-shadow: inset 3px 0 0 var(--acento); }
div[class*="st-key-regbtn_"] [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
/* Capa de clic: un st.button transparente que cubre todo su recuadro (regiones y renglones de leyenda) */
div[class*="st-key-regclic_"], div[class*="st-key-leyclic_"] { position: absolute !important; inset: 0;
  width: 100% !important; height: 100% !important; z-index: 2; margin: 0 !important; }
div[class*="st-key-regclic_"] *, div[class*="st-key-leyclic_"] * { width: 100% !important; height: 100% !important; }
div[class*="st-key-regclic_"] button, div[class*="st-key-leyclic_"] button { opacity: 0; cursor: pointer; }
.rb-h { display: flex; align-items: center; gap: 8px; }
.rb-h .sig { font-size: .66rem; font-weight: 700; color: var(--card); border-radius: 5px; padding: 2px 6px;
  letter-spacing: .03em; flex-shrink: 0; }
.rb-h b { color: var(--texto); font-size: .84rem; font-weight: 600; }
.rb-h small { color: var(--tenue); font-size: .72rem; margin-left: auto; }
.rb-u { font-size: .75rem; color: var(--texto3); margin-top: 3px; line-height: 1.45; }

/* Leyenda de la dona de productos: producto estrella y artículo estrella */
.estrellas { list-style: none; margin: 8px 0 0 0; padding: 0; }
.estrellas li { display: flex; gap: 10px; align-items: flex-start; padding: 8px 0; border-bottom: 1px solid var(--linea);
  font-size: .82rem; color: var(--texto2); }
.estrellas li:last-child { border-bottom: none; }
.estrellas.lado { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 0; }
.estrellas.lado li { border: 1px solid var(--borde); border-radius: 10px; padding: 10px 12px; background: var(--sutil); }
@media (max-width: 640px) { .estrellas.lado { grid-template-columns: 1fr; } }
.estrellas svg { color: var(--oro); flex-shrink: 0; margin-top: 1px; }
.estrellas .l { display: block; font-size: .68rem; font-weight: 600; color: var(--tenue); text-transform: uppercase;
  letter-spacing: .04em; }
.estrellas b { color: var(--texto); font-weight: 600; }
.estrellas small { display: block; color: var(--tenue); font-size: .74rem; }

/* Índice de secciones */
.indice { display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 8px 0; }
.indice a { display: inline-flex; align-items: center; gap: 8px; text-decoration: none !important; font-size: .82rem;
  font-weight: 500; color: var(--texto2) !important; background: var(--card); border: 1px solid var(--borde);
  border-radius: 999px; padding: 6px 14px 6px 6px; transition: all .15s; }
.indice a:hover { border-color: var(--acento); color: var(--acento) !important; }
.indice a span, .seccion .paso { width: 22px; height: 22px; border-radius: 50%; display: grid; place-items: center;
  background: var(--acento); color: #fff; font-size: .72rem; font-weight: 700; flex-shrink: 0; }
.seccion .paso { width: 30px; height: 30px; font-size: .9rem; }
.seccion { scroll-margin-top: 70px; }

/* Lectura rápida en tarjetas */
.hallazgos { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
.hz { display: flex; gap: 12px; align-items: flex-start; padding: 16px; font-size: .86rem; color: var(--texto2);
  border-left: 4px solid var(--c); }
.hz .ico { width: 36px; height: 36px; border-radius: 50%; flex-shrink: 0; display: grid; place-items: center;
  color: #fff; background: var(--c); }
.hz b { color: var(--texto); font-weight: 600; }

/* Barra con lo que está seleccionado */
div[class*="st-key-sel_bar"] { background: rgb(var(--acento-rgb) / .06); border: 1px dashed rgb(var(--acento-rgb) / .45);
  border-radius: 10px; padding: 6px 12px !important; }
.selbar { font-size: .86rem; color: var(--texto2); }
.selbar b { color: var(--acento); }
div[class*="st-key-sel_bar"] button { background: var(--card) !important; border: 1px solid var(--borde) !important;
  color: var(--texto2) !important; border-radius: 8px !important; font-size: .8rem; }
div[class*="st-key-sel_bar"] button:hover { border-color: var(--acento) !important; color: var(--acento) !important; }
div[class*="st-key-sel_bar"] button * { color: inherit !important; }

/* Pie de página */
.pie { text-align: center; margin-top: 40px; font-size: .75rem; color: var(--tenue); }

/* ── Responsivo ─────────────────────────────────────── */
@media (max-width: 1100px) {
  .banner { grid-template-columns: repeat(2, 1fr); gap: 16px; }
  .hallazgos { grid-template-columns: repeat(2, 1fr); }
  .kpis { grid-template-columns: repeat(2, 1fr); gap: 16px; }
}
@media (max-width: 900px) {
  .titulo h1 { font-size: 1.9rem; }
  .topbar .crumb span { display: none; }
}
@media (max-width: 640px) {
  .block-container { padding-left: 1rem !important; padding-right: 1rem !important; }
  .topbar { height: auto; min-height: 56px; flex-wrap: wrap; gap: 8px; padding: 10px 0; margin-bottom: 18px; }
  .titulo { margin-bottom: 16px; }
  .banner, .banner.dos { grid-template-columns: 1fr; gap: 12px; }
  .hallazgos { grid-template-columns: 1fr; }
  .kpis { gap: 12px; margin-bottom: 16px; }
  .kpi { padding: 14px; gap: 10px; }
  .kpi .ico { width: 32px; height: 32px; }
  .kpi .v { font-size: 1.35rem; }
  div[class*="st-key-card_"] { padding: 0 14px 12px 14px !important; }
  .card-h { margin: 0 -14px 8px -14px; padding: 14px; flex-wrap: wrap; }
  .tabla { font-size: .8rem; }
}
@media (max-width: 420px) { .kpis { grid-template-columns: 1fr; } }
"""
md(f"<style>{VARIABLES}{CSS}</style>")


# ──────────────────────────────────────────────────────────────
# COMPONENTES DE PRESENTACIÓN
# ──────────────────────────────────────────────────────────────
def fecha_larga(d) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def dinero(x) -> str:
    return "–" if pd.isna(x) else f"${x:,.2f}"


def alfa(color, a):
    c = color.lstrip("#")
    return f"rgba({int(c[0:2], 16)},{int(c[2:4], 16)},{int(c[4:6], 16)},{a})"


@contextmanager
def tarjeta(titulo, clave, sub="", extra=""):
    """Tarjeta de Mosaic: fondo de tarjeta, rounded-xl, encabezado con línea inferior."""
    extras = [extra] if isinstance(extra, str) else extra   # una etiqueta o varias, una al lado de otra
    chips = "".join(f"<span class=extra>{e}</span>" for e in extras if e)
    with st.container(border=True, key=f"card_{clave}"):
        md(f'<div class="card-h"><div><h2>{escape(titulo)}</h2>{f"<p>{sub}</p>" if sub else ""}</div>'
           f'{f"<div class=extras>{chips}</div>" if chips else ""}</div>')
        yield


def kpi(etiqueta, valor, sub="", badge=None, icono="", tono="ok"):
    b = f'<span class="badge {tono}">{badge}</span>' if badge else ""
    i = f'<div class="ico">{svg(icono, 20)}</div>' if icono else ""
    return (f'<div class="m-card kpi">{i}<div><div class="l">{etiqueta}</div>'
            f'<div class="row"><div class="v">{valor}</div>{b}</div><div class="s">{sub}</div></div></div>')


def fila_kpis(items):
    md('<div class="kpis">' + "".join(kpi(**i) for i in items) + "</div>")


def porcentaje(parte, total):
    """Badge de proporción; se omite cuando la selección es el total."""
    return f"{parte / total:.0%} del total" if total and parte != total else None


# Filtra los nombres del listado mientras se escribe (sin acentos ni mayúsculas). st.markdown no ejecuta
# JavaScript, por eso el banner se dibuja con st.html; el script se registra una sola vez por página.
BUSCADOR_JS = """<script>
if (!window.__buscadorListas) {
  window.__buscadorListas = true;
  const norm = t => t.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  document.addEventListener("input", e => {
    const caja = e.target.closest && e.target.closest(".buscador input");
    if (!caja) return;
    const det = caja.closest("details"), q = norm(caja.value);
    let n = 0;
    det.querySelectorAll(".nombres li:not(.vacio)").forEach(li => {
      const ok = !q || norm(li.textContent).includes(q);
      li.hidden = !ok;
      n += ok;
    });
    det.querySelector(".nombres li.vacio").hidden = n > 0;
    det.querySelector(".buscador .cuenta").textContent = q ? n.toLocaleString("es-MX") + " de " + det.dataset.total : "";
  });
}
</script>"""


def banner(recuadros, clase=""):
    """Recuadros grandes (<details>): al hacer clic se despliega su listado de nombres, con un buscador propio.
    recuadros: (título, ícono, valor, subtítulo, nombres)."""
    html = []
    for titulo, icono, valor, sub, nombres in recuadros:
        nombres = sorted(set(nombres))
        html.append(
            f'<details class="m-card kpi-det" data-total="{len(nombres):,}"><summary><div class="kpi">'
            f'<div class="ico">{icono_mask(icono, 30)}</div><div><div class="l" title="{escape(titulo)}">{escape(titulo)}</div>'
            f'<div class="row"><div class="v">{valor:,}</div></div>{f"<div class=s>{sub}</div>" if sub else ""}'
            f'<div class="pista">Clic para ver los nombres</div></div>'
            f'<span class="flecha">{icono_mask("flecha", 18)}</span></div></summary>'
            f'<label class="buscador">{icono_mask("buscar", 16)}<input type="search" placeholder="Buscar en {escape(titulo.lower())}…"'
            f' autocomplete="off"><span class="cuenta"></span></label>'
            f'<ul class="nombres">{"".join(f"<li>{escape(n)}</li>" for n in nombres)}'
            f'<li class="vacio" hidden>Sin coincidencias.</li></ul></details>')
    st.html(f'<div class="banner {clase}">{"".join(html)}</div>{BUSCADOR_JS}', unsafe_allow_javascript=True)


def guia(texto):
    md(f'<div class="guia">{svg("idea", 16)}<span>{texto}</span></div>')


def tabla(encabezados, filas, numericas=(), alto=None):
    th = "".join(f'<th class="{"num" if i in numericas else ""}">{h}</th>' for i, h in enumerate(encabezados))
    tr = "".join(
        "<tr>" + "".join(
            f'<td class="{"num" if i in numericas else ""} {"fuerte" if i == 0 else ""}">{c}</td>'
            for i, c in enumerate(fila)) + "</tr>"
        for fila in filas)
    marco = f' scroll" style="max-height:{alto}px' if alto else ""
    md(f'<div class="tabla-wrap{marco}"><table class="tabla"><thead><tr>{th}</tr></thead>'
       f'<tbody>{tr}</tbody></table></div>')


def barras_clic(serie, clave, destino, color=None, total=None, alto_max=None, etiquetas=None, vaciable=True,
                n_corto=30):
    """Lista de barras horizontales en la que se puede hacer clic, igual que en las gráficas: el elemento elegido
    queda en st.session_state[destino] y se resalta. Con vaciable, otro clic sobre el mismo lo quita."""
    if serie.empty:
        md('<div class="lb-mas">Sin datos.</div>')
        return None
    color = color or T["acento"]
    total = total or serie.sum()
    claves = list(serie.index)
    nombres = etiquetas or [str(k) for k in claves]
    sel = st.session_state.get(destino)
    sel = sel if sel in claves else None
    fig = go.Figure(go.Bar(
        x=serie.values, y=etiquetas_cortas(nombres, n_corto), orientation="h",
        marker_color=[color if sel in (None, k) else alfa(color, .3) for k in claves],
        text=[f"{v:,} · {v / total:.0%}" for v in serie.values], textposition="outside", cliponaxis=False,
        textfont=dict(color=T["texto2"], size=11), customdata=nombres,
        hovertemplate="<b>%{customdata}</b><br>%{x:,} artículos · clic para elegir<extra></extra>"))
    alto = alto_barras(len(serie), 24, 120)
    estilo(fig, alto)
    fig.update_layout(bargap=.3, margin=dict(t=4, b=4, r=8))
    fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)", tickfont=dict(size=11, color=T["texto2"]),
                     ticksuffix="  ")
    fig.update_xaxes(visible=False, range=[0, serie.max() * 1.25])

    def al_clic():
        indices = st.session_state[clave]["selection"]["point_indices"]
        if not indices:
            return
        elegido = claves[indices[0]]
        st.session_state[destino] = None if vaciable and elegido == st.session_state.get(destino) else elegido

    with st.container(height=alto_max if alto_max and alto > alto_max else "content", border=False):
        st.plotly_chart(fig, width="stretch", theme=None, key=clave, on_select=al_clic, selection_mode="points",
                        config={"displayModeBar": False})
    return sel


def chip(texto, color, n=None, sel=False):
    cuenta = f" <b>{n:,}</b>" if n is not None else ""
    return f'<span class="chip{" sel" if sel else ""}"><i style="background:{color}"></i>{escape(str(texto))}{cuenta}</span>'


def corto(texto, n=28):
    texto = str(texto)
    return texto if len(texto) <= n else texto[:n - 1].rstrip() + "…"


def etiquetas_cortas(valores, n=28):
    """Recorta nombres largos para los ejes sin que dos categorías terminen con la misma etiqueta."""
    vistas, salida = set(), []
    for v in valores:
        e = corto(v, n)
        while e in vistas:
            e += "​"
        vistas.add(e)
        salida.append(e)
    return salida


def alto_barras(n, por_barra=30, minimo=240):
    return max(minimo, por_barra * n + 50)


def estilo(fig, alto=380, leyenda=False):
    """Estilo común. automargin reserva espacio para las etiquetas de los ejes y evita que se corten."""
    fig.update_layout(
        height=alto,
        margin=dict(l=8, r=24, t=20, b=8),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=T["texto3"], size=12),
        showlegend=leyenda, colorway=T["paleta"],
        legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center", title=None, font=dict(color=T["texto2"])),
        hoverlabel=dict(bgcolor=T["card"], bordercolor=T["borde"], font_size=12, font_family="Inter, sans-serif",
                        font_color=T["texto"]),
        bargap=0.3, barcornerradius=4,
    )
    fig.update_xaxes(gridcolor=T["linea"], zeroline=False, title=None, showline=False, automargin=True,
                     tickfont=dict(color=T["tenue"]))
    fig.update_yaxes(gridcolor=T["linea"], zeroline=False, title=None, showline=False, automargin=True,
                     tickfont=dict(color=T["texto3"]))
    return fig


def mostrar(fig, clave):
    st.plotly_chart(fig, width="stretch", theme=None, key=clave, config={"displayModeBar": False})


RAIZ = "__todos__"


def dona_clic(serie, clave, destino, alto=240, colores=None, max_items=9, unidad="artículos", detalle=None,
              vaciable=True, centro=None, base=None, agrupar=True, unidad_centro=None, alto_ley=None, ovalo=None):
    """Dona en la que se puede hacer clic; el elemento elegido queda en st.session_state[destino].

    Streamlit solo avisa de los clics en barras, puntos y gráficas jerárquicas (no en go.Pie), así que la
    dona se dibuja como un sunburst de un nivel: el centro es el total y cada sector un elemento. Clic en
    el centro (o doble clic) quita la selección, salvo con vaciable=False.
    centro: número al centro (p. ej. artículos distintos, que no es la suma de los sectores si un artículo
    está en varios); base: total para los porcentajes de la leyenda; agrupar=False deja solo los
    max_items primeros, sin sector «Otros»; unidad_centro: texto bajo el número del centro si cuenta otra cosa
    (p. ej. productos); alto_ley: alto máximo de la leyenda, con desplazamiento; ovalo: dato extra por elemento,
    en un óvalo del color del sector a la izquierda de las cifras.
    La leyenda también se puede elegir con un clic en cada renglón.
    """
    s = serie[serie > 0].sort_values(ascending=False)
    if len(s) > max_items and not agrupar:
        s = s.head(max_items)
    elif len(s) > max_items:
        s = pd.concat([s.head(max_items - 1), pd.Series({OTROS: s.iloc[max_items - 1:].sum()})])
    nombres = [str(k) for k in s.index]
    if isinstance(colores, dict):
        libres = [c for c in T["paleta"] if c not in colores.values()] or T["paleta"]
        cols = [colores.get(k) or libres[i % len(libres)] for i, k in enumerate(nombres)]
    else:
        lista = list(colores or T["paleta"])
        cols = [lista[i % len(lista)] for i in range(len(s))]
    if OTROS in nombres:
        cols[nombres.index(OTROS)] = T["otros"]

    sel = st.session_state.get(destino)
    sel = sel if sel in nombres else None
    rellenos = [c if sel in (None, k) else alfa(c, .3) for c, k in zip(cols, nombres)]
    total = int(s.sum())
    centro = total if centro is None else int(centro)
    base = base or total
    fig = go.Figure(go.Sunburst(
        ids=[RAIZ, *nombres], labels=["Todos", *nombres], parents=["", *[RAIZ] * len(s)],
        values=[total, *[int(v) for v in s.values]], branchvalues="total", sort=False,
        marker=dict(colors=[T["card"], *rellenos], line=dict(color=T["card"], width=3)),
        leaf=dict(opacity=1),
        texttemplate=[f"<b>{centro:,}</b><br><span style='font-size:11px'>{unidad_centro or unidad}</span>",
                      *[""] * len(s)],
        insidetextorientation="horizontal",
        textfont=dict(family="Inter, sans-serif", size=20, color=T["texto"]),
        customdata=[0, *[v / base for v in s.values]],
        hovertemplate=["Clic aquí para ver todo<extra></extra>" if vaciable else "<extra></extra>",
                       *[f"<b>%{{label}}</b><br>%{{value:,}} {unidad} · %{{customdata:.0%}}<extra></extra>"] * len(s)],
    ))
    estilo(fig, alto)
    fig.update_layout(margin=dict(l=0, r=0, t=4, b=4))

    def elegir(elegido):
        if elegido == OTROS and not vaciable:   # «Otros» agrupa varios: no es un elemento que se pueda elegir
            return
        if vaciable and elegido == st.session_state.get(destino):   # otro clic sobre el mismo lo quita
            elegido = None
        if elegido or vaciable:
            st.session_state[destino] = elegido

    def al_clic():
        puntos = st.session_state[clave]["selection"]["points"]
        elegido = puntos[0].get("id") if puntos else None
        elegir(None if elegido == RAIZ else elegido)

    st.plotly_chart(fig, width="stretch", theme=None, key=clave, on_select=al_clic, selection_mode="points",
                    config={"displayModeBar": False})
    md(f'<div class="pista-clic">{"Clic en el centro o en el mismo renglón para ver todo" if sel and vaciable else "Haz clic en un color o en un renglón para ver el detalle"}</div>')
    detalle, ovalo = detalle or {}, ovalo or {}
    with st.container(key=f"leyenda_{clave}", height=alto_ley if alto_ley else "content", border=False):
        for n, (k, v, c) in enumerate(zip(nombres, s.values, cols)):
            ov = (f'<span class="ovalo" style="background:{alfa(c, .16)};border-color:{alfa(c, .55)}">{ovalo[k]}</span>'
                  if k in ovalo else "")
            with st.container(key=f"leyitem_{clave}_{n}{'__sel' if k == sel else ''}"):
                md(f'<div class="ley"><i style="background:{c}"></i><span class="n" title="{escape(k)}">{escape(k)}</span>{ov}'
                   f'<span class="v">{v:,}{f"<small>{detalle[k]}</small>" if k in detalle else ""}'
                   f'<em>{v / base:.0%}</em></span></div>')
                st.button(f"Elegir {k}", key=f"leyclic_{clave}_{n}", on_click=elegir, args=(k,))
    return sel


def estrellas_region(todos) -> dict:
    """Por región: productos regionales distintos, producto estrella (el de más artículos regionales) y artículo
    estrella (el artículo de ese producto que más veces se negoció). Se usan solo los regionales: los nacionales
    llegan a todas y harían que el primer lugar fuera el mismo producto en las 12 regiones."""
    e = todos[todos["ALCANCE"] == "Regional"].explode("REGIONES").dropna(subset=["REGIONES"])
    salida = {}
    for sig, g in e.groupby("REGIONES"):
        prods_r = conteo(g, "PRODUCTO")
        estrella = prods_r.index[0]
        arts = g[g["PRODUCTO"] == estrella].groupby(["ARTICULO", "PROV_CORTO"]).size()
        arts = arts.sort_index().sort_values(ascending=False, kind="stable")
        salida[sig] = dict(n_prod=len(prods_r), n_art=n_art(g), producto=estrella, n_estrella=int(prods_r.iloc[0]),
                           articulo=arts.index[0][0], proveedor=arts.index[0][1])
    return salida


def mapa(info: dict, clave, resaltar=None, alto=440):
    """Estados coloreados por gerencia regional; al pasar el cursor se ve el producto estrella de la región.
    Con resaltar=sigla, las demás regiones quedan en gris; con resaltar=NAL, todo el país va en negro."""
    nacional = resaltar == NAL
    gj = geo_estados()
    if not gj:
        st.caption(f"No encontré {MAPA.name} junto a la app: el mapa no se puede dibujar.")
        return
    fig = go.Figure()
    for sig, (nombre, clave_r, unidades) in REGIONES.items():
        estados = list(dict.fromkeys(e for _, e, _ in unidades))
        color = NEGRO if nacional else T["otros"] if resaltar and sig != resaltar else COLOR_REGION[sig]
        uos = {e: ", ".join(u.removeprefix("U.O. ") for u, ee, _ in unidades if ee == e) for e in estados}
        i = info.get(sig)
        texto = (f"★ <b>{escape(i['producto'])}</b><br>{i['n_prod']:,} productos regionales" if i
                 else "Sin artículos regionales")
        fig.add_trace(go.Choropleth(
            geojson=gj, featureidkey="properties.name", locations=estados, z=[1] * len(estados),
            colorscale=[[0, color], [1, color]], showscale=False,
            marker_line_color=T["card"], marker_line_width=1,
            customdata=[[nombre, sig, clave_r, uos[e], texto] for e in estados],
            hovertemplate="<b>%{customdata[1]} · %{customdata[0]}</b> (clave %{customdata[2]})"
                          "<br>%{location}: %{customdata[3]}<br>%{customdata[4]}<extra></extra>"))
    sin = [f["properties"]["name"] for f in gj["features"] if f["properties"]["name"] not in ESTADO_REGION]
    if sin:   # p. ej. Ciudad de México, que no aparece en la estructura operativa
        fig.add_trace(go.Choropleth(
            geojson=gj, featureidkey="properties.name", locations=sin, z=[1] * len(sin),
            colorscale=[[0, NEGRO if nacional else T["sutil"]], [1, NEGRO if nacional else T["sutil"]]],
            showscale=False, marker_line_color=T["card"] if nacional else T["borde"], marker_line_width=1,
            hovertemplate="<b>%{location}</b><br>Sin unidad operativa en la estructura<extra></extra>"))
    centros = centros_region()
    siglas = [s for s in REGIONES if s in centros]
    fig.add_trace(go.Scattergeo(
        lon=[centros[s][0] for s in siglas], lat=[centros[s][1] for s in siglas], text=siglas, mode="text",
        textfont=dict(family="Inter, sans-serif", size=11,
                      color=["#ffffff" if nacional else T["texto3"] if resaltar and s != resaltar else T["card"]
                             for s in siglas]),
        hoverinfo="skip"))
    fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)", projection_type="mercator")
    estilo(fig, alto)
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False)
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    st.plotly_chart(fig, width="stretch", theme=None, key=clave,
                    config={"displayModeBar": False, "scrollZoom": False})


def texto_cobertura(g) -> str:
    """Cobertura de un proveedor: Nacional, Regional o ambas si tiene artículos de los dos tipos."""
    return "".join(chip(a, COLOR_ALCANCE[a]) for a in ["Nacional", "Regional"] if (g["ALCANCE"] == a).any()) or "–"


# ──────────────────────────────────────────────────────────────
# DATOS
# ──────────────────────────────────────────────────────────────
def codigo_share(link: str) -> str:
    """Codifica un vínculo para compartir en el formato de la API "shares" de OneDrive/Graph."""
    return "u!" + base64.urlsafe_b64encode(link.strip().encode()).decode().rstrip("=")


def url_descarga(link: str) -> str:
    """Vínculo para compartir de OneDrive/SharePoint -> URL de descarga directa del archivo."""
    link = link.strip()
    if "sharepoint.com" in link:   # OneDrive de trabajo (Microsoft 365): basta con download=1
        partes = urlsplit(link)
        query = [(k, v) for k, v in parse_qsl(partes.query) if k != "download"] + [("download", "1")]
        return urlunsplit(partes._replace(query=urlencode(query)))
    # OneDrive personal (1drv.ms / onedrive.live.com)
    return f"https://api.onedrive.com/v1.0/shares/{codigo_share(link)}/root/content"


def config_onedrive() -> dict:
    try:
        return dict(st.secrets.get("onedrive", {}))
    except Exception:  # noqa: BLE001  (no existe secrets.toml)
        return {}


@st.cache_data(max_entries=2, show_spinner="Descargando el Excel de OneDrive…")
def descargar_excel(link: str, bloque: int):
    """Descarga el Excel. `bloque` cambia cada N minutos, así el archivo se vuelve a bajar solo."""
    cfg = config_onedrive()
    encabezados = {"User-Agent": "Mozilla/5.0"}
    if all(cfg.get(k) for k in ("tenant_id", "client_id", "client_secret")):
        # Archivo NO público: se descarga con una aplicación registrada en Microsoft Entra (Graph)
        token = requests.post(
            f"https://login.microsoftonline.com/{cfg['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": cfg["client_id"], "client_secret": cfg["client_secret"],
                  "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"},
            timeout=30)
        token.raise_for_status()
        encabezados["Authorization"] = f"Bearer {token.json()['access_token']}"
        url = f"https://graph.microsoft.com/v1.0/shares/{codigo_share(link)}/driveItem/content"
    else:
        url = url_descarga(link)
    r = requests.get(url, headers=encabezados, timeout=60, allow_redirects=True)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):   # un .xlsx es un zip; si no, OneDrive devolvió una página de acceso
        raise ValueError("OneDrive no devolvió un Excel: revisa que el vínculo permita ver el archivo "
                         "sin iniciar sesión, o configura el acceso con aplicación de Microsoft.")
    return r.content, ahora()


def motivo(e: Exception) -> str:
    """Explicación corta del error, sin exponer la URL ni credenciales."""
    if isinstance(e, requests.HTTPError) and e.response is not None:
        codigo = e.response.status_code
        pistas = {401: "sin autorización", 403: "acceso denegado", 404: "no se encontró el archivo"}
        return f"HTTP {codigo}, {pistas.get(codigo, 'error del servidor')}"
    if isinstance(e, requests.RequestException):
        return "no hubo conexión con OneDrive"
    return str(e)


cfg_od = config_onedrive()
if cfg_od.get("link"):
    minutos = max(1, int(cfg_od.get("minutos", 10)))
    try:
        contenido, descargado = descargar_excel(cfg_od["link"], int(ahora().timestamp() // (minutos * 60)))
        origen = f"OneDrive · {descargado:%H:%M}"
    except Exception as e:  # noqa: BLE001
        if not ARCHIVO.exists():
            st.error(f"No pude descargar el Excel de OneDrive ({motivo(e)}).")
            st.stop()
        st.warning(f"No pude descargar el Excel de OneDrive ({motivo(e)}). Muestro la copia local.",
                   icon=":material/cloud_off:")
        contenido, origen = ARCHIVO.read_bytes(), "Copia local"
elif ARCHIVO.exists():
    contenido, origen = ARCHIVO.read_bytes(), "Archivo local"
else:
    st.error(f"No hay datos: agrega el vínculo de OneDrive en los secretos (`[onedrive] link`) "
             f"o coloca **{ARCHIVO.name}** junto a app_mosaic.py.")
    st.stop()

try:
    datos = cargar(contenido)
except Exception as e:  # noqa: BLE001
    st.error(f"No pude leer el Excel con la estructura esperada: {e}")
    st.stop()
if datos.empty:
    st.warning("El Excel no tiene artículos capturados.")
    st.stop()

hoy = pd.Timestamp(ahora().date())


SECCIONES = {"resumen": "Lectura rápida", "cobertura": "Proveedores y cobertura", "productos": "Productos y artículos",
             "unidades": "Unidades operativas"}


def seccion(ancla, sub):
    """Encabezado numerado; el id permite saltar a la sección desde el índice de arriba."""
    n = list(SECCIONES).index(ancla) + 1
    md(f'<div class="seccion" id="{ancla}"><span class="paso">{n}</span><div><h2>{SECCIONES[ancla]}</h2>'
       f'<p>{sub}</p></div></div>')


def barra_seleccion(clave, partes, limpiar):
    """Muestra lo que está elegido y un botón para volver a ver todo."""
    def quitar():
        for k in limpiar:
            st.session_state.pop(k, None)

    with st.container(key=f"sel_bar_{clave}", horizontal=True, vertical_alignment="center"):
        md('<div class="selbar">Estás viendo: ' + " › ".join(f"<b>{escape(str(p))}</b>" for p in partes) + "</div>")
        st.button("Quitar selección", icon=":material/close:", key=f"quitar_{clave}", on_click=quitar)


# ──────────────────────────────────────────────────────────────
# ENCABEZADO
# ──────────────────────────────────────────────────────────────
md(f'<div class="topbar"><div class="crumb"><span>Alimentación para el Bienestar · </span>'
   f'<b>Programa de Abasto Rural 2026</b></div>'
   f'<div class="fecha"><i></i>Situación al {fecha_larga(hoy)}<span class="origen">· {origen}</span></div></div>')

a, b = st.columns([4, 1], vertical_alignment="bottom")
with a:
    md('<div class="titulo"><div><h1>Subgerencia de Abarrotes</h1>'
       '<p>Programa de Abasto Rural 2026 · proveedores, productos, artículos, marcas y unidades operativas.</p></div></div>')
with b, st.container(key="controles"):
    st.toggle(":material/dark_mode: Modo oscuro", key="oscuro")
    if cfg_od.get("link"):
        # El callback borra la caché antes de la siguiente ejecución, así el Excel se vuelve a descargar
        st.button("Actualizar datos", icon=":material/refresh:", key="actualizar", on_click=descargar_excel.clear)

# ──────────────────────────────────────────────────────────────
# BANNER PRINCIPAL: proveedores (CVE), productos (PROD.), artículos (ART.) y ofertas vigentes
# Cada clave cuenta una vez aunque se repita; al hacer clic en el recuadro se despliegan los nombres.
# Oferta vigente: artículo con inventario Oferta cuya vigencia no ha vencido (o no trae fecha).
# ──────────────────────────────────────────────────────────────
ofertas = datos[(datos["INVENTARIO"] == "Oferta") & ~(datos["VIG_FECHA"] < hoy)]
BANNER = [
    ("Proveedores", "proveedor", "CVE", datos.groupby("CVE")["PROV_CORTO"].first()),
    ("Productos", "producto", "PROD.", datos.groupby("PROD")["PRODUCTO"].first()),
    ("Artículos", "articulo", "ART.", datos.groupby("ART_ID")["ARTICULO"].first()),
    ("Ofertas vigentes", "estrella", "ART.", ofertas.groupby("ART_ID")["ARTICULO"].first()),
]
banner([(titulo, icono, len(por_clave), "", por_clave.dropna()) for titulo, icono, campo, por_clave in BANNER])

# ──────────────────────────────────────────────────────────────
# LECTURA RÁPIDA
# ──────────────────────────────────────────────────────────────
total = n_art(datos)
prov_n = datos.groupby("CVE")["ART_ID"].nunique().sort_values(ascending=False)
nombre_prov = datos.groupby("CVE")["PROV_CORTO"].first()
prods = (datos.groupby("PROD").agg(nombre=("PRODUCTO", "first"), art=("ART_ID", "nunique"), prov=("CVE", "nunique"))
         .sort_values(["art", "nombre"], ascending=[False, True]))
marcas = conteo(datos, "MARCA")
prov_marcas = datos.groupby("PROV_CORTO")["MARCA"].nunique().sort_values(ascending=False)
marca_prods = datos.groupby("MARCA")["PROD"].nunique().sort_values(ascending=False)
mas_prov = prods.sort_values(["prov", "art"], ascending=False).iloc[0]

hallazgos = [
    (T["acento"], "proveedor", f"El proveedor con más artículos es <b>{escape(nombre_prov[prov_n.index[0]])}</b>, "
                               f"con <b>{prov_n.iloc[0]:,}</b>."),
    (T["verde"], "producto", f"El producto con más artículos es <b>{escape(prods['nombre'].iloc[0])}</b>, "
                             f"con <b>{prods['art'].iloc[0]:,}</b>."),
    (T["oro"], "marca", f"La marca con más artículos es <b>{escape(marcas.index[0])}</b>, con <b>{marcas.iloc[0]:,}</b>."),
    (T["azul"], "marca", f"<b>{escape(prov_marcas.index[0])}</b> es el proveedor con más marcas: maneja "
                         f"<b>{prov_marcas.iloc[0]:,}</b>."),
    (T["verde_claro"], "producto", f"<b>{escape(marca_prods.index[0])}</b> es la marca con más productos distintos: "
                                   f"<b>{marca_prods.iloc[0]:,}</b>."),
    (GUINDA, "estrella", f"<b>{escape(mas_prov['nombre'])}</b> es el producto con más proveedores: "
                         f"<b>{mas_prov['prov']:,}</b> lo surten."),
]

md('<nav class="indice">' + "".join(f'<a href="#{k}"><span>{i}</span>{v}</a>'
                                    for i, (k, v) in enumerate(SECCIONES.items(), start=1)) + "</nav>")
seccion("resumen", "Lo más importante del programa de un vistazo.")
md('<div class="hallazgos">' + "".join(
    f'<div class="m-card hz" style="--c:{c}"><div class="ico">{svg(i, 17)}</div><div>{t}</div></div>'
    for c, i, t in hallazgos) + "</div>")

# ──────────────────────────────────────────────────────────────
# PROVEEDORES Y COBERTURA
# ──────────────────────────────────────────────────────────────
seccion("cobertura", "Haz clic en un recuadro para ver sus proveedores y búscalos por nombre · abajo, cada "
        "proveedor con sus tipos de inventario y su cobertura.")
# Un proveedor con artículos nacionales y regionales cuenta en los dos recuadros
banner([(f"Proveedores · Cobertura {alc.lower()}", icono, d["CVE"].nunique(), f"<b>{n_art(d):,}</b> artículos",
         d["PROV_CORTO"].dropna())
        for alc, icono in [("Nacional", "nacional"), ("Regional", "regional")]
        for d in [datos[datos["ALCANCE"] == alc]]], clase="dos")

with st.container():
    resumen = (datos.groupby("CVE").agg(prov=("PROV_CORTO", "first"), art=("ART_ID", "nunique"))
               .sort_values(["art", "prov"], ascending=[False, True]))
    tipos = datos.groupby(["CVE", "INVENTARIO"])["ART_ID"].nunique()
    cobertura = datos.groupby("CVE")[["ALCANCE", "REGIONES"]].apply(texto_cobertura)
    with tarjeta("Proveedores · todos", "c_prov",
                 "Cada proveedor con sus tipos de inventario y si su cobertura es nacional o regional.",
                 extra=f"{len(resumen):,} proveedores"):
        filas = []
        for cve, r in resumen.iterrows():
            fichas = "".join(chip(t, COLOR_INV.get(t, T["texto3"]), int(n))
                             for t, n in tipos[cve].sort_values(ascending=False).items())
            filas.append([escape(r.prov), fichas, cobertura[cve]])
        tabla(["Proveedor", "Tipos de inventario", "Cobertura"], filas, alto=520)

# ──────────────────────────────────────────────────────────────
# PRODUCTOS Y ARTÍCULOS: el tipo de inventario elegido filtra los productos y sus artículos
# ──────────────────────────────────────────────────────────────
seccion("productos", "<b>①</b> Si quieres, haz clic en un tipo de inventario · <b>②</b> elige un producto haciendo "
        "clic en su barra o búscalo por nombre · <b>③</b> haz clic en un proveedor de la dona para ver solo sus "
        "artículos.")
inv_n = conteo(datos, "INVENTARIO")
sel_inv = st.session_state.get("sel_inv")
sel_inv = sel_inv if sel_inv in inv_n.index else None
d_inv = datos[datos["INVENTARIO"] == sel_inv] if sel_inv else datos
prods_inv = (d_inv.groupby("PROD").agg(nombre=("PRODUCTO", "first"), art=("ART_ID", "nunique"))
             .sort_values(["art", "nombre"], ascending=[False, True]))
if st.session_state.get("sel_prod") not in prods_inv.index:
    st.session_state["sel_prod"] = prods_inv.index[0]
sel_prod = st.session_state["sel_prod"]
nombre_prod = prods_inv.at[sel_prod, "nombre"]
dp = d_inv[d_inv["PROD"] == sel_prod]
st.selectbox("Buscar producto", list(prods_inv.index), key="sel_prod", label_visibility="collapsed",
             format_func=lambda k: f"🔎  {prods_inv.at[k, 'nombre']}  ·  {prods_inv.at[k, 'art']:,} artículos")
st.write("")

clave_pp = f"sel_pp_{sel_inv or 'todos'}_{sel_prod}"
i, a, b = st.columns([2, 3, 2], gap="medium")
with i:
    prov_inv = datos.groupby("INVENTARIO")["CVE"].nunique()
    neg_inv = datos.groupby("INVENTARIO")["NEGOCIACION"].nunique()
    with tarjeta("Tipo de inventario · toda la cobertura", "c_inv", "Artículos por tipo de inventario.",
                 extra=[f"{total:,} artículos", f"{datos['NEGOCIACION'].nunique():,} hojas de negociación"]):
        # El clic se toma antes de dibujar: si cambia el tipo, la siguiente ejecución ya filtra con él
        dona_clic(inv_n, "c_inv_g", "sel_inv", colores=COLOR_INV, centro=total,
                  detalle={k: f"{prov_inv.get(k, 0):,} prov." for k in inv_n.index},
                  ovalo={k: f"{neg_inv.get(k, 0):,} {'hoja' if neg_inv.get(k, 0) == 1 else 'hojas'} neg."
                         for k in inv_n.index})
with a:
    with tarjeta("Todos los productos" + (f" · {sel_inv}" if sel_inv else ""), "pd_barras",
                 "Artículos distintos por producto, de mayor a menor.", extra=f"{len(prods_inv)} productos"):
        barras_clic(prods_inv["art"], f"pd_barras_g_{sel_inv or 'todos'}", "sel_prod",
                    etiquetas=list(prods_inv["nombre"]), vaciable=False, alto_max=470, n_corto=24)
with b:
    with tarjeta(f"Artículos de {nombre_prod}", "pd_dona", "Por proveedor.", extra=f"{n_art(dp):,} artículos"):
        sel_pp = dona_clic(conteo(dp, "PROV_CORTO"), f"pd_dona_{sel_inv or 'todos'}_{sel_prod}", clave_pp, alto=220,
                           max_items=10, centro=n_art(dp))
da = dp[dp["PROV_CORTO"] == sel_pp] if sel_pp and sel_pp != OTROS else dp
titulo_art = nombre_prod + (f" · {sel_inv}" if sel_inv else "") + (f" · {sel_pp}" if da is not dp else "")
if sel_inv or da is not dp:
    st.write("")
    barra_seleccion("prod", [x for x in [sel_inv, nombre_prod if da is not dp else None,
                                         sel_pp if da is not dp else None] if x], ["sel_inv", clave_pp])

st.write("")
with tarjeta(f"Artículos de {titulo_art}", "pd_tabla", "Presentación, inventario, cobertura y costo de cada artículo.",
             extra=f"{n_art(da):,} artículos"):
    filas = [[f"{escape(r.ARTICULO)}<small>ART. {escape(r.ART_ID)}</small>", escape(r.MARCA),
              escape(r.PRESENTACION), escape(r.PROV_CORTO),
              chip(r.INVENTARIO, COLOR_INV.get(r.INVENTARIO, T["texto3"])),
              "Nacional" if r.ALCANCE == "Nacional" else escape(r.COBERTURA), dinero(r.COSTO_UNIT)]
             for r in da.sort_values(["ARTICULO", "ART_ID"]).itertuples()]
    tabla(["Artículo", "Marca", "Presentación", "Proveedor", "Inventario", "Cobertura", "Costo unitario"],
          filas, numericas=(6,), alto=440)

# ──────────────────────────────────────────────────────────────
# UNIDADES OPERATIVAS: mapa y dona de productos en una tarjeta, regiones como botones y artículos de la región
# ──────────────────────────────────────────────────────────────
seccion("unidades", f"Haz clic en una <b>región</b> (el recuadro con sus unidades operativas) para ver sus productos "
        f"y la lista de sus artículos. Son {len(REGIONES)} gerencias regionales y {N_UNIDADES} unidades operativas.")
info_reg = estrellas_region(datos)
sel_reg = st.session_state.get("sel_region")
sel_reg = sel_reg if sel_reg in REGIONES or sel_reg == NAL else None
clave_up = f"sel_uprod_{sel_reg or 'todas'}"
regional = datos[datos["ALCANCE"] == "Regional"]
if sel_reg == NAL:
    d_reg = datos[datos["ALCANCE"] == "Nacional"]
elif sel_reg:
    d_reg = regional[regional["REGIONES"].map(lambda r: sel_reg in r)]
else:
    d_reg = regional
tipo_art = "nacionales" if sel_reg == NAL else "regionales"
sel_up = st.session_state.get(clave_up)
sel_up = sel_up if sel_up in set(d_reg["PRODUCTO"]) else None
nombre_reg = ("Nacional" if sel_reg == NAL else f"{sel_reg} · {REGIONES[sel_reg][0]}" if sel_reg
              else "todo el país")
if sel_reg or sel_up:
    barra_seleccion("reg", [x for x in [nombre_reg if sel_reg else None, sel_up] if x], ["sel_region", clave_up])
    st.write("")


def elegir_region(sig):
    """Clic en el recuadro de una región: la elige; otro clic sobre la misma la quita."""
    st.session_state["sel_region"] = None if st.session_state.get("sel_region") == sig else sig


a, b = st.columns([7, 3], gap="medium")
with a:
    with tarjeta(f"Mapa y productos {tipo_art} · {nombre_reg}", "u_mapa",
                 "Pasa el cursor sobre un estado para ver el producto estrella de su región; haz clic en un producto "
                 "de la dona para ver solo sus artículos abajo.", extra=f"{n_art(d_reg):,} artículos {tipo_art}"):
        g = conteo(d_reg, "PRODUCTO")
        m, d = st.columns([3, 2], gap="small")
        with m:
            mapa(info_reg, "u_mapa_g", resaltar=sel_reg, alto=430)
        with d:
            if d_reg.empty:
                md('<div class="lb-mas">Esta región no tiene artículos regionales.</div>')
            else:
                # Todos los productos, sin agrupar en «Otros»; al centro, cuántos productos son
                dona_clic(g, f"u_dona_{sel_reg or 'todas'}", clave_up, alto=230, max_items=len(g), agrupar=False,
                          centro=len(g), unidad_centro="productos", detalle={k: "art." for k in g.index},
                          alto_ley=330)
with b:
    with tarjeta("Regiones y unidades operativas", "u_botones", "Haz clic en un recuadro.", extra=f"{len(REGIONES)} regiones"):
        with st.container(height=560, border=False):
            with st.container(key=f"regbtn_{NAL}_{'on' if sel_reg == NAL else 'off'}"):
                md(f'<div class="rb-h"><span class="sig" style="background:{NEGRO};color:#fff">{NAL}</span>'
                   f'<b>Nacional</b><small>todo el país</small></div>'
                   f'<div class="rb-u">Artículos que llegan a las {len(REGIONES)} regiones y sus {N_UNIDADES} '
                   f'unidades operativas</div>')
                st.button("Elegir Nacional", key=f"regclic_{NAL}", on_click=elegir_region, args=(NAL,))
            for sig, (nombre, clave_r, unidades) in REGIONES.items():
                with st.container(key=f"regbtn_{sig}_{'on' if sig == sel_reg else 'off'}"):
                    md(f'<div class="rb-h"><span class="sig" style="background:{COLOR_REGION[sig]}">{sig}</span>'
                       f'<b>{nombre}</b><small>clave {clave_r}</small></div>'
                       f'<div class="rb-u">{" · ".join(escape(u.removeprefix("U.O. ")) for u, _, _ in unidades)}</div>')
                    st.button(f"Elegir {nombre}", key=f"regclic_{sig}", on_click=elegir_region, args=(sig,))

st.write("")
da_reg = d_reg[d_reg["PRODUCTO"] == sel_up] if sel_up else d_reg
with tarjeta(f"Artículos {tipo_art} · {nombre_reg}" + (f" · {sel_up}" if sel_up else ""), "u_tabla",
             "Artículos con cobertura nacional: llegan a todas las regiones." if sel_reg == NAL else
             "Artículos con cobertura regional que llegan a esta región (los nacionales llegan a todas).",
             extra=f"{n_art(da_reg):,} artículos"):
    filas = [[f"{escape(r.ARTICULO)}<small>ART. {escape(r.ART_ID)}</small>", escape(r.PRODUCTO), escape(r.MARCA),
              escape(r.PRESENTACION), escape(r.PROV_CORTO), chip(r.INVENTARIO, COLOR_INV.get(r.INVENTARIO, T["texto3"])),
              "Nacional" if r.ALCANCE == "Nacional" else escape(" ".join(r.REGIONES)), dinero(r.COSTO_UNIT)]
             for r in da_reg.sort_values(["PRODUCTO", "ARTICULO", "ART_ID"]).itertuples()]
    tabla(["Artículo", "Producto", "Marca", "Presentación", "Proveedor", "Inventario", "Regiones", "Costo unitario"],
          filas, numericas=(7,), alto=440)

md('<div class="pie">Subgerencia de Abarrotes · Alimentación para el Bienestar · Gobierno de México</div>')
