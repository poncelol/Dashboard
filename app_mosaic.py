"""
Programa de Abasto Rural 2026 · Tablero ejecutivo (diseño estilo Mosaic)
Subgerencia de Abarrotes · Alimentación para el Bienestar

Ejecutar:   streamlit run app_mosaic.py

Origen de los datos (en este orden):
  1. OneDrive / SharePoint: si en los secretos (.streamlit/secrets.toml o "Secrets" en
     Streamlit Community Cloud) existe [onedrive] link = "...", el Excel se descarga de ahí
     y se vuelve a descargar solo cada [onedrive] minutos (10 por omisión).
  2. El archivo PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx junto a este archivo.

Diseño inspirado en Mosaic Lite (cruip/tailwind-dashboard-template): sidebar con
navegación, header superior, tarjetas rounded-xl con encabezado, KPIs con etiqueta
en mayúsculas y modo oscuro. El acento violeta de Mosaic se sustituye por el
guinda institucional.
"""
import base64
import io
import re
from contextlib import contextmanager
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
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
    initial_sidebar_state="expanded",
)

ARCHIVO = Path(__file__).parent / "PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx"

# En la nube el servidor corre en UTC: fechas y horas se muestran en hora del centro de México
try:
    ZONA = ZoneInfo("America/Mexico_City")
except Exception:  # noqa: BLE001  (Windows sin el paquete tzdata)
    ZONA = None


def ahora():
    return datetime.now(ZONA) if ZONA else datetime.now()

# hoja del Excel -> nombre en el tablero (orden de "embudo")
HOJAS = {
    "LISTA DE ESPERA": "Lista de espera",
    "ELABORACION": "En elaboración",
    "PROCESO DE FIRMA": "Proceso de firma",
    "VIGENTES": "Vigentes",
    "HISTORICAS": "Históricas",
    "CANCELADAS": "Canceladas",
}
GLOBAL = "Global"
CANASTA_LABEL = {"CB": "Canasta básica", "COMP": "Complementario", "PAR": "PAR"}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Paleta institucional gob.mx
GUINDA, GUINDA_MED = "#611232", "#9D2449"
VERDE_MED, VERDE_CLARO = "#235B4E", "#3E8E7E"
ORO, ORO_CLARO = "#A57F2C", "#BC955C"
AZUL = "#2E6B9E"

# Íconos de línea (trazos de Lucide, licencia ISC): heredan el color del texto con currentColor
ICONOS = {
    "proveedor": '<path d="M2 20a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8l-7 5V8l-7 5V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/>'
                 '<path d="M17 18h1"/><path d="M12 18h1"/><path d="M7 18h1"/>',
    "negociacion": '<rect width="8" height="4" x="8" y="2" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6'
                   'a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="M12 11h4"/><path d="M12 16h4"/><path d="M8 11h.01"/>'
                   '<path d="M8 16h.01"/>',
    "producto": '<circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 '
                '2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/>',
    "articulo": '<path d="m7.5 4.27 9 5.15"/><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 '
                '0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>',
    "calendario": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
    "marca": '<path d="M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 '
             '2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r=".5" fill="currentColor"/>',
    "precio": '<line x1="12" x2="12" y1="2" y2="22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
    "idea": '<path d="M15 14c.2-1 .7-1.7 1.5-2.5 1-.9 1.5-2.2 1.5-3.5A6 6 0 0 0 6 8c0 1 .2 2.2 1.5 3.5.7.7 1.3 1.5 1.5 2.5"/>'
            '<path d="M9 18h6"/><path d="M10 22h4"/>',
    "tendencia": '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
    "estrella": '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
    "pastel": '<path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
    "enlace": '<path d="M9 17H7A5 5 0 0 1 7 7h2"/><path d="M15 7h2a5 5 0 1 1 0 10h-2"/><line x1="8" x2="16" y1="12" y2="12"/>',
    "alerta": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/>'
              '<path d="M12 17h.01"/>',
    "ok": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
}


def svg(nombre, tam=18):
    return (f'<svg width="{tam}" height="{tam}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICONOS[nombre]}</svg>')


PAGINAS = {
    "Panorama": ":material/dashboard:",
    "Proveedores": ":material/factory:",
    "Productos": ":material/shopping_cart:",
    "Cobertura": ":material/map:",
    "Vigencia": ":material/event:",
    "Duplicados": ":material/content_copy:",
}

# Campos por los que se pueden buscar duplicados (nombre visible -> columnas que forman el valor)
CLAVES_DUP = {
    "Código de barras": ["CODIGO_BARRAS"],
    "Artículo": ["ARTICULO"],
    "Artículo + marca": ["ARTICULO", "MARCA"],
    "Artículo + proveedor": ["ARTICULO", "PROVEEDOR"],
    "Clave de artículo": ["ART_ID"],
}
MODOS_DUP = ["Todos", "Solo duplicados", "Sin duplicados"]

# Íconos de Material Symbols para los botones (Streamlit los dibuja con la sintaxis :material/nombre:)
ICONOS_BOTON = {
    "Global": "public",
    "Lista de espera": "hourglass_empty",
    "En elaboración": "edit_note",
    "Proceso de firma": "draw",
    "Vigentes": "verified",
    "Históricas": "history",
    "Canceladas": "cancel",
    "Todos": "select_all",
    "Solo duplicados": "content_copy",
    "Sin duplicados": "filter_1",
}


def con_icono(texto):
    icono = ICONOS_BOTON.get(texto)
    return f":material/{icono}: {texto}" if icono else texto

# Tokens de tema (grises de Mosaic; en oscuro, colores de datos más claros para que contrasten)
TEMAS = {
    "claro": dict(
        fondo="#f3f4f6", card="#ffffff", sutil="#f9fafb", linea="#f3f4f6", borde="#e5e7eb",
        texto="#1f2937", texto2="#4b5563", texto3="#6b7280", tenue="#9ca3af",
        acento=GUINDA_MED, acento_rgb="157 36 73", ok="#239f52", alerta="#e63939", otros="#d1d5db",
        paleta=[GUINDA_MED, VERDE_MED, ORO, AZUL, VERDE_CLARO, GUINDA, "#7A5C1E", "#B5547A", "#5F6B6D"],
        verde=VERDE_MED, oro=ORO, azul=AZUL, verde_claro=VERDE_CLARO,
        heat=["#f9fafb", ORO_CLARO, GUINDA], sombra="0 1px 1px 0 rgb(0 0 0 / .05), 0 1px 2px 0 rgb(0 0 0 / .02)",
    ),
    "oscuro": dict(
        fondo="#111827", card="#1f2937", sutil="#263142", linea="#2d3748", borde="#374151",
        texto="#f3f4f6", texto2="#d1d5db", texto3="#9ca3af", tenue="#6b7280",
        acento="#E0698C", acento_rgb="224 105 140", ok="#4bd37d", alerta="#ff7474", otros="#4b5563",
        paleta=["#E0698C", "#4FB39A", "#D9AE55", "#6AA6DE", "#8FD3C2", "#C77D93", "#B8955A", "#E59AB3", "#94A3B8"],
        verde="#4FB39A", oro="#D9AE55", azul="#6AA6DE", verde_claro="#8FD3C2",
        heat=["#263142", GUINDA_MED, ORO_CLARO], sombra="0 1px 2px 0 rgb(0 0 0 / .35)",
    ),
}

# ──────────────────────────────────────────────────────────────
# CARGA Y LIMPIEZA (misma lógica que app.py)
# ──────────────────────────────────────────────────────────────
RENOMBRAR = {
    "NEGOCIACION": "NEGOCIACION", "CVE": "CVE", "PROVEEDOR": "PROVEEDOR",
    "PRODUCTOS": "PRODUCTO", "ART.": "ART_ID", "ARTICULOS": "ARTICULO",
    "MARCA": "MARCA", "CAP": "CAP", "GRAM": "GRAM", "U/M": "UM",
    "COSTO DE ADQUISICIÓN 2026": "COSTO_CAJA",
    "COSTO UNITARIO DE ADQUISICIÓN 2026": "COSTO_UNIT",
    "UTILIDAD": "UTILIDAD", "VENTA S/IMP.": "VENTA_SIN", "VENTA C/IMP.": "VENTA_CON",
    "IMPUESTO": "IMPUESTO", "REPRESENTANTE": "REPRESENTANTE",
    "CODIGO BARRAS": "CODIGO_BARRAS", "CANASTA BÁSICA": "CANASTA", "CB": "CANASTA",
    "TIPO DE INVENTARIO": "INVENTARIO", "COBERTURA": "COBERTURA",
    "VIGENCIA": "VIGENCIA", "JUSTIFICACIÓN": "JUSTIFICACION",
}
NUMERICAS = ["CAP", "GRAM", "COSTO_CAJA", "COSTO_UNIT", "UTILIDAD", "VENTA_SIN", "VENTA_CON"]


def nombre_corto(nombre: str) -> str:
    corto = re.sub(r",?\s+S\.\s?(A\.|DE\s?R\.L\.).*$", "", str(nombre).strip(), flags=re.I)
    return corto.strip(" ,").title() or str(nombre)


@st.cache_data(show_spinner="Preparando el tablero…")
def cargar(contenido: bytes):
    xls = pd.ExcelFile(io.BytesIO(contenido))
    articulos, negociaciones = [], []

    for hoja, estatus in HOJAS.items():
        if hoja not in xls.sheet_names:
            continue
        raw = pd.read_excel(xls, sheet_name=hoja, header=3)
        cols = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]
        raw.columns = cols
        raw = raw.rename(columns=RENOMBRAR)
        raw = raw.loc[:, ~raw.columns.str.startswith("Unnamed")]
        raw = raw.loc[:, ~raw.columns.duplicated()]
        raw["ESTATUS"] = estatus
        raw["NEGOCIACION"] = pd.to_numeric(raw["NEGOCIACION"], errors="coerce")

        es_articulo = raw["ART_ID"].notna()
        cab = raw[~es_articulo & raw["NEGOCIACION"].notna()]
        negociaciones.append(
            cab[[c for c in ["ESTATUS", "NEGOCIACION", "PROVEEDOR", "REPRESENTANTE", "COBERTURA"] if c in cab.columns]]
        )
        articulos.append(raw[es_articulo])

    df = pd.concat(articulos, ignore_index=True)
    neg = pd.concat(negociaciones, ignore_index=True)

    for c in NUMERICAS:
        if c not in df:
            df[c] = pd.NA
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["UTILIDAD"] = df["UTILIDAD"].fillna(0.15)
    df["VENTA_SIN"] = df["VENTA_SIN"].fillna(df["COSTO_UNIT"] / (1 - df["UTILIDAD"]))

    for c in ["PROVEEDOR", "PRODUCTO", "ARTICULO", "MARCA", "UM", "IMPUESTO"]:
        df[c] = df[c].astype(str).str.strip().str.upper().replace({"NAN": pd.NA, "": pd.NA})
    df["MARCA"] = df["MARCA"].fillna("SIN MARCA").str.title()
    df["PRODUCTO"] = df["PRODUCTO"].fillna("SIN CLASIFICAR").str.capitalize()
    df["IMPUESTO"] = df["IMPUESTO"].fillna("Sin dato")
    df["PROV_CORTO"] = df["PROVEEDOR"].map(nombre_corto)
    df["NEGOCIACION"] = df["NEGOCIACION"].astype("Int64").astype(str)
    df["CANASTA"] = (
        df["CANASTA"].astype(str).str.strip().str.upper()
        .map(lambda x: CANASTA_LABEL.get(x, "Sin clasificar" if x in ("NAN", "<NA>", "") else x))
    )
    df["COBERTURA"] = df["COBERTURA"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["COBERTURA"] = df["COBERTURA"].str.replace(r"\s+", " ", regex=True).str.strip().str.upper()

    df["VIG_FECHA"] = pd.to_datetime(df["VIGENCIA"], errors="coerce")

    # Código de barras como texto; "0" o vacío es un dato faltante, no un código repetido
    if "CODIGO_BARRAS" not in df:
        df["CODIGO_BARRAS"] = pd.NA
    cb = df["CODIGO_BARRAS"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    vacio = cb.isna() | cb.str.fullmatch(r"0*|NAN|nan|<NA>|None").fillna(True).astype(bool)
    df["CODIGO_BARRAS"] = cb.where(~vacio, pd.NA)

    neg["NEGOCIACION"] = neg["NEGOCIACION"].astype("Int64").astype(str)
    return df, neg


# ──────────────────────────────────────────────────────────────
# SIDEBAR (va primero: el tema elegido define el CSS)
# ──────────────────────────────────────────────────────────────
def md(texto, destino=st):
    destino.markdown(texto, unsafe_allow_html=True)


with st.sidebar:
    md('<div class="sb-logo"><div class="mark">AR</div>'
       '<div><div class="t">Abasto Rural</div><div class="s">Programa 2026</div></div></div>')
    # El tema se guarda en la URL (?tema=oscuro): sobrevive a recargas y se puede compartir
    if "oscuro" not in st.session_state:
        st.session_state["oscuro"] = st.query_params.get("tema") == "oscuro"
    oscuro = st.toggle(":material/dark_mode: Modo oscuro", key="oscuro")
    st.query_params["tema"] = "oscuro" if oscuro else "claro"

T = TEMAS["oscuro" if oscuro else "claro"]
COLOR_ESTATUS = {
    "Lista de espera": T["texto3"], "En elaboración": T["oro"], "Proceso de firma": T["azul"],
    "Vigentes": T["verde"], "Históricas": "#B8955A" if oscuro else "#7A5C1E", "Canceladas": T["acento"],
}

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

/* ── Sidebar ─────────────────────────────────────────── */
[data-testid="stSidebar"] { background: var(--card); border-right: none; box-shadow: var(--sombra);
  border-radius: 0 16px 16px 0; }
[data-testid="stSidebarUserContent"] { padding-top: 1rem; }
.sb-logo { display: flex; align-items: center; gap: 12px; padding: 0 4px; margin-bottom: 16px; }
.sb-logo .mark { width: 34px; height: 34px; border-radius: 9px; flex-shrink: 0;
  background: linear-gradient(135deg, #611232 0%, #9D2449 100%); color: #fff;
  display: grid; place-items: center; font-weight: 700; font-size: .95rem; }
.sb-logo .t { font-weight: 700; color: var(--texto); font-size: .95rem; line-height: 1.2; }
.sb-logo .s { font-size: .75rem; color: var(--tenue); font-weight: 500; }
.sb-h { font-size: .7rem; text-transform: uppercase; letter-spacing: .05em; color: var(--tenue);
  font-weight: 600; padding-left: 12px; margin: 22px 0 8px 0; }
.sb-foot { font-size: .75rem; color: var(--tenue); padding: 18px 12px 0 12px; line-height: 1.5; }

/* Controles de Streamlit 1.66 (react-aria). Se apuntan por data-testid, role y estados
   (aria-checked / data-selected), no por clases internas, para que sigan el tema claro/oscuro. */

/* Navegación (st.radio con apariencia de enlaces de Mosaic) */
[data-testid="stSidebar"] [data-testid="stRadioGroup"] { gap: 2px; }
[data-testid="stSidebar"] [data-testid="stRadioOption"] { width: 100%; padding: 8px 12px; border-radius: 8px;
  margin: 0; cursor: pointer; transition: background .15s; }
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:not([data-testid]) { display: none; }
[data-testid="stSidebar"] [data-testid="stRadioOption"] p { font-size: .875rem; font-weight: 500; color: var(--texto2); }
[data-testid="stSidebar"] [data-testid="stRadioOption"] span[role="img"] { color: var(--tenue); margin-right: 10px; }
[data-testid="stSidebar"] [data-testid="stRadioOption"]:hover { background: var(--sutil); }
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] {
  background: linear-gradient(90deg, rgb(var(--acento-rgb) / .14) 0%, rgb(var(--acento-rgb) / .04) 100%); }
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] p { color: var(--texto); }
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] span[role="img"] { color: var(--acento); }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { font-size: .8rem; font-weight: 500; color: var(--texto2); }

/* Pills de etapa: botones tipo Mosaic */
[data-testid="stButtonGroup"] > div { gap: 6px; }
[data-testid="stButtonGroup"] button[data-variant] { background: var(--card) !important;
  border: 1px solid var(--borde) !important; color: var(--texto2) !important; border-radius: 8px !important;
  font-size: .8rem; font-weight: 500; padding: 4px 10px !important; min-height: 0; transition: all .15s; }
[data-testid="stButtonGroup"] button[data-variant] * { color: inherit !important; }
[data-testid="stButtonGroup"] button[data-variant] span[role="img"] { font-size: 1.05rem; margin-right: 2px; opacity: .85; }
[data-testid="stButtonGroup"] button[data-variant][aria-checked="true"] span[role="img"] { opacity: 1; }
[data-testid="stButtonGroup"] button[data-variant]:hover { border-color: var(--tenue) !important;
  color: var(--texto) !important; }
[data-testid="stButtonGroup"] button[data-variant][aria-checked="true"] { background: var(--acento) !important;
  border-color: var(--acento) !important; color: #fff !important; }

/* Interruptor (modo oscuro) */
[data-testid="stCheckbox"] label > div:not([data-testid]) { background: var(--borde) !important; }
[data-testid="stCheckbox"][data-selected="true"] label > div:not([data-testid]) { background: var(--acento) !important; }
[data-testid="stCheckbox"] label > div:not([data-testid]) > div { background: #fff !important; }
[data-testid="stCheckbox"] p, [data-testid="stCheckbox"] span[role="img"] { color: var(--texto2) !important; }

/* Entradas: multiselect, selectbox y texto */
[data-testid="stMultiSelect"] div[role="group"][data-rac], [data-testid="stSelectbox"] div[role="group"][data-rac],
[data-testid="stTextInputRootElement"] { background: var(--sutil) !important; border: 1px solid var(--borde) !important;
  border-radius: 8px !important; }
[data-testid="stMultiSelect"] div[role="group"][data-focus-within="true"],
[data-testid="stSelectbox"] div[role="group"][data-focus-within="true"],
[data-testid="stTextInputRootElement"]:focus-within { border-color: var(--acento) !important; }
[data-testid="stMultiSelect"] input, [data-testid="stSelectbox"] input, [data-testid="stTextInputField"] {
  color: var(--texto) !important; background: transparent !important; caret-color: var(--acento); }
[data-testid="stMultiSelect"] input::placeholder, [data-testid="stSelectbox"] input::placeholder,
[data-testid="stTextInputField"]::placeholder { color: var(--tenue) !important; opacity: 1; }
[data-testid="stMultiSelect"] button, [data-testid="stSelectbox"] button { color: var(--texto3) !important;
  background: transparent !important; }
[data-testid="stMultiSelectTagsContainer"] [data-tag] { background: var(--acento) !important; color: #fff !important;
  border-radius: 6px !important; }
[data-testid="stMultiSelectTagsContainer"] [data-tag] * { color: #fff !important; }

/* Listas desplegables (se abren fuera de la app, en un portal) */
[data-testid$="Dropdown"] { background: var(--card) !important; border: 1px solid var(--borde) !important;
  border-radius: 8px !important; box-shadow: 0 8px 24px rgb(0 0 0 / .18) !important; }
[data-testid$="Dropdown"] [role="listbox"] { background: var(--card) !important; }
[data-testid$="Dropdown"] [role="option"] { color: var(--texto2) !important; background: transparent !important; }
[data-testid$="Dropdown"] [role="option"][data-focused="true"],
[data-testid$="Dropdown"] [role="option"]:hover { background: var(--sutil) !important; color: var(--texto) !important; }
[data-testid$="Dropdown"] [role="option"][aria-selected="true"] { color: var(--acento) !important; font-weight: 600; }

/* Botón para abrir/cerrar la barra lateral */
[data-testid="stSidebarCollapseButton"] button, [data-testid="stExpandSidebarButton"],
[data-testid="stBaseButton-headerNoPadding"] { color: var(--texto3) !important; }

/* ── Header superior (ancho completo) ─────────────────── */
.topbar { background: var(--card); display: flex; align-items: center; justify-content: space-between;
  gap: 16px; height: 64px; margin-bottom: 28px;
  box-shadow: 0 0 0 100vmax var(--card), 0 1px 0 100vmax var(--borde); clip-path: inset(0 -100vmax -1px -100vmax); }
.topbar .crumb { font-size: .875rem; color: var(--texto3); }
.topbar .crumb b { color: var(--texto); font-weight: 600; }
.topbar .fecha { display: inline-flex; align-items: center; gap: 8px; font-size: .8rem; color: var(--texto2);
  font-weight: 500; border: 1px solid var(--borde); border-radius: 8px; padding: 6px 12px; white-space: nowrap; }
.topbar .fecha .origen { color: var(--tenue); }
[data-testid="stSidebar"] button[data-testid^="stBaseButton"]:not([data-testid*="header"]) {
  background: var(--sutil) !important; border: 1px solid var(--borde) !important; color: var(--texto2) !important;
  border-radius: 8px !important; font-size: .8rem; }
[data-testid="stSidebar"] button[data-testid^="stBaseButton"]:not([data-testid*="header"]):hover {
  border-color: var(--acento) !important; color: var(--acento) !important; }
[data-testid="stSidebar"] button[data-testid^="stBaseButton"]:not([data-testid*="header"]) * { color: inherit !important; }
.topbar .fecha i { width: 8px; height: 8px; border-radius: 50%; background: var(--ok);
  box-shadow: 0 0 0 3px rgb(62 201 114 / .2); }

/* ── Título de página ───────────────────────────────── */
.titulo { display: flex; justify-content: space-between; align-items: flex-end; gap: 16px;
  flex-wrap: wrap; margin-bottom: 24px; }
.titulo h1 { font-size: 1.88rem; font-weight: 700; margin: 0; padding: 0; line-height: 1.33; }
.titulo p { margin: 2px 0 0 0; font-size: .875rem; color: var(--texto3); }
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
.badge.alerta { color: var(--alerta); background: rgb(255 86 86 / .15); }

/* Embudo */
.etapas { display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; padding: 4px 0 8px 0; }
.etapa { border-radius: 10px; padding: 10px 12px; border: 1px solid transparent; }
.etapa.sel { border-color: var(--acento); background: rgb(var(--acento-rgb) / .06); }
.etapa .t { display: flex; align-items: center; gap: 8px; font-size: .8rem; color: var(--texto3); font-weight: 500; }
.etapa .t i { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.etapa .n { font-size: 1.5rem; font-weight: 700; color: var(--texto); margin-top: 4px; }
.etapa .n small { font-size: .75rem; font-weight: 500; color: var(--tenue); margin-left: 4px; }
.etapa .pista { height: 6px; background: var(--linea); border-radius: 999px; margin-top: 8px; overflow: hidden; }
.etapa .pista div { height: 100%; border-radius: 999px; }

/* Lista estilo "Recent Activity" */
.lista { list-style: none; margin: 0; padding: 0; }
.lista li { display: flex; gap: 12px; padding: 12px 0; border-bottom: 1px solid var(--linea);
  font-size: .875rem; color: var(--texto2); align-items: center; }
.lista li:last-child { border-bottom: none; }
.lista .ico { width: 36px; height: 36px; border-radius: 50%; flex-shrink: 0; display: grid; place-items: center;
  color: #fff; font-weight: 700; font-size: .9rem; }
.lista b { color: var(--texto); font-weight: 600; }

/* Barras con nombre (estilo "Reasons for Refunds"): cada producto con nombre, cantidad y % */
.lb { list-style: none; margin: 0; padding: 0; }
.lb.scroll { overflow-y: auto; padding-right: 6px; }
.lb li { padding: 6px 0; }
.lb-top { display: flex; justify-content: space-between; gap: 12px; font-size: .84rem; }
.lb-n { color: var(--texto); font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.lb-n .rk { display: inline-block; width: 22px; color: var(--tenue); font-weight: 600; font-size: .75rem; }
.lb-v { color: var(--texto2); font-weight: 600; white-space: nowrap; }
.lb-v em { font-style: normal; color: var(--tenue); font-weight: 500; margin-left: 6px; }
.lb-pista { height: 6px; background: var(--linea); border-radius: 999px; margin-top: 5px; overflow: hidden; }
.lb-pista div { height: 100%; border-radius: 999px; }
.lb-mas { font-size: .78rem; color: var(--tenue); padding-top: 8px; margin-top: 4px; border-top: 1px dashed var(--borde); }
.lb.scroll::-webkit-scrollbar { width: 6px; }
.lb.scroll::-webkit-scrollbar-thumb { background: var(--borde); border-radius: 999px; }

/* Rejilla de mini tarjetas (mezcla por proveedor) */
.grid3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
.mix { padding: 16px 18px; border: 1px solid var(--linea); box-shadow: none !important; }
.mix-h { display: flex; flex-direction: column; margin-bottom: 8px; }
.mix-h b { color: var(--texto); font-size: .9rem; font-weight: 600; }
.mix-h span { font-size: .75rem; color: var(--tenue); }

/* Tabla estilo "Top Channels" */
.tabla-wrap { overflow-x: auto; }
.tabla { width: 100%; border-collapse: collapse; font-size: .875rem; }
.tabla thead th { font-size: .7rem; text-transform: uppercase; color: var(--tenue); font-weight: 600;
  background: var(--sutil); padding: 8px; text-align: left; border: none; }
.tabla thead th:first-child { border-radius: 6px 0 0 6px; }
.tabla thead th:last-child { border-radius: 0 6px 6px 0; }
.tabla td { padding: 10px 8px; border: none; border-bottom: 1px solid var(--linea); color: var(--texto2); font-weight: 500; }
.tabla tr:last-child td { border-bottom: none; }
.tabla td.fuerte { color: var(--texto); }
.tabla .num { text-align: right; }

/* Guía de lectura */
.guia { display: flex; gap: 10px; align-items: flex-start; background: rgb(188 149 92 / .12);
  border-radius: 8px; padding: 10px 14px; font-size: .82rem; color: var(--texto2); margin: 0 0 6px 0; }
.guia svg { color: var(--oro); flex-shrink: 0; margin-top: 1px; }
.guia b { color: var(--texto); }

/* Leyenda HTML de las donas (se acomoda sola al ancho, nunca tapa la gráfica) */
.ley { list-style: none; margin: 4px 0 0 0; padding: 0; }
.ley li { display: flex; align-items: center; gap: 8px; padding: 7px 0; border-bottom: 1px solid var(--linea); font-size: .84rem; }
.ley li:last-child { border-bottom: none; }
.ley i { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
.ley .n { flex: 1; min-width: 0; color: var(--texto); font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ley .v { color: var(--texto2); font-weight: 600; white-space: nowrap; }
.ley em { font-style: normal; color: var(--tenue); font-weight: 500; display: inline-block; min-width: 44px; text-align: right; }

/* ── Responsivo ─────────────────────────────────────── */
@media (max-width: 1100px) {
  .etapas { grid-template-columns: repeat(3, 1fr); }
  .grid3 { grid-template-columns: repeat(2, 1fr); }
  .kpis { grid-template-columns: repeat(2, 1fr); gap: 16px; }
}
@media (max-width: 900px) {
  .grid3 { grid-template-columns: 1fr; }
  .titulo h1 { font-size: 1.5rem; }
  .topbar .crumb span { display: none; }
}
@media (max-width: 640px) {
  .block-container { padding-left: 1rem !important; padding-right: 1rem !important; }
  .topbar { height: auto; min-height: 56px; flex-wrap: wrap; gap: 8px; padding: 10px 0; margin-bottom: 18px; }
  .titulo { margin-bottom: 16px; }
  .titulo > div:last-child .etiq:first-child { margin-left: 0; }
  .etapas { grid-template-columns: repeat(2, 1fr); gap: 8px; }
  .kpis { gap: 12px; margin-bottom: 16px; }
  .kpi { padding: 14px; gap: 10px; }
  .kpi .ico { width: 32px; height: 32px; }
  .kpi .v { font-size: 1.35rem; }
  div[class*="st-key-card_"] { padding: 0 14px 12px 14px !important; }
  .card-h { margin: 0 -14px 8px -14px; padding: 14px; flex-wrap: wrap; }
  .tabla { font-size: .8rem; }
}
@media (max-width: 420px) { .kpis { grid-template-columns: 1fr; } .etapas { grid-template-columns: 1fr 1fr; } }
"""
md(f"<style>{VARIABLES}{CSS}</style>")


# ──────────────────────────────────────────────────────────────
# COMPONENTES DE PRESENTACIÓN
# ──────────────────────────────────────────────────────────────
def fecha_larga(d) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def dinero(x) -> str:
    return "–" if pd.isna(x) else f"${x:,.2f}"


@contextmanager
def tarjeta(titulo, clave, sub="", extra=""):
    """Tarjeta de Mosaic: fondo de tarjeta, rounded-xl, encabezado con línea inferior."""
    with st.container(border=True, key=f"card_{clave}"):
        md(f'<div class="card-h"><div><h2>{escape(titulo)}</h2>{f"<p>{sub}</p>" if sub else ""}</div>'
           f'{f"<span class=extra>{extra}</span>" if extra else ""}</div>')
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


def guia(texto):
    md(f'<div class="guia">{svg("idea", 16)}<span>{texto}</span></div>')


def tabla(encabezados, filas, numericas=()):
    th = "".join(f'<th class="{"num" if i in numericas else ""}">{h}</th>' for i, h in enumerate(encabezados))
    tr = "".join(
        "<tr>" + "".join(
            f'<td class="{"num" if i in numericas else ""} {"fuerte" if i == 0 else ""}">{c}</td>'
            for i, c in enumerate(fila)) + "</tr>"
        for fila in filas)
    md(f'<div class="tabla-wrap"><table class="tabla"><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table></div>')


def lista_barras(serie, color=None, total=None, alto=None, ranking=True):
    """Cada elemento con su nombre, cantidad y % del total: nada se esconde en "Otros"."""
    if serie.empty:
        return f'<div class="lb-mas">Sin datos.</div>'
    color = color or T["acento"]
    total = total or serie.sum()
    tope = serie.max()
    filas = "".join(
        f'<li><div class="lb-top"><span class="lb-n" title="{escape(str(k))}">'
        f'{f"<span class=rk>{n}</span>" if ranking else ""}{escape(str(k))}</span>'
        f'<span class="lb-v">{v:,}<em>{v / total:.1%}</em></span></div>'
        f'<div class="lb-pista"><div style="width:{v / tope:.1%};background:{color}"></div></div></li>'
        for n, (k, v) in enumerate(serie.items(), start=1))
    alto_css = f' style="max-height:{alto}px"' if alto else ""
    return f'<ul class="lb{" scroll" if alto else ""}"{alto_css}>{filas}</ul>'


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
        hoverlabel=dict(bgcolor=T["card"], bordercolor=T["borde"], font_size=12, font_family="Inter",
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


def dona(serie, clave, alto=220, colores=None, max_items=8):
    """Dona + leyenda HTML (nombre, cantidad y %). Solo para categorías con pocas opciones."""
    s = serie.sort_values(ascending=False)
    if len(s) > max_items:
        s = pd.concat([s.head(max_items - 1), pd.Series({"Otros": s.iloc[max_items - 1:].sum()})])
    cols = list(colores or T["paleta"])
    cols = [cols[i % len(cols)] for i in range(len(s))]
    if "Otros" in s.index:
        cols[list(s.index).index("Otros")] = T["otros"]
    fig = go.Figure(go.Pie(
        labels=s.index, values=s.values, hole=0.72, sort=False,
        marker=dict(colors=cols, line=dict(color=T["card"], width=3)),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value:,} artículos · %{percent}<extra></extra>",
    ))
    fig.add_annotation(
        text=f"<b>{int(s.sum()):,}</b><br><span style='font-size:11px;color:{T['tenue']}'>artículos</span>",
        showarrow=False, font=dict(family="Inter, sans-serif", size=20, color=T["texto"]))
    estilo(fig, alto)
    fig.update_layout(margin=dict(l=0, r=0, t=4, b=4))
    mostrar(fig, clave)
    md('<ul class="ley">' + "".join(
        f'<li><i style="background:{c}"></i><span class="n" title="{escape(str(k))}">{escape(str(k))}</span>'
        f'<span class="v">{v:,}<em>{v / s.sum():.0%}</em></span></li>'
        for (k, v), c in zip(s.items(), cols)) + "</ul>")


def resumen_productos(df):
    """Una fila por producto: artículos, proveedores, venta mediana y canasta predominante."""
    p = df.groupby("PRODUCTO").agg(art=("ART_ID", "count"), prov=("PROV_CORTO", "nunique"),
                                   precio=("VENTA_SIN", "median"))
    p["canasta"] = df.groupby("PRODUCTO")["CANASTA"].agg(lambda s: s.mode().iat[0])
    return p


def concentracion(serie, alto=340):
    """Curva acumulada: con cuántos productos (de mayor a menor) se junta qué % del surtido."""
    acum = serie.cumsum() / serie.sum()
    x = list(range(1, len(serie) + 1))
    relleno = f"rgba({T['acento_rgb'].replace(' ', ',')},0.12)"
    fig = go.Figure(go.Scatter(
        x=x, y=acum.values, mode="lines", line=dict(color=T["acento"], width=2.5),
        fill="tozeroy", fillcolor=relleno,
        customdata=list(zip(serie.index, serie.values)),
        hovertemplate="<b>#%{x} · %{customdata[0]}</b><br>%{customdata[1]:,} artículos"
                      "<br>Acumulado: %{y:.0%}<extra></extra>"))
    for meta, color in [(0.5, T["oro"]), (0.8, T["azul"])]:
        n = int((acum < meta).sum()) + 1
        linea = dict(color=color, dash="dot", width=1.5)
        fig.add_shape(type="line", x0=0, x1=n, y0=meta, y1=meta, line=linea)
        fig.add_shape(type="line", x0=n, x1=n, y0=0, y1=meta, line=linea)
        fig.add_scatter(x=[n], y=[meta], mode="markers", marker=dict(size=9, color=color,
                        line=dict(color=T["card"], width=2)), hoverinfo="skip", showlegend=False)
        fig.add_annotation(x=n, y=meta, text=f"<b>{n} productos</b> = {meta:.0%}", showarrow=False,
                           xanchor="left", yanchor="top", xshift=8, yshift=-4,
                           font=dict(color=color, size=12), bgcolor=T["card"])
    estilo(fig, alto)
    fig.update_yaxes(tickformat=".0%", range=[0, 1.04], dtick=0.25)
    fig.update_xaxes(range=[0, len(x) + 1], showgrid=False,
                     title=dict(text="Productos, de mayor a menor número de artículos",
                                font=dict(color=T["tenue"], size=11)))
    return fig


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
    with st.sidebar:
        if st.button("Actualizar datos", icon=":material/refresh:", key="actualizar", width="stretch"):
            descargar_excel.clear()
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
    todo, negs = cargar(contenido)
except Exception as e:  # noqa: BLE001
    st.error(f"No pude leer el Excel con la estructura esperada: {e}")
    st.stop()

hoy = pd.Timestamp(ahora().date())
orden = list(HOJAS.values())

# ──────────────────────────────────────────────────────────────
# SIDEBAR: navegación y filtros
# ──────────────────────────────────────────────────────────────
with st.sidebar:
    md('<div class="sb-h">Páginas</div>')
    pagina = st.radio("Página", list(PAGINAS), format_func=lambda p: f"{PAGINAS[p]} {p}",
                      label_visibility="collapsed", key="pagina")

    md('<div class="sb-h">Etapa</div>')
    etapa = st.pills("Etapa", [GLOBAL, *orden], selection_mode="single", default="Vigentes", key="f_etapa",
                     format_func=con_icono,
                     label_visibility="collapsed")
    etapa = etapa or GLOBAL   # si se desmarca la etapa, se ve todo el programa
    base = todo if etapa == GLOBAL else todo[todo["ESTATUS"] == etapa]

    md('<div class="sb-h">Filtros</div>')
    f_producto = st.multiselect("Producto", sorted(base["PRODUCTO"].dropna().unique()), placeholder="Todos")
    f_prov = st.multiselect("Proveedor", sorted(base["PROV_CORTO"].dropna().unique()), placeholder="Todos")
    f_canasta = st.multiselect("Canasta", sorted(base["CANASTA"].dropna().unique()), placeholder="Todas")

    md('<div class="sb-h">Duplicados</div>')
    dup_por = st.selectbox("Buscar repetidos por", list(CLAVES_DUP), key="dup_por")
    dup_modo = st.segmented_control("Mostrar", MODOS_DUP, default="Todos", key="dup_modo",
                                    format_func=con_icono) or "Todos"

    md('<div class="sb-foot">Subgerencia de Abarrotes<br>Alimentación para el Bienestar · Gobierno de México</div>')

# Duplicados: se detectan dentro de la etapa elegida (en Global, también entre etapas)
cols_dup = CLAVES_DUP[dup_por]
valido = base[cols_dup].notna().all(axis=1)
clave = (base.loc[valido, cols_dup].astype(str).agg(" · ".join, axis=1).reindex(base.index)
         if valido.any() else pd.Series(pd.NA, index=base.index, dtype="object"))
base = base.assign(CLAVE_DUP=clave, ES_DUP=valido & clave.duplicated(keep=False))
if dup_modo == "Solo duplicados":
    base_vista = base[base["ES_DUP"]]
elif dup_modo == "Sin duplicados":
    # De cada valor repetido se queda la fila más actual: primero por etapa (Vigentes antes que
    # Históricas o Canceladas) y luego por la vigencia más reciente. Las filas sin valor se conservan.
    prioridad = {"Vigentes": 0, "Proceso de firma": 1, "En elaboración": 2, "Lista de espera": 3,
                 "Históricas": 4, "Canceladas": 5}
    ordenada = base.assign(_p=base["ESTATUS"].map(prioridad)).sort_values(["_p", "VIG_FECHA"], ascending=[True, False])
    base_vista = ordenada[~valido.reindex(ordenada.index) | ~ordenada["CLAVE_DUP"].duplicated()].drop(columns="_p")
else:
    base_vista = base

f_previo, f = base, base_vista   # f_previo: sin aplicar el modo de duplicados (lo usa la página Duplicados)
for col, sel in [("PRODUCTO", f_producto), ("PROV_CORTO", f_prov), ("CANASTA", f_canasta)]:
    if sel:
        f = f[f[col].isin(sel)]
        f_previo = f_previo[f_previo[col].isin(sel)]

# ──────────────────────────────────────────────────────────────
# HEADER + TÍTULO
# ──────────────────────────────────────────────────────────────
md(f'<div class="topbar"><div class="crumb"><span>Alimentación para el Bienestar · </span>'
   f'<b>Programa de Abasto Rural 2026</b></div>'
   f'<div class="fecha"><i></i>Situación al {fecha_larga(hoy)}<span class="origen">· {origen}</span></div></div>')

SUBTITULOS = {
    "Panorama": "Negociaciones, proveedores y surtido de un vistazo.",
    "Proveedores": "Quién surte el programa y qué aporta cada uno.",
    "Productos": "Precios y proveedores por categoría de surtido.",
    "Cobertura": "Regiones que cubre el surtido de la selección.",
    "Vigencia": "Fechas de vigencia y próximos vencimientos.",
    "Duplicados": f"Valores repetidos por {dup_por.lower()} y en qué etapas, proveedores y negociaciones aparecen.",
}
color_etapa = T["acento"] if etapa == GLOBAL else COLOR_ESTATUS[etapa]
etiquetas = f'<span class="etiq"><i style="background:{color_etapa}"></i>{"Todo el programa" if etapa == GLOBAL else etapa}</span>'
etiquetas += "".join(f'<span class="etiq">{escape(x)}</span>' for x in [*f_producto, *f_prov, *f_canasta][:4])
if dup_modo != "Todos":
    etiquetas += (f'<span class="etiq"><i style="background:{T["oro"]}"></i>{dup_modo.lower()} · '
                  f'{escape(dup_por.lower())}</span>')
md(f'<div class="titulo"><div><h1>{pagina}</h1><p>{SUBTITULOS[pagina]}</p></div><div>{etiquetas}</div></div>')

if base.empty:
    n_neg = negs.loc[negs["ESTATUS"] == etapa, "NEGOCIACION"].nunique()
    detalle = (f"tiene {n_neg:,} negociaciones, pero todavía sin artículos capturados" if n_neg
               else "no tiene registros")
    st.info(f"La etapa **{etapa}** {detalle} en el Excel. Elige otra etapa o **Global** en la barra lateral.",
            icon=":material/hourglass_empty:")
    st.stop()
if f.empty:
    if dup_modo == "Solo duplicados" and not f_previo.empty:
        st.success(f"No hay valores repetidos por **{dup_por.lower()}** en esta selección.",
                   icon=":material/check_circle:")
    else:
        st.warning("Con esa selección no hay artículos. Prueba quitando algún filtro en la barra lateral.")
    st.stop()

# Métricas comunes a todas las páginas
dias = (f["VIG_FECHA"] - hoy).dt.days
vencidos, por_vencer = int((dias < 0).sum()), int(dias.between(0, 30).sum())
prov_n = f["PROV_CORTO"].value_counts()
prod_n = f["PRODUCTO"].value_counts()

fila_kpis([
    dict(etiqueta="Proveedores", valor=f"{f['PROVEEDOR'].nunique():,}", sub="distintos en la selección",
         badge=porcentaje(f["PROVEEDOR"].nunique(), todo["PROVEEDOR"].nunique()), icono="proveedor"),
    dict(etiqueta="Negociaciones", valor=f"{f['NEGOCIACION'].nunique():,}", sub="con artículos",
         badge=porcentaje(f["NEGOCIACION"].nunique(), todo["NEGOCIACION"].nunique()), icono="negociacion"),
    dict(etiqueta="Productos", valor=f"{f['PRODUCTO'].nunique():,}", sub="categorías de surtido",
         badge=porcentaje(f["PRODUCTO"].nunique(), todo["PRODUCTO"].nunique()), icono="producto"),
    dict(etiqueta="Artículos", valor=f"{len(f):,}", sub="referencias individuales",
         badge=porcentaje(len(f), len(todo)), icono="articulo"),
])

# ──────────────────────────────────────────────────────────────
# PANORAMA
# ──────────────────────────────────────────────────────────────
if pagina == "Panorama":
    # Embudo de negociaciones (siempre sin filtros: es el panorama general)
    est = (todo.groupby("ESTATUS").agg(Negociaciones=("NEGOCIACION", "nunique"), Articulos=("ART_ID", "count"))
           .reindex(orden).fillna(0))
    extra = negs.groupby("ESTATUS")["NEGOCIACION"].nunique().reindex(orden).fillna(0)
    est["Negociaciones"] = [int(max(a, b)) for a, b in zip(est["Negociaciones"], extra.values)]
    tope = max(est["Negociaciones"].max(), 1)
    with tarjeta("Negociaciones por etapa", "embudo",
                 "Todo el programa, de la lista de espera a las canceladas. La etapa elegida aparece resaltada.",
                 extra=f"{int(est['Negociaciones'].sum()):,} negociaciones"):
        md('<div class="etapas">' + "".join(
            f'<div class="etapa {"sel" if e == etapa else ""}"><div class="t">'
            f'<i style="background:{COLOR_ESTATUS[e]}"></i>{e}</div>'
            f'<div class="n">{int(r.Negociaciones):,}<small>{int(r.Articulos):,} art.</small></div>'
            f'<div class="pista"><div style="width:{r.Negociaciones / tope:.0%};background:{COLOR_ESTATUS[e]}"></div></div></div>'
            for e, r in est.iterrows()) + "</div>")

    st.write("")
    top5 = prov_n.head(5).sum() / len(f)
    cb = (f["CANASTA"] == "Canasta básica").mean()
    mitad = int((prod_n.cumsum() / prod_n.sum() < 0.5).sum()) + 1
    resumen = resumen_productos(f)
    uno = resumen[resumen["prov"] == 1]
    hallazgos = [
        (T["acento"], "tendencia", f"Los <b>5 principales proveedores</b> concentran el <b>{top5:.0%}</b> de los "
                                   f"artículos; el mayor es <b>{escape(prov_n.index[0])}</b>."),
        (T["verde"], "estrella", f"<b>{escape(prod_n.index[0])}</b> es el producto con más referencias "
                                 f"({prod_n.iloc[0]:,}); <b>{cb:.0%}</b> del surtido pertenece a la canasta básica."),
        (T["azul"], "pastel", f"El surtido está muy repartido: se necesitan <b>{mitad} de {len(prod_n)} productos</b> "
                              f"para juntar la mitad de los artículos."),
    ]
    if len(uno):
        hallazgos.append((T["oro"], "enlace", f"<b>{len(uno)} productos</b> dependen de <b>un solo proveedor</b>"
                                              f" ({uno['art'].sum() / len(f):.0%} de los artículos)."))
    if vencidos:
        hallazgos.append((T["alerta"], "alerta", f"<b>{vencidos / len(f):.0%}</b> de los artículos ya rebasó su fecha "
                          "de vigencia" + (f" y <b>{por_vencer:,}</b> más vencen en los próximos 30 días."
                                           if por_vencer else ".")))
    elif por_vencer:
        hallazgos.append((T["oro"], "alerta", f"<b>{por_vencer:,}</b> artículos vencen en los próximos 30 días."))
    else:
        hallazgos.append((T["ok"], "ok", "Sin vencimientos en los próximos 30 días."))

    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Lectura rápida", "lectura", "Lo más importante de la selección."):
            md('<ul class="lista">' + "".join(
                f'<li><div class="ico" style="background:{c}">{svg(i, 17)}</div><div>{t}</div></li>'
                for c, i, t in hallazgos) + "</ul>")
    with b:
        with tarjeta("Tipo de canasta", "r_canasta"):
            dona(f["CANASTA"].value_counts(), "r_canasta", 200)

    st.write("")
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Concentración del surtido", "r_conc",
                     "Cuánto del surtido juntan los productos más grandes. Pasa el cursor sobre la línea para ver "
                     "cada producto.", extra=f"{len(prod_n)} productos"):
            mostrar(concentracion(prod_n), "r_conc")
    with b:
        with tarjeta("Dependen de un solo proveedor", "r_uno",
                     "Productos que hoy surte una sola empresa: un riesgo para el abasto.",
                     extra=f"{len(uno)} de {len(resumen)}"):
            if uno.empty:
                md(f'<div class="lb-mas">Todos los productos tienen más de un proveedor.</div>')
            else:
                unico = f[f["PRODUCTO"].isin(uno.index)].groupby("PRODUCTO")["PROV_CORTO"].first()
                serie = uno["art"].sort_values(ascending=False)
                serie.index = [f"{p} · {unico[p]}" for p in serie.index]
                md(lista_barras(serie, color=T["oro"], total=len(f), alto=330))

    st.write("")
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Todos los productos", "r_todos", "Ordenados por número de artículos, con su % del surtido.",
                     extra=f"{len(prod_n)} productos"):
            buscar = st.text_input("Buscar", placeholder="Buscar producto…", label_visibility="collapsed",
                                   key="buscar_prod")
            serie = prod_n[prod_n.index.str.contains(re.escape(buscar), case=False)] if buscar else prod_n
            md(lista_barras(serie, total=prod_n.sum(), alto=460))
    with b:
        with tarjeta("Impuesto", "r_imp", "Porcentaje del surtido por régimen de impuesto."):
            dona(f["IMPUESTO"].value_counts(), "r_imp", 200,
                 colores=[T["oro"], T["verde"], T["acento"], T["azul"], T["texto3"]])
        st.write("")
        with tarjeta("Principales proveedores", "r_tabla", "Los 6 con más artículos en la selección."):
            prods_por_prov = f.groupby("PROV_CORTO")["PRODUCTO"].nunique()
            tabla(["Proveedor", "Productos", "Artículos", "Part."],
                  [[escape(p), f"{prods_por_prov[p]:,}", f"{n:,}",
                    f'<span style="color:{T["ok"]}">{n / len(f):.1%}</span>'] for p, n in prov_n.head(6).items()],
                  numericas=(1, 2, 3))

# ──────────────────────────────────────────────────────────────
# PROVEEDORES
# ──────────────────────────────────────────────────────────────
elif pagina == "Proveedores":
    n_par = min(25, len(prov_n))
    acum = prov_n.cumsum() / prov_n.sum()
    with tarjeta("Cuánto aporta cada proveedor", "p_pareto",
                 f"Los {n_par} proveedores con más artículos. La línea dorada es el % acumulado del surtido.",
                 extra=f"{len(prov_n)} proveedores"):
        top = prov_n.head(n_par)
        x = etiquetas_cortas(top.index, 18)
        fig = go.Figure()
        fig.add_bar(x=x, y=top.values, marker_color=T["acento"], name="Artículos", customdata=top.index,
                    hovertemplate="<b>%{customdata}</b><br>%{y:,} artículos<extra></extra>")
        fig.add_scatter(x=x, y=acum.head(n_par).values, yaxis="y2", mode="lines+markers",
                        line=dict(color=T["oro"], width=2.5), marker=dict(size=6), name="% acumulado",
                        customdata=top.index, hovertemplate="<b>%{customdata}</b><br>acumulado: %{y:.0%}<extra></extra>")
        estilo(fig, 460, leyenda=True)
        fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 1.04], tickformat=".0%",
                                      showgrid=False, automargin=True, tickfont=dict(color=T["tenue"])),
                          legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0, xanchor="left"),
                          margin=dict(t=40))
        fig.update_xaxes(tickangle=-45, tickfont=dict(size=10))
        mostrar(fig, "p_pareto")

    st.write("")
    n_mix = min(9, len(prov_n))
    with tarjeta("Qué surte cada proveedor", "p_mezcla",
                 f"Los {n_mix} proveedores con más artículos y sus productos principales, con nombre."):
        bloques = []
        for p in prov_n.head(n_mix).index:
            c = f[f["PROV_CORTO"] == p]["PRODUCTO"].value_counts()
            resto = c.iloc[6:]
            mas = (f'<div class="lb-mas">+ {len(resto)} productos más · {resto.sum() / c.sum():.0%} de sus artículos</div>'
                   if len(resto) else "")
            bloques.append(f'<div class="m-card mix"><div class="mix-h"><b>{escape(p)}</b>'
                           f'<span>{c.sum():,} artículos · {len(c)} productos</span></div>'
                           f'{lista_barras(c.head(6), total=c.sum(), ranking=False)}{mas}</div>')
        md('<div class="grid3">' + "".join(bloques) + "</div>")

    st.write("")
    with tarjeta("Ficha de proveedor", "p_ficha_sel", "Elige uno para ver su perfil completo."):
        prov_sel = st.selectbox("Proveedor", list(prov_n.index), label_visibility="collapsed", key="sel_prov")
    fp = f[f["PROV_CORTO"] == prov_sel]
    prox = fp.loc[fp["VIG_FECHA"] >= hoy, "VIG_FECHA"].min()
    st.write("")
    fila_kpis([
        dict(etiqueta="Artículos", valor=f"{len(fp):,}", badge=porcentaje(len(fp), len(f)), icono="articulo"),
        dict(etiqueta="Productos", valor=f"{fp['PRODUCTO'].nunique():,}", icono="producto"),
        dict(etiqueta="Negociaciones", valor=f"{fp['NEGOCIACION'].nunique():,}", icono="negociacion"),
        dict(etiqueta="Próxima vigencia", valor=prox.strftime("%d/%m/%Y") if pd.notna(prox) else "–",
             sub="fecha más cercana" if pd.notna(prox) else "sin fecha próxima", icono="calendario"),
    ])
    a, b = st.columns([3, 2], gap="medium")
    with a:
        c = fp["PRODUCTO"].value_counts()
        with tarjeta(f"Qué surte {prov_sel}", "p_ficha", "Todos sus productos.", extra=f"{len(c)} productos"):
            md(lista_barras(c, alto=420))
    with b:
        with tarjeta("Marcas que maneja", "p_marcas", "Las 10 con más artículos."):
            marcas = fp["MARCA"].value_counts().head(10)
            tabla(["Marca", "Artículos"], [[escape(m), f"{n:,}"] for m, n in marcas.items()], numericas=(1,))

# ──────────────────────────────────────────────────────────────
# PRODUCTOS
# ──────────────────────────────────────────────────────────────
elif pagina == "Productos":
    n_prec = min(20, len(prod_n))
    with tarjeta("Precio por producto", "pd_med",
                 f"Venta mediana por pieza, sin impuestos, de los {n_prec} productos con más referencias."):
        med = (f[f["PRODUCTO"].isin(prod_n.head(n_prec).index)].groupby("PRODUCTO")["VENTA_SIN"].median()
               .dropna().sort_values().reset_index())
        fig = go.Figure(go.Bar(
            x=med["VENTA_SIN"], y=etiquetas_cortas(med["PRODUCTO"], 26), orientation="h", marker_color=T["oro"],
            text=med["VENTA_SIN"].map(lambda v: f"${v:,.2f}"), textposition="outside", cliponaxis=False,
            textfont=dict(color=T["texto2"]), customdata=med["PRODUCTO"],
            hovertemplate="<b>%{customdata}</b><br>$%{x:,.2f} por pieza<extra></extra>"))
        fig.update_xaxes(visible=False, range=[0, med["VENTA_SIN"].max() * 1.3 if len(med) else 1])
        fig.update_yaxes(gridcolor="rgba(0,0,0,0)")
        mostrar(estilo(fig, alto_barras(len(med), 28)), "pd_med")

    st.write("")
    with tarjeta("Ficha de producto", "pd_sel", "Elige un producto para ver quién lo surte y a qué precio."):
        prod_sel = st.selectbox("Producto", list(prod_n.index), label_visibility="collapsed", key="sel_prod")
    fq = f[f["PRODUCTO"] == prod_sel]
    st.write("")
    fila_kpis([
        dict(etiqueta="Artículos", valor=f"{len(fq):,}", badge=porcentaje(len(fq), len(f)), icono="articulo"),
        dict(etiqueta="Proveedores", valor=f"{fq['PROVEEDOR'].nunique():,}", icono="proveedor"),
        dict(etiqueta="Marcas", valor=f"{fq['MARCA'].nunique():,}", icono="marca"),
        dict(etiqueta="Venta mediana", valor=dinero(fq["VENTA_SIN"].median()), sub="por pieza, sin imp.", icono="precio"),
    ])
    a, b = st.columns(2, gap="medium")
    with a:
        c = fq["PROV_CORTO"].value_counts()
        with tarjeta(f"Quién surte {prod_sel}", "pd_prov", "Todos los proveedores de este producto.",
                     extra=f"{len(c)} proveedores"):
            md(lista_barras(c, color=T["verde"], alto=420))
    with b:
        with tarjeta("Precio de venta por proveedor", "pd_strip", "Cada punto es un artículo."):
            top_p = fq["PROV_CORTO"].value_counts().head(10).index
            d = fq[fq["PROV_CORTO"].isin(top_p) & fq["VENTA_SIN"].notna()]
            if d.empty:
                st.caption("Sin precios para este producto.")
            else:
                cortos = dict(zip(top_p, etiquetas_cortas(top_p, 20)))
                d = d.assign(Proveedor=d["PROV_CORTO"].map(cortos))
                fig = px.strip(d, x="VENTA_SIN", y="Proveedor", hover_name="ARTICULO",
                               hover_data={"PROV_CORTO": True, "MARCA": True, "VENTA_SIN": ":$.2f", "Proveedor": False},
                               labels={"PROV_CORTO": "Proveedor", "MARCA": "Marca", "VENTA_SIN": "Venta"},
                               color_discrete_sequence=[T["acento"]])
                fig.update_traces(marker=dict(size=9, opacity=.75))
                fig.update_xaxes(tickprefix="$")
                mostrar(estilo(fig, alto_barras(d["Proveedor"].nunique(), 40, 300)), "pd_strip")

# ──────────────────────────────────────────────────────────────
# COBERTURA
# ──────────────────────────────────────────────────────────────
elif pagina == "Cobertura":
    reg = f.assign(REGION=f["COBERTURA"].str.split()).explode("REGION")
    reg = reg[reg["REGION"].notna() & ~reg["REGION"].isin(["", "NAN"])]
    if reg.empty:
        st.info("No hay información de cobertura para esta selección.")
    else:
        guia("Siglas de región tal como vienen en el Excel (NAL = nacional). Cada artículo cuenta una vez por "
             "región que cubre. En el mapa de calor, <b>mientras más intenso el cuadro, más artículos</b>.")
        st.write("")
        reg_n = reg["REGION"].value_counts()
        a, b = st.columns([2, 3], gap="medium")
        with a:
            with tarjeta("Artículos por región", "c_reg", extra=f"{len(reg_n)} regiones"):
                md(lista_barras(reg_n, color=T["verde_claro"], total=len(f), alto=500))
        with b:
            n_heat = min(15, len(prod_n))
            with tarjeta("Producto × región", "c_heat", f"Los {n_heat} productos con más referencias."):
                top_p = prod_n.head(n_heat).index
                mat = reg[reg["PRODUCTO"].isin(top_p)].groupby(["PRODUCTO", "REGION"]).size().unstack(fill_value=0)
                mat = mat.reindex([p for p in top_p if p in mat.index])
                mat = mat[[r for r in reg_n.index if r in mat.columns]]
                nombres = list(mat.index)
                mat.index = etiquetas_cortas(nombres, 24)
                fig = px.imshow(mat, aspect="auto", color_continuous_scale=T["heat"], text_auto=True)
                fig.update_xaxes(side="top", tickangle=0)
                fig.update_layout(coloraxis_showscale=False)
                fig.update_traces(xgap=2, ygap=2, textfont=dict(size=10),
                                  customdata=[[n] * mat.shape[1] for n in nombres],
                                  hovertemplate="<b>%{customdata}</b> · %{x}<br>%{z} artículos<extra></extra>")
                mostrar(estilo(fig, alto_barras(len(mat), 32, 320)), "c_heat")

# ──────────────────────────────────────────────────────────────
# VIGENCIA
# ──────────────────────────────────────────────────────────────
elif pagina == "Vigencia":
    v = f.assign(DIAS=dias)
    v["RANGO"] = pd.cut(v["DIAS"], [-float("inf"), -1, 30, 90, float("inf")],
                        labels=["Vencida", "≤ 30 días", "31–90 días", "> 90 días"]).astype(str)
    v.loc[v["VIG_FECHA"].isna(), "RANGO"] = "Sin fecha"
    cols_r = {"Vencida": T["alerta"], "≤ 30 días": T["oro"], "31–90 días": T["azul"],
              "> 90 días": T["verde"], "Sin fecha": T["tenue"]}
    a, b = st.columns([2, 3], gap="medium")
    with a:
        with tarjeta("Situación de vigencia", "v_dona", "Rojo = vencida, dorado = vence pronto, verde = con tiempo."):
            r = v["RANGO"].value_counts()
            r = r.reindex([x for x in cols_r if x in r.index])
            dona(r, "v_dona", 220, colores=[cols_r[x] for x in r.index])
    with b:
        with tarjeta("Vencimientos por mes", "v_mes", "En rojo, los meses que ya pasaron."):
            m = v.dropna(subset=["VIG_FECHA"]).copy()
            if m.empty:
                st.info("Ningún artículo de la selección tiene fecha de vigencia.")
            else:
                m["Mes"] = m["VIG_FECHA"].dt.to_period("M").dt.to_timestamp()
                m = m.groupby("Mes").size().reset_index(name="Artículos")
                m["Color"] = [T["alerta"] if x < hoy.to_period("M").to_timestamp() else T["verde"] for x in m["Mes"]]
                fig = go.Figure(go.Bar(x=m["Mes"], y=m["Artículos"], marker_color=m["Color"],
                                       hovertemplate="%{x|%B %Y}<br>%{y:,} artículos<extra></extra>"))
                fig.update_xaxes(tickformat="%b %y", gridcolor="rgba(0,0,0,0)")
                mostrar(estilo(fig, 380), "v_mes")

    st.write("")
    prox = (v[v["DIAS"].between(0, 90)].groupby(["VIG_FECHA", "PROV_CORTO"]).size()
            .reset_index(name="Artículos").sort_values("VIG_FECHA").head(12))
    with tarjeta("Próximos vencimientos", "v_prox", "Los siguientes 90 días."):
        if prox.empty:
            st.success("No hay vencimientos en los próximos 90 días para esta selección.")
        else:
            filas = []
            for _, r in prox.iterrows():
                d_ = int((r["VIG_FECHA"] - hoy).days)
                tono = "alerta" if d_ <= 30 else "neutro"
                filas.append([escape(r["PROV_CORTO"]), r["VIG_FECHA"].strftime("%d/%m/%Y"),
                              f'<span class="badge {tono}">en {d_} días</span>', f'{r["Artículos"]:,}'])
            tabla(["Proveedor", "Fecha", "Plazo", "Artículos"], filas, numericas=(3,))

# ──────────────────────────────────────────────────────────────
# DUPLICADOS
# ──────────────────────────────────────────────────────────────
elif pagina == "Duplicados":
    d0 = f_previo   # la selección sin el modo de duplicados, para poder ver todos los repetidos
    rep = d0[d0["ES_DUP"]]
    sin_valor = int(d0["CLAVE_DUP"].isna().sum())
    guia(f"Un valor está <b>duplicado</b> si su {dup_por.lower()} aparece en más de una fila. "
         + ("En <b>Global</b> se comparan todas las etapas: un artículo en <b>Históricas</b> y <b>Vigentes</b> "
            "suele ser el mismo artículo renegociado. " if etapa == GLOBAL else
            f"Solo se compara dentro de <b>{etapa}</b>; elige <b>Global</b> para comparar entre etapas. ")
         + "En la barra lateral, <b>Solo duplicados</b> muestra únicamente esas filas en todas las páginas y "
           "<b>Sin duplicados</b> deja una por valor (la más actual).")
    st.write("")

    if rep.empty:
        st.success(f"No hay valores repetidos por **{dup_por.lower()}** en esta selección.",
                   icon=":material/check_circle:")
        st.stop()

    orden_etapas = {e: i for i, e in enumerate(orden)}
    resumen_dup = (rep.groupby("CLAVE_DUP")
                   .agg(Veces=("CLAVE_DUP", "size"),
                        Etapas=("ESTATUS", lambda s: ", ".join(sorted(set(s), key=orden_etapas.get))),
                        Proveedores=("PROV_CORTO", "nunique"), Negociaciones=("NEGOCIACION", "nunique"),
                        Producto=("PRODUCTO", "first"))
                   .sort_values("Veces", ascending=False))

    fila_kpis([
        dict(etiqueta="Valores repetidos", valor=f"{len(resumen_dup):,}",
             sub=f"de {d0['CLAVE_DUP'].nunique():,} valores distintos", icono="enlace"),
        dict(etiqueta="Filas con repetido", valor=f"{len(rep):,}", sub="aparecen más de una vez",
             badge=f"{len(rep) / len(d0):.0%} de las filas", tono="alerta"),
        dict(etiqueta="Filas sobrantes", valor=f"{len(rep) - len(resumen_dup):,}",
             sub="se quitan con «Sin duplicados»", icono="articulo"),
        dict(etiqueta=f"Sin {dup_por.lower()}", valor=f"{sin_valor:,}",
             sub="filas sin dato, no se comparan", icono="alerta"),
    ])

    a, b = st.columns(2, gap="medium")
    with a:
        if etapa == GLOBAL:
            with tarjeta("En qué etapas se repiten", "d_etapas",
                         "Combinación de etapas donde aparece cada valor repetido.",
                         extra=f"{len(resumen_dup):,} valores"):
                md(lista_barras(resumen_dup["Etapas"].value_counts(), total=len(resumen_dup), alto=330))
        else:
            with tarjeta("Productos con más repetidos", "d_prod", "Filas repetidas por producto."):
                md(lista_barras(rep["PRODUCTO"].value_counts(), total=len(rep), alto=330))
    with b:
        with tarjeta("Proveedores con más repetidos", "d_prov", "Filas repetidas por proveedor.",
                     extra=f"{rep['PROV_CORTO'].nunique():,} proveedores"):
            md(lista_barras(rep["PROV_CORTO"].value_counts(), color=T["oro"], total=len(rep), alto=330))

    st.write("")
    with tarjeta("Valores más repetidos", "d_tabla", f"Ordenados por número de veces. Busca un {dup_por.lower()}, "
                 "producto o etapa.", extra=f"{len(resumen_dup):,} valores"):
        buscar = st.text_input("Buscar", placeholder="Buscar valor, producto o etapa…",
                               label_visibility="collapsed", key="buscar_dup")
        vista = resumen_dup
        if buscar:
            texto = vista.index.to_series() + " " + vista["Producto"].astype(str) + " " + vista["Etapas"]
            vista = vista[texto.str.contains(re.escape(buscar), case=False).values]
        st.caption(f"Mostrando {min(50, len(vista)):,} de {len(vista):,} valores.")
        tabla([dup_por, "Producto", "Veces", "Etapas", "Prov.", "Neg."],
              [[escape(k), escape(str(r.Producto)), f"{r.Veces:,}", escape(r.Etapas),
                f"{r.Proveedores:,}", f"{r.Negociaciones:,}"] for k, r in vista.head(50).iterrows()],
              numericas=(2, 4, 5))

    st.write("")
    with tarjeta("Detalle de un valor", "d_detalle", "Todas las filas donde aparece el valor elegido."):
        valor = st.selectbox("Valor", list(resumen_dup.index), key="sel_dup", label_visibility="collapsed",
                             format_func=lambda v: f"{v}  ·  {resumen_dup.at[v, 'Veces']} veces")
        filas_valor = rep[rep["CLAVE_DUP"] == valor].sort_values(["ESTATUS", "VIG_FECHA"])
        tabla(["Etapa", "Negociación", "Proveedor", "Artículo", "Marca", "Venta s/imp.", "Vigencia"],
              [[f'<span class="etiq" style="margin:0"><i style="background:{COLOR_ESTATUS[r.ESTATUS]}"></i>'
                f'{r.ESTATUS}</span>', escape(str(r.NEGOCIACION)), escape(str(r.PROV_CORTO)),
                escape(str(r.ARTICULO)), escape(str(r.MARCA)), dinero(r.VENTA_SIN),
                r.VIG_FECHA.strftime("%d/%m/%Y") if pd.notna(r.VIG_FECHA) else "–"]
               for r in filas_valor.itertuples()],
              numericas=(5,))
