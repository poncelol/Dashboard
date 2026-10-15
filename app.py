"""
Programa de Abasto Rural 2026 · Tablero ejecutivo
Subgerencia de Abarrotes · Alimentación para el Bienestar

Ejecutar:   streamlit run app.py
Requiere el Excel PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx junto a este archivo.
"""
import io
import re
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ──────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Abasto Rural 2026 · Alimentación para el Bienestar",
    page_icon="🌽",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ARCHIVO = Path(__file__).parent / "PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx"

# hoja del Excel -> nombre en el tablero (orden de "embudo")
HOJAS = {
    "LISTA DE ESPERA": "Lista de espera",
    "ELABORACION": "En elaboración",
    "PROCESO DE FIRMA": "Proceso de firma",
    "VIGENTES": "Vigentes",
    "HISTORICAS": "Históricas",
    "CANCELADAS": "Canceladas",
}
CANASTA_LABEL = {"CB": "Canasta básica", "COMP": "Complementario", "PAR": "PAR"}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Paleta institucional gob.mx (guinda, verde, dorado, gris)
GUINDA, GUINDA_MED = "#611232", "#9D2449"
VERDE, VERDE_MED, VERDE_CLARO = "#10312B", "#235B4E", "#3E8E7E"
ORO, ORO_CLARO = "#A57F2C", "#BC955C"
AZUL = "#2E6B9E"
BARRA_NAV = "#393C3E"
FONDO, ARENA, TEXTO, GRIS = "#F4F4F2", "#E3E1DC", "#161A1D", "#6F7271"
OTROS = "#B9B9B7"

# Colores para pasteles (oscuros, para que el % en blanco siempre se lea)
PALETA = [GUINDA, VERDE_MED, ORO, GUINDA_MED, VERDE_CLARO, AZUL, "#7A5C1E", "#B5547A", "#5F6B6D"]
COLOR_ESTATUS = {
    "Lista de espera": GRIS, "En elaboración": ORO, "Proceso de firma": AZUL,
    "Vigentes": VERDE_MED, "Históricas": "#7A5C1E", "Canceladas": GUINDA_MED,
}

# ──────────────────────────────────────────────────────────────
# ESTILO
# ──────────────────────────────────────────────────────────────
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"], .stApp {{ font-family: 'Noto Sans', 'Open Sans', sans-serif; color: {TEXTO}; }}
.stApp {{ background: {FONDO}; }}
.block-container {{ padding-top: 0 !important; padding-bottom: 3rem; max-width: 1320px; }}
#MainMenu, footer, header[data-testid="stHeader"], [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stSidebar"], [data-testid="collapsedControl"] {{ display: none !important; }}

/* Barra superior estilo gob.mx */
.gobbar {{ background: {BARRA_NAV}; color: #fff; margin: 0 -9999px 0 -9999px; padding: 0 9999px;
  display: flex; align-items: center; justify-content: space-between; height: 52px; }}
.gobbar .logo {{ font-weight: 300; font-size: 1.15rem; letter-spacing: .02em; }}
.gobbar .logo b {{ font-weight: 700; }}
.gobbar .dep {{ font-size: .85rem; opacity: .85; }}
.gobbar-wrap {{ background: {BARRA_NAV}; margin: 0 -9999px 18px -9999px; padding: 0 9999px; }}

/* Banner */
.banner {{
  background: linear-gradient(115deg, {GUINDA} 0%, {GUINDA_MED} 100%);
  border-bottom: 5px solid {ORO_CLARO};
  border-radius: 0 0 18px 18px; padding: 32px 38px; color: #fff; margin-bottom: 20px;
  display: flex; justify-content: space-between; align-items: flex-end; gap: 20px;
}}
.banner .eyebrow {{ font-size: .9rem; opacity: .85; }}
.banner h1 {{ font-weight: 700; font-size: 2.3rem; margin: 4px 0 0 0; color: #fff; line-height: 1.15; padding: 0; }}
.banner p {{ margin: 8px 0 0 0; opacity: .9; font-size: 1.02rem; }}
.banner .fecha {{ text-align: right; font-size: .85rem; opacity: .9; white-space: nowrap; }}
.banner .fecha b {{ display:block; font-size: 1.15rem; font-weight: 600; }}

/* Embudo */
.etapas {{ display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin-bottom: 8px; }}
.etapa {{ background: #fff; border-radius: 10px; padding: 14px 16px; border: 1px solid {ARENA}; }}
.etapa .n {{ font-size: 2rem; font-weight: 700; line-height: 1; }}
.etapa .t {{ font-size: .82rem; color: {GRIS}; margin-top: 6px; font-weight: 500; }}
.etapa .barra {{ height: 4px; border-radius: 4px; margin-top: 10px; }}

/* KPIs */
.kpi {{ background: #fff; border-radius: 10px; padding: 18px 22px; border: 1px solid {ARENA};
  border-top: 4px solid {GUINDA}; height: 100%; }}
.kpi .l {{ font-size: .85rem; color: {GRIS}; font-weight: 600; }}
.kpi .v {{ font-size: 2.3rem; font-weight: 700; line-height: 1.1; margin-top: 4px; color: {VERDE_MED}; }}
.kpi .s {{ font-size: .8rem; color: {GRIS}; margin-top: 2px; }}

/* Lectura rápida */
.hallazgo {{ background: #fff; border-radius: 10px; border: 1px solid {ARENA}; border-left: 5px solid {ORO};
  padding: 12px 16px; font-size: .95rem; line-height: 1.5; height: 100%; }}
.hallazgo b {{ color: {GUINDA}; }}
.hallazgo.alerta {{ border-left-color: {GUINDA_MED}; background: #FBF3F5; }}

/* Títulos con la línea guinda característica de gob.mx */
h2, h3, h4 {{ font-family: 'Noto Sans', sans-serif !important; color: {TEXTO}; font-weight: 600 !important; }}
.seccion {{ font-size: 1.35rem; font-weight: 600; color: {TEXTO}; margin: 30px 0 0 0; }}
.seccion::after {{ content: ""; display: block; width: 56px; height: 3px; background: {GUINDA}; margin: 6px 0 8px 0; }}
.sub {{ color: {GRIS}; font-size: .9rem; margin-bottom: 10px; }}
.guia {{ background: #fff; border: 1px dashed {ORO_CLARO}; border-radius: 10px; padding: 10px 14px;
  font-size: .88rem; color: {TEXTO}; margin: 4px 0 12px 0; }}

.stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {ARENA}; }}
.stTabs [data-baseweb="tab"] {{ background: transparent; padding: 10px 20px; font-weight: 500; color: {GRIS}; border-radius: 8px 8px 0 0; }}
.stTabs [aria-selected="true"] {{ color: {GUINDA} !important; background: #fff; font-weight: 600; }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {GUINDA}; height: 3px; }}

/* Tarjetas para gráficos */
[data-testid="stPlotlyChart"] {{ background: #fff; border-radius: 10px; padding: 8px 8px 0 8px; border: 1px solid {ARENA}; }}
[data-testid="stExpander"] {{ background:#fff; border-radius:10px; border: 1px solid {ARENA}; }}

/* Leyenda compartida y chips */
.chip {{ display:inline-flex; align-items:center; gap:7px; background:#fff; border:1px solid {ARENA}; color:{TEXTO};
  border-radius:20px; padding:4px 12px; margin:3px 6px 3px 0; font-size:.84rem; font-weight:500; }}
.chip i {{ display:inline-block; width:12px; height:12px; border-radius:50%; }}

@media (max-width: 900px) {{
  .etapas {{ grid-template-columns: repeat(2, 1fr); }}
  .banner {{ flex-direction: column; align-items: flex-start; }}
  .banner h1 {{ font-size: 1.7rem; }}
}}
</style>
""",
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────
# CARGA Y LIMPIEZA
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

    neg["NEGOCIACION"] = neg["NEGOCIACION"].astype("Int64").astype(str)
    return df, neg


# ──────────────────────────────────────────────────────────────
# AYUDAS DE PRESENTACIÓN
# ──────────────────────────────────────────────────────────────
def fecha_larga(d) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def dinero(x) -> str:
    return "–" if pd.isna(x) else f"${x:,.2f}"


def kpi(etiqueta, valor, sub=""):
    return (f'<div class="kpi"><div class="l">{etiqueta}</div>'
            f'<div class="v">{valor}</div><div class="s">{sub}</div></div>')


def seccion(titulo, sub=""):
    st.markdown(f'<div class="seccion">{titulo}</div><div class="sub">{sub}</div>', unsafe_allow_html=True)


def guia(texto):
    """Pequeña nota de 'cómo leer' para que la gráfica se entienda sin explicación."""
    st.markdown(f'<div class="guia">💡 {texto}</div>', unsafe_allow_html=True)


def estilo(fig, alto=380, leyenda=False, titulo=None):
    fig.update_layout(
        height=alto,
        margin=dict(l=16, r=16, t=56 if titulo else 16, b=16),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Noto Sans, sans-serif", color=TEXTO, size=13),
        title=dict(text=titulo, font=dict(family="Noto Sans, sans-serif", size=16, color=TEXTO), x=0.02) if titulo else None,
        showlegend=leyenda, colorway=PALETA,
        legend=dict(orientation="h", y=-0.08, title=None),
        hoverlabel=dict(bgcolor="white", font_size=13, font_family="Noto Sans"),
    )
    fig.update_xaxes(gridcolor=ARENA, zeroline=False, title=None, showline=False)
    fig.update_yaxes(gridcolor=ARENA, zeroline=False, title=None, showline=False)
    return fig


def mostrar(fig, clave):
    st.plotly_chart(fig, width="stretch", theme=None, key=clave, config={"displayModeBar": False})


def barras_h(serie, titulo, color=VERDE_MED, alto=400, n=10):
    d = serie.head(n).rename_axis("k").reset_index(name="v").iloc[::-1]
    fig = go.Figure(go.Bar(
        x=d["v"], y=d["k"], orientation="h", marker_color=color,
        text=d["v"].map("{:,}".format), textposition="outside", cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x:,} artículos<extra></extra>",
    ))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(gridcolor="rgba(0,0,0,0)")
    return estilo(fig, alto, titulo=titulo)


def _centro(fig, total):
    fig.add_annotation(
        text=f"<b>{int(total):,}</b><br><span style='font-size:11px;color:{GRIS}'>artículos</span>",
        showarrow=False, font=dict(family="Noto Sans, sans-serif", size=22, color=GUINDA))


def dona(serie, titulo, alto=300, colores=None, max_items=6):
    """Dona genérica con leyenda (para cuando hay una sola gráfica a la vista)."""
    s = serie.sort_values(ascending=False)
    if len(s) > max_items:
        s = pd.concat([s.head(max_items - 1), pd.Series({"Otros": s.iloc[max_items - 1:].sum()})])
    cols = list(colores or PALETA)
    cols = [cols[i % len(cols)] for i in range(len(s))]
    if "Otros" in s.index:
        cols[list(s.index).index("Otros")] = OTROS
    fig = go.Figure(go.Pie(
        labels=s.index, values=s.values, hole=0.6, sort=False,
        marker=dict(colors=cols, line=dict(color="white", width=2)),
        texttemplate="%{percent:.0%}", textposition="inside", insidetextorientation="horizontal",
        textfont=dict(size=13, color="white"),
        hovertemplate="<b>%{label}</b><br>%{value:,} artículos · %{percent}<extra></extra>",
    ))
    fig.update_layout(uniformtext=dict(minsize=10, mode="hide"))
    _centro(fig, s.sum())
    return estilo(fig, alto, leyenda=True, titulo=titulo)


# ── Pasteles de "qué surte cada proveedor" ────────────────────
def mapa_colores(prod_conteo, n=8):
    """Un color fijo por producto principal: el mismo color significa lo mismo en todos los pasteles."""
    return {p: PALETA[i] for i, p in enumerate(prod_conteo.head(n).index)}


def agrupar_productos(conteo, mapa):
    s = {}
    for p, v in conteo.items():
        k = p if p in mapa else "Otros"
        s[k] = s.get(k, 0) + int(v)
    s = pd.Series(s)
    orden_ = [k for k in mapa if k in s.index] + (["Otros"] if "Otros" in s.index else [])
    return s.reindex(orden_)


def pastel_proveedor(conteo, mapa, titulo, alto=310):
    s = agrupar_productos(conteo, mapa)
    fig = go.Figure(go.Pie(
        labels=s.index, values=s.values, hole=0.55, sort=False,
        marker=dict(colors=[mapa.get(k, OTROS) for k in s.index], line=dict(color="white", width=2)),
        texttemplate="%{percent:.0%}", textposition="inside", insidetextorientation="horizontal",
        textfont=dict(size=12, color="white"),
        hovertemplate="<b>%{label}</b><br>%{value:,} artículos · %{percent}<extra></extra>",
    ))
    fig.update_layout(uniformtext=dict(minsize=10, mode="hide"))
    _centro(fig, s.sum())
    t = titulo if len(titulo) <= 30 else titulo[:28] + "…"
    return estilo(fig, alto, leyenda=False, titulo=f"<b>{t}</b>")


def leyenda_colores(mapa, incluir_otros=True):
    chips = "".join(f'<span class="chip"><i style="background:{c}"></i>{p}</span>' for p, c in mapa.items())
    if incluir_otros:
        chips += f'<span class="chip"><i style="background:{OTROS}"></i>Otros productos</span>'
    st.markdown(chips, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# DATOS
# ──────────────────────────────────────────────────────────────
if not ARCHIVO.exists():
    st.error(f"No encuentro el archivo de datos **{ARCHIVO.name}** junto a app.py.")
    st.stop()

try:
    todo, negs = cargar(ARCHIVO.read_bytes())
except Exception as e:  # noqa: BLE001
    st.error(f"No pude leer el Excel con la estructura esperada: {e}")
    st.stop()

hoy = pd.Timestamp(date.today())

# ──────────────────────────────────────────────────────────────
# ENCABEZADO
# ──────────────────────────────────────────────────────────────
st.markdown(
    f"""<div class="gobbar-wrap"><div class="gobbar" style="margin:0;padding:0;">
        <span class="logo">GOBIERNO DE <b>MÉXICO</b></span>
        <span class="dep">Alimentación para el Bienestar</span>
    </div></div>
    <div class="banner">
      <div>
        <div class="eyebrow">Subgerencia de Abarrotes</div>
        <h1>Programa de Abasto Rural 2026</h1>
        <p>Negociaciones, proveedores y surtido de un vistazo</p>
      </div>
      <div class="fecha">Situación al<b>{fecha_larga(hoy)}</b></div>
    </div>""",
    unsafe_allow_html=True,
)

# Embudo de negociaciones (siempre sin filtros: es el panorama general)
orden = list(HOJAS.values())
est = (todo.groupby("ESTATUS").agg(Negociaciones=("NEGOCIACION", "nunique"), Artículos=("ART_ID", "count"))
       .reindex(orden).fillna(0))
extra = negs.groupby("ESTATUS")["NEGOCIACION"].nunique().reindex(orden).fillna(0)
est["Negociaciones"] = [int(max(a, b)) for a, b in zip(est["Negociaciones"], extra.values)]
tarjetas = "".join(
    f'<div class="etapa"><div class="n" style="color:{COLOR_ESTATUS[e]}">{int(r.Negociaciones):,}</div>'
    f'<div class="t">{e}</div><div class="barra" style="background:{COLOR_ESTATUS[e]}"></div></div>'
    for e, r in est.iterrows()
)
st.markdown(f'<div class="etapas">{tarjetas}</div>'
            '<div class="sub">Negociaciones en cada etapa del programa, de la lista de espera a las canceladas.</div>',
            unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# FILTROS
# ──────────────────────────────────────────────────────────────
st.markdown('<div class="sub" style="margin-top:6px">Elige qué etapas quieres ver (puedes marcar varias):</div>',
            unsafe_allow_html=True)
f_estatus = st.pills("Ver", orden, selection_mode="multi", default=["Vigentes"], key="f_estatus",
                     label_visibility="collapsed")
base = todo[todo["ESTATUS"].isin(f_estatus)] if f_estatus else todo

with st.expander("🔎 Filtrar por producto, proveedor o canasta"):
    c1, c2, c3 = st.columns(3)
    f_producto = c1.multiselect("Producto", sorted(base["PRODUCTO"].dropna().unique()), placeholder="Todos")
    f_prov = c2.multiselect("Proveedor", sorted(base["PROV_CORTO"].dropna().unique()), placeholder="Todos")
    f_canasta = c3.multiselect("Canasta", sorted(base["CANASTA"].dropna().unique()), placeholder="Todas")

f = base
for col, sel in [("PRODUCTO", f_producto), ("PROV_CORTO", f_prov), ("CANASTA", f_canasta)]:
    if sel:
        f = f[f[col].isin(sel)]

if f.empty:
    st.warning("Con esa selección no hay artículos. Prueba quitando algún filtro.")
    st.stop()

# ──────────────────────────────────────────────────────────────
# KPIs + LECTURA RÁPIDA
# ──────────────────────────────────────────────────────────────
st.write("")
k = st.columns(4)
k[0].markdown(kpi("Proveedores", f"{f['PROVEEDOR'].nunique():,}", "distintos en la selección"), unsafe_allow_html=True)
k[1].markdown(kpi("Negociaciones", f"{f['NEGOCIACION'].nunique():,}", "con artículos"), unsafe_allow_html=True)
k[2].markdown(kpi("Productos", f"{f['PRODUCTO'].nunique():,}", "categorías de surtido"), unsafe_allow_html=True)
k[3].markdown(kpi("Artículos", f"{len(f):,}", "referencias individuales"), unsafe_allow_html=True)

dias = (f["VIG_FECHA"] - hoy).dt.days
vencidos, por_vencer = int((dias < 0).sum()), int(dias.between(0, 30).sum())
prov_n = f["PROV_CORTO"].value_counts()
prod_n = f["PRODUCTO"].value_counts()
top5 = prov_n.head(5).sum() / len(f)
cb = (f["CANASTA"] == "Canasta básica").mean()

hallazgos = [
    ("", f"🏭 Los <b>5 principales proveedores</b> concentran el <b>{top5:.0%}</b> de los artículos; "
         f"el mayor es <b>{prov_n.index[0]}</b>."),
    ("", f"🛒 <b>{prod_n.index[0]}</b> es el producto con más referencias ({prod_n.iloc[0]:,}); "
         f"<b>{cb:.0%}</b> del surtido pertenece a la canasta básica."),
]
if vencidos:
    hallazgos.append(("alerta", f"⚠️ <b>{vencidos / len(f):.0%}</b> de los artículos ya rebasó su fecha de vigencia"
                                + (f" y <b>{por_vencer:,}</b> más vencen en los próximos 30 días." if por_vencer else ".")))
elif por_vencer:
    hallazgos.append(("alerta", f"⚠️ <b>{por_vencer:,}</b> artículos vencen en los próximos 30 días."))
else:
    hallazgos.append(("", "✅ Sin vencimientos en los próximos 30 días."))

seccion("Lectura rápida", "Lo más importante de la selección, en una frase.")
h = st.columns(len(hallazgos))
for col, (cls, txt) in zip(h, hallazgos):
    col.markdown(f'<div class="hallazgo {cls}">{txt}</div>', unsafe_allow_html=True)

st.write("")
tabs = st.tabs(["📊 Panorama", "🏭 Proveedores", "🛒 Productos", "🗺️ Cobertura", "📅 Vigencia"])

# ──────────────────────────────────────────────────────────────
# PANORAMA
# ──────────────────────────────────────────────────────────────
with tabs[0]:
    guia("Las barras muestran <b>cuántos artículos</b> tiene cada producto. Los círculos muestran qué <b>porcentaje</b> "
         "del surtido corresponde a cada tipo de canasta e impuesto.")
    a, b = st.columns([3, 2])
    with a:
        mostrar(barras_h(prod_n, "Productos con más referencias", VERDE_MED, 470, n=12), "r_prod")
    with b:
        mostrar(dona(f["CANASTA"].value_counts(), "Tipo de canasta", 228), "r_canasta")
        mostrar(dona(f["IMPUESTO"].value_counts(), "Impuesto", 228,
                     colores=[ORO, VERDE_MED, GUINDA_MED, AZUL, GRIS, VERDE_CLARO]), "r_imp")

# ──────────────────────────────────────────────────────────────
# PROVEEDORES
# ──────────────────────────────────────────────────────────────
with tabs[1]:
    guia("Cada <b>círculo</b> es un proveedor y cada color es un producto. Entre más grande el pedazo, "
         "más artículos de ese producto surte. El mismo color significa lo mismo en todos los círculos.")

    a, b = st.columns([3, 2])
    with a:
        mostrar(barras_h(prov_n, "Proveedores con más artículos", GUINDA, 440, n=10), "p_top")
    with b:
        mostrar(dona(prov_n, "Cuánto aporta cada proveedor", 440, max_items=7), "p_share")

    # ── Qué surte cada uno: pasteles lado a lado ──
    seccion("Qué surte cada proveedor", "Los 6 proveedores con más artículos y la mezcla de productos de cada uno.")
    mapa = mapa_colores(prod_n, 8)
    leyenda_colores(mapa, incluir_otros=len(prod_n) > len(mapa))
    st.write("")
    top6 = list(prov_n.head(6).index)
    for fila in range(0, len(top6), 3):
        cols = st.columns(3)
        for col, (i, p) in zip(cols, list(enumerate(top6))[fila:fila + 3]):
            conteo = f[f["PROV_CORTO"] == p]["PRODUCTO"].value_counts()
            with col:
                mostrar(pastel_proveedor(conteo, mapa, p), f"p_pie_{i}")

    # ── Ficha individual ──
    seccion("Ficha de proveedor", "Elige uno para ver su perfil completo.")
    prov_sel = st.selectbox("Proveedor", list(prov_n.index), label_visibility="collapsed", key="sel_prov")
    fp = f[f["PROV_CORTO"] == prov_sel]
    k = st.columns(4)
    k[0].markdown(kpi("Artículos", f"{len(fp):,}"), unsafe_allow_html=True)
    k[1].markdown(kpi("Productos", f"{fp['PRODUCTO'].nunique():,}"), unsafe_allow_html=True)
    k[2].markdown(kpi("Negociaciones", f"{fp['NEGOCIACION'].nunique():,}"), unsafe_allow_html=True)
    prox = fp.loc[fp["VIG_FECHA"] >= hoy, "VIG_FECHA"].min()
    k[3].markdown(kpi("Próxima vigencia", prox.strftime("%d/%m/%Y") if pd.notna(prox) else "–",
                      "fecha más cercana" if pd.notna(prox) else "sin fecha próxima"), unsafe_allow_html=True)
    st.write("")
    a, b = st.columns([3, 2])
    with a:
        mostrar(dona(fp["PRODUCTO"].value_counts(), f"Qué surte {prov_sel}", 400, max_items=8), "p_ficha")
    with b:
        marcas = fp["MARCA"].value_counts().head(10)
        st.markdown("#### Marcas que maneja")
        st.markdown("".join(f'<span class="chip">{m} · {n}</span>' for m, n in marcas.items()),
                    unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# PRODUCTOS
# ──────────────────────────────────────────────────────────────
with tabs[2]:
    seccion("Precio por producto", "Venta mediana por pieza, sin impuestos, de los productos con más referencias.")
    med = (f[f["PRODUCTO"].isin(prod_n.head(12).index)].groupby("PRODUCTO")["VENTA_SIN"].median()
           .dropna().sort_values().reset_index())
    fig = go.Figure(go.Bar(
        x=med["VENTA_SIN"], y=med["PRODUCTO"], orientation="h", marker_color=ORO,
        text=med["VENTA_SIN"].map(lambda v: f"${v:,.2f}"), textposition="outside", cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>$%{x:,.2f} por pieza<extra></extra>"))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(gridcolor="rgba(0,0,0,0)")
    mostrar(estilo(fig, 440), "pd_med")

    seccion("Ficha de producto", "Elige un producto para ver quién lo surte y a qué precio.")
    prod_sel = st.selectbox("Producto", list(prod_n.index), label_visibility="collapsed", key="sel_prod")
    fq = f[f["PRODUCTO"] == prod_sel]
    k = st.columns(4)
    k[0].markdown(kpi("Artículos", f"{len(fq):,}"), unsafe_allow_html=True)
    k[1].markdown(kpi("Proveedores", f"{fq['PROVEEDOR'].nunique():,}"), unsafe_allow_html=True)
    k[2].markdown(kpi("Marcas", f"{fq['MARCA'].nunique():,}"), unsafe_allow_html=True)
    k[3].markdown(kpi("Venta mediana", dinero(fq["VENTA_SIN"].median()), "por pieza, sin imp."), unsafe_allow_html=True)
    st.write("")
    a, b = st.columns(2)
    with a:
        mostrar(dona(fq["PROV_CORTO"].value_counts(), f"Quién surte {prod_sel}", 420, max_items=7), "pd_prov")
    with b:
        top_p = fq["PROV_CORTO"].value_counts().head(8).index
        d = fq[fq["PROV_CORTO"].isin(top_p) & fq["VENTA_SIN"].notna()]
        if d.empty:
            st.caption("Sin precios para este producto.")
        else:
            fig = px.strip(d, x="VENTA_SIN", y="PROV_CORTO", hover_name="ARTICULO",
                           hover_data={"MARCA": True, "VENTA_SIN": ":$.2f", "PROV_CORTO": False},
                           color_discrete_sequence=[VERDE_MED])
            fig.update_traces(marker=dict(size=9, opacity=.75))
            fig.update_xaxes(tickprefix="$")
            mostrar(estilo(fig, 420, titulo="Precio de venta por proveedor (cada punto es un artículo)"), "pd_strip")

# ──────────────────────────────────────────────────────────────
# COBERTURA
# ──────────────────────────────────────────────────────────────
with tabs[3]:
    reg = f.assign(REGION=f["COBERTURA"].str.split()).explode("REGION")
    reg = reg[reg["REGION"].notna() & ~reg["REGION"].isin(["", "NAN"])]
    if reg.empty:
        st.info("No hay información de cobertura para esta selección.")
    else:
        guia("Siglas de región tal como vienen en el Excel (NAL = nacional). Cada artículo cuenta una vez por "
             "región que cubre. En el mapa de calor, <b>mientras más oscuro el cuadro, más artículos</b>.")
        a, b = st.columns([2, 3])
        with a:
            mostrar(barras_h(reg["REGION"].value_counts(), "Artículos por región", VERDE_CLARO, 460, n=12), "c_reg")
        with b:
            top_p = f["PRODUCTO"].value_counts().head(10).index
            mat = reg[reg["PRODUCTO"].isin(top_p)].groupby(["PRODUCTO", "REGION"]).size().unstack(fill_value=0)
            fig = px.imshow(mat, aspect="auto", color_continuous_scale=["#F4EFE6", ORO_CLARO, GUINDA])
            fig.update_xaxes(side="top")
            fig.update_layout(coloraxis_showscale=False)
            fig.update_traces(hovertemplate="<b>%{y}</b> · %{x}<br>%{z} artículos<extra></extra>")
            mostrar(estilo(fig, 460, titulo="Producto × región"), "c_heat")

# ──────────────────────────────────────────────────────────────
# VIGENCIA
# ──────────────────────────────────────────────────────────────
with tabs[4]:
    v = f.assign(DIAS=dias)

    def rango(r):
        if pd.isna(r.VIG_FECHA):
            return "Sin fecha"
        if r.DIAS < 0:
            return "Vencida"
        if r.DIAS <= 30:
            return "≤ 30 días"
        if r.DIAS <= 90:
            return "31–90 días"
        return "> 90 días"

    v["RANGO"] = v.apply(rango, axis=1)
    cols_r = {"Vencida": GUINDA_MED, "≤ 30 días": ORO, "31–90 días": AZUL, "> 90 días": VERDE_MED, "Sin fecha": GRIS}
    guia("Rojo = ya venció, dorado = vence pronto, verde = vigente con tiempo.")
    a, b = st.columns([2, 3])
    with a:
        r = v["RANGO"].value_counts()
        r = r.reindex([x for x in cols_r if x in r.index])
        mostrar(dona(r, "Situación de vigencia", 380, colores=[cols_r[x] for x in r.index]), "v_dona")
    with b:
        m = v.dropna(subset=["VIG_FECHA"]).copy()
        if m.empty:
            st.info("Ningún artículo de la selección tiene fecha de vigencia.")
        else:
            m["Mes"] = m["VIG_FECHA"].dt.to_period("M").dt.to_timestamp()
            m = m.groupby("Mes").size().reset_index(name="Artículos")
            m["Color"] = [GUINDA_MED if x < hoy.to_period("M").to_timestamp() else VERDE_MED for x in m["Mes"]]
            fig = go.Figure(go.Bar(x=m["Mes"], y=m["Artículos"], marker_color=m["Color"],
                                   hovertemplate="%{x|%B %Y}<br>%{y:,} artículos<extra></extra>"))
            fig.update_xaxes(tickformat="%b %y")
            mostrar(estilo(fig, 380, titulo="Vencimientos por mes (en rojo, ya pasados)"), "v_mes")

    prox = (v[v["DIAS"].between(0, 90)].groupby(["VIG_FECHA", "PROV_CORTO"]).size()
            .reset_index(name="Artículos").sort_values("VIG_FECHA").head(8))
    seccion("Próximos vencimientos", "Los siguientes 90 días.")
    if prox.empty:
        st.success("No hay vencimientos en los próximos 90 días para esta selección.")
    else:
        for _, r in prox.iterrows():
            d_ = int((r["VIG_FECHA"] - hoy).days)
            color = GUINDA_MED if d_ <= 30 else AZUL
            st.markdown(
                f'<div class="hallazgo" style="border-left-color:{color};margin-bottom:8px">'
                f'<b>{r["VIG_FECHA"].strftime("%d/%m/%Y")}</b>'
                f' · en {d_} días · {r["PROV_CORTO"]} · {r["Artículos"]:,} artículos</div>',
                unsafe_allow_html=True)