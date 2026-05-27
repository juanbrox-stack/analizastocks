import streamlit as st
import pandas as pd
import numpy as np
import requests
import io
from datetime import datetime

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Stock Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #0f1117; color: #e0e0e0; }
    .main-header {
        background: linear-gradient(135deg, #1a1f35 0%, #0d1b2a 100%);
        border: 1px solid #2a3550;
        border-radius: 12px;
        padding: 20px 30px;
        margin-bottom: 24px;
    }
    .metric-card {
        background: #1a1f35;
        border: 1px solid #2a3550;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
    }
    .metric-card .label { font-size: 12px; color: #8899aa; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
    .metric-card .value { font-size: 28px; font-weight: 700; }
    .badge-danger { background: #3d1515; color: #ff5555; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .badge-warning { background: #3d2d0a; color: #ffaa33; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .badge-success { background: #0d2d1a; color: #33cc66; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .badge-info { background: #0d1f3d; color: #3399ff; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .section-title { font-size: 16px; font-weight: 600; color: #c0d0e8; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 1px solid #2a3550; }
    div[data-testid="stDataFrame"] { border-radius: 8px; border: 1px solid #2a3550; }
    .stTabs [data-baseweb="tab"] { color: #8899aa; }
    .stTabs [aria-selected="true"] { color: #3399ff; border-bottom-color: #3399ff !important; }
    .stSelectbox > div > div { background-color: #1a1f35 !important; border-color: #2a3550 !important; }
    .url-box { background: #151b2e; border: 1px dashed #3a4a6a; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ──────────────────────────────────────────────────────────────────
TARIFA_DEFAULT_SHEET = "T_AMZ"

@st.cache_data(show_spinner=False)
def load_tarifa_from_url(url: str):
    """Download tarifa Excel from a URL."""
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        xls = pd.ExcelFile(io.BytesIO(r.content))
        # Use first T_ sheet available
        sheet = next((s for s in xls.sheet_names if s.startswith("T_")), xls.sheet_names[0])
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = df.columns.str.strip()
        ref_col = "sku" if "sku" in df.columns else "REFERENCIA"
        df = df.rename(columns={ref_col: "REFERENCIA"})
        df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip()
        return df, None
    except Exception as e:
        return None, str(e)


@st.cache_data(show_spinner=False)
def load_stock_from_url(url: str):
    """Download stock Excel from a URL."""
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        xls = pd.ExcelFile(io.BytesIO(r.content))
        sheets = {}
        for sh in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sh)
            df.columns = df.columns.str.strip()
            if "Referencia" in df.columns:
                df["Referencia"] = df["Referencia"].astype(str).str.strip()
            sheets[sh] = df
        return sheets, None
    except Exception as e:
        return None, str(e)


def load_uploaded(file_obj):
    """Load Excel from uploaded file object."""
    xls = pd.ExcelFile(file_obj)
    return xls


def get_tarifa_df(xls):
    # Use first T_ sheet available (T_AMZ default)
    sheet = next((s for s in xls.sheet_names if s.startswith("T_")), xls.sheet_names[0])
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = df.columns.str.strip()
    ref_col = "sku" if "sku" in df.columns else "REFERENCIA"
    df = df.rename(columns={ref_col: "REFERENCIA"})
    df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip()
    return df


# Mapping: country -> sheet name and stock column names
COUNTRY_CONFIG = {
    "España":    {"sheet": "España",   "disponible": "Stock Disponible", "real": "Stock Fisico",  "extra": ["Stock Operativo", "Mar", "Puerto", "Despachado"], "has_transit": True},
    "Alemania":  {"sheet": "Alemania", "disponible": "StockDisponible", "real": "StockReal",     "extra": [], "has_transit": False},
    "Francia":   {"sheet": "Francia",  "disponible": "StockDisponible", "real": "StockReal",     "extra": [], "has_transit": False},
    "Italia":    {"sheet": "Italia",   "disponible": "StockDisponible", "real": "StockReal",     "extra": [], "has_transit": False},
}

def get_stock_for_country(xls, country: str):
    cfg = COUNTRY_CONFIG[country]
    df = pd.read_excel(xls, sheet_name=cfg["sheet"])
    df.columns = df.columns.str.strip()
    df["Referencia"] = df["Referencia"].astype(str).str.strip()
    # Normalise to unified column names
    df = df.rename(columns={
        cfg["disponible"]: "Stock Disponible",
        cfg["real"]: "Stock Fisico",
    })
    # Ensure transit columns exist
    for col in ["Mar", "Puerto", "Despachado", "Stock Operativo"]:
        if col not in df.columns:
            df[col] = 0
    df["Stock Disponible"] = pd.to_numeric(df["Stock Disponible"], errors="coerce").fillna(0)
    df["Stock Fisico"] = pd.to_numeric(df["Stock Fisico"], errors="coerce").fillna(0)
    return df


def merge_tarifa_stock(df_tarifa, df_stock):
    # Pick only columns that exist in this country's stock sheet
    stock_cols = ["Referencia", "Stock Disponible", "Stock Fisico",
                  "Mar", "Puerto", "Despachado", "Stock Operativo",
                  "Familia Padre", "Familia", "Subfamilia"]
    stock_cols = [c for c in stock_cols if c in df_stock.columns]
    df = df_tarifa.merge(df_stock[stock_cols], left_on="REFERENCIA", right_on="Referencia", how="left")
    for col in ["Stock Disponible", "Stock Fisico", "Stock Operativo", "Mar", "Puerto", "Despachado"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Sin stock"] = df["Stock Disponible"] == 0
    df["En mar"] = df["Mar"] > 0
    df["En puerto"] = df["Puerto"] > 0
    df["Total entrante"] = df["Mar"] + df["Puerto"] + df["Despachado"]
    df["Riesgo rotura"] = (df["Sin stock"]) & (df["Total entrante"] == 0)
    df["Stock bajo"] = (df["Stock Disponible"] > 0) & (df["Stock Disponible"] <= 3)
    return df


def status_label(row):
    if row["Sin stock"] and row["Riesgo rotura"]:
        return "🔴 Sin stock"
    elif row["Sin stock"] and row["Total entrante"] > 0:
        return "🟡 Sin stock (entrante)"
    elif row["Stock bajo"]:
        return "🟠 Stock bajo"
    else:
        return "🟢 OK"


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")

    st.markdown("---")
    st.markdown("#### 📁 Fuente de datos")
    data_mode = st.radio(
        "Modo",
        ["Subir archivos", "URLs remotas"],
        label_visibility="collapsed",
    )

    tarifa_xls = None
    stock_xls = None
    tarifa_dfs_remote = None
    stock_sheets_remote = None
    data_ready = False

    if data_mode == "Subir archivos":
        f_tarifa = st.file_uploader("📊 Tarifa Nacional (.xlsx)", type=["xlsx"])
        f_stock = st.file_uploader("📦 Stock Global (.xlsx)", type=["xlsx"])
        if f_tarifa and f_stock:
            tarifa_xls = load_uploaded(f_tarifa)
            stock_xls = load_uploaded(f_stock)
            data_ready = True
    else:
        st.markdown('<div class="url-box">', unsafe_allow_html=True)
        tarifa_url = st.text_input(
            "🔗 URL Tarifa Nacional",
            placeholder="https://example.com/tarifa.xlsx",
            help="URL pública de descarga del Excel de tarifa",
        )
        stock_url = st.text_input(
            "🔗 URL Stock Global",
            placeholder="https://example.com/stock.xlsx",
            help="URL pública de descarga del Excel de stock",
        )
        st.markdown('</div>', unsafe_allow_html=True)
        if tarifa_url and stock_url:
            with st.spinner("Descargando datos..."):
                tarifa_dfs_remote, err1 = load_tarifa_from_url(tarifa_url)
                stock_sheets_remote, err2 = load_stock_from_url(stock_url)
            if err1:
                st.error(f"Error tarifa: {err1}")
            elif err2:
                st.error(f"Error stock: {err2}")
            else:
                data_ready = True
                st.success("✅ Datos cargados")

    st.markdown("---")
    st.markdown("#### 🌍 País")
    selected_country = st.selectbox(
        "País",
        ["España", "Alemania", "Francia", "Italia"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown("#### 🔍 Filtros")

    # Filters are rendered below after data loads (options depend on data)
    filter_placeholder = st.container()

    show_only = st.multiselect(
        "Mostrar sólo",
        ["Sin stock", "Stock bajo", "En mar", "En puerto", "Riesgo rotura"],
        default=[],
    )

    depos_filter = st.checkbox("Ocultar desposicionados", value=False)

    st.markdown("---")
    st.caption(f"🕒 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

# ── Main ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="main-header">', unsafe_allow_html=True)
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("## 📦 Dashboard de Stock vs Tarifa")
    st.markdown("Visibilidad en tiempo real · Anticipación a roturas de stock")
with col_h2:
    pass
st.markdown("</div>", unsafe_allow_html=True)

if not data_ready:
    st.info("👈 Carga los datos desde la barra lateral para comenzar (archivos o URLs).")
    st.markdown("""
    **¿Qué muestra este dashboard?**
    - 🔴 SKUs **sin stock** y sin reposición prevista  
    - 🟡 SKUs **sin stock** pero con mercancía **en mar o puerto**  
    - 🟠 SKUs con **stock bajo** (≤ 3 unidades)  
    - 📊 Análisis por **familia y subfamilia**
    - 🚢 Stock **en tránsito** (mar, puerto, despachado)
    """)
    st.stop()

# ── Load & merge data ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def build_merged(country, _tarifa_xls=None, _stock_xls=None,
                 _tarifa_df_remote=None, _stock_sheets_remote=None):
    if _tarifa_xls is not None:
        df_t = get_tarifa_df(_tarifa_xls)
        df_s = get_stock_for_country(_stock_xls, country)
    else:
        df_t = _tarifa_df_remote.copy()
        cfg = COUNTRY_CONFIG[country]
        df_s = _stock_sheets_remote[cfg["sheet"]].copy()
        df_s.columns = df_s.columns.str.strip()
        df_s["Referencia"] = df_s["Referencia"].astype(str).str.strip()
        df_s = df_s.rename(columns={cfg["disponible"]: "Stock Disponible", cfg["real"]: "Stock Fisico"})
        for col in ["Mar", "Puerto", "Despachado", "Stock Operativo"]:
            if col not in df_s.columns:
                df_s[col] = 0
        df_s["Stock Disponible"] = pd.to_numeric(df_s["Stock Disponible"], errors="coerce").fillna(0)
        df_s["Stock Fisico"] = pd.to_numeric(df_s["Stock Fisico"], errors="coerce").fillna(0)
    return merge_tarifa_stock(df_t, df_s)


with st.spinner("Procesando datos..."):
    df_full = build_merged(
        selected_country,
        _tarifa_xls=tarifa_xls,
        _stock_xls=stock_xls,
        _tarifa_df_remote=tarifa_dfs_remote,
        _stock_sheets_remote=stock_sheets_remote,
    )

# ── Dynamic sidebar filters (need data to populate options) ──────────────────
familias = sorted(df_full["FAMILIA"].dropna().unique().tolist()) if "FAMILIA" in df_full.columns else []

with filter_placeholder:
    familia_filter = st.multiselect("Familia", familias, placeholder="Todas")

# Cascade: subfamilia options depend on selected familias
if familia_filter:
    sub_df = df_full[df_full["FAMILIA"].isin(familia_filter)]
else:
    sub_df = df_full
subfamilias = sorted(sub_df["SUBFAMILIA"].dropna().unique().tolist()) if "SUBFAMILIA" in sub_df.columns else []

with filter_placeholder:
    subfamilia_filter = st.multiselect("Subfamilia", subfamilias, placeholder="Todas")

# Apply filters
df = df_full.copy()

if depos_filter and "DESPOSICIONADO" in df.columns:
    df = df[df["DESPOSICIONADO"] != True]

if familia_filter:
    df = df[df["FAMILIA"].isin(familia_filter)]
if subfamilia_filter:
    df = df[df["SUBFAMILIA"].isin(subfamilia_filter)]

if "Sin stock" in show_only:
    df = df[df["Sin stock"]]
if "Stock bajo" in show_only:
    df = df[df["Stock bajo"]]
if "En mar" in show_only:
    df = df[df["En mar"]]
if "En puerto" in show_only:
    df = df[df["En puerto"]]
if "Riesgo rotura" in show_only:
    df = df[df["Riesgo rotura"]]

df["Estado"] = df.apply(status_label, axis=1)

# ── KPI Cards ─────────────────────────────────────────────────────────────────
total = len(df)
sin_stock = int(df["Sin stock"].sum())
riesgo = int(df["Riesgo rotura"].sum())
bajo = int(df["Stock bajo"].sum())
en_mar = int(df["En mar"].sum())
en_puerto = int(df["En puerto"].sum())
ok_pct = round((total - sin_stock - bajo) / total * 100, 1) if total > 0 else 0

c1, c2, c3, c4, c5, c6 = st.columns(6)
cards = [
    (c1, "Total SKUs", total, "#3399ff"),
    (c2, "🔴 Sin stock", sin_stock, "#ff5555"),
    (c3, "⚠️ Riesgo rotura", riesgo, "#ff3333"),
    (c4, "🟠 Stock bajo", bajo, "#ffaa33"),
    (c5, "🚢 En mar", en_mar, "#33aaff"),
    (c6, "🏭 En puerto", en_puerto, "#aa88ff"),
]
for col, label, value, color in cards:
    with col:
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value" style="color:{color}">{value:,}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Vista General",
    "🔴 Sin Stock",
    "🚢 Tránsito (Mar / Puerto)",
    "📈 Por Categoría",
    "🔎 Búsqueda SKU",
])

# ─────────────────────── TAB 1: Vista general ────────────────────────────────
with tab1:
    col_a, col_b = st.columns([3, 2])

    with col_a:
        st.markdown('<div class="section-title">Estado por Familia</div>', unsafe_allow_html=True)
        if "FAMILIA" in df.columns:
            summary = df.groupby("FAMILIA").agg(
                Total=("REFERENCIA", "count"),
                Sin_Stock=("Sin stock", "sum"),
                Riesgo=("Riesgo rotura", "sum"),
                Stock_Bajo=("Stock bajo", "sum"),
                En_Mar=("En mar", "sum"),
            ).reset_index()
            summary["% Sin stock"] = (summary["Sin_Stock"] / summary["Total"] * 100).round(1)
            summary = summary.sort_values("% Sin stock", ascending=False)

            def color_pct(val):
                if val >= 60:
                    return "color: #ff5555; font-weight:600"
                elif val >= 30:
                    return "color: #ffaa33; font-weight:600"
                elif val >= 10:
                    return "color: #ffdd44"
                return "color: #33cc66"

            styled = summary.style.map(color_pct, subset=["% Sin stock"])
            st.dataframe(styled, use_container_width=True, height=420)

    with col_b:
        st.markdown('<div class="section-title">Distribución de Estado</div>', unsafe_allow_html=True)
        estado_counts = df["Estado"].value_counts()
        estado_df = pd.DataFrame({
            "Estado": estado_counts.index,
            "SKUs": estado_counts.values,
        })

        try:
            import plotly.express as px
            color_map = {
                "🔴 Sin stock": "#ff5555",
                "🟡 Sin stock (entrante)": "#ffdd00",
                "🟠 Stock bajo": "#ffaa33",
                "🟢 OK": "#33cc66",
            }
            fig = px.pie(
                estado_df,
                names="Estado",
                values="SKUs",
                color="Estado",
                color_discrete_map=color_map,
                hole=0.45,
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#c0d0e8",
                margin=dict(t=20, b=20, l=10, r=10),
                legend=dict(orientation="v", font_size=12),
                showlegend=True,
            )
            st.plotly_chart(fig, use_container_width=True)
        except ImportError:
            st.dataframe(estado_df, use_container_width=True)

    # Full table preview
    st.markdown('<div class="section-title">Detalle completo</div>', unsafe_allow_html=True)
    cols_show = ["REFERENCIA", "NOMBRE COMPLETO", "FAMILIA", "SUBFAMILIA",
                 "Stock Disponible", "Stock Fisico", "Mar", "Puerto", "Despachado", "Estado"]
    cols_available = [c for c in cols_show if c in df.columns]
    st.dataframe(df[cols_available].head(500), use_container_width=True, height=300)
    if len(df) > 500:
        st.caption(f"Mostrando 500 de {len(df)} filas. Usa filtros para reducir.")


# ─────────────────────── TAB 2: Sin Stock ────────────────────────────────────
with tab2:
    df_sin = df[df["Sin stock"]].copy()
    df_sin_riesgo = df_sin[df_sin["Riesgo rotura"]]
    df_sin_entrante = df_sin[~df_sin["Riesgo rotura"]]

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.markdown(f'<div class="metric-card"><div class="label">🔴 Sin stock sin reposición</div><div class="value" style="color:#ff5555">{len(df_sin_riesgo):,}</div></div>', unsafe_allow_html=True)
    with col_r2:
        st.markdown(f'<div class="metric-card"><div class="label">🟡 Sin stock CON entrante</div><div class="value" style="color:#ffdd00">{len(df_sin_entrante):,}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    subtab_a, subtab_b = st.tabs(["🔴 Riesgo real de rotura", "🟡 Sin stock pero con entrante"])

    def render_sin_stock_table(df_sub, btn_key, filename):
        cols = ["REFERENCIA", "NOMBRE COMPLETO", "FAMILIA", "SUBFAMILIA",
                "PVPR ", "NETO", "Stock Fisico", "Stock Disponible",
                "Mar", "Puerto", "Despachado", "Total entrante"]
        cols_ok = [c for c in cols if c in df_sub.columns]
        st.dataframe(df_sub[cols_ok].reset_index(drop=True), use_container_width=True, height=420)
        csv = df_sub[cols_ok].to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Exportar CSV", csv, filename, "text/csv", key=btn_key)

    with subtab_a:
        st.markdown('<div class="section-title">SKUs sin stock y sin ningún entrante (riesgo inmediato)</div>', unsafe_allow_html=True)
        render_sin_stock_table(df_sin_riesgo, "dl_riesgo", "sin_stock_riesgo.csv")

    with subtab_b:
        st.markdown('<div class="section-title">SKUs sin stock pero con mercancía en tránsito</div>', unsafe_allow_html=True)
        render_sin_stock_table(df_sin_entrante, "dl_entrante", "sin_stock_entrante.csv")


# ─────────────────────── TAB 3: Tránsito ─────────────────────────────────────
with tab3:
    col_m, col_p = st.columns(2)

    with col_m:
        st.markdown('<div class="section-title">🚢 Stock en Mar</div>', unsafe_allow_html=True)
        df_mar = df[df["Mar"] > 0].copy()
        st.markdown(f"**{len(df_mar):,} SKUs** con stock en tránsito marítimo")
        cols_mar = ["REFERENCIA", "NOMBRE COMPLETO", "FAMILIA", "Stock Disponible", "Mar", "Puerto", "Despachado", "Sin stock"]
        cols_ok = [c for c in cols_mar if c in df_mar.columns]
        st.dataframe(df_mar[cols_ok].reset_index(drop=True), use_container_width=True, height=400)
        csv = df_mar[cols_ok].to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Exportar Mar CSV", csv, "en_mar.csv", "text/csv", key="dl_mar")

    with col_p:
        st.markdown('<div class="section-title">🏭 Stock en Puerto</div>', unsafe_allow_html=True)
        df_puerto = df[df["Puerto"] > 0].copy()
        if len(df_puerto) == 0:
            st.info("No hay stock en puerto actualmente.")
        else:
            st.markdown(f"**{len(df_puerto):,} SKUs** en puerto")
            cols_p = ["REFERENCIA", "NOMBRE COMPLETO", "FAMILIA", "Stock Disponible", "Puerto", "Despachado", "Sin stock"]
            cols_ok = [c for c in cols_p if c in df_puerto.columns]
            st.dataframe(df_puerto[cols_ok].reset_index(drop=True), use_container_width=True, height=400)
            csv = df_puerto[cols_ok].to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Exportar Puerto CSV", csv, "en_puerto.csv", "text/csv", key="dl_puerto")

    # Combined table: sin stock but incoming
    st.markdown("---")
    st.markdown('<div class="section-title">📋 Sin stock almacén pero con entrante (mar + puerto + despachado)</div>', unsafe_allow_html=True)
    df_combo = df[(df["Sin stock"]) & (df["Total entrante"] > 0)].copy()
    cols_c = ["REFERENCIA", "NOMBRE COMPLETO", "FAMILIA", "Stock Disponible", "Mar", "Puerto", "Despachado", "Total entrante"]
    cols_ok = [c for c in cols_c if c in df_combo.columns]
    st.dataframe(df_combo[cols_ok].reset_index(drop=True), use_container_width=True, height=300)


# ─────────────────────── TAB 4: Por Categoría ────────────────────────────────
with tab4:
    try:
        import plotly.express as px
        import plotly.graph_objects as go

        if "FAMILIA" in df.columns:
            st.markdown('<div class="section-title">Top familias con mayor % sin stock</div>', unsafe_allow_html=True)
            fam_stats = df.groupby("FAMILIA").agg(
                Total=("REFERENCIA", "count"),
                Sin_Stock=("Sin stock", "sum"),
                Riesgo=("Riesgo rotura", "sum"),
                En_Mar=("En mar", "sum"),
                Stock_Bajo=("Stock bajo", "sum"),
            ).reset_index()
            fam_stats = fam_stats[fam_stats["Total"] >= 3]
            fam_stats["pct_sin_stock"] = (fam_stats["Sin_Stock"] / fam_stats["Total"] * 100).round(1)
            fam_stats = fam_stats.sort_values("pct_sin_stock", ascending=True).tail(25)

            fig_bar = go.Figure(go.Bar(
                x=fam_stats["pct_sin_stock"],
                y=fam_stats["FAMILIA"],
                orientation="h",
                marker_color=[
                    "#ff5555" if x >= 60 else "#ffaa33" if x >= 30 else "#33cc66"
                    for x in fam_stats["pct_sin_stock"]
                ],
                text=[f"{x}%" for x in fam_stats["pct_sin_stock"]],
                textposition="outside",
            ))
            fig_bar.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#c0d0e8",
                xaxis=dict(title="% Sin stock", color="#c0d0e8", gridcolor="#2a3550"),
                yaxis=dict(color="#c0d0e8"),
                height=600,
                margin=dict(l=10, r=60, t=20, b=20),
            )
            st.plotly_chart(fig_bar, use_container_width=True)

            # Stacked bar: stock vs sin stock por familia
            st.markdown('<div class="section-title">Composición de stock por familia (Top 20)</div>', unsafe_allow_html=True)
            top_fam = fam_stats.sort_values("Total", ascending=False).tail(20)
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                name="Con stock",
                x=top_fam["FAMILIA"],
                y=top_fam["Total"] - top_fam["Sin_Stock"],
                marker_color="#33cc66",
            ))
            fig2.add_trace(go.Bar(
                name="Sin stock (entrante)",
                x=top_fam["FAMILIA"],
                y=top_fam["En_Mar"],
                marker_color="#ffdd00",
            ))
            fig2.add_trace(go.Bar(
                name="Sin stock sin entrante",
                x=top_fam["FAMILIA"],
                y=top_fam["Riesgo"],
                marker_color="#ff5555",
            ))
            fig2.update_layout(
                barmode="stack",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#c0d0e8",
                xaxis=dict(color="#c0d0e8", tickangle=-35, gridcolor="#2a3550"),
                yaxis=dict(color="#c0d0e8", gridcolor="#2a3550"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, bgcolor="rgba(0,0,0,0)"),
                height=420,
                margin=dict(l=10, r=10, t=40, b=80),
            )
            st.plotly_chart(fig2, use_container_width=True)

    except ImportError:
        st.warning("Instala plotly para ver gráficos: `pip install plotly`")

    # Pivot table subfamilia
    if "SUBFAMILIA" in df.columns and "FAMILIA" in df.columns:
        st.markdown('<div class="section-title">Tabla resumen Subfamilia</div>', unsafe_allow_html=True)
        sub_stats = df.groupby(["FAMILIA", "SUBFAMILIA"]).agg(
            Total=("REFERENCIA", "count"),
            Sin_Stock=("Sin stock", "sum"),
            Riesgo=("Riesgo rotura", "sum"),
            En_Mar=("En mar", "sum"),
            Stock_Bajo=("Stock bajo", "sum"),
        ).reset_index()
        sub_stats["% Sin stock"] = (sub_stats["Sin_Stock"] / sub_stats["Total"] * 100).round(1)
        sub_stats = sub_stats.sort_values(["% Sin stock", "Total"], ascending=[False, False])
        st.dataframe(sub_stats, use_container_width=True, height=350)


# ─────────────────────── TAB 5: Búsqueda SKU ────────────────────────────────
with tab5:
    st.markdown('<div class="section-title">🔎 Búsqueda por Referencia o Nombre</div>', unsafe_allow_html=True)
    query = st.text_input("Buscar SKU, referencia o nombre de producto", placeholder="Ej: 120, cafeter, licuadora...")

    if query:
        mask = (
            df["REFERENCIA"].str.contains(query, case=False, na=False) |
            df.get("NOMBRE COMPLETO", pd.Series(dtype=str)).str.contains(query, case=False, na=False)
        )
        df_search = df[mask].copy()
        st.markdown(f"**{len(df_search)} resultados**")
        cols_s = ["REFERENCIA", "EAN", "NOMBRE COMPLETO", "FAMILIA", "SUBFAMILIA",
                  "PVPR ", "NETO", "Stock Fisico", "Stock Disponible",
                  "Mar", "Puerto", "Despachado", "Total entrante", "Estado"]
        cols_ok = [c for c in cols_s if c in df_search.columns]
        st.dataframe(df_search[cols_ok].reset_index(drop=True), use_container_width=True, height=420)
    else:
        st.info("Introduce una referencia, EAN o nombre para buscar.")

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"Dashboard generado · {datetime.now().strftime('%d/%m/%Y %H:%M')} · SKUs analizados: **{total:,}**")
