import json
from pathlib import Path
from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st

# =================python -m streamlit run app.py========================================
# 1. KONFIGURASI DASHBOARD
# =========================================================

st.set_page_config(
    page_title="SIPENA HIPERTENSI | Kota Bandung",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = BASE_DIR / "data_hipertensi.xlsx"
GEOJSON_PATH = BASE_DIR / "3273-kota-bandung-level-kecamatan.json"

INK = "#243B53"
MUTED = "#8996A5"
PURPLE = "#159A9C"       # variabel lama dipertahankan, kini teal
CYAN = "#6B9AC4"
BG = "#F4F8F9"

# =========================================================
# 2. DESAIN ANTARMUKA — MODERN HEALTHCARE
# =========================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');

:root {
    --navy: #183B4E;
    --teal: #168C89;
    --teal-dark: #087775;
    --mint: #DDF2EE;
    --mint-light: #F0F8F7;
    --text: #263C4A;
    --muted: #83939D;
    --border: #E3ECEC;
    --background: #F3F7F7;
}

.stApp {
    background: var(--background);
    color: var(--text);
    font-family: 'DM Sans', 'Segoe UI', sans-serif;
}

.block-container {
    max-width: 100%;
    padding: 1.5rem 2rem 3rem;
}

[data-testid="stHeader"] { background: transparent; }

/* SIDEBAR */
section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid var(--border);
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label {
    color: var(--navy) !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border: 1px solid #D9E7E6 !important;
    border-radius: 10px !important;
}

/* MULTISELECT TAG — OVERRIDE WARNA MERAH STREAMLIT */
/* Target spesifik pada komponen Tahun dan Kecamatan di sidebar. */
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] div[data-baseweb="tag"],
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"],
section[data-testid="stSidebar"] [data-baseweb="select"] div[data-baseweb="tag"],
section[data-testid="stSidebar"] [data-baseweb="select"] [data-baseweb="tag"] {
    background: #DDF2EE !important;
    background-color: #DDF2EE !important;
    border: 1px solid #A9D8CF !important;
    border-radius: 8px !important;
    box-shadow: none !important;
}
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] span,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] div,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] p {
    background: transparent !important;
    background-color: transparent !important;
    color: #087775 !important;
}
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] svg,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] path {
    fill: #087775 !important;
    color: #087775 !important;
}
/* Fallback untuk variasi struktur BaseWeb pada versi Streamlit lain. */
[data-baseweb="tag"] {
    background: #DDF2EE !important;
    background-color: #DDF2EE !important;
    color: #087775 !important;
    border: 1px solid #A9D8CF !important;
}
[data-baseweb="tag"] * {
    color: #087775 !important;
    fill: #087775 !important;
}
:root { --primary-color: #168C89; }

/* BUTTON */
.stButton > button {
    background: #168C89 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 10px 15px !important;
    font-weight: 700 !important;
    transition: all .2s ease;
}
.stButton > button:hover {
    background: #087775 !important;
    color: #FFFFFF !important;
    transform: translateY(-1px);
}

/* HERO HEADER */
.hero {
    position: relative;
    overflow: hidden;
    background: linear-gradient(112deg, #FFFFFF 0%, #E8F7F4 34%, #CDECE6 72%, #B8E1DA 100%);
    border: 1px solid #D8EAE7;
    border-radius: 22px;
    padding: 36px 38px;
    margin-bottom: 30px;
    box-shadow: 0 10px 30px rgba(24, 59, 78, .055);
}
.hero::after {
    content: "";
    position: absolute;
    right: -55px;
    top: -100px;
    width: 270px;
    height: 270px;
    border-radius: 50%;
    background: rgba(22, 140, 137, .15);
    box-shadow: 0 0 0 28px rgba(22, 140, 137, .045),
                0 0 0 58px rgba(22, 140, 137, .025);
}
.hero::before {
    content: "";
    position: absolute;
    right: 155px;
    bottom: -115px;
    width: 210px;
    height: 210px;
    border-radius: 50%;
    background: rgba(255, 255, 255, .30);
}
.hero-kicker {
    position: relative;
    z-index: 1;
    display: inline-block;
    color: #075F60;
    background: #BDE5DC;
    border: 1px solid #9ED4C9;
    border-left: 4px solid #168C89;
    border-radius: 8px;
    padding: 10px 15px;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-bottom: 15px;
    box-shadow: 0 3px 10px rgba(22, 140, 137, .08);
}
.hero-title {
    position: relative;
    z-index: 1;
    color: #183B4E;
    font-size: 37px;
    font-weight: 800;
    letter-spacing: -1.1px;
    margin: 0;
}
.hero-subtitle {
    position: relative;
    z-index: 1;
    color: #526F7B;
    font-size: 13px;
    margin-top: 13px;
}

/* SECTION HEADINGS */
.section-heading {
    color: #183B4E;
    font-size: 19px;
    font-weight: 800;
    margin: 27px 0 5px;
    letter-spacing: -.3px;
}
.section-caption {
    color: #83939D;
    font-size: 12px;
    margin-bottom: 16px;
}

/* KPI CARDS */
div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E3ECEC;
    border-top: 3px solid #168C89;
    border-radius: 16px;
    padding: 21px;
    min-height: 115px;
    box-shadow: 0 5px 18px rgba(24, 59, 78, .035);
}
div[data-testid="stMetricLabel"] {
    color: #83939D !important;
    font-size: 12px !important;
    font-weight: 600 !important;
}
div[data-testid="stMetricValue"] {
    color: #183B4E !important;
    font-size: 28px !important;
    font-weight: 800 !important;
}

/* CHART CARDS */
div[data-testid="stPlotlyChart"] {
    background: #FFFFFF;
    border: 1px solid #E3ECEC;
    border-radius: 17px;
    padding: 9px;
    box-shadow: 0 5px 18px rgba(24, 59, 78, .035);
}

/* DATA TABLE */
div[data-testid="stDataFrame"] {
    background: #FFFFFF;
    border: 1px solid #DDE9E8;
    border-radius: 15px;
    padding: 7px;
    overflow: hidden;
    box-shadow: 0 5px 18px rgba(24, 59, 78, .035);
}

/* DATA HASIL PENGAMATAN — CUSTOM TABLE */
.data-table-shell {
    background: #FFFFFF;
    border: 1px solid #DCEAE8;
    border-radius: 18px;
    padding: 13px;
    box-shadow: 0 12px 30px rgba(24, 59, 78, .065);
}
.data-table-scroll {
    max-height: 540px;
    overflow: auto;
    border-radius: 12px;
}
.data-table-scroll table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-family: 'DM Sans', sans-serif;
    font-size: 12px;
    color: #294452;
}
.data-table-scroll thead th {
    position: sticky;
    top: 0;
    z-index: 2;
    background: #183B4E;
    color: #FFFFFF;
    text-align: left;
    font-weight: 700;
    letter-spacing: .25px;
    padding: 16px 17px;
    border: none;
}
.data-table-scroll thead th:first-child { border-radius: 9px 0 0 0; }
.data-table-scroll thead th:last-child { border-radius: 0 9px 0 0; }
.data-table-scroll tbody td {
    padding: 14px 17px;
    border-bottom: 1px solid #EAF0EF;
    background: #FFFFFF;
    vertical-align: middle;
}
.data-table-scroll tbody tr:nth-child(even) td {
    background: #F1F8F6;
}
.data-table-scroll tbody tr:hover td {
    background: #DDF2EE;
    transition: background .15s ease;
}
.data-table-scroll tbody td:last-child {
    text-align: right;
    font-weight: 700;
    color: #087775;
    font-variant-numeric: tabular-nums;
}
.table-note {
    color: #83939D;
    font-size: 11px;
    padding: 11px 4px 1px;
}

/* FOOTER */
.footer {
    text-align: center;
    color: #8B9BA3;
    font-size: 11px;
    padding: 27px 0 8px;
    border-top: 1px solid var(--border);
    margin-top: 30px;
}

@media (max-width: 900px) {
    .block-container { padding-left: 1rem; padding-right: 1rem; }
    .hero { padding: 23px; }
    .hero-title { font-size: 28px; }
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# 3. MEMUAT DAN MEMVALIDASI DATA
# =========================================================

if not EXCEL_PATH.exists():
    st.error(
        "File data_hipertensi.xlsx tidak ditemukan. "
        "Letakkan file Excel di folder yang sama dengan app.py."
    )
    st.stop()

try:
    df = pd.read_excel(EXCEL_PATH)

except Exception as exc:
    st.error(f"File Excel gagal dibaca: {exc}")
    st.stop()

required_columns = [
    "tahun",
    "jumlah_penderita_hipertensi",
    "bps_nama_kecamatan",
    "fasilitas_kesehatan",
]

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]

if missing_columns:
    st.error(
        "Kolom berikut tidak ditemukan di Excel: "
        + ", ".join(missing_columns)
    )
    st.stop()

df["tahun"] = pd.to_numeric(
    df["tahun"],
    errors="coerce"
)

df["jumlah_penderita_hipertensi"] = pd.to_numeric(
    df["jumlah_penderita_hipertensi"],
    errors="coerce"
)

df["bps_nama_kecamatan"] = (
    df["bps_nama_kecamatan"]
    .astype("string")
    .str.strip()
)

df["fasilitas_kesehatan"] = (
    df["fasilitas_kesehatan"]
    .astype("string")
    .str.strip()
)

df = df.dropna(subset=required_columns).copy()

df = df[
    (df["bps_nama_kecamatan"] != "")
    & (df["fasilitas_kesehatan"] != "")
].copy()

if df.empty:
    st.error("Tidak terdapat baris data valid di dalam file Excel.")
    st.stop()

df["tahun"] = df["tahun"].astype(int)

# =========================================================
# 4. MEMUAT DATA GEOJSON
# =========================================================

geojson_data = None

if GEOJSON_PATH.exists():

    try:
        with GEOJSON_PATH.open("r", encoding="utf-8") as file:
            geojson_data = json.load(file)

    except (json.JSONDecodeError, OSError) as exc:
        st.warning(f"File peta tidak dapat dibaca: {exc}")
        # =========================================================
# =========================================================
# 5. FILTER DATA
# =========================================================

st.sidebar.markdown("## FILTER DATA")

st.sidebar.caption(
    "Atur wilayah dan periode pengamatan."
)

years = sorted(
    df["tahun"].unique().tolist()
)

districts = sorted(
    df["bps_nama_kecamatan"].dropna().unique().tolist()
)

# 1. WILAYAH PENGAMATAN (DILETAKKAN DI ATAS)

with st.sidebar.expander("Wilayah Pengamatan", expanded=True):

    selected_districts = st.multiselect(
        "Pilih Kecamatan",
        options=districts,
        default=districts,
        key="filter_kecamatan"
    )

# 2. PERIODE PENGAMATAN

with st.sidebar.expander("Periode Pengamatan", expanded=True):

    selected_years = st.multiselect(
        "Pilih Tahun",
        options=years,
        default=years,
        key="filter_tahun"
    )

# 3. RESET FILTER

if st.sidebar.button(
    "Reset Filter",
    use_container_width=True
):
    st.session_state["filter_tahun"] = years
    st.session_state["filter_kecamatan"] = districts
    st.rerun()

st.sidebar.divider()

st.sidebar.caption(
    "Sumber: Dinas Kesehatan Kota Bandung melalui Galura Data Jabar"
)

# 4. VALIDASI FILTER

if not selected_years or not selected_districts:

    st.warning(
        "Silakan pilih minimal satu tahun dan satu kecamatan."
    )
    st.stop()

# 5. FILTER DATA

filtered = df[
    df["tahun"].isin(selected_years)
    & df["bps_nama_kecamatan"].isin(selected_districts)
].copy()

if filtered.empty:

    st.warning(
        "Tidak terdapat data untuk kombinasi filter yang dipilih."
    )
    st.stop()
# =========================================================
# 6. HEADER DASHBOARD
# =========================================================

year_text = (
    f"{min(selected_years)}–{max(selected_years)}"
    if selected_years
    else "Tidak ada tahun dipilih"
)

header_html = (
    '<div class="hero">'
    '<div class="hero-kicker">'
    'SISTEM INFORMASI KESEHATAN KOTA BANDUNG'
    '</div>'
    '<h1 class="hero-title">'
    'SIPENA HIPERTENSI'
    '</h1>'
    '<p class="hero-subtitle">'
    'Sistem Informasi Pemantauan Hipertensi Kota Bandung'
    f' · {year_text}'
    '</p>'
    '</div>'
)

st.markdown(header_html, unsafe_allow_html=True)

# =========================================================
# 7. INDIKATOR RINGKASAN DATA
# =========================================================

st.markdown(
    '<div class="section-heading">Ringkasan Data</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Indikator ringkasan berdasarkan periode dan wilayah yang dipilih.'
    '</div>',
    unsafe_allow_html=True
)

k1, k2, k3, k4 = st.columns(4)

k1.metric(
    "Jumlah Observasi",
    f"{len(filtered):,}"
)

k2.metric(
    "Jumlah Kecamatan",
    f"{filtered['bps_nama_kecamatan'].nunique():,}"
)

k3.metric(
    "Fasilitas Kesehatan",
    f"{filtered['fasilitas_kesehatan'].nunique():,}"
)

k4.metric(
    "Cakupan Tahun",
    f"{filtered['tahun'].nunique():,}"
)

# =========================================================
# 8. DISTRIBUSI DATA MENURUT KECAMATAN
# FILTER JUMLAH KECAMATAN
# =========================================================

st.markdown(
    '<div class="section-heading">'
    '01. Distribusi Data Menurut Kecamatan'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Distribusi jumlah penderita hipertensi berdasarkan kecamatan '
    'selama periode pengamatan.'
    '</div>',
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# A. AGREGASI DATA
# ---------------------------------------------------------

by_district = (
    filtered.groupby(
        "bps_nama_kecamatan",
        as_index=False
    )["jumlah_penderita_hipertensi"]
    .sum()
    .sort_values(
        "jumlah_penderita_hipertensi",
        ascending=False
    )
)

total_kecamatan = len(by_district)

# ---------------------------------------------------------
# B. FILTER JUMLAH KECAMATAN
# ---------------------------------------------------------

col_info, col_filter = st.columns([2, 1])

with col_info:
    st.caption(
        f"Total tersedia: {total_kecamatan} kecamatan"
    )

with col_filter:
    pilihan_kecamatan = st.selectbox(
        "Jumlah Kecamatan Ditampilkan",
        options=[
            "Semua Kecamatan",
            5,
            10,
            15,
            20,
            25,
            30
        ],
        index=3,
        key="filter_jumlah_kecamatan"
    )

# ---------------------------------------------------------
# C. MENENTUKAN DATA GRAFIK
# ---------------------------------------------------------

if pilihan_kecamatan == "Semua Kecamatan":

    district_display = by_district.copy()

else:

    district_display = by_district.head(
        pilihan_kecamatan
    ).copy()

district_display = district_display.sort_values(
    "jumlah_penderita_hipertensi",
    ascending=True
)

jumlah_ditampilkan = len(district_display)

# ---------------------------------------------------------
# D. MEMBUAT GRAFIK
# ---------------------------------------------------------

fig_district = px.bar(
    district_display,

    x="jumlah_penderita_hipertensi",

    y="bps_nama_kecamatan",

    orientation="h",

    text="jumlah_penderita_hipertensi",

    color="jumlah_penderita_hipertensi",

    color_continuous_scale=[
        "#C5E9E2",
        "#65BDB3",
        "#168C89",
        "#075D66"
    ],

    labels={
        "jumlah_penderita_hipertensi": "Jumlah Penderita",
        "bps_nama_kecamatan": "Kecamatan"
    }
)

# ---------------------------------------------------------
# E. PENGATURAN UKURAN DAN TAMPILAN
# ---------------------------------------------------------

tinggi_grafik = max(
    350,
    min(700, 180 + jumlah_ditampilkan * 17)
)

fig_district.update_layout(
    template="plotly_white",

    height=tinggi_grafik,

    margin=dict(
        l=20,
        r=65,
        t=20,
        b=45
    ),

    showlegend=False,

    coloraxis_showscale=False,

    font=dict(
        family="DM Sans, Segoe UI, sans-serif",
        size=11,
        color="#243B53"
    ),

    xaxis=dict(
        title="Jumlah Penderita",
        gridcolor="#EDF2F3",
        zeroline=False
    ),

    yaxis=dict(
        title="",
        automargin=True
    ),

    bargap=0.30,

    plot_bgcolor="#FFFFFF",

    paper_bgcolor="#FFFFFF"
)

# ---------------------------------------------------------
# F. LABEL DAN INFORMASI GRAFIK
# ---------------------------------------------------------

fig_district.update_traces(
    texttemplate="%{x:,.0f}",

    textposition="outside",

    cliponaxis=False,

    marker_line_width=0,

    hovertemplate=(
        "<b>%{y}</b><br>"
        "Jumlah penderita: %{x:,.0f}"
        "<extra></extra>"
    )
)

# ---------------------------------------------------------
# G. TAMPILKAN GRAFIK
# ---------------------------------------------------------

st.plotly_chart(
    fig_district,
    use_container_width=True,
    config={
        "displayModeBar": False,
        "responsive": True
    }
)

st.caption(
    f"Menampilkan {jumlah_ditampilkan} dari "
    f"{total_kecamatan} kecamatan berdasarkan "
    "filter periode dan wilayah yang dipilih."
)

# =========================================================
# 9. DISTRIBUSI FASILITAS DAN TREN TAHUNAN
# =========================================================

left, right = st.columns(2, gap="large")

# ---------------------------------------------------------
# 9A. DISTRIBUSI FASILITAS KESEHATAN
# MODERN HEALTHCARE — VERTICAL COLUMN CHART
# ---------------------------------------------------------

with left:

    st.markdown(
        '<div class="section-heading">'
        '02. Distribusi Menurut Fasilitas Kesehatan'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-caption">'
        'Lima belas fasilitas kesehatan dengan jumlah penderita '
        'hipertensi tertinggi.'
        '</div>',
        unsafe_allow_html=True
    )

    # 1. AGREGASI DATA
    by_facility = (
        filtered.groupby(
            "fasilitas_kesehatan",
            as_index=False
        )["jumlah_penderita_hipertensi"]
        .sum()
        .nlargest(
            15,
            "jumlah_penderita_hipertensi"
        )
        .sort_values(
            "jumlah_penderita_hipertensi",
            ascending=False
        )
    )

    # 2. MEMBUAT VERTICAL COLUMN CHART
    fig_facility = px.bar(
        by_facility,
        x="fasilitas_kesehatan",
        y="jumlah_penderita_hipertensi",
        text="jumlah_penderita_hipertensi",
        color="jumlah_penderita_hipertensi",
        color_continuous_scale=[
            "#CDECE6",
            "#8BD2C5",
            "#36A99D",
            "#087775",
            "#075D66"
        ],
        labels={
            "fasilitas_kesehatan": "Fasilitas Kesehatan",
            "jumlah_penderita_hipertensi": "Jumlah Penderita"
        }
    )

    # 3. PENGATURAN KOLOM
    fig_facility.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        cliponaxis=False,
        marker_line_width=0,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Jumlah penderita: %{y:,.0f}"
            "<extra></extra>"
        )
    )

    # 4. PENGATURAN LAYOUT
    fig_facility.update_layout(
        template="plotly_white",
        height=560,
        margin=dict(
            l=15,
            r=25,
            t=55,
            b=150
        ),
        showlegend=False,
        coloraxis_showscale=False,
        font=dict(
            family="DM Sans, Segoe UI, sans-serif",
            size=11,
            color="#263C4A"
        ),
        xaxis=dict(
            title="Fasilitas Kesehatan",
            tickangle=-45,
            automargin=True,
            tickfont=dict(
                size=9,
                color="#526F7B"
            ),
            showgrid=False,
            linecolor="#DCE8E7"
        ),
        yaxis=dict(
            title="Jumlah Penderita",
            tickformat=",",
            rangemode="tozero",
            gridcolor="#EDF2F3",
            zeroline=False
        ),
        bargap=0.35,
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF"
    )

    # 5. TAMPILKAN GRAFIK
    st.plotly_chart(
        fig_facility,
        use_container_width=True,
        config={
            "displayModeBar": False,
            "responsive": True,
            "scrollZoom": False
        }
    )


# ---------------------------------------------------------
# 9B. TREN TAHUNAN JUMLAH PENDERITA HIPERTENSI
# MODERN HEALTHCARE — LINE + FILLED AREA
# ---------------------------------------------------------

with right:

    st.markdown(
        '<div class="section-heading">'
        '03. Tren Tahunan Jumlah Penderita Hipertensi'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-caption">'
        'Perkembangan jumlah penderita hipertensi setiap tahun '
        'berdasarkan wilayah pengamatan yang dipilih.'
        '</div>',
        unsafe_allow_html=True
    )

    # 1. AGREGASI DATA
    by_year = (
        filtered.groupby(
            "tahun",
            as_index=False
        )["jumlah_penderita_hipertensi"]
        .sum()
        .sort_values("tahun")
    )

    # 2. MEMBUAT GRAFIK AREA
    fig_year = px.area(
        by_year,
        x="tahun",
        y="jumlah_penderita_hipertensi",
        color_discrete_sequence=["#168C89"],
        labels={
            "tahun": "Tahun",
            "jumlah_penderita_hipertensi": "Jumlah Penderita"
        }
    )

    # 3. MEMBUAT AREA LEBIH LEMBUT
    fig_year.update_traces(
        fill="tozeroy",
        fillcolor="rgba(22, 140, 137, 0.13)",
        line=dict(
            color="#168C89",
            width=3.5,
            shape="spline"
        ),
        mode="lines+markers+text",
        marker=dict(
            size=10,
            color="#FFFFFF",
            line=dict(
                color="#168C89",
                width=3
            )
        ),
        text=by_year["jumlah_penderita_hipertensi"],
        texttemplate="%{text:,.0f}",
        textposition="top center",
        textfont=dict(
            family="DM Sans",
            size=10,
            color="#183B4E"
        ),
        cliponaxis=False,
        hovertemplate=(
            "<b>Tahun %{x}</b><br>"
            "Jumlah penderita: %{y:,.0f}"
            "<extra></extra>"
        )
    )

    # 4. PENGATURAN LAYOUT
    fig_year.update_layout(
        template="plotly_white",
        height=460,
        margin=dict(
            l=15,
            r=25,
            t=55,
            b=45
        ),
        font=dict(
            family="DM Sans, Segoe UI, sans-serif",
            size=11,
            color="#526F7B"
        ),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        showlegend=False,
        hovermode="x unified"
    )

    # 5. SUMBU X
    fig_year.update_xaxes(
        title="Tahun",
        tickmode="linear",
        dtick=1,
        showgrid=False,
        showline=True,
        linecolor="#DCE8E7",
        tickfont=dict(
            size=11,
            color="#718793"
        )
    )

    # 6. SUMBU Y
    fig_year.update_yaxes(
        title="Jumlah Penderita",
        rangemode="tozero",
        tickformat=",",
        gridcolor="#EAF0EF",
        gridwidth=1,
        zeroline=False,
        tickfont=dict(
            size=10,
            color="#718793"
        )
    )

    # 7. TAMPILKAN GRAFIK
    st.plotly_chart(
        fig_year,
        use_container_width=True,
        config={
            "displayModeBar": False,
            "responsive": True,
            "scrollZoom": False
        }
    )

# =========================================================
# 10. DISTRIBUSI SPASIAL MENURUT KECAMATAN
# =========================================================

st.markdown(
    '<div class="section-heading">'
    '04. Distribusi Spasial Menurut Kecamatan'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Peta tematik menunjukkan distribusi jumlah penderita '
    'hipertensi menurut kecamatan di Kota Bandung.'
    '</div>',
    unsafe_allow_html=True
)

if geojson_data is None:

    st.warning(
        "File GeoJSON tidak tersedia sehingga peta tidak dapat ditampilkan."
    )

elif "bps_kode_kecamatan" not in filtered.columns:

    st.warning(
        "Kolom bps_kode_kecamatan tidak ditemukan di Excel."
    )

else:

    def normalize_code(value):

        if pd.isna(value):
            return None

        code = str(value).strip()

        if code.endswith(".0"):
            code = code[:-2]

        return code

    geo_features = geojson_data.get("features", [])

    geo_names = {}

    for feature in geo_features:

        props = feature.get("properties", {})

        code = normalize_code(
            props.get("id_kecamatan")
        )

        name = props.get("nama_kecamatan")

        if code and name:
            geo_names[code] = str(name).strip()

    map_source = filtered.copy()

    map_source["id_kecamatan"] = (
        map_source["bps_kode_kecamatan"]
        .apply(normalize_code)
    )

    map_data = (
        map_source.groupby(
            "id_kecamatan",
            as_index=False
        )["jumlah_penderita_hipertensi"]
        .sum()
    )

    map_data["nama_kecamatan"] = (
        map_data["id_kecamatan"]
        .map(geo_names)
    )

    map_data = map_data.dropna(
        subset=[
            "id_kecamatan",
            "nama_kecamatan"
        ]
    )

    if map_data.empty:

        st.warning(
            "Kode kecamatan pada Excel belum cocok "
            "dengan kode pada GeoJSON."
        )

    else:

        fig_map = px.choropleth_map(
            map_data,
            geojson=geojson_data,
            locations="id_kecamatan",
            featureidkey="properties.id_kecamatan",
            color="jumlah_penderita_hipertensi",
            hover_name="nama_kecamatan",
            hover_data={
                "id_kecamatan": False,
                "nama_kecamatan": False,
                "jumlah_penderita_hipertensi": ":,.0f",
            },
            color_continuous_scale=[
                "#DDF2EE",
                "#A9DCD3",
                "#168C89",
                "#075D66"
            ],
            map_style="carto-positron",
            center={
    "lat": -6.9147,
    "lon": 107.6098
},
zoom=11.5,
            opacity=0.78,
            labels={
                "jumlah_penderita_hipertensi": "Jumlah Penderita"
            },
        )

        fig_map.update_traces(
            marker_line_color="#FFFFFF",
            marker_line_width=1.2,
            hovertemplate=(
                "<b>%{hovertext}</b><br>"
                "Jumlah penderita: %{z:,.0f}"
                "<extra></extra>"
            ),
        )

        fig_map.update_layout(
            height=750,
            margin=dict(
                l=0,
                r=0,
                t=5,
                b=0
            ),
            font=dict(
                family="DM Sans, Segoe UI, sans-serif",
                color=INK
            ),
            coloraxis_colorbar=dict(
                title="Jumlah<br>tercatat"
            ),
        )

        st.plotly_chart(
            fig_map,
            use_container_width=True,
            config={"displayModeBar": False}
        )

# =========================================================
# 11. TABEL DATA
# =========================================================

st.markdown(
    '<div class="section-heading">'
    '05. Data Hasil Pengamatan'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Data rinci berdasarkan tahun, kecamatan, dan fasilitas kesehatan '
    'sesuai kriteria filter.'
    '</div>',
    unsafe_allow_html=True
)

display_columns = [
    col for col in [
        "tahun",
        "bps_nama_kecamatan",
        "fasilitas_kesehatan",
        "jumlah_penderita_hipertensi",
    ]
    if col in filtered.columns
]

table_data = filtered[display_columns].sort_values(
    [
        "tahun",
        "bps_nama_kecamatan",
        "fasilitas_kesehatan"
    ]
)

table_labels = {
    "tahun": "Tahun",
    "bps_nama_kecamatan": "Kecamatan",
    "fasilitas_kesehatan": "Fasilitas Kesehatan",
    "jumlah_penderita_hipertensi": "Jumlah Penderita",
}

pretty_table = table_data.rename(columns=table_labels).copy()
pretty_table["Jumlah Penderita"] = pretty_table["Jumlah Penderita"].map(
    lambda value: f"{value:,.0f}"
)

# HTML table digunakan agar warna header, baris selang-seling,
# hover, dan sticky header tampil konsisten di browser Streamlit.
table_html = pretty_table.to_html(
    index=False,
    border=0,
    classes="healthcare-table",
    escape=True
)

st.markdown(
    '<div class="data-table-shell">'
    '<div class="data-table-scroll">'
    + table_html +
    '</div>'
    '<div class="table-note">'
    f'Menampilkan {len(pretty_table):,} baris data sesuai filter yang dipilih.'
    '</div>'
    '</div>',
    unsafe_allow_html=True
)

# =========================================================
# 12. FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer">
       
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# 7A. ANALISIS TAMBAHAN DATA HIPERTENSI
# =========================================================

st.markdown(
    '<div class="section-heading">'
    'Analisis Perkembangan dan Perbandingan'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Analisis perubahan jumlah penderita hipertensi '
    'dan perbandingan antarwilayah berdasarkan filter.'
    '</div>',
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# 1. ANALISIS PERUBAHAN TAHUNAN
# ---------------------------------------------------------

annual_analysis = (
    filtered.groupby("tahun", as_index=False)
    ["jumlah_penderita_hipertensi"]
    .sum()
    .sort_values("tahun")
)

if len(annual_analysis) >= 2:

    first_year = annual_analysis.iloc[0]
    last_year = annual_analysis.iloc[-1]

    first_value = first_year["jumlah_penderita_hipertensi"]
    last_value = last_year["jumlah_penderita_hipertensi"]

    difference = last_value - first_value

    if first_value != 0:
        percentage = (difference / first_value) * 100
    else:
        percentage = 0

    col_a, col_b, col_c = st.columns(3)

    col_a.metric(
        "Jumlah Tahun Awal",
        f"{first_value:,.0f}"
    )

    col_b.metric(
        "Jumlah Tahun Akhir",
        f"{last_value:,.0f}"
    )

    col_c.metric(
        "Perubahan",
        f"{difference:+,.0f}",
        f"{percentage:+.2f}%"
    )

    if difference > 0:
        st.warning(
            f"Jumlah penderita hipertensi meningkat "
            f"sebanyak {difference:,.0f} "
            f"({percentage:.2f}%) dari tahun "
            f"{int(first_year['tahun'])} hingga "
            f"{int(last_year['tahun'])}."
        )

    elif difference < 0:
        st.success(
            f"Jumlah penderita hipertensi menurun "
            f"sebanyak {abs(difference):,.0f} "
            f"({abs(percentage):.2f}%) dari tahun "
            f"{int(first_year['tahun'])} hingga "
            f"{int(last_year['tahun'])}."
        )

    else:
        st.info(
            "Jumlah penderita hipertensi tidak berubah "
            "antara tahun awal dan tahun akhir."
        )

else:
    st.info(
        "Pilih minimal dua tahun untuk melihat "
        "perubahan jumlah penderita hipertensi."
    )


# ---------------------------------------------------------
# 2. PERINGKAT KECAMATAN INTERAKTIF
# ---------------------------------------------------------

st.markdown(
    '<div class="section-heading">'
    'Peringkat Kecamatan Berdasarkan Jumlah Penderita'
    '</div>',
    unsafe_allow_html=True
)

top_n = st.select_slider(
    "Tampilkan peringkat kecamatan",
    options=[5, 10, 15, 20, 30],
    value=10,
    key="top_n_kecamatan"
)

district_ranking = (
    filtered.groupby("bps_nama_kecamatan", as_index=False)
    ["jumlah_penderita_hipertensi"]
    .sum()
    .sort_values(
        "jumlah_penderita_hipertensi",
        ascending=False
    )
    .head(top_n)
)

fig_ranking = px.bar(
    district_ranking.sort_values(
        "jumlah_penderita_hipertensi",
        ascending=True
    ),
    x="jumlah_penderita_hipertensi",
    y="bps_nama_kecamatan",
    orientation="h",
    text="jumlah_penderita_hipertensi",
    color_discrete_sequence=["#159A9C"],
    labels={
        "bps_nama_kecamatan": "Kecamatan",
        "jumlah_penderita_hipertensi": "Jumlah Penderita"
    }
)

fig_ranking.update_traces(
    texttemplate="%{text:,.0f}",
    textposition="outside",
    cliponaxis=False,
    hovertemplate=(
        "<b>%{y}</b><br>"
        "Jumlah penderita: %{x:,.0f}"
        "<extra></extra>"
    )
)

fig_ranking.update_layout(
    template="plotly_white",
    height=max(400, top_n * 38),
    margin=dict(l=10, r=65, t=20, b=35),
    font=dict(
        family="DM Sans, Segoe UI, sans-serif",
        color=INK
    ),
    xaxis=dict(
        title="Jumlah Penderita",
        rangemode="tozero",
        gridcolor="#EDF2F3"
    ),
    yaxis=dict(
        title="",
        automargin=True
    )
)

st.plotly_chart(
    fig_ranking,
    use_container_width=True,
    config={"displayModeBar": False}
)


# ---------------------------------------------------------
# 3. PERBANDINGAN ANTARKECAMATAN
# ---------------------------------------------------------

st.markdown(
    '<div class="section-heading">'
    'Perbandingan Antar Kecamatan'
    '</div>',
    unsafe_allow_html=True
)

available_districts = sorted(
    filtered["bps_nama_kecamatan"].unique().tolist()
)

comparison_districts = st.multiselect(
    "Pilih kecamatan yang ingin dibandingkan",
    options=available_districts,
    default=available_districts[:3],
    key="comparison_districts"
)

if comparison_districts:

    comparison_data = (
        filtered[
            filtered["bps_nama_kecamatan"].isin(
                comparison_districts
            )
        ]
        .groupby(
            ["tahun", "bps_nama_kecamatan"],
            as_index=False
        )["jumlah_penderita_hipertensi"]
        .sum()
    )

    fig_comparison = px.line(
        comparison_data,
        x="tahun",
        y="jumlah_penderita_hipertensi",
        color="bps_nama_kecamatan",
        markers=True,
        labels={
            "tahun": "Tahun",
            "jumlah_penderita_hipertensi": "Jumlah Penderita",
            "bps_nama_kecamatan": "Kecamatan"
        }
    )

    fig_comparison.update_layout(
        template="plotly_white",
        height=450,
        margin=dict(l=10, r=15, t=20, b=35),
        font=dict(
            family="DM Sans, Segoe UI, sans-serif",
            color=INK
        ),
        xaxis=dict(
            title="Tahun",
            dtick=1,
            gridcolor="#EDF2F3"
        ),
        yaxis=dict(
            title="Jumlah Penderita",
            rangemode="tozero",
            gridcolor="#EDF2F3"
        ),
        hovermode="x unified"
    )

    st.plotly_chart(
        fig_comparison,
        use_container_width=True,
        config={"displayModeBar": False}
    )

else:
    st.info(
        "Silakan pilih minimal satu kecamatan "
        "untuk menampilkan perbandingan."
    )

    # =========================================================
# 10A. INSIGHT UTAMA OTOMATIS
# =========================================================

st.markdown(
    '<div class="section-heading">'
    'Insight Utama Otomatis'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Ringkasan otomatis berdasarkan data dan filter yang dipilih.'
    '</div>',
    unsafe_allow_html=True
)

insight_district = (
    filtered.groupby("bps_nama_kecamatan")
    ["jumlah_penderita_hipertensi"]
    .sum()
    .sort_values(ascending=False)
)

insight_year = (
    filtered.groupby("tahun")
    ["jumlah_penderita_hipertensi"]
    .sum()
    .sort_index()
)

total_cases = filtered["jumlah_penderita_hipertensi"].sum()

top_district = insight_district.index[0]
top_district_value = insight_district.iloc[0]

top_share = (
    top_district_value / total_cases * 100
    if total_cases > 0 else 0
)

insight_col1, insight_col2 = st.columns(2)

with insight_col1:
    st.info(
        f"**Kecamatan dengan jumlah tertinggi**\n\n"
        f"{top_district} mencatat "
        f"{top_district_value:,.0f} penderita "
        f"({top_share:.1f}% dari total data yang ditampilkan)."
    )

with insight_col2:
    if len(insight_year) >= 2:

        first_value = insight_year.iloc[0]
        last_value = insight_year.iloc[-1]

        change = last_value - first_value

        if first_value != 0:
            change_pct = change / first_value * 100
        else:
            change_pct = 0

        if change > 0:
            message = "meningkat"
        elif change < 0:
            message = "menurun"
        else:
            message = "tidak berubah"

        st.info(
            f"**Perubahan antarperiode**\n\n"
            f"Jumlah tercatat {message} sebesar "
            f"{abs(change):,.0f} ({abs(change_pct):.1f}%) "
            f"dari {insight_year.index[0]} "
            f"hingga {insight_year.index[-1]}."
        )

    else:
        st.info(
            "Pilih minimal dua tahun untuk melihat "
            "perubahan antarperiode."
        )


# =========================================================
# 10B. HEATMAP KECAMATAN × TAHUN
# =========================================================

st.markdown(
    '<div class="section-heading">'
    'Heatmap Kecamatan dan Tahun'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Perbandingan jumlah penderita hipertensi berdasarkan kecamatan dan tahun.'
    '</div>',
    unsafe_allow_html=True
)

heatmap_data = filtered.pivot_table(
    index="bps_nama_kecamatan",
    columns="tahun",
    values="jumlah_penderita_hipertensi",
    aggfunc="sum",
    fill_value=0
)

heatmap_data = heatmap_data.sort_index()

fig_heatmap = px.imshow(
    heatmap_data,
    labels={
        "x": "Tahun",
        "y": "Kecamatan",
        "color": "Jumlah Penderita"
    },
    color_continuous_scale=[
        "#F0F8F7",
        "#A9DCD3",
        "#168C89",
        "#075D66"
    ],
    aspect="auto",
    text_auto=".0f"
)

fig_heatmap.update_layout(
    template="plotly_white",
    height=max(500, len(heatmap_data) * 27),
    margin=dict(l=10, r=15, t=20, b=35),
    font=dict(
        family="DM Sans, Segoe UI, sans-serif",
        color=INK
    ),
    xaxis_title="Tahun",
    yaxis_title="Kecamatan"
)

st.plotly_chart(
    fig_heatmap,
    use_container_width=True,
    config={"displayModeBar": False}
)



# =========================================================
# 10D. DOWNLOAD DATA EXCEL / CSV
# =========================================================

st.markdown(
    '<div class="section-heading">'
    'Download Data'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Unduh data berdasarkan filter tahun dan kecamatan '
    'yang sedang digunakan.'
    '</div>',
    unsafe_allow_html=True
)

download_columns = [
    "tahun",
    "bps_nama_kecamatan",
    "fasilitas_kesehatan",
    "jumlah_penderita_hipertensi"
]

export_data = filtered[
    download_columns
].copy()

export_data = export_data.rename(
    columns={
        "tahun": "Tahun",
        "bps_nama_kecamatan": "Kecamatan",
        "fasilitas_kesehatan": "Fasilitas Kesehatan",
        "jumlah_penderita_hipertensi": "Jumlah Penderita Hipertensi"
    }
)

download_col1, download_col2 = st.columns(2)

# DOWNLOAD EXCEL
excel_buffer = BytesIO()

with pd.ExcelWriter(
    excel_buffer,
    engine="openpyxl"
) as writer:

    export_data.to_excel(
        writer,
        index=False,
        sheet_name="Data Hipertensi"
    )

excel_buffer.seek(0)

with download_col1:

    st.download_button(
        label="Download Excel (.xlsx)",
        data=excel_buffer,
        file_name="data_hipertensi_bandung.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

# DOWNLOAD CSV
csv_data = export_data.to_csv(
    index=False,
    encoding="utf-8-sig"
)

with download_col2:

    st.download_button(
        label="Download CSV (.csv)",
        data=csv_data,
        file_name="data_hipertensi_bandung.csv",
        mime="text/csv",
        use_container_width=True
    )


# =========================================================
# DASHBOARD HIPERTENSI KOTA BANDUNG
# MULTI-PAGE VERSION
# =========================================================

import json
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# 1. KONFIGURASI APLIKASI
# =========================================================

st.set_page_config(
    page_title="SIPENA HIPERTENSI | Kota Bandung",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = BASE_DIR / "data_hipertensi.xlsx"

GEOJSON_CANDIDATES = [
    BASE_DIR / "3273-kota-bandung-level-kecamatan.json",
    BASE_DIR / "3273-kota-bandung-level-kecamatan(1).json"
]

GEOJSON_PATH = next(
    (p for p in GEOJSON_CANDIDATES if p.exists()),
    None
)


# =========================================================
# 2. WARNA DASHBOARD
# =========================================================

NAVY = "#183B4E"
TEAL = "#168C89"
TEAL_DARK = "#087775"
MINT = "#DDF2EE"
BG = "#F3F7F7"
TEXT = "#263C4A"
MUTED = "#83939D"

PALETTE = [
    "#DDF2EE",
    "#A9DCD3",
    "#65BEB1",
    "#168C89",
    "#075D66"
]


# =========================================================
# 3. CSS MODERN HEALTHCARE
# =========================================================

st.markdown("""
<style>

@import url(
'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap'
);

.stApp {
    background: #F3F7F7;
    font-family: 'DM Sans', sans-serif;
    color: #263C4A;
}

.block-container {
    max-width: 100%;
    padding: 1.5rem 2rem 3rem;
}

[data-testid="stHeader"] {
    background: transparent;
}

/* SIDEBAR */

section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid #E3ECEC;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label {
    color: #183B4E !important;
}

/* TAG FILTER */

[data-baseweb="tag"] {
    background: #DDF2EE !important;
    border: 1px solid #A9D8CF !important;
    color: #087775 !important;
}

[data-baseweb="tag"] * {
    color: #087775 !important;
    fill: #087775 !important;
}

/* BUTTON */

.stButton > button,
.stDownloadButton > button {
    background: #168C89 !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
}

/* HERO */

.hero {
    background: linear-gradient(
        112deg,
        #FFFFFF,
        #E8F7F4,
        #BDE6DE
    );
    border: 1px solid #D8EAE7;
    border-radius: 22px;
    padding: 35px 38px;
    margin-bottom: 28px;
    box-shadow: 0 10px 30px rgba(24,59,78,.05);
}

.hero-kicker {
    color: #075F60;
    background: #BDE5DC;
    border-radius: 8px;
    padding: 9px 14px;
    display: inline-block;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 1.5px;
}

.hero-title {
    color: #183B4E;
    font-size: 36px;
    font-weight: 800;
    margin-top: 14px;
}

.hero-subtitle {
    color: #526F7B;
    font-size: 13px;
}

/* SECTION */

.section-heading {
    color: #183B4E;
    font-size: 20px;
    font-weight: 800;
    margin-top: 25px;
    margin-bottom: 5px;
}

.section-caption {
    color: #83939D;
    font-size: 12px;
    margin-bottom: 15px;
}

/* KPI */

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E3ECEC;
    border-top: 4px solid #168C89;
    border-radius: 16px;
    padding: 20px;
    box-shadow: 0 5px 18px rgba(24,59,78,.04);
}

div[data-testid="stMetricLabel"] {
    color: #83939D !important;
}

div[data-testid="stMetricValue"] {
    color: #183B4E !important;
    font-weight: 800 !important;
}

/* CHART */

div[data-testid="stPlotlyChart"] {
    background: #FFFFFF;
    border: 1px solid #E3ECEC;
    border-radius: 16px;
    padding: 8px;
}

/* INSIGHT */

.insight-box {
    background: linear-gradient(
        110deg,
        #FFFFFF,
        #E8F7F4
    );
    border-left: 5px solid #168C89;
    border-radius: 13px;
    padding: 18px;
    margin: 15px 0;
    line-height: 1.7;
}

/* FOOTER */

.footer {
    text-align: center;
    color: #83939D;
    font-size: 11px;
    padding: 25px;
    margin-top: 35px;
    border-top: 1px solid #E3ECEC;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# 4. MEMUAT DATA
# =========================================================

if not EXCEL_PATH.exists():

    st.error(
        "File data_hipertensi.xlsx tidak ditemukan. "
        "Pastikan file berada dalam folder yang sama dengan app.py."
    )

    st.stop()


try:

    df = pd.read_excel(EXCEL_PATH)

except Exception as e:

    st.error(f"Gagal membaca file Excel: {e}")

    st.stop()


# =========================================================
# 5. VALIDASI KOLOM
# =========================================================

required_columns = [
    "tahun",
    "jumlah_penderita_hipertensi",
    "bps_nama_kecamatan",
    "fasilitas_kesehatan"
]

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]

if missing_columns:

    st.error(
        "Kolom tidak ditemukan: "
        + ", ".join(missing_columns)
    )

    st.stop()


# =========================================================
# 6. PEMBERSIHAN DATA
# =========================================================

df["tahun"] = pd.to_numeric(
    df["tahun"],
    errors="coerce"
)

df["jumlah_penderita_hipertensi"] = pd.to_numeric(
    df["jumlah_penderita_hipertensi"],
    errors="coerce"
)

df["bps_nama_kecamatan"] = (
    df["bps_nama_kecamatan"]
    .astype("string")
    .str.strip()
)

df["fasilitas_kesehatan"] = (
    df["fasilitas_kesehatan"]
    .astype("string")
    .str.strip()
)

df = df.dropna(
    subset=required_columns
).copy()

df["tahun"] = df["tahun"].astype(int)

df["jumlah_penderita_hipertensi"] = (
    df["jumlah_penderita_hipertensi"]
    .clip(lower=0)
)

if df.empty:

    st.error("Dataset tidak memiliki data yang valid.")

    st.stop()


# =====================================================
# NAVIGASI MULTI-PAGE
# =====================================================

page = st.sidebar.radio(
    "NAVIGASI HALAMAN",
    [
        "Overview",
        "Analisis Wilayah",
        "Analisis Tren",
        "Peta Persebaran",
        "Data & Download"
    ],
    key="dashboard_page"
)

# =====================================================
# HALAMAN 1: OVERVIEW
# =====================================================

if page == "Overview":
    pass


# =====================================================
# HALAMAN 2: ANALISIS WILAYAH
# =====================================================

elif page == "Analisis Wilayah":

    st.title("Analisis Wilayah")

    # PINDAHKAN KODE DISTRIBUSI KECAMATAN
    # PINDAHKAN KODE RANKING KECAMATAN


# =====================================================
# HALAMAN 3: ANALISIS TREN
# =====================================================

elif page == "Analisis Tren":

    st.title("Analisis Tren")

    # PINDAHKAN KODE TREN TAHUNAN
    # PINDAHKAN KODE HEATMAP


# =====================================================
# HALAMAN 4: PETA PERSEBARAN
# =====================================================

elif page == "Peta Persebaran":

    st.title("Peta Persebaran Hipertensi")

    # PINDAHKAN KODE PETA GEOJSON


# =====================================================
# HALAMAN 5: DATA DAN DOWNLOAD
# =====================================================

elif page == "Data & Download":

    st.title("Data & Download")

    # PINDAHKAN KODE TABEL DATA
    # PINDAHKAN KODE DOWNLOAD EXCEL DAN CSV

  # =========================================================
# 11. SUMBER DATA
# =========================================================

st.markdown("---")

st.markdown(
    '<div class="section-heading">Sumber Data</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    **Sumber:** Dinas Kesehatan Kota Bandung melalui
    Galura Data Jabar.

    **Periode Data:** 2019–2025.

    """
)
  