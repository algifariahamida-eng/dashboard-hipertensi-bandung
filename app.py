import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# 1. KONFIGURASI DASHBOARD
# =========================================================

st.set_page_config(
    page_title="Dashboard Data Hipertensi Kota Bandung",
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
    max-width: 1500px;
    padding: 1.5rem 2.3rem 3rem;
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
# 5. FILTER DATA
# =========================================================

st.sidebar.markdown("## FILTER DATA")

st.sidebar.markdown(
    "Pilih periode dan wilayah untuk menampilkan data."
)

years = sorted(
    df["tahun"].unique().tolist()
)

districts = sorted(
    df["bps_nama_kecamatan"]
    .dropna()
    .unique()
    .tolist()
)

selected_years = st.sidebar.multiselect(
    "Tahun",
    options=years,
    default=years,
)

selected_districts = st.sidebar.multiselect(
    "Kecamatan",
    options=districts,
    default=districts,
)

# Override tambahan setelah komponen filter dibuat.
# Ini menargetkan chip pilihan Tahun dan Kecamatan secara langsung.
st.markdown("""
<style>
section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="tag"] {
    background: #DDF2EE !important;
    background-color: #DDF2EE !important;
    border: 1px solid #A9D8CF !important;
    border-radius: 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="tag"] span {
    color: #087775 !important;
    background: transparent !important;
}
section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="tag"] svg,
section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="tag"] path {
    color: #087775 !important;
    fill: #087775 !important;
}
</style>
""", unsafe_allow_html=True)

if st.sidebar.button(
    "Atur ulang filter",
    use_container_width=True
):
    st.rerun()

st.sidebar.divider()

st.sidebar.caption(
    "Sumber: Dinas Kesehatan Kota Bandung melalui Galura Data Jabar"
)

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
    'VISUALISASI DATA KESEHATAN MASYARAKAT'
    '</div>'
    '<h1 class="hero-title">'
    'Dashboard Data Hipertensi'
    '</h1>'
    '<p class="hero-subtitle">'
    'Distribusi dan perkembangan data hipertensi penduduk '
    'usia 15 tahun ke atas di Kota Bandung'
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
# =========================================================

st.markdown(
    '<div class="section-heading">'
    '01. Distribusi Data Menurut Kecamatan'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-caption">'
    'Jumlah penderita hipertensi yang tercatat pada setiap kecamatan '
    'selama periode pengamatan.'
    '</div>',
    unsafe_allow_html=True
)

by_district = (
    filtered.groupby(
        "bps_nama_kecamatan",
        as_index=False
    )["jumlah_penderita_hipertensi"]
    .sum()
    .sort_values(
        "jumlah_penderita_hipertensi",
        ascending=True
    )
)

fig_district = px.bar(
    by_district,
    x="jumlah_penderita_hipertensi",
    y="bps_nama_kecamatan",
    orientation="h",
    color="jumlah_penderita_hipertensi",
    color_continuous_scale=[
        "#DDF2EE",
        "#79C9C2",
        "#168C89",
        "#087775"
    ],
    labels={
        "bps_nama_kecamatan": "Kecamatan",
        "jumlah_penderita_hipertensi": "Jumlah Penderita",
    },
)

fig_district.update_layout(
    template="plotly_white",
    height=max(420, 24 * len(by_district) + 120),
    coloraxis_showscale=False,
    margin=dict(
        l=12,
        r=18,
        t=20,
        b=35
    ),
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
    ),
    hovermode="closest",
)

fig_district.update_traces(
    marker_line_width=0,
    hovertemplate=(
        "<b>%{y}</b><br>"
        "Jumlah penderita: %{x:,.0f}"
        "<extra></extra>"
    ),
)

st.plotly_chart(
    fig_district,
    use_container_width=True,
    config={"displayModeBar": False}
)

# =========================================================
# 9. DISTRIBUSI FASILITAS DAN TREN TAHUNAN
# =========================================================

left, right = st.columns(2, gap="large")

# ---------------------------------------------------------
# 9A. DISTRIBUSI FASILITAS KESEHATAN
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
            ascending=True
        )
    )

    fig_facility = px.bar(
        by_facility,
        x="jumlah_penderita_hipertensi",
        y="fasilitas_kesehatan",
        orientation="h",
        color_discrete_sequence=["#6D9EAD"],
        labels={
            "fasilitas_kesehatan": "Fasilitas Kesehatan",
            "jumlah_penderita_hipertensi": "Jumlah Penderita",
        },
    )

    fig_facility.update_layout(
        template="plotly_white",
        height=540,
        margin=dict(
            l=8,
            r=15,
            t=15,
            b=35
        ),
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
        ),
    )

    fig_facility.update_traces(
        marker_line_width=0,
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Jumlah penderita: %{x:,.0f}"
            "<extra></extra>"
        ),
    )

    st.plotly_chart(
        fig_facility,
        use_container_width=True,
        config={"displayModeBar": False}
    )

# ---------------------------------------------------------
# 9B. TREN TAHUNAN
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
        'Jumlah penderita hipertensi yang tercatat setiap tahun '
        'pada wilayah yang dipilih.'
        '</div>',
        unsafe_allow_html=True
    )

    by_year = (
        filtered.groupby(
            "tahun",
            as_index=False
        )["jumlah_penderita_hipertensi"]
        .sum()
        .sort_values("tahun")
    )

    fig_year = px.line(
        by_year,
        x="tahun",
        y="jumlah_penderita_hipertensi",
        markers=True,
        color_discrete_sequence=[PURPLE],
        labels={
            "tahun": "Tahun",
            "jumlah_penderita_hipertensi": "Jumlah Penderita",
        },
    )

    fig_year.update_traces(
        line=dict(width=3.5),
        marker=dict(
            size=9,
            color="#FFFFFF",
            line=dict(
                width=3,
                color=PURPLE
            )
        ),
        hovertemplate=(
            "Tahun: %{x}<br>"
            "Jumlah penderita: %{y:,.0f}"
            "<extra></extra>"
        ),
    )

    fig_year.update_layout(
        template="plotly_white",
        height=540,
        margin=dict(
            l=12,
            r=15,
            t=15,
            b=35
        ),
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
            tickformat=",",
            gridcolor="#EDF2F3"
        ),
        hovermode="x unified",
    )

    st.plotly_chart(
        fig_year,
        use_container_width=True,
        config={"displayModeBar": False}
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
            zoom=9.5,
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
            height=610,
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
        DASHBOARD DATA HIPERTENSI · KOTA BANDUNG<br>
        Visualisasi Data Kesehatan Masyarakat
    </div>
    """,
    unsafe_allow_html=True,
)