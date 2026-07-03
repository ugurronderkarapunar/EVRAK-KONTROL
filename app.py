"""
⚓ Gemi Personeli Nitelik Belgesi Takip Sistemi
================================================
Versiyon   : 1.4.3
Açıklama   : Personel belge takip, uyarılar, grafikler.
             v1.4.3: components.html -> st.html, date parse uyarısı giderildi.
Bağımlılıklar: streamlit, pandas, openpyxl, xlrd, plotly
"""

import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import plotly.express as px
import plotly.graph_objects as go
import json

# ──────────────────────────────────────────────────────────────────────────────
# SAYFA YAPILANDIRMASI
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Belge Takip Sistemi",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# ÖZEL CSS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #0B1929 !important;
    color: #E8EDF3 !important;
    font-family: 'Inter', sans-serif !important;
}
[data-testid="stSidebar"] {
    background: #0D2137 !important;
    border-right: 1px solid #1E3A52 !important;
}
.hero-block {
    background: linear-gradient(135deg, #0D2137 0%, #1A3A5C 60%, #0D2137 100%);
    border: 1px solid #2A5280;
    border-radius: 16px;
    padding: 36px 40px 28px 40px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.hero-block::before {
    content: "⚓";
    position: absolute;
    right: 32px; top: 50%;
    transform: translateY(-50%);
    font-size: 96px;
    opacity: 0.06;
    pointer-events: none;
}
.hero-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.1rem;
    font-weight: 800;
    color: #F0C040;
    letter-spacing: -0.5px;
    margin: 0 0 6px 0;
    line-height: 1.15;
}
.hero-sub { font-size: 0.95rem; color: #8BAEC8; margin: 0; max-width: 620px; }

[data-testid="metric-container"] {
    background: #112236 !important;
    border: 1px solid #1E3A52 !important;
    border-radius: 12px !important;
    padding: 18px 22px !important;
}
[data-testid="metric-container"] label {
    color: #8BAEC8 !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.04em !important;
    text-transform: uppercase !important;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #F0C040 !important;
    font-family: 'Syne', sans-serif !important;
    font-size: 2rem !important;
}
.alert-critical {
    background: linear-gradient(90deg, #3D0B0B, #1C0B0B);
    border-left: 5px solid #E74C3C;
    border-radius: 10px; padding: 18px 22px; margin-bottom: 20px;
    color: #FADADD; font-size: 0.97rem;
}
.alert-warning {
    background: linear-gradient(90deg, #3D2200, #1C1100);
    border-left: 5px solid #F39C12;
    border-radius: 10px; padding: 18px 22px; margin-bottom: 20px;
    color: #FDEBD0; font-size: 0.97rem;
}
.alert-ok {
    background: linear-gradient(90deg, #0B2D1A, #061A0E);
    border-left: 5px solid #2ECC71;
    border-radius: 10px; padding: 18px 22px; margin-bottom: 20px;
    color: #D5F5E3; font-size: 0.97rem;
}
.alert-info {
    background: linear-gradient(90deg, #0B1E30, #061018);
    border-left: 5px solid #3498DB;
    border-radius: 10px; padding: 18px 22px; margin-bottom: 20px;
    color: #D6EAF8; font-size: 0.97rem;
}
.section-label {
    font-size: 0.72rem; font-weight: 600; letter-spacing: 0.12em;
    text-transform: uppercase; color: #4A7FA5;
    margin-bottom: 12px; border-bottom: 1px solid #1E3A52; padding-bottom: 6px;
}
.card-box {
    background: #112236;
    border: 1px solid #1E3A52;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 16px;
}
[data-testid="stSidebar"] label { color: #8BAEC8 !important; font-size: 0.82rem !important; font-weight: 500 !important; }
[data-testid="stDataFrame"] { border: 1px solid #1E3A52 !important; border-radius: 12px !important; overflow: hidden !important; }
hr { border-color: #1E3A52 !important; }
.stButton > button {
    background: #1A3A5C; color: #F0C040;
    border: 1px solid #2A5280; border-radius: 8px;
    font-weight: 600; padding: 6px 18px;
}
.stButton > button:hover { background: #2A5280; color: #FFD966; border-color: #4A7FA5; }
.js-plotly-plot { border-radius: 12px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# SABİTLER
# ──────────────────────────────────────────────────────────────────────────────
UNVAN_LISTESI = [
    "— Atanmadı —",
    "Kaptan",
    "Başmakinist",
    "Güverte L.",
    "Gemici",
    "Yağcı",
    "Diğer",
]

SAGLIK_ANAHTAR_KELIMELER = ["sağlık", "saglik", "health", "yoklama"]

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(17,34,54,0.0)",
    plot_bgcolor="rgba(17,34,54,0.0)",
    font=dict(color="#E8EDF3", family="Inter"),
    margin=dict(l=16, r=16, t=36, b=16),
)

DURUM_RENKLERI = {
    "🔴 Süresi Dolmuş":   "#E74C3C",
    "🟠 Bu Hafta Bitiyor": "#E67E22",
    "🟡 Bu Ay Bitiyor":    "#F1C40F",
    "🟢 Geçerli":          "#2ECC71",
    "⚪ Bilinmiyor":        "#95A5A6",
}

# ──────────────────────────────────────────────────────────────────────────────
# YARDIMCI FONKSİYONLAR
# ──────────────────────────────────────────────────────────────────────────────
def get_month_range(year: int, month: int):
    first_day = datetime.date(year, month, 1)
    last_day  = datetime.date(year, month, calendar.monthrange(year, month)[1])
    return first_day, last_day

def get_month_options() -> list:
    today = datetime.date.today()
    TR_AYLAR = ["","Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
                "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"]
    options = []
    for i in range(12):
        mo = today.month - 1 + i
        y, m = today.year + mo // 12, mo % 12 + 1
        label = f"{TR_AYLAR[m]} {y}" + ("  ← Cari Ay" if i == 0 else "")
        options.append((label, y, m))
    return options

def is_saglik_belgesi(evrak_adi: str) -> bool:
    ad = str(evrak_adi).lower()
    return any(k in ad for k in SAGLIK_ANAHTAR_KELIMELER)

def hesapla_bitis(baslangic: datetime.date, evrak_adi: str) -> datetime.date:
    yil = 2 if is_saglik_belgesi(evrak_adi) else 5
    try:
        return baslangic.replace(year=baslangic.year + yil)
    except ValueError:
        return baslangic.replace(year=baslangic.year + yil, day=28)

def load_excel(file):
    raw = file.read(); file.seek(0)
    magic = raw[:8]
    if magic == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        try: return pd.read_excel(io.BytesIO(raw), engine="xlrd")
        except Exception as e: st.error(f"XLS okunamadı: {e}"); return None
    if raw[:4] == b'PK\x03\x04':
        try: return pd.read_excel(io.BytesIO(raw), engine="openpyxl")
        except Exception as e: st.error(f"XLSX okunamadı: {e}"); return None
    for enc in ["windows-1254","iso-8859-9","utf-8-sig","utf-8","latin-1"]:
        try: text = raw.decode(enc); break
        except: continue
    else:
        st.error("Dosya kodlaması tanınamadı."); return None
    first = text.split("\n")[0]
    sep = "\t" if first.count("\t") >= first.count(",") else ","
    try:
        df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Metin dosyası okunamadı: {e}"); return None

def normalize_columns(df: pd.DataFrame):
    def norm(s):
        return (str(s).strip().lower()
                .replace("ı","i").replace("ğ","g").replace("ü","u")
                .replace("ş","s").replace("ö","o").replace("ç","c")
                .replace(" ",""))
    col_map = {
        "ad":"Adı Soyadı","adsoyadi":"Adı Soyadı","adisoyadı":"Adı Soyadı",
        "adisoyadi":"Adı Soyadı","adsoyad":"Adı Soyadı","personel":"Adı Soyadı",
        "sicilno":"Sicil No","sicil":"Sicil No","sicilnumarasi":"Sicil No",
        "evrakadi":"Evrak Adı","evrak":"Evrak Adı","nitelik":"Evrak Adı",
        "nitelikbelgesi":"Evrak Adı","belge":"Evrak Adı","belgeadi":"Evrak Adı",
        "baslangiçtarihi":"Başlangıç Tarihi","baslangic":"Başlangıç Tarihi",
        "baslangictarihi":"Başlangıç Tarihi","baslangiç":"Başlangıç Tarihi",
        "baslangic tarihi":"Başlangıç Tarihi",
        "bitistarihi":"Bitiş Tarihi","bitis":"Bitiş Tarihi","bitiş":"Bitiş Tarihi",
        "bitiştarihi":"Bitiş Tarihi","bitis tarihi":"Bitiş Tarihi",
        "gecerlilik":"Bitiş Tarihi","gecerliliktarihi":"Bitiş Tarihi",
    }
    renamed = {c: col_map[norm(c)] for c in df.columns if norm(c) in col_map}
    df = df.rename(columns=renamed)
    missing = [c for c in ["Adı Soyadı","Evrak Adı","Bitiş Tarihi"] if c not in df.columns]
    if missing:
        st.error(f"Zorunlu sütunlar bulunamadı: **{', '.join(missing)}**")
        return None
    for opt in ["Sicil No","Başlangıç Tarihi"]:
        if opt not in df.columns:
            df[opt] = "-"
    return df

def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    def parse_one(val):
        if pd.isna(val) or str(val).strip() in ("","nan","None","-"): return None
        s = str(val).strip()
        for fmt in ["%d.%m.%Y","%d/%m/%Y","%Y-%m-%d","%d.%m.%y"]:
            try: return datetime.datetime.strptime(s, fmt).date()
            except: pass
        return None
    for col in ["Bitiş Tarihi","Başlangıç Tarihi"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_one)
    return df

def compute_remaining_days(df: pd.DataFrame) -> pd.DataFrame:
    today = datetime.date.today()
    df["Kalan Gün"] = df["Bitiş Tarihi"].apply(
        lambda d: (d - today).days if d is not None else None)
    return df

def get_status_label(days) -> str:
    if days is None:       return "⚪ Bilinmiyor"
    if days < 0:           return "🔴 Süresi Dolmuş"
    if days <= 7:          return "🟠 Bu Hafta Bitiyor"
    if days <= 31:         return "🟡 Bu Ay Bitiyor"
    return "🟢 Geçerli"

def filter_by_month(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    fd, ld = get_month_range(year, month)
    mask = df["Bitiş Tarihi"].apply(lambda d: d is not None and fd <= d <= ld)
    return df[mask].copy()

def make_download_excel(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Evraklar")
    return out.getvalue()

def get_unvan(ad):
    return st.session_state.get("unvan_map", {}).get(ad, "— Atanmadı —")

# ══════════════════════════════════════════════════════════════════════════════
# localStorage KÖPRÜSÜ (st.html kullanıldı, date parse uyarısı giderildi)
# ══════════════════════════════════════════════════════════════════════════════
if "unvan_map" not in st.session_state:
    st.session_state["unvan_map"] = {}
if "app_initialized" not in st.session_state:
    st.session_state["app_initialized"] = False

if not st.session_state["app_initialized"]:
    # localStorage'dan URL'ye ilk aktarım
    st.html("""
    <script>
    (function() {
        const url = new URL(window.parent.location.href);
        if (!url.searchParams.has('_u')) {
            const stored = localStorage.getItem('nitelik_unvan_map');
            if (stored) {
                url.searchParams.set('_u', stored);
                window.parent.history.replaceState({}, '', url.toString());
                window.parent.location.reload();
            }
        }
    })();
    </script>
    """)
    st.session_state["app_initialized"] = True
    st.stop()

# URL'den unvan_map yükle
query_data = st.query_params.get("_u", "")
if query_data:
    try:
        data = json.loads(query_data)
        if isinstance(data, dict):
            st.session_state["unvan_map"] = data
    except Exception:
        pass

# Güncel unvan_map'i localStorage'a yaz
unvan_map_json = json.dumps(st.session_state["unvan_map"], ensure_ascii=False)
st.html(f"""
<script>
(function() {{
    try {{
        localStorage.setItem('nitelik_unvan_map', {json.dumps(unvan_map_json)});
    }} catch(e) {{
        console.warn('localStorage yazılamadı:', e);
    }}
}})();
</script>
""")

def reset_localstorage():
    st.html("""
    <script>
    (function() {
        localStorage.removeItem('nitelik_unvan_map');
    })();
    </script>
    """)

# ──────────────────────────────────────────────────────────────────────────────
# GRAFİK FONKSİYONLARI
# ──────────────────────────────────────────────────────────────────────────────
def grafik_durum_pasta(df: pd.DataFrame):
    counts = df["Durum"].value_counts().reset_index()
    counts.columns = ["Durum", "Adet"]
    renkler = [DURUM_RENKLERI.get(d, "#95A5A6") for d in counts["Durum"]]
    fig = go.Figure(go.Pie(
        labels=counts["Durum"], values=counts["Adet"],
        marker_colors=renkler,
        hole=0.55,
        textinfo="percent+value",
        textfont_size=12,
        hovertemplate="<b>%{label}</b><br>%{value} evrak (%{percent})<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="📊 Evrak Durum Dağılımı", font=dict(size=14, color="#F0C040")),
        showlegend=True,
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11),
                    bgcolor="rgba(0,0,0,0)", bordercolor="rgba(0,0,0,0)"),
        **PLOTLY_LAYOUT,
    )
    return fig

def grafik_aylik_bitis(df: pd.DataFrame):
    today = datetime.date.today()
    TR_AYLAR = ["","Oca","Şub","Mar","Nis","May","Haz","Tem","Ağu","Eyl","Eki","Kas","Ara"]
    ay_labels, ay_counts, ay_renkler = [], [], []
    for i in range(12):
        mo = today.month - 1 + i
        y, m = today.year + mo // 12, mo % 12 + 1
        fd, ld = get_month_range(y, m)
        cnt = df["Bitiş Tarihi"].apply(lambda d: d is not None and fd <= d <= ld).sum()
        ay_labels.append(f"{TR_AYLAR[m]}\n{y}")
        ay_counts.append(int(cnt))
        ay_renkler.append("#E74C3C" if i == 0 else "#2980B9" if i <= 2 else "#1A5276")
    fig = go.Figure(go.Bar(
        x=ay_labels, y=ay_counts, marker_color=ay_renkler,
        text=ay_counts, textposition="outside",
        textfont=dict(color="#E8EDF3", size=11),
        hovertemplate="<b>%{x}</b><br>%{y} evrak bitiyor<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="📅 Aylık Bitiş Takvimi (12 Ay)", font=dict(size=14, color="#F0C040")),
        xaxis=dict(gridcolor="#1E3A52", tickfont=dict(size=10)),
        yaxis=dict(gridcolor="#1E3A52", tickfont=dict(size=10)),
        **PLOTLY_LAYOUT,
    )
    return fig

def grafik_unvan_dagilim(df: pd.DataFrame):
    unvan_counts = (df[df["Ünvan"] != "— Atanmadı —"]
                    .groupby("Ünvan").size().reset_index(name="Adet")
                    .sort_values("Adet", ascending=True))
    if unvan_counts.empty:
        return None
    fig = go.Figure(go.Bar(x=unvan_counts["Adet"], y=unvan_counts["Ünvan"],
                           orientation="h", marker_color="#2980B9",
                           text=unvan_counts["Adet"], textposition="outside",
                           textfont=dict(color="#E8EDF3", size=11)))
    fig.update_layout(
        title=dict(text="👤 Ünvana Göre Evrak Dağılımı", font=dict(size=14, color="#F0C040")),
        xaxis=dict(gridcolor="#1E3A52"), yaxis=dict(gridcolor="#1E3A52"),
        **PLOTLY_LAYOUT,
    )
    return fig

def grafik_kritik_personel(df: pd.DataFrame):
    kritik = df[df["Durum"].isin(["🔴 Süresi Dolmuş","🟠 Bu Hafta Bitiyor","🟡 Bu Ay Bitiyor"])]
    if kritik.empty:
        return None
    top = (kritik.groupby("Adı Soyadı").size()
           .reset_index(name="Kritik Evrak")
           .sort_values("Kritik Evrak", ascending=True).tail(10))
    fig = go.Figure(go.Bar(x=top["Kritik Evrak"], y=top["Adı Soyadı"],
                           orientation="h", marker_color="#E74C3C",
                           text=top["Kritik Evrak"], textposition="outside",
                           textfont=dict(color="#E8EDF3", size=11)))
    fig.update_layout(
        title=dict(text="🚨 En Kritik 10 Personel", font=dict(size=14, color="#F0C040")),
        xaxis=dict(gridcolor="#1E3A52"), yaxis=dict(gridcolor="#1E3A52"),
        **PLOTLY_LAYOUT,
    )
    return fig

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div style='font-family:Syne,sans-serif;font-size:1.25rem;font-weight:800;color:#F0C040;padding:8px 0 18px 0;letter-spacing:-0.3px;'>⚓ Belge Takip</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>📂 Veri Kaynağı</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Excel dosyasını buraya sürükleyin", type=["xlsx","xls"])
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>📅 Kontrol Dönemi</div>", unsafe_allow_html=True)
    month_options = get_month_options()
    month_labels  = [x[0] for x in month_options]
    selected_month_label = st.selectbox("Hangi ayı kontrol etmek istiyorsunuz?", month_labels, index=0)
    sel_idx   = month_labels.index(selected_month_label)
    sel_year  = month_options[sel_idx][1]
    sel_month = month_options[sel_idx][2]
    first_day, last_day = get_month_range(sel_year, sel_month)

    today = datetime.date.today()
    st.markdown(f"<div style='font-size:0.82rem;color:#8BAEC8;line-height:1.9;margin-top:10px;'>🗓 Bugün: <b style='color:#E8EDF3;'>{today.strftime('%d.%m.%Y')}</b><br>🔍 Aralık: <b style='color:#F0C040;'>{first_day.strftime('%d.%m')} – {last_day.strftime('%d.%m.%Y')}</b></div>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<div style='font-size:0.72rem;color:#4A7FA5;text-align:center;'>v1.4.3 · Nitelik Belgesi Takip Sistemi</div>", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# HERO BAŞLIK
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(f"""<div class="hero-block"><div class="hero-title">Nitelik Belgesi Takip Sistemi</div><p class="hero-sub">Gemi personeline ait nitelik belgelerinin bitiş tarihlerini izler. Cari ay içinde süresi dolacak kritik evrakları anında tespit eder.<span style="color:#F0C040;font-weight:600;">&nbsp;·&nbsp; {first_day.strftime('%B %Y')} dönemi aktif.</span></p></div>""", unsafe_allow_html=True)

if uploaded_file is None:
    st.markdown("""<div class="alert-info"><b>📂 Başlamak için sol panelden Excel dosyanızı yükleyin.</b><br><br>Dosyanızda şu sütunlar bulunmalıdır:<br>&nbsp;&nbsp;• <b>Adı Soyadı</b> — personelin tam adı<br>&nbsp;&nbsp;• <b>Sicil No</b> — personel sicil numarası (isteğe bağlı)<br>&nbsp;&nbsp;• <b>Evrak Adı</b> — belgenin adı<br>&nbsp;&nbsp;• <b>Başlangıç Tarihi</b> — belgenin başlangıç tarihi (isteğe bağlı)<br>&nbsp;&nbsp;• <b>Bitiş Tarihi</b> — belgenin bitiş tarihi <span style="color:#F0C040;">(zorunlu)</span></div>""", unsafe_allow_html=True)
    sample_df = pd.DataFrame({
        "Adı Soyadı": ["Ahmet Yılmaz","Mehmet Demir","Ali Kaya"],
        "Sicil No": ["001","002","003"],
        "Evrak Adı": ["STCW Güvenlik","Gemiadamları Sağlık Yoklama Belgesi","Vardiya Zabiti"],
        "Başlangıç Tarihi": ["01.07.2023","15.03.2024","10.01.2022"],
        "Bitiş Tarihi": [today.strftime("%d.%m.%Y"), (today+datetime.timedelta(days=10)).strftime("%d.%m.%Y"), (today+datetime.timedelta(days=90)).strftime("%d.%m.%Y")],
    })
    st.download_button("⬇️ Örnek şablonu indir (.xlsx)", make_download_excel(sample_df), "ornek_sablon.xlsx")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# VERİ YÜKLEME
# ──────────────────────────────────────────────────────────────────────────────
df_raw = load_excel(uploaded_file)
if df_raw is None: st.stop()
df = normalize_columns(df_raw)
if df is None: st.stop()
df = parse_dates(df)
df = compute_remaining_days(df)
df["Durum"] = df["Kalan Gün"].apply(get_status_label)

personel_listesi = sorted(df["Adı Soyadı"].dropna().unique().tolist())
df.insert(0, "Ünvan", df["Adı Soyadı"].apply(get_unvan))
df_month   = filter_by_month(df, sel_year, sel_month)
df_display = df_month.copy()

# ──────────────────────────────────────────────────────────────────────────────
# SEKMELER
# ──────────────────────────────────────────────────────────────────────────────
tab_panel, tab_evrak, tab_grafik = st.tabs(["📋 Evrak Takip Paneli", "✍️ Evrak / Tarih Girişi", "📊 Grafikler & Analiz"])

# ╔══════════════════════════════════════════════════════════════════════════════
# SEKME 1 — EVRAK TAKİP PANELİ
# ╚══════════════════════════════════════════════════════════════════════════════
with tab_panel:
    atanmamis_sayisi = sum(1 for p in personel_listesi if st.session_state["unvan_map"].get(p,"— Atanmadı —") == "— Atanmadı —")
    with st.expander(f"👤 Personel Ünvan Ataması  —  {len(personel_listesi)} kişi  {'✅ Tümü atandı' if atanmamis_sayisi==0 else f'⚠️ {atanmamis_sayisi} kişiye henüz ünvan atanmadı'}", expanded=(atanmamis_sayisi>0)):
        st.markdown("<div style='font-size:0.85rem;color:#8BAEC8;margin-bottom:14px;'>Her personel için doğru ünvanı seçin.</div>", unsafe_allow_html=True)
        # toplu seçili atama
        c1,c2,c3 = st.columns([3,2,2])
        with c1: secili_kisiler = st.multiselect("Kişi(ler) seç:", options=personel_listesi, key="ms")
        with c2: secim_unvan = st.selectbox("Ünvan:", ["(Seçin)"]+UNVAN_LISTESI[1:], key="su")
        with c3:
            if st.button("✅ Seçililere Ata"):
                if secim_unvan != "(Seçin)" and secili_kisiler:
                    for p in secili_kisiler: st.session_state["unvan_map"][p] = secim_unvan
                    st.success(f"{len(secili_kisiler)} kişiye {secim_unvan} atandı."); st.rerun()
                else: st.warning("Seçim yapın.")
        st.markdown("---")
        # toplu tüm liste
        ca1,ca2,ca3 = st.columns([2,2,3])
        with ca1: bulk_unvan = st.selectbox("Ünvan:", ["(Seçin)"]+UNVAN_LISTESI[1:], key="bulk_u")
        with ca2: bulk_target = st.selectbox("Kimlere?", ["Atanmamış olanlara","Tüm personele"], key="bulk_t")
        with ca3:
            if st.button("⚡ Toplu Uygula"):
                if bulk_unvan != "(Seçin)":
                    for p in personel_listesi:
                        if bulk_target == "Tüm personele" or st.session_state["unvan_map"].get(p,"— Atanmadı —") == "— Atanmadı —":
                            st.session_state["unvan_map"][p] = bulk_unvan
                    st.rerun()
        st.markdown("---")
        cols_per_row = 3
        for row_start in range(0, len(personel_listesi), cols_per_row):
            row_p = personel_listesi[row_start:row_start+cols_per_row]
            cols  = st.columns(cols_per_row)
            for col, personel in zip(cols, row_p):
                with col:
                    mevcut = st.session_state["unvan_map"].get(personel,"— Atanmadı —")
                    idx = UNVAN_LISTESI.index(mevcut) if mevcut in UNVAN_LISTESI else 0
                    secim = col.selectbox(personel, UNVAN_LISTESI, index=idx, key=f"unvan_{personel}")
                    if secim != st.session_state["unvan_map"].get(personel):
                        st.session_state["unvan_map"][personel] = secim
                        st.rerun()
        if st.button("🗑 Tüm ünvan atamalarını sıfırla"):
            st.session_state["unvan_map"] = {}
            reset_localstorage()
            st.query_params.clear()
            st.rerun()

    df["Ünvan"] = df["Adı Soyadı"].apply(get_unvan)
    df_month["Ünvan"] = df_month["Adı Soyadı"].apply(get_unvan)
    df_display["Ünvan"] = df_display["Adı Soyadı"].apply(get_unvan)

    total_personel = df["Adı Soyadı"].nunique()
    total_evrak = len(df)
    ay_biten = len(df_month)
    dolmus = int(df["Kalan Gün"].apply(lambda x: x is not None and x < 0).sum())
    bu_hafta = int(df["Kalan Gün"].apply(lambda x: x is not None and 0 <= x <= 7).sum())

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: st.metric("👤 Toplam Personel", total_personel)
    with c2: st.metric("📄 Toplam Evrak", total_evrak)
    with c3: st.metric("📅 Bu Ay Bitecek", ay_biten, delta=f"-{ay_biten}" if ay_biten else None, delta_color="inverse")
    with c4: st.metric("🔴 Süresi Dolmuş", dolmus, delta=f"-{dolmus}" if dolmus else None, delta_color="inverse")
    with c5: st.metric("⚡ Bu Hafta Bitiyor", bu_hafta, delta=f"-{bu_hafta}" if bu_hafta else None, delta_color="inverse")

    if dolmus>0: st.markdown(f"<div class='alert-critical'>🚨 {dolmus} evrakın süresi dolmuş!</div>", unsafe_allow_html=True)
    if bu_hafta>0: st.markdown(f"<div class='alert-warning'>⚠️ {bu_hafta} evrak bu hafta bitiyor.</div>", unsafe_allow_html=True)
    if ay_biten>0 and dolmus==0 and bu_hafta==0: st.markdown(f"<div class='alert-warning'>⚠️ Bu ay {ay_biten} evrak bitecek.</div>", unsafe_allow_html=True)
    if ay_biten==0: st.markdown("<div class='alert-ok'>✅ Bu ay süresi dolacak evrak yok.</div>", unsafe_allow_html=True)

    # filtreler
    fc1,fc2,fc3,fc4 = st.columns([3,2,2,2])
    with fc1: search_name = st.text_input("Personel adı ara")
    with fc2:
        evrak_lst = ["(Hepsi)"] + sorted(df_month["Evrak Adı"].dropna().unique().tolist())
        sel_evrak = st.selectbox("Evrak türü", evrak_lst)
    with fc3: sel_durum = st.selectbox("Durum", ["(Hepsi)","🔴 Süresi Dolmuş","🟠 Bu Hafta Bitiyor","🟡 Bu Ay Bitiyor"])
    with fc4: sel_unvan_f = st.selectbox("Ünvana göre", ["(Hepsi)"]+UNVAN_LISTESI[1:])

    df_filtered = df_display.copy()
    if search_name.strip(): df_filtered = df_filtered[df_filtered["Adı Soyadı"].str.contains(search_name.strip(), case=False, na=False)]
    if sel_evrak != "(Hepsi)": df_filtered = df_filtered[df_filtered["Evrak Adı"] == sel_evrak]
    if sel_durum != "(Hepsi)": df_filtered = df_filtered[df_filtered["Durum"] == sel_durum]
    if sel_unvan_f != "(Hepsi)": df_filtered = df_filtered[df_filtered["Ünvan"] == sel_unvan_f]

    st.markdown(f"<div class='section-label'>📋 {first_day.strftime('%B %Y')} — Süresi Dolan / Dolacak Evraklar ({len(df_filtered)} kayıt)</div>", unsafe_allow_html=True)
    if len(df_filtered)==0:
        st.markdown("<div class='alert-info'>Kayıt bulunamadı.</div>", unsafe_allow_html=True)
    else:
        show_cols = [c for c in ["Ünvan","Adı Soyadı","Sicil No","Evrak Adı","Başlangıç Tarihi","Bitiş Tarihi","Kalan Gün","Durum"] if c in df_filtered.columns]
        df_show = df_filtered[show_cols].sort_values("Kalan Gün", ascending=True).copy()
        for dc in ["Bitiş Tarihi","Başlangıç Tarihi"]:
            if dc in df_show.columns: df_show[dc] = df_show[dc].apply(lambda d: d.strftime("%d.%m.%Y") if d else "-")
        df_show["Kalan Gün"] = df_show["Kalan Gün"].apply(lambda x: int(x) if x is not None else "-")
        st.dataframe(df_show, use_container_width=True, hide_index=True, height=min(60+len(df_show)*36, 520))
        st.download_button("⬇️ Excel indir", make_download_excel(df_show), f"kritik_evraklar_{today.strftime('%Y%m%d')}.xlsx")

    with st.expander("📄 Tüm evrak listesi (tüm aylar)"):
        show_all = [c for c in ["Ünvan","Adı Soyadı","Sicil No","Evrak Adı","Başlangıç Tarihi","Bitiş Tarihi","Kalan Gün","Durum"] if c in df.columns]
        df_all = df[show_all].copy()
        for dc in ["Bitiş Tarihi","Başlangıç Tarihi"]:
            if dc in df_all.columns: df_all[dc] = df_all[dc].apply(lambda d: d.strftime("%d.%m.%Y") if d else "-")
        df_all["Kalan Gün"] = df_all["Kalan Gün"].apply(lambda x: int(x) if x is not None else "-")
        st.dataframe(df_all.sort_values("Kalan Gün"), use_container_width=True, hide_index=True, height=400)
        st.download_button("⬇️ Tüm liste Excel", make_download_excel(df_all), f"tum_evraklar_{today.strftime('%Y%m%d')}.xlsx", key="dl_all")

# ╔══════════════════════════════════════════════════════════════════════════════
# SEKME 2 — EVRAK / TARİH GİRİŞİ
# ╚══════════════════════════════════════════════════════════════════════════════
with tab_evrak:
    st.markdown("<div class='section-label'>✍️ Personele Evrak / Başlangıç Tarihi Gir</div>", unsafe_allow_html=True)
    st.markdown("<div class='alert-info'>Sağlık belgesi → 2 yıl, diğer → 5 yıl</div>", unsafe_allow_html=True)
    if "manuel_tarihler" not in st.session_state: st.session_state["manuel_tarihler"] = {}

    col_g1, col_g2 = st.columns([1,1])
    with col_g1:
        st.markdown("<div class='card-box'>", unsafe_allow_html=True)
        secilen_personel = st.selectbox("Personel", ["(Seçin)"]+personel_listesi, key="giris_personel")
        YENI_EVRAK_SECENEGI = "✏️ (Yeni belge adı gir...)"
        evrak_secenekleri = [YENI_EVRAK_SECENEGI] + (sorted(df[df["Adı Soyadı"]==secilen_personel]["Evrak Adı"].dropna().unique()) if secilen_personel!="(Seçin)" else [])
        secilen_evrak_goster = st.selectbox("Evrak", evrak_secenekleri, key="giris_evrak_secenek")
        if secilen_evrak_goster == YENI_EVRAK_SECENEGI:
            yeni_evrak_adi = st.text_input("Yeni belge adını yazın:", placeholder="örn: İleri Yangın Eğitimi", key="yeni_evrak_adi_input").strip()
            secilen_evrak = yeni_evrak_adi if yeni_evrak_adi else None
        else:
            secilen_evrak = secilen_evrak_goster

        if secilen_personel!="(Seçin)" and secilen_evrak:
            key_mt = (secilen_personel, secilen_evrak)
            mevcut_bas = st.session_state["manuel_tarihler"].get(key_mt, {}).get("baslangic")
            if mevcut_bas is None:
                satir = df[(df["Adı Soyadı"]==secilen_personel)&(df["Evrak Adı"]==secilen_evrak)]
                mevcut_bas = satir.iloc[0]["Başlangıç Tarihi"] if (not satir.empty and satir.iloc[0]["Başlangıç Tarihi"] is not None) else today
            yenileme_tarihi = st.date_input("📅 Giriş / Yenileme Tarihi", value=mevcut_bas, format="DD.MM.YYYY", key="giris_tarih")
            sure_yil = 2 if is_saglik_belgesi(secilen_evrak) else 5
            hesaplanan_bitis = hesapla_bitis(yenileme_tarihi, secilen_evrak)
            st.markdown(f"<div style='margin-top:12px;padding:14px;background:#0B2D1A;border:1px solid #2ECC71;border-radius:10px;'>⏳ Geçerlilik süresi: <b style='color:#F0C040;'>{sure_yil} yıl</b><br>📌 Hesaplanan bitiş: <b style='color:#2ECC71;'>{hesaplanan_bitis.strftime('%d.%m.%Y')}</b></div>", unsafe_allow_html=True)
            if st.button("💾 Kaydet", key="giris_kaydet"):
                st.session_state["manuel_tarihler"][key_mt] = {"baslangic": yenileme_tarihi, "bitis": hesaplanan_bitis}
                st.success(f"✅ {secilen_personel} — {secilen_evrak}: {yenileme_tarihi.strftime('%d.%m.%Y')} → {hesaplanan_bitis.strftime('%d.%m.%Y')} kaydedildi.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with col_g2:
        st.markdown("<div class='card-box'>", unsafe_allow_html=True)
        st.markdown("**📋 Manuel Girilen Tarihler**")
        if not st.session_state["manuel_tarihler"]:
            st.markdown("<div style='color:#8BAEC8;text-align:center;'>Henüz manuel tarih girilmedi.</div>", unsafe_allow_html=True)
        else:
            mt_rows = []
            for (ad, evrak), vals in st.session_state["manuel_tarihler"].items():
                sure_yil = 2 if is_saglik_belgesi(evrak) else 5
                kalan = (vals["bitis"] - today).days
                mt_rows.append({"Personel": ad, "Evrak": evrak, "Başlangıç": vals["baslangic"].strftime("%d.%m.%Y"), "Bitiş": vals["bitis"].strftime("%d.%m.%Y"), "Süre": f"{sure_yil} yıl", "Kalan Gün": kalan, "Durum": get_status_label(kalan)})
            mt_df = pd.DataFrame(mt_rows).sort_values("Kalan Gün")
            st.dataframe(mt_df, use_container_width=True, hide_index=True, height=340)
            st.download_button("⬇️ Manuel tarihleri Excel indir", make_download_excel(mt_df), f"manuel_tarihler_{today.strftime('%Y%m%d')}.xlsx", key="dl_manuel")
            if st.button("🗑 Tüm manuel tarihleri sıfırla"): st.session_state["manuel_tarihler"] = {}; st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# ╔══════════════════════════════════════════════════════════════════════════════
# SEKME 3 — GRAFİKLER & ANALİZ
# ╚══════════════════════════════════════════════════════════════════════════════
with tab_grafik:
    st.markdown("<div class='section-label'>📊 Evrak Durum Analizi</div>", unsafe_allow_html=True)
    g1, g2 = st.columns([1, 1.6])
    with g1: st.plotly_chart(grafik_durum_pasta(df), use_container_width=True)
    with g2: st.plotly_chart(grafik_aylik_bitis(df), use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)
    g3, g4 = st.columns(2)
    with g3:
        fig_unvan = grafik_unvan_dagilim(df)
        if fig_unvan: st.plotly_chart(fig_unvan, use_container_width=True)
        else: st.info("Ünvan ataması yapılınca grafik görünür.")
    with g4:
        fig_kritik = grafik_kritik_personel(df)
        if fig_kritik: st.plotly_chart(fig_kritik, use_container_width=True)
        else: st.success("Kritik personel yok.")

    st.markdown("<div class='section-label'>📋 Evrak Türüne Göre Durum Dağılımı</div>", unsafe_allow_html=True)
    evrak_durum = df.groupby(["Evrak Adı","Durum"]).size().reset_index(name="Adet")
    top_evraklar = df["Evrak Adı"].value_counts().head(15).index.tolist()
    evrak_durum_top = evrak_durum[evrak_durum["Evrak Adı"].isin(top_evraklar)]
    if not evrak_durum_top.empty:
        fig_stacked = px.bar(evrak_durum_top, x="Adet", y="Evrak Adı", color="Durum", orientation="h", color_discrete_map=DURUM_RENKLERI, barmode="stack")
        fig_stacked.update_layout(title="En Yaygın 15 Evrak Türü", height=480, **PLOTLY_LAYOUT)
        st.plotly_chart(fig_stacked, use_container_width=True)

    ozet = pd.DataFrame({
        "Metrik": ["Toplam Evrak Kaydı","Benzersiz Personel","Benzersiz Evrak Türü","Süresi Dolmuş","Bu Hafta Bitiyor","Bu Ay Bitiyor","3 Ay İçinde Bitiyor","Geçerli (>90 gün)"],
        "Değer": [len(df), df["Adı Soyadı"].nunique(), df["Evrak Adı"].nunique(),
                  int(df["Kalan Gün"].apply(lambda x: x is not None and x<0).sum()),
                  int(df["Kalan Gün"].apply(lambda x: x is not None and 0<=x<=7).sum()),
                  int(df["Kalan Gün"].apply(lambda x: x is not None and 0<=x<=31).sum()),
                  int(df["Kalan Gün"].apply(lambda x: x is not None and 0<=x<=90).sum()),
                  int(df["Kalan Gün"].apply(lambda x: x is not None and x>90).sum())]
    })
    st.dataframe(ozet, use_container_width=True, hide_index=True)
