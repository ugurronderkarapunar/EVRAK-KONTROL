"""
⚓ Gemi Personeli Nitelik Belgesi Takip Sistemi
================================================
Yazar      : Claude (Anthropic)
Versiyon   : 1.2.0
Açıklama   : Personele ait nitelik belgelerinin bitiş tarihlerini izler,
             cari ay içinde süresi dolacak evrakları dinamik uyarılarla
             ve filtrelenebilir tablolarla kullanıcıya sunar.
Bağımlılıklar: streamlit, pandas, openpyxl, xlrd
"""

import streamlit as st
import pandas as pd
import datetime
import calendar
import io

# ──────────────────────────────────────────────────────────────────────────────
# SAYFA YAPLANDIRMASI (en üste yazılmalı — herhangi bir st. çağrısından önce)
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Belge Takip Sistemi",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# ÖZEL CSS — Denizcilik teması: lacivert / altın / çelik grisi palet
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ---------- Google Fonts ---------- */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap');

/* ---------- Genel arka plan ---------- */
html, body, [data-testid="stAppViewContainer"] {
    background: #0B1929 !important;
    color: #E8EDF3 !important;
    font-family: 'Inter', sans-serif !important;
}

[data-testid="stSidebar"] {
    background: #0D2137 !important;
    border-right: 1px solid #1E3A52 !important;
}

/* ---------- Başlık bloğu ---------- */
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
    right: 32px;
    top: 50%;
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
.hero-sub {
    font-size: 0.95rem;
    color: #8BAEC8;
    margin: 0;
    max-width: 620px;
}

/* ---------- Metrik kartları ---------- */
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
[data-testid="metric-container"] [data-testid="stMetricDelta"] {
    font-size: 0.78rem !important;
}

/* ---------- Uyarı kutusu (custom) ---------- */
.alert-critical {
    background: linear-gradient(90deg, #3D0B0B, #1C0B0B);
    border-left: 5px solid #E74C3C;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 20px;
    color: #FADADD;
    font-size: 0.97rem;
}
.alert-warning {
    background: linear-gradient(90deg, #3D2200, #1C1100);
    border-left: 5px solid #F39C12;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 20px;
    color: #FDEBD0;
    font-size: 0.97rem;
}
.alert-ok {
    background: linear-gradient(90deg, #0B2D1A, #061A0E);
    border-left: 5px solid #2ECC71;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 20px;
    color: #D5F5E3;
    font-size: 0.97rem;
}
.alert-info {
    background: linear-gradient(90deg, #0B1E30, #061018);
    border-left: 5px solid #3498DB;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 20px;
    color: #D6EAF8;
    font-size: 0.97rem;
}

/* ---------- Bölüm başlıkları ---------- */
.section-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #4A7FA5;
    margin-bottom: 12px;
    border-bottom: 1px solid #1E3A52;
    padding-bottom: 6px;
}

/* ---------- Filtre satırı ---------- */
.filter-row {
    background: #112236;
    border: 1px solid #1E3A52;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
}

/* ---------- Sidebar etiketleri ---------- */
[data-testid="stSidebar"] label {
    color: #8BAEC8 !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
}
[data-testid="stSidebar"] .stFileUploader {
    background: #112236 !important;
    border: 1px dashed #2A5280 !important;
    border-radius: 10px !important;
}

/* ---------- Tablo ---------- */
[data-testid="stDataFrame"] {
    border: 1px solid #1E3A52 !important;
    border-radius: 12px !important;
    overflow: hidden !important;
}

/* ---------- İnce çizgi ayırıcı ---------- */
hr { border-color: #1E3A52 !important; }

/* ---------- Streamlit buton ---------- */
.stButton > button {
    background: #1A3A5C;
    color: #F0C040;
    border: 1px solid #2A5280;
    border-radius: 8px;
    font-weight: 600;
    padding: 6px 18px;
}
.stButton > button:hover {
    background: #2A5280;
    color: #FFD966;
    border-color: #4A7FA5;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# YARDIMCI FONKSİYONLAR
# ──────────────────────────────────────────────────────────────────────────────

def get_month_range(year: int, month: int) -> tuple[datetime.date, datetime.date]:
    """Verilen yıl ve ay için ilk ve son günü döndürür."""
    first_day = datetime.date(year, month, 1)
    last_day_num = calendar.monthrange(year, month)[1]
    last_day = datetime.date(year, month, last_day_num)
    return first_day, last_day


def get_month_options() -> list:
    """
    Bugünden başlayarak 12 aylık seçim listesi döndürür.
    Her eleman: (görünen ad, yıl, ay)
    """
    today = datetime.date.today()
    TR_AYLAR = [
        "", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
        "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
    ]
    options = []
    for i in range(12):
        month_offset = today.month - 1 + i
        y = today.year + month_offset // 12
        m = month_offset % 12 + 1
        label = f"{TR_AYLAR[m]} {y}"
        if i == 0:
            label += "  ← Cari Ay"
        options.append((label, y, m))
    return options


def load_excel(file) -> pd.DataFrame | None:
    """
    Yüklenen dosyayı akıllıca okur.
    Gerçek format ne olursa olsun (.xlsx, .xls, TSV/CSV .xls uzantılı) handle eder.

    Strateji:
    1. İlk 8 bayta bakarak gerçek formatı tespit et (magic bytes).
    2. Gerçek Excel (OLE2/OOXML) ise ilgili engine ile oku.
    3. Değilse metin dosyası say; sekme (\\t) veya virgül (,) ayırıcısını dene.
       Türkçe karakterler için windows-1254 → iso-8859-9 → utf-8 sırasıyla dene.
    """
    raw_bytes = file.read()  # Tüm içeriği bir kez oku
    file.seek(0)             # Pointer'ı başa al

    filename = file.name.lower()

    # ── Magic byte tespiti ────────────────────────────────────────────────────
    magic = raw_bytes[:8]
    is_ole2  = magic[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'   # Gerçek .xls (BIFF)
    is_ooxml = magic[:4] == b'PK\x03\x04'                           # .xlsx (ZIP/OOXML)

    # ── Gerçek Excel formatı ──────────────────────────────────────────────────
    if is_ole2:
        try:
            return pd.read_excel(io.BytesIO(raw_bytes), engine="xlrd")
        except Exception as e:
            st.error(f"XLS dosyası okunamadı: {e}")
            return None

    if is_ooxml:
        try:
            return pd.read_excel(io.BytesIO(raw_bytes), engine="openpyxl")
        except Exception as e:
            st.error(f"XLSX dosyası okunamadı: {e}")
            return None

    # ── Metin tabanlı dosya (TSV / CSV, .xls uzantılı olsa bile) ─────────────
    # Türkçe Windows encoding sırası
    ENCODINGS = ["windows-1254", "iso-8859-9", "utf-8-sig", "utf-8", "latin-1"]
    text_content = None
    used_encoding = None

    for enc in ENCODINGS:
        try:
            text_content = raw_bytes.decode(enc)
            used_encoding = enc
            break
        except (UnicodeDecodeError, LookupError):
            continue

    if text_content is None:
        st.error("Dosya kodlaması tanınamadı. Lütfen dosyayı UTF-8 veya Windows-1254 ile kaydedin.")
        return None

    # Ayırıcı tespiti: ilk satırdaki sekme sayısı > virgül sayısı → TSV
    first_line = text_content.split("\n")[0]
    sep = "\t" if first_line.count("\t") >= first_line.count(",") else ","

    try:
        df = pd.read_csv(io.StringIO(text_content), sep=sep, dtype=str)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Metin dosyası okunamadı: {e}")
        return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Sütun adlarını standartlaştırır.
    Boşluk, büyük/küçük harf ve Türkçe karakter farklarını tolere eder.
    """
    # Normalize helper
    def norm(s):
        return (str(s).strip().lower()
                .replace("ı", "i").replace("ğ", "g")
                .replace("ü", "u").replace("ş", "s")
                .replace("ö", "o").replace("ç", "c")
                .replace(" ", ""))

    col_map = {
        # Ad / Adı Soyadı varyantları
        "adisoyadı": "Adı Soyadı",
        "adisoyadi": "Adı Soyadı",
        "adsoyad":   "Adı Soyadı",
        "ad":        "Adı Soyadı",
        "adsoyadi":  "Adı Soyadı",
        "personel":  "Adı Soyadı",
        # Sicil No
        "sicilno":   "Sicil No",
        "sicil":     "Sicil No",
        "sicilnumarasi": "Sicil No",
        # Evrak Adı / Nitelik
        "evrakadi":       "Evrak Adı",
        "evrak":          "Evrak Adı",
        "nitelikbelgesi": "Evrak Adı",
        "nitelik":        "Evrak Adı",
        "belge":          "Evrak Adı",
        "belgeadi":       "Evrak Adı",
        "sertifika":      "Evrak Adı",
        # Başlangıç Tarihi varyantları
        "baslangiçtarihi":  "Başlangıç Tarihi",
        "baslangiç":        "Başlangıç Tarihi",
        "baslangictarihi":  "Başlangıç Tarihi",
        "baslangic":        "Başlangıç Tarihi",
        "baslangictarih":   "Başlangıç Tarihi",
        "baslangic tarihi": "Başlangıç Tarihi",
        # Bitiş Tarihi varyantları
        "bitistarihi":  "Bitiş Tarihi",
        "bitiştarihi":  "Bitiş Tarihi",
        "bitis":        "Bitiş Tarihi",
        "bitiş":        "Bitiş Tarihi",
        "bitistarihi":  "Bitiş Tarihi",
        "bitis tarihi": "Bitiş Tarihi",
        "gecerlilik":   "Bitiş Tarihi",
        "gecerliliktarihi": "Bitiş Tarihi",
        "son gecerlilik":   "Bitiş Tarihi",
        "songecerlilik":    "Bitiş Tarihi",
    }

    renamed = {}
    for col in df.columns:
        key = norm(col)
        if key in col_map:
            renamed[col] = col_map[key]

    df = df.rename(columns=renamed)

    required = ["Adı Soyadı", "Evrak Adı", "Bitiş Tarihi"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(
            f"Excel dosyasında şu zorunlu sütunlar bulunamadı: **{', '.join(missing)}**\n\n"
            "Lütfen sütun adlarını kontrol edin. Beklenen sütunlar: "
            "**Adı Soyadı, Sicil No, Evrak Adı, Başlangıç Tarihi, Bitiş Tarihi**"
        )
        return None

    # Eksik isteğe bağlı sütunları ekle
    for opt in ["Sicil No", "Başlangıç Tarihi"]:
        if opt not in df.columns:
            df[opt] = "-"

    return df


def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    'Bitiş Tarihi' ve 'Başlangıç Tarihi' sütunlarını datetime.date tipine dönüştürür.
    Desteklenen formatlar: d.M.YYYY · dd.MM.YYYY · YYYY-MM-DD · d/M/YYYY
    Hatalı değerler NaT olarak işaretlenir.
    """
    DATE_FORMATS = [
        "%d.%m.%Y",   # 02.07.2026
        "%-d.%-m.%Y", # 2.7.2026  (Unix only)
        "%d/%m/%Y",   # 02/07/2026
        "%Y-%m-%d",   # 2026-07-02
        "%d.%m.%y",   # 02.07.26
    ]

    def parse_one(val):
        if pd.isna(val) or str(val).strip() in ("", "nan", "None", "-"):
            return None
        s = str(val).strip()
        # Önce pandas'ın genel parser'ını dene (dayfirst=True)
        parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
        # Sonra bilinen formatları teker teker dene
        for fmt in DATE_FORMATS:
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except (ValueError, TypeError):
                continue
        return None

    for col in ["Bitiş Tarihi", "Başlangıç Tarihi"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_one)
    return df


def compute_remaining_days(df: pd.DataFrame) -> pd.DataFrame:
    """'Bitiş Tarihi' ile bugün arasındaki gün farkını hesaplar."""
    today = datetime.date.today()
    df["Kalan Gün"] = df["Bitiş Tarihi"].apply(
        lambda d: (d - today).days if pd.notna(d) else None
    )
    return df


def get_status_label(days) -> str:
    """Kalan güne göre Türkçe durum etiketi döndürür."""
    if days is None:
        return "⚪ Bilinmiyor"
    if days < 0:
        return "🔴 Süresi Dolmuş"
    if days <= 7:
        return "🟠 Bu Hafta Bitiyor"
    if days <= 31:
        return "🟡 Bu Ay Bitiyor"
    return "🟢 Geçerli"


def filter_by_month(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    """Belirtilen ay içinde süresi dolan kayıtları filtreler."""
    first_day, last_day = get_month_range(year, month)
    mask = df["Bitiş Tarihi"].apply(
        lambda d: d is not None and pd.notna(d) and first_day <= d <= last_day
    )
    return df[mask].copy()


def make_download_excel(df: pd.DataFrame) -> bytes:
    """DataFrame'i indirilebilir Excel dosyasına dönüştürür."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Kritik Evraklar")
    return output.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(
        "<div style='font-family:Syne,sans-serif;font-size:1.25rem;"
        "font-weight:800;color:#F0C040;padding:8px 0 18px 0;letter-spacing:-0.3px;'>"
        "⚓ Belge Takip</div>",
        unsafe_allow_html=True,
    )

    # — Dosya yükleyici
    st.markdown("<div class='section-label'>📂 Veri Kaynağı</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        label="Excel dosyasını buraya sürükleyin",
        type=["xlsx", "xls"],
        help="Beklenen sütunlar: Adı Soyadı, Sicil No, Evrak Adı, Başlangıç Tarihi, Bitiş Tarihi",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # — Ay seçimi
    st.markdown("<div class='section-label'>📅 Kontrol Dönemi</div>", unsafe_allow_html=True)
    month_options = get_month_options()
    month_labels  = [x[0] for x in month_options]
    selected_month_label = st.selectbox(
        "Hangi ayı kontrol etmek istiyorsunuz?",
        options=month_labels,
        index=0,
        help="Varsayılan: bugünün bulunduğu ay. İlerideki ayları da seçebilirsiniz.",
    )
    # Seçilen etiketin yıl ve ay bilgisini al
    sel_idx   = month_labels.index(selected_month_label)
    sel_year  = month_options[sel_idx][1]
    sel_month = month_options[sel_idx][2]
    first_day, last_day = get_month_range(sel_year, sel_month)

    today = datetime.date.today()
    st.markdown(
        f"<div style='font-size:0.82rem;color:#8BAEC8;line-height:1.9;margin-top:10px;'>"
        f"🗓 Bugün: <b style='color:#E8EDF3;'>{today.strftime('%d.%m.%Y')}</b><br>"
        f"🔍 Aralık: <b style='color:#F0C040;'>{first_day.strftime('%d.%m')} – {last_day.strftime('%d.%m.%Y')}</b>"
        f"</div>",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.markdown(
        "<div style='font-size:0.72rem;color:#4A7FA5;text-align:center;'>"
        "v1.2 · Nitelik Belgesi Takip Sistemi<br>"
        "Streamlit + pandas"
        "</div>",
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────────────────────
# ANA İÇERİK
# ──────────────────────────────────────────────────────────────────────────────

# ── Başlık (Hero) bloğu ──────────────────────────────────────────────────────
st.markdown(
    f"""
    <div class="hero-block">
        <div class="hero-title">Nitelik Belgesi Takip Sistemi</div>
        <p class="hero-sub">
            Gemi personeline ait nitelik belgelerinin bitiş tarihlerini izler.
            Cari ay içinde süresi dolacak kritik evrakları anında tespit eder.
            <span style="color:#F0C040;font-weight:600;">
                &nbsp;·&nbsp; {first_day.strftime('%B %Y')} dönemi aktif.
            </span>
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Excel yüklenmemişse yönlendirme mesajı ───────────────────────────────────
if uploaded_file is None:
    st.markdown(
        """
        <div class="alert-info">
        <b>📂 Başlamak için sol panelden Excel dosyanızı yükleyin.</b><br><br>
        Dosyanızda şu sütunlar bulunmalıdır:<br>
        &nbsp;&nbsp;• <b>Adı Soyadı</b> — personelin tam adı<br>
        &nbsp;&nbsp;• <b>Sicil No</b> — personel sicil numarası (isteğe bağlı)<br>
        &nbsp;&nbsp;• <b>Evrak Adı</b> — belgenin adı (örn. STCW Güvenlik, Sağlık Sertifikası)<br>
        &nbsp;&nbsp;• <b>Başlangıç Tarihi</b> — belgenin başlangıç tarihi (isteğe bağlı)<br>
        &nbsp;&nbsp;• <b>Bitiş Tarihi</b> — belgenin bitiş tarihi <span style="color:#F0C040;">(zorunlu)</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Örnek şablon indirme
    st.markdown("#### 📥 Örnek Excel Şablonu")
    sample_df = pd.DataFrame({
        "Adı Soyadı":      ["Ahmet Yılmaz", "Mehmet Demir", "Ali Kaya"],
        "Sicil No":        ["001", "002", "003"],
        "Evrak Adı":       ["STCW Güvenlik", "Sağlık Sertifikası", "Vardiya Zabiti Belgesi"],
        "Başlangıç Tarihi":["01.07.2023", "15.03.2024", "10.01.2022"],
        "Bitiş Tarihi":    [
            today.strftime("%d.%m.%Y"),
            (today + datetime.timedelta(days=10)).strftime("%d.%m.%Y"),
            (today + datetime.timedelta(days=90)).strftime("%d.%m.%Y"),
        ],
    })
    sample_excel = make_download_excel(sample_df)
    st.download_button(
        label="⬇️ Örnek şablonu indir (.xlsx)",
        data=sample_excel,
        file_name="ornek_sablon.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.stop()  # Dosya yoksa sayfanın geri kalanını render etme


# ──────────────────────────────────────────────────────────────────────────────
# VERİ YÜKLEME VE İŞLEME
# ──────────────────────────────────────────────────────────────────────────────

# Dosyayı oku
df_raw = load_excel(uploaded_file)
if df_raw is None:
    st.stop()

# Sütunları normalize et
df = normalize_columns(df_raw)
if df is None:
    st.stop()

# Tarihleri parse et
df = parse_dates(df)

# Kalan gün hesapla
df = compute_remaining_days(df)

# Durum etiketi ekle
df["Durum"] = df["Kalan Gün"].apply(get_status_label)

# ── Ünvan atamaları (session_state'ten) ──────────────────────────────────────
# Benzersiz personel listesi
personel_listesi = sorted(df["Adı Soyadı"].dropna().unique().tolist())

# session_state başlat
if "unvan_map" not in st.session_state:
    st.session_state["unvan_map"] = {}

# Her personele atanmış ünvanı sütun olarak ekle
def get_unvan(ad):
    return st.session_state["unvan_map"].get(ad, "— Atanmadı —")

df.insert(0, "Ünvan", df["Adı Soyadı"].apply(get_unvan))

# Seçilen ay filtresi
df_month = filter_by_month(df, sel_year, sel_month)
df_display = df_month.copy()


# ──────────────────────────────────────────────────────────────────────────────
# ÜNVAN ATAMA EKRANI
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

with st.expander(
    f"👤 Personel Ünvan Ataması  —  {len(personel_listesi)} kişi  "
    f"({'✅ Tümü atandı' if all(p in st.session_state['unvan_map'] and st.session_state['unvan_map'][p] != '— Atanmadı —' for p in personel_listesi) else f'⚠️ {sum(1 for p in personel_listesi if st.session_state["unvan_map"].get(p,"— Atanmadı —") == "— Atanmadı —")} kişiye henüz ünvan atanmadı'})",
    expanded=any(
        st.session_state["unvan_map"].get(p, "— Atanmadı —") == "— Atanmadı —"
        for p in personel_listesi
    ),
):
    st.markdown(
        "<div style='font-size:0.85rem;color:#8BAEC8;margin-bottom:14px;'>"
        "Her personel için doğru ünvanı seçin. Değişiklikler anında tabloya yansır."
        "</div>",
        unsafe_allow_html=True,
    )

    # Toplu atama kısayolu
    ca1, ca2, ca3 = st.columns([2, 2, 3])
    with ca1:
        bulk_unvan = st.selectbox(
            "Toplu ata:",
            ["(Seçin)"] + UNVAN_LISTESI[1:],
            key="bulk_unvan_sel",
        )
    with ca2:
        # Sadece atanmamışlara veya hepsine uygula
        bulk_target = st.selectbox(
            "Kimlere?",
            ["Atanmamış olanlara", "Tüm personele"],
            key="bulk_target_sel",
        )
    with ca3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⚡ Toplu Uygula", key="bulk_apply"):
            if bulk_unvan != "(Seçin)":
                for p in personel_listesi:
                    mevcut = st.session_state["unvan_map"].get(p, "— Atanmadı —")
                    if bulk_target == "Tüm personele" or mevcut == "— Atanmadı —":
                        st.session_state["unvan_map"][p] = bulk_unvan
                st.rerun()

    st.markdown("---")

    # Kişi başı ünvan atama — 3 sütunlu grid
    cols_per_row = 3
    for row_start in range(0, len(personel_listesi), cols_per_row):
        row_personel = personel_listesi[row_start:row_start + cols_per_row]
        cols = st.columns(cols_per_row)
        for col, personel in zip(cols, row_personel):
            with col:
                mevcut = st.session_state["unvan_map"].get(personel, "— Atanmadı —")
                idx = UNVAN_LISTESI.index(mevcut) if mevcut in UNVAN_LISTESI else 0
                secim = col.selectbox(
                    label=personel,
                    options=UNVAN_LISTESI,
                    index=idx,
                    key=f"unvan_{personel}",
                )
                if secim != st.session_state["unvan_map"].get(personel):
                    st.session_state["unvan_map"][personel] = secim

    st.markdown("---")
    # Sıfırla butonu
    if st.button("🗑 Tüm ünvan atamalarını sıfırla", key="reset_unvanlar"):
        st.session_state["unvan_map"] = {}
        st.rerun()

# Ünvan ataması yapılmış personeli df'e yansıt
df["Ünvan"] = df["Adı Soyadı"].apply(get_unvan)
df_month["Ünvan"] = df_month["Adı Soyadı"].apply(get_unvan)
df_display["Ünvan"] = df_display["Adı Soyadı"].apply(get_unvan)

st.markdown("<br>", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# METRİK KARTLARI
# ──────────────────────────────────────────────────────────────────────────────

total_personel   = df["Adı Soyadı"].nunique()
total_evrak      = len(df)
ay_biten_evrak   = len(df_month)
bugun_biten      = len(df[df["Kalan Gün"] == 0])
dolmus           = len(df[df["Kalan Gün"].apply(lambda x: x is not None and x < 0)])
bu_hafta_biten   = len(df[df["Kalan Gün"].apply(lambda x: x is not None and 0 < x <= 7)])

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("👤 Toplam Personel",   total_personel)
with col2:
    st.metric("📄 Toplam Evrak",      total_evrak)
with col3:
    st.metric("📅 Bu Ay Bitecek",     ay_biten_evrak,
              delta=f"-{ay_biten_evrak}" if ay_biten_evrak > 0 else None,
              delta_color="inverse")
with col4:
    st.metric("🔴 Süresi Dolmuş",     dolmus,
              delta=f"-{dolmus}" if dolmus > 0 else None,
              delta_color="inverse")
with col5:
    st.metric("⚡ Bu Hafta Bitiyor",  bu_hafta_biten,
              delta=f"-{bu_hafta_biten}" if bu_hafta_biten > 0 else None,
              delta_color="inverse")

st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# UYARI KUTULARI
# ──────────────────────────────────────────────────────────────────────────────

if dolmus > 0:
    st.markdown(
        f"<div class='alert-critical'>"
        f"🚨 <b>ACİL:</b> <b>{dolmus}</b> adet evrakın süresi dolmuş! "
        f"Bu evrakların ivedilikle yenilenmesi gerekmektedir."
        f"</div>",
        unsafe_allow_html=True,
    )

if bu_hafta_biten > 0:
    st.markdown(
        f"<div class='alert-warning'>"
        f"⚠️ <b>DİKKAT:</b> <b>{bu_hafta_biten}</b> adet evrak bu hafta sona eriyor. "
        f"Hemen bildirim yapılmalıdır."
        f"</div>",
        unsafe_allow_html=True,
    )

if ay_biten_evrak > 0 and dolmus == 0 and bu_hafta_biten == 0:
    st.markdown(
        f"<div class='alert-warning'>"
        f"⚠️ <b>UYARI:</b> Bu ay toplam <b>{ay_biten_evrak}</b> adet evrakın süresi dolacak. "
        f"İlgili personel bilgilendirilmelidir."
        f"</div>",
        unsafe_allow_html=True,
    )

if ay_biten_evrak == 0:
    st.markdown(
        "<div class='alert-ok'>"
        "✅ <b>Bu ay süresi dolacak evrak bulunmuyor.</b> Tüm belgeler geçerli görünüyor."
        "</div>",
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────────────────────
# FİLTRELEME SATIRLARI
# ──────────────────────────────────────────────────────────────────────────────

st.markdown("<div class='section-label'>🔍 Tablo Filtreleri</div>", unsafe_allow_html=True)

with st.container():
    fc1, fc2, fc3, fc4 = st.columns([3, 2, 2, 2])

    with fc1:
        search_name = st.text_input(
            "Personel adına göre ara",
            placeholder="örn. Ahmet",
        )
    with fc2:
        evrak_listesi = ["(Hepsi)"] + sorted(df_month["Evrak Adı"].dropna().unique().tolist())
        selected_evrak = st.selectbox("Evrak türü", evrak_listesi)
    with fc3:
        durum_listesi = ["(Hepsi)", "🔴 Süresi Dolmuş", "🟠 Bu Hafta Bitiyor", "🟡 Bu Ay Bitiyor"]
        selected_durum = st.selectbox("Durum", durum_listesi)
    with fc4:
        unvan_filtre_listesi = ["(Hepsi)"] + UNVAN_LISTESI[1:]
        selected_unvan_filtre = st.selectbox("Ünvana göre", unvan_filtre_listesi)

# Filtreleri uygula
df_filtered = df_display.copy()

if search_name.strip():
    df_filtered = df_filtered[
        df_filtered["Adı Soyadı"].str.contains(search_name.strip(), case=False, na=False)
    ]
if selected_evrak != "(Hepsi)":
    df_filtered = df_filtered[df_filtered["Evrak Adı"] == selected_evrak]
if selected_durum != "(Hepsi)":
    df_filtered = df_filtered[df_filtered["Durum"] == selected_durum]
if selected_unvan_filtre != "(Hepsi)":
    df_filtered = df_filtered[df_filtered["Ünvan"] == selected_unvan_filtre]


# ──────────────────────────────────────────────────────────────────────────────
# TABLO — Cari ay kritik evrakları
# ──────────────────────────────────────────────────────────────────────────────

st.markdown(
    f"<div class='section-label'>"
    f"📋 {first_day.strftime('%B %Y')} — Süresi Dolan / Dolacak Evraklar "
    f"({len(df_filtered)} kayıt)"
    f"</div>",
    unsafe_allow_html=True,
)

if len(df_filtered) == 0:
    st.markdown(
        "<div class='alert-info'>"
        "Seçili filtrelere uygun kayıt bulunamadı. Filtreleri değiştirmeyi deneyin."
        "</div>",
        unsafe_allow_html=True,
    )
else:
    # Gösterilecek sütunları seç ve sırala
    show_cols = ["Ünvan", "Adı Soyadı", "Sicil No", "Evrak Adı",
                 "Başlangıç Tarihi", "Bitiş Tarihi", "Kalan Gün", "Durum"]
    show_cols = [c for c in show_cols if c in df_filtered.columns]

    df_show = df_filtered[show_cols].sort_values("Kalan Gün", ascending=True)

    # Tarihleri okunabilir formata çevir
    for date_col in ["Bitiş Tarihi", "Başlangıç Tarihi"]:
        if date_col in df_show.columns:
            df_show[date_col] = df_show[date_col].apply(
                lambda d: d.strftime("%d.%m.%Y") if pd.notna(d) and d else "-"
            )

    # Kalan Gün — None değerlerini göster
    df_show["Kalan Gün"] = df_show["Kalan Gün"].apply(
        lambda x: int(x) if pd.notna(x) and x is not None else "-"
    )

    st.dataframe(
        df_show,
        use_container_width=True,
        hide_index=True,
        height=min(60 + len(df_show) * 36, 520),
        column_config={
            "Ünvan":           st.column_config.TextColumn("Ünvan", width="medium"),
            "Adı Soyadı":      st.column_config.TextColumn("Adı Soyadı", width="medium"),
            "Sicil No":        st.column_config.TextColumn("Sicil No", width="small"),
            "Evrak Adı":       st.column_config.TextColumn("Evrak Adı", width="large"),
            "Başlangıç Tarihi":st.column_config.TextColumn("Başlangıç", width="small"),
            "Bitiş Tarihi":    st.column_config.TextColumn("Bitiş Tarihi", width="small"),
            "Kalan Gün":       st.column_config.NumberColumn("Kalan Gün", format="%d gün", width="small"),
            "Durum":           st.column_config.TextColumn("Durum", width="medium"),
        },
    )

    # İndir butonu
    excel_bytes = make_download_excel(df_show)
    st.download_button(
        label="⬇️ Filtrelenmiş listeyi Excel olarak indir",
        data=excel_bytes,
        file_name=f"kritik_evraklar_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ──────────────────────────────────────────────────────────────────────────────
# TAM LİSTE (isteğe bağlı genişlet)
# ──────────────────────────────────────────────────────────────────────────────

with st.expander("📄 Tüm evrak listesini göster (tüm aylar)"):
    show_all_cols = ["Ünvan", "Adı Soyadı", "Sicil No", "Evrak Adı",
                     "Başlangıç Tarihi", "Bitiş Tarihi", "Kalan Gün", "Durum"]
    show_all_cols = [c for c in show_all_cols if c in df.columns]

    df_all = df[show_all_cols].copy()
    for date_col in ["Bitiş Tarihi", "Başlangıç Tarihi"]:
        if date_col in df_all.columns:
            df_all[date_col] = df_all[date_col].apply(
                lambda d: d.strftime("%d.%m.%Y") if pd.notna(d) and d else "-"
            )
    df_all["Kalan Gün"] = df_all["Kalan Gün"].apply(
        lambda x: int(x) if pd.notna(x) and x is not None else "-"
    )
    df_all_sorted = df_all.sort_values("Kalan Gün", ascending=True)

    st.dataframe(
        df_all_sorted,
        use_container_width=True,
        hide_index=True,
        height=400,
    )

    all_excel = make_download_excel(df_all_sorted)
    st.download_button(
        label="⬇️ Tüm listeyi Excel olarak indir",
        data=all_excel,
        file_name=f"tum_evraklar_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_all",
    )
