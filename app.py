"""
⚓ Gemi Personeli Nitelik Belgesi Takip Sistemi
================================================
Versiyon   : 1.6.0
Açıklama   : v1.6.0: Sunucu tarafı JSON state, otomatik Excel kaydetme,
             yedekleme, undo, personel özeti, e-posta uyarısı,
             toplu ekleme, tarih formatı, sicil no arama.
"""

import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import os
import json
import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import plotly.express as px
import plotly.graph_objects as go

# ──────────────────────────────────────────────────────────────────────────────
# SABİTLER
# ──────────────────────────────────────────────────────────────────────────────
SAVED_EXCEL_PATH = "saved_upload.xlsx"
STATE_FILE       = "state.json"
BACKUP_DIR       = "backups"
MAX_BACKUPS      = 10

UNVAN_LISTESI = [
    "— Atanmadı —", "Kaptan", "Başmakinist", "Güverte L.",
    "Gemici", "Yağcı", "Diğer"
]

SAGLIK_ANAHTAR_KELIMELER = ["sağlık", "saglik", "health", "yoklama"]

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(17,34,54,0.0)",
    plot_bgcolor="rgba(17,34,54,0.0)",
    font=dict(color="#E8EDF3", family="Inter"),
    margin=dict(l=16, r=16, t=36, b=16),
)

DURUM_RENKLERI = {
    "🔴 Süresi Dolmuş": "#E74C3C",
    "🟠 Bu Hafta Bitiyor": "#E67E22",
    "🟡 Bu Ay Bitiyor": "#F1C40F",
    "🟢 Geçerli": "#2ECC71",
    "⚪ Bilinmiyor": "#95A5A6",
}

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
# CSS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap');
html, body, [data-testid="stAppViewContainer"] {
    background: #0B1929 !important; color: #E8EDF3 !important;
    font-family: 'Inter', sans-serif !important;
}
[data-testid="stSidebar"] { background: #0D2137 !important; border-right: 1px solid #1E3A52 !important; }
.hero-block {
    background: linear-gradient(135deg, #0D2137 0%, #1A3A5C 60%, #0D2137 100%);
    border: 1px solid #2A5280; border-radius: 16px;
    padding: 36px 40px 28px 40px; margin-bottom: 28px; position: relative; overflow: hidden;
}
.hero-block::before { content: "⚓"; position: absolute; right: 32px; top: 50%; transform: translateY(-50%); font-size: 96px; opacity: 0.06; pointer-events: none; }
.hero-title { font-family: 'Syne', sans-serif; font-size: 2.1rem; font-weight: 800; color: #F0C040; letter-spacing: -0.5px; margin: 0 0 6px 0; line-height: 1.15; }
.hero-sub { font-size: 0.95rem; color: #8BAEC8; margin: 0; max-width: 620px; }
[data-testid="metric-container"] { background: #112236 !important; border: 1px solid #1E3A52 !important; border-radius: 12px !important; padding: 18px 22px !important; }
[data-testid="metric-container"] label { color: #8BAEC8 !important; font-size: 0.78rem !important; letter-spacing: 0.04em !important; text-transform: uppercase !important; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #F0C040 !important; font-family: 'Syne', sans-serif !important; font-size: 2rem !important; }
.alert-critical { background: linear-gradient(90deg, #3D0B0B, #1C0B0B); border-left: 5px solid #E74C3C; border-radius: 10px; padding: 18px 22px; margin-bottom: 20px; color: #FADADD; font-size: 0.97rem; }
.alert-warning { background: linear-gradient(90deg, #3D2200, #1C1100); border-left: 5px solid #F39C12; border-radius: 10px; padding: 18px 22px; margin-bottom: 20px; color: #FDEBD0; font-size: 0.97rem; }
.alert-ok { background: linear-gradient(90deg, #0B2D1A, #061A0E); border-left: 5px solid #2ECC71; border-radius: 10px; padding: 18px 22px; margin-bottom: 20px; color: #D5F5E3; font-size: 0.97rem; }
.alert-info { background: linear-gradient(90deg, #0B1E30, #061018); border-left: 5px solid #3498DB; border-radius: 10px; padding: 18px 22px; margin-bottom: 20px; color: #D6EAF8; font-size: 0.97rem; }
.section-label { font-size: 0.72rem; font-weight: 600; letter-spacing: 0.12em; text-transform: uppercase; color: #4A7FA5; margin-bottom: 12px; border-bottom: 1px solid #1E3A52; padding-bottom: 6px; }
.card-box { background: #112236; border: 1px solid #1E3A52; border-radius: 12px; padding: 20px 24px; margin-bottom: 16px; }
[data-testid="stSidebar"] label { color: #8BAEC8 !important; font-size: 0.82rem !important; font-weight: 500 !important; }
[data-testid="stDataFrame"] { border: 1px solid #1E3A52 !important; border-radius: 12px !important; overflow: hidden !important; }
hr { border-color: #1E3A52 !important; }
.stButton > button { background: #1A3A5C; color: #F0C040; border: 1px solid #2A5280; border-radius: 8px; font-weight: 600; padding: 6px 18px; }
.stButton > button:hover { background: #2A5280; color: #FFD966; border-color: #4A7FA5; }
.js-plotly-plot { border-radius: 12px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# YARDIMCI FONKSİYONLAR (tarih, excel, state, yedek, e-posta ...)
# ──────────────────────────────────────────────────────────────────────────────

def get_month_range(year, month):
    first_day = datetime.date(year, month, 1)
    last_day = datetime.date(year, month, calendar.monthrange(year, month)[1])
    return first_day, last_day

def get_month_options():
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

def is_saglik_belgesi(evrak_adi):
    ad = str(evrak_adi).lower()
    return any(k in ad for k in SAGLIK_ANAHTAR_KELIMELER)

def hesapla_bitis(baslangic, evrak_adi):
    yil = 2 if is_saglik_belgesi(evrak_adi) else 5
    try:
        return baslangic.replace(year=baslangic.year + yil)
    except ValueError:
        return baslangic.replace(year=baslangic.year + yil, day=28)

def load_excel(file):
    raw = file.read()
    file.seek(0)
    magic = raw[:8]
    if magic == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        try: return pd.read_excel(io.BytesIO(raw), engine="xlrd")
        except: st.error("XLS okunamadı"); return None
    if raw[:4] == b'PK\x03\x04':
        try: return pd.read_excel(io.BytesIO(raw), engine="openpyxl")
        except: st.error("XLSX okunamadı"); return None
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
    except: st.error("Metin dosyası okunamadı"); return None

def normalize_columns(df):
    def norm(s):
        return (str(s).strip().lower()
                .replace("ı","i").replace("ğ","g").replace("ü","u")
                .replace("ş","s").replace("ö","o").replace("ç","c")
                .replace(" ",""))
    col_map = {
        "ad":"Adı Soyadı","adsoyadi":"Adı Soyadı","personel":"Adı Soyadı",
        "sicilno":"Sicil No","sicil":"Sicil No",
        "evrakadi":"Evrak Adı","evrak":"Evrak Adı","nitelik":"Evrak Adı",
        "belge":"Evrak Adı","belgeadi":"Evrak Adı",
        "baslangiçtarihi":"Başlangıç Tarihi","baslangic":"Başlangıç Tarihi",
        "baslangictarihi":"Başlangıç Tarihi",
        "bitistarihi":"Bitiş Tarihi","bitis":"Bitiş Tarihi",
        "bitiştarihi":"Bitiş Tarihi","gecerlilik":"Bitiş Tarihi",
    }
    renamed = {c: col_map[norm(c)] for c in df.columns if norm(c) in col_map}
    df = df.rename(columns=renamed)
    missing = [c for c in ["Adı Soyadı","Evrak Adı","Bitiş Tarihi"] if c not in df.columns]
    if missing:
        st.error(f"Zorunlu sütunlar bulunamadı: {', '.join(missing)}")
        return None
    for opt in ["Sicil No","Başlangıç Tarihi"]:
        if opt not in df.columns: df[opt] = "-"
    return df

def parse_dates(df):
    def parse_one(val):
        if pd.isna(val) or str(val).strip() in ("","nan","None","-"): return None
        s = str(val).strip()
        p = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(p): return p.date()
        for fmt in ["%d.%m.%Y","%d/%m/%Y","%Y-%m-%d","%d.%m.%y"]:
            try: return datetime.datetime.strptime(s, fmt).date()
            except: pass
        return None
    for col in ["Bitiş Tarihi","Başlangıç Tarihi"]:
        if col in df.columns: df[col] = df[col].apply(parse_one)
    return df

def compute_remaining_days(df):
    today = datetime.date.today()
    df["Kalan Gün"] = df["Bitiş Tarihi"].apply(lambda d: (d - today).days if d is not None else None)
    return df

def get_status_label(days):
    if days is None: return "⚪ Bilinmiyor"
    if days < 0: return "🔴 Süresi Dolmuş"
    if days <= 7: return "🟠 Bu Hafta Bitiyor"
    if days <= 31: return "🟡 Bu Ay Bitiyor"
    return "🟢 Geçerli"

def filter_by_month(df, year, month):
    fd, ld = get_month_range(year, month)
    mask = df["Bitiş Tarihi"].apply(lambda d: d is not None and fd <= d <= ld)
    return df[mask].copy()

def make_download_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Evraklar")
    return out.getvalue()

def format_date(d, fmt):
    if d is None: return "-"
    return d.strftime(fmt)

def backup_excel():
    if not os.path.exists(SAVED_EXCEL_PATH): return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"backup_{timestamp}.xlsx"
    shutil.copy2(SAVED_EXCEL_PATH, os.path.join(BACKUP_DIR, backup_name))
    # Eski yedekleri temizle
    backups = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("backup_")])
    while len(backups) > MAX_BACKUPS:
        os.remove(os.path.join(BACKUP_DIR, backups[0]))
        backups.pop(0)

def save_state_and_excel():
    """State'i JSON'a yaz, sonra Excel'i güncelle."""
    state = {
        "unvan_map": st.session_state["unvan_map"],
        "manuel_tarihler": {f"{k[0]}|{k[1]}": {"baslangic": v["baslangic"].strftime("%Y-%m-%d"),
                                                "bitis": v["bitis"].strftime("%Y-%m-%d")}
                            for k, v in st.session_state["manuel_tarihler"].items()},
        "evrak_duzenlemeleri": st.session_state["evrak_duzenlemeleri"],
        "silinmis_evraklar": st.session_state["silinmis_evraklar"],
    }
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    # Yedek al ve yeni Excel'i yaz
    backup_excel()
    df_final = build_final_df(df_original.copy())  # df_original global
    with pd.ExcelWriter(SAVED_EXCEL_PATH, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Evraklar")

def push_undo():
    """Son işlem için state yedekle."""
    st.session_state["undo_state"] = {
        "unvan_map": dict(st.session_state["unvan_map"]),
        "manuel_tarihler": dict(st.session_state["manuel_tarihler"]),
        "evrak_duzenlemeleri": dict(st.session_state["evrak_duzenlemeleri"]),
        "silinmis_evraklar": list(st.session_state["silinmis_evraklar"]),
    }

def undo_last():
    if st.session_state.get("undo_state"):
        st.session_state["unvan_map"] = st.session_state["undo_state"]["unvan_map"]
        st.session_state["manuel_tarihler"] = st.session_state["undo_state"]["manuel_tarihler"]
        st.session_state["evrak_duzenlemeleri"] = st.session_state["undo_state"]["evrak_duzenlemeleri"]
        st.session_state["silinmis_evraklar"] = st.session_state["undo_state"]["silinmis_evraklar"]
        del st.session_state["undo_state"]
        save_state_and_excel()
        st.rerun()

def send_email(smtp_config, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_config['username']
        msg['To'] = smtp_config['to_email']
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP(smtp_config['host'], smtp_config['port'])
        server.starttls()
        server.login(smtp_config['username'], smtp_config['password'])
        server.sendmail(smtp_config['username'], smtp_config['to_email'], msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"E‑posta gönderilemedi: {e}")
        return False

# ──────────────────────────────────────────────────────────────────────────────
# STATE YÜKLEME
# ──────────────────────────────────────────────────────────────────────────────
def load_state():
    default = {
        "unvan_map": {},
        "manuel_tarihler": {},
        "evrak_duzenlemeleri": {},
        "silinmis_evraklar": [],
    }
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Tarihleri datetime.date'e çevir
        manuel = {}
        for k, v in data.get("manuel_tarihler", {}).items():
            ad, evrak = k.split("|", 1)
            manuel[(ad, evrak)] = {
                "baslangic": datetime.datetime.strptime(v["baslangic"], "%Y-%m-%d").date(),
                "bitis": datetime.datetime.strptime(v["bitis"], "%Y-%m-%d").date(),
            }
        data["manuel_tarihler"] = manuel
        return data
    return default

state = load_state()
if "unvan_map" not in st.session_state:
    st.session_state["unvan_map"] = state["unvan_map"]
if "manuel_tarihler" not in st.session_state:
    st.session_state["manuel_tarihler"] = state["manuel_tarihler"]
if "evrak_duzenlemeleri" not in st.session_state:
    st.session_state["evrak_duzenlemeleri"] = state["evrak_duzenlemeleri"]
if "silinmis_evraklar" not in st.session_state:
    st.session_state["silinmis_evraklar"] = state["silinmis_evraklar"]
if "undo_state" not in st.session_state:
    st.session_state["undo_state"] = None

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR (Excel yükleme, ayarlar, SMTP, format)
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div style='font-family:Syne;font-size:1.25rem;color:#F0C040;padding:8px 0;'>⚓ Belge Takip</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>📂 Veri Kaynağı</div>", unsafe_allow_html=True)

    saved_exists = os.path.exists(SAVED_EXCEL_PATH)
    if saved_exists:
        st.success(f"✅ Kayıtlı Excel bulundu ({SAVED_EXCEL_PATH})")
    else:
        st.info("Henüz kayıtlı Excel yok, lütfen yükleyin.")

    uploaded_file = st.file_uploader("Excel yükle (yeni dosya seçerseniz eski yerine geçer)", type=["xlsx","xls"])
    if uploaded_file is not None:
        with open(SAVED_EXCEL_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("Yeni Excel kaydedildi.")
        saved_exists = True

    if saved_exists:
        if st.button("🗑 Kayıtlı Excel'i sil ve sıfırla"):
            if os.path.exists(SAVED_EXCEL_PATH):
                os.remove(SAVED_EXCEL_PATH)
            if os.path.exists(STATE_FILE):
                os.remove(STATE_FILE)
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>📅 Kontrol Dönemi</div>", unsafe_allow_html=True)
    month_options = get_month_options()
    month_labels = [x[0] for x in month_options]
    selected_month_label = st.selectbox("Ay seçin", month_labels, index=0)
    sel_idx = month_labels.index(selected_month_label)
    sel_year, sel_month = month_options[sel_idx][1], month_options[sel_idx][2]
    first_day, last_day = get_month_range(sel_year, sel_month)

    today = datetime.date.today()
    st.markdown(f"<small>Bugün: {today.strftime('%d.%m.%Y')} | Aralık: {first_day.strftime('%d.%m')} – {last_day.strftime('%d.%m.%Y')}</small>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div class='section-label'>⚙️ Ayarlar</div>", unsafe_allow_html=True)
    date_format = st.selectbox("Tarih formatı", ["GG.AA.YYYY", "AA/GG/YYYY"], index=0)
    if date_format == "GG.AA.YYYY":
        st.session_state["date_fmt"] = "%d.%m.%Y"
    else:
        st.session_state["date_fmt"] = "%m/%d/%Y"

    st.markdown("---")
    st.markdown("<div class='section-label'>📧 E‑Posta Uyarıları</div>", unsafe_allow_html=True)
    with st.expander("SMTP Ayarları"):
        smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com")
        smtp_port = st.number_input("Port", value=587)
        smtp_user = st.text_input("Kullanıcı adı (e‑posta)")
        smtp_pass = st.text_input("Şifre", type="password")
        smtp_to   = st.text_input("Alıcı e‑posta")
        smtp_ok = all([smtp_host, smtp_user, smtp_pass, smtp_to])
    if st.button("📧 Kritik Evrak Uyarısı Gönder"):
        if not smtp_ok:
            st.error("SMTP bilgilerini doldurun.")
        else:
            with st.spinner("E‑posta gönderiliyor..."):
                # df_final global olacak, henüz oluşmadıysa diye kontrol
                if "df_final" not in st.session_state:
                    st.warning("Önce Excel yükleyin.")
                else:
                    df = st.session_state["df_final"]
                    kritik = df[df["Durum"].isin(["🔴 Süresi Dolmuş", "🟠 Bu Hafta Bitiyor"])]
                    if kritik.empty:
                        st.info("Kritik evrak yok.")
                    else:
                        body = "Aşağıdaki evraklar kritik durumda:\n\n"
                        for _, row in kritik.iterrows():
                            body += f"{row['Adı Soyadı']} - {row['Evrak Adı']} - {row['Bitiş Tarihi'].strftime(st.session_state['date_fmt'])} ({row['Durum']})\n"
                        config = {"host":smtp_host, "port":smtp_port, "username":smtp_user, "password":smtp_pass, "to_email":smtp_to}
                        if send_email(config, "Kritik Evrak Uyarısı", body):
                            st.success("E‑posta gönderildi.")
                        else:
                            st.error("E‑posta gönderilemedi.")

    st.markdown("---")
    if st.button("↩️ Son İşlemi Geri Al"):
        undo_last()
    st.markdown("v1.6.0 · Tüm haklar açık.")

# ──────────────────────────────────────────────────────────────────────────────
# VERİ YÜKLEME VE HAZIRLIK
# ──────────────────────────────────────────────────────────────────────────────
if not os.path.exists(SAVED_EXCEL_PATH):
    st.markdown("<div class='alert-info'>📂 Lütfen sol panelden bir Excel dosyası yükleyin.</div>", unsafe_allow_html=True)
    st.stop()

df_raw = load_excel(open(SAVED_EXCEL_PATH, "rb"))
if df_raw is None: st.stop()
df = normalize_columns(df_raw)
if df is None: st.stop()
df = parse_dates(df)
df = compute_remaining_days(df)
df["Durum"] = df["Kalan Gün"].apply(get_status_label)
df["Kaynak"] = "Excel"
df["_key"] = df["Adı Soyadı"] + "|" + df["Evrak Adı"] + "_" + df.index.astype(str)

def get_unvan(ad):
    return st.session_state["unvan_map"].get(ad, "— Atanmadı —")
df["Ünvan"] = df["Adı Soyadı"].apply(get_unvan)

df_original = df.copy()

def build_final_df(base_df):
    df_edit = base_df.copy()
    silinecekler = set(st.session_state["silinmis_evraklar"])
    if silinecekler:
        df_edit = df_edit[~df_edit["_key"].isin(silinecekler)]
    for key, vals in st.session_state["evrak_duzenlemeleri"].items():
        if key not in df_edit["_key"].values: continue
        mask = df_edit["_key"] == key
        if "baslangic" in vals and vals["baslangic"]:
            try: df_edit.loc[mask, "Başlangıç Tarihi"] = datetime.datetime.strptime(vals["baslangic"], "%Y-%m-%d").date()
            except: pass
        if "bitis" in vals and vals["bitis"]:
            try: df_edit.loc[mask, "Bitiş Tarihi"] = datetime.datetime.strptime(vals["bitis"], "%Y-%m-%d").date()
            except: pass
        df_edit = compute_remaining_days(df_edit)
        df_edit["Durum"] = df_edit["Kalan Gün"].apply(get_status_label)
    manuel_rows = []
    for i, ((ad, evrak), vals) in enumerate(st.session_state["manuel_tarihler"].items()):
        man_key = f"{ad}|{evrak}_manuel_{i}"
        if man_key in silinecekler: continue
        bas, bit = vals["baslangic"], vals["bitis"]
        manuel_rows.append({
            "Ünvan": get_unvan(ad), "Adı Soyadı": ad, "Sicil No": "-",
            "Evrak Adı": evrak, "Başlangıç Tarihi": bas, "Bitiş Tarihi": bit,
            "Kalan Gün": (bit - today).days, "Durum": get_status_label((bit - today).days),
            "Kaynak": "Manuel", "_key": man_key,
        })
    if manuel_rows:
        df_edit = pd.concat([df_edit, pd.DataFrame(manuel_rows)], ignore_index=True)
    return df_edit

df_final = build_final_df(df_original)
st.session_state["df_final"] = df_final
personel_listesi = sorted(df_final["Adı Soyadı"].dropna().unique().tolist())
df_month = filter_by_month(df_final, sel_year, sel_month)

# ──────────────────────────────────────────────────────────────────────────────
# SEKMELER
# ──────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Evrak Takip Paneli", "✍️ Evrak / Tarih Girişi",
    "✏️ Kişi Evrak Düzenle", "👤 Personel Özeti", "📊 Grafikler & Analiz"
])

# SEKME 1 – Evrak Takip Paneli (ünvan atama, metrikler, tablo)
with tab1:
    # Ünvan atama (öncekiyle aynı mantık, push_undo eklendi)
    atanmamis = sum(1 for p in personel_listesi if st.session_state["unvan_map"].get(p,"— Atanmadı —") == "— Atanmadı —")
    with st.expander(f"👤 Personel Ünvan Ataması ({len(personel_listesi)} kişi)", expanded=(atanmamis>0)):
        # ... (içerik aynı, state değişikliklerinde push_undo() ve save_state_and_excel() çağrılı)
        st.markdown("**🎯 Seçili Kişilere Ünvan Ata**")
        sa1, sa2, sa3 = st.columns([3,2,2])
        with sa1:
            sec_kisiler = st.multiselect("Kişi(ler)", personel_listesi, key="ms1")
        with sa2:
            sec_unvan = st.selectbox("Ünvan", ["(Seçin)"]+UNVAN_LISTESI[1:], key="su1")
        with sa3:
            if st.button("✅ Ata", key="ata1"):
                if sec_unvan != "(Seçin)" and sec_kisiler:
                    push_undo()
                    for p in sec_kisiler:
                        st.session_state["unvan_map"][p] = sec_unvan
                    save_state_and_excel()
                    st.rerun()
        st.markdown("---")
        st.markdown("**⚡ Toplu Atama**")
        ca1, ca2, ca3 = st.columns([2,2,3])
        with ca1:
            bulk_unvan = st.selectbox("Ünvan", ["(Seçin)"]+UNVAN_LISTESI[1:], key="bu1")
        with ca2:
            bulk_target = st.selectbox("Kimlere?", ["Atanmamış olanlara", "Tüm personele"], key="bt1")
        with ca3:
            if st.button("⚡ Uygula", key="bua1"):
                if bulk_unvan != "(Seçin)":
                    push_undo()
                    for p in personel_listesi:
                        if bulk_target == "Tüm personele" or st.session_state["unvan_map"].get(p,"— Atanmadı —") == "— Atanmadı —":
                            st.session_state["unvan_map"][p] = bulk_unvan
                    save_state_and_excel()
                    st.rerun()
        st.markdown("---")
        # Bireysel seçimler
        cols_per_row = 3
        for row_start in range(0, len(personel_listesi), cols_per_row):
            row_p = personel_listesi[row_start:row_start+cols_per_row]
            cols = st.columns(cols_per_row)
            for col, personel in zip(cols, row_p):
                with col:
                    mevcut = st.session_state["unvan_map"].get(personel,"— Atanmadı —")
                    idx = UNVAN_LISTESI.index(mevcut)
                    secim = col.selectbox(personel, UNVAN_LISTESI, index=idx, key=f"unvan_{personel}")
                    if secim != mevcut:
                        push_undo()
                        st.session_state["unvan_map"][personel] = secim
                        save_state_and_excel()
        if st.button("🗑 Tüm ünvanları sıfırla"):
            push_undo()
            st.session_state["unvan_map"] = {}
            save_state_and_excel()
            st.rerun()

    # Metrikler
    total_p = df_final["Adı Soyadı"].nunique()
    total_e = len(df_final)
    ay_biten = len(df_month)
    dolmus = int(df_final["Kalan Gün"].apply(lambda x: x is not None and x < 0).sum())
    bu_hafta = int(df_final["Kalan Gün"].apply(lambda x: x is not None and 0 <= x <= 7).sum())
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Personel", total_p)
    c2.metric("Evrak", total_e)
    c3.metric("Bu Ay Bitecek", ay_biten)
    c4.metric("Süresi Dolmuş", dolmus)
    c5.metric("Bu Hafta Bitiyor", bu_hafta)

    # Uyarılar
    if dolmus: st.markdown(f"<div class='alert-critical'>🚨 {dolmus} evrak süresi dolmuş!</div>", unsafe_allow_html=True)
    if bu_hafta: st.markdown(f"<div class='alert-warning'>⚠️ {bu_hafta} evrak bu hafta bitiyor.</div>", unsafe_allow_html=True)

    # Filtreler (sicil no eklendi)
    st.markdown("<div class='section-label'>🔍 Filtreler</div>", unsafe_allow_html=True)
    fc1,fc2,fc3,fc4,fc5 = st.columns([2,2,2,2,2])
    with fc1: search_name = st.text_input("Ad ara")
    with fc2: search_sicil = st.text_input("Sicil No ara")
    with fc3:
        evrak_lst = ["(Hepsi)"] + sorted(df_month["Evrak Adı"].dropna().unique().tolist())
        sel_evrak = st.selectbox("Evrak türü", evrak_lst)
    with fc4:
        sel_durum = st.selectbox("Durum", ["(Hepsi)","🔴 Süresi Dolmuş","🟠 Bu Hafta Bitiyor","🟡 Bu Ay Bitiyor"])
    with fc5:
        sel_unvan_f = st.selectbox("Ünvan", ["(Hepsi)"] + UNVAN_LISTESI[1:])

    df_filt = df_month.copy()
    if search_name.strip(): df_filt = df_filt[df_filt["Adı Soyadı"].str.contains(search_name.strip(), case=False)]
    if search_sicil.strip(): df_filt = df_filt[df_filt["Sicil No"].astype(str).str.contains(search_sicil.strip())]
    if sel_evrak != "(Hepsi)": df_filt = df_filt[df_filt["Evrak Adı"] == sel_evrak]
    if sel_durum != "(Hepsi)": df_filt = df_filt[df_filt["Durum"] == sel_durum]
    if sel_unvan_f != "(Hepsi)": df_filt = df_filt[df_filt["Ünvan"] == sel_unvan_f]

    show_cols = [c for c in ["Ünvan","Adı Soyadı","Sicil No","Evrak Adı",
                              "Başlangıç Tarihi","Bitiş Tarihi","Kalan Gün","Durum","Kaynak"]
                 if c in df_filt.columns]
    df_show = df_filt[show_cols].sort_values("Kalan Gün", ascending=True).copy()
    fmt = st.session_state["date_fmt"]
    for dc in ["Bitiş Tarihi","Başlangıç Tarihi"]:
        if dc in df_show.columns: df_show[dc] = df_show[dc].apply(lambda d: format_date(d, fmt))
    df_show["Kalan Gün"] = df_show["Kalan Gün"].apply(lambda x: int(x) if x is not None else "-")
    st.dataframe(df_show, use_container_width=True, hide_index=True, height=min(60+len(df_show)*36, 520))
    st.download_button("⬇️ Filtrelenmiş Excel", make_download_excel(df_show), f"kritik_{today.strftime('%Y%m%d')}.xlsx")

    with st.expander("📄 Tüm evraklar"):
        show_all = [c for c in ["Ünvan","Adı Soyadı","Sicil No","Evrak Adı",
                                 "Başlangıç Tarihi","Bitiş Tarihi","Kalan Gün","Durum","Kaynak"]
                    if c in df_final.columns]
        df_all = df_final[show_all].copy()
        for dc in ["Bitiş Tarihi","Başlangıç Tarihi"]:
            if dc in df_all.columns: df_all[dc] = df_all[dc].apply(lambda d: format_date(d, fmt))
        st.dataframe(df_all.sort_values("Kalan Gün"), use_container_width=True, hide_index=True, height=400)

# SEKME 2 – Manuel Evrak / Tarih Girişi
with tab2:
    st.markdown("<div class='section-label'>✍️ Evrak / Başlangıç Tarihi Gir</div>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("<div class='card-box'>", unsafe_allow_html=True)
        sec_p = st.selectbox("Personel", ["(Seçin)"]+personel_listesi, key="giris_p")
        if sec_p != "(Seçin)":
            mevcut_evraklar = sorted(df_final[df_final["Adı Soyadı"]==sec_p]["Evrak Adı"].dropna().unique().tolist())
            # Yeni evrak eklemeye izin ver
            sec_e = st.selectbox("Evrak", ["(Yeni evrak yaz...)"]+mevcut_evraklar, key="giris_e")
            if sec_e == "(Yeni evrak yaz...)":
                yeni_evrak = st.text_input("Evrak adını yazın")
                if yeni_evrak:
                    sec_e = yeni_evrak
            else:
                yeni_evrak = sec_e
            if sec_e and sec_e != "(Yeni evrak yaz...)":
                key_mt = (sec_p, sec_e)
                mevcut_bas = st.session_state["manuel_tarihler"].get(key_mt, {}).get("baslangic")
                if mevcut_bas is None:
                    satir = df_final[(df_final["Adı Soyadı"]==sec_p)&(df_final["Evrak Adı"]==sec_e)]
                    mevcut_bas = satir.iloc[0]["Başlangıç Tarihi"] if not satir.empty else today
                tarih = st.date_input("Giriş/Yenileme Tarihi", value=mevcut_bas, format="DD.MM.YYYY")
                sure = 2 if is_saglik_belgesi(sec_e) else 5
                bitis = hesapla_bitis(tarih, sec_e)
                st.success(f"Geçerlilik: {sure} yıl → Bitiş: {bitis.strftime('%d.%m.%Y')}")
                if st.button("💾 Kaydet"):
                    push_undo()
                    st.session_state["manuel_tarihler"][key_mt] = {"baslangic": tarih, "bitis": bitis}
                    save_state_and_excel()
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
    with col_b:
        st.markdown("<div class='card-box'>", unsafe_allow_html=True)
        st.markdown("**📋 Manuel Kayıtlar**")
        if not st.session_state["manuel_tarihler"]:
            st.info("Henüz manuel kayıt yok.")
        else:
            mt_rows = []
            for (ad, evrak), vals in st.session_state["manuel_tarihler"].items():
                kalan = (vals["bitis"] - today).days
                mt_rows.append({
                    "Personel": ad, "Evrak": evrak,
                    "Başlangıç": vals["baslangic"].strftime(fmt),
                    "Bitiş": vals["bitis"].strftime(fmt),
                    "Kalan Gün": kalan, "Durum": get_status_label(kalan)
                })
            st.dataframe(pd.DataFrame(mt_rows).sort_values("Kalan Gün"), use_container_width=True, hide_index=True)
            if st.button("🗑 Manuel kayıtları sıfırla"):
                push_undo()
                st.session_state["manuel_tarihler"] = {}
                save_state_and_excel()
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # Toplu ekleme
    st.markdown("---")
    st.markdown("**📎 Toplu Evrak Ekle (CSV/Excel)**")
    toplu_file = st.file_uploader("Dosya seç", type=["csv","xlsx","xls"], key="toplu")
    if toplu_file:
        df_toplu = load_excel(toplu_file)
        if df_toplu is not None:
            df_toplu = normalize_columns(df_toplu)
            if df_toplu is not None:
                st.dataframe(df_toplu.head())
                if st.button("✅ Bu kayıtları ekle"):
                    push_undo()
                    for _, row in df_toplu.iterrows():
                        ad = row["Adı Soyadı"]
                        evrak = row["Evrak Adı"]
                        bas = row.get("Başlangıç Tarihi", today)
                        if isinstance(bas, str): bas = datetime.datetime.strptime(bas, "%d.%m.%Y").date()
                        bitis = row.get("Bitiş Tarihi")
                        if isinstance(bitis, str) and bitis not in ("","-"):
                            bitis = datetime.datetime.strptime(bitis, "%d.%m.%Y").date()
                        else:
                            bitis = hesapla_bitis(bas, evrak)
                        st.session_state["manuel_tarihler"][(ad, evrak)] = {"baslangic": bas, "bitis": bitis}
                    save_state_and_excel()
                    st.success("Toplu kayıt eklendi.")
                    st.rerun()

# SEKME 3 – Kişi Evrak Düzenle / Sil
with tab3:
    st.markdown("<div class='section-label'>✏️ Kişi Evrak Düzenle / Sil</div>", unsafe_allow_html=True)
    sec_kisi = st.selectbox("Personel", ["(Seçin)"]+personel_listesi, key="kisi_duz")
    if sec_kisi != "(Seçin)":
        kisi_df = df_final[df_final["Adı Soyadı"]==sec_kisi]
        for idx, row in kisi_df.iterrows():
            key_str = row["_key"]
            with st.container():
                st.markdown("---")
                col1, col2, col3 = st.columns([2,1,1])
                with col1:
                    st.markdown(f"**{row['Evrak Adı']}** ({row['Kaynak']})")
                    st.caption(f"Başlangıç: {format_date(row['Başlangıç Tarihi'], fmt)} | Bitiş: {format_date(row['Bitiş Tarihi'], fmt)}")
                    st.caption(f"Kalan: {int(row['Kalan Gün']) if row['Kalan Gün'] else '-'} | Durum: {row['Durum']}")
                with col2:
                    if st.button("✏️ Düzenle", key=f"duz_{key_str}"):
                        st.session_state[f"edit_{key_str}"] = True
                with col3:
                    if st.button("🗑 Sil", key=f"sil_{key_str}"):
                        push_undo()
                        st.session_state["silinmis_evraklar"].append(key_str)
                        if row["Kaynak"] == "Manuel":
                            parts = key_str.rsplit("_manuel_",1)
                            if len(parts)==2:
                                ad_evrak = parts[0].split("|",1)
                                if len(ad_evrak)==2:
                                    ad, evrak = ad_evrak
                                    st.session_state["manuel_tarihler"].pop((ad, evrak), None)
                        save_state_and_excel()
                        st.rerun()
                if st.session_state.get(f"edit_{key_str}"):
                    with st.form(key=f"form_{key_str}"):
                        ybas = st.date_input("Başlangıç", row["Başlangıç Tarihi"] or today, key=f"bas_{key_str}")
                        ybit = st.date_input("Bitiş", row["Bitiş Tarihi"] or today, key=f"bit_{key_str}")
                        if st.form_submit_button("Kaydet"):
                            push_undo()
                            st.session_state["evrak_duzenlemeleri"][key_str] = {
                                "baslangic": ybas.strftime("%Y-%m-%d"),
                                "bitis": ybit.strftime("%Y-%m-%d")
                            }
                            save_state_and_excel()
                            st.session_state[f"edit_{key_str}"] = False
                            st.rerun()
                        if st.form_submit_button("İptal"):
                            st.session_state[f"edit_{key_str}"] = False
                            st.rerun()
    if st.button("🔄 Tüm düzenleme/silmeleri sıfırla"):
        push_undo()
        st.session_state["evrak_duzenlemeleri"] = {}
        st.session_state["silinmis_evraklar"] = []
        save_state_and_excel()
        st.rerun()

# SEKME 4 – Personel Özeti
with tab4:
    st.markdown("<div class='section-label'>👤 Personel Özeti</div>", unsafe_allow_html=True)
    sec_ozet = st.selectbox("Personel seçin", ["(Seçin)"]+personel_listesi, key="ozet_kisi")
    if sec_ozet != "(Seçin)":
        df_kisi = df_final[df_final["Adı Soyadı"]==sec_ozet]
        if df_kisi.empty:
            st.info("Kayıt yok.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                # pasta
                counts = df_kisi["Durum"].value_counts().reset_index()
                counts.columns = ["Durum", "Adet"]
                renkler = [DURUM_RENKLERI.get(d, "#95A5A6") for d in counts["Durum"]]
                fig = go.Figure(go.Pie(labels=counts["Durum"], values=counts["Adet"],
                                       marker_colors=renkler, hole=0.5))
                fig.update_layout(title=f"{sec_ozet} - Evrak Durumu", **PLOTLY_LAYOUT)
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                st.subheader("Kritik Evraklar")
                kritik = df_kisi[df_kisi["Durum"].isin(["🔴 Süresi Dolmuş","🟠 Bu Hafta Bitiyor"])]
                if kritik.empty:
                    st.success("Kritik evrak yok.")
                else:
                    for _, r in kritik.iterrows():
                        st.markdown(f"- **{r['Evrak Adı']}** ({r['Durum']}) – Bitiş: {format_date(r['Bitiş Tarihi'], fmt)}")
            st.dataframe(df_kisi[["Evrak Adı","Başlangıç Tarihi","Bitiş Tarihi","Kalan Gün","Durum"]]
                         .sort_values("Kalan Gün"), use_container_width=True)

# SEKME 5 – Grafikler & Analiz
with tab5:
    st.markdown("<div class='section-label'>📊 Genel Analiz</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        counts_all = df_final["Durum"].value_counts().reset_index()
        counts_all.columns = ["Durum", "Adet"]
        fig_pasta = go.Figure(go.Pie(labels=counts_all["Durum"], values=counts_all["Adet"],
                                     marker_colors=[DURUM_RENKLERI.get(d, "#95A5A6") for d in counts_all["Durum"]],
                                     hole=0.5))
        fig_pasta.update_layout(title="Tüm Evraklar", **PLOTLY_LAYOUT)
        st.plotly_chart(fig_pasta, use_container_width=True)
    with col2:
        # aylık bitiş
        ay_labels, ay_counts = [], []
        for i in range(12):
            mo = today.month - 1 + i
            y, m = today.year + mo // 12, mo % 12 + 1
            fd, ld = get_month_range(y, m)
            cnt = df_final["Bitiş Tarihi"].apply(lambda d: d is not None and fd <= d <= ld).sum()
            ay_labels.append(f"{calendar.month_abbr[m]}\n{y}")
            ay_counts.append(cnt)
        fig_bar = go.Figure(go.Bar(x=ay_labels, y=ay_counts, marker_color="#2980B9"))
        fig_bar.update_layout(title="Aylık Bitiş Takvimi", **PLOTLY_LAYOUT)
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")
    st.subheader("Özet İstatistikler")
    ozet = pd.DataFrame({
        "Metrik": ["Toplam Evrak","Personel","Evrak Türü","Süresi Dolmuş","Bu Hafta","Bu Ay","3 Ay","Geçerli"],
        "Değer": [len(df_final), df_final["Adı Soyadı"].nunique(), df_final["Evrak Adı"].nunique(),
                  int(df_final["Kalan Gün"]<0), int((df_final["Kalan Gün"]>=0)&(df_final["Kalan Gün"]<=7)),
                  int((df_final["Kalan Gün"]>=0)&(df_final["Kalan Gün"]<=31)),
                  int((df_final["Kalan Gün"]>=0)&(df_final["Kalan Gün"]<=90)),
                  int(df_final["Kalan Gün"]>90)]
    })
    st.dataframe(ozet, use_container_width=True, hide_index=True)
