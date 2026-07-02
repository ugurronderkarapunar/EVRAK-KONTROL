"""
Gemi / Personel Ehliyet ve Evrak Takip Uygulaması
-------------------------------------------------
Bu uygulama, yüklenen Excel dosyasındaki personel evraklarının
bitiş tarihlerini analiz ederek, içinde bulunulan ay içinde
süresi dolacak kritik belgeleri vurgular.
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime
import calendar

# ----------------------------- Sayfa Yapılandırması -----------------------------
st.set_page_config(
    page_title="Evrak Takip Sistemi",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------- Yardımcı Fonksiyonlar -----------------------------
@st.cache_data
def load_excel(file):
    """Yüklenen Excel dosyasını okur ve DataFrame olarak döndürür."""
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Excel dosyası okunurken hata oluştu: {e}")
        return None

def validate_columns(df, required_cols):
    """Gerekli sütunların DataFrame'de olup olmadığını kontrol eder."""
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        st.error(f"Excel dosyasında şu zorunlu sütunlar eksik: {', '.join(missing)}")
        st.info("Beklenen sütunlar: Adı Soyadı, Sicil No, Evrak Adı, Başlangıç Tarihi, Bitiş Tarihi")
        return False
    return True

def prepare_data(df):
    """Tarih sütunlarını dönüştürür, geçersiz değerleri temizler."""
    df = df.copy()
    # Bitiş Tarihi'ni datetime formatına çevir, hataları NaT yap
    df['Bitiş Tarihi'] = pd.to_datetime(df['Bitiş Tarihi'], dayfirst=True, errors='coerce')
    # Başlangıç Tarihi'ni de dönüştürelim (ileride belki kullanılır)
    df['Başlangıç Tarihi'] = pd.to_datetime(df['Başlangıç Tarihi'], dayfirst=True, errors='coerce')
    # Geçerli tarih olmayan satırları temizle
    df = df.dropna(subset=['Bitiş Tarihi']).reset_index(drop=True)
    return df

def filter_current_month(df):
    """Sadece içinde bulunulan aya ait bitiş tarihlerini filtreler."""
    today = date.today()
    first_day = date(today.year, today.month, 1)
    last_day = date(today.year, today.month, calendar.monthrange(today.year, today.month)[1])
    
    # Filtre: Bitiş Tarihi bu ay içinde (ilk ve son gün dahil)
    mask = (df['Bitiş Tarihi'].dt.date >= first_day) & (df['Bitiş Tarihi'].dt.date <= last_day)
    return df[mask].copy(), first_day, last_day

# ----------------------------- Arayüz: Üst Bilgi -----------------------------
st.title("🚢 Gemi / Personel Evrak Takip Paneli")
st.markdown("""
Bu uygulama, personelinizin ehliyet, sertifika ve diğer evraklarının **geçerlilik sürelerini** 
takip etmenizi sağlar. Sol panelden Excel dosyanızı yükleyin, personel ünvanını girin;
sistem **bu ay içinde süresi dolacak** tüm evrakları otomatik olarak listelesin.
""")

# ----------------------------- Kenar Çubuğu (Sidebar) -----------------------------
with st.sidebar:
    st.header("📁 Veri Yükleme ve Ayarlar")
    
    # Excel dosya yükleyici
    uploaded_file = st.file_uploader(
        "Personel evrak listesi (Excel)",
        type=["xlsx", "xls"],
        help="'Adı Soyadı', 'Sicil No', 'Evrak Adı', 'Başlangıç Tarihi', 'Bitiş Tarihi' sütunlarını içermelidir."
    )
    
    st.divider()
    
    # Personel Ünvanı seçimi
    st.subheader("👤 Personel Ünvanı")
    unvan_options = [
        "Seçiniz",
        "Kaptan",
        "Başmühendis",
        "İkinci Kaptan",
        "İkinci Mühendis",
        "Güverte Lostromosu",
        "Yağcı",
        "Aşçı",
        "Diğer (Manuel)"
    ]
    unvan_secim = st.selectbox("Ünvan seçiniz", unvan_options)
    
    # Manuel giriş alanı
    if unvan_secim == "Diğer (Manuel)":
        custom_unvan = st.text_input("Ünvanı yazınız", value="", placeholder="Örn: Elektrik Zabiti")
        final_unvan = custom_unvan.strip() if custom_unvan.strip() else "Belirtilmedi"
    elif unvan_secim == "Seçiniz":
        final_unvan = ""  # Boş bırakılacak, tabloda "Belirtilmedi" yazabilir
    else:
        final_unvan = unvan_secim

# ----------------------------- Ana Panel -----------------------------
if uploaded_file is None:
    # Dosya yüklenmemişse kullanıcıyı yönlendir
    st.info("ℹ️ Lütfen sol panelden personel evrak listesini içeren Excel dosyasını yükleyin.")
else:
    # 1. Excel'i oku
    df_raw = load_excel(uploaded_file)
    
    if df_raw is not None:
        # 2. Sütun kontrolü
        required_columns = ["Adı Soyadı", "Sicil No", "Evrak Adı", "Başlangıç Tarihi", "Bitiş Tarihi"]
        if validate_columns(df_raw, required_columns):
            # 3. Veriyi hazırla (tarih dönüşümleri)
            df_clean = prepare_data(df_raw)
            
            if df_clean.empty:
                st.warning("⚠️ Excel dosyasında geçerli 'Bitiş Tarihi' bilgisi bulunamadı.")
            else:
                # 4. Bu ay süresi dolanları filtrele
                df_this_month, first_day, last_day = filter_current_month(df_clean)
                
                # 5. Özet metrikleri hesapla
                # Toplam personel sayısı (benzersiz Sicil No veya Adı Soyadı)
                if 'Sicil No' in df_clean.columns:
                    total_personel = df_clean['Sicil No'].nunique()
                else:
                    total_personel = df_clean['Adı Soyadı'].nunique()
                
                # Bu ay süresi dolan evrak sayısı
                expired_count = len(df_this_month)
                
                # 6. Metrik kartlarını göster
                col1, col2, col3 = st.columns(3)
                col1.metric("👥 Toplam Personel", total_personel)
                col2.metric("📅 Bu Ay Takip Edilen Dönem", f"{first_day.strftime('%d.%m.%Y')} - {last_day.strftime('%d.%m.%Y')}")
                col3.metric("⚠️ Bu Ay Süresi Dolan Evrak", expired_count)
                
                st.divider()
                
                # 7. Kritik uyarı kutusu
                if expired_count > 0:
                    st.error(f"🚨 DİKKAT: Bu ay süresi dolan **{expired_count}** adet evrak bulunmaktadır! "
                             f"Aşağıdaki listede detayları inceleyip gerekli yenileme işlemlerini başlatın.")
                    
                    # 8. Görüntülenecek tabloyu oluştur
                    display_df = df_this_month.copy()
                    
                    # Kalan gün hesapla (bugüne göre)
                    today_ts = pd.Timestamp(date.today())
                    display_df['Kalan Gün Sayısı'] = (display_df['Bitiş Tarihi'] - today_ts).dt.days
                    
                    # Ünvan sütununu ekle (kullanıcının seçtiği / girdiği değer)
                    display_df['Ünvan'] = final_unvan if final_unvan else "Belirtilmedi"
                    
                    # Bitiş Tarihi'ni okunaklı formata çevir
                    display_df['Bitiş Tarihi'] = display_df['Bitiş Tarihi'].dt.strftime('%d.%m.%Y')
                    
                    # İstenen sütun sıralaması
                    column_order = ['Ünvan', 'Adı Soyadı', 'Sicil No', 'Evrak Adı', 'Bitiş Tarihi', 'Kalan Gün Sayısı']
                    # Varsa tüm sütunlar seç, yoksa eksik olanları atla (hata almamak için)
                    available_cols = [col for col in column_order if col in display_df.columns]
                    display_df = display_df[available_cols]
                    
                    # 9. İnteraktif tablo gösterimi
                    st.subheader("📋 Bu Ay Süresi Dolacak Evrak Listesi")
                    st.dataframe(
                        display_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Kalan Gün Sayısı": st.column_config.NumberColumn(
                                "Kalan Gün",
                                help="Pozitif: henüz dolmadı, 0: bugün doluyor, Negatif: süresi geçmiş"
                            )
                        }
                    )
                else:
                    # Eğer bu ay süresi dolan evrak yoksa bilgi mesajı
                    st.success("✅ Tebrikler! Bu ay içinde süresi dolacak herhangi bir evrak bulunmamaktadır.")
                    st.balloons()
