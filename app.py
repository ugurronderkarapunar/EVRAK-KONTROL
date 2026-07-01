"""
Gemiadamı Belgeleri Takip Sistemi - Streamlit Dashboard
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64
import re

# Sayfa yapılandırması
st.set_page_config(
    page_title="Gemiadamı Belge Takip Sistemi",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS özel stilleri
st.markdown("""
<style>
    .main-header { font-size: 2rem; font-weight: bold; color: #1a5276; }
    .stat-card { background: #f0f2f6; border-radius: 10px; padding: 15px; text-align: center; }
    .stat-number { font-size: 2.5rem; font-weight: bold; }
    .stat-label { font-size: 0.9rem; color: #555; }
    .traffic-green { color: #27ae60; }
    .traffic-yellow { color: #f1c40f; }
    .traffic-orange { color: #e67e22; }
    .traffic-red { color: #e74c3c; }
    .traffic-expired { color: #2c3e50; }
    .cert-card { background: #f8f9fa; border-radius: 8px; padding: 10px; margin: 5px 0; border-left: 4px solid #1a5276; }
    .filter-chip { display: inline-block; padding: 4px 12px; border-radius: 20px; background: #e9ecef; margin: 2px; cursor: pointer; }
    .filter-chip:hover { background: #dee2e6; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# VERİ YÜKLEME VE İŞLEME
# ============================================================

@st.cache_data
def load_data():
    """Excel verisini yükle ve işle"""
    # Veriyi doğrudan metinden oluştur
    data = []
    
    # Veri satırları - bu kısımda tüm veriyi işliyoruz
    # Veriyi parse etmek için regex kullanıyoruz
    
    raw_text = """FATİH AKKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-03|2026-07-02
    CAN ÇER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-07-05|2026-07-04
    MURAT AVŞAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-05|2027-11-04
    CAN ÇER|31-Gemi Adamı Cüzdan Belgesi|2021-07-05|2026-07-04
    ALPER BATUMOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-05|2026-07-04
    MURAT AVŞAR|02-Temel İlk Yardım Belgesi|2025-05-21|2030-05-20
    İSMAİL OCAK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-07-05|2026-07-04
    MUSTAFA DURGUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-07-05|2026-07-04
    MURAT AVŞAR|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-09-14|2029-09-13
    MURAT AVŞAR|31-Gemi Adamı Cüzdan Belgesi|2024-08-24|2029-08-23
    MUSTAFA DURGUN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-07-05|2026-07-04
    MUSTAFA DURGUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-07-05|2026-07-04
    EKREM ÖZÇELİK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-09|2027-04-08
    EKREM ÖZÇELİK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-01-21|2029-01-20
    MUSTAFA DURGUN|02-Temel İlk Yardım Belgesi|2021-07-05|2026-07-04
    MUSTAFA DURGUN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-07-05|2026-07-04
    MUSTAFA DURGUN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-07-05|2026-07-04
    MUSTAFA DURGUN|31-Gemi Adamı Cüzdan Belgesi|2021-07-05|2026-07-04
    VOLKAN KURTULUŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-07-07|2026-07-06
    EKREM ÖZÇELİK|31-Gemi Adamı Cüzdan Belgesi|2024-03-04|2029-03-03
    EMİN YILDIRIM|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-13|2029-12-12
    EMİN YILDIRIM|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-13|2029-12-12
    EMİN YILDIRIM|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-17|2027-06-16
    EMİN YILDIRIM|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-12-13|2029-12-12
    EMİN YILDIRIM|10-İleri Yangınla Mücadele Belgesi|2024-12-13|2029-12-12
    EMİN YILDIRIM|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-13|2029-12-12
    EMİN YILDIRIM|02-Temel İlk Yardım Belgesi|2024-12-13|2029-12-12
    EMİN YILDIRIM|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-13|2029-12-12
    EMİN YILDIRIM|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-13|2029-12-12
    EMİN YILDIRIM|31-Gemi Adamı Cüzdan Belgesi|2025-01-08|2030-01-07
    HAKAN AYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-01-26|2031-01-25
    HAKAN AYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-01-26|2031-01-25
    VOLKAN KURTULUŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-07-07|2026-07-06
    HAKAN AYDIN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-01-26|2031-01-25
    HAKAN AYDIN|10-İleri Yangınla Mücadele Belgesi|2026-01-26|2031-01-25
    HAKAN AYDIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-01-26|2031-01-25
    HAKAN AYDIN|02-Temel İlk Yardım Belgesi|2026-01-26|2031-01-25
    HAKAN AYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-01-26|2031-01-25
    HAKAN AYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-01-26|2031-01-25
    HAKAN AYDIN|31-Gemi Adamı Cüzdan Belgesi|2026-02-18|2031-02-17
    MAKSUT KEMERBAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|10-İleri Yangınla Mücadele Belgesi|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-21|2030-05-20
    MAKSUT KEMERBAŞ|02-Temel İlk Yardım Belgesi|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|20-Seyir Vardiyası Tutma Belgesi|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|30-Gemi Güvenlik Zabiti|2025-06-04|2030-06-03
    MAKSUT KEMERBAŞ|31-Gemi Adamı Cüzdan Belgesi|2025-06-05|2030-06-04
    MURAT HASAN KARAKUŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-31|2027-01-30
    MURAT HASAN KARAKUŞ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|10-İleri Yangınla Mücadele Belgesi|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|02-Temel İlk Yardım Belgesi|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-08-14|2028-08-13
    MURAT HASAN KARAKUŞ|31-Gemi Adamı Cüzdan Belgesi|2023-08-15|2028-08-14
    MUHAMMET KAKIL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-04-27|2031-04-26
    MUHAMMET KAKIL|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-04-27|2031-04-26
    MUHAMMET KAKIL|00-Gemiadamları Sağlık Yoklama Belgesi|2026-03-10|2028-03-09
    MUHAMMET KAKIL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-04-27|2031-04-26
    MUHAMMET KAKIL|02-Temel İlk Yardım Belgesi|2026-04-27|2031-04-26
    MUHAMMET KAKIL|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-04-27|2031-04-26
    MUHAMMET KAKIL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-04-27|2031-04-26
    MUHAMMET KAKIL|31-Gemi Adamı Cüzdan Belgesi|2026-05-06|2031-05-05
    ADNAN KARTAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-03-01|2027-02-28
    ADNAN KARTAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-03-01|2027-02-28
    ADNAN KARTAL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-09|2027-04-08
    ADNAN KARTAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-03-01|2027-02-28
    ADNAN KARTAL|02-Temel İlk Yardım Belgesi|2022-03-01|2027-02-28
    ADNAN KARTAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-03-01|2027-02-28
    ADNAN KARTAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-29|2030-05-28
    ADNAN KARTAL|31-Gemi Adamı Cüzdan Belgesi|2023-03-01|2028-02-29
    ALİ İHSAN ÇOM|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-12|2027-06-11
    ALİ İHSAN ÇOM|10-İleri Yangınla Mücadele Belgesi|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|02-Temel İlk Yardım Belgesi|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-14|2030-05-13
    ALİ İHSAN ÇOM|31-Gemi Adamı Cüzdan Belgesi|2025-05-14|2030-05-13
    ARİF TÜYSÜZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-17|2027-09-16
    FATİH KARTOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-08-07|2030-08-06
    FATİH KARTOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-08-07|2030-08-06
    FATİH KARTOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-17|2030-09-16
    FATİH KARTOĞLU|02-Temel İlk Yardım Belgesi|2025-08-07|2030-08-06
    FATİH KARTOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-08-07|2030-08-06
    FATİH KARTOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-08-07|2030-08-06
    FATİH KARTOĞLU|31-Gemi Adamı Cüzdan Belgesi|2025-09-09|2030-05-20
    FEDAİ SOMUNCU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-18|2027-07-17
    FEDAİ SOMUNCU|10-İleri Yangınla Mücadele Belgesi|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|02-Temel İlk Yardım Belgesi|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-11-28|2030-11-27
    FEDAİ SOMUNCU|31-Gemi Adamı Cüzdan Belgesi|2025-12-02|2030-12-01
    HAMİT BİLGİLİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-01-21|2030-01-20
    HAMİT BİLGİLİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-01-20|2030-01-19
    VOLKAN KURTULUŞ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-07-07|2026-07-06
    HAMİT BİLGİLİ|10-İleri Yangınla Mücadele Belgesi|2025-01-21|2030-01-20
    HAMİT BİLGİLİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-20|2030-05-19
    HAMİT BİLGİLİ|02-Temel İlk Yardım Belgesi|2025-01-21|2030-01-20
    HAMİT BİLGİLİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-20|2030-05-19
    HAMİT BİLGİLİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-20|2030-05-19
    HAMİT BİLGİLİ|31-Gemi Adamı Cüzdan Belgesi|2025-05-20|2030-05-19
    İBRAHİM OLGUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-11|2031-02-10
    İBRAHİM OLGUN|10-İleri Yangınla Mücadele Belgesi|2026-02-11|2031-02-10
    İBRAHİM OLGUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-11|2031-02-10
    İBRAHİM OLGUN|02-Temel İlk Yardım Belgesi|2026-02-11|2031-02-10
    İBRAHİM OLGUN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-11|2031-02-10
    İBRAHİM OLGUN|31-Gemi Adamı Cüzdan Belgesi|2026-02-13|2031-02-12
    MACİT AKMERMER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-04|2030-06-03
    MACİT AKMERMER|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-04|2030-06-03
    MACİT AKMERMER|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-06|2027-10-05
    MACİT AKMERMER|10-İleri Yangınla Mücadele Belgesi|2025-06-04|2030-06-03
    MACİT AKMERMER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-17|2030-09-16
    MACİT AKMERMER|02-Temel İlk Yardım Belgesi|2025-06-04|2030-06-03
    MACİT AKMERMER|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-04|2030-06-03
    MACİT AKMERMER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-06-04|2030-06-03
    MACİT AKMERMER|31-Gemi Adamı Cüzdan Belgesi|2025-06-05|2030-06-04
    MUZAFFER MORGÜL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-11|2027-09-10
    MUZAFFER MORGÜL|10-İleri Yangınla Mücadele Belgesi|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|02-Temel İlk Yardım Belgesi|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-17|2029-09-16
    MUZAFFER MORGÜL|31-Gemi Adamı Cüzdan Belgesi|2024-09-17|2029-09-16
    NURETTİN ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-21|2030-05-20
    NURETTİN ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-31|2029-12-30
    NURETTİN ÖZTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-13|2027-06-12
    NURETTİN ÖZTÜRK|10-İleri Yangınla Mücadele Belgesi|2024-12-31|2029-12-30
    NURETTİN ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-31|2029-12-30
    NURETTİN ÖZTÜRK|02-Temel İlk Yardım Belgesi|2024-12-31|2029-12-30
    NURETTİN ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-31|2029-12-30
    NURETTİN ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-31|2029-12-30
    NURETTİN ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2025-01-17|2030-01-16
    VOLKAN KURTULUŞ|10-İleri Yangınla Mücadele Belgesi|2021-07-07|2026-07-06
    VOLKAN KURTULUŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-07-07|2026-07-06
    RAHMİ GÖK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-15|2027-10-14
    VOLKAN KURTULUŞ|02-Temel İlk Yardım Belgesi|2021-07-07|2026-07-06
    VOLKAN KURTULUŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-07-07|2026-07-06
    VOLKAN KURTULUŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-07-07|2026-07-06
    VOLKAN KURTULUŞ|31-Gemi Adamı Cüzdan Belgesi|2021-07-07|2026-07-06
    İBRAHİM MARAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-07-09|2026-07-08
    İBRAHİM MARAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-07-09|2026-07-08
    SERKAN AKTEPE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-26|2030-09-25
    SERKAN AKTEPE|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-26|2030-09-25
    SERKAN AKTEPE|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-17|2028-06-16
    SERKAN AKTEPE|10-İleri Yangınla Mücadele Belgesi|2025-09-26|2030-09-25
    SERKAN AKTEPE|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-26|2030-09-25
    SERKAN AKTEPE|02-Temel İlk Yardım Belgesi|2025-09-26|2030-09-25
    SERKAN AKTEPE|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-26|2030-09-25
    SERKAN AKTEPE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-26|2030-09-25
    SERKAN AKTEPE|31-Gemi Adamı Cüzdan Belgesi|2025-09-29|2030-09-28
    SERKAN YÜKSEL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-19|2027-06-18
    YAKUP BABAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-05-30|2027-05-29
    YAKUP BABAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-05-30|2027-05-29
    YAKUP BABAN|10-İleri Yangınla Mücadele Belgesi|2022-05-30|2027-05-29
    YAKUP BABAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-05-30|2027-05-29
    YAKUP BABAN|02-Temel İlk Yardım Belgesi|2022-05-30|2027-05-29
    YAKUP BABAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-05-30|2027-05-29
    YAKUP BABAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-05-30|2027-05-29
    YAKUP YARDIMCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-12|2028-04-11
    YAKUP YARDIMCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-12|2028-04-11
    YAKUP YARDIMCI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-26|2027-09-25
    YAKUP YARDIMCI|10-İleri Yangınla Mücadele Belgesi|2023-04-12|2028-04-11
    YAKUP YARDIMCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-12|2028-04-11
    YAKUP YARDIMCI|02-Temel İlk Yardım Belgesi|2023-04-12|2028-04-11
    YAKUP YARDIMCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-12|2028-04-11
    YAKUP YARDIMCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-12|2028-04-11
    YAKUP YARDIMCI|31-Gemi Adamı Cüzdan Belgesi|2023-04-12|2028-04-11
    YÜKSEL TÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2029-11-05|2030-05-20
    YÜKSEL TÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2029-11-05|2030-05-20
    İBRAHİM MARAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-07-09|2026-07-08
    YÜKSEL TÜRK|10-İleri Yangınla Mücadele Belgesi|2029-11-05|2030-05-20
    YÜKSEL TÜRK|02-Temel İlk Yardım Belgesi|2029-11-05|2030-05-20
    YÜKSEL TÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2029-11-05|2030-05-20
    YÜKSEL TÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2029-11-05|2030-05-20
    YÜKSEL TÜRK|31-Gemi Adamı Cüzdan Belgesi|2024-12-10|2029-12-09
    ABDUL BÜYÜK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-11|2027-04-10
    ADEM KARAKUŞ|01-Güverte Kısmı Gemi Adamı Belgesi 1|2024-09-20|2029-09-19
    ADEM KARAKUŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-20|2029-09-19
    ADEM KARAKUŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-20|2029-09-19
    ADEM KARAKUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-08|2027-08-07
    ADEM KARAKUŞ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-05-20|2030-05-19
    ADEM KARAKUŞ|10-İleri Yangınla Mücadele Belgesi|2024-09-20|2029-09-19
    ADEM KARAKUŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-20|2030-05-19
    ADEM KARAKUŞ|02-Temel İlk Yardım Belgesi|2024-09-20|2029-09-19
    ADEM KARAKUŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-20|2029-09-19
    ADEM KARAKUŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-20|2029-09-19
    ADEM KARAKUŞ|31-Gemi Adamı Cüzdan Belgesi|2024-09-23|2029-09-22
    AHMET ERKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-04|2029-12-03
    AHMET ERKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-04|2029-12-03
    AHMET ERKAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-10|2027-12-09
    AHMET ERKAN|10-İleri Yangınla Mücadele Belgesi|2024-12-04|2029-12-03
    AHMET ERKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-04|2029-12-03
    AHMET ERKAN|02-Temel İlk Yardım Belgesi|2024-12-04|2029-12-03
    AHMET ERKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-04|2029-12-03
    AHMET ERKAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-04|2029-12-03
    AHMET ERKAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-03-02|2027-03-01
    AHMET ERKAN|31-Gemi Adamı Cüzdan Belgesi|2024-12-05|2029-12-04
    ALİ ALPER GÖRKEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-02|2027-10-01
    ALİ ALPER GÖRKEN|10-İleri Yangınla Mücadele Belgesi|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|02-Temel İlk Yardım Belgesi|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-12-03|2029-12-02
    ALİ ALPER GÖRKEN|31-Gemi Adamı Cüzdan Belgesi|2024-12-04|2029-12-03
    ALPASLAN BULUT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-05-24|2028-05-23
    ALPASLAN BULUT|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-24|2028-05-23
    ALPASLAN BULUT|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-24|2027-01-23
    ALPASLAN BULUT|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-02-03|2027-02-02
    ALPASLAN BULUT|10-İleri Yangınla Mücadele Belgesi|2023-05-24|2028-05-23
    ALPASLAN BULUT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-24|2028-05-23
    ALPASLAN BULUT|02-Temel İlk Yardım Belgesi|2023-05-24|2028-05-23
    ALPASLAN BULUT|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-24|2028-05-23
    ALPASLAN BULUT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-05-24|2028-05-23
    ALPASLAN BULUT|31-Gemi Adamı Cüzdan Belgesi|2022-02-03|2027-02-02
    İBRAHİM MARAŞ|02-Temel İlk Yardım Belgesi|2021-07-09|2026-07-08
    İBRAHİM MARAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-07-09|2026-07-08
    ARSLAN MURAT GÜL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-14|2027-07-13
    İBRAHİM MARAŞ|31-Gemi Adamı Cüzdan Belgesi|2021-07-09|2026-07-08
    SABRİ KILIÇ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-10|2026-07-09
    YUNUS EMRE ZAMAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-10|2026-07-09
    MUSTAFA DURGUN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-10|2026-07-09
    NAZİF İLYAS|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-11|2026-07-10
    ÜMİT ŞEN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-12|2026-07-11
    GÖKHAN ERSOY|10-İleri Yangınla Mücadele Belgesi|2021-07-12|2026-07-11
    AYDIN OĞUZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-01-14|2031-01-13
    AYDIN OĞUZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-01-14|2031-01-13
    AYDIN OĞUZ|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-14|2028-01-13
    AYDIN OĞUZ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-03-14|2029-03-13
    AYDIN OĞUZ|10-İleri Yangınla Mücadele Belgesi|2026-01-14|2031-01-13
    AYDIN OĞUZ|02-Temel İlk Yardım Belgesi|2026-01-14|2031-01-13
    AYDIN OĞUZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-01-14|2031-01-13
    AYDIN OĞUZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-01-14|2031-01-13
    AYDIN OĞUZ|31-Gemi Adamı Cüzdan Belgesi|2026-01-14|2031-01-13
    HARUN DEMİRCİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-07-13|2026-07-12
    HARUN DEMİRCİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-07-13|2026-07-12
    HARUN DEMİRCİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-07-13|2026-07-12
    AYHAN PİŞKİNER|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-02-02|2029-02-01
    HARUN DEMİRCİ|02-Temel İlk Yardım Belgesi|2021-07-13|2026-07-12
    HARUN DEMİRCİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-07-13|2026-07-12
    HARUN DEMİRCİ|31-Gemi Adamı Cüzdan Belgesi|2021-07-13|2026-07-12
    MEHMET AKİF KALAFAT|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-07-16|2026-07-15
    MEHMET AKİF KALAFAT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-07-16|2026-07-15
    MEHMET AKİF KALAFAT|02-Temel İlk Yardım Belgesi|2021-07-16|2026-07-15
    AYKUN EMANET|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-04|2027-02-03
    MEHMET AKİF KALAFAT|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-07-16|2026-07-15
    AYKUN EMANET|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-04|2027-02-03
    AYKUN EMANET|02-Temel İlk Yardım Belgesi|2022-02-04|2027-02-03
    AYKUN EMANET|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-04|2027-02-03
    AYKUN EMANET|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-04|2027-02-03
    AYKUN EMANET|31-Gemi Adamı Cüzdan Belgesi|2022-02-04|2027-02-03
    BEDİR YANARDAĞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|10-İleri Yangınla Mücadele Belgesi|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|02-Temel İlk Yardım Belgesi|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|30-Gemi Güvenlik Zabiti|2026-02-03|2031-02-02
    BEDİR YANARDAĞ|31-Gemi Adamı Cüzdan Belgesi|2026-02-04|2031-02-03
    BİLAL KURU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-08|2030-07-07
    BİLAL KURU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-08|2030-07-07
    BİLAL KURU|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-20|2028-01-19
    BİLAL KURU|10-İleri Yangınla Mücadele Belgesi|2025-07-08|2030-07-07
    BİLAL KURU|02-Temel İlk Yardım Belgesi|2025-07-08|2030-07-07
    BİLAL KURU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-08|2030-07-07
    BİLAL KURU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-07-08|2030-07-07
    BİLAL KURU|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-07-06|2027-07-05
    BİLAL KURU|30-Gemi Güvenlik Zabiti|2025-07-08|2030-07-07
    BİLAL KURU|31-Gemi Adamı Cüzdan Belgesi|2022-07-06|2027-07-05
    BURHAN KADİR EREL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-11-02|2028-11-01
    BURHAN KADİR EREL|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-11-02|2028-11-01
    BURHAN KADİR EREL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-08|2027-04-07
    BURHAN KADİR EREL|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-11-02|2028-11-01
    BURHAN KADİR EREL|10-İleri Yangınla Mücadele Belgesi|2023-11-02|2028-11-01
    BURHAN KADİR EREL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-11-02|2028-11-01
    BURHAN KADİR EREL|02-Temel İlk Yardım Belgesi|2023-11-02|2028-11-01
    BURHAN KADİR EREL|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-11-02|2028-11-01
    BURHAN KADİR EREL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-11-02|2028-11-01
    BURHAN KADİR EREL|31-Gemi Adamı Cüzdan Belgesi|2023-11-06|2028-11-05
    CAN ÇER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-16|2029-12-15
    CAN ÇER|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-16|2029-12-15
    CAN ÇER|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-28|2027-02-27
    CAN ÇER|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-01-23|2029-01-22
    CAN ÇER|10-İleri Yangınla Mücadele Belgesi|2024-12-16|2029-12-15
    CAN ÇER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-16|2029-12-15
    CAN ÇER|02-Temel İlk Yardım Belgesi|2024-12-16|2029-12-15
    CAN ÇER|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-16|2029-12-15
    MEHMET AKİF KALAFAT|31-Gemi Adamı Cüzdan Belgesi|2021-07-16|2026-07-15
    HAKAN AKSU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-17|2026-07-16
    CUMHUR KAZMAZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-17|2029-12-13
    CUMHUR KAZMAZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-17|2029-12-13
    CUMHUR KAZMAZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-16|2027-05-15
    CUMHUR KAZMAZ|10-İleri Yangınla Mücadele Belgesi|2024-12-17|2029-12-13
    FURGAN GÜNÇALDI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-17|2026-07-17
    CUMHUR KAZMAZ|21-Seyir ve Gemi İdaresi Belgesi|2024-12-24|2029-12-23
    CUMHUR KAZMAZ|02-Temel İlk Yardım Belgesi|2024-12-17|2029-12-13
    CUMHUR KAZMAZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-17|2029-12-16
    CUMHUR KAZMAZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-17|2029-12-13
    CUMHUR KAZMAZ|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-07-09|2028-07-08
    CUMHUR KAZMAZ|31-Gemi Adamı Cüzdan Belgesi|2024-12-24|2029-12-23
    EMRE YILMAZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-05-09|2028-05-08
    EMRE YILMAZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-09|2028-05-08
    EMRE YILMAZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-09|2028-05-08
    EMRE YILMAZ|02-Temel İlk Yardım Belgesi|2023-05-09|2028-05-08
    EMRE YILMAZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-09|2028-05-08
    EMRE YILMAZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-05-09|2028-05-08
    EMRE ZOR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-30|2030-06-29
    EMRE ZOR|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-30|2030-06-29
    FATİH ORAL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-19|2026-07-18
    EMRE ZOR|10-İleri Yangınla Mücadele Belgesi|2025-06-30|2030-06-29
    EMRE ZOR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-18|2030-09-17
    EMRE ZOR|02-Temel İlk Yardım Belgesi|2025-06-30|2030-06-29
    EMRE ZOR|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-30|2030-06-29
    EMRE ZOR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-06-30|2030-06-29
    EMRE ZOR|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-04-17|2028-04-16
    EMRE ZOR|30-Gemi Güvenlik Zabiti|2025-06-30|2030-06-29
    EMRE ZOR|31-Gemi Adamı Cüzdan Belgesi|2023-04-17|2028-04-16
    MUSTAFA ÖZTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-19|2026-07-18
    ENVER AVCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-18|2030-09-17
    ENVER AVCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-18|2030-09-17
    ENVER AVCI|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-25|2028-06-24
    ENVER AVCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-18|2030-09-17
    ENVER AVCI|02-Temel İlk Yardım Belgesi|2025-09-18|2030-09-17
    ENVER AVCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-18|2030-09-17
    ENVER AVCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-18|2030-09-17
    ENVER AVCI|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-03-12|2029-03-11
    ENVER AVCI|31-Gemi Adamı Cüzdan Belgesi|2025-10-12|2030-09-17
    ERAY KIZILKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-23|2030-12-22
    ERAY KIZILKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-23|2030-12-22
    FATİH SAMED GÖRMÜŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-22|2026-07-21
    ERAY KIZILKAYA|10-İleri Yangınla Mücadele Belgesi|2025-12-23|2030-12-22
    ERAY KIZILKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-23|2030-12-22
    ERAY KIZILKAYA|02-Temel İlk Yardım Belgesi|2025-12-23|2030-12-22
    ERAY KIZILKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-23|2030-12-22
    ERAY KIZILKAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-23|2030-12-22
    ERAY KIZILKAYA|30-Gemi Güvenlik Zabiti|2025-12-23|2030-12-22
    ERAY KIZILKAYA|31-Gemi Adamı Cüzdan Belgesi|2025-12-23|2030-12-22
    ERBİL EKİZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-07|2029-08-06
    ERBİL EKİZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-08-07|2029-08-06
    ERBİL EKİZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-13|2027-06-12
    ERBİL EKİZ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-08-07|2029-08-06
    ERBİL EKİZ|10-İleri Yangınla Mücadele Belgesi|2024-08-07|2029-08-06
    ERBİL EKİZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-08-07|2029-08-06
    ERBİL EKİZ|02-Temel İlk Yardım Belgesi|2024-08-07|2029-08-06
    ERBİL EKİZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-08-07|2029-08-06
    ERBİL EKİZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-08-07|2029-08-06
    ERBİL EKİZ|31-Gemi Adamı Cüzdan Belgesi|2024-08-15|2029-08-14
    ERCAN AYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-03-13|2030-03-12
    ERCAN AYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-03-13|2030-03-12
    GÖKHAN GÖZTOK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-23|2026-07-22
    ERCAN AYDIN|10-İleri Yangınla Mücadele Belgesi|2025-03-13|2030-03-12
    ERCAN AYDIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-03-13|2030-03-12
    ERCAN AYDIN|02-Temel İlk Yardım Belgesi|2025-03-13|2030-03-12
    ERCAN AYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-03-13|2030-03-12
    ERCAN AYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-03-13|2030-03-12
    ERCAN AYDIN|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-03-13|2030-03-12
    ERCAN AYDIN|31-Gemi Adamı Cüzdan Belgesi|2025-03-14|2030-03-13
    ERCAN DÖNMEZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-30|2030-06-29
    ERCAN DÖNMEZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-30|2030-06-29
    TOLGA GÖKSU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-23|2026-07-22
    ERCAN DÖNMEZ|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2022-10-21|2027-10-20
    ERCAN DÖNMEZ|02-Temel İlk Yardım Belgesi|2025-06-30|2030-06-29
    ERCAN DÖNMEZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-30|2030-06-29
    ERCAN DÖNMEZ|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-05-16|2027-05-15
    ERCAN DÖNMEZ|31-Gemi Adamı Cüzdan Belgesi|2022-05-17|2027-05-16
    ERDEM ALTUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-08|2027-09-07
    ERDEM ALTUN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-08|2027-09-07
    ERDEM ALTUN|10-İleri Yangınla Mücadele Belgesi|2022-09-08|2027-09-07
    ERDEM ALTUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-08|2027-09-07
    ERDEM ALTUN|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2022-09-08|2027-09-07
    ERDEM ALTUN|02-Temel İlk Yardım Belgesi|2022-09-08|2027-09-07
    ERDEM ALTUN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-08|2027-09-07
    EREN EMRAH TAŞAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-10|2027-04-09
    EREN EMRAH TAŞAY|10-İleri Yangınla Mücadele Belgesi|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|02-Temel İlk Yardım Belgesi|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-02-23|2029-02-22
    EREN EMRAH TAŞAY|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-02-14|2027-02-13
    EREN EMRAH TAŞAY|31-Gemi Adamı Cüzdan Belgesi|2025-03-19|2030-03-18
    ERHAN ARSLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-30|2027-07-29
    ERKAN KESKİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-10|2029-09-09
    ERKAN KESKİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-10|2029-09-09
    ERKAN KESKİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-04|2027-07-03
    ERKAN KESKİN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-09-10|2029-09-09
    ERKAN KESKİN|10-İleri Yangınla Mücadele Belgesi|2024-09-10|2029-09-09
    ERKAN KESKİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-10|2029-09-09
    ERKAN KESKİN|02-Temel İlk Yardım Belgesi|2024-09-10|2029-09-09
    ERKAN KESKİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-10|2029-09-09
    ERKAN KESKİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-10|2029-09-09
    ERKAN KESKİN|31-Gemi Adamı Cüzdan Belgesi|2024-10-26|2029-10-25
    ERKAN UYANIK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-08-25|2028-08-24
    ERKAN UYANIK|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-08-25|2028-08-24
    ERKAN UYANIK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-10|2027-03-09
    ERKAN UYANIK|10-İleri Yangınla Mücadele Belgesi|2023-08-25|2028-08-24
    ERKAN UYANIK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-08-25|2028-08-24
    ERKAN UYANIK|02-Temel İlk Yardım Belgesi|2023-08-25|2028-08-24
    ERKAN UYANIK|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-25|2028-08-24
    ERKAN UYANIK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-08-25|2028-08-24
    ERKAN UYANIK|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-08-25|2028-08-24
    ERKAN UYANIK|31-Gemi Adamı Cüzdan Belgesi|2023-08-28|2028-08-27
    ERKAN UZUNER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-10|2028-04-09
    ERKAN UZUNER|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-10|2028-04-09
    AYHAN PİŞKİNER|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-24|2026-07-23
    ERKAN UZUNER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-10|2028-04-09
    ERKAN UZUNER|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2023-04-10|2028-04-09
    ERKAN UZUNER|02-Temel İlk Yardım Belgesi|2023-04-10|2028-04-09
    ERKAN UZUNER|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-10|2028-04-09
    ERKAN UZUNER|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-06-28|2027-06-27
    ERKAN UZUNER|31-Gemi Adamı Cüzdan Belgesi|2026-01-19|2031-01-18
    AYKUN EMANET|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-24|2026-07-23
    SERKAN MAMAT|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-24|2026-07-23
    BARIŞ TERLEMEZ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-25|2026-07-24
    MELİH ERÇAYAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-25|2026-07-24
    KEMAL KOCATÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-26|2026-07-25
    TURGAY MUŞDAL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-26|2026-07-25
    İLKAY KARABUL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-26|2026-07-25
    YUSUF CANDAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-26|2026-07-25
    UĞUR ÇER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-07-29|2026-07-28
    UĞUR ÇER|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-07-29|2026-07-28
    ERTUĞRUL ARPAKÇI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-09|2027-10-08
    ERTUĞRUL ARPAKÇI|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-01-30|2029-01-29
    ERTUĞRUL ARPAKÇI|10-İleri Yangınla Mücadele Belgesi|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|02-Temel İlk Yardım Belgesi|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|20-Seyir Vardiyası Tutma Belgesi|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-23|2031-02-22
    ERTUĞRUL ARPAKÇI|31-Gemi Adamı Cüzdan Belgesi|2026-03-03|2031-03-02
    EYÜP ÇELİKKOL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-23|2030-12-22
    EYÜP ÇELİKKOL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-23|2030-12-22
    UĞUR ÇER|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-07-29|2026-07-28
    EYÜP ÇELİKKOL|10-İleri Yangınla Mücadele Belgesi|2025-12-23|2030-12-22
    EYÜP ÇELİKKOL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-23|2030-12-22
    EYÜP ÇELİKKOL|02-Temel İlk Yardım Belgesi|2025-12-23|2030-12-22
    EYÜP ÇELİKKOL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-23|2030-12-22
    EYÜP ÇELİKKOL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-23|2030-12-22
    EYÜP ÇELİKKOL|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-10-17|2027-10-16
    EYÜP ÇELİKKOL|31-Gemi Adamı Cüzdan Belgesi|2025-12-24|2030-12-23
    FATİH AKKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-19|2030-09-18
    FATİH AKKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-19|2030-09-18
    UĞUR ÇER|10-İleri Yangınla Mücadele Belgesi|2021-07-29|2026-07-28
    FATİH AKKAYA|10-İleri Yangınla Mücadele Belgesi|2025-09-19|2030-09-18
    FATİH AKKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-19|2030-09-18
    FATİH AKKAYA|02-Temel İlk Yardım Belgesi|2025-09-19|2030-09-18
    FATİH AKKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-19|2030-09-18
    FATİH AKKAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-19|2030-09-18
    FATİH AKKAYA|31-Gemi Adamı Cüzdan Belgesi|2025-09-22|2030-09-21
    FATİH ORAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-01-08|2030-01-07
    FATİH ORAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-01-08|2030-01-07
    UĞUR ÇER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-07-29|2026-07-28
    FATİH ORAL|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-01-08|2030-01-07
    FATİH ORAL|10-İleri Yangınla Mücadele Belgesi|2025-01-08|2030-01-07
    FATİH ORAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-01-08|2030-01-07
    FATİH ORAL|02-Temel İlk Yardım Belgesi|2025-01-08|2030-01-07
    FATİH ORAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-01-08|2030-01-07
    FATİH ORAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-01-08|2030-01-07
    FATİH ORAL|31-Gemi Adamı Cüzdan Belgesi|2025-01-08|2030-01-07
    FATİH TÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-14|2030-07-13
    FATİH TÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-14|2030-07-13
    FATİH TÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-08|2027-04-07
    FATİH TÜRK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-01-29|2029-01-28
    FATİH TÜRK|10-İleri Yangınla Mücadele Belgesi|2025-07-14|2030-07-13
    FATİH TÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-18|2030-09-17
    FATİH TÜRK|02-Temel İlk Yardım Belgesi|2025-07-14|2030-07-13
    FATİH TÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-14|2030-07-13
    FATİH TÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-07-14|2030-07-13
    FATİH TÜRK|30-Gemi Güvenlik Zabiti|2025-07-14|2030-07-13
    FATİH TÜRK|31-Gemi Adamı Cüzdan Belgesi|2025-07-16|2030-07-15
    FERHAT SARIKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-17|2030-12-16
    FERHAT SARIKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-17|2030-12-16
    FERHAT SARIKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-15|2027-01-14
    FERHAT SARIKAYA|10-İleri Yangınla Mücadele Belgesi|2025-12-17|2030-12-16
    FERHAT SARIKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-17|2030-12-16
    FERHAT SARIKAYA|02-Temel İlk Yardım Belgesi|2025-12-17|2030-12-16
    FERHAT SARIKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-17|2030-12-16
    FERHAT SARIKAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-17|2030-12-16
    FERHAT SARIKAYA|31-Gemi Adamı Cüzdan Belgesi|2025-12-22|2030-12-21
    GÖKHAN GÖZTOK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-11-29|2028-11-28
    GÖKHAN GÖZTOK|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-11-29|2028-11-28
    UĞUR ÇER|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-07-29|2026-07-28
    GÖKHAN GÖZTOK|10-İleri Yangınla Mücadele Belgesi|2023-11-29|2028-11-28
    GÖKHAN GÖZTOK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-11-29|2028-11-28
    GÖKHAN GÖZTOK|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2023-11-29|2028-11-28
    GÖKHAN GÖZTOK|02-Temel İlk Yardım Belgesi|2023-11-29|2028-11-28
    GÖKHAN GÖZTOK|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-11-29|2028-11-28
    GÖKHAN GÖZTOK|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-06-02|2027-06-01
    GÜVEN HENDEM|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-04-12|2030-05-19
    GÜVEN HENDEM|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-05-16|2027-05-15
    UĞUR ÇER|02-Temel İlk Yardım Belgesi|2021-07-29|2026-07-28
    UĞUR ÇER|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-07-29|2026-07-28
    HAKAN KAMAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-24|2027-04-23
    UĞUR ÇER|31-Gemi Adamı Cüzdan Belgesi|2021-07-29|2026-07-28
    ENVER AVCI|01-Güverte Kısmı Gemi Adamı Belgesi 1|2021-07-30|2026-07-29
    MUHAMMED ATIF YÜKSELOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-07-31|2026-07-30
    HÜSEYİN ŞİŞMAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-02|2026-08-01
    HÜSEYİN ŞİŞMAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-02|2026-08-01
    HÜSEYİN ŞİŞMAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-02|2026-08-01
    HAKAN ŞENOL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-22|2029-08-21
    HAKAN ŞENOL|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-23|2028-05-22
    HAKAN ŞENOL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-22|2027-09-21
    HAKAN ŞENOL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-23|2028-05-22
    HAKAN ŞENOL|02-Temel İlk Yardım Belgesi|2023-05-23|2028-05-22
    HAKAN ŞENOL|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-23|2028-05-22
    HAKAN ŞENOL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-11-20|2028-11-19
    HAKAN ŞENOL|31-Gemi Adamı Cüzdan Belgesi|2023-10-06|2028-10-05
    HALUK KIZILPINAR|01-Güverte Kısmı Gemi Adamı Belgesi 1|2025-01-21|2030-01-20
    HALUK KIZILPINAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-12|2029-02-11
    HALUK KIZILPINAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-02-12|2029-02-11
    HALUK KIZILPINAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-12|2027-09-22
    HALUK KIZILPINAR|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-01-21|2030-01-20
    HALUK KIZILPINAR|10-İleri Yangınla Mücadele Belgesi|2024-02-12|2029-02-11
    HALUK KIZILPINAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-01-06|2028-01-05
    HALUK KIZILPINAR|02-Temel İlk Yardım Belgesi|2024-02-12|2029-02-11
    HALUK KIZILPINAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-12|2029-02-11
    HALUK KIZILPINAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-02-12|2029-02-11
    HALUK KIZILPINAR|31-Gemi Adamı Cüzdan Belgesi|2025-01-21|2030-01-20
    HASAN AYDEMİR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-06-17|2027-06-16
    HASAN AYDEMİR|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-17|2027-06-16
    HASAN AYDEMİR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-10|2027-09-09
    HASAN AYDEMİR|10-İleri Yangınla Mücadele Belgesi|2022-06-17|2027-06-16
    HASAN AYDEMİR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-06-17|2027-06-16
    HASAN AYDEMİR|02-Temel İlk Yardım Belgesi|2022-06-17|2027-06-16
    HASAN AYDEMİR|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-17|2027-06-16
    HASAN AYDEMİR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-06-17|2027-06-16
    HASAN AYDEMİR|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-12-13|2029-12-12
    HASAN AYDEMİR|31-Gemi Adamı Cüzdan Belgesi|2024-12-13|2029-12-12
    HASAN DALCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-31|2030-12-30
    HASAN DALCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-31|2030-12-30
    HASAN DALCI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-29|2027-07-28
    HASAN DALCI|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-12-31|2030-12-30
    HASAN DALCI|10-İleri Yangınla Mücadele Belgesi|2025-12-31|2030-12-30
    HASAN DALCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-31|2030-12-30
    HASAN DALCI|02-Temel İlk Yardım Belgesi|2025-12-31|2030-12-30
    HASAN DALCI|20-Seyir Vardiyası Tutma Belgesi|2025-12-31|2030-12-30
    HASAN DALCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-31|2030-12-30
    HASAN DALCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-31|2030-12-30
    HASAN DALCI|31-Gemi Adamı Cüzdan Belgesi|2025-12-31|2030-12-30
    HÜSEYİN DEMİRKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-30|2030-06-29
    HÜSEYİN DEMİRKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-30|2030-06-29
    HÜSEYİN DEMİRKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-20|2027-06-19
    HÜSEYİN DEMİRKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-06-30|2030-06-29
    HÜSEYİN DEMİRKAYA|02-Temel İlk Yardım Belgesi|2025-06-30|2030-06-29
    HÜSEYİN DEMİRKAYA|20-Seyir Vardiyası Tutma Belgesi|2025-06-30|2030-06-29
    HÜSEYİN DEMİRKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-30|2030-06-29
    HÜSEYİN DEMİRKAYA|31-Gemi Adamı Cüzdan Belgesi|2023-09-21|2028-09-20
    HÜSEYİN TUĞCU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-02-14|2030-02-13
    HÜSEYİN TUĞCU|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-14|2027-02-13
    HÜSEYİN TUĞCU|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|10-İleri Yangınla Mücadele Belgesi|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|02-Temel İlk Yardım Belgesi|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-09-15|2028-09-14
    HÜSEYİN TUĞCU|31-Gemi Adamı Cüzdan Belgesi|2023-09-18|2028-09-17
    HÜSEYİN ŞİŞMAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-02|2026-08-01
    İBRAHİM SARI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-03-24|2029-03-23
    İBRAHİM SARI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-14|2027-05-13
    İBRAHİM SARI|10-İleri Yangınla Mücadele Belgesi|2024-03-24|2029-03-23
    İBRAHİM SARI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-03-24|2029-03-23
    İBRAHİM SARI|02-Temel İlk Yardım Belgesi|2024-03-24|2029-03-23
    İBRAHİM SARI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-03-24|2029-03-23
    İBRAHİM SARI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-03-24|2029-03-23
    İBRAHİM SARI|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-03-24|2029-03-23
    İBRAHİM SARI|31-Gemi Adamı Cüzdan Belgesi|2024-03-24|2029-03-23
    İDRİS USTALAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-01-05|2029-01-04
    İDRİS USTALAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-01-05|2029-01-04
    İDRİS USTALAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-03|2027-01-02
    İDRİS USTALAR|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-01-05|2029-01-04
    İDRİS USTALAR|10-İleri Yangınla Mücadele Belgesi|2024-01-05|2029-01-04
    İDRİS USTALAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-01-05|2029-01-04
    İDRİS USTALAR|02-Temel İlk Yardım Belgesi|2024-01-05|2029-01-04
    İDRİS USTALAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-01-05|2029-01-04
    İDRİS USTALAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-01-05|2029-01-04
    İDRİS USTALAR|31-Gemi Adamı Cüzdan Belgesi|2024-01-08|2029-01-07
    İLKER BAYSAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-09-22|2028-09-21
    İLKER BAYSAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-09-22|2028-09-21
    İLKER BAYSAL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-23|2027-09-22
    İLKER BAYSAL|10-İleri Yangınla Mücadele Belgesi|2023-09-22|2028-09-21
    İLKER BAYSAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-09-22|2028-09-21
    İLKER BAYSAL|02-Temel İlk Yardım Belgesi|2023-09-22|2028-09-21
    İLKER BAYSAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-09-22|2028-09-21
    İLKER BAYSAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-09-22|2028-09-21
    İLKER BAYSAL|31-Gemi Adamı Cüzdan Belgesi|2022-03-07|2027-03-06
    İLKER HAYTA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-04|2029-12-03
    İLKER HAYTA|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-04|2029-12-03
    HÜSEYİN ŞİŞMAN|02-Temel İlk Yardım Belgesi|2021-08-02|2026-08-01
    İLKER HAYTA|10-İleri Yangınla Mücadele Belgesi|2024-12-04|2029-12-03
    İLKER HAYTA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-04|2029-12-03
    İLKER HAYTA|02-Temel İlk Yardım Belgesi|2024-12-04|2029-12-03
    İLKER HAYTA|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-04|2029-12-03
    İLKER HAYTA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-04|2029-12-03
    İLKER HAYTA|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-04-14|2028-04-13
    İLKER HAYTA|31-Gemi Adamı Cüzdan Belgesi|2024-12-05|2029-12-04
    HÜSEYİN ŞİŞMAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-02|2026-08-01
    HÜSEYİN ŞİŞMAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-08-02|2026-08-01
    HÜSEYİN ŞİŞMAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-08-02|2026-08-01
    ÖZGE ALP ÇINAR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-05|2026-08-04
    İMRAN BAYAR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-05|2026-08-04
    MUHAMMET ÇAKIR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-06|2026-08-05
    MUHAMMET ÇAKIR|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-06|2026-08-05
    MUHAMMET ÇAKIR|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-08-06|2026-08-05
    MUHAMMET ÇAKIR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-06|2026-08-05
    MUHAMMET ÇAKIR|02-Temel İlk Yardım Belgesi|2021-08-06|2026-08-05
    İSMAİL ÖZCAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-07|2027-11-06
    İSMAİL ÖZCAN|10-İleri Yangınla Mücadele Belgesi|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|02-Temel İlk Yardım Belgesi|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-05-07|2029-05-06
    İSMAİL ÖZCAN|31-Gemi Adamı Cüzdan Belgesi|2024-05-08|2029-05-07
    KAAN KAHRAMAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-05|2029-09-04
    KAAN KAHRAMAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-05|2029-09-04
    KAAN KAHRAMAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-06|2027-11-05
    KAAN KAHRAMAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-05|2029-09-04
    KAAN KAHRAMAN|02-Temel İlk Yardım Belgesi|2024-09-05|2029-09-04
    KAAN KAHRAMAN|20-Seyir Vardiyası Tutma Belgesi|2024-09-05|2029-09-04
    KAAN KAHRAMAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-05|2029-09-04
    KAAN KAHRAMAN|31-Gemi Adamı Cüzdan Belgesi|2024-09-06|2029-09-05
    KEMAL GÜRKAN ŞAHİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-02|2027-05-01
    KEMAL GÜRKAN ŞAHİN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-04-03|2029-04-02
    KEMAL GÜRKAN ŞAHİN|10-İleri Yangınla Mücadele Belgesi|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|02-Temel İlk Yardım Belgesi|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|20-Seyir Vardiyası Tutma Belgesi|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-27|2029-12-26
    KEMAL GÜRKAN ŞAHİN|31-Gemi Adamı Cüzdan Belgesi|2025-01-01|2029-12-31
    MUHAMMET ÇAKIR|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-06|2026-08-05
    MUHAMMET ÇAKIR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-08-06|2026-08-05
    KORAY TORLAK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-12|2027-03-11
    MUHAMMET ÇAKIR|31-Gemi Adamı Cüzdan Belgesi|2021-08-06|2026-08-05
    RABİA SELİN DELİOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-06|2026-08-05
    GÖKHAN TARLACI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-07|2026-08-06
    ÖZLEM GÜRSU TARLACI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-07|2026-08-06
    ZAFER PAPAKER|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-07|2026-08-06
    SERDAR ŞİMŞEK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-08|2026-08-07
    MEHMET BİLGİÇ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-22|2027-07-21
    MEHMET BİLGİÇ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-01-24|2030-01-23
    MEHMET BİLGİÇ|10-İleri Yangınla Mücadele Belgesi|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|02-Temel İlk Yardım Belgesi|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-16|2031-02-15
    MEHMET BİLGİÇ|31-Gemi Adamı Cüzdan Belgesi|2026-02-17|2031-02-16
    MEHMET DURMUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-09|2026-08-08
    İSMET PERİŞANOĞLU|25-Kısa Mesafe Telsiz Operatörü Belgesi|2021-08-09|2026-08-08
    UĞURCAN SÖNMEZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-10|2026-08-09
    UĞURCAN SÖNMEZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-10|2026-08-09
    UĞURCAN SÖNMEZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-10|2026-08-09
    UĞURCAN SÖNMEZ|02-Temel İlk Yardım Belgesi|2021-08-10|2026-08-09
    UĞURCAN SÖNMEZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-10|2026-08-09
    UĞURCAN SÖNMEZ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-08-11|2026-08-10
    UĞURCAN SÖNMEZ|31-Gemi Adamı Cüzdan Belgesi|2021-08-11|2026-08-10
    ÜMİT ALKANOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-14|2026-08-13
    İSMAİL YALÇINKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-16|2026-08-15
    İSMAİL YALÇINKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-16|2026-08-15
    MEHMET LOR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-04-14|2031-04-13
    İSMAİL YALÇINKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-16|2026-08-15
    MEHMET LOR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-11|2028-04-10
    MEHMET LOR|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-08-07|2028-08-06
    MEHMET LOR|30-Gemi Güvenlik Zabiti|2024-01-18|2029-01-17
    MEHMET LOR|31-Gemi Adamı Cüzdan Belgesi|2023-08-07|2028-08-06
    MİTAT ESER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-20|2030-06-19
    MİTAT ESER|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-20|2030-06-19
    MİTAT ESER|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-23|2027-01-22
    MİTAT ESER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-06-20|2030-06-19
    MİTAT ESER|02-Temel İlk Yardım Belgesi|2025-06-20|2030-06-19
    MİTAT ESER|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-20|2030-06-19
    MİTAT ESER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-06-20|2030-06-19
    MİTAT ESER|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-03-14|2027-03-13
    MİTAT ESER|31-Gemi Adamı Cüzdan Belgesi|2025-06-20|2030-06-19
    MUAMMER KEÇECİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-04|2028-06-03
    MUAMMER KEÇECİ|10-İleri Yangınla Mücadele Belgesi|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|02-Temel İlk Yardım Belgesi|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-09-19|2030-09-18
    MUAMMER KEÇECİ|31-Gemi Adamı Cüzdan Belgesi|2025-09-19|2030-09-18
    MUBİN HUT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-17|2030-07-16
    MUBİN HUT|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-17|2030-07-16
    MUBİN HUT|10-İleri Yangınla Mücadele Belgesi|2025-07-17|2030-07-16
    MUBİN HUT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-07-17|2030-07-16
    MUBİN HUT|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2025-07-15|2030-07-14
    MUBİN HUT|02-Temel İlk Yardım Belgesi|2025-07-17|2030-07-16
    MUBİN HUT|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-17|2030-07-16
    MUBİN HUT|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-08-09|2029-08-08
    MUBİN HUT|31-Gemi Adamı Cüzdan Belgesi|2025-07-23|2030-07-22
    İSMAİL YALÇINKAYA|02-Temel İlk Yardım Belgesi|2021-08-16|2026-08-15
    İSMAİL YALÇINKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-16|2026-08-15
    MUHAMMET ÇAKIR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-25|2027-02-24
    NURETTİN AYDIN|25-Kısa Mesafe Telsiz Operatörü Belgesi|2021-08-16|2026-08-15
    RAHMİ GÖK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-18|2026-08-17
    RAHMİ GÖK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-18|2026-08-17
    RAHMİ GÖK|10-İleri Yangınla Mücadele Belgesi|2021-08-18|2026-08-17
    RAHMİ GÖK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-18|2026-08-17
    RAHMİ GÖK|02-Temel İlk Yardım Belgesi|2021-08-18|2026-08-17
    MURAT ALEV|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-11-25|2030-11-24
    MURAT ALEV|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-11-25|2030-11-24
    MURAT ALEV|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-25|2027-09-24
    MURAT ALEV|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-11-25|2030-11-24
    MURAT ALEV|10-İleri Yangınla Mücadele Belgesi|2025-11-25|2030-11-24
    MURAT ALEV|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-11-25|2030-11-24
    MURAT ALEV|02-Temel İlk Yardım Belgesi|2025-11-25|2030-11-24
    MURAT ALEV|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-11-25|2030-11-24
    MURAT ALEV|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-11-25|2030-11-24
    MURAT ALEV|30-Gemi Güvenlik Zabiti|2025-11-25|2030-11-24
    MURAT ALEV|31-Gemi Adamı Cüzdan Belgesi|2025-11-26|2030-11-25
    MURAT ASLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-05|2027-03-04
    MURAT ASLAN|21-Seyir ve Gemi İdaresi Belgesi|2022-05-16|2027-05-15
    MURAT ASLAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-11-14|2027-11-13
    MURAT ASLAN|31-Gemi Adamı Cüzdan Belgesi|2022-11-22|2027-11-21
    MURAT SEKMEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-23|2029-09-22
    MURAT SEKMEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-23|2029-09-22
    MURAT SEKMEN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-07|2027-05-06
    MURAT SEKMEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-23|2029-09-22
    MURAT SEKMEN|02-Temel İlk Yardım Belgesi|2024-09-23|2029-09-22
    MURAT SEKMEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-23|2029-09-22
    MURAT SEKMEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-23|2029-09-22
    MURAT SEKMEN|31-Gemi Adamı Cüzdan Belgesi|2024-09-23|2029-09-22
    RAHMİ GÖK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-18|2026-08-17
    RAHMİ GÖK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-08-18|2026-08-17
    MURAT TAFRALI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-09|2027-09-08
    FAHRİ YAZICI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-19|2026-08-18
    FAHRİ YAZICI|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-19|2026-08-18
    FAHRİ YAZICI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-19|2026-08-18
    FAHRİ YAZICI|02-Temel İlk Yardım Belgesi|2021-08-19|2026-08-18
    FAHRİ YAZICI|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-19|2026-08-18
    FAHRİ YAZICI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-08-19|2026-08-18
    FAHRİ YAZICI|31-Gemi Adamı Cüzdan Belgesi|2021-08-19|2026-08-18
    MURAT ÜNAL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-02|2027-01-01
    NUSRET İLKER ÇETİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-05|2027-08-04
    NUSRET İLKER ÇETİN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|10-İleri Yangınla Mücadele Belgesi|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|02-Temel İlk Yardım Belgesi|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-24|2027-01-23
    NUSRET İLKER ÇETİN|31-Gemi Adamı Cüzdan Belgesi|2022-01-24|2027-01-23
    MEHMET SEFER|10-İleri Yangınla Mücadele Belgesi|2021-08-19|2026-08-18
    AHMET YILMAZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-23|2026-08-22
    AHMET YILMAZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-23|2026-08-22
    AHMET YILMAZ|10-İleri Yangınla Mücadele Belgesi|2021-08-23|2026-08-22
    AHMET YILMAZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-23|2026-08-22
    AHMET YILMAZ|02-Temel İlk Yardım Belgesi|2021-08-23|2026-08-22
    AHMET YILMAZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-23|2026-08-22
    ORHAN GEÇER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-12-07|2028-12-06
    ORHAN GEÇER|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-12-07|2028-12-06
    HAKAN KAMAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-24|2026-08-23
    ORHAN GEÇER|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-12-07|2028-12-06
    ORHAN GEÇER|10-İleri Yangınla Mücadele Belgesi|2023-12-07|2028-12-06
    ORHAN GEÇER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-12-07|2028-12-06
    ORHAN GEÇER|02-Temel İlk Yardım Belgesi|2023-12-07|2028-12-06
    ORHAN GEÇER|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-12-07|2028-12-06
    ORHAN GEÇER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-12-07|2028-12-06
    ORHAN GEÇER|31-Gemi Adamı Cüzdan Belgesi|2023-12-07|2028-12-06
    HAKAN KAMAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-24|2026-08-23
    HAKAN KAMAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-24|2026-08-23
    HAKAN KAMAN|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-08-24|2026-08-23
    HAKAN KAMAN|02-Temel İlk Yardım Belgesi|2021-08-24|2026-08-23
    HAKAN KAMAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-24|2026-08-23
    HAKAN KAMAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-08-24|2026-08-23
    AHMET YILMAZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-08-24|2026-08-23
    AHMET YILMAZ|31-Gemi Adamı Cüzdan Belgesi|2021-08-24|2026-08-23
    HAKAN KAMAN|31-Gemi Adamı Cüzdan Belgesi|2021-08-25|2026-08-24
    RAMAZAN KONYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-22|2030-12-21
    RAMAZAN KONYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-22|2030-12-21
    RAMAZAN KONYA|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-08|2027-04-07
    RAMAZAN KONYA|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-09-08|2029-09-07
    RAMAZAN KONYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-22|2030-12-21
    RAMAZAN KONYA|02-Temel İlk Yardım Belgesi|2025-12-22|2030-12-21
    RAMAZAN KONYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-22|2030-12-21
    RAMAZAN KONYA|30-Gemi Güvenlik Zabiti|2025-12-22|2030-12-21
    RAMAZAN KONYA|31-Gemi Adamı Cüzdan Belgesi|2025-12-22|2030-12-21
    RASİM BARUT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-03-18|2030-03-17
    RASİM BARUT|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-03-18|2030-03-17
    RASİM BARUT|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-10|2027-12-09
    RASİM BARUT|10-İleri Yangınla Mücadele Belgesi|2025-03-18|2030-03-17
    RASİM BARUT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-03-18|2030-03-17
    RASİM BARUT|02-Temel İlk Yardım Belgesi|2025-03-18|2030-03-17
    RASİM BARUT|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-03-18|2030-03-17
    RASİM BARUT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-03-18|2030-03-17
    RASİM BARUT|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-03-18|2030-03-17
    RASİM BARUT|31-Gemi Adamı Cüzdan Belgesi|2025-03-19|2030-03-18
    RECEP ÇORUHLİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-04|2027-02-03
    RECEP ÇORUHLİ|10-İleri Yangınla Mücadele Belgesi|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|02-Temel İlk Yardım Belgesi|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-08-07|2028-08-06
    RECEP ÇORUHLİ|31-Gemi Adamı Cüzdan Belgesi|2023-08-07|2028-08-06
    MERT CEMİL AKÇAY|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-26|2026-08-25
    METİN ÇİMEN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-28|2026-08-27
    RECEP MICIK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-09|2027-05-08
    YUSUF ÖZKAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-08-29|2026-08-28
    NAZİF İLYAS|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-08-31|2026-08-30
    NAZİF İLYAS|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-08-31|2026-08-30
    NAZİF İLYAS|10-İleri Yangınla Mücadele Belgesi|2021-08-31|2026-08-30
    NAZİF İLYAS|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-08-31|2026-08-30
    NAZİF İLYAS|02-Temel İlk Yardım Belgesi|2021-08-31|2026-08-30
    RESÜL ATAMAN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-06-09|2030-05-20
    SADULLAH ÇİFTCİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-18|2027-08-17
    SADULLAH ÇİFTCİ|10-İleri Yangınla Mücadele Belgesi|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|02-Temel İlk Yardım Belgesi|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-01-24|2027-01-23
    SADULLAH ÇİFTCİ|31-Gemi Adamı Cüzdan Belgesi|2022-01-24|2027-01-23
    NAZİF İLYAS|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-08-31|2026-08-30
    NAZİF İLYAS|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-08-31|2026-08-30
    SALİH ACUN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-19|2027-11-18
    MURAT KESKİN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-02|2026-09-01
    MUSTAFA SOYTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-03|2026-09-02
    MUSTAFA SOYTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-03|2026-09-02
    MUSTAFA SOYTÜRK|10-İleri Yangınla Mücadele Belgesi|2021-09-03|2026-09-02
    MUSTAFA SOYTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-03|2026-09-02
    MUSTAFA SOYTÜRK|02-Temel İlk Yardım Belgesi|2021-09-03|2026-09-02
    MUSTAFA SOYTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-03|2026-09-02
    SERHAT AKBAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-03-04|2031-03-03
    SERHAT AKBAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-03-04|2031-03-03
    SERHAT AKBAY|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-10|2027-12-09
    SERHAT AKBAY|10-İleri Yangınla Mücadele Belgesi|2026-03-04|2031-03-03
    SERHAT AKBAY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-03-04|2031-03-03
    SERHAT AKBAY|02-Temel İlk Yardım Belgesi|2026-03-04|2031-03-03
    SERHAT AKBAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-03-04|2031-03-03
    SERHAT AKBAY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-03-04|2031-03-03
    SERHAT AKBAY|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-05-21|2030-05-20
    SERHAT AKBAY|30-Gemi Güvenlik Zabiti|2024-02-26|2029-02-25
    SERHAT AKBAY|31-Gemi Adamı Cüzdan Belgesi|2026-03-05|2031-03-04
    SERKAN AYTAÇ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-01-28|2031-01-27
    SERKAN AYTAÇ|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-01-28|2031-01-27
    SERKAN AYTAÇ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-22|2027-05-20
    MUSTAFA SOYTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-03|2026-09-02
    SERKAN AYTAÇ|10-İleri Yangınla Mücadele Belgesi|2026-01-28|2031-01-27
    SERKAN AYTAÇ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-01-28|2031-01-27
    SERKAN AYTAÇ|02-Temel İlk Yardım Belgesi|2026-01-28|2031-01-27
    SERKAN AYTAÇ|20-Seyir Vardiyası Tutma Belgesi|2026-01-28|2031-01-27
    SERKAN AYTAÇ|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-01-28|2031-01-27
    SERKAN AYTAÇ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-01-28|2031-01-27
    SERKAN AYTAÇ|31-Gemi Adamı Cüzdan Belgesi|2026-01-28|2031-01-27
    SERKAN ŞAHİN|01-Güverte Kısmı Gemi Adamı Belgesi 1|2024-09-20|2029-09-19
    SERKAN ŞAHİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-12-08|2027-12-07
    SERKAN ŞAHİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-12-08|2027-12-07
    SERKAN ŞAHİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-04|2027-09-03
    SERKAN ŞAHİN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-12-08|2027-12-07
    SERKAN ŞAHİN|10-İleri Yangınla Mücadele Belgesi|2022-12-08|2027-12-07
    SERKAN ŞAHİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-12-08|2027-12-07
    SERKAN ŞAHİN|02-Temel İlk Yardım Belgesi|2022-12-08|2027-12-07
    SERKAN ŞAHİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-12-08|2027-12-07
    SERKAN ŞAHİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-12-08|2027-12-07
    SERKAN ŞAHİN|31-Gemi Adamı Cüzdan Belgesi|2022-12-09|2027-12-08
    SONER HUTOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-02-06|2030-02-05
    SONER HUTOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-02-06|2030-02-05
    MUSTAFA SOYTÜRK|31-Gemi Adamı Cüzdan Belgesi|2021-09-03|2026-09-02
    SONER HUTOĞLU|10-İleri Yangınla Mücadele Belgesi|2025-02-06|2030-02-05
    SONER HUTOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-02-06|2030-02-05
    SONER HUTOĞLU|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2025-02-06|2030-02-05
    SONER HUTOĞLU|02-Temel İlk Yardım Belgesi|2025-02-06|2030-02-05
    SONER HUTOĞLU|20-Seyir Vardiyası Tutma Belgesi|2025-02-06|2030-02-05
    SONER HUTOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-02-06|2030-02-05
    SONER HUTOĞLU|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-02-06|2030-02-05
    SONER HUTOĞLU|31-Gemi Adamı Cüzdan Belgesi|2025-02-07|2030-02-06
    SÜLEYMAN ERKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-09|2027-09-08
    SÜLEYMAN ERKAN|10-İleri Yangınla Mücadele Belgesi|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|02-Temel İlk Yardım Belgesi|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-02-06|2030-02-05
    SÜLEYMAN ERKAN|31-Gemi Adamı Cüzdan Belgesi|2025-02-18|2030-02-17
    ŞABAN ÇELİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-04-09|2031-04-08
    ŞABAN ÇELİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-04-09|2031-04-08
    ŞABAN ÇELİK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-28|2027-03-27
    ŞABAN ÇELİK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-04-09|2031-04-08
    ŞABAN ÇELİK|10-İleri Yangınla Mücadele Belgesi|2026-04-09|2031-04-08
    ŞABAN ÇELİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-04-09|2031-04-08
    ŞABAN ÇELİK|02-Temel İlk Yardım Belgesi|2026-04-09|2031-04-08
    ŞABAN ÇELİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-04-09|2031-04-08
    ŞABAN ÇELİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-04-09|2031-04-08
    ŞABAN ÇELİK|30-Gemi Güvenlik Zabiti|2026-04-09|2031-04-08
    ŞABAN ÇELİK|31-Gemi Adamı Cüzdan Belgesi|2026-04-10|2031-04-09
    TAMER ARSLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-15|2027-12-14
    TAMER ARSLAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-12-14|2027-12-13
    TAMER ARSLAN|31-Gemi Adamı Cüzdan Belgesi|2022-12-27|2027-12-26
    TANER METE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-10-04|2028-10-03
    TANER METE|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-10-04|2028-10-03
    TANER METE|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-07|2027-01-06
    TANER METE|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-10-04|2028-10-03
    TANER METE|10-İleri Yangınla Mücadele Belgesi|2023-10-04|2028-10-03
    TANER METE|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-10-04|2028-10-03
    TANER METE|02-Temel İlk Yardım Belgesi|2023-10-04|2028-10-03
    TANER METE|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-10-04|2028-10-03
    TANER METE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-10-04|2028-10-03
    TANER METE|31-Gemi Adamı Cüzdan Belgesi|2023-10-06|2028-10-05
    TAYFUN ERENDAĞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-06|2027-11-05
    FURKAN ÇAKMAKTAŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-03|2026-09-02
    YÜKSEL TÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-06|2026-09-05
    İSHAK TAŞCI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-06|2026-09-05
    YUSUF YASİN KADEM SARI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-09|2026-09-08
    MUSTAFA KÖKSEL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-09|2026-09-08
    HÜSEYİN KÖMÜRCÜ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-11|2026-09-08
    HÜSEYİN KÖMÜRCÜ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-11|2026-09-08
    HÜSEYİN KÖMÜRCÜ|02-Temel İlk Yardım Belgesi|2021-10-11|2026-09-08
    HÜSEYİN KÖMÜRCÜ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-11|2026-09-08
    ERSEN DİKMEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-10|2026-09-09
    TURGUT ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-27|2027-11-26
    TURGUT ÖZTÜRK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-12-04|2028-12-03
    TURGUT ÖZTÜRK|10-İleri Yangınla Mücadele Belgesi|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|02-Temel İlk Yardım Belgesi|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-01-16|2031-01-15
    TURGUT ÖZTÜRK|30-Gemi Güvenlik Zabiti|2024-09-10|2029-09-09
    TURGUT ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2024-01-29|2029-01-28
    ERSEN DİKMEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-10|2026-09-09
    ERSEN DİKMEN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-09-10|2026-09-09
    UĞUR ÇER|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-04|2027-03-03
    ERSEN DİKMEN|10-İleri Yangınla Mücadele Belgesi|2021-09-10|2026-09-09
    ERSEN DİKMEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-10|2026-09-09
    ERSEN DİKMEN|02-Temel İlk Yardım Belgesi|2021-09-10|2026-09-09
    ERSEN DİKMEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-10|2026-09-09
    ERSEN DİKMEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-10|2026-09-09
    ERSEN DİKMEN|31-Gemi Adamı Cüzdan Belgesi|2021-09-10|2026-09-09
    ARSLAN MURAT GÜL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-13|2026-09-12
    UĞUR FETTAHOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2026-03-13|2028-03-12
    UĞUR FETTAHOĞLU|10-İleri Yangınla Mücadele Belgesi|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|02-Temel İlk Yardım Belgesi|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-12-12|2029-12-11
    UĞUR FETTAHOĞLU|31-Gemi Adamı Cüzdan Belgesi|2024-12-12|2029-12-11
    ULAŞ YÜCE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-04-13|2031-04-12
    ULAŞ YÜCE|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-04-13|2031-04-12
    ULAŞ YÜCE|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-24|2027-09-23
    ULAŞ YÜCE|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-04-13|2031-04-12
    ULAŞ YÜCE|10-İleri Yangınla Mücadele Belgesi|2026-04-13|2031-04-12
    ULAŞ YÜCE|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-04-13|2031-04-12
    ULAŞ YÜCE|02-Temel İlk Yardım Belgesi|2026-04-13|2031-04-12
    ULAŞ YÜCE|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-04-13|2031-04-12
    ULAŞ YÜCE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-04-13|2031-04-12
    ULAŞ YÜCE|31-Gemi Adamı Cüzdan Belgesi|2026-04-14|2031-04-13
    ÜMİT ŞEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-05-14|2031-05-13
    ÜMİT ŞEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-06-14|2031-06-13
    ARSLAN MURAT GÜL|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-13|2026-09-12
    ÜMİT ŞEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-05-14|2031-05-13
    ÜMİT ŞEN|02-Temel İlk Yardım Belgesi|2026-05-14|2031-05-13
    ÜMİT ŞEN|20-Seyir Vardiyası Tutma Belgesi|2025-10-09|2030-10-08
    ÜMİT ŞEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-05-14|2031-05-13
    ÜMİT ŞEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-05-14|2031-05-13
    ÜMİT ŞEN|31-Gemi Adamı Cüzdan Belgesi|2025-10-09|2030-10-08
    VOLKAN DEPE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-15|2030-05-14
    VOLKAN DEPE|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-05-15|2030-05-14
    VOLKAN DEPE|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-05|2027-08-04
    VOLKAN DEPE|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-05-15|2030-05-14
    VOLKAN DEPE|10-İleri Yangınla Mücadele Belgesi|2025-05-15|2030-05-14
    VOLKAN DEPE|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-15|2030-05-14
    VOLKAN DEPE|02-Temel İlk Yardım Belgesi|2025-05-15|2030-05-14
    VOLKAN DEPE|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-15|2030-05-14
    VOLKAN DEPE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-15|2030-05-14
    VOLKAN DEPE|31-Gemi Adamı Cüzdan Belgesi|2025-05-16|2030-05-15
    ARSLAN MURAT GÜL|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-09-13|2026-09-12
    ARSLAN MURAT GÜL|10-İleri Yangınla Mücadele Belgesi|2021-09-13|2026-09-12
    VOLKAN KURTULUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-17|2027-10-16
    ARSLAN MURAT GÜL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-13|2026-09-12
    ARSLAN MURAT GÜL|02-Temel İlk Yardım Belgesi|2021-09-13|2026-09-12
    ARSLAN MURAT GÜL|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-13|2026-09-12
    ARSLAN MURAT GÜL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-13|2026-09-12
    ARSLAN MURAT GÜL|31-Gemi Adamı Cüzdan Belgesi|2021-09-13|2026-09-12
    RIFAT MERT ÖZAYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-14|2026-09-13
    RIFAT MERT ÖZAYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-14|2026-09-13
    YASİN GÖZLEMECİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-15|2027-01-14
    YASİN GÖZLEMECİ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-02-02|2028-02-01
    YASİN GÖZLEMECİ|10-İleri Yangınla Mücadele Belgesi|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|02-Temel İlk Yardım Belgesi|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|20-Seyir Vardiyası Tutma Belgesi|2025-05-20|2030-05-19
    YASİN GÖZLEMECİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-31|2030-12-30
    YASİN GÖZLEMECİ|31-Gemi Adamı Cüzdan Belgesi|2023-02-02|2028-02-01
    RIFAT MERT ÖZAYDIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-14|2026-09-13
    RIFAT MERT ÖZAYDIN|02-Temel İlk Yardım Belgesi|2021-09-14|2026-09-13
    RIFAT MERT ÖZAYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-14|2026-09-13
    RIFAT MERT ÖZAYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-14|2026-09-13
    RIFAT MERT ÖZAYDIN|31-Gemi Adamı Cüzdan Belgesi|2021-09-14|2026-09-13
    RECEP ASLAN|10-İleri Yangınla Mücadele Belgesi|2021-09-15|2026-09-14
    EYÜP CÜCÜ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-16|2026-09-15
    EYÜP CÜCÜ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-16|2026-09-15
    EYÜP CÜCÜ|10-İleri Yangınla Mücadele Belgesi|2021-09-16|2026-09-15
    EYÜP CÜCÜ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-16|2026-09-15
    ZAFER KAMBAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-11-28|2027-11-25
    ZAFER KAMBAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-11-28|2027-11-25
    ZAFER KAMBAY|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-24|2027-07-23
    ZAFER KAMBAY|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-02-12|2029-02-11
    ZAFER KAMBAY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-11-28|2027-11-25
    ZAFER KAMBAY|02-Temel İlk Yardım Belgesi|2022-11-28|2027-11-25
    ZAFER KAMBAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-11-28|2027-11-25
    ZEKİ ÖZGÜR SARIKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-18|2027-04-17
    ZEKİ ÖZGÜR SARIKAYA|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|10-İleri Yangınla Mücadele Belgesi|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|02-Temel İlk Yardım Belgesi|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-12|2029-02-11
    ZEKİ ÖZGÜR SARIKAYA|31-Gemi Adamı Cüzdan Belgesi|2024-02-13|2029-02-12
    ADEM GÖDEK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-03|2029-12-02
    ADEM GÖDEK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-03|2029-12-02
    ADEM GÖDEK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-04|2027-03-03
    ADEM GÖDEK|10-İleri Yangınla Mücadele Belgesi|2024-12-03|2029-12-02
    ADEM GÖDEK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-03|2029-12-02
    ADEM GÖDEK|02-Temel İlk Yardım Belgesi|2024-12-03|2029-12-02
    ADEM GÖDEK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-03|2029-12-02
    ADEM GÖDEK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-03|2029-12-02
    ADEM GÖDEK|31-Gemi Adamı Cüzdan Belgesi|2024-12-04|2029-12-03
    AHMET ENGİN TEKE|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-16|2027-12-15
    AHMET ENGİN TEKE|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-21|2030-05-20
    AHMET ENGİN TEKE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-21|2030-05-20
    AHMET ENGİN TEKE|31-Gemi Adamı Cüzdan Belgesi|2025-05-21|2030-05-20
    AYDIN TAFRALI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-03-06|2031-03-05
    AYDIN TAFRALI|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-03-06|2031-03-05
    AYDIN TAFRALI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-08|2027-01-07
    AYDIN TAFRALI|10-İleri Yangınla Mücadele Belgesi|2026-03-06|2031-03-05
    AYDIN TAFRALI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-03-06|2031-03-05
    AYDIN TAFRALI|02-Temel İlk Yardım Belgesi|2026-03-06|2031-03-05
    AYDIN TAFRALI|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-03-06|2031-03-05
    AYDIN TAFRALI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-03-06|2031-03-05
    AYDIN TAFRALI|31-Gemi Adamı Cüzdan Belgesi|2026-03-13|2031-03-12
    EYÜP CÜCÜ|02-Temel İlk Yardım Belgesi|2021-09-16|2026-09-15
    BİROL BAYRAKTAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-30|2027-09-29
    CEM ARIKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-12-11|2028-12-10
    CEM ARIKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-12-11|2028-12-10
    CEM ARIKAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-03|2027-10-02
    CEM ARIKAN|10-İleri Yangınla Mücadele Belgesi|2022-12-13|2027-12-12
    CEM ARIKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-12-11|2028-12-10
    CEM ARIKAN|02-Temel İlk Yardım Belgesi|2023-12-11|2028-12-10
    CEM ARIKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-12-11|2028-12-10
    CEM ARIKAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-12-11|2028-12-10
    CEM ARIKAN|31-Gemi Adamı Cüzdan Belgesi|2022-12-13|2027-12-12
    CEMALETTİN ARSLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-31|2030-07-30
    CEMALETTİN ARSLAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-31|2030-07-30
    CEMALETTİN ARSLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-18|2027-03-17
    CEMALETTİN ARSLAN|10-İleri Yangınla Mücadele Belgesi|2025-09-17|2030-09-16
    CEMALETTİN ARSLAN|02-Temel İlk Yardım Belgesi|2025-07-31|2030-07-30
    CEMALETTİN ARSLAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-31|2030-07-30
    CEMALETTİN ARSLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-17|2030-09-16
    CEMALETTİN ARSLAN|31-Gemi Adamı Cüzdan Belgesi|2025-08-05|2030-08-04
    EYÜP CÜCÜ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-16|2026-09-15
    EYÜP CÜCÜ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-22|2026-09-15
    CENGİZ İLİTLİ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-06|2027-02-05
    CENGİZ İLİTLİ|10-İleri Yangınla Mücadele Belgesi|2023-11-07|2028-11-06
    EYÜP CÜCÜ|31-Gemi Adamı Cüzdan Belgesi|2021-09-17|2026-09-16
    YUSUF ARDA KAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-18|2026-09-17
    İSMAİL BARBAROS ÖZ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-19|2026-09-18
    ERDAL ERTEKİN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-20|2026-09-19
    SİNAN SÜLÜ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-20|2026-09-19
    EMREHAN UZUNOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-04-03|2029-04-02
    EMREHAN UZUNOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-04-03|2029-04-02
    EMREHAN UZUNOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-02|2027-01-01
    EMREHAN UZUNOĞLU|10-İleri Yangınla Mücadele Belgesi|2023-08-24|2028-08-23
    EMREHAN UZUNOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-04-03|2029-04-02
    EMREHAN UZUNOĞLU|02-Temel İlk Yardım Belgesi|2024-04-03|2029-04-02
    EMREHAN UZUNOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-04-03|2029-04-02
    EMREHAN UZUNOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-04-03|2029-04-02
    EMREHAN UZUNOĞLU|31-Gemi Adamı Cüzdan Belgesi|2025-11-03|2030-11-02
    İBRAHİM SARI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-22|2026-09-21
    AHMET KIRAÇ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-22|2026-09-21
    AHMET KIRAÇ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-22|2026-09-21
    AHMET KIRAÇ|10-İleri Yangınla Mücadele Belgesi|2021-09-22|2026-09-21
    AHMET KIRAÇ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-22|2026-09-21
    AHMET KIRAÇ|02-Temel İlk Yardım Belgesi|2021-09-22|2026-09-21
    AHMET KIRAÇ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-22|2026-09-21
    EROL İPEK|31-Gemi Adamı Cüzdan Belgesi|2022-02-02|2027-02-01
    AHMET KIRAÇ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-22|2026-09-21
    AHMET KIRAÇ|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-09-22|2026-09-21
    EYÜP CÜCÜ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-23|2027-01-22
    UĞUR AY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-23|2026-09-22
    UĞUR AY|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-23|2026-09-22
    UĞUR AY|10-İleri Yangınla Mücadele Belgesi|2021-09-23|2026-09-22
    UĞUR AY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-23|2026-09-22
    UĞUR AY|02-Temel İlk Yardım Belgesi|2021-09-23|2026-09-22
    UĞUR AY|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-23|2026-09-22
    UĞUR AY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-23|2026-09-22
    UĞUR AY|31-Gemi Adamı Cüzdan Belgesi|2021-09-23|2026-09-22
    HARUN DEMİRCİ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-15|2027-04-14
    AHMET KIRAÇ|31-Gemi Adamı Cüzdan Belgesi|2021-09-23|2026-09-22
    İBRAHİM AKÇAY|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-24|2026-09-23
    ÇAĞLA ÖZKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-24|2026-09-23
    HARUN DEMİRCİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-05-31|2029-05-30
    EMRE ZOR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-26|2026-09-25
    HAŞİM EMİR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-02|2030-07-01
    HAŞİM EMİR|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-02|2030-07-01
    HAŞİM EMİR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-19|2027-05-18
    HAŞİM EMİR|10-İleri Yangınla Mücadele Belgesi|2025-07-02|2030-07-01
    HAŞİM EMİR|02-Temel İlk Yardım Belgesi|2025-07-02|2030-07-01
    HAŞİM EMİR|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-02|2030-07-01
    HAŞİM EMİR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-07-02|2030-07-01
    HAŞİM EMİR|31-Gemi Adamı Cüzdan Belgesi|2025-07-03|2030-07-02
    AHMET EMRAH YILDIRAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-27|2026-09-26
    AHMET EMRAH YILDIRAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-27|2026-09-26
    HIZIR ÇOLAK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-23|2027-07-22
    AHMET EMRAH YILDIRAN|10-İleri Yangınla Mücadele Belgesi|2021-09-27|2026-09-26
    AHMET EMRAH YILDIRAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-27|2026-09-26
    AHMET EMRAH YILDIRAN|02-Temel İlk Yardım Belgesi|2021-09-27|2026-09-26
    AHMET EMRAH YILDIRAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-27|2026-09-26
    AHMET EMRAH YILDIRAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-27|2026-09-26
    AHMET EMRAH YILDIRAN|31-Gemi Adamı Cüzdan Belgesi|2021-09-27|2026-09-26
    HÜSEYİN TAŞKIN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-28|2027-04-27
    METİN ÇİMEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-21|2030-05-20
    METİN ÇİMEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-11-05|2029-11-04
    ÜMİT İRİTAŞ|25-Kısa Mesafe Telsiz Operatörü Belgesi|2021-09-27|2026-09-26
    METİN ÇİMEN|10-İleri Yangınla Mücadele Belgesi|2024-11-05|2029-11-04
    METİN ÇİMEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-11-05|2029-11-04
    METİN ÇİMEN|02-Temel İlk Yardım Belgesi|2024-11-05|2029-11-04
    METİN ÇİMEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-11-05|2029-11-04
    METİN ÇİMEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-11-05|2029-11-04
    METİN ÇİMEN|31-Gemi Adamı Cüzdan Belgesi|2024-11-07|2029-11-06
    MEVLÜT BAYRAMÇAVUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-10|2027-07-09
    EROL İPEK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-28|2026-09-27
    EROL İPEK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-28|2026-09-27
    MUSTAFA SOYTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2026-02-16|2028-02-15
    EROL İPEK|10-İleri Yangınla Mücadele Belgesi|2021-09-28|2026-09-27
    EROL İPEK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-28|2026-09-27
    EROL İPEK|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-09-28|2026-09-27
    EROL İPEK|02-Temel İlk Yardım Belgesi|2021-09-28|2026-09-27
    EROL İPEK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-28|2026-09-27
    TAYFUN TÜYSÜZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-28|2026-09-27
    MUSTEKİM AVCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-20|2031-02-19
    MUSTEKİM AVCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-20|2031-02-19
    MUSTEKİM AVCI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-08|2027-10-07
    MUSTEKİM AVCI|10-İleri Yangınla Mücadele Belgesi|2026-02-20|2031-02-19
    MUSTEKİM AVCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-20|2031-02-19
    MUSTEKİM AVCI|02-Temel İlk Yardım Belgesi|2026-02-20|2031-02-19
    MUSTEKİM AVCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-20|2031-02-19
    MUSTEKİM AVCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-20|2031-02-19
    MUSTEKİM AVCI|31-Gemi Adamı Cüzdan Belgesi|2026-02-25|2031-02-24
    MUSTAKİM ÇAKMAKÇI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-14|2027-08-13
    MUSTAKİM ÇAKMAKÇI|10-İleri Yangınla Mücadele Belgesi|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|02-Temel İlk Yardım Belgesi|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-07-17|2029-07-16
    MUSTAKİM ÇAKMAKÇI|31-Gemi Adamı Cüzdan Belgesi|2024-07-19|2029-07-18
    NAZİF ALARÇİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-19|2030-09-18
    NAZİF ALARÇİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-19|2030-09-18
    NAZİF ALARÇİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-27|2027-06-26
    NAZİF ALARÇİN|10-İleri Yangınla Mücadele Belgesi|2025-09-19|2030-09-18
    NAZİF ALARÇİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-19|2030-09-18
    TAYFUN TÜYSÜZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-28|2026-09-27
    NAZİF ALARÇİN|02-Temel İlk Yardım Belgesi|2025-09-19|2030-09-18
    NAZİF ALARÇİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-19|2030-09-18
    NAZİF ALARÇİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-19|2030-09-18
    NAZİF ALARÇİN|31-Gemi Adamı Cüzdan Belgesi|2025-09-19|2030-09-18
    NURETTİN KARAL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-02|2027-07-01
    ÖZGÜR SAHTİYAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-16|2027-04-15
    ÖZGÜR SAHTİYAN|10-İleri Yangınla Mücadele Belgesi|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|02-Temel İlk Yardım Belgesi|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-04-15|2030-04-14
    ÖZGÜR SAHTİYAN|31-Gemi Adamı Cüzdan Belgesi|2025-04-16|2030-04-15
    RECEP ASLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-16|2029-10-15
    RECEP ASLAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-16|2029-10-15
    TAYFUN TÜYSÜZ|10-İleri Yangınla Mücadele Belgesi|2021-09-28|2026-09-27
    TAYFUN TÜYSÜZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-28|2026-09-27
    RECEP ASLAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-10-16|2029-10-15
    RECEP ASLAN|02-Temel İlk Yardım Belgesi|2024-10-16|2029-10-15
    RECEP ASLAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-16|2029-10-15
    RECEP ASLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-16|2029-10-15
    RECEP ASLAN|31-Gemi Adamı Cüzdan Belgesi|2024-10-17|2029-10-16
    REFİK İPEK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-14|2030-05-13
    REFİK İPEK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-05-14|2030-05-13
    REFİK İPEK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-14|2027-05-13
    REFİK İPEK|10-İleri Yangınla Mücadele Belgesi|2025-05-14|2030-05-13
    REFİK İPEK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-14|2030-05-13
    REFİK İPEK|02-Temel İlk Yardım Belgesi|2025-05-14|2030-05-13
    REFİK İPEK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-14|2030-05-13
    REFİK İPEK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-14|2030-05-13
    REFİK İPEK|31-Gemi Adamı Cüzdan Belgesi|2025-05-14|2030-05-13
    RESUL ŞEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-03-04|2031-03-03
    RESUL ŞEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-03-04|2031-03-03
    RESUL ŞEN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-19|2027-08-18
    RESUL ŞEN|10-İleri Yangınla Mücadele Belgesi|2026-03-04|2031-03-03
    RESUL ŞEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-03-04|2031-03-03
    RESUL ŞEN|02-Temel İlk Yardım Belgesi|2026-03-04|2031-03-03
    RESUL ŞEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-03-04|2031-03-03
    RESUL ŞEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-03-04|2031-03-03
    RESUL ŞEN|31-Gemi Adamı Cüzdan Belgesi|2026-03-04|2031-03-03
    TAYFUN TÜYSÜZ|02-Temel İlk Yardım Belgesi|2021-09-28|2026-09-27
    SALİH YAVUZ ELMAS|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-25|2029-12-24
    SALİH YAVUZ ELMAS|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-21|2027-05-20
    SALİH YAVUZ ELMAS|10-İleri Yangınla Mücadele Belgesi|2025-05-21|2030-05-20
    SALİH YAVUZ ELMAS|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-21|2030-05-20
    SALİH YAVUZ ELMAS|02-Temel İlk Yardım Belgesi|2025-05-21|2030-05-20
    SALİH YAVUZ ELMAS|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-21|2030-05-20
    SALİH YAVUZ ELMAS|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-21|2030-05-20
    SALİH YAVUZ ELMAS|31-Gemi Adamı Cüzdan Belgesi|2025-05-21|2030-05-20
    TAYFUN TÜYSÜZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-28|2026-09-27
    ŞEREF ZAFER EREK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-30|2030-12-29
    ŞEREF ZAFER EREK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-30|2030-12-29
    ŞEREF ZAFER EREK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-30|2027-12-29
    ŞEREF ZAFER EREK|10-İleri Yangınla Mücadele Belgesi|2025-12-30|2030-12-29
    ŞEREF ZAFER EREK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-30|2030-12-29
    ŞEREF ZAFER EREK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-30|2030-12-29
    ŞEREF ZAFER EREK|31-Gemi Adamı Cüzdan Belgesi|2025-12-31|2030-12-30
    TAYFUN TÜYSÜZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-28|2026-09-27
    EYÜP NAMOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-09-28|2026-09-27
    TAYFUN TÜYSÜZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-21|2027-01-20
    EYÜP NAMOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-28|2026-09-27
    EYÜP NAMOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-28|2026-09-27
    EYÜP NAMOĞLU|02-Temel İlk Yardım Belgesi|2021-09-28|2026-09-27
    EYÜP NAMOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-28|2026-09-27
    EYÜP NAMOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-09-28|2026-09-27
    TAYFUN TÜYSÜZ|31-Gemi Adamı Cüzdan Belgesi|2021-09-29|2026-09-28
    EYÜP NAMOĞLU|31-Gemi Adamı Cüzdan Belgesi|2021-09-29|2026-09-28
    ZEKİ ŞAHİN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-09-30|2026-09-29
    UĞUR AY|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-21|2027-05-20
    BURAK GEYİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-09-30|2026-09-29
    BURAK GEYİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-09-30|2026-09-29
    BURAK GEYİK|02-Temel İlk Yardım Belgesi|2021-09-30|2026-09-29
    BURAK GEYİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-09-30|2026-09-29
    GÖKTUĞ AY|31-Gemi Adamı Cüzdan Belgesi|2021-09-30|2026-09-29
    SALİH KARACA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-01|2026-09-30
    ÜMİT ALKANOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-17|2030-07-16
    ÜMİT ALKANOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-17|2030-07-16
    SALİH KARACA|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-01|2026-09-30
    ÜMİT ALKANOĞLU|10-İleri Yangınla Mücadele Belgesi|2022-11-23|2027-11-22
    ÜMİT ALKANOĞLU|02-Temel İlk Yardım Belgesi|2025-07-17|2030-07-16
    ÜMİT ALKANOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-17|2030-07-16
    ÜMİT ALKANOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-07-17|2030-07-16
    ÜMİT ALKANOĞLU|31-Gemi Adamı Cüzdan Belgesi|2025-07-17|2030-07-16
    YAVUZ SELİM BÜYÜK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-03-10|2031-03-09
    YAVUZ SELİM BÜYÜK|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-03-10|2031-03-09
    YAVUZ SELİM BÜYÜK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-21|2027-03-20
    YAVUZ SELİM BÜYÜK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-03-10|2031-03-09
    YAVUZ SELİM BÜYÜK|02-Temel İlk Yardım Belgesi|2026-03-10|2031-03-09
    YAVUZ SELİM BÜYÜK|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-03-10|2031-03-09
    YAVUZ SELİM BÜYÜK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-03-10|2031-03-09
    YAVUZ SELİM BÜYÜK|31-Gemi Adamı Cüzdan Belgesi|2026-03-11|2031-03-10
    YUSUF ÖZKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-27|2027-01-26
    YUSUF ÖZKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-01-27|2027-01-26
    SALİH KARACA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-01|2026-09-30
    YUSUF ÖZKAN|10-İleri Yangınla Mücadele Belgesi|2022-01-27|2027-01-26
    YUSUF ÖZKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-27|2027-01-26
    YUSUF ÖZKAN|02-Temel İlk Yardım Belgesi|2022-01-27|2027-01-26
    YUSUF ÖZKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-27|2027-01-26
    YUSUF ÖZKAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-27|2027-01-26
    YUSUF ÖZKAN|31-Gemi Adamı Cüzdan Belgesi|2022-01-27|2027-01-26
    ZEKERİYA HAKAN USTA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-02-02|2028-02-01
    ZEKERİYA HAKAN USTA|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-02|2028-02-01
    ZEKERİYA HAKAN USTA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-02|2028-02-01
    ZEKERİYA HAKAN USTA|02-Temel İlk Yardım Belgesi|2023-02-02|2028-02-01
    ZEKERİYA HAKAN USTA|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-02|2028-02-01
    ZEKERİYA HAKAN USTA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-02-02|2028-02-01
    ALİ EMRE KUL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-26|2028-04-25
    ALİ EMRE KUL|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-26|2028-04-25
    SALİH KARACA|02-Temel İlk Yardım Belgesi|2021-10-01|2026-09-30
    ALİ EMRE KUL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-26|2028-04-25
    ALİ EMRE KUL|02-Temel İlk Yardım Belgesi|2023-04-26|2028-04-25
    ALİ EMRE KUL|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-26|2028-04-25
    ALİ EMRE KUL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-26|2028-04-25
    ALİ EMRE KUL|31-Gemi Adamı Cüzdan Belgesi|2023-04-27|2028-04-26
    SALİH KARACA|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-01|2026-09-30
    SALİH KARACA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-01|2026-09-30
    SALİH KARACA|31-Gemi Adamı Cüzdan Belgesi|2021-10-01|2026-09-30
    ERCAN AYDIN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-04|2026-10-03
    SALİH KARACA|10-İleri Yangınla Mücadele Belgesi|2022-11-18|2027-11-17
    MUSTAFA KÖKSEL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-04|2026-10-03
    MUSTAFA KÖKSEL|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-04|2026-10-03
    MUSTAFA KÖKSEL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-04|2026-10-03
    MUSTAFA KÖKSEL|02-Temel İlk Yardım Belgesi|2021-10-04|2026-10-03
    MUSTAFA KÖKSEL|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-04|2026-10-03
    MUSTAFA KÖKSEL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-04|2026-10-03
    MUSTAFA KÖKSEL|31-Gemi Adamı Cüzdan Belgesi|2021-10-04|2026-10-03
    ELİF ÖZLEM SEZEK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-04|2026-10-03
    ELİF ÖZLEM SEZEK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-04|2026-10-03
    ELİF ÖZLEM SEZEK|02-Temel İlk Yardım Belgesi|2021-10-04|2026-10-03
    ELİF ÖZLEM SEZEK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-04|2026-10-03
    ELİF ÖZLEM SEZEK|31-Gemi Adamı Cüzdan Belgesi|2021-10-04|2026-10-03
    HASAN ASLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-04|2026-10-03
    TUNCAY ALBAYRAK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-06|2026-10-05
    TUNCAY ALBAYRAK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-06|2026-10-05
    TUNCAY ALBAYRAK|02-Temel İlk Yardım Belgesi|2021-10-06|2026-10-05
    TUNCAY ALBAYRAK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-06|2026-10-05
    TUNCAY ALBAYRAK|31-Gemi Adamı Cüzdan Belgesi|2021-10-06|2026-10-05
    SONER HUTOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-08|2026-10-07
    HASAN KÜÇÜKİSLAMOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-08|2026-10-07
    HASAN KÜÇÜKİSLAMOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-08|2026-10-07
    HASAN KÜÇÜKİSLAMOĞLU|10-İleri Yangınla Mücadele Belgesi|2021-10-08|2026-10-07
    HASAN KÜÇÜKİSLAMOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-08|2026-10-07
    FAHRİ YAZICI|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-24|2028-06-23
    HASAN KÜÇÜKİSLAMOĞLU|02-Temel İlk Yardım Belgesi|2021-10-08|2026-10-07
    HASAN KÜÇÜKİSLAMOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-08|2026-10-07
    HASAN KÜÇÜKİSLAMOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-08|2026-10-07
    UMUT CAN YİĞİT|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-10|2026-10-09
    HÜSEYİN BATUHAN BAYRAKTAR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-10|2026-10-09
    HAMZA SAFER ÇINAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-02-14|2027-02-13
    HAMZA SAFER ÇINAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-14|2027-02-13
    MUSTAFA BURAK ÖZTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-10|2026-10-09
    HAMZA SAFER ÇINAR|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-12-27|2029-12-26
    HAMZA SAFER ÇINAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-14|2027-02-13
    HAMZA SAFER ÇINAR|02-Temel İlk Yardım Belgesi|2022-02-14|2027-02-13
    HAMZA SAFER ÇINAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-14|2027-02-13
    HAMZA SAFER ÇINAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-14|2027-02-13
    HAMZA SAFER ÇINAR|31-Gemi Adamı Cüzdan Belgesi|2022-02-14|2027-02-13
    SALİH ACUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-11|2026-10-10
    SALİH ACUN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-11|2026-10-10
    SALİH ACUN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-10-11|2026-10-10
    SALİH ACUN|10-İleri Yangınla Mücadele Belgesi|2021-10-11|2026-10-10
    SALİH ACUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-11|2026-10-10
    SALİH ACUN|02-Temel İlk Yardım Belgesi|2021-10-11|2026-10-10
    SALİH ACUN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-11|2026-10-10
    SALİH ACUN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|02-Temel İlk Yardım Belgesi|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-11|2026-10-10
    RECEPALİ BAYRAKTAR|31-Gemi Adamı Cüzdan Belgesi|2021-10-11|2026-10-10
    HÜSEYİN KÖMÜRCÜ|31-Gemi Adamı Cüzdan Belgesi|2021-10-11|2026-10-10
    HASAN KÜÇÜKİSLAMOĞLU|31-Gemi Adamı Cüzdan Belgesi|2021-10-11|2026-10-10
    RECEP MICIK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-12|2026-10-11
    RECEP MICIK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-12|2026-10-11
    RECEP MICIK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-10-12|2026-10-11
    RECEP MICIK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-12|2026-10-11
    RECEP MICIK|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-10-12|2026-10-11
    RECEP MICIK|02-Temel İlk Yardım Belgesi|2021-10-12|2026-10-11
    RECEP MICIK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-12|2026-10-11
    SELİM BEKAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-03-16|2027-03-15
    SELİM BEKAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-03-16|2027-03-15
    SELİM BEKAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-04|2027-06-03
    SELİM BEKAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-03-16|2027-03-15
    SELİM BEKAR|02-Temel İlk Yardım Belgesi|2022-03-16|2027-03-15
    SELİM BEKAR|20-Seyir Vardiyası Tutma Belgesi|2022-03-16|2027-03-15
    SELİM BEKAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-03-16|2027-03-15
    SELİM BEKAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-03-16|2027-03-15
    SELİM BEKAR|31-Gemi Adamı Cüzdan Belgesi|2022-03-16|2027-03-15
    ONUR BAYAT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-01-03|2030-01-02
    ONUR BAYAT|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-01-03|2030-01-02
    ONUR BAYAT|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-19|2028-01-18
    ONUR BAYAT|10-İleri Yangınla Mücadele Belgesi|2025-01-03|2030-01-02
    ONUR BAYAT|02-Temel İlk Yardım Belgesi|2025-01-03|2030-01-02
    ONUR BAYAT|20-Seyir Vardiyası Tutma Belgesi|2025-01-03|2030-01-02
    ONUR BAYAT|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-01-03|2030-01-02
    ONUR BAYAT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-01-03|2030-01-02
    ONUR BAYAT|31-Gemi Adamı Cüzdan Belgesi|2023-10-30|2028-10-29
    ADNAN KELEŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-02|2027-07-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-26|2028-06-25
    İSMAİL HAKKI BAYRAMÇAVUŞ|10-İleri Yangınla Mücadele Belgesi|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|02-Temel İlk Yardım Belgesi|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-01-02|2028-01-01
    İSMAİL HAKKI BAYRAMÇAVUŞ|31-Gemi Adamı Cüzdan Belgesi|2023-01-02|2028-01-01
    İSMET PERİŞANOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-27|2031-02-26
    İSMET PERİŞANOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-27|2031-02-26
    İSMET PERİŞANOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-12|2027-11-11
    İSMET PERİŞANOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-27|2031-02-26
    İSMET PERİŞANOĞLU|02-Temel İlk Yardım Belgesi|2026-02-27|2031-02-26
    İSMET PERİŞANOĞLU|20-Seyir Vardiyası Tutma Belgesi|2026-02-27|2031-02-26
    İSMET PERİŞANOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-27|2031-02-26
    İSMET PERİŞANOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-27|2031-02-26
    RECEP MICIK|31-Gemi Adamı Cüzdan Belgesi|2021-10-12|2026-10-11
    İSMET PERİŞANOĞLU|31-Gemi Adamı Cüzdan Belgesi|2026-03-02|2031-03-01
    KORAY KARACA|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-10-08|2030-10-07
    KORAY KARACA|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2024-02-19|2029-02-18
    KORAY KARACA|02-Temel İlk Yardım Belgesi|2025-10-08|2030-10-07
    KORAY KARACA|20-Seyir Vardiyası Tutma Belgesi|2025-10-08|2030-10-07
    KORAY KARACA|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-10-08|2030-10-07
    KORAY KARACA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-10-08|2030-10-07
    KORAY KARACA|30-Gemi Güvenlik Zabiti|2025-10-08|2030-10-07
    KORAY KARACA|31-Gemi Adamı Cüzdan Belgesi|2025-10-08|2030-10-07
    MUSTAFA ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-11-18|2030-11-17
    MUSTAFA ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-11-18|2030-11-17
    SALİH ACUN|31-Gemi Adamı Cüzdan Belgesi|2021-10-13|2026-10-12
    MUSTAFA ÖZTÜRK|10-İleri Yangınla Mücadele Belgesi|2025-11-18|2030-11-17
    MUSTAFA ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-11-18|2030-11-17
    MUSTAFA ÖZTÜRK|02-Temel İlk Yardım Belgesi|2025-11-18|2030-11-17
    MUSTAFA ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-11-18|2030-11-17
    MUSTAFA ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-11-18|2030-11-17
    MURAT KESKİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-14|2026-10-13
    MUSTAFA ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2025-11-18|2030-11-17
    HAKAN AKSU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-22|2030-05-21
    HAKAN AKSU|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-15|2028-05-14
    MURAT KESKİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-14|2026-10-13
    HAKAN AKSU|10-İleri Yangınla Mücadele Belgesi|2023-05-15|2028-05-14
    HAKAN AKSU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-15|2028-05-14
    HAKAN AKSU|02-Temel İlk Yardım Belgesi|2023-05-15|2028-05-14
    HAKAN AKSU|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-15|2028-05-14
    HAKAN AKSU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-05-15|2028-05-14
    HAKAN AKSU|31-Gemi Adamı Cüzdan Belgesi|2023-05-15|2028-05-14
    HALİT HİLMİ SÜSLÜ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-19|2027-12-18
    HALİT HİLMİ SÜSLÜ|10-İleri Yangınla Mücadele Belgesi|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|02-Temel İlk Yardım Belgesi|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-12-24|2030-12-23
    HALİT HİLMİ SÜSLÜ|31-Gemi Adamı Cüzdan Belgesi|2025-12-26|2030-12-25
    RAMAZAN ÇELİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|10-İleri Yangınla Mücadele Belgesi|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|02-Temel İlk Yardım Belgesi|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-11-20|2029-11-19
    RAMAZAN ÇELİK|31-Gemi Adamı Cüzdan Belgesi|2026-02-02|2031-02-01
    MURAT KESKİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-14|2026-10-13
    MURAT KESKİN|02-Temel İlk Yardım Belgesi|2021-10-14|2026-10-13
    AHMET KIRAÇ|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-05|2028-06-04
    MURAT KESKİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-14|2026-10-13
    MURAT KESKİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-14|2026-10-13
    MURAT KESKİN|31-Gemi Adamı Cüzdan Belgesi|2021-10-14|2026-10-13
    ÜMİT EKİZ|31-Gemi Adamı Cüzdan Belgesi|2021-10-14|2026-10-13
    FATİH YAŞAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-15|2026-10-14
    FATİH YAŞAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-15|2026-10-14
    FATİH YAŞAR|10-İleri Yangınla Mücadele Belgesi|2021-10-15|2026-10-14
    ERKAN BÜYÜKASLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-23|2030-06-22
    ERKAN BÜYÜKASLAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-06-23|2030-06-22
    ERKAN BÜYÜKASLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-06|2027-08-05
    ERKAN BÜYÜKASLAN|10-İleri Yangınla Mücadele Belgesi|2025-07-23|2030-07-22
    ERKAN BÜYÜKASLAN|02-Temel İlk Yardım Belgesi|2025-06-23|2030-06-22
    ERKAN BÜYÜKASLAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-06-23|2030-06-22
    ERKAN BÜYÜKASLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-06-23|2030-06-22
    ERKAN BÜYÜKASLAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-05-23|2030-05-22
    ERKAN BÜYÜKASLAN|31-Gemi Adamı Cüzdan Belgesi|2025-06-23|2030-06-22
    SÜLEYMAN NAROĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|10-İleri Yangınla Mücadele Belgesi|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|02-Temel İlk Yardım Belgesi|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|24-Genel Telsiz Operatörü (GOC) Belgesi|2025-04-30|2030-04-29
    SÜLEYMAN NAROĞLU|31-Gemi Adamı Cüzdan Belgesi|2025-05-22|2030-05-21
    FATİH YAŞAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-15|2026-10-14
    FATİH YAŞAR|02-Temel İlk Yardım Belgesi|2021-10-15|2026-10-14
    FATİH YAŞAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-15|2026-10-14
    FATİH YAŞAR|31-Gemi Adamı Cüzdan Belgesi|2021-10-21|2026-10-14
    AHMET ZAFER DİNÇ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-15|2026-10-14
    AHMET ZAFER DİNÇ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-15|2026-10-14
    AHMET ZAFER DİNÇ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-15|2026-10-14
    KEMAL KOCATÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-03-13|2031-03-12
    KEMAL KOCATÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-03-13|2031-03-12
    AHMET ZAFER DİNÇ|02-Temel İlk Yardım Belgesi|2021-10-15|2026-10-14
    KEMAL KOCATÜRK|10-İleri Yangınla Mücadele Belgesi|2026-03-13|2031-03-12
    KEMAL KOCATÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-03-13|2031-03-12
    KEMAL KOCATÜRK|02-Temel İlk Yardım Belgesi|2026-03-13|2031-03-12
    KEMAL KOCATÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-03-13|2031-03-12
    KEMAL KOCATÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-03-13|2031-03-12
    KEMAL KOCATÜRK|31-Gemi Adamı Cüzdan Belgesi|2026-03-16|2031-03-15
    MEHMET MURAT ÇİÇEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-25|2029-10-24
    MEHMET MURAT ÇİÇEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-25|2029-10-24
    MEHMET MURAT ÇİÇEN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-23|2027-07-22
    MEHMET MURAT ÇİÇEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-10-25|2029-10-24
    MEHMET MURAT ÇİÇEN|02-Temel İlk Yardım Belgesi|2024-10-25|2029-10-24
    MEHMET MURAT ÇİÇEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-25|2029-10-24
    MEHMET MURAT ÇİÇEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-25|2029-10-24
    MEHMET MURAT ÇİÇEN|31-Gemi Adamı Cüzdan Belgesi|2024-10-25|2029-10-24
    SALİH BÜLBÜL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-12-20|2027-12-19
    SALİH BÜLBÜL|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-12-20|2027-12-19
    SALİH BÜLBÜL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-21|2027-02-20
    SALİH BÜLBÜL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-12-20|2027-12-19
    SALİH BÜLBÜL|02-Temel İlk Yardım Belgesi|2022-12-20|2027-12-19
    SALİH BÜLBÜL|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-12-20|2027-12-19
    SALİH BÜLBÜL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-12-20|2027-12-19
    SALİH BÜLBÜL|31-Gemi Adamı Cüzdan Belgesi|2022-12-21|2027-12-20
    AHMET ZAFER DİNÇ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-15|2026-10-14
    AHMET ZAFER DİNÇ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-15|2026-10-14
    AHMET ZAFER DİNÇ|31-Gemi Adamı Cüzdan Belgesi|2021-10-15|2026-10-14
    CİHAT SARIHASAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-15|2026-10-14
    CİHAT SARIHASAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-15|2026-10-14
    CİHAT SARIHASAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-15|2026-10-14
    CİHAT SARIHASAN|02-Temel İlk Yardım Belgesi|2021-10-15|2026-10-14
    CİHAT SARIHASAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-15|2026-10-14
    MUHAMMED ALİ AVCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-10-15|2030-10-14
    MUHAMMED ALİ AVCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-10-15|2030-10-14
    CİHAT SARIHASAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-15|2026-10-14
    MUHAMMED ALİ AVCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-10-15|2030-10-14
    MUHAMMED ALİ AVCI|02-Temel İlk Yardım Belgesi|2025-10-15|2030-10-14
    MUHAMMED ALİ AVCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-10-15|2030-10-14
    MUHAMMED ALİ AVCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-10-15|2030-10-14
    MUHAMMED ALİ AVCI|31-Gemi Adamı Cüzdan Belgesi|2025-10-15|2030-10-14
    FATİH SAMED GÖRMÜŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-01-14|2030-01-13
    FATİH SAMED GÖRMÜŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-01-14|2030-01-13
    CİHAT SARIHASAN|31-Gemi Adamı Cüzdan Belgesi|2021-10-15|2026-10-14
    FATİH SAMED GÖRMÜŞ|10-İleri Yangınla Mücadele Belgesi|2025-01-14|2030-01-13
    FATİH SAMED GÖRMÜŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-01-14|2030-01-13
    FATİH SAMED GÖRMÜŞ|02-Temel İlk Yardım Belgesi|2025-01-14|2030-01-13
    FATİH SAMED GÖRMÜŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-01-14|2030-01-13
    FATİH SAMED GÖRMÜŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-01-14|2030-01-13
    FATİH SAMED GÖRMÜŞ|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-02-26|2029-02-25
    FATİH SAMED GÖRMÜŞ|31-Gemi Adamı Cüzdan Belgesi|2025-01-16|2030-01-15
    FURKAN ÖKSÜZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-22|2029-08-21
    FURKAN ÖKSÜZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-23|2028-05-22
    FURKAN ÖKSÜZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-19|2027-09-18
    FURKAN ÖKSÜZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-23|2028-05-22
    FURKAN ÖKSÜZ|02-Temel İlk Yardım Belgesi|2023-05-23|2028-05-22
    FURKAN ÖKSÜZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-23|2028-05-22
    FURKAN ÖKSÜZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-11-20|2028-11-19
    FURKAN ÖKSÜZ|31-Gemi Adamı Cüzdan Belgesi|2023-10-06|2028-10-05
    EYÜP GÜNDOĞDU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-03-20|2028-03-19
    EYÜP GÜNDOĞDU|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-03-20|2028-03-19
    ALİ EMRE KUL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-16|2026-10-15
    EYÜP GÜNDOĞDU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-03-20|2028-03-19
    EYÜP GÜNDOĞDU|02-Temel İlk Yardım Belgesi|2023-03-20|2028-03-19
    EYÜP GÜNDOĞDU|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-03-20|2028-03-19
    EYÜP GÜNDOĞDU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-30|2028-04-29
    EYÜP GÜNDOĞDU|31-Gemi Adamı Cüzdan Belgesi|2023-03-20|2028-03-19
    SÜLEYMAN AYDIN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-16|2026-10-15
    ÖZGÜR TEKİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-04-16|2029-04-15
    ÖZGÜR TEKİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-04-16|2029-04-15
    ÖZGÜR TEKİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-12|2027-02-11
    ÖZGÜR TEKİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-04-16|2029-04-15
    ÖZGÜR TEKİN|02-Temel İlk Yardım Belgesi|2024-04-16|2029-04-15
    ÖZGÜR TEKİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-04-16|2029-04-15
    ÖZGÜR TEKİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-04-16|2029-04-15
    ÖZGÜR TEKİN|31-Gemi Adamı Cüzdan Belgesi|2022-02-21|2027-02-20
    AHMET DENİZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-03-11|2027-03-10
    AHMET DENİZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-03-11|2027-03-10
    AHMET DENİZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-03-11|2027-03-10
    AHMET DENİZ|02-Temel İlk Yardım Belgesi|2022-03-11|2027-03-10
    AHMET DENİZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-03-11|2027-03-10
    AHMET DENİZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-03-11|2027-03-10
    AHMET DENİZ|31-Gemi Adamı Cüzdan Belgesi|2022-10-11|2027-10-10
    CUMHUR KENARCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-01-24|2029-01-23
    CUMHUR KENARCI|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-28|2028-01-27
    CUMHUR KENARCI|10-İleri Yangınla Mücadele Belgesi|2024-01-24|2029-01-23
    CUMHUR KENARCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-01-24|2029-01-23
    CUMHUR KENARCI|02-Temel İlk Yardım Belgesi|2024-01-24|2029-01-23
    CUMHUR KENARCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-01-24|2029-01-23
    CUMHUR KENARCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-02-15|2029-02-14
    EMRAH TOKER|01-Güverte Kısmı Gemi Adamı Belgesi 1|2025-07-19|2030-07-18
    EMRAH TOKER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-09|2030-07-08
    EMRAH TOKER|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-09|2030-07-08
    EMRAH TOKER|10-İleri Yangınla Mücadele Belgesi|2023-07-24|2028-07-23
    EMRAH TOKER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-07-09|2030-07-08
    EMRAH TOKER|02-Temel İlk Yardım Belgesi|2025-07-09|2030-07-08
    EMRAH TOKER|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-09|2030-07-08
    EMRAH TOKER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-07-09|2030-07-08
    EMRAH TOKER|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-01-29|2029-01-28
    EMRAH TOKER|31-Gemi Adamı Cüzdan Belgesi|2025-07-14|2030-07-13
    ERDEM KIRBAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-05-08|2028-05-07
    ERDEM KIRBAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-08|2028-05-07
    ERDEM KIRBAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-08|2028-05-07
    ERDEM KIRBAŞ|02-Temel İlk Yardım Belgesi|2023-05-08|2028-05-07
    ERDEM KIRBAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-08|2028-05-07
    ERDEM KIRBAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-03-01|2027-02-28
    ERDEM KIRBAŞ|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-10-30|2028-10-29
    ERDEM KIRBAŞ|31-Gemi Adamı Cüzdan Belgesi|2022-03-01|2027-02-28
    SEMİH AKSOY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-18|2029-09-17
    SEMİH AKSOY|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-18|2029-09-17
    SEMİH AKSOY|10-İleri Yangınla Mücadele Belgesi|2024-09-18|2029-09-17
    SEMİH AKSOY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-18|2029-09-17
    SEMİH AKSOY|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2024-09-18|2029-09-17
    SEMİH AKSOY|02-Temel İlk Yardım Belgesi|2024-09-18|2029-09-17
    SEMİH AKSOY|20-Seyir Vardiyası Tutma Belgesi|2024-09-18|2029-09-17
    SEMİH AKSOY|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-18|2029-09-17
    SEMİH AKSOY|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-12-15|2028-12-14
    SEMİH AKSOY|31-Gemi Adamı Cüzdan Belgesi|2024-09-19|2029-09-18
    YUSUF İSLAM SAĞIR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-11|2027-07-10
    YUSUF İSLAM SAĞIR|10-İleri Yangınla Mücadele Belgesi|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2022-02-07|2027-02-06
    YUSUF İSLAM SAĞIR|02-Temel İlk Yardım Belgesi|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|20-Seyir Vardiyası Tutma Belgesi|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-08|2028-02-07
    YUSUF İSLAM SAĞIR|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-02-07|2027-02-06
    YUSUF İSLAM SAĞIR|31-Gemi Adamı Cüzdan Belgesi|2022-06-06|2027-06-05
    EMRAH DEMİRKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|02-Temel İlk Yardım Belgesi|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-06-07|2027-06-06
    EMRAH DEMİRKAN|31-Gemi Adamı Cüzdan Belgesi|2022-06-07|2027-06-06
    MEHMET MAHMUTOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-07-19|2029-07-18
    MEHMET MAHMUTOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-07-19|2029-07-18
    MEHMET MAHMUTOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-23|2028-01-22
    MEHMET MAHMUTOĞLU|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-03-07|2027-03-06
    MEHMET MAHMUTOĞLU|10-İleri Yangınla Mücadele Belgesi|2024-07-19|2029-07-18
    MEHMET MAHMUTOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-07-19|2029-07-18
    MEHMET MAHMUTOĞLU|02-Temel İlk Yardım Belgesi|2024-07-19|2029-07-18
    MEHMET MAHMUTOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-07-19|2029-07-18
    MEHMET MAHMUTOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-03-07|2027-03-06
    MEHMET MAHMUTOĞLU|31-Gemi Adamı Cüzdan Belgesi|2022-03-07|2027-03-06
    TURGAY MUŞDAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-05|2030-09-04
    TURGAY MUŞDAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-05|2030-09-04
    MERT AKAR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-18|2026-10-17
    TURGAY MUŞDAL|10-İleri Yangınla Mücadele Belgesi|2025-09-05|2030-09-04
    TURGAY MUŞDAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-09-15|2028-09-14
    TURGAY MUŞDAL|02-Temel İlk Yardım Belgesi|2025-09-05|2030-09-04
    TURGAY MUŞDAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-05|2030-09-04
    TURGAY MUŞDAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-05|2030-09-04
    TURGAY MUŞDAL|31-Gemi Adamı Cüzdan Belgesi|2025-09-10|2030-09-09
    KEREM KAPTANAĞASI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-02-03|2027-02-02
    KEREM KAPTANAĞASI|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-03|2027-02-02
    KEREM KAPTANAĞASI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-20|2027-05-19
    KEREM KAPTANAĞASI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-03|2027-02-02
    KEREM KAPTANAĞASI|02-Temel İlk Yardım Belgesi|2022-02-03|2027-02-02
    KEREM KAPTANAĞASI|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-03|2027-02-02
    KEREM KAPTANAĞASI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-03|2027-02-02
    KEREM KAPTANAĞASI|31-Gemi Adamı Cüzdan Belgesi|2022-02-03|2027-02-02
    EMRE BEKAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-02-24|2028-02-23
    EMRE BEKAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-24|2028-02-23
    EMRE BEKAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-21|2027-08-20
    EMRE BEKAR|10-İleri Yangınla Mücadele Belgesi|2022-06-10|2027-06-09
    EMRE BEKAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-24|2028-02-23
    EMRE BEKAR|02-Temel İlk Yardım Belgesi|2023-02-24|2028-02-23
    EMRE BEKAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-24|2028-02-23
    EMRE BEKAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-03-31|2027-03-30
    EMRE BEKAR|31-Gemi Adamı Cüzdan Belgesi|2025-02-11|2030-02-10
    UĞUR ÇÜLBAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-11-27|2028-11-26
    UĞUR ÇÜLBAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-11-27|2028-11-26
    UĞUR ÇÜLBAN|02-Temel İlk Yardım Belgesi|2023-11-27|2028-11-26
    UĞUR ÇÜLBAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-21|2030-05-20
    UĞUR ÇÜLBAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-11-27|2028-11-26
    UĞUR ÇÜLBAN|31-Gemi Adamı Cüzdan Belgesi|2023-11-27|2028-11-26
    RAMAZAN ERKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-12-11|2028-12-10
    RAMAZAN ERKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-12-11|2028-12-10
    RAMAZAN ERKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-12-11|2028-12-10
    RAMAZAN ERKAN|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2023-12-11|2028-12-10
    RAMAZAN ERKAN|02-Temel İlk Yardım Belgesi|2023-12-11|2028-12-10
    RAMAZAN ERKAN|20-Seyir Vardiyası Tutma Belgesi|2023-12-11|2028-12-10
    RAMAZAN ERKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-12-11|2028-12-10
    RAMAZAN ERKAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-12-11|2028-12-10
    RAMAZAN ERKAN|30-Gemi Güvenlik Zabiti|2023-12-11|2028-12-10
    RAMAZAN ERKAN|31-Gemi Adamı Cüzdan Belgesi|2023-12-13|2028-12-12
    HÜSEYİN KÖMÜRCÜ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-10-10|2027-10-09
    HÜSEYİN KÖMÜRCÜ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-10-10|2027-10-09
    HÜSEYİN KÖMÜRCÜ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-18|2027-11-17
    HÜSEYİN KÖMÜRCÜ|10-İleri Yangınla Mücadele Belgesi|2023-10-10|2028-10-09
    HÜSEYİN KÖMÜRCÜ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-10-10|2027-10-09
    HÜSEYİN KÖMÜRCÜ|02-Temel İlk Yardım Belgesi|2022-10-10|2027-10-09
    HÜSEYİN KÖMÜRCÜ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-10-10|2027-10-09
    HÜSEYİN KÖMÜRCÜ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-10-10|2027-10-09
    HÜSEYİN KÖMÜRCÜ|31-Gemi Adamı Cüzdan Belgesi|2022-10-11|2027-10-10
    ÜMİT YAVUZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-05-25|2027-05-24
    ÜMİT YAVUZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-05-25|2027-05-24
    ÜMİT YAVUZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-12|2027-06-11
    ÜMİT YAVUZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-05-25|2027-05-24
    ÜMİT YAVUZ|02-Temel İlk Yardım Belgesi|2022-05-25|2027-05-24
    ÜMİT YAVUZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-05-25|2027-05-24
    ÜMİT YAVUZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-05-25|2027-05-24
    ÜMİT YAVUZ|31-Gemi Adamı Cüzdan Belgesi|2022-05-25|2027-05-24
    BEHLÜL ÇEKEN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-08|2027-05-07
    BEHLÜL ÇEKEN|10-İleri Yangınla Mücadele Belgesi|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|02-Temel İlk Yardım Belgesi|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-24|2027-01-23
    BEHLÜL ÇEKEN|31-Gemi Adamı Cüzdan Belgesi|2023-07-28|2028-07-27
    OSMAN KOÇAK|20-Seyir Vardiyası Tutma Belgesi|2021-10-18|2026-10-17
    HIZIR ÇOLAK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-19|2026-10-18
    HIZIR ÇOLAK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-19|2026-10-18
    HIZIR ÇOLAK|10-İleri Yangınla Mücadele Belgesi|2021-10-19|2026-10-18
    HIZIR ÇOLAK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-19|2026-10-18
    HIZIR ÇOLAK|02-Temel İlk Yardım Belgesi|2021-10-19|2026-10-18
    HIZIR ÇOLAK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-19|2026-10-18
    MURAT CAN ŞAFFAK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-11|2027-04-10
    CÜNEYT HAMARAT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-22|2030-05-20
    CÜNEYT HAMARAT|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-12-20|2027-12-19
    HIZIR ÇOLAK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-19|2026-10-18
    CÜNEYT HAMARAT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-12-20|2027-12-19
    CÜNEYT HAMARAT|02-Temel İlk Yardım Belgesi|2022-12-20|2027-12-19
    CÜNEYT HAMARAT|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-12-20|2027-12-19
    CÜNEYT HAMARAT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-06-12|2028-06-11
    CÜNEYT HAMARAT|31-Gemi Adamı Cüzdan Belgesi|2022-12-21|2027-12-20
    TAHA KARTOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-02-09|2027-02-08
    TAHA KARTOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-09|2027-02-08
    TAHA KARTOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-22|2027-05-21
    TAHA KARTOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-09|2027-02-08
    TAHA KARTOĞLU|02-Temel İlk Yardım Belgesi|2022-02-09|2027-02-08
    TAHA KARTOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-09|2027-02-08
    TAHA KARTOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-09|2027-02-08
    TAHA KARTOĞLU|24-Genel Telsiz Operatörü (GOC) Belgesi|2022-02-08|2027-02-07
    TAHA KARTOĞLU|31-Gemi Adamı Cüzdan Belgesi|2022-02-09|2027-02-07
    FURKAN ÇETİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-10-10|2027-10-09
    FURKAN ÇETİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-10-10|2027-10-09
    FURKAN ÇETİN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-10|2027-10-09
    FURKAN ÇETİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-10-10|2027-10-09
    FURKAN ÇETİN|02-Temel İlk Yardım Belgesi|2022-10-10|2027-10-09
    FURKAN ÇETİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-10-10|2027-10-09
    FURKAN ÇETİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-10-10|2027-10-09
    FURKAN ÇETİN|31-Gemi Adamı Cüzdan Belgesi|2022-10-10|2027-10-09
    NAZİF ALARÇİN|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|02-Temel İlk Yardım Belgesi|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-19|2026-10-18
    RAMAZAN SANDIKÇI|31-Gemi Adamı Cüzdan Belgesi|2021-10-19|2026-10-18
    SADIK İNANDI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-07-03|2030-07-02
    SADIK İNANDI|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-07-03|2030-07-02
    SADIK İNANDI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-26|2027-02-25
    SADIK İNANDI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-07-03|2030-07-02
    SADIK İNANDI|02-Temel İlk Yardım Belgesi|2025-07-03|2030-07-02
    SADIK İNANDI|20-Seyir Vardiyası Tutma Belgesi|2024-11-17|2029-11-16
    SADIK İNANDI|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-07-03|2030-07-02
    SADIK İNANDI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-31|2027-01-30
    SADIK İNANDI|31-Gemi Adamı Cüzdan Belgesi|2025-07-03|2030-07-02
    MUSA KAŞIKÇI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-06-12|2029-06-11
    MUSA KAŞIKÇI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-06-12|2029-06-11
    MUSA KAŞIKÇI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-24|2027-06-23
    MUSA KAŞIKÇI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-06-12|2029-06-11
    MUSA KAŞIKÇI|02-Temel İlk Yardım Belgesi|2024-06-12|2029-06-11
    MUSA KAŞIKÇI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-06-12|2029-06-11
    MUSA KAŞIKÇI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-06-12|2029-06-11
    MUSA KAŞIKÇI|31-Gemi Adamı Cüzdan Belgesi|2024-06-12|2029-06-11
    RECEP HARDAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-28|2030-09-22
    RECEP HARDAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-28|2030-09-22
    HIZIR ÇOLAK|31-Gemi Adamı Cüzdan Belgesi|2021-10-20|2026-10-19
    RECEP HARDAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-28|2030-09-22
    RECEP HARDAL|02-Temel İlk Yardım Belgesi|2025-09-28|2030-09-22
    RECEP HARDAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-28|2030-09-22
    RECEP HARDAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-28|2030-09-22
    RECEP HARDAL|31-Gemi Adamı Cüzdan Belgesi|2025-09-28|2030-09-22
    MERT CEMİL AKÇAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-17|2029-10-16
    MERT CEMİL AKÇAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-24|2029-09-23
    DOĞUKAN KARATOP|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-20|2026-10-19
    MERT CEMİL AKÇAY|10-İleri Yangınla Mücadele Belgesi|2023-09-20|2028-09-19
    MERT CEMİL AKÇAY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-24|2029-09-23
    MERT CEMİL AKÇAY|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2023-09-20|2028-09-19
    MERT CEMİL AKÇAY|02-Temel İlk Yardım Belgesi|2024-09-24|2029-09-23
    MERT CEMİL AKÇAY|20-Seyir Vardiyası Tutma Belgesi|2024-09-24|2029-09-23
    MERT CEMİL AKÇAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-24|2029-09-23
    MERT CEMİL AKÇAY|31-Gemi Adamı Cüzdan Belgesi|2024-10-02|2029-10-01
    ERDAL ERTEKİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-21|2029-08-20
    ERDAL ERTEKİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-30|2028-05-29
    İLKER HAYTA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-22|2026-10-21
    ERDAL ERTEKİN|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-06-09|2028-06-08
    ERDAL ERTEKİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-30|2028-05-29
    ERDAL ERTEKİN|02-Temel İlk Yardım Belgesi|2023-05-30|2028-05-29
    ERDAL ERTEKİN|20-Seyir Vardiyası Tutma Belgesi|2024-08-21|2029-08-20
    ERDAL ERTEKİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-30|2028-05-29
    ERDAL ERTEKİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-05-30|2028-05-29
    ERDAL ERTEKİN|31-Gemi Adamı Cüzdan Belgesi|2023-05-30|2028-05-29
    ERTAN ÇETİNKAPLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-04-25|2030-04-24
    ERTAN ÇETİNKAPLAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-04-25|2030-04-24
    ERTAN ÇETİNKAPLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-15|2027-08-14
    ERTAN ÇETİNKAPLAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-25|2030-04-24
    ERTAN ÇETİNKAPLAN|02-Temel İlk Yardım Belgesi|2025-04-25|2030-04-24
    ERTAN ÇETİNKAPLAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-25|2030-04-24
    ERTAN ÇETİNKAPLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-04-25|2030-04-24
    ERTAN ÇETİNKAPLAN|31-Gemi Adamı Cüzdan Belgesi|2022-05-20|2027-05-19
    TUNCAY ALBAYRAK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-11-22|2029-11-21
    MUSTAFA ÖZTÜRK|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-10-22|2026-10-21
    MURAT GÜLMEZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-22|2026-10-21
    İSMAİL YALÇINKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-23|2026-10-22
    MUHAMMET PAPAKER|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-24|2026-10-23
    TUNCAY ALBAYRAK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-01-23|2028-01-22
    MEHMET DURMUŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-25|2026-10-24
    MERT BERKAY BAYDERE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-05-31|2027-05-30
    MERT BERKAY BAYDERE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-05-31|2027-05-30
    EYÜP HIZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-10-02|2028-10-01
    EYÜP HIZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-10-02|2028-10-01
    EYÜP HIZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-10-02|2028-10-01
    EYÜP HIZ|02-Temel İlk Yardım Belgesi|2023-10-02|2028-10-01
    EYÜP HIZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-10-02|2028-10-01
    EYÜP HIZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-08-02|2029-08-01
    EYÜP HIZ|31-Gemi Adamı Cüzdan Belgesi|2023-10-24|2028-10-23
    TALHA CİVELEK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-19|2029-09-18
    TALHA CİVELEK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-19|2029-09-18
    TALHA CİVELEK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-20|2027-11-19
    TALHA CİVELEK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-19|2029-09-18
    TALHA CİVELEK|02-Temel İlk Yardım Belgesi|2024-09-19|2029-09-18
    TALHA CİVELEK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-19|2029-09-18
    TALHA CİVELEK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-19|2029-09-18
    TALHA CİVELEK|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-09-06|2029-09-05
    TALHA CİVELEK|31-Gemi Adamı Cüzdan Belgesi|2024-09-20|2029-09-19
    MUSTAFA ÖZDEMİR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-21|2027-05-20
    MUSTAFA ÖZDEMİR|10-İleri Yangınla Mücadele Belgesi|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|02-Temel İlk Yardım Belgesi|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-04-14|2028-04-13
    MUSTAFA ÖZDEMİR|31-Gemi Adamı Cüzdan Belgesi|2023-04-18|2028-04-13
    UMUT CAN YİĞİT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-22|2029-12-21
    UMUT CAN YİĞİT|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-08|2027-06-07
    MEHMET DURMUŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-25|2026-10-24
    UMUT CAN YİĞİT|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-11-07|2027-11-06
    UMUT CAN YİĞİT|10-İleri Yangınla Mücadele Belgesi|2022-12-30|2027-12-29
    UMUT CAN YİĞİT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-06-08|2027-06-07
    UMUT CAN YİĞİT|02-Temel İlk Yardım Belgesi|2022-06-08|2027-06-07
    UMUT CAN YİĞİT|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-08|2027-06-07
    UMUT CAN YİĞİT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-11-07|2027-11-06
    UMUT CAN YİĞİT|31-Gemi Adamı Cüzdan Belgesi|2024-12-23|2029-12-22
    ÖMER KETENCİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-20|2030-05-19
    ÖMER KETENCİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-05-20|2030-05-19
    ÖMER KETENCİ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-05|2027-05-04
    ÖMER KETENCİ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-03-21|2028-03-20
    ÖMER KETENCİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-20|2030-05-19
    ÖMER KETENCİ|02-Temel İlk Yardım Belgesi|2025-05-20|2030-05-19
    ÖMER KETENCİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-20|2030-05-19
    ÖMER KETENCİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-20|2030-05-19
    ÖMER KETENCİ|31-Gemi Adamı Cüzdan Belgesi|2025-05-20|2030-05-19
    AYNUR AVCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-30|2030-09-29
    AYNUR AVCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-30|2030-09-29
    AYNUR AVCI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-03|2027-07-02
    AYNUR AVCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-10-08|2030-09-22
    AYNUR AVCI|02-Temel İlk Yardım Belgesi|2025-09-30|2030-09-29
    AYNUR AVCI|20-Seyir Vardiyası Tutma Belgesi|2025-09-30|2030-09-29
    AYNUR AVCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-30|2030-09-29
    AYNUR AVCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-30|2030-09-29
    AYNUR AVCI|31-Gemi Adamı Cüzdan Belgesi|2025-10-03|2030-10-02
    ÖZGE ALP ÇINAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-16|2030-09-15
    MEHMET DURMUŞ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-10-25|2026-10-24
    ÖZGE ALP ÇINAR|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-11-22|2029-11-21
    ÖZGE ALP ÇINAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-11-22|2029-11-20
    ÖZGE ALP ÇINAR|31-Gemi Adamı Cüzdan Belgesi|2025-09-16|2030-09-15
    ŞABAN ALİ KUL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-02-13|2030-02-12
    ŞABAN ALİ KUL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-02-13|2030-02-12
    ŞABAN ALİ KUL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-03|2027-02-02
    ŞABAN ALİ KUL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-02-13|2030-02-12
    ŞABAN ALİ KUL|02-Temel İlk Yardım Belgesi|2025-02-13|2030-02-12
    ŞABAN ALİ KUL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-02-13|2030-02-12
    ŞABAN ALİ KUL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-02-13|2030-02-12
    ŞABAN ALİ KUL|31-Gemi Adamı Cüzdan Belgesi|2025-02-18|2030-02-17
    DİLARA USTA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-26|2027-09-25
    DİLARA USTA|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-26|2027-09-25
    DİLARA USTA|00-Gemiadamları Sağlık Yoklama Belgesi|2026-02-03|2028-02-02
    DİLARA USTA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-26|2027-09-25
    DİLARA USTA|02-Temel İlk Yardım Belgesi|2022-09-26|2027-09-25
    DİLARA USTA|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-26|2027-09-25
    DİLARA USTA|31-Gemi Adamı Cüzdan Belgesi|2023-06-16|2028-06-15
    İMRAN BAYAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-03-01|2027-02-28
    İMRAN BAYAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-12|2030-09-11
    MEHMET DURMUŞ|10-İleri Yangınla Mücadele Belgesi|2021-10-25|2026-10-24
    İMRAN BAYAR|10-İleri Yangınla Mücadele Belgesi|2026-04-21|2031-04-20
    İMRAN BAYAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-12|2030-09-11
    İMRAN BAYAR|02-Temel İlk Yardım Belgesi|2025-09-12|2030-09-11
    İMRAN BAYAR|20-Seyir Vardiyası Tutma Belgesi|2022-03-01|2027-02-28
    İMRAN BAYAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-12|2030-09-11
    İMRAN BAYAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-12|2030-09-11
    İMRAN BAYAR|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-12-05|2028-12-04
    MEHMET DURMUŞ|07-Otm.Radar Plot. Aygıt. (ARPA) Kullan.|2021-10-25|2026-10-24
    SERKAN AKÇAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-01|2029-07-31
    SERKAN AKÇAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-08-01|2029-07-31
    SERKAN AKÇAY|02-Temel İlk Yardım Belgesi|2024-08-01|2029-07-31
    SERKAN AKÇAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-08-01|2029-07-31
    SERKAN AKÇAY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-21|2030-05-20
    SERKAN AKÇAY|31-Gemi Adamı Cüzdan Belgesi|2024-10-31|2029-10-30
    DOĞANCAN SOYDAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-14|2029-10-13
    DOĞANCAN SOYDAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-14|2029-10-13
    DOĞANCAN SOYDAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-10|2027-10-09
    DOĞANCAN SOYDAN|10-İleri Yangınla Mücadele Belgesi|2024-10-14|2029-10-13
    DOĞANCAN SOYDAN|02-Temel İlk Yardım Belgesi|2024-10-14|2029-10-13
    DOĞANCAN SOYDAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-14|2029-10-13
    DOĞANCAN SOYDAN|31-Gemi Adamı Cüzdan Belgesi|2024-10-14|2029-10-13
    ANIL COŞKUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-31|2029-10-30
    ANIL COŞKUN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-31|2029-10-30
    ANIL COŞKUN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-19|2027-09-18
    ANIL COŞKUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-10-31|2029-10-30
    ANIL COŞKUN|02-Temel İlk Yardım Belgesi|2024-10-31|2029-10-30
    ANIL COŞKUN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-31|2029-10-30
    ANIL COŞKUN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-02-16|2029-02-15
    ANIL COŞKUN|31-Gemi Adamı Cüzdan Belgesi|2025-07-01|2030-06-30
    KUDRET GİRGİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-08-04|2030-08-03
    KUDRET GİRGİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-08-04|2030-08-03
    MEHMET DURMUŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-25|2026-10-24
    KUDRET GİRGİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-08-04|2030-08-03
    KUDRET GİRGİN|02-Temel İlk Yardım Belgesi|2025-08-04|2030-08-03
    KUDRET GİRGİN|20-Seyir Vardiyası Tutma Belgesi|2025-08-04|2030-08-03
    KUDRET GİRGİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-08-04|2030-08-03
    KUDRET GİRGİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-08-04|2030-08-03
    KUDRET GİRGİN|31-Gemi Adamı Cüzdan Belgesi|2025-08-04|2030-08-03
    YAVUZ KIROĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-10|2027-02-09
    YAVUZ KIROĞLU|10-İleri Yangınla Mücadele Belgesi|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|02-Temel İlk Yardım Belgesi|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-08-16|2029-08-15
    YAVUZ KIROĞLU|31-Gemi Adamı Cüzdan Belgesi|2024-08-16|2029-08-15
    ERAY DİRİARIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-04-17|2029-04-16
    ERAY DİRİARIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-04-17|2029-04-16
    ERAY DİRİARIN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-13|2027-10-12
    ERAY DİRİARIN|10-İleri Yangınla Mücadele Belgesi|2025-04-17|2030-04-16
    ERAY DİRİARIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-17|2030-04-16
    ERAY DİRİARIN|02-Temel İlk Yardım Belgesi|2025-04-17|2030-04-16
    ERAY DİRİARIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-17|2030-04-16
    ERAY DİRİARIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-04-17|2030-04-16
    ERAY DİRİARIN|31-Gemi Adamı Cüzdan Belgesi|2024-04-18|2029-04-17
    AYBERK ÇELİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-12|2029-09-11
    AYBERK ÇELİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-10-02|2028-10-01
    AYBERK ÇELİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-10-02|2028-10-01
    AYBERK ÇELİK|02-Temel İlk Yardım Belgesi|2023-10-02|2028-10-01
    AYBERK ÇELİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-10-02|2028-10-01
    AYBERK ÇELİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-15|2030-09-14
    AYBERK ÇELİK|31-Gemi Adamı Cüzdan Belgesi|2024-09-12|2029-09-11
    DURSUN ÇOLAK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-04-12|2027-04-11
    DURSUN ÇOLAK|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-04-12|2027-04-11
    DURSUN ÇOLAK|10-İleri Yangınla Mücadele Belgesi|2022-04-12|2027-04-11
    DURSUN ÇOLAK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-04-12|2027-04-11
    DURSUN ÇOLAK|02-Temel İlk Yardım Belgesi|2022-04-12|2027-04-11
    DURSUN ÇOLAK|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-04-12|2027-04-11
    DURSUN ÇOLAK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-04-12|2027-04-11
    DURSUN ÇOLAK|31-Gemi Adamı Cüzdan Belgesi|2022-04-12|2027-04-11
    MUSTAFA BOSTAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-08-02|2028-08-01
    MUSTAFA BOSTAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-08-02|2028-08-01
    MUSTAFA BOSTAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-08-02|2028-08-01
    MUSTAFA BOSTAN|02-Temel İlk Yardım Belgesi|2023-08-02|2028-08-01
    MUSTAFA BOSTAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-02|2028-08-01
    MUSTAFA BOSTAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-01-10|2029-01-09
    MUSTAFA BOSTAN|25-Kısa Mesafe Telsiz Operatörü Belgesi|2023-08-02|2028-08-01
    MEHMET DURMUŞ|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-10-25|2026-10-24
    MEHMET DURMUŞ|02-Temel İlk Yardım Belgesi|2021-10-25|2026-10-24
    MEHMET DURMUŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-25|2026-10-24
    İSMAİL YALÇINKAYA|10-İleri Yangınla Mücadele Belgesi|2026-01-04|2030-05-20
    MEHMET DURMUŞ|31-Gemi Adamı Cüzdan Belgesi|2021-10-25|2026-10-24
    TOLGA GÖKSU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-25|2026-10-24
    TOLGA GÖKSU|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-25|2026-10-24
    İSMAİL YALÇINKAYA|31-Gemi Adamı Cüzdan Belgesi|2022-04-29|2027-04-28
    ADNAN CAN SELÇUK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-17|2028-04-16
    ADNAN CAN SELÇUK|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-17|2028-04-16
    ADNAN CAN SELÇUK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-21|2027-05-20
    ADNAN CAN SELÇUK|10-İleri Yangınla Mücadele Belgesi|2023-04-17|2028-04-16
    ADNAN CAN SELÇUK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-17|2028-04-16
    ADNAN CAN SELÇUK|02-Temel İlk Yardım Belgesi|2023-04-17|2028-04-16
    ADNAN CAN SELÇUK|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-17|2028-04-16
    ADNAN CAN SELÇUK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-17|2028-04-16
    TOLGA GÖKSU|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-10-25|2026-10-24
    ALİ YILMAZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-31|2027-01-30
    ALİ YILMAZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-04-21|2030-04-20
    ALİ YILMAZ|10-İleri Yangınla Mücadele Belgesi|2025-04-21|2030-04-20
    ALİ YILMAZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-21|2030-04-20
    ALİ YILMAZ|02-Temel İlk Yardım Belgesi|2025-04-21|2030-04-20
    ALİ YILMAZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-21|2030-04-20
    ALİ YILMAZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-06-13|2027-06-12
    ALİ YILMAZ|31-Gemi Adamı Cüzdan Belgesi|2025-04-28|2030-04-27
    ORHUN DOĞRUYOL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-05-16|2029-05-15
    ORHUN DOĞRUYOL|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-05-16|2029-05-15
    ORHUN DOĞRUYOL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-07|2027-05-06
    ORHUN DOĞRUYOL|10-İleri Yangınla Mücadele Belgesi|2024-08-27|2029-08-26
    ORHUN DOĞRUYOL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-05-16|2029-05-15
    ORHUN DOĞRUYOL|02-Temel İlk Yardım Belgesi|2024-05-16|2029-05-15
    ORHUN DOĞRUYOL|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-05-16|2029-05-15
    ORHUN DOĞRUYOL|31-Gemi Adamı Cüzdan Belgesi|2025-05-16|2030-05-15
    METEHAN ŞEKER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-09-12|2028-09-11
    METEHAN ŞEKER|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-09|2027-09-08
    METEHAN ŞEKER|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-20|2027-10-19
    METEHAN ŞEKER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-09|2027-09-08
    METEHAN ŞEKER|02-Temel İlk Yardım Belgesi|2022-09-09|2027-09-08
    METEHAN ŞEKER|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-09|2027-09-08
    METEHAN ŞEKER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-09-09|2027-09-08
    METEHAN ŞEKER|31-Gemi Adamı Cüzdan Belgesi|2022-09-12|2027-09-11
    UĞURCAN ARI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-06-10|2029-06-09
    UĞURCAN ARI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-16|2027-12-15
    UĞURCAN ARI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-06-10|2029-06-09
    UĞURCAN ARI|02-Temel İlk Yardım Belgesi|2024-06-10|2029-06-09
    UĞURCAN ARI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-06-10|2029-06-09
    UĞURCAN ARI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-10-25|2028-10-24
    UĞURCAN ARI|31-Gemi Adamı Cüzdan Belgesi|2024-06-11|2029-06-10
    HÜSEYİN KÖMÜRCÜ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-09|2029-12-04
    TOLGA GÖKSU|10-İleri Yangınla Mücadele Belgesi|2021-10-25|2026-10-24
    HÜSEYİN KÖMÜRCÜ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-08|2027-09-07
    HÜSEYİN KÖMÜRCÜ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-09-07|2028-09-06
    TOLGA GÖKSU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-25|2026-10-24
    TOLGA GÖKSU|02-Temel İlk Yardım Belgesi|2021-10-25|2026-10-24
    HÜSEYİN KÖMÜRCÜ|20-Seyir Vardiyası Tutma Belgesi|2023-08-24|2028-08-23
    TOLGA GÖKSU|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-25|2026-10-24
    HÜSEYİN KÖMÜRCÜ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-11-29|2029-11-27
    TOLGA GÖKSU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-25|2026-10-24
    ORHAN ÖZKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-09-08|2028-09-07
    ORHAN ÖZKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-09-08|2028-09-07
    ORHAN ÖZKAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-11|2027-06-10
    ORHAN ÖZKAN|02-Temel İlk Yardım Belgesi|2023-09-08|2028-09-07
    ORHAN ÖZKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-09-08|2028-09-07
    ORHAN ÖZKAN|31-Gemi Adamı Cüzdan Belgesi|2023-09-11|2028-09-10
    DOĞUKAN KARATOP|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-10-18|2028-10-17
    DOĞUKAN KARATOP|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-23|2027-09-22
    DOĞUKAN KARATOP|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-22|2027-09-21
    DOĞUKAN KARATOP|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-02-04|2031-01-16
    DOĞUKAN KARATOP|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-29|2027-09-28
    DOĞUKAN KARATOP|02-Temel İlk Yardım Belgesi|2022-09-29|2027-09-28
    DOĞUKAN KARATOP|20-Seyir Vardiyası Tutma Belgesi|2023-10-18|2028-10-17
    DOĞUKAN KARATOP|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-29|2027-09-28
    TOLGA GÖKSU|31-Gemi Adamı Cüzdan Belgesi|2021-10-25|2026-10-24
    DOĞUKAN KARATOP|31-Gemi Adamı Cüzdan Belgesi|2023-12-05|2028-12-04
    ETHEM HACISALİHOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-08-06|2030-08-05
    ETHEM HACISALİHOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-08-06|2030-08-05
    ETHEM HACISALİHOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-06|2027-08-05
    ETHEM HACISALİHOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-08-06|2030-08-05
    ETHEM HACISALİHOĞLU|02-Temel İlk Yardım Belgesi|2025-08-06|2030-08-05
    ETHEM HACISALİHOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-08-06|2030-08-05
    ETHEM HACISALİHOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-08-26|2030-08-25
    ETHEM HACISALİHOĞLU|31-Gemi Adamı Cüzdan Belgesi|2022-12-29|2027-12-28
    GÖKHAN ERSOY|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-11|2027-08-10
    CEYHUN PIRLANT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-26|2026-10-25
    CEYHUN PIRLANT|31-Gemi Adamı Cüzdan Belgesi|2021-10-26|2026-10-25
    GÖKHAN ERSOY|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-05-03|2029-05-02
    GÖKHAN ERSOY|31-Gemi Adamı Cüzdan Belgesi|2022-03-01|2027-02-28
    İLKAY KARABUL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-04-18|2029-04-17
    İLKAY KARABUL|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-28|2027-06-27
    EKREM ÖZÇELİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-10-27|2026-10-26
    İLKAY KARABUL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-06-28|2027-06-27
    İLKAY KARABUL|02-Temel İlk Yardım Belgesi|2022-06-28|2027-06-27
    İLKAY KARABUL|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-28|2027-06-27
    İLKAY KARABUL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-04-18|2029-04-17
    İLKAY KARABUL|31-Gemi Adamı Cüzdan Belgesi|2023-06-01|2028-05-31
    MURAT GÜLMEZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-12|2028-04-11
    MURAT GÜLMEZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-12|2028-04-11
    MURAT GÜLMEZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-28|2027-07-27
    MURAT GÜLMEZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-12|2028-04-11
    MURAT GÜLMEZ|02-Temel İlk Yardım Belgesi|2023-04-12|2028-04-11
    MURAT GÜLMEZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-12|2028-04-11
    EKREM ÖZÇELİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-10-27|2026-10-26
    MURAT GÜLMEZ|31-Gemi Adamı Cüzdan Belgesi|2023-04-14|2028-04-13
    EKREM ÖZÇELİK|10-İleri Yangınla Mücadele Belgesi|2021-10-27|2026-10-26
    PELİN KARAARSLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-11-04|2027-03-21
    EKREM ÖZÇELİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-10-27|2026-10-26
    PELİN KARAARSLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-01|2027-09-30
    EKREM ÖZÇELİK|02-Temel İlk Yardım Belgesi|2021-10-27|2026-10-26
    EKREM ÖZÇELİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-10-27|2026-10-26
    PELİN KARAARSLAN|20-Seyir Vardiyası Tutma Belgesi|2024-02-08|2029-02-07
    EKREM ÖZÇELİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-10-27|2026-10-26
    MUHAMMED ALİ AVCI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-10-30|2026-10-29
    ERKAN UZUNER|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-01|2026-10-31
    ZEKİ ŞAHİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-06-02|2027-06-01
    ZEKİ ŞAHİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-02|2027-06-01
    ATAKAN ÇELİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-01|2026-10-31
    ZEKİ ŞAHİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-06-02|2027-06-01
    ZEKİ ŞAHİN|02-Temel İlk Yardım Belgesi|2022-06-02|2027-06-01
    ZEKİ ŞAHİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-02|2027-06-01
    ZEKİ ŞAHİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-06-02|2027-06-01
    ZEKİ ŞAHİN|31-Gemi Adamı Cüzdan Belgesi|2022-06-02|2027-06-01
    ATAKAN ÇELİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-01|2026-10-31
    CEYHUN PIRLANT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-28|2029-02-27
    CEYHUN PIRLANT|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-02-28|2029-02-27
    CEYHUN PIRLANT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-02-28|2029-02-27
    CEYHUN PIRLANT|02-Temel İlk Yardım Belgesi|2024-02-28|2029-02-27
    CEYHUN PIRLANT|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-28|2029-02-27
    ATAKAN ÇELİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-01|2026-10-31
    ATAKAN ÇELİK|02-Temel İlk Yardım Belgesi|2021-11-01|2026-10-31
    ATAKAN ÇELİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-01|2026-10-31
    YASİN METE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-20|2029-12-19
    YASİN METE|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-20|2029-12-19
    YASİN METE|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-20|2029-12-19
    YASİN METE|02-Temel İlk Yardım Belgesi|2024-12-20|2029-12-19
    YASİN METE|20-Seyir Vardiyası Tutma Belgesi|2025-05-28|2030-05-27
    YASİN METE|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-20|2029-12-19
    YASİN METE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-20|2029-12-19
    YASİN METE|31-Gemi Adamı Cüzdan Belgesi|2026-11-11|2031-02-02
    ATAKAN ÇELİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-01|2026-10-31
    ATAKAN ÇELİK|31-Gemi Adamı Cüzdan Belgesi|2021-11-01|2026-10-31
    BAHATTİN ÇELİK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-01|2027-11-30
    GÖKHAN ERSOY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-02|2026-11-01
    CENGİZ İLİTLİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-04|2026-11-03
    CENGİZ İLİTLİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-04|2026-11-03
    CENGİZ İLİTLİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-04|2026-11-03
    CENGİZ İLİTLİ|02-Temel İlk Yardım Belgesi|2021-11-04|2026-11-03
    BASRİ KASAP|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-12|2027-01-11
    BASRİ KASAP|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-01-12|2027-01-11
    BASRİ KASAP|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-08|2027-10-07
    BASRİ KASAP|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-12|2027-01-11
    BASRİ KASAP|02-Temel İlk Yardım Belgesi|2022-01-12|2027-01-11
    BASRİ KASAP|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-12|2027-01-11
    BASRİ KASAP|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-06-06|2028-06-05
    BASRİ KASAP|31-Gemi Adamı Cüzdan Belgesi|2022-01-12|2027-01-11
    ÖMER FARUK ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-19|2029-12-18
    ÖMER FARUK ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-01-03|2029-01-02
    ÖMER FARUK ÖZTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-22|2027-10-21
    ÖMER FARUK ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-01-03|2029-01-02
    ÖMER FARUK ÖZTÜRK|02-Temel İlk Yardım Belgesi|2024-01-03|2029-01-02
    ÖMER FARUK ÖZTÜRK|20-Seyir Vardiyası Tutma Belgesi|2024-01-03|2029-01-02
    ÖMER FARUK ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-01-03|2029-01-02
    ÖMER FARUK ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-10|2029-09-09
    ÖMER FARUK ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2024-01-04|2029-01-03
    CENGİZ İLİTLİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-04|2026-11-03
    CENGİZ İLİTLİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-04|2026-11-03
    CENGİZ İLİTLİ|31-Gemi Adamı Cüzdan Belgesi|2021-11-04|2026-11-03
    EYÜP NAMOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-04|2026-11-03
    RAMAZAN SANDIKÇI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-04|2026-11-03
    AHMET ZAFER DİNÇ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-04|2026-11-03
    PELİN KARAARSLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-04|2026-11-03
    HÜSEYİN BATUHAN BAYRAKTAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-04-11|2030-04-10
    PELİN KARAARSLAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-04|2026-11-03
    HÜSEYİN BATUHAN BAYRAKTAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-11|2030-04-10
    HÜSEYİN BATUHAN BAYRAKTAR|02-Temel İlk Yardım Belgesi|2025-04-11|2030-04-10
    HÜSEYİN BATUHAN BAYRAKTAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-11|2030-04-10
    HÜSEYİN BATUHAN BAYRAKTAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-04-11|2030-04-10
    HÜSEYİN BATUHAN BAYRAKTAR|31-Gemi Adamı Cüzdan Belgesi|2025-04-15|2030-04-14
    ELİF ÖZLEM SEZEK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-07|2029-08-06
    PELİN KARAARSLAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-04|2026-11-03
    ELİF ÖZLEM SEZEK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-12|2027-09-11
    ELİF ÖZLEM SEZEK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-03-02|2027-03-01
    PELİN KARAARSLAN|02-Temel İlk Yardım Belgesi|2021-11-04|2026-11-03
    PELİN KARAARSLAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-04|2026-11-03
    PELİN KARAARSLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-04|2026-11-03
    PELİN KARAARSLAN|31-Gemi Adamı Cüzdan Belgesi|2021-11-04|2026-11-03
    HÜSEYİN ATLI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-05|2026-11-04
    MERT AKAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-20|2029-02-19
    MERT AKAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-02-20|2029-02-19
    HASAN KÜÇÜKİSLAMOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-05|2026-11-04
    MERT AKAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-02-20|2029-02-19
    MERT AKAR|02-Temel İlk Yardım Belgesi|2024-02-20|2029-02-19
    MERT AKAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-20|2029-02-19
    MERT AKAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-21|2030-05-20
    BATU HAN DÜLGE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-09|2030-12-08
    BATU HAN DÜLGE|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-09|2030-12-08
    BATU HAN DÜLGE|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-05|2027-09-04
    BATU HAN DÜLGE|10-İleri Yangınla Mücadele Belgesi|2025-12-09|2030-12-08
    BATU HAN DÜLGE|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-09|2030-12-08
    BATU HAN DÜLGE|02-Temel İlk Yardım Belgesi|2025-12-09|2030-12-08
    BATU HAN DÜLGE|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-09|2030-12-08
    BATU HAN DÜLGE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-17|2029-12-16
    BATU HAN DÜLGE|31-Gemi Adamı Cüzdan Belgesi|2025-12-26|2030-12-25
    GÖKHAN KAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-10-10|2027-10-09
    GÖKHAN KAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2025-12-01|2027-11-30
    GÖKHAN KAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-12-18|2028-12-17
    GÖKHAN KAYA|02-Temel İlk Yardım Belgesi|2022-01-07|2027-01-06
    GÖKHAN KAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-07|2027-01-06
    GÖKHAN KAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-07|2027-01-06
    GÖKHAN KAYA|31-Gemi Adamı Cüzdan Belgesi|2022-01-07|2027-01-06
    VOLKAN UZUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-01-07|2027-01-06
    RECEP HARDAL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-07|2026-11-06
    VOLKAN UZUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-07|2027-01-06
    ELİF ÖZLEM SEZEK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-09|2026-11-08
    FEYZA ÖZKAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-11|2026-11-10
    VOLKAN UZUN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-07|2027-01-06
    VOLKAN UZUN|31-Gemi Adamı Cüzdan Belgesi|2025-11-13|2030-11-12
    ONUR AYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-18|2029-10-17
    ONUR AYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-09-19|2028-09-18
    ONUR AYDIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-09-19|2028-09-18
    ONUR AYDIN|02-Temel İlk Yardım Belgesi|2023-09-19|2028-09-18
    ONUR AYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-09-19|2028-09-18
    FEYZA ÖZKAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-11|2026-11-10
    ONUR AYDIN|31-Gemi Adamı Cüzdan Belgesi|2023-09-19|2028-09-18
    AHMET SERHAN BÜYÜKKERMAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-07-27|2028-07-26
    AHMET SERHAN BÜYÜKKERMAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-07-27|2028-07-26
    AHMET SERHAN BÜYÜKKERMAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-03|2027-01-02
    AHMET SERHAN BÜYÜKKERMAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-07-27|2028-07-26
    AHMET SERHAN BÜYÜKKERMAN|02-Temel İlk Yardım Belgesi|2023-07-27|2028-07-26
    AHMET SERHAN BÜYÜKKERMAN|22-VHF Haberleşme Belgesi|2022-02-25|2027-02-24
    AHMET SERHAN BÜYÜKKERMAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-07-27|2028-07-26
    AHMET SERHAN BÜYÜKKERMAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-02|2029-10-01
    HÜSEYİN ATLI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-04-18|2029-04-17
    FEYZA ÖZKAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-11|2026-11-10
    HÜSEYİN ATLI|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2025-08-22|2030-08-21
    HÜSEYİN ATLI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-05-22|2030-05-21
    HÜSEYİN ATLI|02-Temel İlk Yardım Belgesi|2024-04-18|2029-04-17
    HÜSEYİN ATLI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-04-18|2029-04-17
    HÜSEYİN ATLI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-04-18|2029-04-17
    HÜSEYİN ATLI|31-Gemi Adamı Cüzdan Belgesi|2024-04-24|2029-04-23
    BAYKAN CİNOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-20|2029-12-19
    BAYKAN CİNOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-30|2029-12-29
    BAYKAN CİNOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-30|2029-12-29
    BAYKAN CİNOĞLU|02-Temel İlk Yardım Belgesi|2024-12-20|2029-12-19
    BAYKAN CİNOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-20|2029-12-19
    BAYKAN CİNOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-20|2029-12-19
    BAYKAN CİNOĞLU|31-Gemi Adamı Cüzdan Belgesi|2024-12-20|2029-12-19
    BURAK YÜMLÜ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-21|2029-02-20
    BURAK YÜMLÜ|00-Gemiadamları Sağlık Yoklama Belgesi|2026-05-06|2028-05-05
    BURAK YÜMLÜ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-04-07|2027-04-06
    BURAK YÜMLÜ|10-İleri Yangınla Mücadele Belgesi|2024-02-21|2029-02-20
    BURAK YÜMLÜ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-02-21|2029-02-20
    BURAK YÜMLÜ|02-Temel İlk Yardım Belgesi|2024-02-21|2029-02-20
    BURAK YÜMLÜ|20-Seyir Vardiyası Tutma Belgesi|2024-02-21|2029-02-20
    BURAK YÜMLÜ|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-21|2029-02-20
    BURAK YÜMLÜ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-02-21|2029-02-20
    BURAK YÜMLÜ|31-Gemi Adamı Cüzdan Belgesi|2024-02-21|2029-02-20
    EMRE YEŞİL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-08-14|2030-08-13
    EMRE YEŞİL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-08-14|2030-08-13
    EMRE YEŞİL|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-02|2028-06-01
    EMRE YEŞİL|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-05-06|2027-05-05
    EMRE YEŞİL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-08-14|2030-08-13
    EMRE YEŞİL|02-Temel İlk Yardım Belgesi|2025-08-14|2030-08-13
    EMRE YEŞİL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-08-14|2030-08-13
    EMRE YEŞİL|31-Gemi Adamı Cüzdan Belgesi|2025-08-14|2030-08-13
    KAHRAMAN GÜMÜŞLER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-11-29|2027-11-28
    KAHRAMAN GÜMÜŞLER|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-11-29|2027-11-28
    KAHRAMAN GÜMÜŞLER|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-03-29|2029-03-28
    KAHRAMAN GÜMÜŞLER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-11-29|2027-11-28
    KAHRAMAN GÜMÜŞLER|02-Temel İlk Yardım Belgesi|2022-11-29|2027-11-28
    KAHRAMAN GÜMÜŞLER|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-11-29|2027-11-28
    KAHRAMAN GÜMÜŞLER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-11-29|2027-11-28
    KAHRAMAN GÜMÜŞLER|31-Gemi Adamı Cüzdan Belgesi|2022-11-29|2027-11-28
    GÜLSÜM İNANÇ KORKMAZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-16|2027-09-15
    FEYZA ÖZKAN|02-Temel İlk Yardım Belgesi|2021-11-11|2026-11-10
    GÜLSÜM İNANÇ KORKMAZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-26|2027-11-25
    GÜLSÜM İNANÇ KORKMAZ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-02-18|2028-02-17
    FEYZA ÖZKAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-11|2026-11-10
    FEYZA ÖZKAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-11|2026-11-10
    GÜLSÜM İNANÇ KORKMAZ|20-Seyir Vardiyası Tutma Belgesi|2022-09-16|2027-09-15
    KORAY TORLAK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-12|2026-11-11
    GÜLSÜM İNANÇ KORKMAZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-25|2027-02-24
    KORAY TORLAK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-12|2026-11-11
    UĞUR DURMUŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-19|2028-04-18
    UĞUR DURMUŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-19|2028-04-18
    UĞUR DURMUŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-07|2027-04-06
    UĞUR DURMUŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-19|2028-04-18
    UĞUR DURMUŞ|02-Temel İlk Yardım Belgesi|2023-04-19|2028-04-18
    UĞUR DURMUŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-19|2028-04-18
    UĞUR DURMUŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-19|2028-04-18
    SÜLEYMAN AYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-11-07|2029-11-06
    SÜLEYMAN AYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-11-07|2029-11-06
    KORAY TORLAK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-11-12|2026-11-11
    SÜLEYMAN AYDIN|10-İleri Yangınla Mücadele Belgesi|2024-11-07|2029-11-06
    SÜLEYMAN AYDIN|02-Temel İlk Yardım Belgesi|2024-11-07|2029-11-06
    SÜLEYMAN AYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-11-07|2029-11-06
    SÜLEYMAN AYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-11-07|2029-11-06
    AYDIN SAZDAĞI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-23|2027-07-22
    AYDIN SAZDAĞI|10-İleri Yangınla Mücadele Belgesi|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|02-Temel İlk Yardım Belgesi|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-07-02|2030-07-01
    AYDIN SAZDAĞI|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-07-17|2029-07-16
    AYDIN SAZDAĞI|31-Gemi Adamı Cüzdan Belgesi|2024-07-19|2029-07-18
    AZİZ KAAN ÇELİKAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-23|2030-05-22
    AZİZ KAAN ÇELİKAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-11|2028-05-10
    AZİZ KAAN ÇELİKAY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-11|2028-05-10
    AZİZ KAAN ÇELİKAY|02-Temel İlk Yardım Belgesi|2023-05-11|2028-05-10
    AZİZ KAAN ÇELİKAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-11|2028-05-10
    AZİZ KAAN ÇELİKAY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-23|2030-05-22
    AZİZ KAAN ÇELİKAY|31-Gemi Adamı Cüzdan Belgesi|2023-05-26|2028-05-25
    KUBİLAY ŞENTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-28|2027-02-27
    KUBİLAY ŞENTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2026-02-17|2028-02-16
    KUBİLAY ŞENTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-28|2027-02-27
    KUBİLAY ŞENTÜRK|02-Temel İlk Yardım Belgesi|2022-02-28|2027-02-27
    KUBİLAY ŞENTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-28|2027-02-27
    KUBİLAY ŞENTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-28|2027-02-27
    KUBİLAY ŞENTÜRK|31-Gemi Adamı Cüzdan Belgesi|2022-02-28|2027-02-27
    KORAY TORLAK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-12|2026-11-11
    KORAY TORLAK|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-11-12|2026-11-11
    KORAY TORLAK|02-Temel İlk Yardım Belgesi|2021-11-12|2026-11-11
    KORAY TORLAK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-12|2026-11-11
    KORAY TORLAK|31-Gemi Adamı Cüzdan Belgesi|2021-11-12|2026-11-11
    FEYZA ÖZKAN|31-Gemi Adamı Cüzdan Belgesi|2021-11-12|2026-11-11
    ÜMİT EKİZ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-12|2026-11-11
    SELÇUK ATAKURU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-14|2026-11-13
    RABİA SELİN DELİOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-12-03|2027-12-02
    RABİA SELİN DELİOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-15|2027-06-14
    RIDVAN PAŞALI|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-14|2026-11-13
    RABİA SELİN DELİOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-06-15|2027-06-14
    RABİA SELİN DELİOĞLU|02-Temel İlk Yardım Belgesi|2022-06-15|2027-06-14
    RABİA SELİN DELİOĞLU|20-Seyir Vardiyası Tutma Belgesi|2024-12-03|2029-12-02
    RABİA SELİN DELİOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-15|2027-06-14
    RABİA SELİN DELİOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-06-29|2027-06-28
    RECEP ASLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-15|2026-11-14
    İSMET MOLLA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-01-26|2029-01-25
    İSMET MOLLA|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-06-06|2027-06-05
    İSMET MOLLA|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-25|2027-09-22
    İSMET MOLLA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-08-08|2027-08-07
    İSMET MOLLA|02-Temel İlk Yardım Belgesi|2022-06-06|2027-06-05
    İSMET MOLLA|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-06-06|2027-06-05
    İSMET MOLLA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-01-26|2029-01-25
    İSMET MOLLA|31-Gemi Adamı Cüzdan Belgesi|2022-06-06|2027-06-05
    ATAKAN MÜLDÜR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-03-19|2030-03-18
    ATAKAN MÜLDÜR|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-03|2029-10-02
    ATAKAN MÜLDÜR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-10-03|2029-10-02
    ATAKAN MÜLDÜR|02-Temel İlk Yardım Belgesi|2024-10-03|2029-10-02
    ATAKAN MÜLDÜR|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-03|2029-10-02
    ATAKAN MÜLDÜR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-03|2029-10-02
    ATAKAN MÜLDÜR|31-Gemi Adamı Cüzdan Belgesi|2024-10-04|2029-10-03
    HARUN MARAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-07-18|2027-07-17
    HARUN MARAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-07-18|2027-07-17
    GÖKHAN TARLACI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-02-02|2027-02-01
    GÖKHAN TARLACI|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-02|2027-02-01
    ATAKAN ÇELİK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-15|2026-11-14
    GÖKHAN TARLACI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-02|2027-02-01
    GÖKHAN TARLACI|02-Temel İlk Yardım Belgesi|2022-02-02|2027-02-01
    GÖKHAN TARLACI|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-02|2027-02-01
    GÖKHAN TARLACI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-10-24|2027-10-23
    GÖKHAN TARLACI|31-Gemi Adamı Cüzdan Belgesi|2022-02-02|2027-02-01
    MUHAMMED ATIF YÜKSELOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-01-10|2030-01-09
    MUHAMMED ATIF YÜKSELOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-01-10|2030-01-09
    BARIŞ ELİBAL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-15|2026-11-14
    MUHAMMED ATIF YÜKSELOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-01-10|2030-01-09
    MUHAMMED ATIF YÜKSELOĞLU|02-Temel İlk Yardım Belgesi|2025-01-10|2030-01-09
    MUHAMMED ATIF YÜKSELOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-01-10|2030-01-09
    MUHAMMED ATIF YÜKSELOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-01-10|2030-01-09
    MUHAMMED ATIF YÜKSELOĞLU|31-Gemi Adamı Cüzdan Belgesi|2025-01-11|2030-01-10
    MUHAMMET EMRE KURU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-10-11|2027-10-10
    MUHAMMET EMRE KURU|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-05-14|2029-05-13
    MUHAMMET EMRE KURU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-05-14|2029-05-13
    MUHAMMET EMRE KURU|02-Temel İlk Yardım Belgesi|2024-05-14|2029-05-13
    MUHAMMET EMRE KURU|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-05-14|2029-05-13
    MUHAMMET EMRE KURU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-05-14|2029-05-13
    MUHAMMET EMRE KURU|31-Gemi Adamı Cüzdan Belgesi|2024-05-17|2029-05-16
    MEHMET AKİF KALAFAT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-05-14|2029-05-13
    RIDVAN PAŞALI|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-16|2026-11-15
    MEHMET AKİF KALAFAT|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-07|2027-01-06
    MEHMET AKİF KALAFAT|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2022-03-21|2027-03-20
    MEHMET AKİF KALAFAT|10-İleri Yangınla Mücadele Belgesi|2024-11-16|2029-11-15
    RIDVAN PAŞALI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-16|2026-11-15
    RIDVAN PAŞALI|02-Temel İlk Yardım Belgesi|2021-11-16|2026-11-15
    RIDVAN PAŞALI|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-16|2026-11-15
    MEHMET AKİF KALAFAT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-05-14|2029-05-13
    SERKAN KAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-21|2026-11-20
    HASRET CAN KÖSE|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-09-15|2028-09-14
    HASRET CAN KÖSE|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-23|2027-09-22
    HASRET CAN KÖSE|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-10-12|2027-10-11
    MEVLÜT TAŞCIOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-18|2028-04-17
    MEVLÜT TAŞCIOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-18|2028-04-17
    MEVLÜT TAŞCIOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-18|2028-04-17
    MEVLÜT TAŞCIOĞLU|02-Temel İlk Yardım Belgesi|2023-04-18|2028-04-17
    MEVLÜT TAŞCIOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-18|2028-04-17
    MEVLÜT TAŞCIOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-18|2028-04-17
    MEVLÜT TAŞCIOĞLU|31-Gemi Adamı Cüzdan Belgesi|2023-04-24|2028-04-23
    İBRAHİM ONUR PERİŞANOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-02|2027-09-01
    İBRAHİM ONUR PERİŞANOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-02|2027-09-01
    İBRAHİM ONUR PERİŞANOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-22|2027-05-21
    İBRAHİM ONUR PERİŞANOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-02|2027-09-01
    İBRAHİM ONUR PERİŞANOĞLU|02-Temel İlk Yardım Belgesi|2022-09-02|2027-09-01
    İBRAHİM ONUR PERİŞANOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-02|2027-09-01
    İBRAHİM ONUR PERİŞANOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-05-22|2028-05-21
    İBRAHİM ONUR PERİŞANOĞLU|31-Gemi Adamı Cüzdan Belgesi|2022-09-02|2027-09-01
    MUHAMMED ALİ ERTİK|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-21|2026-11-20
    ÖMER KANBUR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-22|2026-11-21
    İBRAHİM MARAŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-10|2027-01-09
    ÖMER KANBUR|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-22|2026-11-21
    ÖMER KANBUR|10-İleri Yangınla Mücadele Belgesi|2021-11-22|2026-11-21
    ÖMER KANBUR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-22|2026-11-21
    İBRAHİM MARAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-06-07|2029-06-06
    ÖMER KANBUR|02-Temel İlk Yardım Belgesi|2021-11-22|2026-11-21
    ÖMER KANBUR|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-22|2026-11-21
    ÖMER KANBUR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-22|2026-11-21
    DENİZ ARSLANBAŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-08-25|2027-08-24
    ÖMER KANBUR|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-11-22|2026-11-21
    ÖMER KANBUR|31-Gemi Adamı Cüzdan Belgesi|2021-11-22|2026-11-21
    OKTAY GÖÇ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-23|2026-11-22
    OKTAY GÖÇ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-23|2026-11-22
    DENİZ ARSLANBAŞ|31-Gemi Adamı Cüzdan Belgesi|2022-01-11|2027-01-10
    ORHAN ATAMAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-19|2029-10-18
    ORHAN ATAMAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-07-31|2029-07-30
    ORHAN ATAMAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-29|2027-07-28
    ORHAN ATAMAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-07-31|2029-07-30
    ORHAN ATAMAN|02-Temel İlk Yardım Belgesi|2024-07-31|2029-07-30
    ORHAN ATAMAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-07-31|2029-07-30
    ORHAN ATAMAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-03-16|2027-03-15
    ORHAN ATAMAN|31-Gemi Adamı Cüzdan Belgesi|2023-07-27|2028-07-26
    ÖZLEM GÜRSU TARLACI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-14|2027-09-13
    ÖZLEM GÜRSU TARLACI|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-14|2027-09-13
    OKTAY GÖÇ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-23|2026-11-22
    ÖZLEM GÜRSU TARLACI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-14|2027-09-13
    ÖZLEM GÜRSU TARLACI|02-Temel İlk Yardım Belgesi|2022-09-14|2027-09-13
    ÖZLEM GÜRSU TARLACI|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-14|2027-09-13
    ÖZLEM GÜRSU TARLACI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-08-02|2029-08-01
    ÖZLEM GÜRSU TARLACI|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-08-02|2029-08-01
    ÖZLEM GÜRSU TARLACI|31-Gemi Adamı Cüzdan Belgesi|2023-07-26|2028-07-25
    HAKAN ÖZÇINAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-13|2027-03-12
    HAKAN ÖZÇINAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|02-Temel İlk Yardım Belgesi|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|20-Seyir Vardiyası Tutma Belgesi|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-05-31|2029-05-30
    HAKAN ÖZÇINAR|31-Gemi Adamı Cüzdan Belgesi|2024-06-04|2029-06-03
    ESAT KÖNEÇOĞLU|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-08-22|2029-08-21
    ESAT KÖNEÇOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-08-22|2027-08-21
    ESAT KÖNEÇOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-04|2027-09-03
    ESAT KÖNEÇOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-08-22|2027-08-21
    ESAT KÖNEÇOĞLU|02-Temel İlk Yardım Belgesi|2022-08-22|2027-08-21
    ESAT KÖNEÇOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-08-22|2027-08-21
    ESAT KÖNEÇOĞLU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-08-22|2027-08-21
    ESAT KÖNEÇOĞLU|31-Gemi Adamı Cüzdan Belgesi|2023-09-19|2028-09-18
    OSMAN KOÇAK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-05-12|2027-05-11
    OKTAY GÖÇ|02-Temel İlk Yardım Belgesi|2021-11-23|2026-11-22
    HASAN ASLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-06|2028-01-05
    OKTAY GÖÇ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-23|2026-11-22
    UMUT KURUOĞLU|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-13|2027-09-12
    UMUT KURUOĞLU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-12|2027-03-11
    UMUT KURUOĞLU|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-13|2027-09-12
    UMUT KURUOĞLU|02-Temel İlk Yardım Belgesi|2022-09-13|2027-09-12
    UMUT KURUOĞLU|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-13|2027-09-12
    OKTAY GÖÇ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-23|2026-11-22
    OKTAY GÖÇ|31-Gemi Adamı Cüzdan Belgesi|2021-11-23|2026-11-22
    RIFAT MERT ÖZAYDIN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-17|2027-09-16
    MURAT AVŞAR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-11-25|2026-11-24
    MURAT AVŞAR|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-11-25|2026-11-24
    MURAT AVŞAR|10-İleri Yangınla Mücadele Belgesi|2021-11-25|2026-11-24
    MURAT AVŞAR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-11-25|2026-11-24
    MURAT AVŞAR|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-11-25|2026-11-24
    MURAT AVŞAR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-11-25|2026-11-24
    RAHMİ GÖK|31-Gemi Adamı Cüzdan Belgesi|2021-11-25|2026-11-24
    AHMET YILMAZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-01-15|2027-01-14
    SALİH KARACA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-25|2026-11-24
    CÜNEYT HAMARAT|00-Gemiadamları Sağlık Yoklama Belgesi|2024-11-25|2026-11-24
    VOLKAN UZUN|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-01|2026-11-30
    VOLKAN UZUN|02-Temel İlk Yardım Belgesi|2021-12-01|2026-11-30
    VOLKAN UZUN|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-01|2026-11-30
    HAMZA SAFER ÇINAR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-03|2026-12-02
    UFUK BİTİŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-04-11|2027-04-10
    UFUK BİTİŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-04-11|2027-04-10
    UFUK BİTİŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-17|2027-09-16
    UFUK BİTİŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-04-11|2027-04-10
    UFUK BİTİŞ|02-Temel İlk Yardım Belgesi|2022-04-11|2027-04-10
    UFUK BİTİŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-04-11|2027-04-10
    UFUK BİTİŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-12|2029-12-11
    UFUK BİTİŞ|31-Gemi Adamı Cüzdan Belgesi|2024-12-19|2029-12-18
    ABDULAZİZ ÇELİK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-21|2027-05-20
    YUSUF TANRITANIR|00-Gemiadamları Sağlık Yoklama Belgesi|2025-04-14|2027-04-13
    ÜMİT EKİZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-22|2030-05-21
    ÜMİT EKİZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-05-22|2030-05-21
    GÖKHAN CEYLAN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-03|2026-12-02
    ÜMİT EKİZ|02-Temel İlk Yardım Belgesi|2025-05-22|2030-05-21
    ÜMİT EKİZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-22|2030-05-21
    ÜMİT EKİZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-22|2030-05-21
    ERAY KIZILKAYA|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-05|2026-12-04
    İSMAİL HIZARCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-14|2027-09-13
    İSMAİL HIZARCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-14|2027-09-13
    İSMAİL HIZARCI|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-30|2028-06-29
    İSMAİL HIZARCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-14|2027-09-13
    İSMAİL HIZARCI|02-Temel İlk Yardım Belgesi|2022-09-14|2027-09-13
    İSMAİL HIZARCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-14|2027-09-13
    İSMAİL HIZARCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-07-16|2029-07-15
    İSMAİL HIZARCI|31-Gemi Adamı Cüzdan Belgesi|2024-07-19|2029-07-18
    SUDE DARDAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-01-26|2029-01-25
    SUDE DARDAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-08-24|2027-08-23
    SUDE DARDAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-08-24|2027-08-23
    SUDE DARDAN|02-Temel İlk Yardım Belgesi|2022-08-24|2027-08-23
    SUDE DARDAN|20-Seyir Vardiyası Tutma Belgesi|2024-01-26|2029-01-25
    SUDE DARDAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-08-24|2027-08-23
    SUDE DARDAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-08-24|2027-08-23
    SUDE DARDAN|31-Gemi Adamı Cüzdan Belgesi|2022-08-31|2027-08-30
    ORHAN GEÇER|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-05|2026-12-04
    FURKAN ÇAKMAKTAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-09-23|2030-09-22
    FURKAN ÇAKMAKTAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-17|2028-02-16
    RABİA SELİN DELİOĞLU|31-Gemi Adamı Cüzdan Belgesi|2021-12-06|2026-12-05
    FURKAN ÇAKMAKTAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-17|2028-02-16
    FURKAN ÇAKMAKTAŞ|02-Temel İlk Yardım Belgesi|2023-02-17|2028-02-16
    FURKAN ÇAKMAKTAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-17|2028-02-16
    FURKAN ÇAKMAKTAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-09-23|2030-09-22
    FURKAN ÇAKMAKTAŞ|31-Gemi Adamı Cüzdan Belgesi|2023-03-21|2028-03-20
    İBRAHİM EKİZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-03-06|2028-03-05
    İBRAHİM EKİZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-17|2027-11-16
    İBRAHİM EKİZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-03-06|2028-03-05
    İBRAHİM EKİZ|02-Temel İlk Yardım Belgesi|2023-03-06|2028-03-05
    İBRAHİM EKİZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-03-06|2028-03-05
    İBRAHİM EKİZ|31-Gemi Adamı Cüzdan Belgesi|2023-03-15|2028-03-14
    RIDVAN PAŞALI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-06-25|2029-06-24
    İSHAK TAŞCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-07|2026-12-06
    İSHAK TAŞCI|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-07|2026-12-06
    RIDVAN PAŞALI|10-İleri Yangınla Mücadele Belgesi|2024-06-25|2029-06-24
    İSHAK TAŞCI|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-12-07|2026-12-06
    İSHAK TAŞCI|10-İleri Yangınla Mücadele Belgesi|2021-12-07|2026-12-06
    İSHAK TAŞCI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-07|2026-12-06
    RIDVAN PAŞALI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-06-25|2029-06-24
    RIDVAN PAŞALI|31-Gemi Adamı Cüzdan Belgesi|2024-06-28|2029-06-27
    İSHAK TAŞCI|02-Temel İlk Yardım Belgesi|2021-12-07|2026-12-06
    İSHAK TAŞCI|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-07|2026-12-06
    İSHAK TAŞCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-07|2026-12-06
    İSHAK TAŞCI|31-Gemi Adamı Cüzdan Belgesi|2021-12-07|2026-12-06
    MURAT TAFRALI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-07|2026-12-06
    MURAT TAFRALI|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-07|2026-12-06
    MURAT TAFRALI|10-İleri Yangınla Mücadele Belgesi|2021-12-07|2026-12-06
    MURAT TAFRALI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-07|2026-12-06
    YUSUF ARDA KAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-08-04|2027-08-03
    YUSUF ARDA KAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-08-04|2027-08-03
    MURAT TAFRALI|17-Ro-Ro Yolcu Gemileri Gemiadamları Blg|2021-12-07|2026-12-06
    YUSUF ARDA KAYA|10-İleri Yangınla Mücadele Belgesi|2023-08-15|2028-08-14
    YUSUF ARDA KAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2028-08-15|2030-05-19
    YUSUF ARDA KAYA|02-Temel İlk Yardım Belgesi|2023-08-15|2028-08-14
    YUSUF ARDA KAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-15|2028-08-14
    YUSUF ARDA KAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-05-16|2027-05-15
    YUSUF ARDA KAYA|31-Gemi Adamı Cüzdan Belgesi|2024-07-06|2029-07-05
    İBRAHİM AKÇAY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-03-09|2028-03-08
    İBRAHİM AKÇAY|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-03-09|2028-03-08
    MURAT TAFRALI|02-Temel İlk Yardım Belgesi|2021-12-07|2026-12-06
    İBRAHİM AKÇAY|10-İleri Yangınla Mücadele Belgesi|2025-05-20|2030-05-19
    İBRAHİM AKÇAY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-03-09|2028-03-08
    İBRAHİM AKÇAY|02-Temel İlk Yardım Belgesi|2023-03-09|2028-03-08
    İBRAHİM AKÇAY|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-03-09|2028-03-08
    İBRAHİM AKÇAY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-03-09|2028-03-08
    MELİH ERÇAYAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-08-16|2028-08-15
    MELİH ERÇAYAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-08-16|2028-08-15
    MURAT TAFRALI|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-07|2026-12-06
    MELİH ERÇAYAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-08-16|2028-08-15
    MELİH ERÇAYAN|02-Temel İlk Yardım Belgesi|2023-08-16|2028-08-15
    MELİH ERÇAYAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-16|2028-08-15
    MELİH ERÇAYAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-08-16|2028-08-15
    MELİH ERÇAYAN|31-Gemi Adamı Cüzdan Belgesi|2023-08-17|2028-08-16
    ZAFER PAPAKER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-02-10|2029-02-09
    ZAFER PAPAKER|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-02-10|2029-02-09
    MURAT TAFRALI|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-12-07|2026-12-06
    ZAFER PAPAKER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-02-10|2029-02-09
    ZAFER PAPAKER|02-Temel İlk Yardım Belgesi|2024-02-10|2029-02-09
    ZAFER PAPAKER|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-02-10|2029-02-09
    ZAFER PAPAKER|31-Gemi Adamı Cüzdan Belgesi|2024-02-10|2029-02-09
    MURAT TAFRALI|31-Gemi Adamı Cüzdan Belgesi|2021-12-07|2026-12-06
    GÖKHAN CEYLAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-07|2026-12-06
    CİHAT SARIHASAN|00-Gemiadamları Sağlık Yoklama Belgesi|2026-06-16|2028-06-15
    ERSEN DİKMEN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-09|2026-12-08
    EYÜP GÜNDOĞDU|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-09|2026-12-08
    BARIŞ ELİBAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-09|2026-12-08
    BARIŞ ELİBAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-09|2026-12-08
    BARIŞ ELİBAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-09|2026-12-08
    BARIŞ ELİBAL|02-Temel İlk Yardım Belgesi|2021-12-09|2026-12-08
    GÖKHAN CEYLAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-08-14|2029-08-13
    BARIŞ ELİBAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-09|2026-12-08
    GÖKHAN CEYLAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-08-14|2029-08-13
    GÖKHAN CEYLAN|02-Temel İlk Yardım Belgesi|2024-08-14|2029-08-13
    GÖKHAN CEYLAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-08-14|2029-08-13
    GÖKHAN CEYLAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-06-05|2028-06-04
    GÖKHAN CEYLAN|31-Gemi Adamı Cüzdan Belgesi|2024-08-15|2029-08-14
    EMRE SOYDAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-20|2030-05-19
    EMRE SOYDAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-09-18|2028-09-17
    EMRE SOYDAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-12|2027-02-11
    EMRE SOYDAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-09-18|2028-09-17
    EMRE SOYDAN|02-Temel İlk Yardım Belgesi|2023-09-18|2028-09-17
    EMRE SOYDAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-09-18|2028-09-17
    EMRE SOYDAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-09-18|2028-09-17
    EYÜP KAAN KANBUR|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-06-26|2030-06-25
    EYÜP KAAN KANBUR|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-11-28|2027-11-27
    BARIŞ ELİBAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-09|2026-12-08
    EYÜP KAAN KANBUR|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-11-28|2027-11-27
    EYÜP KAAN KANBUR|02-Temel İlk Yardım Belgesi|2022-12-15|2027-12-14
    EYÜP KAAN KANBUR|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-11-28|2027-11-27
    EYÜP KAAN KANBUR|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-06-26|2030-06-25
    EYÜP KAAN KANBUR|31-Gemi Adamı Cüzdan Belgesi|2022-12-15|2027-12-14
    EMRE BEKTAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-31|2029-10-30
    EMRE BEKTAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-31|2029-10-30
    EMRE BEKTAŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-05|2027-11-04
    EMRE BEKTAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-10-31|2029-10-30
    EMRE BEKTAŞ|02-Temel İlk Yardım Belgesi|2024-10-31|2029-10-30
    EMRE BEKTAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-31|2029-10-30
    EMRE BEKTAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-31|2029-10-30
    EMRE BEKTAŞ|31-Gemi Adamı Cüzdan Belgesi|2024-10-31|2029-10-30
    MUSTAFA ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-24|2029-10-23
    MUSTAFA ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-24|2029-10-23
    MUSTAFA ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-10-24|2029-10-23
    MUSTAFA ÖZTÜRK|02-Temel İlk Yardım Belgesi|2024-10-24|2029-10-23
    MUSTAFA ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-24|2029-10-23
    MUSTAFA ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-24|2029-10-23
    MUSTAFA ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2024-10-30|2029-10-29
    CAN ŞENGÜL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-05-29|2028-05-28
    CAN ŞENGÜL|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-05-29|2028-05-28
    CAN ŞENGÜL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-12|2027-05-11
    CAN ŞENGÜL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-05-29|2028-05-28
    CAN ŞENGÜL|02-Temel İlk Yardım Belgesi|2023-05-29|2028-05-28
    CAN ŞENGÜL|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-05-29|2028-05-28
    CAN ŞENGÜL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-07-04|2028-07-03
    CAN ŞENGÜL|31-Gemi Adamı Cüzdan Belgesi|2023-07-05|2028-07-04
    SERKAN KAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-13|2029-12-12
    SERKAN KAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-13|2029-12-12
    DENİZ ARSLANBAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-09|2026-12-08
    SERKAN KAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-13|2029-12-12
    SERKAN KAYA|02-Temel İlk Yardım Belgesi|2024-12-13|2029-12-12
    SERKAN KAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-13|2029-12-12
    SERKAN KAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-12-13|2029-12-12
    SERKAN KAYA|31-Gemi Adamı Cüzdan Belgesi|2024-12-16|2029-12-15
    İSMAİL OCAK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-08|2029-10-07
    İSMAİL OCAK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-08|2029-10-07
    İSMAİL OCAK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-10|2027-07-09
    İSMAİL OCAK|02-Temel İlk Yardım Belgesi|2024-10-08|2029-10-07
    İSMAİL OCAK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-08|2029-10-07
    DENİZ ARSLANBAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-09|2026-12-08
    İSMAİL OCAK|31-Gemi Adamı Cüzdan Belgesi|2025-09-03|2030-09-02
    DENİZ ARSLANBAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-09|2026-12-08
    DENİZ ARSLANBAŞ|02-Temel İlk Yardım Belgesi|2021-12-09|2026-12-08
    DENİZ ARSLANBAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-09|2026-12-08
    DENİZ ARSLANBAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-09|2026-12-08
    ADNAN CAN SELÇUK|31-Gemi Adamı Cüzdan Belgesi|2021-12-10|2026-12-09
    BARIŞ ELİBAL|31-Gemi Adamı Cüzdan Belgesi|2021-12-10|2026-12-09
    HAKAN AYDIN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-11|2026-12-10
    ERCAN DÖNMEZ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-12|2026-12-11
    EYÜP KAAN KANBUR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-12|2026-12-11
    YUNUS EMRE ZAMAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-07-13|2028-07-12
    YUNUS EMRE ZAMAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-07-13|2028-07-12
    RECEPALİ BAYRAKTAR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-13|2026-12-12
    YUNUS EMRE ZAMAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-07-13|2028-07-12
    YUNUS EMRE ZAMAN|02-Temel İlk Yardım Belgesi|2023-07-13|2028-07-12
    YUNUS EMRE ZAMAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-07-13|2028-07-12
    YUNUS EMRE ZAMAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-07-13|2028-07-12
    YUNUS EMRE ZAMAN|31-Gemi Adamı Cüzdan Belgesi|2023-07-13|2028-07-12
    FATİH KARAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-20|2029-12-19
    FATİH KARAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-08-31|2028-08-30
    FATİH KARAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-21|2027-03-20
    FATİH KARAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-08-31|2028-08-30
    FATİH KARAN|02-Temel İlk Yardım Belgesi|2023-08-31|2028-08-30
    FATİH KARAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-08-31|2028-08-30
    FATİH KARAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-04-15|2027-04-14
    FATİH KARAN|31-Gemi Adamı Cüzdan Belgesi|2023-08-31|2028-08-30
    HİLAL AY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-08|2029-10-07
    HİLAL AY|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-10-04|2027-10-03
    HİLAL AY|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-10-08|2029-10-07
    HİLAL AY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-10-04|2027-10-03
    HİLAL AY|02-Temel İlk Yardım Belgesi|2022-10-04|2027-10-03
    HİLAL AY|20-Seyir Vardiyası Tutma Belgesi|2024-10-08|2029-10-07
    HİLAL AY|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-10-04|2027-10-03
    HİLAL AY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-10|2027-01-09
    HİLAL AY|31-Gemi Adamı Cüzdan Belgesi|2022-10-06|2027-10-05
    AYHAN PİŞKİNER|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|10-İleri Yangınla Mücadele Belgesi|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|02-Temel İlk Yardım Belgesi|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-14|2026-12-13
    AYHAN PİŞKİNER|31-Gemi Adamı Cüzdan Belgesi|2021-12-14|2026-12-13
    HÜSEYİN ŞİŞMAN|31-Gemi Adamı Cüzdan Belgesi|2024-08-08|2029-08-07
    ÖMER İLBEY DOĞAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-11-04|2027-11-03
    ÖMER İLBEY DOĞAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-11-04|2027-11-03
    ÖMER İLBEY DOĞAN|02-Temel İlk Yardım Belgesi|2022-11-04|2027-11-03
    ÖMER İLBEY DOĞAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-11-04|2027-11-03
    ÖMER İLBEY DOĞAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-07-27|2028-07-26
    ÖMER İLBEY DOĞAN|31-Gemi Adamı Cüzdan Belgesi|2022-11-07|2027-11-06
    ALPASLAN AK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-12-05|2030-12-04
    ALPASLAN AK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-12-05|2030-12-04
    ALPASLAN AK|00-Gemiadamları Sağlık Yoklama Belgesi|2026-01-22|2028-01-21
    ALPASLAN AK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-12-05|2030-12-04
    ALPASLAN AK|02-Temel İlk Yardım Belgesi|2025-12-05|2030-12-04
    ALPASLAN AK|20-Seyir Vardiyası Tutma Belgesi|2025-12-05|2030-12-04
    ALPASLAN AK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-12-05|2030-12-04
    ALPASLAN AK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-30|2030-05-29
    ALPASLAN AK|31-Gemi Adamı Cüzdan Belgesi|2025-12-25|2030-12-24
    EFE GÖKBERK EMEKSİZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2026-02-27|2031-02-26
    EFE GÖKBERK EMEKSİZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-27|2031-02-26
    EFE GÖKBERK EMEKSİZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-22|2027-05-21
    EFE GÖKBERK EMEKSİZ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-02-27|2031-02-26
    EFE GÖKBERK EMEKSİZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-27|2031-02-26
    EFE GÖKBERK EMEKSİZ|02-Temel İlk Yardım Belgesi|2026-02-27|2031-02-26
    EFE GÖKBERK EMEKSİZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-27|2031-02-26
    EFE GÖKBERK EMEKSİZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-08-31|2028-08-30
    EFE GÖKBERK EMEKSİZ|31-Gemi Adamı Cüzdan Belgesi|2026-02-28|2031-02-27
    ATALAY KOÇALAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-08-08|2028-08-07
    ATALAY KOÇALAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-11|2027-02-10
    ATALAY KOÇALAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-07|2027-07-06
    ATALAY KOÇALAN|10-İleri Yangınla Mücadele Belgesi|2025-11-05|2030-11-04
    ATALAY KOÇALAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-11|2027-02-10
    ATALAY KOÇALAN|02-Temel İlk Yardım Belgesi|2022-02-11|2027-02-10
    ATALAY KOÇALAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-11|2027-02-10
    ATALAY KOÇALAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-11|2027-02-10
    ATALAY KOÇALAN|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-05-17|2029-05-16
    ATALAY KOÇALAN|31-Gemi Adamı Cüzdan Belgesi|2022-02-11|2027-02-10
    YASİN ECEVİT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-05-24|2028-05-23
    YASİN ECEVİT|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-12-06|2029-12-05
    YASİN ECEVİT|00-Gemiadamları Sağlık Yoklama Belgesi|2025-05-21|2027-05-20
    YASİN ECEVİT|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2023-09-06|2028-09-05
    YASİN ECEVİT|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-12-06|2029-12-05
    YASİN ECEVİT|02-Temel İlk Yardım Belgesi|2024-12-06|2029-12-05
    YASİN ECEVİT|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-12-06|2029-12-05
    YASİN ECEVİT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-10-27|2027-10-26
    YASİN ECEVİT|31-Gemi Adamı Cüzdan Belgesi|2024-12-06|2029-12-05
    ENVER MUŞDAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-04-29|2030-04-28
    ENVER MUŞDAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-04-29|2030-04-28
    ENVER MUŞDAL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-18|2027-07-17
    ENVER MUŞDAL|10-İleri Yangınla Mücadele Belgesi|2025-04-29|2030-04-28
    ENVER MUŞDAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-04-29|2030-04-28
    ENVER MUŞDAL|02-Temel İlk Yardım Belgesi|2025-04-29|2030-04-28
    ENVER MUŞDAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-04-29|2030-04-28
    ENVER MUŞDAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-04-18|2029-04-17
    ENVER MUŞDAL|24-Genel Telsiz Operatörü (GOC) Belgesi|2023-10-02|2028-10-01
    ENVER MUŞDAL|31-Gemi Adamı Cüzdan Belgesi|2025-05-12|2030-05-11
    BURAK GEYİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-06-05|2028-06-04
    YUSUF YASİN KADEM SARI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-14|2026-12-13
    BURAK GEYİK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-04|2027-06-03
    YUSUF YASİN KADEM SARI|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-14|2026-12-13
    YUSUF YASİN KADEM SARI|10-İleri Yangınla Mücadele Belgesi|2021-12-14|2026-12-13
    BURAK GEYİK|20-Seyir Vardiyası Tutma Belgesi|2023-06-05|2028-06-04
    YUSUF YASİN KADEM SARI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-14|2026-12-13
    BURAK GEYİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-07-07|2027-07-06
    BURAK GEYİK|31-Gemi Adamı Cüzdan Belgesi|2022-07-07|2027-07-06
    KENAN ÖLMEZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-03-23|2027-03-22
    KENAN ÖLMEZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-03-23|2027-03-22
    KENAN ÖLMEZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-03-23|2027-03-22
    KENAN ÖLMEZ|02-Temel İlk Yardım Belgesi|2022-03-23|2027-03-22
    KENAN ÖLMEZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-03-23|2027-03-22
    KENAN ÖLMEZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-09-27|2028-09-26
    KENAN ÖLMEZ|31-Gemi Adamı Cüzdan Belgesi|2022-03-23|2027-03-22
    YASİN GÜNDOĞDU|00-Gemiadamları Sağlık Yoklama Belgesi|2025-07-22|2027-07-21
    YASİN GÜNDOĞDU|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-08-23|2028-08-22
    ENGİN ATILGAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-10-15|2029-10-14
    ENGİN ATILGAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-10-15|2029-10-14
    ENGİN ATILGAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-19|2027-09-18
    ENGİN ATILGAN|02-Temel İlk Yardım Belgesi|2024-10-15|2029-10-14
    ENGİN ATILGAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-10-15|2029-10-14
    ENGİN ATILGAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-15|2029-10-14
    ENGİN ATILGAN|31-Gemi Adamı Cüzdan Belgesi|2022-06-01|2027-05-31
    ÇAĞLAR YALIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-08|2030-05-07
    ÇAĞLAR YALIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-05-08|2030-05-07
    ÇAĞLAR YALIN|02-Temel İlk Yardım Belgesi|2025-05-08|2030-05-07
    ÇAĞLAR YALIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-05-08|2030-05-07
    ÇAĞLAR YALIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-08|2030-05-07
    ÇAĞLAR YALIN|31-Gemi Adamı Cüzdan Belgesi|2025-05-13|2030-05-12
    SAMET SARIÇİÇEK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-10-19|2027-10-18
    SAMET SARIÇİÇEK|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-10-19|2027-10-18
    SAMET SARIÇİÇEK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-09|2027-09-08
    SAMET SARIÇİÇEK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-10-19|2027-10-18
    SAMET SARIÇİÇEK|02-Temel İlk Yardım Belgesi|2022-10-19|2027-10-18
    SAMET SARIÇİÇEK|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-10-19|2027-10-18
    SAMET SARIÇİÇEK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-10-19|2027-10-18
    SAMET SARIÇİÇEK|31-Gemi Adamı Cüzdan Belgesi|2022-10-20|2027-10-19
    HALİL ÖZSEVİNÇ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-22|2030-05-20
    HALİL ÖZSEVİNÇ|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-02-21|2027-02-20
    HALİL ÖZSEVİNÇ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-02-21|2027-02-20
    HALİL ÖZSEVİNÇ|02-Temel İlk Yardım Belgesi|2022-02-21|2027-02-20
    HALİL ÖZSEVİNÇ|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-02-21|2027-02-20
    HALİL ÖZSEVİNÇ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-02-21|2027-02-20
    HALİL ÖZSEVİNÇ|31-Gemi Adamı Cüzdan Belgesi|2022-02-21|2027-02-20
    BERK TURAN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-06-11|2027-06-10
    BERK TURAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-07-24|2028-07-23
    ARSLAN YUSUF KATIRCI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-12-25|2027-12-24
    ARSLAN YUSUF KATIRCI|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-03|2027-02-02
    ARSLAN YUSUF KATIRCI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-01-26|2028-01-25
    CANSU AKBAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-13|2029-12-12
    CANSU AKBAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-02|2028-02-01
    YUSUF YASİN KADEM SARI|02-Temel İlk Yardım Belgesi|2021-12-14|2026-12-13
    CANSU AKBAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-02|2028-02-01
    CANSU AKBAŞ|02-Temel İlk Yardım Belgesi|2023-02-02|2028-02-01
    CANSU AKBAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-02|2028-02-01
    CANSU AKBAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-07-27|2028-07-26
    CANSU AKBAŞ|31-Gemi Adamı Cüzdan Belgesi|2023-02-08|2028-02-07
    HARUN ANGUN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-12-02|2029-12-01
    HARUN ANGUN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-09-13|2028-09-12
    HARUN ANGUN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-09-18|2027-09-16
    HARUN ANGUN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-09-13|2028-09-12
    HARUN ANGUN|02-Temel İlk Yardım Belgesi|2023-09-13|2028-09-12
    HARUN ANGUN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-09-13|2028-09-12
    HARUN ANGUN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-09-13|2028-09-12
    HARUN ANGUN|31-Gemi Adamı Cüzdan Belgesi|2023-09-14|2028-09-13
    MUSTAFA BURAK ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-17|2029-09-16
    MUSTAFA BURAK ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-17|2029-09-16
    YUSUF YASİN KADEM SARI|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-14|2026-12-13
    MUSTAFA BURAK ÖZTÜRK|10-İleri Yangınla Mücadele Belgesi|2024-09-17|2029-09-16
    MUSTAFA BURAK ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-17|2029-09-16
    MUSTAFA BURAK ÖZTÜRK|02-Temel İlk Yardım Belgesi|2024-09-17|2029-09-16
    MUSTAFA BURAK ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-17|2029-09-16
    YUSUF YASİN KADEM SARI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-14|2026-12-13
    MUSTAFA BURAK ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2022-01-12|2027-01-11
    ALİ RIZA CAN HANTAL|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-20|2027-09-19
    ALİ RIZA CAN HANTAL|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-01-13|2027-01-12
    ALİ RIZA CAN HANTAL|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-07|2027-11-06
    ALİ RIZA CAN HANTAL|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-01-13|2027-01-12
    ALİ RIZA CAN HANTAL|02-Temel İlk Yardım Belgesi|2022-01-13|2027-01-12
    ALİ RIZA CAN HANTAL|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-01-13|2027-01-12
    ALİ RIZA CAN HANTAL|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2022-01-13|2027-01-12
    ALİ RIZA CAN HANTAL|31-Gemi Adamı Cüzdan Belgesi|2022-01-13|2027-01-12
    YUSUF CANDAN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-23|2027-09-22
    YUSUF CANDAN|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-23|2027-09-22
    YUSUF YASİN KADEM SARI|31-Gemi Adamı Cüzdan Belgesi|2021-12-14|2026-12-13
    YUSUF CANDAN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-23|2027-09-22
    YUSUF CANDAN|02-Temel İlk Yardım Belgesi|2022-09-23|2027-09-22
    YUSUF CANDAN|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-23|2027-09-22
    YUSUF CANDAN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-04-16|2029-04-15
    YUSUF CANDAN|31-Gemi Adamı Cüzdan Belgesi|2022-09-06|2027-09-05
    ÇAĞLA ÖZKAYA|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-04-27|2028-04-26
    ÇAĞLA ÖZKAYA|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-04-27|2028-04-26
    YUSUF YASİN KADEM SARI|24-Genel Telsiz Operatörü (GOC) Belgesi|2021-12-15|2026-12-14
    ÇAĞLA ÖZKAYA|10-İleri Yangınla Mücadele Belgesi|2023-06-19|2028-06-18
    ÇAĞLA ÖZKAYA|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-04-27|2028-04-26
    ÇAĞLA ÖZKAYA|02-Temel İlk Yardım Belgesi|2023-04-27|2028-04-26
    ÇAĞLA ÖZKAYA|20-Seyir Vardiyası Tutma Belgesi|2023-04-27|2028-04-26
    ÇAĞLA ÖZKAYA|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-04-27|2028-04-26
    ÇAĞLA ÖZKAYA|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-03-02|2031-03-01
    ÇAĞLA ÖZKAYA|31-Gemi Adamı Cüzdan Belgesi|2025-12-22|2030-12-21
    MUHAMMED ALİ ERTİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-07-06|2028-07-05
    MUHAMMED ALİ ERTİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-09-07|2030-09-06
    BAHATTİN ÇELİK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2021-12-16|2026-12-15
    MUHAMMED ALİ ERTİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-09-07|2030-09-06
    MUHAMMED ALİ ERTİK|02-Temel İlk Yardım Belgesi|2025-09-07|2030-09-06
    MUHAMMED ALİ ERTİK|20-Seyir Vardiyası Tutma Belgesi|2023-07-06|2028-07-05
    MUHAMMED ALİ ERTİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-09-07|2030-09-06
    MUHAMMED ALİ ERTİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-10-17|2029-10-16
    MUHAMMED ALİ ERTİK|31-Gemi Adamı Cüzdan Belgesi|2025-09-07|2030-09-06
    BAHATTİN ÇELİK|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-16|2026-12-15
    BAHATTİN ÇELİK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-16|2026-12-15
    BAHATTİN ÇELİK|02-Temel İlk Yardım Belgesi|2021-12-16|2026-12-15
    BAHATTİN ÇELİK|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-16|2026-12-15
    BAHATTİN ÇELİK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-16|2026-12-15
    BAHATTİN ÇELİK|31-Gemi Adamı Cüzdan Belgesi|2021-12-16|2026-12-15
    CUMHUR KAZMAZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-17|2026-12-16
    GÜLSÜM İNANÇ KORKMAZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2021-12-17|2026-12-16
    HALİT ŞAHİN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-05-27|2030-05-26
    HALİT ŞAHİN|01-Denizde Kişisel Can Kurtarma Teknikl.|2026-02-27|2031-02-26
    HALİT ŞAHİN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2026-02-27|2031-02-26
    HALİT ŞAHİN|02-Temel İlk Yardım Belgesi|2026-02-27|2031-02-26
    HALİT ŞAHİN|03-Yangın Önleme ve Yangınla Mücadele Bl|2026-02-27|2031-02-26
    HALİT ŞAHİN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2026-02-27|2031-02-26
    HALİT ŞAHİN|13-Petrol Tankerleri İşlemleri Belgesi|2026-02-27|2031-02-26
    HALİT ŞAHİN|31-Gemi Adamı Cüzdan Belgesi|2026-03-02|2031-03-01
    NURETTİN AYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-07-06|2028-07-05
    NURETTİN AYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-07-06|2028-07-05
    NURETTİN AYDIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-07-06|2028-07-05
    NURETTİN AYDIN|02-Temel İlk Yardım Belgesi|2023-07-06|2028-07-05
    NURETTİN AYDIN|20-Seyir Vardiyası Tutma Belgesi|2023-07-06|2028-07-05
    NURETTİN AYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-07-06|2028-07-05
    NURETTİN AYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-10-27|2028-10-26
    GÜLSÜM İNANÇ KORKMAZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2021-12-17|2026-12-16
    SERKAN MAMAT|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-11-20|2029-11-19
    SERKAN MAMAT|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-11-20|2029-11-19
    GÜLSÜM İNANÇ KORKMAZ|02-Temel İlk Yardım Belgesi|2021-12-17|2026-12-16
    SERKAN MAMAT|10-İleri Yangınla Mücadele Belgesi|2024-11-20|2029-11-19
    SERKAN MAMAT|02-Temel İlk Yardım Belgesi|2024-11-20|2029-11-19
    SERKAN MAMAT|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-11-20|2029-11-19
    SERKAN MAMAT|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-11-20|2029-11-19
    SERKAN MAMAT|31-Gemi Adamı Cüzdan Belgesi|2024-12-26|2029-12-25
    GÖKTUĞ AY|00-Gemiadamları Sağlık Yoklama Belgesi|2025-03-07|2027-03-06
    GÖKTUĞ AY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-05-30|2030-05-29
    GÜLSÜM İNANÇ KORKMAZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2021-12-17|2026-12-16
    MESUT SOY|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-11-14|2028-11-13
    RAMAZAN ÖZTÜRK|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-10-09|2027-09-16
    RAMAZAN ÖZTÜRK|10-İleri Yangınla Mücadele Belgesi|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|02-Temel İlk Yardım Belgesi|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-08-04|2030-08-03
    RAMAZAN ÖZTÜRK|31-Gemi Adamı Cüzdan Belgesi|2025-08-03|2030-08-02
    FATİH GÜVELİ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-05-26|2027-05-25
    FATİH GÜVELİ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-06-07|2028-06-06
    FATİH GÜVELİ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2024-09-17|2029-09-16
    FATİH GÜVELİ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-06-07|2028-06-06
    FATİH GÜVELİ|02-Temel İlk Yardım Belgesi|2023-06-07|2028-06-06
    FATİH GÜVELİ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-06-07|2028-06-06
    FATİH GÜVELİ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-04-19|2028-04-18
    FATİH GÜVELİ|31-Gemi Adamı Cüzdan Belgesi|2023-06-02|2028-06-01
    SERKAN SELİM CANBAZ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-02-01|2028-01-31
    SERKAN SELİM CANBAZ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-01|2028-01-31
    SERKAN SELİM CANBAZ|00-Gemiadamları Sağlık Yoklama Belgesi|2025-02-19|2027-02-18
    SERKAN SELİM CANBAZ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-01|2028-01-31
    SERKAN SELİM CANBAZ|02-Temel İlk Yardım Belgesi|2023-02-01|2028-01-31
    SERKAN SELİM CANBAZ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-01|2028-01-31
    SERKAN SELİM CANBAZ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2023-02-01|2028-01-31
    SERKAN SELİM CANBAZ|31-Gemi Adamı Cüzdan Belgesi|2023-02-02|2028-02-01
    CEMAL ERDEM|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2024-09-23|2029-09-22
    CEMAL ERDEM|01-Denizde Kişisel Can Kurtarma Teknikl.|2024-09-09|2029-09-08
    CEMAL ERDEM|04-Personel Güvenliği ve Sosyal Sor.Bl.|2024-09-09|2029-09-08
    CEMAL ERDEM|02-Temel İlk Yardım Belgesi|2024-09-09|2029-09-08
    CEMAL ERDEM|03-Yangın Önleme ve Yangınla Mücadele Bl|2024-09-09|2029-09-08
    CEMAL ERDEM|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-09-09|2029-09-08
    CEMAL ERDEM|25-Kısa Mesafe Telsiz Operatörü Belgesi|2024-09-04|2029-09-03
    MELİKE KABAK|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-12-25|2028-12-24
    MELİKE KABAK|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-21|2027-11-20
    MELİKE KABAK|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2026-03-03|2031-03-02
    MELİKE KABAK|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-12-25|2028-12-24
    MELİKE KABAK|02-Temel İlk Yardım Belgesi|2023-12-25|2028-12-24
    MELİKE KABAK|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-12-25|2028-12-24
    MELİKE KABAK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-01-15|2029-01-14
    MELİKE KABAK|31-Gemi Adamı Cüzdan Belgesi|2023-12-25|2028-12-24
    BAHTİYAR GÜNEY|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2022-09-14|2027-09-13
    BAHTİYAR GÜNEY|01-Denizde Kişisel Can Kurtarma Teknikl.|2022-09-14|2027-09-13
    BAHTİYAR GÜNEY|10-İleri Yangınla Mücadele Belgesi|2024-01-23|2029-01-22
    BAHTİYAR GÜNEY|04-Personel Güvenliği ve Sosyal Sor.Bl.|2022-09-14|2027-09-13
    BAHTİYAR GÜNEY|02-Temel İlk Yardım Belgesi|2022-09-14|2027-09-13
    BAHTİYAR GÜNEY|03-Yangın Önleme ve Yangınla Mücadele Bl|2022-09-14|2027-09-13
    BAHTİYAR GÜNEY|24-Genel Telsiz Operatörü (GOC) Belgesi|2024-01-23|2029-01-22
    FURGAN GÜNÇALDI|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-07-31|2028-07-20
    FURGAN GÜNÇALDI|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-07-31|2028-07-20
    GÜLSÜM İNANÇ KORKMAZ|31-Gemi Adamı Cüzdan Belgesi|2021-12-17|2026-12-16
    FURGAN GÜNÇALDI|10-İleri Yangınla Mücadele Belgesi|2023-07-31|2028-07-26
    FURGAN GÜNÇALDI|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-07-31|2028-07-20
    FURGAN GÜNÇALDI|02-Temel İlk Yardım Belgesi|2023-07-31|2028-07-20
    FURGAN GÜNÇALDI|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-07-31|2028-07-20
    FURGAN GÜNÇALDI|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2024-08-19|2029-08-18
    FURGAN GÜNÇALDI|31-Gemi Adamı Cüzdan Belgesi|2023-08-01|2028-07-31
    ÜMİT İRİTAŞ|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2023-02-15|2028-02-14
    ÜMİT İRİTAŞ|01-Denizde Kişisel Can Kurtarma Teknikl.|2023-02-15|2028-02-14
    ÜMİT İRİTAŞ|10-İleri Yangınla Mücadele Belgesi|2023-02-15|2028-02-14
    ÜMİT İRİTAŞ|04-Personel Güvenliği ve Sosyal Sor.Bl.|2023-02-15|2028-02-14
    ÜMİT İRİTAŞ|02-Temel İlk Yardım Belgesi|2023-02-15|2028-02-14
    ÜMİT İRİTAŞ|03-Yangın Önleme ve Yangınla Mücadele Bl|2023-02-15|2028-02-14
    ÜMİT İRİTAŞ|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-03-10|2030-03-09
    EYÜP ÇELİKKOL|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-18|2026-12-17
    ÜMİT İRİTAŞ|31-Gemi Adamı Cüzdan Belgesi|2023-02-15|2028-02-14
    MUSTAFA BURAK ÖZTÜRK|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-20|2026-12-19
    MEHMET SEFER|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-20|2026-12-19
    CEYHUN PIRLANT|25-Kısa Mesafe Telsiz Operatörü Belgesi|2021-12-21|2026-12-20
    HAMİT BİLGİLİ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-23|2026-12-22
    MEHMET LOR|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-26|2026-12-25
    CANSU AKBAŞ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-26|2026-12-25
    SERKAN AYTAÇ|23-GMDSS Sınırlı Telsiz Opert.Blg.(ROC)|2021-12-27|2026-12-26
    KUDRET GİRGİN|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-27|2026-12-26
    MEHMET EKİCİ|00-Gemiadamları Sağlık Yoklama Belgesi|2024-12-30|2026-12-29
    ONUR AYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2021-12-31|2026-12-30
    NAZİF İLYAS|31-Gemi Adamı Cüzdan Belgesi|2024-10-30|2029-10-29
    GÖKHAN AYDIN|05-Cankurtarma Araçlarını Kullanm.Yt.Bl.|2025-10-03|2030-10-02
    GÖKHAN AYDIN|01-Denizde Kişisel Can Kurtarma Teknikl.|2025-10-03|2030-10-02
    GÖKHAN AYDIN|00-Gemiadamları Sağlık Yoklama Belgesi|2025-11-25|2027-11-24
    GÖKHAN AYDIN|10-İleri Yangınla Mücadele Belgesi|2025-10-03|2030-10-02
    GÖKHAN AYDIN|04-Personel Güvenliği ve Sosyal Sor.Bl.|2025-10-03|2030-10-02
    GÖKHAN AYDIN|02-Temel İlk Yardım Belgesi|2025-10-03|2030-10-02
    GÖKHAN AYDIN|03-Yangın Önleme ve Yangınla Mücadele Bl|2025-10-03|2030-10-02
    GÖKHAN AYDIN|18-Yolcu Gemisinde Çalışma Yeterlil.Blg.|2025-03-20|2030-03-19
    GÖKHAN AYDIN|31-Gemi Adamı Cüzdan Belgesi|2025-10-07|2030-10-06"""

    for line in raw_text.strip().split('\n'):
        parts = line.strip().split('|')
        if len(parts) >= 4:
            name = parts[0].strip()
            cert = parts[1].strip()
            start = parts[2].strip()
            end = parts[3].strip()
            data.append([name, cert, start, end])
    
    df = pd.DataFrame(data, columns=['Ad', 'Nitelik', 'Başlangıç tarihi', 'Bitiş tarihi'])
    
    # Tarihleri dönüştür
    df['Başlangıç tarihi'] = pd.to_datetime(df['Başlangıç tarihi'])
    df['Bitiş tarihi'] = pd.to_datetime(df['Bitiş tarihi'])
    
    # Bugünün tarihi
    today = datetime.now().date()
    
    # Kalan gün hesapla
    df['Kalan Gün'] = (df['Bitiş tarihi'].dt.date - today).dt.days
    
    # Durum belirleme
    def get_status(days):
        if days < 0:
            return 'Süresi Doldu'
        elif days <= 30:
            return 'Kritik'
        elif days <= 60:
            return 'Uyarı'
        elif days <= 90:
            return 'Yaklaşıyor'
        else:
            return 'Geçerli'
    
    df['Durum'] = df['Kalan Gün'].apply(get_status)
    
    # Renk kodları
    def get_color(days):
        if days < 0:
            return '⚫'  # Siyah - Süresi doldu
        elif days <= 30:
            return '🔴'  # Kırmızı - Kritik
        elif days <= 60:
            return '🟠'  # Turuncu - Uyarı
        elif days <= 90:
            return '🟡'  # Sarı - Yaklaşıyor
        else:
            return '🟢'  # Yeşil - Geçerli
    
    df['Renk'] = df['Kalan Gün'].apply(get_color)
    
    # Trafik lambası metni
    def get_traffic_text(days):
        if days < 0:
            return 'Süresi Doldu'
        elif days <= 30:
            return 'Kritik (0-30 gün)'
        elif days <= 60:
            return 'Uyarı (31-60 gün)'
        elif days <= 90:
            return 'Yaklaşıyor (61-90 gün)'
        else:
            return 'Geçerli (90+ gün)'
    
    df['Trafik Durumu'] = df['Kalan Gün'].apply(get_traffic_text)
    
    # Ek alanlar (kullanıcı tarafından doldurulacak)
    df['Bildirildi mi?'] = 'Hayır'
    df['Bildirim Tarihi'] = ''
    df['Sorumlu Kişi'] = ''
    df['Notlar'] = ''
    df['Arşiv'] = False
    
    return df

# ============================================================
# SESSION STATE YÖNETİMİ
# ============================================================

def init_session_state():
    """Session state başlatma"""
    if 'data' not in st.session_state:
        st.session_state.data = load_data()
    if 'selected_person' not in st.session_state:
        st.session_state.selected_person = None
    if 'filter_status' not in st.session_state:
        st.session_state.filter_status = 'Tümü'
    if 'search_term' not in st.session_state:
        st.session_state.search_term = ''
    if 'show_archived' not in st.session_state:
        st.session_state.show_archived = False
    if 'notification_history' not in st.session_state:
        st.session_state.notification_history = []

# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

def get_person_list(df):
    """Kişi listesini döndür"""
    return sorted(df['Ad'].unique())

def filter_data(df, status_filter='Tümü', search_term='', show_archived=False):
    """Veriyi filtrele"""
    filtered = df.copy()
    
    # Arşiv filtresi
    if not show_archived:
        filtered = filtered[filtered['Arşiv'] == False]
    
    # Durum filtresi
    if status_filter != 'Tümü':
        filtered = filtered[filtered['Durum'] == status_filter]
    
    # Arama filtresi
    if search_term:
        filtered = filtered[filtered['Ad'].str.contains(search_term, case=False, na=False)]
    
    return filtered

def get_status_counts(df):
    """Durum sayılarını döndür"""
    counts = df['Durum'].value_counts().to_dict()
    total = len(df)
    return counts, total

def get_expiring_by_month(df):
    """Aylara göre sona eren belge sayısı"""
    df_copy = df.copy()
    df_copy['Bitiş Ayı'] = df_copy['Bitiş tarihi'].dt.strftime('%Y-%m')
    monthly = df_copy.groupby('Bitiş Ayı').size().reset_index(name='Adet')
    monthly = monthly.sort_values('Bitiş Ayı')
    return monthly

def get_person_certificates(df, person_name):
    """Belirli bir kişinin tüm belgelerini döndür"""
    return df[df['Ad'] == person_name]

def get_traffic_light(days):
    """Trafik lambası emojisi"""
    if days < 0:
        return '⚫'
    elif days <= 30:
        return '🔴'
    elif days <= 60:
        return '🟠'
    elif days <= 90:
        return '🟡'
    else:
        return '🟢'

# ============================================================
# SAYFALAR
# ============================================================

def page_dashboard(df):
    """Dashboard sayfası"""
    st.markdown('<h1 class="main-header">📊 Yönetici Dashboard</h1>', unsafe_allow_html=True)
    
    # Bugünün tarihi
    today = datetime.now().date()
    st.info(f"📅 Bugün: {today.strftime('%d.%m.%Y')}")
    
    # İstatistikler
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total_docs = len(df)
    expired = len(df[df['Kalan Gün'] < 0])
    critical = len(df[(df['Kalan Gün'] >= 0) & (df['Kalan Gün'] <= 30)])
    warning = len(df[(df['Kalan Gün'] > 30) & (df['Kalan Gün'] <= 60)])
    upcoming = len(df[(df['Kalan Gün'] > 60) & (df['Kalan Gün'] <= 90)])
    valid = len(df[df['Kalan Gün'] > 90])
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{total_docs}</div>
            <div class="stat-label">📄 Toplam Belge</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="stat-card" style="background:#fee2e2;">
            <div class="stat-number" style="color:#dc2626;">{expired}</div>
            <div class="stat-label">❌ Süresi Dolan</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="stat-card" style="background:#fef3c7;">
            <div class="stat-number" style="color:#d97706;">{critical}</div>
            <div class="stat-label">🔴 Kritik (0-30 gün)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="stat-card" style="background:#fde68a;">
            <div class="stat-number" style="color:#b45309;">{warning}</div>
            <div class="stat-label">🟠 Uyarı (31-60 gün)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class="stat-card" style="background:#d1fae5;">
            <div class="stat-number" style="color:#059669;">{valid}</div>
            <div class="stat-label">🟢 Geçerli (90+ gün)</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    # Grafikler
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📈 Durum Dağılımı")
        status_counts = df['Durum'].value_counts().reset_index()
        status_counts.columns = ['Durum', 'Adet']
        
        color_map = {
            'Süresi Doldu': '#e74c3c',
            'Kritik': '#e67e22',
            'Uyarı': '#f1c40f',
            'Yaklaşıyor': '#f39c12',
            'Geçerli': '#27ae60'
        }
        
        fig = px.pie(
            status_counts, 
            values='Adet', 
            names='Durum',
            color='Durum',
            color_discrete_map=color_map,
            title='Belge Durum Dağılımı'
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📅 Aylara Göre Sona Eren Belgeler")
        monthly = get_expiring_by_month(df)
        if len(monthly) > 0:
            fig = px.bar(
                monthly,
                x='Bitiş Ayı',
                y='Adet',
                title='Aylara Göre Sona Eren Belge Sayısı',
                color='Adet',
                color_continuous_scale='Reds'
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sona eren belge bulunmuyor.")
    
    st.divider()
    
    # Yakında sona erecekler
    st.subheader("⏳ Yakında Sona Erecek Belgeler (90 gün ve altı)")
    
    expiring_soon = df[(df['Kalan Gün'] >= 0) & (df['Kalan Gün'] <= 90)].sort_values('Kalan Gün')
    
    if len(expiring_soon) > 0:
        display_df = expiring_soon[['Ad', 'Nitelik', 'Bitiş tarihi', 'Kalan Gün', 'Renk', 'Trafik Durumu']].copy()
        display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
        display_df.columns = ['Ad', 'Nitelik', 'Bitiş Tarihi', 'Kalan Gün', '', 'Durum']
        
        # Renkli gösterim için
        st.dataframe(
            display_df,
            column_config={
                'Ad': st.column_config.TextColumn('Ad'),
                'Nitelik': st.column_config.TextColumn('Nitelik'),
                'Bitiş Tarihi': st.column_config.TextColumn('Bitiş Tarihi'),
                'Kalan Gün': st.column_config.NumberColumn('Kalan Gün'),
                '': st.column_config.TextColumn(''),
                'Durum': st.column_config.TextColumn('Durum'),
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.success("✅ Yakında sona erecek belge bulunmuyor.")

def page_person_card(df):
    """Personel kartı sayfası"""
    st.markdown('<h1 class="main-header">👤 Personel Kartı</h1>', unsafe_allow_html=True)
    
    persons = get_person_list(df)
    
    col1, col2 = st.columns([1, 3])
    
    with col1:
        selected = st.selectbox(
            "👤 Personel Seçin",
            options=[''] + persons,
            format_func=lambda x: x if x else 'Seçiniz...'
        )
        
        if selected:
            st.session_state.selected_person = selected
    
    with col2:
        if st.session_state.selected_person:
            person = st.session_state.selected_person
            person_data = get_person_certificates(df, person)
            
            st.subheader(f"📋 {person} - Belgeler")
            
            # İstatistikler
            total = len(person_data)
            expired = len(person_data[person_data['Kalan Gün'] < 0])
            critical = len(person_data[(person_data['Kalan Gün'] >= 0) & (person_data['Kalan Gün'] <= 30)])
            
            c1, c2, c3 = st.columns(3)
            c1.metric("📄 Toplam Belge", total)
            c2.metric("❌ Süresi Dolan", expired)
            c3.metric("🔴 Kritik", critical)
            
            st.divider()
            
            # Belgeleri göster
            display_df = person_data[['Nitelik', 'Başlangıç tarihi', 'Bitiş tarihi', 'Kalan Gün', 'Renk', 'Trafik Durumu', 'Bildirildi mi?', 'Sorumlu Kişi', 'Notlar']].copy()
            display_df['Başlangıç tarihi'] = display_df['Başlangıç tarihi'].dt.strftime('%d.%m.%Y')
            display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
            display_df.columns = ['Nitelik', 'Başlangıç', 'Bitiş', 'Kalan Gün', '', 'Durum', 'Bildirildi mi?', 'Sorumlu Kişi', 'Notlar']
            
            st.dataframe(
                display_df,
                column_config={
                    'Nitelik': st.column_config.TextColumn('Nitelik', width='large'),
                    'Başlangıç': st.column_config.TextColumn('Başlangıç'),
                    'Bitiş': st.column_config.TextColumn('Bitiş'),
                    'Kalan Gün': st.column_config.NumberColumn('Kalan Gün'),
                    '': st.column_config.TextColumn(''),
                    'Durum': st.column_config.TextColumn('Durum'),
                    'Bildirildi mi?': st.column_config.TextColumn('Bildirildi mi?'),
                    'Sorumlu Kişi': st.column_config.TextColumn('Sorumlu Kişi'),
                    'Notlar': st.column_config.TextColumn('Notlar'),
                },
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("👈 Lütfen soldan bir personel seçin.")

def page_certificates(df):
    """Belgeler listesi sayfası"""
    st.markdown('<h1 class="main-header">📋 Tüm Belgeler</h1>', unsafe_allow_html=True)
    
    # Filtreler
    col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
    
    with col1:
        status_filter = st.selectbox(
            "📊 Durum Filtresi",
            options=['Tümü', 'Geçerli', 'Yaklaşıyor', 'Uyarı', 'Kritik', 'Süresi Doldu']
        )
        st.session_state.filter_status = status_filter
    
    with col2:
        search_term = st.text_input("🔍 İsim Ara", value=st.session_state.search_term)
        st.session_state.search_term = search_term
    
    with col3:
        show_archived = st.checkbox("📦 Arşivlenenleri Göster", value=st.session_state.show_archived)
        st.session_state.show_archived = show_archived
    
    with col4:
        if st.button("🔄 Sıfırla"):
            st.session_state.filter_status = 'Tümü'
            st.session_state.search_term = ''
            st.session_state.show_archived = False
            st.rerun()
    
    # Filtrele
    filtered = filter_data(df, status_filter, search_term, show_archived)
    
    st.caption(f"📊 Toplam {len(filtered)} belge gösteriliyor.")
    
    # Veriyi göster
    display_df = filtered[['Ad', 'Nitelik', 'Bitiş tarihi', 'Kalan Gün', 'Renk', 'Trafik Durumu', 'Bildirildi mi?', 'Sorumlu Kişi', 'Notlar', 'Arşiv']].copy()
    display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
    display_df.columns = ['Ad', 'Nitelik', 'Bitiş Tarihi', 'Kalan Gün', '', 'Durum', 'Bildirildi mi?', 'Sorumlu Kişi', 'Notlar', 'Arşiv']
    
    st.dataframe(
        display_df,
        column_config={
            'Ad': st.column_config.TextColumn('Ad', width='medium'),
            'Nitelik': st.column_config.TextColumn('Nitelik', width='large'),
            'Bitiş Tarihi': st.column_config.TextColumn('Bitiş Tarihi'),
            'Kalan Gün': st.column_config.NumberColumn('Kalan Gün'),
            '': st.column_config.TextColumn(''),
            'Durum': st.column_config.TextColumn('Durum'),
            'Bildirildi mi?': st.column_config.TextColumn('Bildirildi mi?'),
            'Sorumlu Kişi': st.column_config.TextColumn('Sorumlu Kişi'),
            'Notlar': st.column_config.TextColumn('Notlar'),
            'Arşiv': st.column_config.CheckboxColumn('Arşiv'),
        },
        use_container_width=True,
        hide_index=True
    )

def page_upcoming(df):
    """Yaklaşan belgeler sayfası"""
    st.markdown('<h1 class="main-header">⏳ Yaklaşan Belgeler</h1>', unsafe_allow_html=True)
    
    # 90 gün ve altı
    upcoming_df = df[(df['Kalan Gün'] >= 0) & (df['Kalan Gün'] <= 90)].sort_values('Kalan Gün')
    
    if len(upcoming_df) > 0:
        st.info(f"📌 {len(upcoming_df)} belge yakında sona eriyor.")
        
        # 15, 30, 60, 90 gün grupları
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            count_15 = len(upcoming_df[upcoming_df['Kalan Gün'] <= 15])
            st.metric("🔴 0-15 Gün", count_15)
        with col2:
            count_30 = len(upcoming_df[(upcoming_df['Kalan Gün'] > 15) & (upcoming_df['Kalan Gün'] <= 30)])
            st.metric("🟠 16-30 Gün", count_30)
        with col3:
            count_60 = len(upcoming_df[(upcoming_df['Kalan Gün'] > 30) & (upcoming_df['Kalan Gün'] <= 60)])
            st.metric("🟡 31-60 Gün", count_60)
        with col4:
            count_90 = len(upcoming_df[(upcoming_df['Kalan Gün'] > 60) & (upcoming_df['Kalan Gün'] <= 90)])
            st.metric("🟢 61-90 Gün", count_90)
        
        st.divider()
        
        # Listeleme
        display_df = upcoming_df[['Ad', 'Nitelik', 'Bitiş tarihi', 'Kalan Gün', 'Renk', 'Trafik Durumu']].copy()
        display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
        display_df.columns = ['Ad', 'Nitelik', 'Bitiş Tarihi', 'Kalan Gün', '', 'Durum']
        
        st.dataframe(
            display_df,
            column_config={
                'Ad': st.column_config.TextColumn('Ad'),
                'Nitelik': st.column_config.TextColumn('Nitelik'),
                'Bitiş Tarihi': st.column_config.TextColumn('Bitiş Tarihi'),
                'Kalan Gün': st.column_config.NumberColumn('Kalan Gün'),
                '': st.column_config.TextColumn(''),
                'Durum': st.column_config.TextColumn('Durum'),
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.success("✅ Yakında sona erecek belge bulunmuyor.")

def page_expired(df):
    """Süresi dolan belgeler sayfası"""
    st.markdown('<h1 class="main-header">❌ Süresi Dolan Belgeler</h1>', unsafe_allow_html=True)
    
    expired_df = df[df['Kalan Gün'] < 0].sort_values('Kalan Gün')
    
    if len(expired_df) > 0:
        st.error(f"⚠️ {len(expired_df)} belgenin süresi dolmuştur!")
        
        display_df = expired_df[['Ad', 'Nitelik', 'Bitiş tarihi', 'Kalan Gün', 'Renk']].copy()
        display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
        display_df['Gün Geçmiş'] = display_df['Kalan Gün'].abs()
        display_df.columns = ['Ad', 'Nitelik', 'Bitiş Tarihi', 'Kalan Gün', '']
        
        st.dataframe(
            display_df,
            column_config={
                'Ad': st.column_config.TextColumn('Ad'),
                'Nitelik': st.column_config.TextColumn('Nitelik'),
                'Bitiş Tarihi': st.column_config.TextColumn('Bitiş Tarihi'),
                'Kalan Gün': st.column_config.NumberColumn('Kalan Gün'),
                '': st.column_config.TextColumn(''),
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.success("✅ Süresi dolan belge bulunmuyor.")

def page_statistics(df):
    """İstatistikler sayfası"""
    st.markdown('<h1 class="main-header">📈 İstatistikler</h1>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("👤 Kişi Başına Belge Sayısı")
        person_counts = df['Ad'].value_counts().reset_index()
        person_counts.columns = ['Ad', 'Belge Sayısı']
        person_counts = person_counts.head(20)
        
        fig = px.bar(
            person_counts,
            x='Ad',
            y='Belge Sayısı',
            title='En Çok Belgeye Sahip Kişiler',
            color='Belge Sayısı',
            color_continuous_scale='Blues'
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📋 Belge Türü Dağılımı")
        cert_counts = df['Nitelik'].value_counts().reset_index()
        cert_counts.columns = ['Nitelik', 'Adet']
        cert_counts = cert_counts.head(15)
        
        fig = px.bar(
            cert_counts,
            x='Nitelik',
            y='Adet',
            title='Belge Türlerine Göre Dağılım',
            color='Adet',
            color_continuous_scale='Greens'
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    st.divider()
    
    st.subheader("📊 Özet Tablo")
    
    # Pivot tablo - Durum vs Kişi
    pivot = pd.crosstab(df['Ad'], df['Durum'])
    st.dataframe(pivot, use_container_width=True)
    
    st.divider()
    
    st.subheader("📊 Belge Durumu Özeti")
    status_summary = df['Durum'].value_counts().reset_index()
    status_summary.columns = ['Durum', 'Adet']
    status_summary['Yüzde'] = (status_summary['Adet'] / len(df) * 100).round(1)
    
    st.dataframe(status_summary, use_container_width=True, hide_index=True)

def page_calendar(df):
    """Takvim görünümü"""
    st.markdown('<h1 class="main-header">📅 Takvim Görünümü</h1>', unsafe_allow_html=True)
    
    # Ay seçimi
    months = df['Bitiş tarihi'].dt.to_period('M').unique()
    months = sorted(months)
    
    if len(months) > 0:
        selected_month = st.selectbox(
            "📅 Ay Seçin",
            options=months,
            format_func=lambda x: x.strftime('%B %Y')
        )
        
        month_data = df[df['Bitiş tarihi'].dt.to_period('M') == selected_month]
        
        if len(month_data) > 0:
            st.subheader(f"📌 {selected_month.strftime('%B %Y')} - {len(month_data)} belge sona eriyor")
            
            # Gün bazında göster
            month_data['Gün'] = month_data['Bitiş tarihi'].dt.day
            day_counts = month_data.groupby('Gün').size().reset_index(name='Adet')
            
            fig = px.bar(
                day_counts,
                x='Gün',
                y='Adet',
                title=f"{selected_month.strftime('%B %Y')} - Günlere Göre Sona Eren Belgeler",
                color='Adet',
                color_continuous_scale='Reds'
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Liste
            display_df = month_data[['Ad', 'Nitelik', 'Bitiş tarihi', 'Kalan Gün', 'Renk']].copy()
            display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
            display_df.columns = ['Ad', 'Nitelik', 'Bitiş Tarihi', 'Kalan Gün', '']
            
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Bu ayda sona eren belge bulunmuyor.")
    else:
        st.info("Veri bulunmuyor.")

def page_pivot(df):
    """Pivot tablolar sayfası"""
    st.markdown('<h1 class="main-header">📊 Pivot Tablolar</h1>', unsafe_allow_html=True)
    
    st.subheader("📌 Kişi × Durum Pivot Tablosu")
    pivot1 = pd.crosstab(df['Ad'], df['Durum'])
    st.dataframe(pivot1, use_container_width=True)
    
    st.divider()
    
    st.subheader("📌 Kişi × Belge Türü Pivot Tablosu")
    pivot2 = pd.crosstab(df['Ad'], df['Nitelik'])
    st.dataframe(pivot2, use_container_width=True)
    
    st.divider()
    
    st.subheader("📌 Durum × Belge Türü Pivot Tablosu")
    pivot3 = pd.crosstab(df['Durum'], df['Nitelik'])
    st.dataframe(pivot3, use_container_width=True)

def page_archive(df):
    """Arşiv sayfası"""
    st.markdown('<h1 class="main-header">📦 Arşiv</h1>', unsafe_allow_html=True)
    
    archived = df[df['Arşiv'] == True]
    
    if len(archived) > 0:
        display_df = archived[['Ad', 'Nitelik', 'Bitiş tarihi', 'Kalan Gün', 'Renk', 'Sorumlu Kişi', 'Notlar']].copy()
        display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
        display_df.columns = ['Ad', 'Nitelik', 'Bitiş Tarihi', 'Kalan Gün', '', 'Sorumlu Kişi', 'Notlar']
        
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("📦 Arşivlenmiş belge bulunmuyor.")

def page_reports(df):
    """Raporlar sayfası"""
    st.markdown('<h1 class="main-header">📑 Raporlar</h1>', unsafe_allow_html=True)
    
    report_type = st.selectbox(
        "📄 Rapor Türü",
        options=['Tüm Belgeler', 'Süresi Dolanlar', 'Yaklaşanlar (90 gün)', 'Kişi Bazında']
    )
    
    if report_type == 'Tüm Belgeler':
        report_df = df.copy()
    elif report_type == 'Süresi Dolanlar':
        report_df = df[df['Kalan Gün'] < 0].copy()
    elif report_type == 'Yaklaşanlar (90 gün)':
        report_df = df[(df['Kalan Gün'] >= 0) & (df['Kalan Gün'] <= 90)].copy()
    else:  # Kişi Bazında
        person = st.selectbox("👤 Personel Seçin", get_person_list(df))
        report_df = df[df['Ad'] == person].copy()
    
    if len(report_df) > 0:
        display_df = report_df[['Ad', 'Nitelik', 'Başlangıç tarihi', 'Bitiş tarihi', 'Kalan Gün', 'Durum']].copy()
        display_df['Başlangıç tarihi'] = display_df['Başlangıç tarihi'].dt.strftime('%d.%m.%Y')
        display_df['Bitiş tarihi'] = display_df['Bitiş tarihi'].dt.strftime('%d.%m.%Y')
        display_df.columns = ['Ad', 'Nitelik', 'Başlangıç', 'Bitiş', 'Kalan Gün', 'Durum']
        
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )
        
        # CSV indir
        csv = display_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 CSV Olarak İndir",
            data=csv,
            file_name=f"rapor_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.warning("Rapor için veri bulunmuyor.")

# ============================================================
# ANA UYGULAMA
# ============================================================

def main():
    # Session state başlat
    init_session_state()
    
    df = st.session_state.data
    
    # Sidebar navigasyon
    st.sidebar.image("https://img.icons8.com/fluency/96/000000/ship.png", width=80)
    st.sidebar.title("⚓ Gemiadamı Belge Takip")
    
    pages = {
        "📊 Dashboard": page_dashboard,
        "👤 Personel Kartı": page_person_card,
        "📅 Takvim Görünümü": page_calendar,
        "📋 Tüm Belgeler": page_certificates,
        "⏳ Yaklaşan Belgeler": page_upcoming,
        "❌ Süresi Dolan Belgeler": page_expired,
        "📈 İstatistikler": page_statistics,
        "📊 Pivot Tablolar": page_pivot,
        "📑 Raporlar": page_reports,
        "📦 Arşiv": page_archive,
    }
    
    # Sayfa seçimi
    selected_page = st.sidebar.radio("📌 Menü", list(pages.keys()))
    
    st.sidebar.divider()
    
    # Otomatik uyarı - sayfa yüklendiğinde
    expired_count = len(df[df['Kalan Gün'] < 0])
    critical_count = len(df[(df['Kalan Gün'] >= 0) & (df['Kalan Gün'] <= 15)])
    
    if expired_count > 0 or critical_count > 0:
        st.sidebar.warning("🚨 Acil Uyarılar!")
        if expired_count > 0:
            st.sidebar.error(f"❌ {expired_count} belgenin süresi doldu!")
        if critical_count > 0:
            st.sidebar.error(f"🔴 {critical_count} belge 15 gün içinde sona eriyor!")
    
    st.sidebar.divider()
    
    # Genel istatistikler
    st.sidebar.subheader("📊 Genel")
    st.sidebar.metric("📄 Toplam Belge", len(df))
    st.sidebar.metric("👤 Toplam Personel", len(df['Ad'].unique()))
    st.sidebar.metric("❌ Süresi Dolan", len(df[df['Kalan Gün'] < 0]))
    
    st.sidebar.divider()
    
    # Hızlı filtreler
    st.sidebar.subheader("🎛️ Hızlı Filtreler")
    
    if st.sidebar.button("🟢 Geçerli"):
        st.session_state.filter_status = 'Geçerli'
        st.rerun()
    
    if st.sidebar.button("🟡 Yaklaşıyor"):
        st.session_state.filter_status = 'Yaklaşıyor'
        st.rerun()
    
    if st.sidebar.button("🟠 Uyarı"):
        st.session_state.filter_status = 'Uyarı'
        st.rerun()
    
    if st.sidebar.button("🔴 Kritik"):
        st.session_state.filter_status = 'Kritik'
        st.rerun()
    
    if st.sidebar.button("⚫ Süresi Doldu"):
        st.session_state.filter_status = 'Süresi Doldu'
        st.rerun()
    
    if st.sidebar.button("📋 Tümü"):
        st.session_state.filter_status = 'Tümü'
        st.rerun()
    
    st.sidebar.divider()
    
    # Outlook email oluşturma
    st.sidebar.subheader("📬 E-posta")
    if st.sidebar.button("📧 Outlook ile E-posta Oluştur"):
        # Kritik belgeleri listele
        critical_docs = df[(df['Kalan Gün'] >= 0) & (df['Kalan Gün'] <= 30)]
        if len(critical_docs) > 0:
            subject = "⚠️ Gemiadamı Belge Uyarısı"
            body = "Aşağıdaki belgelerin süresi yaklaşmaktadır:\n\n"
            for _, row in critical_docs.iterrows():
                body += f"- {row['Ad']}: {row['Nitelik']} - {row['Bitiş tarihi'].strftime('%d.%m.%Y')} ({row['Kalan Gün']} gün kaldı)\n"
            
            # mailto linki
            mailto = f"mailto:?subject={subject}&body={body.replace(' ', '%20').replace('\n', '%0A')}"
            st.sidebar.markdown(f'<a href="{mailto}" target="_blank">📧 Outlook\'da Aç</a>', unsafe_allow_html=True)
        else:
            st.sidebar.info("Kritik belge bulunmuyor.")
    
    # Sayfa içeriğini göster
    pages[selected_page](df)

if __name__ == "__main__":
    main()
