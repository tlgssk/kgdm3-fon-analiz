import datetime as dt
import math
import requests
import streamlit as st
import pandas as pd
import re
from typing import Optional, Tuple

# ============================================================
# V20: KAP HİSSE YOĞUNLUK ANALİZ MOTORU
# ============================================================
st.set_page_config(page_title="Multi-Vade Fon Analizi V20", layout="wide")
st.title("📈 Multi-Vade Fon Analizi V20")
st.caption("KAP Portföy Dağılım Analizi + Otomatik Yoğunlaşma Riski Tespitli")

# ============================================================
# KAP MOTORU: Hisse Yoğunluk Analizi
# ============================================================
def get_fon_concentration_score(code: str) -> float:
    """
    KAP üzerinden fonun en büyük hissesinin ağırlığını çeker.
    Eğer %30'dan fazlaysa ceza puanı (penalty) döndürür.
    """
    try:
        # TEFAS Fon Analiz sayfası üzerinden KAP linkini bul
        url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={code.upper()}"
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
        # Portföy dağılım raporunu içeren bir iframe veya linki yakala
        kap_link = re.search(r'href="([^"]*kap.org.tr[^"]*)"', res.text)
        
        if kap_link:
            kap_url = kap_link.group(1)
            res_kap = requests.get(kap_url, timeout=5)
            # En büyük varlığın ağırlığını yüzdesel olarak yakala
            # Regex: %30.5 gibi veya 0.30 gibi yapıları arar
            weights = re.findall(r'(\d{1,2}[.,]\d{1,2})\s*%', res_kap.text)
            if weights:
                top_weight = max([float(w.replace(',', '.')) for w in weights])
                # %30'u aşan her 1 puanlık yoğunluk için 0.5 ceza puanı
                if top_weight > 30:
                    penalty = (top_weight - 30) * 0.5
                    return min(penalty, 15.0) # Ceza max 15 puan olsun
    except:
        pass
    return 0.0

# ============================================================
# VERİ MOTORU (DİNAMİK)
# ============================================================
@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_data(code, days):
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {"User-Agent": "Mozilla/5.0", "X-Requested-With": "XMLHttpRequest"}
    end = dt.date.today()
    start = end - dt.timedelta(days=days + 15)
    for fontip in ["", "YAT", "EMK", "BYF"]:
        data = {"fonkod": code.upper(), "baslangic": start.strftime("%d.%m.%Y"), "bitis": end.strftime("%d.%m.%Y"), "fontip": fontip}
        try:
            res = requests.post(url, data=data, headers=headers, timeout=5)
            if res.status_code == 200:
                data_list = res.json().get("data", [])
                if data_list:
                    rows = [{"date": pd.Timestamp(dt.datetime.fromtimestamp(x["TARIH"]/1000).date()), "price": float(x["FIYAT"])} for x in data_list]
                    return pd.DataFrame(rows).sort_values("date").drop_duplicates("date").tail(days + 1)
        except: continue
    return None

# ============================================================
# ANA ARAYÜZ
# ============================================================
uploaded_file = st.file_uploader("Excel Yükle (A sütunu fon kodu):", type=["xlsx"])

if uploaded_file:
    df_excel = pd.read_excel(uploaded_file, sheet_name="Fon_Listesi")
    codes = df_excel.iloc[:, 0].dropna().unique().tolist()
    
    results = []
    bar = st.progress(0)
    
    for i, code in enumerate(codes):
        st.write(f"Analiz ediliyor: {code}")
        
        # 1. Fiyat Verisi (10-90 gün arası kademeli)
        df = None
        for d in [90, 30, 10]:
            df = fetch_data(code, d)
            if df is not None: break
            
        if df is not None:
            # 2. KAP Yoğunlaşma Riski (V20 Yeni)
            concentration_penalty = get_fon_concentration_score(code)
            
            # 3. Skorlama
            prices = df["price"].tolist()
            ret = [(prices[i]/prices[i-1]-1)*100 for i in range(1, len(prices))]
            perf_score = 50.0 + (sum(ret)/len(ret)) * 5 # Basit bir momentum skoru
            
            final_score = max(0, min(100, perf_score - concentration_penalty))
            
            results.append({
                "Fon": code,
                "Skor": int(final_score),
                "Yoğunluk Cezası": concentration_penalty,
                "Durum": "GÜÇLÜ" if final_score > 60 else "RİSKLİ"
            })
            
        bar.progress((i + 1) / len(codes))
    
    st.table(pd.DataFrame(results))
