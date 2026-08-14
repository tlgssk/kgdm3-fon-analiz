import datetime as dt
import math
import requests
import streamlit as st
import pandas as pd
import re

# ============================================================
# V22: TAM ENTEGRE İSTİHBARAT MOTORU (Getiri Filtreli)
# ============================================================
st.set_page_config(page_title="Multi-Vade Fon Analizi V22", layout="wide")
st.title("📈 Multi-Vade Fon Analizi V22")
st.caption("Filtreleme (AUM/Yatırımcı/Getiri) + KAP Yoğunluk Analizi")

with st.sidebar:
    st.header("⚙️ Filtreler")
    min_aum = st.number_input("Min. Portföy Büyüklüğü (TL)", value=50_000_000.0, step=10_000_000.0)
    min_inv = st.number_input("Min. Yatırımcı Sayısı", value=100, step=50)
    min_weekly_return = st.number_input("Min. Haftalık Getiri (%)", value=0.0, step=0.1)

# KAP MOTORU
def get_fon_concentration_score(code: str) -> float:
    try:
        url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={code.upper()}"
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
        kap_link = re.search(r'href="([^"]*kap.org.tr[^"]*)"', res.text)
        if kap_link:
            res_kap = requests.get(kap_link.group(1), timeout=5)
            weights = re.findall(r'(\d{1,2}[.,]\d{1,2})\s*%', res_kap.text)
            if weights:
                top_weight = max([float(w.replace(',', '.')) for w in weights])
                if top_weight > 30: return min((top_weight - 30) * 0.5, 15.0)
    except: pass
    return 0.0

# TEFAS API MOTORU
def fetch_data_full(code):
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    end = dt.date.today()
    start = end - dt.timedelta(days=30)
    
    try:
        res = requests.post(url, data={"fonkod": code.upper(), "baslangic": start.strftime("%d.%m.%Y"), "bitis": end.strftime("%d.%m.%Y")}, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        data = res.json().get("data", [])
        if not data: return None, None
        
        df = pd.DataFrame(data)
        df["price"] = df["FIYAT"].astype(float)
        
        # Meta veriler
        meta_url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={code.upper()}"
        m_res = requests.get(meta_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
        aum = float(re.search(r'LabelPortfolioValue">([\d\.]+)', m_res.text).group(1).replace('.','')) if re.search(r'LabelPortfolioValue">([\d\.]+)', m_res.text) else 0
        inv = float(re.search(r'LabelInvestorCount">([\d\.]+)', m_res.text).group(1).replace('.','')) if re.search(r'LabelInvestorCount">([\d\.]+)', m_res.text) else 0
        
        # Haftalık getiri hesaplama (Son 5 gün)
        weekly_return = ((df["price"].iloc[-1] / df["price"].iloc[-5]) - 1) * 100
        
        return {"aum": aum, "inv": inv, "weekly_return": weekly_return}, df
    except: return None, None

# ANA AKIŞ
uploaded_file = st.file_uploader("Excel Yükle (A sütunu fon kodu):", type=["xlsx"])
if uploaded_file:
    df_excel = pd.read_excel(uploaded_file, sheet_name="Fon_Listesi")
    codes = df_excel.iloc[:, 0].dropna().unique().tolist()
    
    results = []
    for code in codes:
        meta, df = fetch_data_full(code)
        
        # FİLTRELEME: AUM, Yatırımcı ve Haftalık Getiri
        if meta and meta["aum"] >= min_aum and meta["inv"] >= min_inv and meta["weekly_return"] >= min_weekly_return:
            penalty = get_fon_concentration_score(code)
            score = 50 + (meta["weekly_return"] * 2) - penalty
            results.append({
                "Fon": code, 
                "Haftalık Getiri (%)": round(meta["weekly_return"], 2),
                "AUM (Milyon TL)": round(meta["aum"]/1_000_000, 1),
                "Yatırımcı": int(meta["inv"]),
                "Yoğunlaşma Cezası": penalty,
                "Skor": int(score)
            })
            
    st.table(pd.DataFrame(results).sort_values("Skor", ascending=False))
