import datetime as dt
import math
import requests
import streamlit as st
import pandas as pd

# ============================================================
# V19: KADEMELİ VERİ İNDİRME MOTORU (90 -> 30 -> 10 Gün)
# ============================================================
st.set_page_config(page_title="Multi-Vade Fon Analizi V19", layout="wide")
st.title("📈 Multi-Vade Fon Analizi V19")
st.caption("Kademeli Veri İndirme (90-30-10 Gün) ile Garantili Çalışma")

def get_tefas_price_data(code, days):
    """TEFAS resmi JSON API'sinden veri çeker."""
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    end = dt.date.today()
    start = end - dt.timedelta(days=days + 15) # biraz pay bırakıyoruz
    
    headers = {"User-Agent": "Mozilla/5.0", "X-Requested-With": "XMLHttpRequest"}
    
    for fontip in ["", "YAT", "EMK", "BYF"]:
        data = {"fonkod": code.upper(), "baslangic": start.strftime("%d.%m.%Y"), "bitis": end.strftime("%d.%m.%Y"), "fontip": fontip}
        try:
            res = requests.post(url, data=data, headers=headers, timeout=10)
            if res.status_code == 200:
                data_list = res.json().get("data", [])
                if data_list:
                    rows = [{"date": pd.Timestamp(dt.datetime.fromtimestamp(x["TARIH"]/1000).date()), "price": float(x["FIYAT"])} for x in data_list]
                    df = pd.DataFrame(rows).sort_values("date").drop_duplicates("date").tail(days + 1)
                    if len(df) >= 2: return df
        except: continue
    return None

# Arayüz
uploaded_file = st.file_uploader("Excel Yükle:", type=["xlsx"])
if uploaded_file:
    # Basit bir okuma (Fon kodlarını A sütunundan al)
    df_excel = pd.read_excel(uploaded_file, sheet_name="Fon_Listesi")
    codes = df_excel.iloc[:, 0].dropna().unique().tolist()
    
    calc_funds = []
    progress_text = st.empty()
    bar = st.progress(0)
    
    # Kademeli deneme listesi
    fallback_days = [90, 30, 10]
    
    for i, code in enumerate(codes):
        df = None
        for days in fallback_days:
            progress_text.text(f"Analiz ediliyor: {code} ({days} günlük veri deneniyor...)")
            df = get_tefas_price_data(code, days)
            if df is not None:
                st.write(f"✅ {code} için {days} günlük veri başarıyla alındı.")
                break # Veri bulunduysa bir sonraki fona geç
        
        if df is not None:
            prices = df["price"].tolist()
            ret = [(prices[i]/prices[i-1]-1)*100 for i in range(1, len(prices))]
            calc_funds.append({"code": code, "prices": prices, "daily_returns": ret})
        else:
            st.error(f"❌ {code} için 10 günlük veri de çekilemedi (Erişim engeli).")
        
        bar.progress((i + 1) / len(codes))
    
    if calc_funds:
        st.success(f"Başarıyla {len(calc_funds)} fonun verisi yüklendi.")
        st.dataframe(pd.DataFrame(calc_funds))
    else:
        st.error("Hiçbir veriye ulaşılamadı. Kurumlar IP'nizi engelliyor olabilir.")
