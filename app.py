import datetime as dt
import io
import math
import requests
import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter

# ============================================================
# V18: AKILLI VERİ KIRPMA VE HATA YÖNETİMİ
# ============================================================
st.set_page_config(page_title="Multi-Vade Fon Analizi V18", layout="wide")
st.title("📈 Multi-Vade Fon Analizi V18")
st.caption("Otomatik Veri Kırpma (3 Ay -> 1 Ay) + İstihbarat API Motoru")

# Başlangıçta 3 aylık veri ile dene, hata olursa otomatik düşür
MAX_DAYS_DEFAULT = 90 # 3 Aylık veri
MAX_DAYS_FALLBACK = 30 # 1 Aylık veri

# TEFAS'ın yeni API endpoint'i
def get_tefas_price_data(code, days):
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    # Sadece son 'days' kadar veri çekiyoruz
    end = dt.date.today()
    start = end - dt.timedelta(days=days + 10)
    
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
                    return df
        except: continue
    return None

# Ana Akış
uploaded_file = st.file_uploader("Excel Yükle:", type=["xlsx"])
if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)
    codes = [r[0].value for r in wb["Fon_Listesi"].iter_rows(min_row=2) if r[0].value]
    
    calc_funds = []
    
    # Adım 1: Önce 3 Aylık veriyle dene
    days_to_try = MAX_DAYS_DEFAULT
    
    with st.spinner(f"Veriler {days_to_try} günlük olarak çekiliyor..."):
        for code in codes:
            df = get_tefas_price_data(code, days_to_try)
            
            # Adım 2: Hata alırsan (veri boşsa) 1 aylık ile tekrar dene
            if df is None:
                st.warning(f"{code} için 3 aylık veri çekilemedi, 1 aylık veriye düşülüyor...")
                df = get_tefas_price_data(code, MAX_DAYS_FALLBACK)
            
            if df is not None:
                # Basit metrikler
                prices = df["price"].tolist()
                ret = [(prices[i]/prices[i-1]-1)*100 for i in range(1, len(prices))]
                calc_funds.append({"code": code, "prices": prices, "daily_returns": ret})
    
    if calc_funds:
        st.success(f"{len(calc_funds)} fon analiz edildi.")
        st.write(calc_funds)
    else:
        st.error("Veri çekilemedi. Bağlantı engelleniyor olabilir.")
