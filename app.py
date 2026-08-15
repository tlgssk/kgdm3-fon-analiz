import datetime as dt
import io
import json
import re
from typing import Optional

import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="KGDM-3 & KAZRİSK® Hibrit Fon Analiz Otomasyonu", page_icon="📊", layout="wide")

st.title("📊 KGDM-3 & KAZRİSK® Hibrit Fon Analiz ve Excel Otomasyonu")
st.caption("Çok Kaynaklı Veri Motoru (TEFAS + İş Yatırım + Fintables + KAP) | 16 Kriterli Hibrit Güvenlik & Momentum Skorlaması")

FUND_KINDS = ("YAT", "EMK", "BYF")
LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
HTTP_TIMEOUT = 12
APP_VERSION = "6.0.0"

# GITHUB ÜZERİNDEKİ EXCEL DOSYANIZIN RAW LİNKİ
GITHUB_EXCEL_URL = "https://github.com/tlgssk/kgdm3-fon-analiz/raw/refs/heads/main/Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"

COLOR_NAVY = "1F4E79"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"

# --- KULLANICI KONTROL PANELİ (KENAR ÇUBUĞU) ---
st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

ENABLE_FILTERS = st.sidebar.checkbox(
    "Filtreleri Etkinleştir (Kapatılırsa Tüm Fonlar Hesaplanır)",
    value=False,
    help="İşaret kaldırıldığında takip modundaki veya kriter altındaki fonlar elenmez; listedeki her fon için tam hesaplama yapılır."
)

TARGET_WEEKLY_RETURN = st.sidebar.slider(
    "Hedef Haftalık Getiri (%)",
    min_value=-5.00,
    max_value=10.00,
    value=0.00,
    step=0.10,
    help="Filtre aktifse baz alınacak minimum haftalık getiri beklentisi."
)

MIN_INVESTOR_COUNT = st.sidebar.slider(
    "Minimum Yatırımcı Sayısı Barajı",
    min_value=0,
    max_value=100000,
    value=0,
    step=500,
    help="Filtre aktifse yatırımcı sayısı bu barajın altında kalanlar filtrelenir."
)

def parse_number(value) -> Optional[float]:
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)):
        try:
            if pd.isna(value): return None
        except Exception: pass
        return float(value)
    
    text = str(value).strip()
    if not text: return None
    
    text = text.replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
    if not text: return None
    
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
        
    try: return float(text)
    except (ValueError, TypeError): return None

def normalize_fund_code(value) -> str:
    if value is None: return ""
    code = str(value).strip().upper()
    if code.endswith(".0"): code = code[:-2]
    return code

def format_percent(value) -> str:
    number = parse_number(value)
    if number is None: return "-"
    if number > 0: return f"+%{number:.2f}"
    if number < 0: return f"-%{abs(number):.2f}"
    return "%0.00"

# --- 1. KAYNAK: TEFAS VERİ ÇEKİCİ ---
@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
        crawler = Crawler(timeout=60, max_retry=3)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is None or df.empty: return pd.DataFrame()
        
        df = df.copy()
        rename_map = {"fund_code": "code", "fund_name": "title", "investor_count": "investors", "portfolio_size": "aum"}
        df.rename(columns=rename_map, inplace=True)
        
        if "date" in df.columns: df["date"] = pd.to_datetime(df["date"], errors="coerce")
        if "price" in df.columns: df["price"] = df["price"].apply(parse_number)
        if "aum" in df.columns: df["aum"] = df["aum"].apply(parse_number).fillna(0.0)
        if "investors" in df.columns: df["investors"] = df["investors"].apply(parse_number).fillna(0.0)
        
        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df = df.dropna(subset=["date", "code", "price"])
        df = df[df["price"] > 0]
        return df.sort_values(["code", "date"]).drop_duplicates(subset=["code", "date"], keep="last").reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

# --- 2. KAYNAK: İŞ YATIRIM REST API ---
def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code: return None
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        response = requests.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200: return None
        payload = response.json()
        values = payload.get("value")
        if not values: return None
        df = pd.DataFrame(values)
        if "Tarih" not in df.columns or "Fiyat" not in df.columns: return None
        df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
        df["price"] = df["Fiyat"].apply(parse_number)
        df["aum"] = 0.0
        df["investors"] = 0.0
        df = df.dropna(subset=["date", "price"])
        df = df[df["price"] > 0]
        if len(df) < 2: return None
        df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
        return df[["date", "price", "aum", "investors"]]
    except Exception:
        return None

# --- 3. KAYNAK: TEFAS DOĞRUDAN POST API ---
def fetch_tefas_direct_api(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code: return None
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    payload = {
        "fontip": "YAT",
        "fonkod": code,
        "bastarih": start.strftime("%d.%m.%Y"),
        "bittarih": end.strftime("%d.%m.%Y")
    }
    headers = {
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.tefas.gov.tr"
    }
    try:
        res = requests.post(url, data=payload, headers=headers, timeout=HTTP_TIMEOUT)
        if res.status_code == 200:
            data = res.json().get("data", [])
            if data:
                df = pd.DataFrame(data)
                df["date"] = pd.to_datetime(df["TARIH"], unit='ms')
                df["price"] = df["FIYAT"].apply(parse_number)
                df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number).fillna(0.0)
                df["investors"] = df["KISISAYISI"].apply(parse_number).fillna(0.0)
                df = df.dropna(subset=["date", "price"])
                df = df[df["price"] > 0]
                if len(df) >= 2:
                    return df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
    except Exception:
        pass
    return None

# --- 4. KAYNAK: FİNTABLES & KAP KONSANTRASYON / LİKİDİTE İSTİHBARATI ---
def fetch_fund_structural_data(fund_code: str) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {
        "top_asset_weight": 0.0,
        "is_bist30": False,
        "emergency_cash_ratio": 5.0  # Varsayılan %5 nakit
    }
    try:
        fintables_url = f"https://fintables.com/fonlar/{code.lower()}"
        headers = {"User-Agent": "Mozilla/5.0"}
        fin_res = requests.get(fintables_url, headers=headers, timeout=HTTP_TIMEOUT)
        if fin_res.status_code == 200:
            text = fin_res.text
            match_top = re.search(r'En Büyük Pay["\s:]+([0-9]+(?:\.[0-9]+)?)', text, re.IGNORECASE)
            if match_top:
                structural["top_asset_weight"] = float(match_top.group(1))
            if "BIST 30" in text or "BIST 100" in text or "Ters Repo" in text:
                structural["is_bist30"] = True
            match_cash = re.search(r'(?:Nakit|Ters Repo|PPF)["\s:]+([0-9]+(?:\.[0-9]+)?)', text, re.IGNORECASE)
            if match_cash:
                structural["emergency_cash_ratio"] = float(match_cash.group(1))
    except Exception:
        pass
    return structural

def get_fund_series(universe: pd.DataFrame, fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code: return None
    
    # 1. TEFAS Universe kontrolü
    if universe is not None and not universe.empty:
        rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
        if not rows.empty:
            rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
            if len(rows) >= 2:
                return rows.tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
                
    # 2. TEFAS Doğrudan API kontrolü
    direct_df = fetch_tefas_direct_api(code)
    if direct_df is not None and len(direct_df) >= 2:
        return direct_df

    # 3. İş Yatırım Veri Servisi kontrolü
    return fetch_isyatirim_series(code)

def compute_fund_metrics(series: Optional[pd.DataFrame], fund_code: str) -> Optional[dict]:
    if series is None or len(series) < 2: return None
    df = series.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["price"] = df["price"].apply(parse_number)
    df["aum"] = df["aum"].apply(parse_number).fillna(0.0)
    df["investors"] = df["investors"].apply(parse_number).fillna(0.0)
    
    df = df.dropna(subset=["date", "price"])
    df = df[df["price"] > 0]
    df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    
    if len(df) < 2: return None
    
    prices = df["price"].astype(float).tolist()
    dates = df["date"].dt.strftime("%d.%m").tolist()
    aums = df["aum"].astype(float).tolist()
    investors = df["investors"].astype(float).tolist()
    
    daily_returns = []
    for previous, current in zip(prices[:-1], prices[1:]):
        if previous <= 0: daily_returns.append(0.0)
        else: daily_returns.append((current / previous - 1) * 100)
            
    if not daily_returns: return None
    
    peak = prices[0]
    max_dd = 0.0
    for price in prices:
        if price > peak: peak = price
        dd = (price - peak) / peak * 100
        if dd < max_dd: max_dd = dd
            
    aum_change = ((aums[-1] - aums[0]) / aums[0] * 100) if aums[0] > 0 else 0.0
    inv_change = ((investors[-1] - investors[0]) / investors[0] * 100) if investors[0] > 0 else 0.0
    recent_weekly_ret = sum(daily_returns[-5:]) if len(daily_returns) >= 5 else sum(daily_returns)

    structural = fetch_fund_structural_data(fund_code)

    return {
        "dates": dates[1:], 
        "prices": prices,
        "daily_returns": daily_returns,
        "n_days": len(daily_returns),
        "aum": aums[-1],
        "investors": int(round(investors[-1])),
        "aum_change": aum_change,
        "inv_change": inv_change,
        "max_dd": abs(max_dd),
        "weekly_return": recent_weekly_ret,
        **structural
    }

def zscore(values: list[float]) -> list[float]:
    if not values: return []
    clean_values = [float(value) for value in values if value is not None and pd.notna(value)]
    if not clean_values: return [0.0 for _ in values]
    mean_value = sum(clean_values) / len(clean_values)
    variance = sum((value - mean_value) ** 2 for value in clean_values) / len(clean_values)
    std = variance ** 0.5
    if std < 1e-12: return [0.0 for _ in values]
    raw_z = [(float(value) - mean_value) / std for value in values]
    return [max(-2.5, min(2.5, z)) for z in raw_z]

def style_excel_sheet(ws):
    thin_gray = Side(style="thin", color="D9E1F2")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin_gray)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False

def auto_fit_columns(ws, min_width: int = 10, max_width: int = 45):
    for column_cells in ws.columns:
        column_index = column_cells[0].column
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        width = max(min_width, min(max_length + 3, max_width))
        ws.column_dimensions[get_column_letter(column_index)].width = width

def create_excel_output(wb, ws_list, calculated_funds, n_days) -> io.BytesIO:
    if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")
    
    day_labels = calculated_funds[0]["dates"]
    headers = [
        "Fon Kodu", "Valör", "Hibrit Skor", "Son 5 Gün Skorlar", "Model Kararı", 
        "Ort. Getiri (%)", "Volatilite (%)", "Sharpe", "Kümülatif Getiri (%)", 
        "En Büyük Varlık (%)", "KAZRİSK® Likidite Notu", "AUM Değişim (%)", "Yatırımcı Değişim (%)", 
        "MaxDD (%)", "Fon Büyüklüğü (AUM ₺)", "Yatırımcı Sayısı"
    ]
    for day in day_labels: headers.append(f"{day} Skor")
    for day in day_labels: headers.append(f"{day} % Getiri")
    ws_scores.append(headers)
    
    header_fill = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color=COLOR_WHITE)
    for cell in ws_scores[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_scores.row_dimensions[1].height = 34
    
    for item in calculated_funds:
        risk_label = "⚠️ Sığ / Yüksek Risk" if item["top_asset_weight"] > 30.0 else "🛡️ Likit / BİST30"
        if item["top_asset_weight"] == 0.0: risk_label = "✅ Dengeli / Normal"
        
        row_data = [
            item["code"], item["valor"], item["kgdm_skor"], item["last_5_scores_str"], item["karar"], 
            round(item["mean_return"], 4), round(item["volatility"], 4), round(item["sharpe_like"], 4), 
            round(item["cumulative_return"], 4), round(item["top_asset_weight"], 2), risk_label,
            round(item["aum_change"], 2), round(item["inv_change"], 2), round(item["max_dd"], 2),
            round(item["aum"], 2) if item["aum"] else None, int(item["investors"]) if item["investors"] else None
        ]
        row_data.extend(item["running_scores"])
        row_data.extend([format_percent(value) for value in item["daily_returns"]])
        ws_scores.append(row_data)
        
    green_font = Font(bold=True, color=COLOR_GREEN)
    red_font = Font(bold=True, color=COLOR_RED)
    yellow_font = Font(bold=True, color=COLOR_YELLOW)
    
    for row_number in range(2, ws_scores.max_row + 1):
        ws_scores.cell(row=row_number, column=15).number_format = '#,##0.00 "₺"' 
        ws_scores.cell(row=row_number, column=16).number_format = '#,##0' 
        
        decision_cell = ws_scores.cell(row=row_number, column=5)
        decision_text = str(decision_cell.value)
        if "GÜÇLÜ AL" in decision_text or "ASIL LİSTE" in decision_text: decision_cell.font = green_font
        elif "NÖTR" in decision_text or "DÜZELTME" in decision_text: decision_cell.font = yellow_font
        elif "ACİL SAT" in decision_text or "KESİN SAT" in decision_text: decision_cell.font = red_font
            
    score_range = f"C2:C{ws_scores.max_row}"
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="greaterThanOrEqual", formula=["70"], fill=PatternFill(start_color="E2F0D9", fill_type="solid")))
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="between", formula=["45", "69"], fill=PatternFill(start_color="EAF4E3", fill_type="solid")))
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="lessThan", formula=["45"], fill=PatternFill(start_color="FCE4D6", fill_type="solid")))
    
    style_excel_sheet(ws_scores)
    auto_fit_columns(ws_scores)
    style_excel_sheet(ws_list)
    auto_fit_columns(ws_list)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- ANA UYGULAMA ARAYÜZÜ ---
st.subheader("📂 Portföy Excel Listesi Seçimi")
col_upload, col_github = st.columns(2)

wb = None
source_mode = None

with col_upload:
    uploaded_file = st.file_uploader("Bilgisayardan Excel Yükle (fonlar.xlsx):", type=["xlsx"])
    if uploaded_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_file)
            source_mode = "upload"
        except Exception as exc: st.error(f"Hata: {exc}")

with col_github:
    st.write("Veya GitHub'daki listeyi kullanın:")
    if st.button("🚀 GitHub'dan Çek ve Analiz Et", use_container_width=True):
        try:
            gh_response = requests.get(GITHUB_EXCEL_URL, timeout=HTTP_TIMEOUT)
            gh_response.raise_for_status()
            wb = openpyxl.load_workbook(io.BytesIO(gh_response.content))
            source_mode = "github"
            st.success("✅ Excel dosyası başarıyla indirildi!")
        except Exception as exc: st.error(f"Bağlantı hatası: {exc}")

if wb is None: st.stop()
ws_list = wb["Fon_Listesi"] if "Fon_Listesi" in wb.sheetnames else wb.active
requested_codes = []
excel_valor_dict = {}

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if not row or row[0].value is None: continue
    code = normalize_fund_code(row[0].value)
    if not code: continue
    requested_codes.append(code)
    try: excel_valor_dict[code] = int(round(parse_number(row[3].value)))
    except: excel_valor_dict[code] = 0

requested_codes = list(dict.fromkeys(requested_codes))

today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

with st.spinner("🔄 Çoklu Veri Ağlarından Fon Verileri Alınıyor..."):
    universe = fetch_tefas_universe(start_date, today)

calculated_funds = []
progress = st.progress(0, text="Bütün fonlar analiz ediliyor (Takip modu dahil)...")
total_funds = len(requested_codes)

for index, code in enumerate(requested_codes, start=1):
    series = get_fund_series(universe, code)
    metrics = compute_fund_metrics(series, code)
    if metrics is not None:
        if not ENABLE_FILTERS or (metrics["investors"] >= MIN_INVESTOR_COUNT and metrics["weekly_return"] >= TARGET_WEEKLY_RETURN):
            calculated_funds.append({"code": code, "valor": excel_valor_dict.get(code, 0), **metrics})
    progress.progress(index / total_funds, text=f"{code} analiz edildi...")

progress.empty()

if not calculated_funds:
    st.error("Hesaplanabilecek geçerli fon verisi bulunamadı. Lütfen girdi listesini kontrol edin.")
    st.stop()

n_days = min(item["n_days"] for item in calculated_funds)
for item in calculated_funds:
    item["dates"] = item["dates"][-n_days:]
    item["daily_returns"] = item["daily_returns"][-n_days:]
    item["prices"] = item["prices"][-(n_days+1):]
    item["running_scores"] = []

# --- HİBRİT SKORLAMA MOTORU (KGDM-3 + KAZRİSK®) ---
for d in range(1, n_days + 1):
    day_means, day_sharpes, day_cums, day_aum_chgs, day_inv_chgs, day_maxdds = [], [], [], [], [], []
    for item in calculated_funds:
        slice_ret = item["daily_returns"][:d]
        m_ret = sum(slice_ret) / len(slice_ret)
        vol = (sum((r - m_ret) ** 2 for r in slice_ret) / len(slice_ret)) ** 0.5
        shp = (m_ret / vol) if vol > 1e-12 else 0.0
        cum = (item["prices"][d] / item["prices"][0] - 1) * 100
        
        day_means.append(m_ret); day_sharpes.append(shp); day_cums.append(cum)
        day_aum_chgs.append(item["aum_change"]); day_inv_chgs.append(item["inv_change"]); day_maxdds.append(item["max_dd"])
        
        if d == n_days:
            item["mean_return"] = m_ret; item["volatility"] = vol; item["sharpe_like"] = shp; item["cumulative_return"] = cum
            
    z_m = zscore(day_means)
    z_s = zscore(day_sharpes)
    z_c = zscore(day_cums)
    z_a = zscore(day_aum_chgs)
    z_i = zscore(day_inv_chgs)
    z_d = zscore(day_maxdds)
    
    for i, item in enumerate(calculated_funds):
        val = item["valor"] if item["valor"] is not None else 0
        v_pen = (val * 0.5)
        
        # 1. KGDM MOMENTUM MOTORU (50 Puan Tabanı)
        kgdm_raw = 50 + 10*z_m[i] + 12*z_s[i] + 12*z_c[i] - 10*z_d[i]
        
        # 2. KAZRİSK GÜVENLİK & LİKİDİTE KALKANI (50 Puan Tabanı)
        kazrisk_bonus = 0.0
        if item.get("is_bist30", False): kazrisk_bonus += 7.0
        if item.get("emergency_cash_ratio", 0.0) >= 15.0: kazrisk_bonus += 5.0
        if item["inv_change"] > 0: kazrisk_bonus += 3.0
        
        kazrisk_penalty = 0.0
        if item["top_asset_weight"] > 30.0:
            kazrisk_penalty += (item["top_asset_weight"] - 30.0) * 1.2
        if item["top_asset_weight"] > 15.0 and not item.get("is_bist30", False):
            kazrisk_penalty += 7.0  # Sığ hisse derinlik cezası
            
        kazrisk_raw = 50 + 10*z_a[i] + 10*z_i[i] + kazrisk_bonus - kazrisk_penalty - v_pen
        
        # HİBRİT BİRLEŞİM (%50 KGDM + %50 KAZRİSK)
        hybrid_score = int(round(max(5.0, min(95.0, (kgdm_raw * 0.5) + (kazrisk_raw * 0.5)))))
        item["running_scores"].append(hybrid_score)

for item in calculated_funds:
    scores = item["running_scores"]
    last_5 = scores[-5:] if len(scores) >= 5 else scores
    item["kgdm_skor"] = int(round(sum(last_5) / len(last_5)))
    item["last_5_scores_str"] = " ➔ ".join([str(s) for s in last_5])
    
    # 4 Kademeli Karar Matrisi
    if item["kgdm_skor"] >= 75: item["karar"] = "GÜÇLÜ AL"
    elif item["kgdm_skor"] >= 50: item["karar"] = "ASIL LİSTE"
    elif item["kgdm_skor"] >= 35: item["karar"] = "DÜZELTME / İZLE"
    else: item["karar"] = "ACİL SAT"

calculated_funds.sort(key=lambda x: (-x["kgdm_skor"], -x["cumulative_return"]))

display_rows = []
for item in calculated_funds:
    risk_status = "⚠️ Sığ Hisse / Yüksek Risk" if item["top_asset_weight"] > 30.0 else "🛡️ Likit / Güvenli"
    if item["top_asset_weight"] == 0.0: risk_status = "✅ Normal"
    
    display_rows.append({
        "Fon Kodu": item["code"], 
        "Hibrit Skor (100)": item["kgdm_skor"], 
        "Model Kararı": item["karar"],
        "Ort. Günlük %": round(item["mean_return"], 3), 
        "Sharpe Oranı": round(item["sharpe_like"], 3), 
        "Kümülatif Getiri %": round(item["cumulative_return"], 3),
        "En Büyük Varlık %": round(item["top_asset_weight"], 2),
        "KAZRİSK Güvenlik": risk_status,
        "Fon Büyüklüğü ₺": round(item["aum"], 0), 
        "Yatırımcı": item["investors"]
    })

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value)
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "Likit" in text or "Normal" in text: 
        return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text: 
        return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text or "Sığ Hisse" in text: 
        return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except AttributeError: styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 Analiz Sonuçları (KAZRİSK & KGDM Hibrit Motoru)")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

output = create_excel_output(wb=wb, ws_list=ws_list, calculated_funds=calculated_funds, n_days=n_days)
st.success("✅ Tüm fonlar (takip modundakiler dahil) 4 farklı veri kaynağından çekilerek KAZRİSK derinlik cezalarıyla hesaplandı.")
st.download_button(label="📥 Güncellenmiş Hibrit Excel'i İndir", data=output, file_name="fonlar_guncel.xlsx")
