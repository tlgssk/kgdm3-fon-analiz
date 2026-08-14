import datetime as dt
import io
import re
from typing import Optional

import openpyxl
import pandas as pd
requests = pd.read_csv  # çakışma önleyici
import requests
import streamlit as st

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="KGDM-3 Fon Analiz Otomasyonu", page_icon="📊", layout="wide")

st.title("📊 KGDM-3 Fon Analiz ve Excel Otomasyonu")
st.caption("Varlık Dağılımı & Odak Hisse Yoğunlaşma Analizi | GitHub Entegrasyonu, 5.000 Yatırımcı Alt Sınırı ve Dengelenmiş Z-Skor.")

FUND_KINDS = ("YAT", "EMK", "BYF")
LOOKBACK_CALENDAR_DAYS = 35
TARGET_TRADING_DAYS = 10
MIN_INVESTOR_COUNT = 5000
HTTP_TIMEOUT = 8
APP_VERSION = "5.0.1"

# GITHUB ÜZERİNDEKİ EXCEL DOSYANIZIN RAW LİNKİNİ BURAYA YAPIŞTIRIN:
GITHUB_EXCEL_URL = "https://github.com/tlgssk/kgdm3-fon-analiz/raw/refs/heads/main/Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"

COLOR_NAVY = "1F4E79"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"
COLOR_LIGHT_GREEN = "E2F0D9"
COLOR_LIGHT_RED = "FCE4D6"
COLOR_LIGHT_YELLOW = "FFF2CC"

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

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try: from pytefas import Crawler
    except ImportError: return pd.DataFrame()

    try:
        crawler = Crawler(timeout=60, max_retry=5)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is None or df.empty: return pd.DataFrame()
        
        df = df.copy()
        rename_map = {"fund_code": "code", "fund_name": "title", "investor_count": "investors", "portfolio_size": "aum"}
        df.rename(columns=rename_map, inplace=True)
        
        required_columns = ["date", "code", "price"]
        for column in required_columns:
            if column not in df.columns: return pd.DataFrame()
            
        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["price"] = df["price"].apply(parse_number)
        
        if "aum" in df.columns: df["aum"] = df["aum"].apply(parse_number)
        else: df["aum"] = 0.0
        
        if "investors" in df.columns: df["investors"] = df["investors"].apply(parse_number)
        else: df["investors"] = 0.0
            
        df = df.dropna(subset=["date", "code", "price"])
        df = df[df["price"] > 0]
        df = df.sort_values(["code", "date"]).drop_duplicates(subset=["code", "date"], keep="last").reset_index(drop=True)
        return df
    except Exception:
        return pd.DataFrame()

def get_fund_series(universe: pd.DataFrame, fund_code: str) -> Optional[pd.DataFrame]:
    if universe is None or universe.empty: return None
    code = normalize_fund_code(fund_code)
    if not code: return None
    
    rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
    if rows.empty: return None
    
    rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
    if len(rows) < 2: return None
    if len(rows) > TARGET_TRADING_DAYS + 1: rows = rows.tail(TARGET_TRADING_DAYS + 1)
    return rows.reset_index(drop=True)

def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code: return None
    
    end = dt.datetime.now()
    start = end - dt.timedelta(days=45)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128 Safari/537.36",
        "Accept": "application/json,text/plain,*/*"
    }
    try:
        response = requests.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
        payload = response.json()
        values = payload.get("value")
        
        if not isinstance(values, list) or not values: return None
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
    except (requests.RequestException, ValueError, TypeError):
        return None

def fetch_fintables_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code: return None
    
    url = f"https://fintables.com/fonlar/{code.lower()}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128 Safari/537.36"}
    try:
        response = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200: return None
        
        html = response.text
        pattern = re.compile(r'"date"\s*:\s*"(\d{4}-\d{2}-\d{2})"[^{}]{0,500}?"price"\s*:\s*([0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
        matches = pattern.findall(html)
        if not matches: return None
        
        rows = []
        for date_text, price_text in matches:
            date_value = pd.to_datetime(date_text, errors="coerce")
            price_value = parse_number(price_text)
            if pd.notna(date_value) and price_value is not None and price_value > 0:
                rows.append({"date": date_value, "price": price_value, "aum": 0.0, "investors": 0.0})
                
        if not rows: return None
        df = pd.DataFrame(rows)
        df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last")
        if len(df) < 2: return None
        return df.tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
    except (requests.RequestException, ValueError, TypeError, re.error):
        return None

def fetch_fintables_asset_allocation(fund_code: str) -> dict:
    code = normalize_fund_code(fund_code)
    url = f"https://fintables.com/fonlar/{code.lower()}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128 Safari/537.36"}
    
    result = {"stock_ratio": 50.0}
    try:
        response = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200: return result
        html = response.text
        match_stock = re.search(r'Hisse Senedi["\s:]+([0-9]+(?:\.[0-9]+)?)', html, re.IGNORECASE)
        if match_stock:
            result["stock_ratio"] = float(match_stock.group(1))
        return result
    except Exception:
        return result

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

    allocation = fetch_fintables_asset_allocation(fund_code)

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
        "stock_ratio": allocation["stock_ratio"]
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
        "Fon Kodu", "Valör (Excel)", "KGDM-3 Skor (Ort5)", "Son 5 Gün Skorlar", "5 Günlük Trend", 
        "Model Kararı", "Ort. Getiri (%)", "Volatilite (%)", "Sharpe", "Kümülatif Getiri (%)", 
        "Hisse Oranı (%)", "AUM Değişim (%)", "Yatırımcı Değişim (%)", "MaxDD (%)", "Fon Büyüklüğü (AUM ₺)", "Yatırımcı Sayısı", "Fiyat Kaynağı"
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
        row_data = [
            item["code"],
            item["valor"] if item["valor"] is not None else None,
            item["kgdm_skor"], 
            item["last_5_scores_str"],
            item["trend_status"],
            item["karar"], 
            round(item["mean_return"], 4), 
            round(item["volatility"], 4),
            round(item["sharpe_like"], 4), 
            round(item["cumulative_return"], 4),
            round(item["stock_ratio"], 2),
            round(item["aum_change"], 2),
            round(item["inv_change"], 2),
            round(item["max_dd"], 2),
            round(item["aum"], 2) if item["aum"] else None,
            int(item["investors"]) if item["investors"] else None, 
            item["data_source"]
        ]
        row_data.extend(item["running_scores"])
        row_data.extend([format_percent(value) for value in item["daily_returns"]])
        ws_scores.append(row_data)
        
    COL_VALOR = 2; COL_SCORE = 3; COL_DECISION = 6
    SCORE_START = 18
    RETURN_START = SCORE_START + n_days
    
    green_font = Font(bold=True, color=COLOR_GREEN)
    red_font = Font(bold=True, color=COLOR_RED)
    yellow_font = Font(bold=True, color=COLOR_YELLOW)
    
    for row_number in range(2, ws_scores.max_row + 1):
        ws_scores.cell(row=row_number, column=15).number_format = '#,##0.00 "₺"' 
        ws_scores.cell(row=row_number, column=16).number_format = '#,##0' 
        ws_scores.cell(row=row_number, column=COL_VALOR).number_format = '0'
        ws_scores.cell(row=row_number, column=COL_SCORE).number_format = '0'
        
        decision_cell = ws_scores.cell(row=row_number, column=COL_DECISION)
        decision_text = str(decision_cell.value)
        if "GÜÇLÜ AL" in decision_text or "ASIL LİSTE" in decision_text: decision_cell.font = green_font
        elif "NÖTR" in decision_text: decision_cell.font = yellow_font
        elif "ACİL SAT" in decision_text: decision_cell.font = red_font
        
        for col in range(RETURN_START, RETURN_START + n_days):
            cell = ws_scores.cell(row=row_number, column=col)
            value = str(cell.value)
            if value.startswith("+"): cell.font = green_font
            elif value.startswith("-"): cell.font = red_font
            
    score_range = f"C2:C{ws_scores.max_row}"
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="greaterThanOrEqual", formula=["60"], fill=PatternFill(start_color=COLOR_LIGHT_GREEN, fill_type="solid")))
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="between", formula=["40", "59"], fill=PatternFill(start_color="EAF4E3", fill_type="solid")))
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="between", formula=["25", "39"], fill=PatternFill(start_color=COLOR_LIGHT_YELLOW, fill_type="solid")))
    ws_scores.conditional_formatting.add(score_range, CellIsRule(operator="lessThan", formula=["25"], fill=PatternFill(start_color=COLOR_LIGHT_RED, fill_type="solid")))
    
    style_excel_sheet(ws_scores)
    auto_fit_columns(ws_scores)
    for col in range(SCORE_START, ws_scores.max_column + 1): ws_scores.column_dimensions[get_column_letter(col)].width = 13
    
    style_excel_sheet(ws_list)
    auto_fit_columns(ws_list)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------------------------------------------------------------------------
# ANA UYGULAMA
# ---------------------------------------------------------------------------

st.subheader("📂 Veri Kaynağı Seçimi")
col_upload, col_github = st.columns(2)

wb = None
source_mode = None

with col_upload:
    uploaded_file = st.file_uploader("Bilgisayardan Excel Yükle (fonlar.xlsx):", type=["xlsx"])
    if uploaded_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_file)
            source_mode = "upload"
        except Exception as exc:
            st.error(f"Excel dosyası okunamadı: {exc}")

with col_github:
    st.write("Veya GitHub'daki listeyi kullanın:")
    if st.button("🚀 GitHub'dan Otomatik Çek ve Analiz Et", use_container_width=True):
        try:
            gh_response = requests.get(GITHUB_EXCEL_URL, timeout=HTTP_TIMEOUT)
            gh_response.raise_for_status()
            wb = openpyxl.load_workbook(io.BytesIO(gh_response.content))
            source_mode = "github"
            st.success("✅ GitHub'daki Excel dosyası başarıyla indirildi!")
        except Exception as exc:
            st.error(f"GitHub'dan dosya çekilemedi: {exc}")

if wb is None:
    st.info("Başlamak için yukarıdan Excel dosyası yükleyin ya da GitHub butonuna basın.")
    st.stop()

if "Fon_Listesi" not in wb.sheetnames: 
    st.error("Yüklenen veya indirilen dosyada 'Fon_Listesi' sayfası bulunamadı!")
    st.stop()

ws_list = wb["Fon_Listesi"]
requested_codes = []
excel_valor_dict = {}

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if not row: continue
    code_cell = row[0]
    valor_cell = row[3] if len(row) > 3 else None
    code = normalize_fund_code(code_cell.value)
    if not code: continue
    requested_codes.append(code)
    
    valor = None
    if valor_cell is not None and valor_cell.value is not None and str(valor_cell.value).strip():
        parsed_valor = parse_number(valor_cell.value)
        if parsed_valor is not None: valor = int(round(parsed_valor))
    excel_valor_dict[code] = valor

requested_codes = list(dict.fromkeys(requested_codes))
if not requested_codes: st.warning("Fon_Listesi sayfasında fon kodu bulunamadı."); st.stop()

st.subheader("📋 Analiz Özeti")
summary_col1, summary_col2, summary_col3 = st.columns(3)
with summary_col1: st.metric("Analiz Edilecek Fon", len(requested_codes))
with summary_col2: st.metric("Hedef İşlem Günü", TARGET_TRADING_DAYS)
with summary_col3: st.metric("Yatırımcı Alt Sınırı", f"{MIN_INVESTOR_COUNT:,} Kişi".replace(",", "."))

today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

with st.spinner("🔄 TEFAS verileri ve Varlık Dağılımları taranıyor..."):
    universe = fetch_tefas_universe(start_date, today)

if universe.empty:
    st.warning("⚠️ Ana veri kaynağı olan TEFAS'a (pytefas) erişilemedi. Yedek kaynaklar (İş Yatırım, Fintables) ile analiz devam ediyor...")

available_codes = set(universe["code"].astype(str).str.upper().unique()) if not universe.empty else set()
found_in_tefas = [code for code in requested_codes if code in available_codes]
missing_from_tefas = [code for code in requested_codes if code not in available_codes]

calculated_funds = []
source_errors = []
investor_filtered_out = []
progress = st.progress(0, text="Fonlar analiz ediliyor...")
total_funds = len(requested_codes)

for index, code in enumerate(requested_codes, start=1):
    series = None
    data_source = "Bulunamadı"
    
    if not universe.empty:
        series = get_fund_series(universe, code)
        if series is not None: data_source = "TEFAS"
        
    if series is None:
        series = fetch_isyatirim_series(code)
        if series is not None: data_source = "İş Yatırım"
        
    if series is None:
        series = fetch_fintables_series(code)
        if series is not None: data_source = "Fintables"
        
    metrics = compute_fund_metrics(series, code)
    if metrics is None:
        source_errors.append({"code": code, "source": data_source, "reason": "Yeterli fiyat verisi bulunamadı."})
        progress.progress(index / total_funds, text=f"{code} işlenemedi...")
        continue
        
    if metrics["investors"] < MIN_INVESTOR_COUNT and data_source == "TEFAS":
        investor_filtered_out.append({"code": code, "investors": metrics["investors"], "reason": f"Yatırımcı sayısı barajın altında (< {MIN_INVESTOR_COUNT:,})".replace(",", ".")})
        progress.progress(index / total_funds, text=f"{code} yatırımcı sayısı nedeniyle elendi...")
        continue

    final_valor = excel_valor_dict.get(code)
    calculated_funds.append({"code": code, "data_source": data_source, "valor": final_valor, **metrics})
    progress.progress(index / total_funds, text=f"{code} işleniyor ({index}/{total_funds})")

progress.empty()

if not calculated_funds:
    st.error("10.000 yatırımcı kriterini sağlayan hiçbir fon bulunamadı.")
    if investor_filtered_out:
        st.subheader("🚫 Yatırımcı Sayısı Nedeniyle Elenen Fonlar")
        st.dataframe(pd.DataFrame(investor_filtered_out), use_container_width=True, hide_index=True)
    if source_errors: st.dataframe(pd.DataFrame(source_errors), use_container_width=True)
    st.stop()

n_days = min(item["n_days"] for item in calculated_funds)
for item in calculated_funds:
    item["dates"] = item["dates"][-n_days:]
    item["daily_returns"] = item["daily_returns"][-n_days:]
    item["prices"] = item["prices"][-(n_days+1):]
    item["running_scores"] = []

for d in range(1, n_days + 1):
    day_means, day_sharpes, day_cums, day_aum_chgs, day_inv_chgs, day_maxdds, day_stocks = [], [], [], [], [], [], []
    for item in calculated_funds:
        slice_ret = item["daily_returns"][:d]
        m_ret = sum(slice_ret) / len(slice_ret)
        var = sum((r - m_ret) ** 2 for r in slice_ret) / len(slice_ret)
        vol = var ** 0.5
        shp = (m_ret / vol) if vol > 1e-12 else 0.0
        cum = (item["prices"][d] / item["prices"][0] - 1) * 100
        
        day_means.append(m_ret)
        day_sharpes.append(shp)
        day_cums.append(cum)
        day_aum_chgs.append(item["aum_change"])
        day_inv_chgs.append(item["inv_change"])
        day_maxdds.append(item["max_dd"])
        day_stocks.append(item["stock_ratio"])
        
        if d == n_days:
            item["mean_return"] = m_ret
            item["volatility"] = vol
            item["sharpe_like"] = shp
            item["cumulative_return"] = cum
            
    z_m = zscore(day_means) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    z_s = zscore(day_sharpes) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    z_c = zscore(day_cums) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    z_a = zscore(day_aum_chgs) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    z_i = zscore(day_inv_chgs) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    z_d = zscore(day_maxdds) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    z_st = zscore(day_stocks) if len(calculated_funds) > 1 else [0.0 for _ in calculated_funds]
    
    for i, item in enumerate(calculated_funds):
        val = item["valor"]
        v_pen = (val * 0.5) if val is not None else 0.0
        
        raw_score = 50 + 10 * z_m[i] + 15 * z_s[i] + 10 * z_c[i] + 10 * z_a[i] + 10 * z_i[i] + 10 * z_st[i] - 15 * z_d[i] - v_pen
        score = int(round(max(5.0, min(95.0, raw_score))))
        item["running_scores"].append(score)

for item in calculated_funds:
    scores = item["running_scores"]
    last_5 = scores[-5:] if len(scores) >= 5 else scores
    
    avg_score_5d = sum(last_5) / len(last_5)
    item["kgdm_skor"] = int(round(avg_score_5d))
    item["last_5_scores_str"] = " ➔ ".join([str(s) for s in last_5])
    
    recent_returns = item["daily_returns"][-5:] if len(item["daily_returns"]) >= 5 else item["daily_returns"]
    recent_cum_return = sum(recent_returns)
    
    if len(last_5) >= 2:
        if last_5[-1] > last_5[0] and recent_cum_return > 0:
            item["trend_status"] = "📈 Yükselişte (Kazançlı)"
        elif last_5[-1] < last_5[0] and recent_cum_return < 0:
            item["trend_status"] = "📉 Düşüşte (Kayıplı)"
        else:
            item["trend_status"] = "➡️ Yatay / Kararsız"
    else:
        item["trend_status"] = "➡️ Veri Sınırlı"

    score = item["kgdm_skor"]
    
    if score >= 60: 
        item["karar"] = "GÜÇLÜ AL (≥60 Puan)"
        item["karar_sira"] = 1
    elif score >= 40: 
        item["karar"] = "ASIL LİSTE (40-59 Puan)"
        item["karar_sira"] = 2
    elif score >= 25: 
        item["karar"] = "NÖTR / İZLEME (25-39 Puan)"
        item["karar_sira"] = 3
    else: 
        if "Yükselişte" in item["trend_status"]:
            item["karar"] = "⚠️ Toparlanma Çabasında (Sat Sinyali Devam Ediyor)"
        else:
            item["karar"] = "ACİL SAT (<25 Puan)"
        item["karar_sira"] = 4

calculated_funds.sort(key=lambda item: (item["karar_sira"], -item["kgdm_skor"], -item["cumulative_return"]))

display_rows = []
for item in calculated_funds:
    display_rows.append({
        "Fon": item["code"], 
        "KGDM-3 (Ort5)": item["kgdm_skor"], 
        "Son 5 Gün Skorlar": item["last_5_scores_str"],
        "5 Günlük Trend": item["trend_status"],
        "Model Kararı": item["karar"],
        "Ort. Getiri %": round(item["mean_return"], 3), 
        "Volatilite %": round(item["volatility"], 3),
        "Sharpe": round(item["sharpe_like"], 3), 
        "Kümülatif Getiri %": round(item["cumulative_return"], 3),
        "Hisse Oranı %": round(item["stock_ratio"], 2),
        "AUM Değişim %": round(item["aum_change"], 2),
        "Yatırımcı Değişim %": round(item["inv_change"], 2),
        "MaxDD %": round(item["max_dd"], 2),
        "AUM ₺": round(item["aum"], 0) if item["aum"] else 0, 
        "Yatırımcı": item["investors"],
        "Valör": item["valor"] if item["valor"] is not None else "-", 
        "Kaynak": item["data_source"]
    })

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value)
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "Yükselişte" in text: return "color: #008000; font-weight: bold;"
    if "NÖTR" in text or "Yatay" in text: return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text or "Toparlanma Çabasında" in text: return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except AttributeError: styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 KGDM-3 Fon Sonuçları (Varlık Dağılımı Entegreli)")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

strong_buy_count = sum(item["kgdm_skor"] >= 60 for item in calculated_funds)
main_list_count = sum(40 <= item["kgdm_skor"] < 60 for item in calculated_funds)
neutral_count = sum(25 <= item["kgdm_skor"] < 40 for item in calculated_funds)
sell_count = sum(item["kgdm_skor"] < 25 for item in calculated_funds)

st.subheader("📌 Karar Dağılımı")
c1, c2, c3, c4 = st.columns(4)
with c1: st.metric("🟢 Güçlü Al", strong_buy_count)
with c2: st.metric("🟢 Asıl Liste", main_list_count)
with c3: st.metric("🟡 Nötr", neutral_count)
with c4: st.metric("🔴 Acil Sat", sell_count)

if investor_filtered_out:
    with st.expander(f"🚫 Yatırımcı Sayısı 10.000'den Az Olduğu İçin Elenen Fonlar ({len(investor_filtered_out)})"):
        st.dataframe(pd.DataFrame(investor_filtered_out), use_container_width=True, hide_index=True)

source_df = pd.DataFrame([{"Fon": item["code"], "Kaynak": item["data_source"]} for item in calculated_funds])
source_distribution = source_df["Kaynak"].value_counts().rename_axis("Kaynak").reset_index(name="Fon Sayısı")

with st.expander("🔎 Veri Kaynağı Dağılımı"):
    st.dataframe(source_distribution, use_container_width=True, hide_index=True)

if source_errors:
    with st.expander(f"⚠️ Veri alınamayan fonlar ({len(source_errors)})"):
        st.dataframe(pd.DataFrame(source_errors), use_container_width=True, hide_index=True)

output = create_excel_output(wb=wb, ws_list=ws_list, calculated_funds=calculated_funds, n_days=n_days)
st.success(f"✅ Analiz tamamlandı. ({source_mode.upper()} kaynağından okundu ve hisse oranları tarandı).")

st.download_button(label="📥 Güncellenmiş Excel'i İndir", data=output, file_name="fonlar_guncel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption(f"KGDM-3 Fon Analiz Otomasyonu v{APP_VERSION} | Analiz tarihi: {today.strftime('%d.%m.%Y')}")
