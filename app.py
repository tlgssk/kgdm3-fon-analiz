import datetime as dt
import io
import math
from typing import Optional, List, Dict, Any

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# SAYFA AYARLARI
# ============================================================

st.set_page_config(
    page_title="Multi-Vade Fon Analizi V5",
    page_icon="📈",
    layout="wide",
)

st.title("📈 Multi-Vade Fon Analizi V5")
st.caption(
    "Tam Düzeltilmiş Z-Skor Mantığı + None/0 Ayrımı + Filtre/Kategori Kusursuzlaştırması"
)

# ============================================================
# GENEL AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")
LOOKBACK_CALENDAR_DAYS = 400
HTTP_TIMEOUT = 15

PERIODS = {
    "Kısa Vade (1 Hafta)": 5,
    "Orta Vade (1 Ay)": 21,
    "Uzun Vade (3 Ay)": 63,
    "Çok Uzun Vade (1 Yıl)": 252,
}

MAX_DAYS = max(PERIODS.values())

FINAL_WEIGHTS = {
    5: 0.10,
    21: 0.20,
    63: 0.30,
    252: 0.40,
}

KUT_GOLD_WEIGHT = 0.45
KUT_SILVER_WEIGHT = 0.45
KUT_CASH_WEIGHT = 0.10

COLOR_NAVY = "1F4E79"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"
COLOR_LIGHT_GREEN = "E2F0D9"
COLOR_LIGHT_RED = "FCE4D6"
COLOR_LIGHT_YELLOW = "FFF2CC"

# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

def parse_number(value) -> Optional[float]:
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)):
        try:
            if pd.isna(value): return None
        except: pass
        return float(value)

    text = str(value).strip()
    if not text: return None
    text = text.replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
    if not text: return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."): text = text.replace(".", "").replace(",", ".")
        else: text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try: return float(text)
    except (ValueError, TypeError): return None

def normalize_fund_code(value) -> str:
    if value is None: return ""
    code = str(value).strip().upper()
    return code[:-2] if code.endswith(".0") else code

def safe_float(value, default=0.0):
    try:
        if value is None: return default
        value = float(value)
        return default if math.isnan(value) or math.isinf(value) else value
    except Exception:
        return default

def infer_category(title: str) -> str:
    if not title: return "Diğer"
    t = title.upper()
    if any(k in t for k in ["ALTIN", "GOLD", "KIYMETLİ MADEN", "GÜMÜŞ", "SILVER", "KITA"]): return "Kıymetli Maden"
    if any(k in t for k in ["HİSSE", "EQUITY", "HİSSE SENEDİ", "HİSSE YOĞUN"]): return "Hisse Senedi"
    if any(k in t for k in ["PARA PİYASASI", "LIKIT", "LİKİT", "PARA PIYASASI"]): return "Para Piyasası"
    if any(k in t for k in ["BORÇLANMA", "TAHVİL", "BONO", "BORÇLANMA ARAÇLARI", "KİRA SERTİFİKASI"]): return "Borçlanma"
    if any(k in t for k in ["KARMA", "DEĞİŞKEN", "DENGELİ", "ÇOKLU VARLIK", "FON SEPETİ"]): return "Karma / Değişken"
    if any(k in t for k in ["YABANCI", "EUROBOND", "DIŞ BORÇ", "USD", "EUR", "DÖVİZ"]): return "Yabancı / Döviz"
    if any(k in t for k in ["SERBEST"]): return "Serbest"
    if any(k in t for k in ["KATILIM"]): return "Katılım"
    return "Diğer"

# ============================================================
# TEFAS VERİSİ
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try: from pytefas import Crawler
    except ImportError: return pd.DataFrame()

    try:
        crawler = Crawler(timeout=60, max_retry=5)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is None or df.empty: return pd.DataFrame()

        df = df.copy()
        api_category_col = next((col for col in ["fund_umbrella_title", "umbrella_fund_type", "fon_turu"] if col in df.columns), None)
        df["spk_category"] = df[api_category_col].astype(str) if api_category_col else ""

        df.rename(columns={"fund_code": "code", "fund_name": "title", "investor_count": "investors", "portfolio_size": "aum"}, inplace=True)
        if not all(c in df.columns for c in ["date", "code", "price"]): return pd.DataFrame()

        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["price"] = df["price"].apply(parse_number)
        df["aum"] = df["aum"].apply(parse_number) if "aum" in df.columns else 0.0
        df["investors"] = df["investors"].apply(parse_number) if "investors" in df.columns else 0.0
        if "title" not in df.columns: df["title"] = ""
            
        df = df.dropna(subset=["date", "code", "price"])
        df = df[df["price"] > 0]
        return df.sort_values(["code", "date"]).drop_duplicates(subset=["code", "date"], keep="last").reset_index(drop=True)
    except Exception: return pd.DataFrame()

def get_fund_series(universe: pd.DataFrame, fund_code: str) -> Optional[pd.DataFrame]:
    if universe is None or universe.empty: return None
    code = normalize_fund_code(fund_code)
    rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
    if rows.empty: return None
    rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
    if len(rows) < 2: return None
    if len(rows) > MAX_DAYS + 1: rows = rows.tail(MAX_DAYS + 1)
    return rows.reset_index(drop=True)

def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code: return None
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    
    try:
        response = requests.get(url, params=params, headers={"User-Agent": "Mozilla/5.0"}, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
        values = response.json().get("value")
        if not values: return None

        df = pd.DataFrame(values)
        if "Tarih" not in df.columns or "Fiyat" not in df.columns: return None

        df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
        df["price"] = df["Fiyat"].apply(parse_number)
        df = df.dropna(subset=["date", "price"])
        df = df[df["price"] > 0]
        if len(df) < 2: return None

        return df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(MAX_DAYS + 1).reset_index(drop=True)[["date", "price"]]
    except Exception: return None

# ============================================================
# FON METRİKLERİ + MDD + SORTINO
# ============================================================

def compute_max_drawdown(prices: List[float]) -> float:
    if not prices or len(prices) < 2: return 0.0
    peak, max_dd = prices[0], 0.0
    for p in prices:
        if p > peak: peak = p
        if peak > 0:
            dd = (p / peak - 1.0) * 100.0
            if dd < max_dd: max_dd = dd
    return max_dd

def compute_sortino(returns: List[float], daily_rf: float = 0.0) -> float:
    if not returns: return 0.0
    mean_excess = (sum(returns) / len(returns)) - daily_rf
    downside = [r - daily_rf for r in returns if r < daily_rf]
    if not downside: return mean_excess * 10 if mean_excess > 0 else 0.0  
    downside_vol = math.sqrt(sum(r ** 2 for r in downside) / len(downside))
    return mean_excess / downside_vol if downside_vol > 1e-12 else 0.0

def compute_fund_metrics(series: Optional[pd.DataFrame]) -> Optional[dict]:
    if series is None or len(series) < 2: return None
    df = series.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["price"] = df["price"].apply(parse_number)
    df = df.dropna(subset=["date", "price"])
    df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    if len(df) < 2: return None

    prices = df["price"].astype(float).tolist()
    daily_returns = [(prices[i] / prices[i-1] - 1) * 100 if prices[i-1] > 0 else 0.0 for i in range(1, len(prices))]
    
    title = str(df["title"].iloc[-1] or "") if "title" in df.columns else ""
    spk_cat = str(df["spk_category"].iloc[-1]).strip() if "spk_category" in df.columns else ""
    final_category = spk_cat if spk_cat and spk_cat.lower() not in ["nan", "none", "null"] else infer_category(title)

    return {
        "prices": prices,
        "daily_returns": daily_returns,
        "title": title,
        "aum": safe_float(df["aum"].iloc[-1]) if "aum" in df.columns else 0.0,
        "investors": safe_float(df["investors"].iloc[-1]) if "investors" in df.columns else 0.0,
        "category": final_category,
    }

# ============================================================
# VADE SKORU
# ============================================================

def calculate_period_scores(funds: List[dict], days: int, daily_rf: float, use_category: bool):
    # 1. Aşama: Tüm fonlar için (filtre dışı olanlar dahil) ham metrikleri hesapla
    for item in funds:
        if len(item["daily_returns"]) < days:
            item[f"score_{days}"] = None
            item[f"karar_{days}"] = "YETERSİZ VERİ"
            for m in ["mean", "vol", "shp", "cum", "mdd", "sortino"]: item[f"{m}_{days}"] = None
            continue

        slice_ret = item["daily_returns"][-days:]
        slice_prices = item["prices"][-(days + 1):]
        mean_ret = sum(slice_ret) / len(slice_ret)
        vol = (sum((r - mean_ret) ** 2 for r in slice_ret) / len(slice_ret)) ** 0.5

        item[f"mean_{days}"] = mean_ret
        item[f"vol_{days}"] = vol
        item[f"shp_{days}"] = (mean_ret - daily_rf) / vol if vol > 1e-12 else 0.0
        item[f"cum_{days}"] = (slice_prices[-1] / slice_prices[0] - 1) * 100
        item[f"mdd_{days}"] = compute_max_drawdown(slice_prices)
        item[f"sortino_{days}"] = compute_sortino(slice_ret, daily_rf)

    # 2. Aşama: Sadece geçerli ve filtreden GEÇEN fonları değerlendirme havuzuna (eval) al
    valid_indices = [i for i, f in enumerate(funds) if not f.get("filtered_out") and f.get(f"cum_{days}") is not None]

    categories = {}
    if use_category:
        for idx in valid_indices:
            cat = funds[idx].get("category", "Diğer")
            categories.setdefault(cat, []).append(idx)
    else:
        categories = {"ALL": valid_indices}

    # İstatistiksel hesaplama yardımcısı
    def get_stats(vals):
        if not vals: return 0.0, 0.0
        m = sum(vals) / len(vals)
        std = (sum((v - m) ** 2 for v in vals) / len(vals)) ** 0.5
        return m, std

    def apply_scoring(eval_indices, target_indices):
        if len(eval_indices) < 2:
            for idx in target_indices:
                funds[idx][f"score_{days}"] = None
                funds[idx][f"karar_{days}"] = "YETERSİZ VERİ"
            return

        m_mean, s_mean = get_stats([funds[i][f"mean_{days}"] for i in eval_indices])
        m_shp, s_shp = get_stats([funds[i][f"shp_{days}"] for i in eval_indices])
        m_cum, s_cum = get_stats([funds[i][f"cum_{days}"] for i in eval_indices])
        m_srt, s_srt = get_stats([funds[i][f"sortino_{days}"] for i in eval_indices])
        m_mdd, s_mdd = get_stats([funds[i][f"mdd_{days}"] for i in eval_indices])

        def z(val, m, std): return (val - m) / std if std > 1e-12 else 0.0

        for idx in target_indices:
            item = funds[idx]
            z_m = z(item[f"mean_{days}"], m_mean, s_mean)
            z_s = z(item[f"shp_{days}"], m_shp, s_shp)
            z_c = z(item[f"cum_{days}"], m_cum, s_cum)
            z_srt = z(item[f"sortino_{days}"], m_srt, s_srt)
            z_mdd = z(item[f"mdd_{days}"], m_mdd, s_mdd)

            raw_score = 50 + 12 * z_m + 15 * z_s + 12 * z_c + 12 * z_srt + 10 * z_mdd - (safe_float(item.get("valor", 0)) * 1.2)
            score = int(round(max(0, min(100, raw_score))))
            
            item[f"score_{days}"] = score
            if score >= 60: item[f"karar_{days}"] = "GÜÇLÜ AL"
            elif score >= 40: item[f"karar_{days}"] = "ASIL LİSTE"
            elif score >= 25: item[f"karar_{days}"] = "NÖTR"
            else: item[f"karar_{days}"] = "ACİL SAT"

    # 3. Aşama: Kategori bazlı skorlama (Tek fonlu kategoriler için Global Havuz kurtarması)
    for cat, indices in categories.items():
        if len(indices) < 2:
            apply_scoring(eval_indices=valid_indices, target_indices=indices) # Kendi kategorisinde tekse, tüm fonlarla yarıştır
        else:
            apply_scoring(eval_indices=indices, target_indices=indices)

    # 4. Aşama: Filtreden elenenlere özel durum ataması (Ham verileri korundu, sadece skor iptal)
    for idx, item in enumerate(funds):
        if item.get("filtered_out") and item.get(f"cum_{days}") is not None:
            item[f"score_{days}"] = None
            item[f"karar_{days}"] = "FİLTRE DIŞI"

# ============================================================
# YAHOO + KUT BENCHMARK
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_yahoo_series(symbol: str, start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    try:
        p1 = int(dt.datetime.combine(start_date, dt.time.min).timestamp())
        p2 = int(dt.datetime.combine(end_date + dt.timedelta(days=1), dt.time.min).timestamp())
        res = requests.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}", params={"period1": p1, "period2": p2, "interval": "1d", "events": "history", "includeAdjustedClose": "true"}, headers={"User-Agent": "Mozilla/5.0"}, timeout=HTTP_TIMEOUT)
        res.raise_for_status()
        result = res.json().get("chart", {}).get("result", [{}])[0]
        timestamps, closes = result.get("timestamp"), result.get("indicators", {}).get("quote", [{}])[0].get("close")
        if not timestamps or not closes: return None

        rows = [{"date": pd.Timestamp(dt.datetime.fromtimestamp(t).date()), "price": float(c)} for t, c in zip(timestamps, closes) if c is not None and float(c) > 0]
        if len(rows) < 2: return None
        return pd.DataFrame(rows).sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    except: return None

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_kut_benchmark(start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    gold, silver, usdtry = fetch_yahoo_series("GC=F", start_date, end_date), fetch_yahoo_series("SI=F", start_date, end_date), fetch_yahoo_series("USDTRY=X", start_date, end_date)
    if gold is None or silver is None or usdtry is None: return None

    df = gold.rename(columns={"price": "gold"}).merge(silver.rename(columns={"price": "silver"}), on="date", how="inner").merge(usdtry.rename(columns={"price": "usd"}), on="date", how="inner")
    if df.empty: return None

    g_try, s_try = df["gold"] * df["usd"], df["silver"] * df["usd"]
    df["benchmark"] = (KUT_GOLD_WEIGHT * (g_try / g_try.iloc[0])) + (KUT_SILVER_WEIGHT * (s_try / s_try.iloc[0])) + KUT_CASH_WEIGHT
    return df[["date", "benchmark"]].reset_index(drop=True)

def calculate_kut_benchmark_metrics(funds, benchmark_df):
    bench_returns = {d: ((benchmark_df["benchmark"].iloc[-1] / benchmark_df["benchmark"].iloc[-(d+1)] - 1) * 100) if benchmark_df is not None and not benchmark_df.empty and len(benchmark_df) > d and benchmark_df["benchmark"].iloc[-(d+1)] > 0 else None for d in PERIODS.values()}

    for item in funds:
        if item["code"] != "KUT":
            item["benchmark_active"] = False
            continue
        item["benchmark_active"] = True
        for days in PERIODS.values():
            fr, br = item.get(f"cum_{days}"), bench_returns.get(days)
            if fr is None or br is None:
                item[f"benchmark_{days}"], item[f"benchmark_diff_{days}"], item[f"benchmark_score_{days}"] = None, None, None
            else:
                diff = fr - br
                item[f"benchmark_{days}"], item[f"benchmark_diff_{days}"], item[f"benchmark_score_{days}"] = br, diff, int(round(max(0, min(100, 50 + diff * 5))))

# ============================================================
# TREND + GÜVEN + NİHAİ SKOR
# ============================================================

def calculate_trends(funds):
    for item in funds:
        s63, s252, s21 = item.get("score_63"), item.get("score_252"), item.get("score_21")
        item["trend_3a"] = "VERİ YOK" if s63 is None or s21 is None else ("YUKARI ↑" if s21 - s63 >= 10 else "AŞAĞI ↓" if s21 - s63 <= -10 else "YATAY →")
        item["trend_1y"] = "VERİ YOK" if s252 is None or s63 is None else ("YUKARI ↑" if s63 - s252 >= 10 else "AŞAĞI ↓" if s63 - s252 <= -10 else "YATAY →")

def calculate_confidence(item):
    scores = [float(item[f"score_{d}"]) for d in PERIODS.values() if item.get(f"score_{d}") is not None]
    if not scores: return "DÜŞÜK"
    data_ratio, mean = len(scores) / len(PERIODS), sum(scores) / len(scores)
    dispersion = (sum((x - mean) ** 2 for x in scores) / len(scores)) ** 0.5
    if data_ratio == 1.0 and dispersion <= 10: return "ÇOK YÜKSEK"
    if data_ratio >= 0.75 and dispersion <= 15: return "YÜKSEK"
    return "ORTA" if data_ratio >= 0.50 else "DÜŞÜK"

def calculate_final_scores(funds):
    for item in funds:
        w_sum = sum(item[f"score_{d}"] * w for d, w in FINAL_WEIGHTS.items() if item.get(f"score_{d}") is not None)
        w_tot = sum(w for d, w in FINAL_WEIGHTS.items() if item.get(f"score_{d}") is not None)
        
        if w_tot == 0:
            item["final_score"], item["final_decision"], item["confidence"] = None, "YETERSİZ VERİ", "DÜŞÜK"
            continue

        base_final = w_sum / w_tot
        if item.get("benchmark_active"):
            b_scores = [item[f"benchmark_score_{d}"] for d in PERIODS.values() if item.get(f"benchmark_score_{d}") is not None]
            final_score = (base_final * 0.80 + (sum(b_scores)/len(b_scores)) * 0.20) if b_scores else base_final
        else:
            final_score = base_final

        final_score = int(round(max(0, min(100, final_score))))
        item["final_score"] = final_score
        item["final_decision"] = "GÜÇLÜ AL" if final_score >= 70 else "AL" if final_score >= 60 else "ASIL LİSTE" if final_score >= 45 else "NÖTR" if final_score >= 30 else "ACİL SAT"
        item["confidence"] = calculate_confidence(item)

# ============================================================
# ARAYÜZ VE ÇALIŞTIRMA MANTIĞI
# ============================================================

with st.sidebar:
    st.header("⚙️ Filtreler & Ayarlar")
    annual_rf_rate = st.number_input("Yıllık Risksiz Getiri Oranı (%)", min_value=0.0, max_value=150.0, value=50.0, step=5.0)
    min_aum = st.number_input("Minimum Portföy Büyüklüğü (TL)", min_value=0.0, value=50_000_000.0, step=10_000_000.0)
    min_investors = st.number_input("Minimum Yatırımcı Sayısı", min_value=0, value=100, step=50)
    use_category_scoring = st.checkbox("Kategori bazlı skorlama kullan", value=True)

daily_rf = annual_rf_rate / 252.0

uploaded_file = st.file_uploader("Excel Dosyanızı Yükleyin (Fon_Listesi içeren):", type=["xlsx"])
if not uploaded_file: st.stop()
wb = openpyxl.load_workbook(uploaded_file)
if "Fon_Listesi" not in wb.sheetnames: st.error("Dosyada 'Fon_Listesi' sayfası yok!"); st.stop()

requested_codes, valor_map = [], {}
for row in wb["Fon_Listesi"].iter_rows(min_row=2, values_only=False):
    if row and row[0].value:
        code = normalize_fund_code(row[0].value)
        if code:
            requested_codes.append(code)
            valor_map[code] = parse_number(row[1].value) if len(row) > 1 and row[1].value is not None else 0.0

requested_codes = list(dict.fromkeys(requested_codes))
today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

with st.spinner("Veriler işleniyor..."):
    universe = fetch_tefas_universe(start_date, today)
    calculated_funds, failed_codes = [], []

    for code in requested_codes:
        series, source = get_fund_series(universe, code) if not universe.empty else None, "TEFAS" if not universe.empty else "Bulunamadı"
        if series is None:
            series = fetch_isyatirim_series(code)
            if series is not None: source = "İş Yatırım"

        metrics = compute_fund_metrics(series)
        if metrics:
            metrics.update({"code": code, "source": source, "valor": valor_map.get(code, 0.0)})
            metrics["filtered_out"] = not ((metrics["aum"] >= min_aum if metrics["aum"] > 0 else True) and (metrics["investors"] >= min_investors if metrics["investors"] > 0 else True))
            calculated_funds.append(metrics)
        else: failed_codes.append(code)

if failed_codes: st.warning(f"Veri bulunamayan fonlar: {', '.join(failed_codes)}")

for period_days in PERIODS.values():
    calculate_period_scores(calculated_funds, period_days, daily_rf, use_category_scoring)

if any(item["code"] == "KUT" for item in calculated_funds):
    calculate_kut_benchmark_metrics(calculated_funds, fetch_kut_benchmark(start_date, today))

calculate_trends(calculated_funds)
calculate_final_scores(calculated_funds)

calculated_funds.sort(key=lambda x: x.get("final_score") if x.get("final_score") is not None else -1, reverse=True)

# ============================================================
# EXCEL ÇIKTISI
# ============================================================

if "Vade_Analizi" in wb.sheetnames: del wb["Vade_Analizi"]
ws_out = wb.create_sheet("Vade_Analizi", 0)

headers = ["Fon Kodu", "Kategori", "AUM (TL)", "Yatırımcı", "1H Skor", "1H Karar", "1H Küm %", "1H MDD %", "1A Skor", "1A Karar", "1A Küm %", "1A MDD %", "3A Skor", "3A Karar", "3A Küm %", "3A MDD %", "1Y Skor", "1Y Karar", "1Y Küm %", "1Y MDD %", "1Y Trend", "3A Trend", "Güven Seviyesi", "Nihai Skor", "Nihai Karar", "Benchmark 1H", "KUT-Bench 1H", "Benchmark 1A", "KUT-Bench 1A", "Benchmark 3A", "KUT-Bench 3A", "Benchmark 1Y", "KUT-Bench 1Y", "Benchmark Ort. Skor", "Kaynak"]
ws_out.append(headers)

for cell in ws_out[1]:
    cell.fill, cell.font, cell.alignment = PatternFill(start_color=COLOR_NAVY, fill_type="solid"), Font(color=COLOR_WHITE, bold=True), Alignment(horizontal="center", vertical="center", wrap_text=True)

def fmt(val): return val if val is not None else "-"

for item in calculated_funds:
    b_scores = [item.get(f"benchmark_score_{d}") for d in PERIODS.values() if item.get(f"benchmark_score_{d}") is not None]
    ws_out.append([
        item["code"], item.get("category", "-"), fmt(item.get("aum")), fmt(item.get("investors")),
        fmt(item.get("score_5")), item.get("karar_5", "-"), fmt(item.get("cum_5")), fmt(item.get("mdd_5")),
        fmt(item.get("score_21")), item.get("karar_21", "-"), fmt(item.get("cum_21")), fmt(item.get("mdd_21")),
        fmt(item.get("score_63")), item.get("karar_63", "-"), fmt(item.get("cum_63")), fmt(item.get("mdd_63")),
        fmt(item.get("score_252")), item.get("karar_252", "-"), fmt(item.get("cum_252")), fmt(item.get("mdd_252")),
        item.get("trend_1y", "-"), item.get("trend_3a", "-"), item.get("confidence", "-"),
        fmt(item.get("final_score")), item.get("final_decision", "-"),
        fmt(item.get("benchmark_5")), fmt(item.get("benchmark_diff_5")),
        fmt(item.get("benchmark_21")), fmt(item.get("benchmark_diff_21")),
        fmt(item.get("benchmark_63")), fmt(item.get("benchmark_diff_63")),
        fmt(item.get("benchmark_252")), fmt(item.get("benchmark_diff_252")),
        round(sum(b_scores)/len(b_scores), 1) if b_scores else "-", item.get("source", "-")
    ])

green, red, yellow = Font(color=COLOR_GREEN, bold=True), Font(color=COLOR_RED, bold=True), Font(color=COLOR_YELLOW, bold=True)
for row in range(2, ws_out.max_row + 1):
    for col in [6, 10, 14, 18, 25]:
        v = str(ws_out.cell(row=row, column=col).value)
        if any(x in v for x in ["GÜÇLÜ AL", "AL", "LİSTE"]): ws_out.cell(row=row, column=col).font = green
        elif "NÖTR" in v: ws_out.cell(row=row, column=col).font = yellow
        elif any(x in v for x in ["SAT", "ACİL", "FİLTRE"]): ws_out.cell(row=row, column=col).font = red

    for col in [21, 22]:
        v = str(ws_out.cell(row=row, column=col).value)
        if "YUKARI" in v: ws_out.cell(row=row, column=col).font = green
        elif "AŞAĞI" in v: ws_out.cell(row=row, column=col).font = red
        elif "YATAY" in v: ws_out.cell(row=row, column=col).font = yellow

    c_conf = ws_out.cell(row=row, column=23)
    if c_conf.value in ["ÇOK YÜKSEK", "YÜKSEK"]: c_conf.font = green
    elif c_conf.value == "ORTA": c_conf.font = yellow
    else: c_conf.font = red

    c_fin = ws_out.cell(row=row, column=24)
    if isinstance(c_fin.value, (int, float)):
        c_fin.fill = PatternFill(start_color=COLOR_LIGHT_GREEN if c_fin.value >= 70 else (COLOR_LIGHT_YELLOW if c_fin.value >= 45 else COLOR_LIGHT_RED), fill_type="solid")

for row in range(2, ws_out.max_row + 1):
    for col in [7, 8, 11, 12, 15, 16, 19, 20, 26, 27, 28, 29, 30, 31, 32, 33]:
        cell = ws_out.cell(row=row, column=col)
        if isinstance(cell.value, (int, float)): cell.number_format = '0.00"%"'

for col in ws_out.columns:
    ws_out.column_dimensions[get_column_letter(col[0].column)].width = min(max(max([len(str(c.value)) for c in col]) + 2, 11), 26)

ws_out.freeze_panes = "B2"
ws_out.auto_filter.ref = ws_out.dimensions

output = io.BytesIO()
wb.save(output)
output.seek(0)

# ============================================================
# STREAMLIT EKRAN ÇIKTILARI
# ============================================================

st.success("✅ Multi-Vade V5 analizi (Kusursuz Filtreleme ve Z-Skor Kurtarması ile) tamamlandı!")
st.download_button(label="📥 V5 Excel Çıktısını İndir", data=output, file_name="fon_vade_analizi_V5.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.subheader("📊 Fon Sıralaması Önizleme")
df_preview = pd.DataFrame([{ "Fon": f["code"], "Kategori": f.get("category", "-"), "1H": fmt(f.get("score_5")), "1A": fmt(f.get("score_21")), "3A": fmt(f.get("score_63")), "1Y": fmt(f.get("score_252")), "Nihai Skor": fmt(f.get("final_score")), "Nihai Karar": f.get("final_decision", "-"), "AUM": fmt(f.get("aum")) } for f in calculated_funds])
st.dataframe(df_preview, use_container_width=True, hide_index=True)
