import datetime as dt
import io
import math
import re
from typing import Optional, List

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# SAYFA AYARLARI
# ============================================================
st.set_page_config(page_title="Multi-Vade Fon Analizi V9", page_icon="📈", layout="wide")
st.title("📈 Multi-Vade Fon Analizi V9")
st.caption("Anti-Blokaj Veri Motoru (Doğrudan TEFAS API) + Gelişmiş Z-Skor + Tanh Benchmark")

# ============================================================
# GENEL AYARLAR
# ============================================================
LOOKBACK_CALENDAR_DAYS = 400
HTTP_TIMEOUT = 10

PERIODS = {"Kısa Vade (1 Hafta)": 5, "Orta Vade (1 Ay)": 21, "Uzun Vade (3 Ay)": 63, "Çok Uzun Vade (1 Yıl)": 252}
MAX_DAYS = max(PERIODS.values())
FINAL_WEIGHTS = {5: 0.10, 21: 0.20, 63: 0.30, 252: 0.40}
MIN_FINAL_PERIODS = 3

W_CUM, W_SHP, W_SRT, W_MDD = 0.25, 0.25, 0.25, 0.25
Z_LIMIT = 2.5
VALOR_PENALTY_MAX, VALOR_SCALE = 5.0, 5.0

KUT_GOLD_WEIGHT, KUT_SILVER_WEIGHT, KUT_CASH_WEIGHT = 0.45, 0.45, 0.10

COLOR_NAVY, COLOR_GREEN, COLOR_RED, COLOR_YELLOW, COLOR_WHITE = "1F4E79", "008000", "FF0000", "B8860B", "FFFFFF"
COLOR_LIGHT_GREEN, COLOR_LIGHT_RED, COLOR_LIGHT_YELLOW = "E2F0D9", "FCE4D6", "FFF2CC"

# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================
def parse_number(value) -> Optional[float]:
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)):
        try:
            if pd.isna(value): return None
        except Exception: pass
        return float(value)
    text = str(value).strip().replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "")
    if not text: return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
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
    except Exception: return default

def infer_category(title: str) -> str:
    if not title: return "Diğer"
    t = title.upper()
    if any(k in t for k in ["ALTIN", "GOLD", "KIYMETLİ MADEN", "GÜMÜŞ", "SILVER"]): return "Kıymetli Maden"
    if any(k in t for k in ["HİSSE", "EQUITY", "HİSSE SENEDİ"]): return "Hisse Senedi"
    if any(k in t for k in ["PARA PİYASASI", "LIKIT", "LİKİT"]): return "Para Piyasası"
    if any(k in t for k in ["BORÇLANMA", "TAHVİL", "BONO", "KİRA SERTİFİKASI"]): return "Borçlanma"
    if any(k in t for k in ["KARMA", "DEĞİŞKEN", "DENGELİ", "FON SEPETİ"]): return "Karma / Değişken"
    if any(k in t for k in ["YABANCI", "EUROBOND", "DIŞ BORÇ", "USD", "EUR", "DÖVİZ"]): return "Yabancı / Döviz"
    if "SERBEST" in t: return "Serbest"
    if "KATILIM" in t: return "Katılım"
    return "Diğer"

# ============================================================
# CERRAHİ VERİ MOTORLARI (V9 - Doğrudan API Bağlantıları)
# ============================================================
@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_fund_metadata(code: str) -> dict:
    """TEFAS fon detay sayfasından HTML kazıyarak AUM ve Yatırımcı çeker (Engellenemez)."""
    url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={code.upper()}"
    headers = {"User-Agent": "Mozilla/5.0"}
    meta = {"aum": 0.0, "investors": 0.0, "title": "", "category": ""}
    try:
        res = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        if res.status_code == 200:
            html = res.text
            m_title = re.search(r'id="MainContent_FormViewMainIndicators_LabelFund"?[^>]*>([^<]+)</span>', html)
            if m_title: 
                meta["title"] = m_title.group(1).strip()
                meta["category"] = infer_category(meta["title"])
                
            m_aum = re.search(r'id="MainContent_FormViewMainIndicators_LabelPortfolioValue"?[^>]*>([^<]+)</span>', html)
            if m_aum: meta["aum"] = parse_number(m_aum.group(1)) or 0.0
            
            m_inv = re.search(r'id="MainContent_FormViewMainIndicators_LabelInvestorCount"?[^>]*>([^<]+)</span>', html)
            if m_inv: meta["investors"] = parse_number(m_inv.group(1)) or 0.0
    except: pass
    return meta

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_prices(code: str, start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    """TEFAS'ın grafik çizmek için kullandığı gizli API'sinden doğrudan fiyat çeker."""
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    data = {
        "fontip": "", "sfontur": "", "fonkod": code.upper(), "fongrup": "",
        "baslangic": start_date.strftime("%d.%m.%Y"), "bitis": end_date.strftime("%d.%m.%Y")
    }
    headers = {
        "User-Agent": "Mozilla/5.0", "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    }
    try:
        res = requests.post(url, data=data, headers=headers, timeout=HTTP_TIMEOUT)
        if res.status_code == 200:
            data_list = res.json().get("data", [])
            if data_list:
                rows = [{"date": pd.Timestamp(dt.datetime.fromtimestamp(item["TARIH"] / 1000.0).date()), "price": float(item["FIYAT"])} for item in data_list if item.get("TARIH") and item.get("FIYAT")]
                df = pd.DataFrame(rows).dropna()
                df = df[df["price"] > 0]
                if len(df) >= 2:
                    return df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(MAX_DAYS + 1).reset_index(drop=True)
    except: pass
    return None

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_isyatirim_series(code: str, start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code.upper(), "baslangic": start_date.strftime("%d-%m-%Y"), "bitis": end_date.strftime("%d-%m-%Y")}
    try:
        res = requests.get(url, params=params, headers={"User-Agent": "Mozilla/5.0"}, timeout=HTTP_TIMEOUT)
        values = res.json().get("value")
        if values:
            df = pd.DataFrame(values)
            df["date"] = pd.to_datetime(df.get("Tarih"), dayfirst=True, errors="coerce")
            df["price"] = df.get("Fiyat").apply(parse_number)
            df = df.dropna(subset=["date", "price"])
            df = df[df["price"] > 0]
            if len(df) >= 2:
                return df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(MAX_DAYS + 1).reset_index(drop=True)[["date", "price"]]
    except: pass
    return None

# ============================================================
# METRİKLER (MDD, SORTINO)
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

def compute_sortino(returns: List[float], daily_rf: float = 0.0, max_sortino: float = 10.0) -> float:
    if not returns: return 0.0
    clean_returns = [float(r) for r in returns if r is not None and math.isfinite(float(r))]
    if not clean_returns: return 0.0
    excess = [r - daily_rf for r in clean_returns]
    mean_excess = sum(excess) / len(excess)
    downside_squared = [min(0.0, x) ** 2 for x in excess]
    downside_deviation = math.sqrt(sum(downside_squared) / len(excess))
    if downside_deviation <= 1e-12: return max_sortino if mean_excess > 0 else 0.0
    return max(-max_sortino, min(max_sortino, mean_excess / downside_deviation))

def compute_fund_metrics(series: Optional[pd.DataFrame], metadata: dict) -> Optional[dict]:
    if series is None or len(series) < 2: return None
    prices = series["price"].astype(float).tolist()
    daily_returns = [(prices[i] / prices[i - 1] - 1) * 100 if prices[i - 1] > 0 else 0.0 for i in range(1, len(prices))]
    return {
        "prices": prices, "daily_returns": daily_returns,
        "title": metadata.get("title", ""), "category": metadata.get("category", "Diğer"),
        "aum": metadata.get("aum", 0.0), "investors": metadata.get("investors", 0.0)
    }

def calculate_valor_penalty(valor) -> float:
    return VALOR_PENALTY_MAX * math.tanh(safe_float(valor, 0.0) / VALOR_SCALE)

# ============================================================
# Z-SKOR MOTORU
# ============================================================
def calculate_period_scores(funds: List[dict], days: int, daily_rf: float, use_category: bool):
    for item in funds:
        returns, prices = item.get("daily_returns", []), item.get("prices", [])
        if len(returns) < days or len(prices) < days + 1:
            item[f"score_{days}"], item[f"karar_{days}"] = None, "YETERSİZ VERİ"
            for m in ["mean", "vol", "shp", "cum", "mdd", "sortino"]: item[f"{m}_{days}"] = None
            continue

        slice_ret, slice_prices = returns[-days:], prices[-(days + 1):]
        mean_ret = sum(slice_ret) / len(slice_ret)
        vol = math.sqrt(max(0.0, sum((r - mean_ret) ** 2 for r in slice_ret) / len(slice_ret)))
        item[f"mean_{days}"] = mean_ret
        item[f"vol_{days}"] = vol
        item[f"shp_{days}"] = ((mean_ret - daily_rf) / vol * math.sqrt(252)) if vol > 1e-12 else 0.0
        item[f"cum_{days}"] = (slice_prices[-1] / slice_prices[0] - 1.0) * 100.0 if slice_prices[0] > 0 else None
        item[f"mdd_{days}"] = compute_max_drawdown(slice_prices)
        item[f"sortino_{days}"] = compute_sortino(slice_ret, daily_rf=daily_rf)

    valid_indices = [i for i, f in enumerate(funds) if not f.get("filtered_out", False) and all(f.get(f"{m}_{days}") is not None for m in ["cum", "shp", "sortino", "mdd"])]
    if not valid_indices:
        for item in funds:
            item[f"score_{days}"] = None
            item[f"karar_{days}"] = "FİLTRE DIŞI" if item.get("filtered_out") else "YETERSİZ VERİ"
        return

    categories = {}
    if use_category:
        for idx in valid_indices: categories.setdefault(funds[idx].get("category") or "Diğer", []).append(idx)
    else: categories = {"ALL": valid_indices}

    def get_stats(indices, field):
        vals = [float(funds[i][f"{field}_{days}"]) for i in indices if funds[i].get(f"{field}_{days}") is not None and math.isfinite(float(funds[i][f"{field}_{days}"]))]
        if not vals: return 0.0, 0.0
        m = sum(vals) / len(vals)
        return m, math.sqrt(max(0.0, sum((x - m) ** 2 for x in vals) / len(vals)))

    def apply_scoring(eval_indices, target_indices):
        if len(eval_indices) < 2:
            for idx in target_indices: funds[idx][f"stat_score_{days}"] = None
            return
        m_cum = get_stats(eval_indices, "cum")
        m_shp = get_stats(eval_indices, "shp")
        m_srt = get_stats(eval_indices, "sortino")
        m_mdd = get_stats(eval_indices, "mdd")

        def z(val, m, std): return max(-Z_LIMIT, min(Z_LIMIT, (float(val) - m) / std)) if val is not None and std > 1e-12 else 0.0
        
        for idx in target_indices:
            f = funds[idx]
            w_z = (W_CUM * z(f.get(f"cum_{days}"), *m_cum) + W_SHP * z(f.get(f"shp_{days}"), *m_shp) + W_SRT * z(f.get(f"sortino_{days}"), *m_srt) + W_MDD * z(f.get(f"mdd_{days}"), *m_mdd))
            f[f"stat_score_{days}"] = int(round(max(0.0, min(100.0, 50.0 + 20.0 * w_z))))

    apply_scoring(valid_indices, valid_indices)
    global_scores = {idx: funds[idx].get(f"stat_score_{days}") for idx in valid_indices}

    for cat, indices in categories.items():
        n = len(indices)
        if n >= 10:
            apply_scoring(indices, indices)
            for idx in indices: funds[idx][f"score_{days}"] = funds[idx].get(f"stat_score_{days}")
        elif n >= 5:
            apply_scoring(indices, indices)
            for idx in indices:
                cat_s, glob_s = funds[idx].get(f"stat_score_{days}"), global_scores.get(idx)
                if cat_s is not None and glob_s is not None: funds[idx][f"score_{days}"] = int(round(0.60 * cat_s + 0.40 * glob_s))
                else: funds[idx][f"score_{days}"] = cat_s if cat_s is not None else glob_s
        else:
            for idx in indices: funds[idx][f"score_{days}"] = global_scores.get(idx)

    for item in funds:
        if item.get("filtered_out", False):
            if item.get(f"cum_{days}") is not None: item[f"score_{days}"], item[f"karar_{days}"] = None, "FİLTRE DIŞI"
            continue
        bs = item.get(f"score_{days}")
        if bs is None:
            item[f"karar_{days}"] = "YETERSİZ VERİ"
            continue
        pen = calculate_valor_penalty(item.get("valor", 0))
        item[f"valor_penalty_{days}"] = pen
        s = int(round(max(0.0, min(100.0, float(bs) - pen))))
        item[f"score_{days}"] = s
        item[f"karar_{days}"] = "GÜÇLÜ AL" if s >= 60 else "ASIL LİSTE" if s >= 40 else "NÖTR" if s >= 25 else "ACİL SAT"

# ============================================================
# YAHOO & BENCHMARK
# ============================================================
@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_yahoo_series(symbol: str, start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    try:
        p1 = int(dt.datetime.combine(start_date, dt.time.min).timestamp())
        p2 = int(dt.datetime.combine(end_date + dt.timedelta(days=1), dt.time.min).timestamp())
        res = requests.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}", params={"period1": p1, "period2": p2, "interval": "1d", "events": "history", "includeAdjustedClose": "true"}, headers={"User-Agent": "Mozilla/5.0"}, timeout=HTTP_TIMEOUT)
        res.raise_for_status()
        result = res.json().get("chart", {}).get("result", [{}])[0]
        t_stamps, closes = result.get("timestamp"), result.get("indicators", {}).get("quote", [{}])[0].get("close")
        if not t_stamps or not closes: return None
        rows = [{"date": pd.Timestamp(dt.datetime.fromtimestamp(t).date()), "price": float(c)} for t, c in zip(t_stamps, closes) if c is not None and float(c) > 0]
        if len(rows) < 2: return None
        return pd.DataFrame(rows).sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    except: return None

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_kut_benchmark(start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    gold, silver, usd = fetch_yahoo_series("GC=F", start_date, end_date), fetch_yahoo_series("SI=F", start_date, end_date), fetch_yahoo_series("USDTRY=X", start_date, end_date)
    if gold is None or silver is None or usd is None: return None
    df = gold.rename(columns={"price": "g"}).merge(silver.rename(columns={"price": "s"}), on="date").merge(usd.rename(columns={"price": "u"}), on="date")
    if df.empty: return None
    gt, st = df["g"] * df["u"], df["s"] * df["u"]
    df["benchmark"] = KUT_GOLD_WEIGHT * (gt / gt.iloc[0]) + KUT_SILVER_WEIGHT * (st / st.iloc[0]) + KUT_CASH_WEIGHT
    return df[["date", "benchmark"]].reset_index(drop=True)

def calculate_kut_benchmark_metrics(funds, benchmark_df):
    b_ret = {d: ((benchmark_df["benchmark"].iloc[-1] / benchmark_df["benchmark"].iloc[-(d+1)] - 1) * 100) if benchmark_df is not None and not benchmark_df.empty and len(benchmark_df) > d and benchmark_df["benchmark"].iloc[-(d+1)] > 0 else None for d in PERIODS.values()}
    for item in funds:
        if item["code"] != "KUT": item["benchmark_active"] = False; continue
        item["benchmark_active"] = True
        for days in PERIODS.values():
            fr, br = item.get(f"cum_{days}"), b_ret.get(days)
            if fr is None or br is None: item[f"benchmark_{days}"], item[f"benchmark_diff_{days}"], item[f"benchmark_score_{days}"] = None, None, None
            else: item[f"benchmark_{days}"], item[f"benchmark_diff_{days}"], item[f"benchmark_score_{days}"] = br, fr - br, int(round(max(0.0, min(100.0, 50.0 + 40.0 * math.tanh((fr - br) / 5.0)))))

# ============================================================
# TREND & NİHAİ SKOR
# ============================================================
def calculate_trends(funds):
    for item in funds:
        s63, s252, s21 = item.get("score_63"), item.get("score_252"), item.get("score_21")
        item["trend_3a"] = "VERİ YOK" if s63 is None or s21 is None else ("YUKARI ↑" if s21 - s63 >= 10 else "AŞAĞI ↓" if s21 - s63 <= -10 else "YATAY →")
        item["trend_1y"] = "VERİ YOK" if s252 is None or s63 is None else ("YUKARI ↑" if s63 - s252 >= 10 else "AŞAĞI ↓" if s63 - s252 <= -10 else "YATAY →")

def calculate_confidence(item):
    scores = [float(item[f"score_{d}"]) for d in PERIODS.values() if item.get(f"score_{d}") is not None and math.isfinite(float(item[f"score_{d}"]))]
    if not scores: return "DÜŞÜK"
    pc, dr = len(scores), len(scores) / len(PERIODS)
    disp = math.sqrt(sum((x - (sum(scores) / len(scores))) ** 2 for x in scores) / len(scores))
    if pc == 4 and disp <= 8: return "ÇOK YÜKSEK"
    if pc == 4 and disp <= 15: return "YÜKSEK"
    if pc >= 3 and disp <= 15: return "YÜKSEK"
    if pc >= 3 and disp <= 25: return "ORTA"
    return "ORTA" if dr >= 0.50 else "DÜŞÜK"

def calculate_final_scores(funds):
    for item in funds:
        av_s = [(d, w, float(item[f"score_{d}"])) for d, w in FINAL_WEIGHTS.items() if item.get(f"score_{d}") is not None and math.isfinite(float(item[f"score_{d}"]))]
        if len(av_s) < MIN_FINAL_PERIODS:
            item["final_score"], item["final_decision"], item["confidence"] = None, "YETERSİZ VERİ", calculate_confidence(item)
            continue
        tw, ws = sum(w for _, w, _ in av_s), sum(s * w for _, w, s in av_s)
        base_f = ws / tw if tw > 0 else None
        if base_f is None:
            item["final_score"], item["final_decision"], item["confidence"] = None, "YETERSİZ VERİ", "DÜŞÜK"
            continue
        
        fs = base_f
        if item.get("benchmark_active"):
            bs = [float(item[f"benchmark_score_{d}"]) for d in PERIODS.values() if item.get(f"benchmark_score_{d}") is not None]
            if bs: fs = base_f * 0.80 + (sum(bs) / len(bs)) * 0.20
        
        item["final_score"] = int(round(max(0.0, min(100.0, fs))))
        item["final_decision"] = "GÜÇLÜ AL" if fs >= 70 else "AL" if fs >= 60 else "ASIL LİSTE" if fs >= 45 else "NÖTR" if fs >= 30 else "ACİL SAT"
        item["confidence"] = calculate_confidence(item)

# ============================================================
# STREAMLIT ARAYÜZ
# ============================================================
with st.sidebar:
    st.header("⚙️ Filtreler & Ayarlar")
    annual_rf_rate = st.number_input("Yıllık Risksiz Getiri Oranı (%)", min_value=0.0, max_value=150.0, value=50.0, step=5.0)
    min_aum = st.number_input("Minimum Portföy Büyüklüğü (TL)", min_value=0.0, value=50_000_000.0, step=10_000_000.0)
    min_investors = st.number_input("Minimum Yatırımcı Sayısı", min_value=0, value=100, step=50)
    use_category_scoring = st.checkbox("Kategori bazlı skorlama kullan", value=True)
    st.divider()

daily_rf = ((1 + annual_rf_rate / 100.0) ** (1 / 252.0) - 1) * 100.0

uploaded_file = st.file_uploader("Excel Dosyanızı Yükleyin (Fon_Listesi içeren):", type=["xlsx"])
if not uploaded_file: st.stop()

wb = openpyxl.load_workbook(uploaded_file)
if "Fon_Listesi" not in wb.sheetnames: st.error("Dosyada 'Fon_Listesi' sayfası yok!"); st.stop()

req_codes, valor_map = [], {}
for row in wb["Fon_Listesi"].iter_rows(min_row=2, values_only=False):
    if row and row[0].value:
        c = normalize_fund_code(row[0].value)
        if c: req_codes.append(c); valor_map[c] = parse_number(row[1].value) if len(row) > 1 and row[1].value is not None else 0.0
req_codes = list(dict.fromkeys(req_codes))

today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

# ============================================================
# ANA VERİ TOPLAMA BLOĞU (V9)
# ============================================================
with st.spinner("Fon verileri doğrudan API'lerden indiriliyor..."):
    calc_funds, fail_codes = [], []
    bar = st.progress(0)
    st_text = st.empty()
    tot = len(req_codes)

    for i, code in enumerate(req_codes):
        st_text.text(f"İndiriliyor ve Analiz Ediliyor: {code} ({i+1}/{tot})")
        
        # 1. Metadata TEFAS'tan HTML kazıyarak alınır (Engellenmez)
        meta = fetch_fund_metadata(code)
        
        # 2. Fiyatlar TEFAS Grafik API'sinden çekilir
        series, src = fetch_tefas_prices(code, start_date, today), "TEFAS"
        
        # 3. Eğer TEFAS başarısız olursa İş Yatırım denenir
        if series is None:
            series, src = fetch_isyatirim_series(code, start_date, today), "İş Yatırım" if fetch_isyatirim_series(code, start_date, today) is not None else "Bulunamadı"

        metrics = compute_fund_metrics(series, meta)
        if metrics:
            metrics.update({"code": code, "source": src, "valor": valor_map.get(code, 0.0)})
            a_ok = metrics["aum"] >= min_aum if metrics["aum"] > 0 else True
            i_ok = metrics["investors"] >= min_investors if metrics["investors"] > 0 else True
            metrics["filtered_out"] = not (a_ok and i_ok)
            calc_funds.append(metrics)
        else: fail_codes.append(code)
        bar.progress((i + 1) / tot)
        
    st_text.empty()
    bar.empty()

st.success(f"✅ {len(calc_funds)} fonun verisi başarıyla indirildi. Skorlamaya geçiliyor...")
if fail_codes: st.warning(f"⚠️ Veri bulunamayan fonlar: {', '.join(fail_codes)}")

for d in PERIODS.values(): calculate_period_scores(calc_funds, d, daily_rf, use_category_scoring)
if any(f["code"] == "KUT" for f in calc_funds): calculate_kut_benchmark_metrics(calc_funds, fetch_kut_benchmark(start_date, today))
calculate_trends(calc_funds)
calculate_final_scores(calc_funds)
calc_funds.sort(key=lambda x: x.get("final_score") if x.get("final_score") is not None else -1, reverse=True)

# ============================================================
# EXCEL ÇIKTISI
# ============================================================
if "Vade_Analizi" in wb.sheetnames: del wb["Vade_Analizi"]
ws_out = wb.create_sheet("Vade_Analizi", 0)

headers = [
    "Fon Kodu", "Kategori", "AUM (TL)", "Yatırımcı",
    "1H Skor", "1H Karar", "1H Küm %", "1H MDD %", "1H Valor Ceza",
    "1A Skor", "1A Karar", "1A Küm %", "1A MDD %", "1A Valor Ceza",
    "3A Skor", "3A Karar", "3A Küm %", "3A MDD %", "3A Valor Ceza",
    "1Y Skor", "1Y Karar", "1Y Küm %", "1Y MDD %", "1Y Valor Ceza",
    "1Y Trend", "3A Trend", "Güven Seviyesi",
    "Nihai Skor", "Nihai Karar",
    "Benchmark 1H", "KUT-Bench 1H", "Benchmark 1A", "KUT-Bench 1A",
    "Benchmark 3A", "KUT-Bench 3A", "Benchmark 1Y", "KUT-Bench 1Y",
    "Benchmark Ort. Skor", "Kaynak"
]
ws_out.append(headers)

for cell in ws_out[1]: cell.fill, cell.font, cell.alignment = PatternFill(start_color=COLOR_NAVY, fill_type="solid"), Font(color=COLOR_WHITE, bold=True), Alignment(horizontal="center", vertical="center", wrap_text=True)

def fmt(v): return v if v is not None else "-"

for f in calc_funds:
    bs = [f.get(f"benchmark_score_{d}") for d in PERIODS.values() if f.get(f"benchmark_score_{d}") is not None]
    b_avg = round(sum(bs) / len(bs), 1) if bs else "-"
    ws_out.append([
        f["code"], f.get("category", "-"), fmt(f.get("aum")), fmt(f.get("investors")),
        fmt(f.get("score_5")), f.get("karar_5", "-"), fmt(f.get("cum_5")), fmt(f.get("mdd_5")), fmt(f.get("valor_penalty_5")),
        fmt(f.get("score_21")), f.get("karar_21", "-"), fmt(f.get("cum_21")), fmt(f.get("mdd_21")), fmt(f.get("valor_penalty_21")),
        fmt(f.get("score_63")), f.get("karar_63", "-"), fmt(f.get("cum_63")), fmt(f.get("mdd_63")), fmt(f.get("valor_penalty_63")),
        fmt(f.get("score_252")), f.get("karar_252", "-"), fmt(f.get("cum_252")), fmt(f.get("mdd_252")), fmt(f.get("valor_penalty_252")),
        f.get("trend_1y", "-"), f.get("trend_3a", "-"), f.get("confidence", "-"),
        fmt(f.get("final_score")), f.get("final_decision", "-"),
        fmt(f.get("benchmark_5")), fmt(f.get("benchmark_diff_5")), fmt(f.get("benchmark_21")), fmt(f.get("benchmark_diff_21")),
        fmt(f.get("benchmark_63")), fmt(f.get("benchmark_diff_63")), fmt(f.get("benchmark_252")), fmt(f.get("benchmark_diff_252")),
        b_avg, f.get("source", "-")
    ])

green, red, yellow = Font(color=COLOR_GREEN, bold=True), Font(color=COLOR_RED, bold=True), Font(color=COLOR_YELLOW, bold=True)
for r in range(2, ws_out.max_row + 1):
    for c in [6, 11, 16, 21, 29]:
        v = str(ws_out.cell(r, c).value)
        if any(x in v for x in ["GÜÇLÜ AL", "AL", "LİSTE"]): ws_out.cell(r, c).font = green
        elif "NÖTR" in v: ws_out.cell(r, c).font = yellow
        elif any(x in v for x in ["SAT", "ACİL", "FİLTRE"]): ws_out.cell(r, c).font = red
    for c in [25, 26]:
        v = str(ws_out.cell(r, c).value)
        if "YUKARI" in v: ws_out.cell(r, c).font = green
        elif "AŞAĞI" in v: ws_out.cell(r, c).font = red
        elif "YATAY" in v: ws_out.cell(r, c).font = yellow
    cf = ws_out.cell(r, 27)
    if cf.value in ["ÇOK YÜKSEK", "YÜKSEK"]: cf.font = green
    elif cf.value == "ORTA": cf.font = yellow
    else: cf.font = red
    cfs = ws_out.cell(r, 28)
    if isinstance(cfs.value, (int, float)): cfs.fill = PatternFill(start_color=COLOR_LIGHT_GREEN if cfs.value >= 70 else (COLOR_LIGHT_YELLOW if cfs.value >= 45 else COLOR_LIGHT_RED), fill_type="solid")

for r in range(2, ws_out.max_row + 1):
    for c in [7, 8, 9, 12, 13, 14, 17, 18, 19, 22, 23, 24, 30, 31, 32, 33, 34, 35, 36, 37]:
        if isinstance(ws_out.cell(r, c).value, (int, float)): ws_out.cell(r, c).number_format = '0.00"%"'

for c in ws_out.columns: ws_out.column_dimensions[get_column_letter(c[0].column)].width = min(max(max([len(str(cl.value)) for cl in c]) + 2, 11), 26)
ws_out.freeze_panes, ws_out.auto_filter.ref = "B2", ws_out.dimensions

out = io.BytesIO()
wb.save(out)
out.seek(0)

# ============================================================
# SONUÇ EKRANI
# ============================================================
st.download_button("📥 V9 Excel Çıktısını İndir", data=out, file_name="fon_vade_analizi_V9.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.subheader("📊 Fon Sıralaması Önizleme")
df_p = pd.DataFrame([{"Fon": f["code"], "Kategori": f.get("category", "-"), "1H": fmt(f.get("score_5")), "1A": fmt(f.get("score_21")), "3A": fmt(f.get("score_63")), "1Y": fmt(f.get("score_252")), "Nihai Skor": fmt(f.get("final_score")), "Nihai Karar": f.get("final_decision", "-"), "Güven": f.get("confidence", "-"), "AUM": fmt(f.get("aum"))} for f in calc_funds])
st.dataframe(df_p, use_container_width=True, hide_index=True)
