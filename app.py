import datetime as dt
import io
import math
import urllib3
from typing import Optional, List, Tuple

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# SAYFA AYARLARI & SABİTLER
# ============================================================
st.set_page_config(page_title="Multi-Vade Fon Analizi V16", page_icon="📈", layout="wide")
st.title("📈 Multi-Vade Fon Analizi V16")
st.caption("Saf TEFAS API Motoru (Sıfır Dış Bağımlılık) + Gelişmiş Z-Skor + Tanh Benchmark")

LOOKBACK_CALENDAR_DAYS = 400
HTTP_TIMEOUT = 15

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
        try: return None if pd.isna(value) else float(value)
        except: return float(value)
    text = str(value).strip().replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "")
    if not text: return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text: text = text.replace(",", ".")
    try: return float(text)
    except: return None

def normalize_fund_code(value) -> str:
    if value is None: return ""
    return str(value).strip().upper()

def safe_float(value, default=0.0):
    try:
        val = float(value)
        return default if math.isnan(val) or math.isinf(val) else val
    except: return default

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
# SAF TEFAS VE İŞ YATIRIM API MOTORU (V16)
# ============================================================
@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_fund_data_safely(code: str, start_date: dt.date, end_date: dt.date) -> Tuple[Optional[pd.DataFrame], dict, str]:
    meta = {"aum": 0.0, "investors": 0.0, "title": "", "category": ""}
    
    # 1. YÖNTEM: Doğrudan TEFAS Grafik / Geçmiş Veri API Uç Noktası
    url_tefas = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.tefas.gov.tr/FonAnaliz.aspx"
    }
    
    for fontip in ["", "YAT", "EMK", "BYF"]:
        data = {
            "fontip": fontip, "sfontur": "", "fonkod": code.upper(), "fongrup": "",
            "baslangic": start_date.strftime("%d.%m.%Y"), "bitis": end_date.strftime("%d.%m.%Y")
        }
        try:
            res = requests.post(url_tefas, data=data, headers=headers, verify=False, timeout=HTTP_TIMEOUT)
            if res.status_code == 200:
                json_data = res.json()
                data_list = json_data.get("data", [])
                if data_list:
                    rows = []
                    for item in data_list:
                        t = item.get("TARIH")
                        p = item.get("FIYAT")
                        if t and p:
                            rows.append({
                                "date": pd.Timestamp(dt.datetime.fromtimestamp(t / 1000.0).date()),
                                "price": float(p)
                            })
                    if rows:
                        df = pd.DataFrame(rows).dropna()
                        df = df[df["price"] > 0]
                        if len(df) >= 2:
                            df_sorted = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(MAX_DAYS + 1).reset_index(drop=True)
                            
                            # Meta Bilgiler (Fon Adı, AUM, Yatırımcı)
                            meta_url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={code.upper()}"
                            m_res = requests.get(meta_url, headers={"User-Agent": "Mozilla/5.0"}, verify=False, timeout=5)
                            if m_res.status_code == 200:
                                html = m_res.text
                                m_title = re.search(r'id="MainContent_FormViewMainIndicators_LabelFund"?[^>]*>([^<]+)</span>', html)
                                if m_title: 
                                    meta["title"] = m_title.group(1).strip()
                                    meta["category"] = infer_category(meta["title"])
                                m_aum = re.search(r'id="MainContent_FormViewMainIndicators_LabelPortfolioValue"?[^>]*>([^<]+)</span>', html)
                                if m_aum: meta["aum"] = parse_number(m_aum.group(1)) or 0.0
                                m_inv = re.search(r'id="MainContent_FormViewMainIndicators_LabelInvestorCount"?[^>]*>([^<]+)</span>', html)
                                if m_inv: meta["investors"] = parse_number(m_inv.group(1)) or 0.0

                            return df_sorted, meta, "TEFAS API"
        except:
            continue

    # 2. YÖNTEM (Yedek): İş Yatırım API
    try:
        session = requests.Session()
        session.verify = False
        session.headers.update({"User-Agent": "Mozilla/5.0"})
        session.get("https://www.isyatirim.com.tr/tr-tr/analiz/fon/Sayfalar/Tarihsel-Fiyat-Bilgileri.aspx", timeout=5)
        
        api_url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
        params = {"fonKod": code.upper(), "baslangic": start_date.strftime("%d-%m-%Y"), "bitis": end_date.strftime("%d-%m-%Y")}
        
        res = session.get(api_url, params=params, headers={"X-Requested-With": "XMLHttpRequest"}, timeout=HTTP_TIMEOUT)
        if res.status_code == 200:
            values = res.json().get("value")
            if values:
                df = pd.DataFrame(values)
                if "Tarih" in df.columns and "Fiyat" in df.columns:
                    df["date"] = pd.to_datetime(df.get("Tarih"), dayfirst=True, errors="coerce")
                    df["price"] = df.get("Fiyat").apply(parse_number)
                    df = df.dropna(subset=["date", "price"])
                    df = df[df["price"] > 0]
                    if len(df) >= 2:
                        df_sorted = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(MAX_DAYS + 1).reset_index(drop=True)
                        meta["title"] = code.upper()
                        meta["category"] = infer_category(code.upper())
                        return df_sorted[["date", "price"]], meta, "İş Yatırım"
    except:
        pass

    return None, {}, "Tüm resmi API kanalları yanıt vermedi."

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
    clean = [float(r) for r in returns if r is not None and math.isfinite(float(r))]
    if not clean: return 0.0
    excess = [r - daily_rf for r in clean]
    mean_ex = sum(excess) / len(excess)
    downside_sq = [min(0.0, x) ** 2 for x in excess]
    down_dev = math.sqrt(sum(downside_sq) / len(excess))
    if down_dev <= 1e-12: return max_sortino if mean_ex > 0 else 0.0
    return max(-max_sortino, min(max_sortino, mean_ex / down_dev))

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

        s_ret, s_prices = returns[-days:], prices[-(days + 1):]
        m_ret = sum(s_ret) / len(s_ret)
        vol = math.sqrt(max(0.0, sum((r - m_ret) ** 2 for r in s_ret) / len(s_ret)))
        item[f"mean_{days}"], item[f"vol_{days}"] = m_ret, vol
        item[f"shp_{days}"] = ((m_ret - daily_rf) / vol * math.sqrt(252)) if vol > 1e-12 else 0.0
        item[f"cum_{days}"] = (s_prices[-1] / s_prices[0] - 1.0) * 100.0 if s_prices[0] > 0 else None
        item[f"mdd_{days}"] = compute_max_drawdown(s_prices)
        item[f"sortino_{days}"] = compute_sortino(s_ret, daily_rf=daily_rf)

    valid_idx = [i for i, f in enumerate(funds) if not f.get("filtered_out") and all(f.get(f"{m}_{days}") is not None for m in ["cum", "shp", "sortino", "mdd"])]
    if not valid_idx:
        for f in funds: f[f"score_{days}"], f[f"karar_{days}"] = None, "FİLTRE DIŞI" if f.get("filtered_out") else "YETERSİZ VERİ"
        return

    cats = {}
    if use_category:
        for idx in valid_idx: cats.setdefault(funds[idx].get("category") or "Diğer", []).append(idx)
    else: cats = {"ALL": valid_idx}

    def get_stats(indices, field):
        vals = [float(funds[i][f"{field}_{days}"]) for i in indices if funds[i].get(f"{field}_{days}") is not None and math.isfinite(float(funds[i][f"{field}_{days}"]))]
        if not vals: return 0.0, 0.0
        m = sum(vals) / len(vals)
        return m, math.sqrt(max(0.0, sum((x - m) ** 2 for x in vals) / len(vals)))

    def apply_scoring(eval_idx, target_idx):
        if len(eval_idx) < 2:
            for i in target_idx: funds[i][f"stat_score_{days}"] = None
            return
        m_cum, m_shp, m_srt, m_mdd = get_stats(eval_idx, "cum"), get_stats(eval_idx, "shp"), get_stats(eval_idx, "sortino"), get_stats(eval_idx, "mdd")
        def z(val, m, std): return max(-Z_LIMIT, min(Z_LIMIT, (float(val) - m) / std)) if val is not None and std > 1e-12 else 0.0
        
        for i in target_idx:
            f = funds[i]
            w_z = (W_CUM * z(f.get(f"cum_{days}"), *m_cum) + W_SHP * z(f.get(f"shp_{days}"), *m_shp) + W_SRT * z(f.get(f"sortino_{days}"), *m_srt) + W_MDD * z(f.get(f"mdd_{days}"), *m_mdd))
            f[f"stat_score_{days}"] = int(round(max(0.0, min(100.0, 50.0 + 20.0 * w_z))))

    apply_scoring(valid_idx, valid_idx)
    g_scores = {i: funds[i].get(f"stat_score_{days}") for i in valid_idx}

    for cat, indices in cats.items():
        if len(indices) >= 10:
            apply_scoring(indices, indices)
            for i in indices: funds[i][f"score_{days}"] = funds[i].get(f"stat_score_{days}")
        elif len(indices) >= 5:
            apply_scoring(indices, indices)
            for i in indices:
                c_s, g_s = funds[i].get(f"stat_score_{days}"), g_scores.get(i)
                if c_s is not None and g_s is not None: funds[i][f"score_{days}"] = int(round(0.60 * c_s + 0.40 * g_s))
                else: funds[i][f"score_{days}"] = c_s if c_s is not None else g_s
        else:
            for i in indices: funds[i][f"score_{days}"] = g_scores.get(i)

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
def fetch_yahoo_series(symbol: str, start_date: dt.date, end_date: dt.date) -> Tuple[Optional[pd.DataFrame], str]:
    try:
        session = requests.Session()
        session.verify = False
        p1, p2 = int(dt.datetime.combine(start_date, dt.time.min).timestamp()), int(dt.datetime.combine(end_date + dt.timedelta(days=1), dt.time.min).timestamp())
        res = session.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}", params={"period1": p1, "period2": p2, "interval": "1d", "events": "history", "includeAdjustedClose": "true"}, headers={"User-Agent": "Mozilla/5.0"}, timeout=HTTP_TIMEOUT)
        res.raise_for_status()
        result = res.json().get("chart", {}).get("result", [{}])[0]
        t_stamps, closes = result.get("timestamp"), result.get("indicators", {}).get("quote", [{}])[0].get("close")
        if not t_stamps or not closes: return None, "Yahoo verisi boş."
        rows = [{"date": pd.Timestamp(dt.datetime.fromtimestamp(t).date()), "price": float(c)} for t, c in zip(t_stamps, closes) if c is not None and float(c) > 0]
        if len(rows) < 2: return None, "Yeterli fiyat yok."
        return pd.DataFrame(rows).sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True), "OK"
    except Exception as e: return None, str(e)

def fetch_kut_benchmark(start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    g, _ = fetch_yahoo_series("GC=F", start_date, end_date)
    s, _ = fetch_yahoo_series("SI=F", start_date, end_date)
    u, _ = fetch_yahoo_series("USDTRY=X", start_date, end_date)
    if g is None or s is None or u is None: return None
    df = g.rename(columns={"price": "g"}).merge(s.rename(columns={"price": "s"}), on="date").merge(u.rename(columns={"price": "u"}), on="date")
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
today, start_date = dt.date.today(), dt.date.today() - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

# ============================================================
# ANA VERİ TOPLAMA BLOĞU (V16 - Saf İstihbarat)
# ============================================================
with st.spinner("Saf TEFAS & İş Yatırım API Motoru Çalışıyor..."):
    calc_funds, fail_codes, error_logs = [], [], []
    bar, st_text = st.progress(0), st.empty()
    tot = len(req_codes)

    for i, code in enumerate(req_codes):
        st_text.text(f"İndiriliyor ve Analiz Ediliyor: {code} ({i+1}/{tot})")
        
        series, meta, src = fetch_fund_data_safely(code, start_date, today)

        if series is not None:
            metrics = compute_fund_metrics(series, meta)
            metrics.update({"code": code, "source": src, "valor": valor_map.get(code, 0.0)})
            a_ok = metrics["aum"] >= min_aum if metrics["aum"] > 0 else True
            i_ok = metrics["investors"] >= min_investors if metrics["investors"] > 0 else True
            metrics["filtered_out"] = not (a_ok and i_ok)
            calc_funds.append(metrics)
        else: 
            fail_codes.append(code)
            error_logs.append(f"[{code}] Tüm resmi API kanalları boş döndü.")
            
        bar.progress((i + 1) / tot)
        
    st_text.empty(); bar.empty()

if error_logs:
    with st.expander("🛠️ Hata Logları (Fonlar Neden İndirilemedi?)", expanded=True):
        st.write("Aşağıdaki fonlar için resmi API kanallarından veri alınamadı.")
        for log in error_logs: st.code(log)

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
        round(sum(bs)/len(bs), 1) if bs else "-", f.get("source", "-")
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
st.download_button("📥 V16 Excel Çıktısını İndir", data=out, file_name="fon_vade_analizi_V16.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.subheader("📊 Fon Sıralaması Önizleme")
df_p = pd.DataFrame([{"Fon": f["code"], "Kategori": f.get("category", "-"), "1H": fmt(f.get("score_5")), "1A": fmt(f.get("score_21")), "3A": fmt(f.get("score_63")), "1Y": fmt(f.get("score_252")), "Nihai Skor": fmt(f.get("final_score")), "Nihai Karar": fmt(f.get("final_decision")), "Güven": fmt(f.get("confidence")), "Kaynak": fmt(f.get("source"))} for f in calc_funds])
st.dataframe(df_p, use_container_width=True, hide_index=True)
