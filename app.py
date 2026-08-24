import concurrent.futures
import datetime as dt
import io
import json
import math
import os
import re
import time
from collections import defaultdict
from dataclasses import asdict, dataclass
from typing import Any, Dict, List, Optional

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ============================================================
# OPENPYXL 'extLst' HATASI İÇİN ÇALIŞMA ZAMANI YAMASI
# ============================================================
original_init = PatternFill.__init__
def new_init(self, *args, **kwargs):
    if 'extLst' in kwargs:
        del kwargs['extLst']
    original_init(self, *args, **kwargs)
PatternFill.__init__ = new_init

# ============================================================
# tlgssk - SÜRÜM V13.15.2 (HATA GİDERME & KAPSAMLI VERİ ÇEKİCİ)
# ============================================================

st.set_page_config(
    page_title="tlgssk Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 tlgssk Hibrit Fon Analizi")
st.caption(
    "TEFAS + İş Yatırım | Canlı Sentiment + Çoklu Girdi Desteği | V13.15.2"
)

# ============================================================
# AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")
DEFAULT_FUND_KIND = "YAT"

LOOKBACK_CALENDAR_DAYS = 60
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 3

HTTP_TIMEOUT = 15
MAX_WORKERS = 6

REQUEST_MAX_RETRIES = 3
REQUEST_BACKOFF_FACTOR = 1.0

MIN_REFERENCE_SAMPLE = 5
OVERHEAT_Z_THRESHOLD = 2.0
OVERHEAT_PENALTY = 6.0

DEFAULT_MOMENTUM_WEIGHTS = {"return": 0.30, "sharpe": 0.25, "cumulative": 0.25, "drawdown": 0.20}
SECURITY_WEIGHTS = {"aum": 0.30, "investor": 0.25, "concentration": 0.25, "liquidity": 0.20}
SECURITY_SCALE = {"aum": 20.0, "investor": 20.0, "aum_flow": 8.0, "investor_change": 6.0, "concentration": 20.0}

DEFAULT_HYBRID_MOMENTUM_WEIGHT = 0.50
DEFAULT_HYBRID_SECURITY_WEIGHT = 0.35
DEFAULT_HYBRID_SENTIMENT_WEIGHT = 0.15

Z_LIMIT = 2.5
STRONG_BUY = 75
WATCH_LIST = 50
CORRECTION = 35

MAX_CONCENTRATION_PENALTY = 20.0
BIST30_BONUS = 5.0
HIGH_LIQUIDITY_BONUS = 5.0
LOW_LIQUIDITY_PENALTY = 3.0
EMA_DECAY = 0.65

GITHUB_FALLBACK_URL = "https://github.com/tlgssk/kgdm3-fon-analiz/raw/refs/heads/main/Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"

# ============================================================
# HTTP OTURUMU
# ============================================================
@dataclass
class SourceStatus:
    source: str; attempted: bool = False; ok: bool = False; status_code: Optional[int] = None
    error_type: str = ""; message: str = ""; elapsed_ms: Optional[int] = None; retry_count: int = 0

def new_status(source: str) -> SourceStatus: return SourceStatus(source=source)

def build_http_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=REQUEST_MAX_RETRIES, connect=REQUEST_MAX_RETRIES, read=REQUEST_MAX_RETRIES,
        status=REQUEST_MAX_RETRIES, backoff_factor=REQUEST_BACKOFF_FACTOR,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"}),
        respect_retry_after_header=True
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
        "X-Requested-With": "XMLHttpRequest"
    })
    return session

HTTP = build_http_session()

def request_with_status(source: str, method: str, url: str, *, params=None, data=None, headers=None, timeout=HTTP_TIMEOUT):
    status = new_status(source)
    status.attempted = True
    started = time.perf_counter()
    try:
        response = HTTP.request(method=method, url=url, params=params, data=data, headers=headers, timeout=timeout)
        status.status_code = response.status_code
        if response.status_code == 200:
            status.ok, status.message = True, "OK"
        else:
            status.error_type, status.message = f"HTTP_{response.status_code}", f"Hata {response.status_code}"
        return response, status
    except Exception as exc:
        status.error_type, status.message = "ERROR", str(exc)[:200]
    finally:
        status.elapsed_ms = int((time.perf_counter() - started) * 1000)
    return None, status

# ============================================================
# RENKLER & SIDEBAR
# ============================================================

COLOR_NAVY, COLOR_GREEN, COLOR_RED, COLOR_YELLOW, COLOR_WHITE = "1F4E79", "008000", "FF0000", "B8860B", "FFFFFF"
COLOR_LIGHT_GREEN, COLOR_LIGHT_YELLOW, COLOR_LIGHT_RED = "E2F0D9", "FFF2CC", "FCE4D6"

st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

env_api_key = os.environ.get("GEMINI_API_KEY", "")
try:
    if not env_api_key and "GEMINI_API_KEY" in st.secrets:
        env_api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

api_key_input = st.sidebar.text_input(
    "🔑 Gemini API Key (Canlı Sentiment)",
    value=env_api_key,
    type="password",
    help="Google AI Studio API anahtarı. Boş bırakılırsa kural tabanlı duyarlılık çalışır.",
)

with st.sidebar.expander("⚖️ Skor Ağırlıkları (V13.15)"):
    w_return = st.slider("Getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_cumulative = st.slider("Kümülatif ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_drawdown = st.slider("Drawdown ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05)

    total_m = w_return + w_sharpe + w_cumulative + w_drawdown
    total_m = 1.0 if total_m <= 0 else total_m

    MOMENTUM_WEIGHTS = {
        "return": w_return / total_m, "sharpe": w_sharpe / total_m,
        "cumulative": w_cumulative / total_m, "drawdown": w_drawdown / total_m,
    }

    st.markdown("---")
    st.markdown("**Hibrit Karar Dağılımı**")
    w_hybrid_mom = st.slider("Momentum Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_MOMENTUM_WEIGHT, 0.05)
    w_hybrid_sec = st.slider("Güvenlik Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SECURITY_WEIGHT, 0.05)
    w_hybrid_sent = st.slider("Sentiment Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SENTIMENT_WEIGHT, 0.05)

    tot_h = w_hybrid_mom + w_hybrid_sec + w_hybrid_sent
    if tot_h <= 0: tot_h = 1.0

    HYBRID_MOMENTUM_WEIGHT = w_hybrid_mom / tot_h
    HYBRID_SECURITY_WEIGHT = w_hybrid_sec / tot_h
    HYBRID_SENTIMENT_WEIGHT = w_hybrid_sent / tot_h

SHOW_DIAGNOSTICS = st.sidebar.checkbox("Kaynak tanılama bilgisini göster", value=True)

# ============================================================
# PARSER VE DÖNÜŞTÜRÜCÜ YARDIMCILARI
# ============================================================

def clamp(value, low, high): return max(low, min(high, value))

def safe_float(value, default=0.0) -> float:
    try:
        if value is None: return default
        n = float(value)
        return default if pd.isna(n) else n
    except: return default

def optional_float(value) -> Optional[float]:
    try:
        if value is None or (isinstance(value, str) and not value.strip()):
            return None
        n = float(value)
        return None if pd.isna(n) else n
    except (TypeError, ValueError):
        return None

def normalize_date_key(value) -> Optional[str]:
    try:
        ts = pd.to_datetime(value, errors="coerce")
        if pd.isna(ts): return None
        return ts.strftime("%Y-%m-%d")
    except Exception: return None

def display_date(date_key) -> str:
    try: return pd.to_datetime(date_key).strftime("%d.%m.%Y")
    except Exception: return str(date_key)

def parse_number(value):
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)): return None if pd.isna(value) else float(value)
    text = str(value).replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
    if not text: return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text: text = text.replace(",", ".")
    elif "." in text and re.match(r"^-?\d{1,3}(\.\d{3})+$", text): text = text.replace(".", "")
    try: return float(text)
    except: return None

def normalize_fund_code(value):
    code = str(value).strip().upper()
    return code[:-2] if code.endswith(".0") else code

def format_percent(value):
    n = parse_number(value)
    if n is None: return "-"
    return f"+%{n:.2f}" if n > 0 else (f"-%{abs(n):.2f}" if n < 0 else "%0.00")

def calculate_compounded_return(returns):
    clean = [parse_number(v) for v in returns if v is not None]
    if not clean: return 0.0
    growth = 1.0
    for r in clean: growth *= 1.0 + r / 100.0
    return (growth - 1.0) * 100.0

def calculate_max_drawdown(prices):
    if not prices or len(prices) < 2: return 0.0
    peak, max_dd = safe_float(prices[0]), 0.0
    for price in prices:
        p = safe_float(price)
        if p <= 0: continue
        if p > peak: peak = p
        if peak > 0:
            dd = (p / peak - 1.0) * 100.0
            if dd < max_dd: max_dd = dd
    return max_dd

def zscore(values):
    clean = [optional_float(v) for v in values]
    valid = [v for v in clean if v is not None]
    if len(valid) < 2: return [0.0 if v is not None else None for v in clean]
    mean_v = sum(valid) / len(valid)
    std = (sum((x - mean_v) ** 2 for x in valid) / len(valid)) ** 0.5
    if std <= 1e-12: return [0.0 if v is not None else None for v in clean]
    out = []
    for v in clean:
        if v is not None: out.append(clamp((v - mean_v) / std, -Z_LIMIT, Z_LIMIT))
        else: out.append(None)
    return out

def population_mean_std(values):
    valid = [optional_float(v) for v in values]
    valid = [v for v in valid if v is not None]
    if len(valid) < 2: return 0.0, 0.0
    mean_v = sum(valid) / len(valid)
    return mean_v, (sum((v - mean_v) ** 2 for v in valid) / len(valid)) ** 0.5

def zscore_against_population(value, mean_v, std_v):
    if value is None or std_v <= 1e-12: return 0.0
    return clamp((value - mean_v) / std_v, -Z_LIMIT, Z_LIMIT)

# ============================================================
# CANLI GEMINI DUYARLILIK MOTORU
# ============================================================

@st.cache_data(ttl=60 * 60 * 4, show_spinner=False)
def fetch_batch_market_sentiment(areas: list, api_key: str) -> dict:
    result_map = {}
    api_key_clean = api_key.strip() if api_key else ""
    
    if not api_key_clean:
        for area in areas:
            a_u = area.upper()
            if "YABANCI TEKNOLOJİ" in a_u or "YABANCI" in a_u:
                result_map[area] = {"score": 38, "label": "Negatif (Kâr Satışı)", "ai_active": False, "ai_reason": "API Anahtarı Yok"}
            elif "ALTIN" in a_u or "GÜMÜŞ" in a_u or "KIYMETLİ" in a_u:
                result_map[area] = {"score": 82, "label": "Güçlü Pozitif (Faiz İndirimi)", "ai_active": False, "ai_reason": "API Anahtarı Yok"}
            elif "PARA PİYASASI" in a_u or "BORÇLANMA" in a_u:
                result_map[area] = {"score": 65, "label": "Pozitif (Sabit Getiri)", "ai_active": False, "ai_reason": "API Anahtarı Yok"}
            elif "HİSSE" in a_u or "BIST" in a_u:
                result_map[area] = {"score": 54, "label": "Dengeli / Pozitif Beklenti", "ai_active": False, "ai_reason": "API Anahtarı Yok"}
            else:
                result_map[area] = {"score": 50, "label": "Nötr / Kural Tabanlı", "ai_active": False, "ai_reason": "API Anahtarı Yok"}
        return result_map

    areas_text = "\n".join([f"- {a}" for a in areas])
    prompt = f"""Sen kıdemli bir fon yöneticisi ve makroekonomik analistsin.
Aşağıdaki yatırım alanları için 0-100 arası duyarlılık puanı ve max 6 kelimelik gerekçe üret:
{areas_text}
SADECE JSON döndür: {{"Alan Adı": {{"score": 75, "label": "Kısa gerekçe"}}}}"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.7-flash:generateContent?key={api_key_clean}"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json"}
    }

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=20)
        if response.status_code == 200:
            raw_text = response.json().get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "{}")
            parsed = json.loads(raw_text.strip("```json\n").strip("```").strip())
            for area in areas:
                if area in parsed:
                    result_map[area] = {
                        "score": int(clamp(safe_float(parsed[area].get("score", 50)), 0.0, 100.0)),
                        "label": str(parsed[area].get("label", "Nötr")),
                        "ai_active": True,
                        "ai_reason": "Canlı API"
                    }
                else:
                    result_map[area] = {"score": 50, "label": "Nötr", "ai_active": True, "ai_reason": "Varsayılan"}
            return result_map
    except Exception:
        pass

    for area in areas:
        result_map[area] = {"score": 50, "label": "Nötr", "ai_active": False, "ai_reason": "API Hatası"}
    return result_map

# ============================================================
# TEFAS DOĞRUDAN VERİ ÇEKİCİLER
# ============================================================

def fetch_tefas_direct_api(fund_code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(fund_code)
    status = new_status("TEFAS Direct API")
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {
        "Origin": "https://www.tefas.gov.tr",
        "Referer": "https://www.tefas.gov.tr/TarihselVeriler.aspx",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    }
    
    kinds_to_try = [fund_kind] if fund_kind in FUND_KINDS else ["YAT", "EMK", "BYF"]
    for kind in kinds_to_try:
        payload = {
            "fontip": kind,
            "fonkod": code,
            "bastarih": start.strftime("%d.%m.%Y"),
            "bittarih": end.strftime("%d.%m.%Y")
        }
        res, stat = request_with_status("TEFAS Direct API", "POST", url, data=payload, headers=headers)
        if res and stat.ok:
            try:
                raw = res.json().get("data", [])
                if raw:
                    df = pd.DataFrame(raw)
                    df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
                    df["price"] = df["FIYAT"].apply(parse_number)
                    df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number) if "PORTFOYBUYUKLUK" in df.columns else None
                    df["investors"] = df["KISISAYISI"].apply(parse_number) if "KISISAYISI" in df.columns else None
                    df = df.dropna(subset=["date", "price"])
                    df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                    if len(df) >= 2:
                        return df, stat
            except Exception:
                pass
    return None, status

def fetch_isyatirim_series(fund_code: str):
    code = normalize_fund_code(fund_code)
    status = new_status("İş Yatırım")
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    response, status = request_with_status("İş Yatırım", "GET", url, params=params)
    if response and status.ok:
        try:
            df = pd.DataFrame(response.json().get("value", []))
            if not df.empty and "Tarih" in df.columns:
                df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
                df["price"] = df["Fiyat"].apply(parse_number)
                df["aum"], df["investors"] = None, None
                df = df.dropna(subset=["date", "price"])
                df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                if len(df) >= 2:
                    return df[["date", "price", "aum", "investors"]], status
        except Exception:
            pass
    return None, status

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
        crawler = Crawler(timeout=30, max_retry=2)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is None or df.empty: return pd.DataFrame()
        df.rename(columns={"fund_code": "code", "fund_name": "title", "investor_count": "investors", "portfolio_size": "aum", "fund_type": "kind"}, inplace=True)
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["price"] = df["price"].apply(parse_number)
        df["aum"] = df["aum"].apply(parse_number) if "aum" in df.columns else None
        df["investors"] = df["investors"].apply(parse_number) if "investors" in df.columns else None
        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df = df.dropna(subset=["date", "code", "price"])
        return df[df["price"] > 0].sort_values(["code", "date"]).drop_duplicates(subset=["code", "date"], keep="last").reset_index(drop=True)
    except Exception:
        return pd.DataFrame()

def build_fund_meta_map(universe: pd.DataFrame):
    meta = {}
    if universe is not None and not universe.empty:
        latest = universe.sort_values("date").drop_duplicates(subset=["code"], keep="last")
        for _, row in latest.iterrows():
            code = str(row.get("code", "")).strip().upper()
            if code: meta[code] = {"kind": str(row.get("kind", DEFAULT_FUND_KIND)), "title": str(row.get("title", ""))}
    return meta

def build_universe_reference(universe: pd.DataFrame, window: int):
    ref = {k: {"mean_return": [], "sharpe": [], "cumulative": [], "max_dd_inv": [], "aum": [], "investors": []} for k in FUND_KINDS}
    if universe is None or universe.empty or window < 2: return ref
    for code, group in universe.groupby("code"):
        group = group.sort_values("date")
        kind = str(group["kind"].iloc[-1]).strip().upper()
        if kind not in FUND_KINDS: continue
        prices = group["price"].astype(float).tolist()
        if len(prices) < window + 1: continue
        w_prices = prices[-(window + 1):]
        rets = [0.0 if p0 <= 0 else (p1 / p0 - 1.0) * 100.0 for p0, p1 in zip(w_prices[:-1], w_prices[1:])]
        mean_r = sum(rets) / len(rets)
        vol = (sum((r - mean_r) ** 2 for r in rets) / len(rets)) ** 0.5
        ref[kind]["mean_return"].append(mean_r)
        ref[kind]["sharpe"].append(mean_r / vol if vol > 1e-12 else 0.0)
        ref[kind]["cumulative"].append((w_prices[-1] / w_prices[0] - 1.0) * 100.0)
        ref[kind]["max_dd_inv"].append(calculate_max_drawdown(w_prices))
    return ref

def reference_sample_size(ref, kind): return len(ref.get(kind, {}).get("mean_return", []))

def fetch_fund_structural_data(fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {"top_asset_weight": None, "asset_class_hhi": None, "is_bist30": False, "emergency_cash_ratio": None, "cash_ratio_known": False, "structural_fetch_ok": False, "structural_source": "YOK", "investment_area": "-"}
    
    t_upper = (fund_title or "").upper()
    if "PARA PİYASASI" in t_upper or "PPF" in t_upper: structural["investment_area"] = "Para Piyasası"
    elif "ALTIN" in t_upper or "GÜMÜŞ" in t_upper or "KIYMETLİ" in t_upper: structural["investment_area"] = "Kıymetli Maden"
    elif "YABANCI TEKNOLOJİ" in t_upper: structural["investment_area"] = "Hisse Senedi (Yabancı Teknoloji)"
    elif "HİSSE" in t_upper: structural["investment_area"] = "Hisse Senedi"
    elif "BORÇLANMA" in t_upper: structural["investment_area"] = "Borçlanma Araçları"
    elif "DEĞİŞKEN" in t_upper or "KARMA" in t_upper: structural["investment_area"] = "Karma / Değişken"

    if "BIST 30" in t_upper or "BIST30" in t_upper: structural["is_bist30"] = True
    return structural

def get_fund_series(universe: pd.DataFrame, fund_code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(fund_code)
    statuses = []

    # 1. Önce TEFAS Direct API (En taze ve güvenilir)
    df_dir, stat_dir = fetch_tefas_direct_api(code, fund_kind)
    statuses.append(stat_dir)
    if df_dir is not None and len(df_dir) >= 2:
        return df_dir, "TEFAS Direct API", statuses

    # 2. Önbelleklenmiş Evren
    if universe is not None and not universe.empty:
        rows = universe[universe["code"].eq(code)].copy()
        if len(rows) >= 2:
            ok_status = new_status("TEFAS Evren")
            ok_status.attempted = True
            ok_status.ok = True
            ok_status.message = "Önbellek evreninden alındı"
            statuses.append(ok_status)
            return rows.sort_values("date").reset_index(drop=True), "TEFAS", statuses

    # 3. İş Yatırım Fallback
    df_is, stat_is = fetch_isyatirim_series(code)
    statuses.append(stat_is)
    if df_is is not None and len(df_is) >= 2:
        return df_is, "İş Yatırım", statuses

    return None, "YOK", statuses

def compute_fund_metrics(series: pd.DataFrame, fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None):
    if series is None or len(series) < 2: return None

    df = series.copy().sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    df["date_key"] = df["date"].apply(normalize_date_key)
    df = df.dropna(subset=["date_key", "price"]).copy()
    df["price"] = df["price"].apply(optional_float)
    df = df[df["price"].notna() & (df["price"] > 0)].reset_index(drop=True)
    if len(df) < 2: return None

    prices = df["price"].tolist()
    date_keys_all = df["date_key"].tolist()
    aums = [optional_float(v) for v in df["aum"].tolist()] if "aum" in df.columns else [None] * len(df)
    invs = [optional_float(v) for v in df["investors"].tolist()] if "investors" in df.columns else [None] * len(df)

    rets = []
    return_dates = []
    for i in range(1, len(prices)):
        if prices[i - 1] > 0 and prices[i] > 0:
            rets.append((prices[i] / prices[i - 1] - 1.0) * 100.0)
            return_dates.append(date_keys_all[i])

    if not rets: return None

    struct = fetch_fund_structural_data(fund_code, fund_kind, fund_title)
    aum_last = next((v for v in reversed(aums) if v is not None and v > 0), None)
    inv_last = next((v for v in reversed(invs) if v is not None and v >= 0), None)

    price_map = dict(zip(date_keys_all, prices))
    return {
        "code": fund_code,
        "dates": return_dates,
        "orig_dates": list(return_dates),
        "orig_returns": list(rets),
        "prices": prices,
        "price_dates": date_keys_all,
        "price_map": price_map,
        "daily_returns": rets,
        "n_days": len(rets),
        "aum": aum_last,
        "investors": int(inv_last) if inv_last is not None else None,
        "max_dd": calculate_max_drawdown(prices),
        "weekly_return": calculate_compounded_return(rets[-5:]),
        "fund_title": fund_title or "-",
        **struct,
    }

def fetch_and_compute_one_fund(code: str, universe: pd.DataFrame, meta_map: dict, valor_dict: dict):
    meta = meta_map.get(code, {})
    series, source, statuses = get_fund_series(universe, code, meta.get("kind"))
    metrics = compute_fund_metrics(series, code, meta.get("kind"), meta.get("title"))
    if metrics is None: return code, None, source
    metrics["valor"], metrics["source"], metrics["kind"] = valor_dict.get(code), source, meta.get("kind", DEFAULT_FUND_KIND)
    metrics["source_statuses"] = [asdict(x) for x in statuses]
    metrics["source_chain"] = " → ".join(x.source for x in statuses if x.attempted)
    return code, metrics, source

def calculate_security_scores(funds: List[dict], reference: dict):
    for f in funds:
        f["security_score"] = 50

def calculate_market_relative_momentum(funds: List[dict], reference, window: int):
    for f in funds:
        k = f.get("kind", DEFAULT_FUND_KIND)
        rets = f.get("orig_returns", [])[-window:]
        prc = f.get("prices", [])[-(len(rets) + 1):] if f.get("prices") else []
        if not rets:
            f["market_momentum"] = 50
            continue

        m_r = sum(rets) / len(rets)
        vol = (sum((x - m_r) ** 2 for x in rets) / len(rets)) ** 0.5 if len(rets) > 1 else 0.0
        cum = (prc[-1] / prc[0] - 1.0) * 100.0 if (len(prc) >= 2 and prc[0] > 0) else calculate_compounded_return(rets)
        dd = calculate_max_drawdown(prc) if len(prc) >= 2 else 0.0

        f["_final_mean_return"] = m_r
        f["_final_sharpe"] = m_r / vol if vol > 1e-12 else 0.0
        f["_final_cumulative"] = cum
        f["_final_max_dd"] = dd
        f["volatility"] = vol
        f["reference_scope"] = f"Piyasa ({k})"
        f["market_momentum"] = int(round(clamp(50.0 + (m_r * 15.0), 10.0, 95.0)))

def calculate_trend_scores(funds: List[dict], batch_sentiments: dict) -> int:
    if not funds: return 0

    all_dates = set()
    for f in funds:
        all_dates.update(f.get("orig_dates", []))
        
    master_dates = sorted(list(all_dates))[-TARGET_TRADING_DAYS:]
    if not master_dates: return 0
    
    for f in funds:
        ret_map = dict(zip(f.get("orig_dates", []), f.get("orig_returns", [])))
        f["dates"] = master_dates
        f["daily_returns"] = [ret_map.get(d) for d in master_dates]
        f["n_days"] = len([r for r in f["daily_returns"] if r is not None])

        sec = safe_float(f.get("security_score"), 50.0)
        sent = clamp(safe_float(batch_sentiments.get(f.get("investment_area", "-"), {}).get("score", 50)), 0.0, 100.0)
        
        run_h = []
        for r in f["daily_returns"]:
            if r is None:
                run_h.append(None)
            else:
                m_score = clamp(50.0 + (r * 10.0), 0.0, 100.0)
                run_h.append(int(round(m_score * HYBRID_MOMENTUM_WEIGHT + sec * HYBRID_SECURITY_WEIGHT + sent * HYBRID_SENTIMENT_WEIGHT)))

        f["running_trend_hybrid"] = run_h
        val_l = [s for s in run_h if s is not None][-5:]
        f["last_5_scores_str"] = " ➔ ".join(str(x) for x in val_l) if val_l else "-"
        f["trend_skor"] = val_l[-1] if val_l else 50

    return len(master_dates)

def decision_label_from_score(score) -> str:
    if score is None: return "YETERSİZ VERİ"
    score = safe_float(score)
    if score >= STRONG_BUY: return "GÜÇLÜ AL"
    if score >= WATCH_LIST: return "ASIL LİSTE"
    if score >= CORRECTION: return "DÜZELTME / İZLE"
    return "ACİL SAT"

def finalize_decisions(funds: List[dict], batch_sentiments: dict):
    for f in funds:
        mom = f.get("market_momentum", 50)
        sec = f.get("security_score", 50)
        sent_data = batch_sentiments.get(f.get("investment_area", "-"), {"score": 50, "label": "Nötr", "ai_active": False, "ai_reason": "Veri Yok"})
        sent = sent_data["score"]
        
        f["sentiment_score"] = sent
        f["sentiment_label"] = sent_data["label"]
        f["sentiment_ai_active"] = sent_data.get("ai_active", False)
        f["sentiment_ai_reason"] = sent_data.get("ai_reason", "Bilinmiyor")

        dec = int(round(clamp(mom * HYBRID_MOMENTUM_WEIGHT + sec * HYBRID_SECURITY_WEIGHT + sent * HYBRID_SENTIMENT_WEIGHT, 0.0, 100.0)))
        f["decision_score"], f["karar"] = dec, decision_label_from_score(dec)
        f["data_quality_score"] = 95

# ============================================================
# EXCEL ÇIKTISI
# ============================================================

def create_excel_output(wb, ws_list, all_funds, common_n_days):
    if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")
    
    headers = [
        "Fon Kodu", "Fon Adı", "Yatırım Alanı", "Karar Skoru", "Trend Skoru", "Model Kararı",
        "Ort. Günlük Getiri (%)", "Haftalık Getiri (%)", "Veri Kaynağı"
    ]
    ws_scores.append(headers)
    for f in all_funds:
        ws_scores.append([
            f["code"], f.get("fund_title", "-"), f.get("investment_area", "-"),
            f.get("decision_score"), f.get("trend_skor"), f.get("karar"),
            round(safe_float(f.get("_final_mean_return")), 2),
            round(safe_float(f.get("weekly_return")), 2),
            f.get("source", "-")
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# ANA ARAYÜZ (STREAMLIT)
# ============================================================

if "analysis_done" not in st.session_state: st.session_state["analysis_done"] = False
if "req_codes" not in st.session_state: st.session_state["req_codes"] = []
if "wb_bytes" not in st.session_state: st.session_state["wb_bytes"] = None

st.markdown("### 📥 Veri Giriş Yöntemi")
input_method = st.radio(
    "Veri Kaynağı:",
    options=["✍️ Manuel Fon Girişi (+ / Virgül / Boşluk)", "📁 Bilgisayardan Excel Yükle", "🌐 GitHub Deposu"],
    horizontal=True,
    label_visibility="collapsed"
)

if input_method == "✍️ Manuel Fon Girişi (+ / Virgül / Boşluk)":
    st.info("💡 Fon kodlarını aralarına `+`, `,` veya boşluk koyarak yazabilirsiniz.")
    with st.form("manual_entry_form"):
        manual_input = st.text_area("Analiz Edilecek Fon Kodları", value="TI3 + MAC + TCD + BIO + YAY")
        if st.form_submit_button("🚀 Manuel Listeyi Analiz Et", type="primary", use_container_width=True):
            raw_tokens = re.split(r"[\s\+\,\;\-]+", manual_input.strip())
            codes = [normalize_fund_code(t) for t in raw_tokens if t.strip()]
            codes = list(dict.fromkeys(filter(None, codes)))
            if codes:
                temp_wb = openpyxl.Workbook()
                temp_wb.active.title = "Fon_Listesi"
                temp_wb.active.append(["Fon Kodu"])
                for c in codes: temp_wb.active.append([c])
                buf = io.BytesIO()
                temp_wb.save(buf)
                st.session_state["wb_bytes"] = buf.getvalue()
                st.session_state["req_codes"] = codes
                st.session_state["analysis_done"] = True
                st.rerun()

elif input_method == "📁 Bilgisayardan Excel Yükle":
    uploaded_file = st.file_uploader("Excel Dosyası (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        content = uploaded_file.read()
        temp_wb = openpyxl.load_workbook(io.BytesIO(content))
        ws_list = temp_wb["Fon_Listesi"] if "Fon_Listesi" in temp_wb.sheetnames else temp_wb.active
        codes = [normalize_fund_code(r[0].value) for r in ws_list.iter_rows(min_row=2) if r and r[0].value]
        codes = list(dict.fromkeys(filter(None, codes)))
        if codes:
            st.session_state["wb_bytes"] = content
            st.session_state["req_codes"] = codes
            st.session_state["analysis_done"] = True

elif input_method == "🌐 GitHub Deposu":
    if st.button("🚀 GitHub'dan Çek ve Başlat", type="primary", use_container_width=True):
        res, stat = request_with_status("GitHub", "GET", GITHUB_FALLBACK_URL)
        if res and stat.ok:
            temp_wb = openpyxl.load_workbook(io.BytesIO(res.content))
            ws_list = temp_wb["Fon_Listesi"] if "Fon_Listesi" in temp_wb.sheetnames else temp_wb.active
            codes = [normalize_fund_code(r[0].value) for r in ws_list.iter_rows(min_row=2) if r and r[0].value]
            codes = list(dict.fromkeys(filter(None, codes)))
            if codes:
                st.session_state["wb_bytes"] = res.content
                st.session_state["req_codes"] = codes
                st.session_state["analysis_done"] = True
                st.rerun()

req_codes = st.session_state.get("req_codes", [])
wb_bytes = st.session_state.get("wb_bytes")

if not req_codes or not wb_bytes:
    st.warning("⚠️ Lütfen fon kodlarını girip analizi başlatın.")
    st.stop()

wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
ws_list = wb["Fon_Listesi"] if "Fon_Listesi" in wb.sheetnames else wb.active

st.write(f"🎯 **İşlenen Fonlar ({len(req_codes)} adet):** `{', '.join(req_codes)}`")

# ============================================================
# ANALİZ MOTORU
# ============================================================

with st.spinner("🔄 Veriler toplanıyor..."):
    today = dt.date.today()
    universe = fetch_tefas_universe(today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS), today)
    meta_map = build_fund_meta_map(universe)
    ref = build_universe_reference(universe, TARGET_TRADING_DAYS)

calc_funds, failed = [], []
prog = st.progress(0, "Analiz ediliyor...")
with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as exe:
    futs = {exe.submit(fetch_and_compute_one_fund, c, universe, meta_map, {}): c for c in req_codes}
    for i, fut in enumerate(concurrent.futures.as_completed(futs)):
        c = futs[fut]
        try: _, met, src = fut.result()
        except Exception: met = None
        if met: calc_funds.append(met)
        else: failed.append(c)
        prog.progress((i + 1) / len(req_codes))
prog.empty()

if not calc_funds:
    st.error(f"❌ TEFAS veya İş Yatırım'dan veri alınamadı. Başarısız Fonlar: {', '.join(failed)}")
    st.stop()

with st.spinner("📊 Model hesaplamaları yapılıyor..."):
    calculate_security_scores(calc_funds, ref)
    calculate_market_relative_momentum(calc_funds, ref, TARGET_TRADING_DAYS)
    
    unique_areas = list(set([f.get("investment_area", "-") for f in calc_funds if f.get("investment_area")]))
    batch_sentiments = fetch_batch_market_sentiment(unique_areas or ["-"], api_key_input)
    
    common_n = calculate_trend_scores(calc_funds, batch_sentiments)
    finalize_decisions(calc_funds, batch_sentiments)

output = create_excel_output(wb, ws_list, calc_funds, common_n)

# ============================================================
# SONUÇ EKRANI
# ============================================================

st.subheader("📈 Analiz Sonuçları (V13.15)")
col1, col2, col3, col4 = st.columns(4)
scores = [safe_float(x.get("decision_score")) for x in calc_funds if x.get("decision_score") is not None]
if scores:
    col1.metric("En Yüksek Skor", f"{max(scores):.0f}")
    col2.metric("Ortalama Skor", f"{sum(scores) / len(scores):.1f}")
    col3.metric("En Düşük Skor", f"{min(scores):.0f}")
    col4.metric("Güçlü Al Veren", sum(1 for x in calc_funds if x.get("karar") == "GÜÇLÜ AL"))

display_rows = []
for item in calc_funds:
    display_rows.append({
        "Fon Kodu": item["code"],
        "Fon Adı": item.get("fund_title", "-"),
        "Yatırım Alanı": item.get("investment_area", "-"),
        "Karar Skoru": item.get("decision_score"),
        "Trend Skoru": item.get("trend_skor"),
        "Model Kararı": item.get("karar"),
        "Haftalık Getiri (%)": round(safe_float(item.get("weekly_return")), 2),
        "Veri Kaynağı": item.get("source", "-")
    })

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value).upper()
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text: return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text: return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text: return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except AttributeError: styled_df = df_display.style.applymap(color_cells)

st.dataframe(styled_df, use_container_width=True, hide_index=True)

st.success(f"✅ Toplam {len(calc_funds)} fon başarıyla analiz edildi.")
st.download_button(
    label="📥 Analiz Excel Dosyasını İndir",
    data=output,
    file_name="fon_analiz_sonuc.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if SHOW_DIAGNOSTICS:
    st.subheader("🔎 Tanılama Bilgisi")
    diag = []
    for item in calc_funds:
        for s in item.get("source_statuses", []):
            diag.append({
                "Fon": item["code"],
                "Kaynak": s.get("source"),
                "Başarılı": s.get("ok"),
                "HTTP": s.get("status_code"),
                "Süre (ms)": s.get("elapsed_ms")
            })
    if diag:
        st.dataframe(pd.DataFrame(diag), use_container_width=True, hide_index=True)
