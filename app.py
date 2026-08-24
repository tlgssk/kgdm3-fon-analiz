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
# tlgssk - SÜRÜM V14.0 (TAM & KARARLI SÜRÜM)
# ============================================================

st.set_page_config(
    page_title="tlgssk Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 tlgssk Hibrit Fon Analizi")
st.caption(
    "TEFAS + İş Yatırım | Gemini 3.7 Sentiment (Toplu Sorgu) + Tam Senkron Tarihler | V14.0"
)

# ============================================================
# AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")
DEFAULT_FUND_KIND = "YAT"

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 3

HTTP_TIMEOUT = 20
MAX_WORKERS = 4

REQUEST_MAX_RETRIES = 2
REQUEST_BACKOFF_FACTOR = 1.5

MIN_REFERENCE_SAMPLE = 5
OVERHEAT_Z_THRESHOLD = 2.0
OVERHEAT_PENALTY = 6.0

APP_VERSION = "14.0.0"

GITHUB_OWNER = "tlgssk"
GITHUB_REPO = "kgdm3-fon-analiz"
GITHUB_BRANCH = "main"
GITHUB_FALLBACK_URL = "https://github.com/tlgssk/kgdm3-fon-analiz/raw/refs/heads/main/Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"

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
    retry = Retry(total=REQUEST_MAX_RETRIES, connect=REQUEST_MAX_RETRIES, read=REQUEST_MAX_RETRIES,
                  status=REQUEST_MAX_RETRIES, backoff_factor=REQUEST_BACKOFF_FACTOR,
                  status_forcelist=(429, 500, 502, 503, 504),
                  allowed_methods=frozenset({"GET", "POST"}),
                  respect_retry_after_header=True)
    adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"User-Agent": "KGDM3-Fon-Analiz/13.15", "Accept": "application/json,text/html"})
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

with st.sidebar.expander("⚖️ Skor Ağırlıkları (V14.0)"):
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

RISK_FREE_ANNUAL = st.sidebar.number_input(
    "Yıllık risksiz getiri (%)",
    min_value=0.0, max_value=100.0, value=0.0, step=0.25,
    help="Sharpe/Sortino hesabında kullanılacak yıllık oran. 0 seçilirse risksiz getiri etkisi yoktur."
)

SHOW_DIAGNOSTICS = st.sidebar.checkbox("Kaynak tanılama bilgisini göster", value=True)

# ============================================================
# CANLI GEMINI DUYARLILIK MOTORU
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

@st.cache_data(ttl=60 * 60 * 4, show_spinner=False)
def fetch_batch_market_sentiment(areas: list, api_key: str) -> dict:
    result_map = {}
    api_key_clean = api_key.strip() if api_key else ""
    
    if not api_key_clean:
        for area in areas:
            a_u = area.upper()
            if "YABANCI TEKNOLOJİ" in a_u or "YABANCI" in a_u:
                result_map[area] = {"score": 38, "label": "Negatif (Kâr Satışı)", "ai_active": False, "ai_reason": "API Anahtarı Girilmedi"}
            elif "ALTIN" in a_u or "GÜMÜŞ" in a_u or "KIYMETLİ" in a_u:
                result_map[area] = {"score": 82, "label": "Güçlü Pozitif (Faiz İndirimi)", "ai_active": False, "ai_reason": "API Anahtarı Girilmedi"}
            elif "PARA PİYASASI" in a_u or "BORÇLANMA" in a_u:
                result_map[area] = {"score": 65, "label": "Pozitif (Sabit Getiri)", "ai_active": False, "ai_reason": "API Anahtarı Girilmedi"}
            elif "HİSSE" in a_u or "BIST" in a_u:
                result_map[area] = {"score": 54, "label": "Dengeli / Pozitif Beklenti", "ai_active": False, "ai_reason": "API Anahtarı Girilmedi"}
            else:
                result_map[area] = {"score": 50, "label": "Nötr / Kural Tabanlı", "ai_active": False, "ai_reason": "API Anahtarı Girilmedi"}
        return result_map

    areas_text = "\n".join([f"- {a}" for a in areas])
    prompt = f"""Sen kıdemli bir fon yöneticisi ve makroekonomik analistsin.
Aşağıdaki yatırım alanları için güncel piyasa duyarlılığını 0-100 arası puanla ve max 6 kelimelik gerekçe üret:
{areas_text}
SADECE geçerli JSON objesi üret:
{{"Alan Adı": {{"score": 75, "label": "Kısa gerekçe"}}}}"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.7-flash:generateContent?key={api_key_clean}"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json"}
    }

    last_err = ""
    for attempt in range(3):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=20)
            if response.status_code == 200:
                data = response.json()
                raw_text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "{}")
                parsed_data = json.loads(raw_text.strip("```json\n").strip("```").strip())
                for area in areas:
                    if area in parsed_data:
                        result_map[area] = {
                            "score": int(clamp(safe_float(parsed_data[area].get("score", 50)), 0.0, 100.0)),
                            "label": str(parsed_data[area].get("label", "Nötr")),
                            "ai_active": True,
                            "ai_reason": "Bağlantı Başarılı"
                        }
                    else:
                        result_map[area] = {"score": 50, "label": "Nötr", "ai_active": True, "ai_reason": "Varsayılan"}
                return result_map
            elif response.status_code == 429:
                time.sleep(10 + attempt * 10)
            else:
                last_err = f"HTTP {response.status_code}"
                time.sleep(2)
        except Exception as exc:
            last_err = str(exc)[:30]
            time.sleep(3)

    for area in areas:
        result_map[area] = {"score": 50, "label": "Nötr", "ai_active": False, "ai_reason": f"API Hatası: {last_err}"}
    return result_map

# ============================================================
# TEFAS API VE METRİKLER 
# ============================================================

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

def audit_fund_data(fund: dict) -> dict:
    score = 100.0
    issues = []
    prices = fund.get("prices") or []
    if len(prices) < 2:
        issues.append("Yetersiz fiyat gözlemi")
        score -= 20
    fund["data_quality_score"] = int(round(clamp(score, 0, 100)))
    fund["data_quality_issues"] = " | ".join(dict.fromkeys(issues)) if issues else "OK"
    return fund

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
                if len(df) >= 2: return df[["date", "price", "aum", "investors"]], status
        except: pass
    return None, status

def fetch_tefas_direct_api(fund_code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(fund_code)
    status = new_status("TEFAS Direct API")
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {"Origin": "https://www.tefas.gov.tr", "Referer": "https://www.tefas.gov.tr/TarihselVeriler.aspx"}
    
    for kind in ([fund_kind] if fund_kind in FUND_KINDS else ["YAT", "EMK", "BYF"]):
        payload = {"fontip": kind, "fonkod": code, "bastarih": start.strftime("%d.%m.%Y"), "bittarih": end.strftime("%d.%m.%Y")}
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
                    df = df.dropna(subset=["date", "price"])[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                    if len(df) >= 2: return df, stat
            except: pass
    return None, status

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

    df_dir, stat_dir = fetch_tefas_direct_api(code, fund_kind)
    statuses.append(stat_dir)
    if df_dir is not None and len(df_dir) >= 2:
        return df_dir, "TEFAS Direct API", statuses

    if universe is not None and not universe.empty:
        rows = universe[universe["code"].eq(code)].copy()
        if len(rows) >= 2:
            ok_status = new_status("TEFAS Evren")
            ok_status.attempted = True; ok_status.ok = True; ok_status.message = "Evrenden alındı"
            statuses.append(ok_status)
            return rows.sort_values("date").reset_index(drop=True), "TEFAS", statuses

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
    aum_first = next((v for v in aums if v is not None and v > 0), None)
    inv_first = next((v for v in invs if v is not None and v > 0), None)

    aum_change = ((aum_last / aum_first) - 1.0) * 100.0 if aum_last and aum_first else None
    inv_change = ((inv_last / inv_first) - 1.0) * 100.0 if inv_last is not None and inv_first else None
    price_cum = ((prices[-1] / prices[0]) - 1.0) * 100.0 if prices[0] > 0 else None
    aum_flow_proxy = (aum_change - price_cum) if aum_change is not None and price_cum is not None else None

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
        "aum_change": aum_change,
        "aum_flow_proxy": aum_flow_proxy,
        "inv_change": inv_change,
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


def percentile_rank(value, population, neutral=50.0):
    """0-100 percentile rank. Returns neutral when population is too small."""
    if value is None:
        return neutral
    vals = [optional_float(x) for x in population]
    vals = [x for x in vals if x is not None and math.isfinite(x)]
    if len(vals) < MIN_REFERENCE_SAMPLE:
        return neutral
    less_equal = sum(1 for x in vals if x <= value)
    return 100.0 * less_equal / len(vals)


def safe_std(values):
    vals = [optional_float(v) for v in values]
    vals = [v for v in vals if v is not None and math.isfinite(v)]
    if len(vals) < 2:
        return 0.0
    mean_v = sum(vals) / len(vals)
    return (sum((x - mean_v) ** 2 for x in vals) / len(vals)) ** 0.5


def annualized_volatility(daily_returns):
    vol = safe_std(daily_returns)
    return vol * math.sqrt(252.0)


def annualized_sharpe(daily_returns, risk_free_annual=0.0):
    vals = [optional_float(v) for v in daily_returns]
    vals = [v for v in vals if v is not None and math.isfinite(v)]
    if len(vals) < 3:
        return None
    rf_daily = (1.0 + risk_free_annual / 100.0) ** (1.0 / 252.0) - 1.0
    rf_daily_pct = rf_daily * 100.0
    excess = [v - rf_daily_pct for v in vals]
    vol = safe_std(excess)
    if vol <= 1e-12:
        return None
    return (sum(excess) / len(excess)) / vol * math.sqrt(252.0)


def annualized_sortino(daily_returns, risk_free_annual=0.0):
    vals = [optional_float(v) for v in daily_returns]
    vals = [v for v in vals if v is not None and math.isfinite(v)]
    if len(vals) < 3:
        return None
    rf_daily = ((1.0 + risk_free_annual / 100.0) ** (1.0 / 252.0) - 1.0) * 100.0
    excess = [v - rf_daily for v in vals]
    downside = [min(v, 0.0) for v in excess]
    downside_dev = (sum(x * x for x in downside) / len(downside)) ** 0.5
    if downside_dev <= 1e-12:
        return None
    return (sum(excess) / len(excess)) / downside_dev * math.sqrt(252.0)


def calmar_ratio(daily_returns, max_dd):
    if max_dd is None or max_dd >= -1e-12:
        return None
    cum = calculate_compounded_return(daily_returns)
    return (cum / abs(max_dd)) if abs(max_dd) > 1e-12 else None


def calculate_data_quality(fund):
    """
    Quality is evidence-based. Missing structural fields are not silently
    converted into a perfect score.
    """
    n = int(fund.get("n_days", 0) or 0)
    source = fund.get("source", "YOK")
    prices = fund.get("prices") or []
    issues = []

    if n < 3:
        issues.append("Çok kısa fiyat geçmişi")
    elif n < 5:
        issues.append("Kısa fiyat geçmişi")
    elif n < TARGET_TRADING_DAYS:
        issues.append("10 günden kısa geçmiş")

    if len(prices) >= 2:
        bad = sum(1 for x in prices if optional_float(x) is None or optional_float(x) <= 0)
        if bad:
            issues.append("Geçersiz fiyat gözlemi")

    if source == "YOK":
        issues.append("Kaynak yok")
    elif source == "İş Yatırım":
        issues.append("TEFAS yerine fallback kaynak")

    if fund.get("aum") is None:
        issues.append("AUM yok")
    if fund.get("investors") is None:
        issues.append("Yatırımcı sayısı yok")
    if not fund.get("structural_fetch_ok", False):
        issues.append("Portföy yapısal verisi yok")

    score = 100.0
    if n < 3: score -= 35
    elif n < 5: score -= 20
    elif n < TARGET_TRADING_DAYS: score -= 8
    if source == "İş Yatırım": score -= 5
    if source == "YOK": score -= 40
    if fund.get("aum") is None: score -= 4
    if fund.get("investors") is None: score -= 4
    if not fund.get("structural_fetch_ok", False): score -= 8

    fund["data_quality_score"] = int(round(clamp(score, 0, 100)))
    fund["data_quality_issues"] = " | ".join(dict.fromkeys(issues)) if issues else "OK"
    return fund


def calculate_security_scores(funds: List[dict], reference: dict):
    """
    Security/liquidity is deliberately conservative. It only rewards
    evidence actually available in the dataset and never fabricates
    concentration or liquidity values.
    """
    for f in funds:
        score = 50.0
        known = 0

        aum = optional_float(f.get("aum"))
        investors = optional_float(f.get("investors"))
        dd = optional_float(f.get("max_dd"))
        aum_change = optional_float(f.get("aum_change"))
        inv_change = optional_float(f.get("inv_change"))

        if aum is not None and aum > 0:
            known += 1
            # Log-scaled AUM: 10m -> ~0, 1bn -> ~20, 10bn -> ~30.
            aum_component = clamp((math.log10(max(aum, 1.0)) - 7.0) * 10.0, -10.0, 20.0)
            score += aum_component * SECURITY_WEIGHTS["aum"]

        if investors is not None and investors > 0:
            known += 1
            investor_component = clamp((math.log10(max(investors, 1.0)) - 2.0) * 12.0, -10.0, 20.0)
            score += investor_component * SECURITY_WEIGHTS["investor"]

        if aum_change is not None:
            known += 1
            score += clamp(aum_change * 0.10, -10.0, 10.0)

        if inv_change is not None:
            known += 1
            score += clamp(inv_change * 0.05, -8.0, 8.0)

        if dd is not None:
            # Less severe drawdown => higher score.
            score += clamp((dd + 10.0) * 0.50, -10.0, 5.0)

        if f.get("is_bist30"):
            score += BIST30_BONUS

        # Missing structural liquidity/concentration is penalized modestly,
        # but never replaced by invented values.
        if not f.get("structural_fetch_ok", False):
            score -= 4.0

        f["security_score"] = int(round(clamp(score, 0.0, 100.0)))
        f["security_known_fields"] = known
        f["security_data_status"] = (
            "Yeterli" if known >= 3 else
            "Kısmi" if known >= 1 else
            "Yetersiz"
        )


def calculate_market_relative_momentum(
    funds: List[dict], reference, window: int, risk_free_annual: float = 0.0
):
    """
    Produces a genuine peer-relative score when the reference sample is
    large enough. Absolute metrics are also retained for transparency.
    """
    for f in funds:
        k = f.get("kind", DEFAULT_FUND_KIND)
        rets = [x for x in (f.get("orig_returns", [])[-window:]) if optional_float(x) is not None]
        prices = f.get("prices") or []
        prc = prices[-(len(rets) + 1):] if prices else []

        if len(rets) < 2:
            f.update({
                "market_momentum": 50,
                "relative_return_score": 50,
                "relative_sharpe_score": 50,
                "relative_cumulative_score": 50,
                "relative_drawdown_score": 50,
                "relative_available": False,
                "reference_scope": f"Piyasa ({k}) - yetersiz örneklem",
            })
            continue

        m_r = sum(rets) / len(rets)
        vol_daily = safe_std(rets)
        cum = ((prc[-1] / prc[0]) - 1.0) * 100.0 if len(prc) >= 2 and prc[0] > 0 else calculate_compounded_return(rets)
        dd = calculate_max_drawdown(prc) if len(prc) >= 2 else 0.0
        sharpe = annualized_sharpe(rets, risk_free_annual)
        sortino = annualized_sortino(rets, risk_free_annual)
        calmar = calmar_ratio(rets, dd)

        rr = percentile_rank(m_r, reference.get(k, {}).get("mean_return", []))
        rs = percentile_rank(
            sharpe, reference.get(k, {}).get("sharpe", [])
        ) if sharpe is not None else 50.0
        rc = percentile_rank(cum, reference.get(k, {}).get("cumulative", []))
        # Reference stores negative drawdown; invert ranking so smaller loss is better.
        rd_raw = percentile_rank(dd, reference.get(k, {}).get("max_dd_inv", []))
        rd = rd_raw if reference_sample_size(reference, k) >= MIN_REFERENCE_SAMPLE else 50.0

        weights = MOMENTUM_WEIGHTS
        relative = (
            rr * weights["return"]
            + rs * weights["sharpe"]
            + rc * weights["cumulative"]
            + rd * weights["drawdown"]
        )
        # If peer reference is unavailable, use a transparent absolute fallback.
        reference_ok = reference_sample_size(reference, k) >= MIN_REFERENCE_SAMPLE
        if not reference_ok:
            z_r = clamp(m_r / max(vol_daily, 0.01), -3, 3)
            relative = clamp(50.0 + z_r * 10.0, 5.0, 95.0)

        f.update({
            "_final_mean_return": m_r,
            "_final_sharpe": sharpe,
            "_final_sortino": sortino,
            "_final_calmar": calmar,
            "_final_cumulative": cum,
            "_final_max_dd": dd,
            "volatility": vol_daily,
            "annualized_volatility": annualized_volatility(rets),
            "relative_return_score": rr,
            "relative_sharpe_score": rs,
            "relative_cumulative_score": rc,
            "relative_drawdown_score": rd,
            "relative_available": reference_ok,
            "reference_sample_size": reference_sample_size(reference, k),
            "reference_scope": f"Piyasa ({k})" if reference_ok else f"Piyasa ({k}) - fallback",
            "market_momentum": int(round(clamp(relative, 0.0, 100.0))),
        })


def calculate_trend_scores(
    funds: List[dict], batch_sentiments: dict, risk_free_annual: float = 0.0
) -> int:
    if not funds:
        return 0

    all_dates = set()
    for f in funds:
        all_dates.update(f.get("orig_dates", []))

    master_dates = sorted(list(all_dates))[-TARGET_TRADING_DAYS:]
    if not master_dates:
        return 0

    for f in funds:
        ret_map = dict(zip(f.get("orig_dates", []), f.get("orig_returns", [])))
        f["dates"] = master_dates
        f["daily_returns"] = [ret_map.get(d) for d in master_dates]
        valid_returns = [r for r in f["daily_returns"] if r is not None]

        sec = safe_float(f.get("security_score"), 50.0)
        sent = clamp(
            safe_float(
                batch_sentiments.get(
                    f.get("investment_area", "-"), {}
                ).get("score", 50)
            ),
            0.0, 100.0
        )

        # Daily trend uses standardized own-history return rather than a
        # raw "1% = +10 points" conversion.
        own_mean = sum(valid_returns) / len(valid_returns) if valid_returns else 0.0
        own_std = safe_std(valid_returns)
        run_h = []
        for r in f["daily_returns"]:
            if r is None:
                run_h.append(None)
                continue
            z = (r - own_mean) / own_std if own_std > 1e-12 else 0.0
            m_score = clamp(50.0 + 18.0 * math.tanh(z / 2.0), 0.0, 100.0)
            daily = (
                m_score * HYBRID_MOMENTUM_WEIGHT
                + sec * HYBRID_SECURITY_WEIGHT
                + sent * HYBRID_SENTIMENT_WEIGHT
            )
            run_h.append(int(round(clamp(daily, 0.0, 100.0))))

        f["running_trend_hybrid"] = run_h
        val_l = [s for s in run_h if s is not None][-5:]
        f["last_5_scores_str"] = " ➔ ".join(str(x) for x in val_l) if val_l else "-"
        f["trend_skor"] = val_l[-1] if val_l else 50
        f["n_days"] = len(valid_returns)

    return len(master_dates)


def calculate_confidence(fund):
    """
    Confidence is not probability of profit. It measures evidence quality:
    history length, source quality, completeness and peer reference quality.
    """
    n = int(fund.get("n_days", 0) or 0)
    quality = safe_float(fund.get("data_quality_score"), 0)
    source = fund.get("source")
    ref_n = int(fund.get("reference_sample_size", 0) or 0)

    history_component = clamp(n / max(TARGET_TRADING_DAYS, 1), 0.0, 1.0) * 35.0
    quality_component = clamp(quality / 100.0, 0.0, 1.0) * 35.0
    source_component = 15.0 if source == "TEFAS Direct API" else (12.0 if source == "TEFAS" else 8.0)
    reference_component = clamp(ref_n / 50.0, 0.0, 1.0) * 15.0

    confidence = clamp(
        history_component + quality_component + source_component + reference_component,
        0.0, 100.0
    )
    fund["confidence_score"] = int(round(confidence))
    return fund


def decision_label_from_score(score) -> str:
    if score is None: return "YETERSİZ VERİ"
    score = safe_float(score)
    if score >= STRONG_BUY: return "GÜÇLÜ AL"
    if score >= WATCH_LIST: return "ASIL LİSTE"
    if score >= CORRECTION: return "DÜZELTME / İZLE"
    return "ACİL SAT"


def finalize_decisions(funds: List[dict], batch_sentiments: dict):
    for f in funds:
        calculate_data_quality(f)
        calculate_confidence(f)

        mom = safe_float(f.get("market_momentum"), 50)
        sec = safe_float(f.get("security_score"), 50)
        sent_data = batch_sentiments.get(
            f.get("investment_area", "-"),
            {"score": 50, "label": "Nötr", "ai_active": False, "ai_reason": "Veri Yok"}
        )
        sent = clamp(safe_float(sent_data.get("score"), 50), 0.0, 100.0)

        f["sentiment_score"] = sent
        f["sentiment_label"] = sent_data.get("label", "Nötr")
        f["sentiment_ai_active"] = sent_data.get("ai_active", False)
        f["sentiment_ai_reason"] = sent_data.get("ai_reason", "Bilinmiyor")

        # A decision is penalized when evidence is weak. This prevents
        # one-day/poor-source observations from producing a falsely strong signal.
        raw_dec = (
            mom * HYBRID_MOMENTUM_WEIGHT
            + sec * HYBRID_SECURITY_WEIGHT
            + sent * HYBRID_SENTIMENT_WEIGHT
        )
        confidence = safe_float(f.get("confidence_score"), 0)
        evidence_factor = 0.70 + 0.30 * (confidence / 100.0)
        dec = 50.0 + (raw_dec - 50.0) * evidence_factor

        f["raw_decision_score"] = int(round(clamp(raw_dec, 0, 100)))
        f["decision_score"] = int(round(clamp(dec, 0.0, 100.0)))
        f["karar"] = decision_label_from_score(f["decision_score"])

        # Explainability
        f["decision_components"] = (
            f"Momentum {mom:.0f} | Güvenlik {sec:.0f} | "
            f"Sentiment {sent:.0f} | Güven {confidence:.0f}"
        )

# ============================================================
# GÜVENLİ EXCEL ÇIKTISI (HATA VERMEYEN DİREKT RENKLENDİRME)
# ============================================================

def create_excel_output(wb, ws_list, all_funds, common_n_days):
    if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

    n_dates = common_n_days if common_n_days > 0 else 5
    all_dates = set()
    for f in all_funds:
        for d in f.get("dates", []):
            if d is not None: all_dates.add(d)

    sorted_dates = sorted(list(all_dates))
    sample_dates = sorted_dates[-n_dates:] if len(sorted_dates) >= n_dates else sorted_dates
    last_5_dates = sample_dates[-5:] if len(sample_dates) >= 5 else sample_dates

    headers = [
        "Fon Kodu", "Fon Adı", "Yatırım Alanı", "Valör", "Karar Skoru (Piyasa-Bağıl)", "Trend Skoru (Liste-Bağıl)",
        "Piyasa Momentum", "Güvenlik/Likidite Skoru", "Sentiment Skoru", "Duyarlılık Yönü", "Referans Kapsamı", "Veri Kalitesi", "Güven Skoru", "Ham Karar Skoru", "Aşırı Isınma",
        "Son 5 Trend Skoru", "Model Kararı", "Ort. Günlük Getiri (%)", "Yıllıklandırılmış Volatilite (%)", "Sharpe", "Sortino", "Calmar",
        "Kümülatif Getiri (%)", "MaxDD (%)", "En Büyük Varlık (%)", "BIST30", "Net Likidite (%)", "KAZRİSK",
        "AUM Değişim (%)", "AUM Akış Proxy (%)", "Yatırımcı Değişim (%)", "AUM (₺)", "Yatırımcı",
        "Haftalık Bileşik (%)", "Veri Kaynağı", "Kalite Uyarıları"
    ]

    daily_headers = []
    for day in reversed(last_5_dates): 
        daily_headers.extend([f"{display_date(day)} Karar Skoru", f"{display_date(day)} Model Kararı"])
    headers[3:3] = daily_headers

    for day in sample_dates: headers.append(f"{display_date(day)} Trend Skor")
    for day in sample_dates: headers.append(f"{display_date(day)} Getiri")

    ws_scores.append(headers)
    header_index = {name: idx + 1 for idx, name in enumerate(headers)}

    fill_header = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    font_header = Font(name="Calibri", bold=True, color=COLOR_WHITE)
    for cell in ws_scores[1]:
        cell.fill, cell.font, cell.alignment = fill_header, font_header, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_scores.row_dimensions[1].height = 55

    for item in all_funds:
        row_data = [item["code"], item.get("fund_title") or "-", item.get("investment_area") or "-"]
        fund_dates = item.get("dates", [])
        fund_scores = item.get("running_trend_hybrid", [])
        fund_rets = item.get("daily_returns", [])
        
        score_map = dict(zip(fund_dates, fund_scores))
        ret_map = dict(zip(fund_dates, fund_rets))

        for day in reversed(last_5_dates):
            s = score_map.get(day)
            row_data.extend([s if s is not None else "Veri Açıklanmadı", decision_label_from_score(s) if s is not None else "Veri Açıklanmadı"])

        row_data.extend([
            item.get("valor"), item.get("decision_score"), item.get("trend_skor"), item.get("market_momentum"),
            item.get("security_score"), item.get("sentiment_score"), item.get("sentiment_label"), item.get("reference_scope", "-"),
            item.get("data_quality_score"), item.get("confidence_score"), item.get("raw_decision_score"),
            item.get("last_5_scores_str", "-"), item.get("karar", "-"),
            round(safe_float(item.get("_final_mean_return")), 4),
            round(safe_float(item.get("annualized_volatility")), 4),
            round(safe_float(item.get("_final_sharpe")), 4),
            round(safe_float(item.get("_final_sortino")), 4),
            round(safe_float(item.get("_final_calmar")), 4),
            round(safe_float(item.get("_final_cumulative")), 4),
            round(safe_float(item.get("_final_max_dd")), 4),
            None, "EVET" if item.get("is_bist30") else "HAYIR",
            "Veri Yok", "Veri Yok",
            round(safe_float(item.get("aum_change")), 2), round(safe_float(item.get("aum_flow_proxy")), 2),
            round(safe_float(item.get("inv_change")), 2), round(safe_float(item.get("aum")), 0), item.get("investors"),
            round(safe_float(item.get("weekly_return")), 4), item.get("source", "-"), item.get("data_quality_issues", "")
        ])

        for day in sample_dates:
            s = score_map.get(day)
            row_data.append(s if s is not None else "Veri Açıklanmadı")
            
        for day in sample_dates:
            r = ret_map.get(day)
            row_data.append(format_percent(r) if r is not None else "Veri Açıklanmadı")

        ws_scores.append(row_data)

    green_font, red_font, yellow_font = Font(bold=True, color=COLOR_GREEN), Font(bold=True, color=COLOR_RED), Font(bold=True, color=COLOR_YELLOW)
    fill_green = PatternFill(start_color=COLOR_LIGHT_GREEN, fill_type="solid")
    fill_yellow = PatternFill(start_color=COLOR_LIGHT_YELLOW, fill_type="solid")
    fill_red = PatternFill(start_color=COLOR_LIGHT_RED, fill_type="solid")

    decision_cols = [idx for name, idx in header_index.items() if "Karar" in name and "Skor" not in name]
    score_cols = [idx for name, idx in header_index.items() if "Skor" in name]

    for row_number in range(2, ws_scores.max_row + 1):
        for col_idx in decision_cols:
            cell = ws_scores.cell(row=row_number, column=col_idx)
            text = str(cell.value or "").upper()
            if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text: cell.font = green_font
            elif "DÜZELTME" in text: cell.font = yellow_font
            elif "ACİL SAT" in text: cell.font = red_font

        for col_idx in score_cols:
            cell = ws_scores.cell(row=row_number, column=col_idx)
            if isinstance(cell.value, (int, float)):
                if cell.value >= 75: cell.fill = fill_green
                elif cell.value >= 50: cell.fill = fill_yellow
                else: cell.fill = fill_red

    thin = Side(style="thin", color="D9E1F2")
    for row in ws_scores.iter_rows():
        for cell in row: cell.alignment, cell.border = Alignment(vertical="center"), Border(bottom=thin)

    ws_scores.freeze_panes = "A2"
    ws_scores.sheet_view.showGridLines = False

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# ANA ARAYÜZ (STREAMLIT) - GELİŞMİŞ GİRİŞ PANELİ
# ============================================================

if "analysis_done" not in st.session_state: st.session_state["analysis_done"] = False
if "req_codes" not in st.session_state: st.session_state["req_codes"] = []
if "wb_bytes" not in st.session_state: st.session_state["wb_bytes"] = None

st.markdown("### 📥 Veri Giriş Yöntemi Seçin")
input_method = st.radio(
    "Veri Kaynağı:",
    options=["✍️ Manuel Fon Girişi (+ / Virgül / Boşluk)", "📁 Bilgisayardan Excel Yükle", "🌐 GitHub Deposu"],
    horizontal=True,
    label_visibility="collapsed"
)

if input_method == "✍️ Manuel Fon Girişi (+ / Virgül / Boşluk)":
    st.info("💡 Fon kodlarını aralarına `+`, `,` (virgül), `;` veya boşluk koyarak yazabilirsiniz.")
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
    uploaded_file = st.file_uploader("Bilgisayardan Excel Seçin (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        try:
            content = uploaded_file.read()
            temp_wb = openpyxl.load_workbook(io.BytesIO(content))
            ws_list = temp_wb["Fon_Listesi"] if "Fon_Listesi" in temp_wb.sheetnames else temp_wb.active
            codes = [normalize_fund_code(r[0].value) for r in ws_list.iter_rows(min_row=2) if r and r[0].value]
            codes = list(dict.fromkeys(filter(None, codes)))
            if codes:
                st.session_state["wb_bytes"] = content
                st.session_state["req_codes"] = codes
                st.session_state["analysis_done"] = True
        except Exception as exc:
            st.error(f"Excel yükleme hatası: {exc}")

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
    st.warning("⚠️ Lütfen analiz başlatmak için en az bir geçerli fon kodu girin veya dosya yükleyin.")
    st.stop()

wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
ws_list = wb["Fon_Listesi"] if "Fon_Listesi" in wb.sheetnames else wb.active

st.write(f"🎯 **Analize Alınan Fonlar ({len(req_codes)} adet):** `{', '.join(req_codes)}`")

# ============================================================
# ANALİZ MOTORU & ÇALIŞTIRMA
# ============================================================

with st.spinner("🔄 TEFAS verileri alınıyor..."):
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

eligible = [f for f in calc_funds if f.get("n_days", 0) >= 1]

if not eligible:
    st.error(f"❌ Belirtilen fonlar için TEFAS/İş Yatırım üzerinden fiyat geçmişi alınamadı. Hatalı Fonlar: {', '.join(failed)}")
    st.stop()

with st.spinner("📊 V13 Modeli (Gemini Toplu Sentiment + Baseline) Hesaplanıyor..."):
    for f in eligible: audit_fund_data(f)
    calculate_security_scores(eligible, ref)
    calculate_market_relative_momentum(eligible, ref, TARGET_TRADING_DAYS, RISK_FREE_ANNUAL)
    
    unique_areas = list(set([f.get("investment_area", "-") for f in eligible if f.get("investment_area")]))
    batch_sentiments = fetch_batch_market_sentiment(unique_areas or ["-"], api_key_input)
    
    common_n = calculate_trend_scores(eligible, batch_sentiments, RISK_FREE_ANNUAL)
    finalize_decisions(eligible, batch_sentiments)

output = create_excel_output(wb, ws_list, eligible, common_n)

# ============================================================
# SKOR ÖZETLERİ VE EKRAN TABLOSU
# ============================================================

st.subheader("📈 KAZRİSK Portföy Özeti (V14.0)")
col1, col2, col3, col4 = st.columns(4)
scores = [safe_float(x.get("decision_score")) for x in eligible if x.get("decision_score") is not None]
if scores:
    col1.metric("En Yüksek Skor", f"{max(scores):.0f}")
    col2.metric("Ortalama Skor", f"{sum(scores) / len(scores):.1f}")
    col3.metric("En Düşük Skor", f"{min(scores):.0f}")
    col4.metric("Güçlü Al Veren", sum(1 for x in eligible if x.get("karar") == "GÜÇLÜ AL"))

display_rows = []
early_alerts = []

all_dates_ui = set()
for f in eligible:
    for d in f.get("dates", []):
        if d is not None: all_dates_ui.add(d)

sorted_dates_ui = sorted(list(all_dates_ui))
sample_dates_ui = sorted_dates_ui[-common_n:] if common_n > 0 else sorted_dates_ui
last_5_dates_web = sample_dates_ui[-5:] if len(sample_dates_ui) >= 5 else sample_dates_ui

for item in eligible:
    row_dict = {
        "Fon Kodu": item["code"],
        "Fon Adı": item.get("fund_title") or "-",
        "Yatırım Alanı": item.get("investment_area") or "-",
    }

    fund_dates = item.get("dates", [])
    own_scores = item.get("running_trend_hybrid") or []
    score_map = dict(zip(fund_dates, own_scores))

    for day in reversed(last_5_dates_web):
        s = score_map.get(day)
        row_dict[f"{display_date(day)} Karar Skoru"] = s if s is not None else "Veri Açıklanmadı"
        row_dict[f"{display_date(day)} Model Kararı"] = decision_label_from_score(s) if s is not None else "Veri Açıklanmadı"

    row_dict.update({
        "Sentiment Skoru": item.get("sentiment_score"),
        "Duyarlılık Yönü": item.get("sentiment_label"),
        "Güncel Karar Skoru": item.get("decision_score"),
        "Trend Skoru": item.get("trend_skor"),
        "Güncel Karar": item.get("karar"),
        "Net Likidite (%)": "Veri Yok",
        "KAZRİSK Konsantrasyon": "🛡️ Dengeli",
        "Haftalık Getiri (%)": round(safe_float(item.get("weekly_return")), 2),
        "Veri Kalite Skoru": item.get("data_quality_score"),
    })
    display_rows.append(row_dict)

    valid_history = [(d, s) for d, s in zip(fund_dates, own_scores) if s is not None]
    if len(valid_history) >= 2:
        (d1, s1), (d2, s2) = valid_history[-2], valid_history[-1]
        lbl1, lbl2 = decision_label_from_score(s1), decision_label_from_score(s2)
        if lbl1 == "ACİL SAT" and lbl2 == "ACİL SAT":
            early_alerts.append({"Tip": "SAT", "Fon Kodu": item["code"], "Fon Adı": item.get("fund_title"), "Alan": item.get("investment_area"), "KAZRİSK Durumu": "🚨 2 GÜN TEYİTLİ ACİL SAT", "Son 2 Gün": f"{display_date(d1)} → {display_date(d2)}", "Son Skor": s2})
        elif lbl1 == "GÜÇLÜ AL" and lbl2 == "GÜÇLÜ AL":
            early_alerts.append({"Tip": "AL", "Fon Kodu": item["code"], "Fon Adı": item.get("fund_title"), "Alan": item.get("investment_area"), "KAZRİSK Durumu": "🚀 2 GÜN TEYİTLİ GÜÇLÜ AL", "Son 2 Gün": f"{display_date(d1)} → {display_date(d2)}", "Son Skor": s2})

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value).upper()
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "DENGELİ" in text: return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text: return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text: return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except AttributeError: styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 Analiz Sonuçları — Son 5 İşlem Günü Kararları (V14.0)")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

# ============================================================
# ALARM TABLOLARI
# ============================================================
sell_alerts = [{k: v for k, v in a.items() if k != "Tip"} for a in early_alerts if a["Tip"] == "SAT"]
buy_alerts = [{k: v for k, v in a.items() if k != "Tip"} for a in early_alerts if a["Tip"] == "AL"]

if sell_alerts or buy_alerts:
    st.subheader("🚨/🚀 KAZRİSK® 2 Günlük Teyitli Alarmlar")
    col_alert1, col_alert2 = st.columns(2)
    with col_alert1:
        st.markdown("### 🚨 Satış Alarmları")
        if sell_alerts: st.dataframe(pd.DataFrame(sell_alerts), use_container_width=True, hide_index=True)
        else: st.info("Şu an teyitli 'Acil Sat' sinyali veren fon yok.")
    with col_alert2:
        st.markdown("### 🚀 Fırsat Alarmları")
        if buy_alerts: st.dataframe(pd.DataFrame(buy_alerts), use_container_width=True, hide_index=True)
        else: st.success("Şu an teyitli 'Güçlü Al' fırsatı veren fon yok.")

st.success(f"✅ V14.0 Analiz tamamlandı. Toplam {len(eligible)} fon işlendi.")
st.download_button(
    label="📥 KAZRİSK V14.0 Excel İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_FINAL_V14_0.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if SHOW_DIAGNOSTICS:
    st.subheader("🔎 Veri Kaynağı Tanılaması & Gemini AI Modu")
    diagnostic_rows = []
    for item in eligible:
        reason = item.get("sentiment_ai_reason", "Bilinmiyor")
        ai_status = "🟢 Aktif (Canlı API)" if item.get("sentiment_ai_active") else f"🔴 Pasif ({reason})"
        for status in item.get("source_statuses", []):
            diagnostic_rows.append({
                "Fon": item["code"],
                "Kaynak": status.get("source"),
                "Denendi": "Evet" if status.get("attempted") else "Hayır",
                "Başarılı": "Evet" if status.get("ok") else "Hayır",
                "HTTP": status.get("status_code"),
                "Hata": status.get("error_type"),
                "Süre ms": status.get("elapsed_ms"),
                "Retry": status.get("retry_count"),
                "Mesaj": status.get("message"),
                "Gemini AI Modu": ai_status
            })
    if diagnostic_rows:
        df_diag = pd.DataFrame(diagnostic_rows)
        def color_ai_status(val):
            if isinstance(val, str):
                if "🟢 Aktif" in val: return 'color: #008000; font-weight: bold;'
                elif "🔴 Pasif" in val: return 'color: #FF0000; font-weight: bold;'
            return ''
        try: styled_diag = df_diag.style.map(color_ai_status)
        except AttributeError: styled_diag = df_diag.style.applymap(color_ai_status)
        st.dataframe(styled_diag, use_container_width=True, hide_index=True)
