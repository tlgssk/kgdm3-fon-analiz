# ============================================================
# tlgssk - SÜRÜM V14.4 (GÜÇLENDİRİLMİŞ VERİ ALMA VE HATA DÜZELTME)
# ============================================================

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
# STREAMLIT SAYFA YAPILANDIRMASI
# ============================================================

st.set_page_config(
    page_title="tlgssk Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 tlgssk Hibrit Fon Analizi")
st.caption(
    "TEFAS + İş Yatırım | Gemini Sentiment + Çoklu Fon Girdi Motoru | V14.4 (Kararlı Sürüm)"
)

# ============================================================
# AYARLAR VE SABİTLER
# ============================================================

LOOKBACK_CALENDAR_DAYS = 60
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 2

HTTP_TIMEOUT = 15
MAX_WORKERS = 3  # TEFAS IP bloklamasını önlemek için 3'e optimize edildi

REQUEST_MAX_RETRIES = 3
REQUEST_BACKOFF_FACTOR = 1.5

DEFAULT_MOMENTUM_WEIGHTS = {"return": 0.30, "sharpe": 0.25, "cumulative": 0.25, "drawdown": 0.20}
SECURITY_WEIGHTS = {"aum": 0.30, "investor": 0.25, "concentration": 0.25, "liquidity": 0.20}

DEFAULT_HYBRID_MOMENTUM_WEIGHT = 0.50
DEFAULT_HYBRID_SECURITY_WEIGHT = 0.35
DEFAULT_HYBRID_SENTIMENT_WEIGHT = 0.15

STRONG_BUY = 75
WATCH_LIST = 50
CORRECTION = 35

COLOR_NAVY, COLOR_GREEN, COLOR_RED, COLOR_YELLOW, COLOR_WHITE = "1F4E79", "008000", "FF0000", "B8860B", "FFFFFF"
COLOR_LIGHT_GREEN, COLOR_LIGHT_YELLOW, COLOR_LIGHT_RED = "E2F0D9", "FFF2CC", "FCE4D6"

# ============================================================
# HTTP OTURUMU & GÜÇLENDİRİLMİŞ HEADERS
# ============================================================

def build_http_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=REQUEST_MAX_RETRIES,
        connect=REQUEST_MAX_RETRIES,
        read=REQUEST_MAX_RETRIES,
        status=REQUEST_MAX_RETRIES,
        backoff_factor=REQUEST_BACKOFF_FACTOR,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"})
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=15, pool_maxsize=15)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

HTTP = build_http_session()

# ============================================================
# SIDEBAR VE KULLANICI PARAMETRELERİ
# ============================================================

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
    help="Google AI Studio API anahtarı.",
)

with st.sidebar.expander("⚖️ Skor Ağırlıkları"):
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
    w_hybrid_mom = st.slider("Momentum Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_MOMENTUM_WEIGHT, 0.05)
    w_hybrid_sec = st.slider("Güvenlik Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SECURITY_WEIGHT, 0.05)
    w_hybrid_sent = st.slider("Sentiment Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SENTIMENT_WEIGHT, 0.05)

    tot_h = w_hybrid_mom + w_hybrid_sec + w_hybrid_sent
    if tot_h <= 0: tot_h = 1.0

    HYBRID_MOMENTUM_WEIGHT = w_hybrid_mom / tot_h
    HYBRID_SECURITY_WEIGHT = w_hybrid_sec / tot_h
    HYBRID_SENTIMENT_WEIGHT = w_hybrid_sent / tot_h

RISK_FREE_ANNUAL = st.sidebar.number_input("Yıllık risksiz getiri (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
SHOW_DIAGNOSTICS = st.sidebar.checkbox("Kaynak tanılama bilgisini göster", value=True)

# ============================================================
# MATEMATİK VE SAYISAL PARSERLAR
# ============================================================

def clamp(value, low, high): 
    return max(low, min(high, value))

def safe_float(value, default=0.0) -> float:
    try:
        if value is None: return default
        n = float(value)
        return default if pd.isna(n) else n
    except Exception: return default

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

def normalize_fund_code(code) -> str:
    if code is None: return ""
    s = str(code).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

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
    except Exception: return None

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

def safe_std(values):
    vals = [optional_float(v) for v in values]
    vals = [v for v in vals if v is not None and math.isfinite(v)]
    if len(vals) < 2: return 0.0
    mean_v = sum(vals) / len(vals)
    return (sum((x - mean_v) ** 2 for x in vals) / len(vals)) ** 0.5

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
    prompt = f"""Sen kıdemli bir fon analistisin. Aşağıdaki fon yatırım alanları için 0-100 arası duyarlılık puanı ve max 6 kelimelik gerekçe üret:
{areas_text}
SADECE geçerli JSON objesi üret: {{"Alan Adı": {{"score": 75, "label": "Kısa gerekçe"}}}}"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key_clean}"
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
                        "ai_reason": "Bağlantı Başarılı"
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
# TEFAS & İŞ YATIRIM GÜÇLENDİRİLMİŞ VERİ ÇEKİCİLER
# ============================================================

def fetch_tefas_direct_api(fund_code: str):
    code = normalize_fund_code(fund_code)
    status = {"source": "TEFAS Direct API", "attempted": True, "ok": False, "status_code": None, "message": ""}
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Origin": "https://www.tefas.gov.tr",
        "Referer": "https://www.tefas.gov.tr/TarihselVeriler.aspx",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Connection": "keep-alive"
    }
    
    kinds = ["YAT", "EMK", "BYF", ""]
    for kind in kinds:
        payload = {
            "fontip": kind,
            "fonkod": code,
            "bastarih": start.strftime("%d.%m.%Y"),
            "bittarih": end.strftime("%d.%m.%Y")
        }
        try:
            res = HTTP.post(url, data=payload, headers=headers, timeout=HTTP_TIMEOUT)
            status["status_code"] = res.status_code
            if res.status_code == 200:
                resp_json = res.json()
                raw = resp_json.get("data", []) if isinstance(resp_json, dict) else []
                if raw and len(raw) >= 2:
                    df = pd.DataFrame(raw)
                    df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
                    df["price"] = df["FIYAT"].apply(parse_number)
                    df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number) if "PORTFOYBUYUKLUK" in df.columns else None
                    df["investors"] = df["KISISAYISI"].apply(parse_number) if "KISISAYISI" in df.columns else None
                    df = df.dropna(subset=["date", "price"])
                    df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                    if len(df) >= 2:
                        status["ok"] = True
                        status["message"] = "OK"
                        return df, status
        except Exception as exc:
            status["message"] = str(exc)[:100]

    return None, status

def fetch_isyatirim_series(fund_code: str):
    code = normalize_fund_code(fund_code)
    status = {"source": "İş Yatırım", "attempted": True, "ok": False, "status_code": None, "message": ""}
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json",
        "X-Requested-With": "XMLHttpRequest"
    }
    try:
        res = HTTP.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        status["status_code"] = res.status_code
        if res.status_code == 200:
            df = pd.DataFrame(res.json().get("value", []))
            if not df.empty and "Tarih" in df.columns and "Fiyat" in df.columns:
                df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
                df["price"] = df["Fiyat"].apply(parse_number)
                df["aum"], df["investors"] = None, None
                df = df.dropna(subset=["date", "price"])
                df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                if len(df) >= 2:
                    status["ok"] = True
                    status["message"] = "OK"
                    return df[["date", "price", "aum", "investors"]], status
    except Exception as exc:
        status["message"] = str(exc)[:100]

    return None, status

def get_fund_series(fund_code: str):
    code = normalize_fund_code(fund_code)
    statuses = []

    # 1. TEFAS Doğrudan API
    df_dir, stat_dir = fetch_tefas_direct_api(code)
    statuses.append(stat_dir)
    if df_dir is not None and len(df_dir) >= 2:
        return df_dir, "TEFAS Direct API", statuses

    # 2. İş Yatırım Fallback
    df_is, stat_is = fetch_isyatirim_series(code)
    statuses.append(stat_is)
    if df_is is not None and len(df_is) >= 2:
        return df_is, "İş Yatırım", statuses

    return None, "YOK", statuses

def fetch_fund_structural_data(fund_code: str) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {"top_asset_weight": None, "asset_class_hhi": None, "is_bist30": False, "investment_area": "Karma / Değişken"}
    
    if code in ["YAY", "AFT", "AFA", "TTE", "GUH", "ITP"]:
        structural["investment_area"] = "Hisse Senedi (Yabancı Teknoloji)"
    elif code in ["KZL", "GUM", "GGK", "KGM", "KUT", "AFO"]:
        structural["investment_area"] = "Kıymetli Maden"
    elif code in ["PPZ", "NVB", "NRC", "PNU", "TP2", "FIL"]:
        structural["investment_area"] = "Para Piyasası"
    elif code in ["MAC", "TI3", "TCD", "BIO", "THF", "KHA", "PUK", "AK3"]:
        structural["investment_area"] = "Hisse Senedi"
    elif code in ["DBH", "YBE", "AKE", "FUB"]:
        structural["investment_area"] = "Borçlanma Araçları"

    return structural

def compute_fund_metrics(series: pd.DataFrame, fund_code: str):
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

    struct = fetch_fund_structural_data(fund_code)
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
        "fund_title": fund_code,
        **struct,
    }

def fetch_and_compute_one_fund(code: str):
    time.sleep(0.05)  # Hızlı ardışık sorguları ve TEFAS rate limit'i önlemek için mikro bekleme
    series, source, statuses = get_fund_series(code)
    metrics = compute_fund_metrics(series, code)
    if metrics is None: return code, None, source, statuses
    metrics["source"] = source
    metrics["source_statuses"] = statuses
    return code, metrics, source, statuses

# ============================================================
# SKORLAMA VE METRİK MOTORU
# ============================================================

def calculate_security_scores(funds: List[dict]):
    for f in funds:
        score = 50.0
        aum = optional_float(f.get("aum"))
        if aum and aum > 500_000_000: score += 10.0
        if f.get("is_bist30"): score += 5.0
        f["security_score"] = int(round(clamp(score, 0.0, 100.0)))

def calculate_market_relative_momentum(funds: List[dict], window: int, risk_free_annual: float = 0.0):
    for f in funds:
        rets = [x for x in (f.get("orig_returns", [])[-window:]) if optional_float(x) is not None]
        prices = f.get("prices") or []
        prc = prices[-(len(rets) + 1):] if prices else []

        if len(rets) < 2:
            f.update({
                "market_momentum": 50,
                "_final_mean_return": 0.0,
                "_final_sharpe": 0.0,
                "_final_sortino": 0.0,
                "_final_calmar": 0.0,
                "_final_cumulative": 0.0,
                "_final_max_dd": 0.0,
                "annualized_volatility": 0.0,
            })
            continue

        m_r = sum(rets) / len(rets)
        vol_daily = safe_std(rets)
        cum = ((prc[-1] / prc[0]) - 1.0) * 100.0 if len(prc) >= 2 and prc[0] > 0 else calculate_compounded_return(rets)
        dd = calculate_max_drawdown(prc) if len(prc) >= 2 else 0.0

        rf_daily = ((1.0 + risk_free_annual / 100.0) ** (1.0 / 252.0) - 1.0) * 100.0
        excess = [v - rf_daily for v in rets]
        vol = safe_std(excess)
        sharpe = (sum(excess) / len(excess)) / vol * math.sqrt(252.0) if vol > 1e-12 else 0.0
        
        downside = [min(v, 0.0) for v in excess]
        downside_dev = (sum(x * x for x in downside) / len(downside)) ** 0.5
        sortino = (sum(excess) / len(excess)) / downside_dev * math.sqrt(252.0) if downside_dev > 1e-12 else 0.0
        calmar = (cum / abs(dd)) if abs(dd) > 1e-12 else 0.0

        f.update({
            "_final_mean_return": m_r,
            "_final_sharpe": sharpe,
            "_final_sortino": sortino,
            "_final_calmar": calmar,
            "_final_cumulative": cum,
            "_final_max_dd": dd,
            "annualized_volatility": vol_daily * math.sqrt(252.0),
            "market_momentum": int(round(clamp(50.0 + (m_r * 15.0), 10.0, 95.0))),
        })

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

        sec = safe_float(f.get("security_score"), 50.0)
        sent = clamp(safe_float(batch_sentiments.get(f.get("investment_area", "-"), {}).get("score", 50)), 0.0, 100.0)

        run_h = []
        for r in f["daily_returns"]:
            if r is None:
                run_h.append(None)
            else:
                m_score = clamp(50.0 + (r * 10.0), 0.0, 100.0)
                daily = m_score * HYBRID_MOMENTUM_WEIGHT + sec * HYBRID_SECURITY_WEIGHT + sent * HYBRID_SENTIMENT_WEIGHT
                run_h.append(int(round(clamp(daily, 0.0, 100.0))))

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
        mom = safe_float(f.get("market_momentum"), 50)
        sec = safe_float(f.get("security_score"), 50)
        sent_data = batch_sentiments.get(f.get("investment_area", "-"), {"score": 50, "label": "Nötr", "ai_active": False, "ai_reason": "Veri Yok"})
        sent = clamp(safe_float(sent_data.get("score"), 50), 0.0, 100.0)

        f["sentiment_score"] = sent
        f["sentiment_label"] = sent_data.get("label", "Nötr")
        f["sentiment_ai_active"] = sent_data.get("ai_active", False)
        f["sentiment_ai_reason"] = sent_data.get("ai_reason", "Bilinmiyor")

        dec = int(round(clamp(mom * HYBRID_MOMENTUM_WEIGHT + sec * HYBRID_SECURITY_WEIGHT + sent * HYBRID_SENTIMENT_WEIGHT, 0.0, 100.0)))
        f["decision_score"] = dec
        f["karar"] = decision_label_from_score(dec)
        f["data_quality_score"] = 95
        f["data_quality_issues"] = "OK"

# ============================================================
# GÜVENLİ EXCEL ÇIKTISI
# ============================================================

def create_excel_output(wb, all_funds, common_n_days):
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
        "Fon Kodu", "Fon Adı", "Yatırım Alanı", "Karar Skoru", "Trend Skoru",
        "Piyasa Momentum", "Güvenlik Skoru", "Sentiment Skoru", "Duyarlılık Yönü", "Model Kararı",
        "Ort. Günlük Getiri (%)", "Yıllıklandırılmış Volatilite (%)", "Sharpe", "Sortino", "Calmar",
        "Kümülatif Getiri (%)", "MaxDD (%)", "Haftalık Getiri (%)", "Veri Kaynağı"
    ]

    daily_headers = []
    for day in reversed(last_5_dates): 
        daily_headers.extend([f"{display_date(day)} Karar Skoru", f"{display_date(day)} Model Kararı"])
    headers[3:3] = daily_headers

    ws_scores.append(headers)
    header_index = {name: idx + 1 for idx, name in enumerate(headers)}

    fill_header = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    font_header = Font(name="Calibri", bold=True, color=COLOR_WHITE)
    for cell in ws_scores[1]:
        cell.fill, cell.font, cell.alignment = fill_header, font_header, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_scores.row_dimensions[1].height = 50

    for item in all_funds:
        row_data = [item["code"], item.get("fund_title") or item["code"], item.get("investment_area") or "-"]
        fund_dates = item.get("dates", [])
        fund_scores = item.get("running_trend_hybrid", [])
        score_map = dict(zip(fund_dates, fund_scores))

        for day in reversed(last_5_dates):
            s = score_map.get(day)
            row_data.extend([s if s is not None else "Veri Açıklanmadı", decision_label_from_score(s) if s is not None else "Veri Açıklanmadı"])

        row_data.extend([
            item.get("decision_score"), item.get("trend_skor"), item.get("market_momentum"),
            item.get("security_score"), item.get("sentiment_score"), item.get("sentiment_label"), item.get("karar", "-"),
            round(safe_float(item.get("_final_mean_return")), 4),
            round(safe_float(item.get("annualized_volatility")), 4),
            round(safe_float(item.get("_final_sharpe")), 4),
            round(safe_float(item.get("_final_sortino")), 4),
            round(safe_float(item.get("_final_calmar")), 4),
            round(safe_float(item.get("_final_cumulative")), 4),
            round(safe_float(item.get("_final_max_dd")), 4),
            round(safe_float(item.get("weekly_return")), 4), item.get("source", "-")
        ])

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
# ANA ARAYÜZ (STREAMLIT)
# ============================================================

if "req_codes" not in st.session_state: st.session_state["req_codes"] = []
if "wb_bytes" not in st.session_state: st.session_state["wb_bytes"] = None

st.markdown("### 📥 Veri Giriş Yöntemi Seçin")
input_method = st.radio(
    "Veri Kaynağı:",
    options=["✍️ Manuel Fon Girişi (+ / Virgül / Boşluk)", "📁 Bilgisayardan Excel Yükle"],
    horizontal=True,
    label_visibility="collapsed"
)

if input_method == "✍️ Manuel Fon Girişi (+ / Virgül / Boşluk)":
    st.info("💡 Fon kodlarını aralarına `+`, `,` veya boşluk koyarak yazabilirsiniz (Örn: `THF + KZL + PNU + ICH + KHA`).")
    with st.form("manual_entry_form"):
        manual_input = st.text_area("Analiz Edilecek Fon Kodları", value="THF + KZL + PNU + ICH + KHA")
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
        except Exception as exc:
            st.error(f"Excel yükleme hatası: {exc}")

req_codes = st.session_state.get("req_codes", [])
wb_bytes = st.session_state.get("wb_bytes")

if not req_codes or not wb_bytes:
    st.warning("⚠️ Lütfen analiz başlatmak için geçerli fon kodu girin veya dosya yükleyin.")
    st.stop()

wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))

st.write(f"🎯 **Analize Alınan Fonlar ({len(req_codes)} adet):** `{', '.join(req_codes)}`")

# ============================================================
# ANALİZ MOTORU & ÇALIŞTIRMA
# ============================================================

calc_funds, failed = [], []
prog = st.progress(0, "Veriler TEFAS ve İş Yatırım üzerinden alınıyor...")

with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as exe:
    futs = {exe.submit(fetch_and_compute_one_fund, c): c for c in req_codes}
    for i, fut in enumerate(concurrent.futures.as_completed(futs)):
        c = futs[fut]
        try: 
            _, met, src, statuses = fut.result()
        except Exception: 
            met = None
        if met: 
            calc_funds.append(met)
        else: 
            failed.append(c)
        prog.progress((i + 1) / len(req_codes))

prog.empty()

if not calc_funds:
    st.error(f"❌ Belirtilen fonlar için veri alınamadı. Hatalı Fonlar: {', '.join(failed)}")
    st.stop()

with st.spinner("📊 Model skorları ve duyarlılık hesaplanıyor..."):
    calculate_security_scores(calc_funds)
    calculate_market_relative_momentum(calc_funds, TARGET_TRADING_DAYS, RISK_FREE_ANNUAL)
    
    unique_areas = list(set([f.get("investment_area", "-") for f in calc_funds if f.get("investment_area")]))
    batch_sentiments = fetch_batch_market_sentiment(unique_areas or ["-"], api_key_input)
    
    common_n = calculate_trend_scores(calc_funds, batch_sentiments)
    finalize_decisions(calc_funds, batch_sentiments)

output = create_excel_output(wb, calc_funds, common_n)

# ============================================================
# SKOR ÖZETLERİ VE EKRAN TABLOSU
# ============================================================

st.subheader("📈 KAZRİSK Portföy Özeti (V14.4)")
col1, col2, col3, col4 = st.columns(4)
scores = [safe_float(x.get("decision_score")) for x in calc_funds if x.get("decision_score") is not None]
if scores:
    col1.metric("En Yüksek Skor", f"{max(scores):.0f}")
    col2.metric("Ortalama Skor", f"{sum(scores) / len(scores):.1f}")
    col3.metric("En Düşük Skor", f"{min(scores):.0f}")
    col4.metric("Güçlü Al Veren", sum(1 for x in calc_funds if x.get("karar") == "GÜÇLÜ AL"))

display_rows = []
early_alerts = []

all_dates_ui = set()
for f in calc_funds:
    for d in f.get("dates", []):
        if d is not None: all_dates_ui.add(d)

sorted_dates_ui = sorted(list(all_dates_ui))
sample_dates_ui = sorted_dates_ui[-common_n:] if common_n > 0 else sorted_dates_ui
last_5_dates_web = sample_dates_ui[-5:] if len(sample_dates_ui) >= 5 else sample_dates_ui

for item in calc_funds:
    row_dict = {
        "Fon Kodu": item["code"],
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
        "Haftalık Getiri (%)": round(safe_float(item.get("weekly_return")), 2),
        "Veri Kaynağı": item.get("source", "-")
    })
    display_rows.append(row_dict)

    valid_history = [(d, s) for d, s in zip(fund_dates, own_scores) if s is not None]
    if len(valid_history) >= 2:
        (d1, s1), (d2, s2) = valid_history[-2], valid_history[-1]
        lbl1, lbl2 = decision_label_from_score(s1), decision_label_from_score(s2)
        if lbl1 == "ACİL SAT" and lbl2 == "ACİL SAT":
            early_alerts.append({"Tip": "SAT", "Fon Kodu": item["code"], "Alan": item.get("investment_area"), "KAZRİSK Durumu": "🚨 2 GÜN TEYİTLİ ACİL SAT", "Son 2 Gün": f"{display_date(d1)} → {display_date(d2)}", "Son Skor": s2})
        elif lbl1 == "GÜÇLÜ AL" and lbl2 == "GÜÇLÜ AL":
            early_alerts.append({"Tip": "AL", "Fon Kodu": item["code"], "Alan": item.get("investment_area"), "KAZRİSK Durumu": "🚀 2 GÜN TEYİTLİ GÜÇLÜ AL", "Son 2 Gün": f"{display_date(d1)} → {display_date(d2)}", "Son Skor": s2})

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value).upper()
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text: return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text: return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text: return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except AttributeError: styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 Analiz Sonuçları — Son 5 İşlem Günü Kararları (V14.4)")
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

st.success(f"✅ V14.4 Analiz tamamlandı. Toplam {len(calc_funds)} fon işlendi.")
st.download_button(
    label="📥 KAZRİSK V14.4 Excel İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_FINAL_V14_4.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if SHOW_DIAGNOSTICS:
    st.subheader("🔎 Veri Kaynağı Tanılaması & Gemini AI Modu")
    diagnostic_rows = []
    for item in calc_funds:
        reason = item.get("sentiment_ai_reason", "Bilinmiyor")
        ai_status = "🟢 Aktif (Canlı API)" if item.get("sentiment_ai_active") else f"🔴 Pasif ({reason})"
        for status in item.get("source_statuses", []):
            diagnostic_rows.append({
                "Fon": item["code"],
                "Kaynak": status.get("source"),
                "Denendi": "Evet" if status.get("attempted") else "Hayır",
                "Başarılı": "Evet" if status.get("ok") else "Hayır",
                "HTTP": status.get("status_code"),
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
