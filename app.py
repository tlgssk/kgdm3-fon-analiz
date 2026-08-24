# ============================================================
# KGDM-3 & KAZRİSK - SÜRÜM V13.16 (GÜVENLİK YAMASI EKLENDİ)
# ============================================================

import concurrent.futures
import datetime as dt
import io
import math
import json
import os
import re
import time
from collections import defaultdict
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any

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
# STREAMLIT SAYFA YAPILANDIRMASI
# ============================================================

st.set_page_config(
    page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK Hibrit Fon Analizi")
st.caption(
    "TEFAS + İş Yatırım | Gemini 3.7 Sentiment (Toplu Sorgu) + Tam Senkron Tarihler | V13.16"
)

# ============================================================
# AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")
DEFAULT_FUND_KIND = "YAT"

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 5

HTTP_TIMEOUT = 20
MAX_WORKERS = 4

REQUEST_MAX_RETRIES = 2
REQUEST_BACKOFF_FACTOR = 1.5

MIN_REFERENCE_SAMPLE = 5
OVERHEAT_Z_THRESHOLD = 2.0
OVERHEAT_PENALTY = 6.0

APP_VERSION = "13.16.0"

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
    session.headers.update({"User-Agent": "KGDM3-Fon-Analiz/13.16", "Accept": "application/json,text/html"})
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

with st.sidebar.expander("⚖️ Skor Ağırlıkları (V13.16)"):
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
# CANLI GEMINI DUYARLILIK MOTORU - V13.16
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
        if pd.isna(ts):
            return None
        return ts.strftime("%Y-%m-%d")
    except Exception:
        return None

def display_date(date_key) -> str:
    try:
        return pd.to_datetime(date_key).strftime("%d.%m.%Y")
    except Exception:
        return str(date_key)

def is_valid_observation(value) -> bool:
    return optional_float(value) is not None

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
    
    prompt = f"""Sen kıdemli bir fon yöneticisi ve makroekonomik duyarlılık analistisin.
Aşağıdaki Türkiye TEFAS fon piyasasında yer alan yatırım alanlarının HER BİRİ için güncel piyasa duyarlılığını değerlendir:
{areas_text}

GÖREV:
Her bir varlık sınıfı için duyarlılığı 0-100 arası puanla:
0-35: Sert Düşüş / Satış Baskısı
36-49: Düzeltme / Belirsizlik
50-74: Pozitif / Dengeli Yükseliş
75-100: Güçlü Alım / Ralli
Her biri için maksimum 6 kelimelik kısa bir gerekçe etiketi üret.

Çıktı SADECE geçerli bir JSON objesi olmalıdır. Şema tam olarak şöyle olmalı:
{{
  "Alan 1 Adı": {{"score": 75, "label": "Kısa gerekçe"}},
  "Alan 2 Adı": {{"score": 40, "label": "Kısa gerekçe"}}
}}
"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.7-flash:generateContent?key={api_key_clean}"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json"}
    }

    last_err = ""
    for attempt in range(5):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=25)
            if response.status_code == 200:
                data = response.json()
                raw_text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "{}")
                parsed_data = json.loads(raw_text.strip("```json\n").strip("
