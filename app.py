import concurrent.futures
import datetime as dt
import io
import json
import os
import re
import statistics
import time
from collections import defaultdict
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Tuple, Any

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

# Google GenAI SDK (Opsiyonel)
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


# ============================================================
# KGDM-3 & KAZRİSK - SÜRÜM V10.3 (API FIX YAMASI)
# ============================================================

st.set_page_config(
    page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK Hibrit Fon Analizi")
st.caption(
    "TEFAS + TEFAS Direct API + İş Yatırım + Fintables | "
    "Gemini Canlı Sentiment + Evrensel Baseline + Kalite Denetimi | V10.3"
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

APP_VERSION = "10.3.0"

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

MAX_VALOR_PENALTY = 8.0
MAX_CONCENTRATION_PENALTY = 20.0

BIST30_BONUS = 5.0
HIGH_LIQUIDITY_BONUS = 5.0
LOW_LIQUIDITY_PENALTY = 3.0
POSITIVE_INVESTOR_FLOW_BONUS = 3.0

EMA_DECAY = 0.65

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

ENABLE_FILTERS = st.sidebar.checkbox("Filtreleri Etkinleştir", value=False)
TARGET_WEEKLY_RETURN = st.sidebar.slider("Hedef Haftalık Getiri (%)", -5.0, 10.0, 0.0, 0.10)
MIN_INVESTOR_COUNT = st.sidebar.slider("Minimum Yatırımcı Sayısı", 0, 100000, 0, 500)

with st.sidebar.expander("⚖️ Skor Ağırlıkları (V10.3)"):
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
# CANLI GEMINI DUYARLILIK (MARKET SENTIMENT) MOTORU - V10.3 API FIX
# ============================================================

@st.cache_data(ttl=60 * 60 * 4, show_spinner=False)
def fetch_market_sentiment(investment_area: str, api_key: str) -> dict:
    area = str(investment_area).strip()

    error_reason = ""
    if not api_key:
        error_reason = "API Anahtarı Girilmedi"
    elif not GENAI_AVAILABLE:
        error_reason = "google-genai Kütüphanesi Bulunamadı"

    if error_reason:
        area_upper = area.upper()
        if "YABANCI TEKNOLOJİ" in area_upper or "YABANCI" in area_upper:
            return {"score": 38, "label": "Negatif (Kâr Satışı)", "ai_active": False, "ai_reason": error_reason}
        elif "ALTIN" in area_upper or "GÜMÜŞ" in area_upper or "KIYMETLİ" in area_upper:
            return {"score": 82, "label": "Güçlü Pozitif (Faiz İndirimi)", "ai_active": False, "ai_reason": error_reason}
        elif "PARA PİYASASI" in area_upper or "BORÇLANMA" in area_upper:
            return {"score": 65, "label": "Pozitif (Sabit Getiri)", "ai_active": False, "ai_reason": error_reason}
        elif "HİSSE" in area_upper or "BIST" in area_upper:
            return {"score": 54, "label": "Dengeli / Pozitif Beklenti", "ai_active": False, "ai_reason": error_reason}
        return {"score": 50, "label": "Nötr / Kural Tabanlı", "ai_active": False, "ai_reason": error_reason}

    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
Sen kıdemli bir fon yöneticisi ve makroekonomik duyarlılık analistisin.
Türkiye TEFAS fon piyasasında yer alan şu yatırım alanı için güncel piyasa, haber ve ekonomist beklentilerini değerlendir:
Yatırım Alanı: "{area}"

GÖREV:
1. Bu varlık sınıfı için güncel piyasa duyarlılığını 0 ile 100 arasında bir tam sayı olarak puanla:
   - 0-35: Sert Düşüş / Satış Baskısı / Yüksek Risk
   - 36-49: Düzeltme / Belirsizlik / Yatay
   - 50-74: Pozitif / Dengeli Yükseliş / Talep Artışı
   - 75-100: Güçlü Alım / Ralli / Yoğun Pozitif Beklenti
2. En fazla 6 kelimelik kısa bir gerekçe etiketi üret.

Sadece ve sadece aşağıdaki JSON şemasında çıktı ver, format dışında hiçbir kelime yazma:
{{"score": 75, "label": "Gerekçe etiketi"}}
"""
        # Hata Fix: Stabil model (1.5-flash) seçildi ve JSON kısıtlaması esnetildi.
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.2)
        )
        
        # Gelen yanıtı güvenli bir şekilde temizleyip JSON'a çeviriyoruz
        raw_text = response.text.strip()
        if raw_text.startswith("```json"):
            raw_text = raw_text[7:]
        elif raw_text.startswith("```"):
            raw_text = raw_text[3:]
        if raw_text.endswith("```"):
            raw_text = raw_text[:-3]
            
        data = json.loads(raw_text.strip())
        
        return {
            "score": int(clamp(safe_float(data.get("score", 50)), 0.0, 100.0)),
            "label": str(data.get("label", "Nötr")),
            "ai_active": True,
            "ai_reason": "Bağlantı Başarılı"
        }
    except Exception as exc:
        return {
            "score": 50, 
            "label": "Nötr", 
            "ai_active": False, 
            "ai_reason": f"API Hatası: {str(exc)[:40]}"
        }

# ============================================================
# VERİ KALİTESİ + MATEMATİKSEL TUTARLILIK DENETİMİ
# ============================================================
PRICE_CONSISTENCY_TOLERANCE = 0.0005

def safe_float(value, default=0.0) -> float:
    try:
        if value is None: return default
        n = float(value)
        return default if pd.isna(n) else n
    except: return default

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

def clamp(value, low, high): return max(low, min(high, value))

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
    clean = [safe_float(v) for v in values if v is not None and not pd.isna(safe_float(v))]
    if len(clean) < 2: return [0.0] * len(values)
    mean_v = sum(clean) / len(clean)
    std = (sum((x - mean_v) ** 2 for x in clean) / len(clean)) ** 0.5
    if std <= 1e-12: return [0.0] * len(values)
    return [clamp((safe_float(v) - mean_v) / std, -Z_LIMIT, Z_LIMIT) if v is not None else 0.0 for v in values]

def population_mean_std(values):
    valid = [v for v in values if v is not None]
    if len(valid) < 2: return 0.0, 0.0
    mean_v = sum(valid) / len(valid)
    return mean_v, (sum((v - mean_v) ** 2 for v in valid) / len(valid)) ** 0.5

def zscore_against_population(value, mean_v, std_v):
    if value is None or std_v <= 1e-12: return 0.0
    return clamp((value - mean_v) / std_v, -Z_LIMIT, Z_LIMIT)

def calculate_valor_penalty(excess_valor):
    return clamp(safe_float(excess_valor) / 3.0, 0.0, 1.0) * MAX_VALOR_PENALTY

def validate_price_series(fund: dict) -> Dict[str, Any]:
    dates, prices, returns, issues = fund.get("dates") or [], fund.get("prices") or [], fund.get("daily_returns") or [], []
    if len(prices) < 2: issues.append("Yetersiz fiyat gözlemi")
    if dates and any(dates[i] >= dates[i + 1] for i in range(len(dates) - 1)): issues.append("Tarih sırası sorunu")
    if any((safe_float(p) is None or safe_float(p) <= 0) for p in prices): issues.append("Pozitif olmayan fiyat")
    if returns and len(returns) != max(0, len(prices) - 1): issues.append("Getiri-fiyat uyumsuzluğu")

    if len(prices) >= 2 and len(returns) >= len(prices) - 1:
        for i in range(len(prices) - 1):
            p0, p1, r = safe_float(prices[i]), safe_float(prices[i + 1]), safe_float(returns[i])
            if p0 > 0 and p1 > 0 and r is not None:
                expected = (p1 / p0 - 1.0) * 100.0
                if abs(expected - r) > (PRICE_CONSISTENCY_TOLERANCE * 100.0):
                    issues.append("Getiri-fiyat tutarsızlığı")
                    break
    return {"ok": not issues, "issues": issues}

def validate_structural_data(fund: dict) -> Dict[str, Any]:
    issues = []
    top_weight = fund.get("top_asset_weight")
    weights = [safe_float(top_weight)] if top_weight else []
    if fund.get("structural_fetch_ok") and not weights: issues.append("Yapısal kaynak başarılı fakat dağılım yok")
    return {"ok": not issues, "issues": issues, "hhi": fund.get("asset_class_hhi")}

def audit_fund_data(fund: dict) -> dict:
    price, structural = validate_price_series(fund), validate_structural_data(fund)
    score = 100.0
    if not price["ok"]: score -= 20
    if fund.get("n_days", 0) < TARGET_TRADING_DAYS: score -= 15
    if not fund.get("structural_fetch_ok", False): score -= 15
    if fund.get("aum") is None and fund.get("investors") is None: score -= 10
    if "Liste-bağıl" in str(fund.get("reference_scope", "")): score -= 10

    issues = price["issues"] + structural["issues"]
    if not fund.get("source"): issues.append("Fiyat kaynağı yok")
    if fund.get("structural_error"): issues.append(str(fund.get("structural_error")))

    fund["price_data_audit"] = price
    fund["structural_data_audit"] = structural
