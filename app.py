# ============================================================
# KGDM-3 & KAZRİSK - SÜRÜM V14.4 (AUTH & MODEL 404 YAMASI)
# ============================================================

import concurrent.futures
import datetime as dt
import io
import math
import json
import os
import random
import re
import time
from collections import defaultdict
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any

import openpyxl
import pandas as pd
import requests
import urllib3
import cloudscraper
import streamlit as st

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from google import genai
    from google.genai import types
    HAS_GOOGLE_GENAI = True
except ImportError:
    HAS_GOOGLE_GENAI = False

try:
    from tefas import Crawler as TefasCrawler
    HAS_TEFAS_CRAWLER = True
except ImportError:
    HAS_TEFAS_CRAWLER = False

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

st.set_page_config(page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi", page_icon="📊", layout="wide")
st.title("📊 KGDM-3 & KAZRİSK Hibrit Fon Analizi")
st.caption("İş Yatırım Çerez Korumalı Hat + Gemini 1.5 Flash | V14.4 Anti-Bot")

# ============================================================
# AYARLAR VE SABİTLER
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF", "KAT", "")
DEFAULT_FUND_KIND = "YAT"

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 5

HTTP_TIMEOUT = 18
MAX_WORKERS = 3

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

COLOR_NAVY, COLOR_GREEN, COLOR_RED, COLOR_YELLOW, COLOR_WHITE = "1F4E79", "008000", "FF0000", "B8860B", "FFFFFF"
COLOR_LIGHT_GREEN, COLOR_LIGHT_YELLOW, COLOR_LIGHT_RED = "E2F0D9", "FFF2CC", "FCE4D6"

# ============================================================
# SIDEBAR VE KULLANICI PARAMETRELERİ
# ============================================================

st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

env_api_key = os.environ.get("GEMINI_API_KEY", "")
try:
    if not env_api_key and "GEMINI_API_KEY" in st.secrets:
        env_api_key = st.secrets["GEMINI_API_KEY"]
except Exception: pass

api_key_input = st.sidebar.text_input("🔑 Gemini API Key (Canlı Sentiment)", value=env_api_key, type="password")

with st.sidebar.expander("⚖️ Skor Ağırlıkları"):
    w_return = st.slider("Getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_cumulative = st.slider("Kümülatif ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_drawdown = st.slider("Drawdown ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05)
    total_m = w_return + w_sharpe + w_cumulative + w_drawdown
    total_m = 1.0 if total_m <= 0 else total_m
    MOMENTUM_WEIGHTS = {"return": w_return / total_m, "sharpe": w_sharpe / total_m, "cumulative": w_cumulative / total_m, "drawdown": w_drawdown / total_m}

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
# YARDIMCI FONKSİYONLAR
# ============================================================
def clamp(value, low, high): return max(low, min(high, value))
def safe_float(value, default=0.0):
    try:
        if value is None: return default
        n = float(value)
        return default if pd.isna(n) else n
    except: return default
def optional_float(value):
    try:
        if value is None or (isinstance(value, str) and not value.strip()): return None
        n = float(value)
        return None if pd.isna(n) else n
    except: return None
def normalize_date_key(value):
    try:
        ts = pd.to_datetime(value, errors="coerce")
        return ts.strftime("%Y-%m-%d") if not pd.isna(ts) else None
    except: return None
def display_date(date_key):
    try: return pd.to_datetime(date_key).strftime("%d.%m.%Y")
    except: return str(date_key)
def parse_number(value):
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)): return None if pd.isna(value) else float(value)
    text = str(value).replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
    if not text: return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text: text = text.replace(",", ".")
    elif "." in text and rematch(r"^-?\d{1,3}(\.\d{3})+$", text): text = text.replace(".", "")
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

# ============================================================
# GEMINI SDK DUYARLILIK MOTORU (V14.4 404 YAMASI)
# ============================================================
@st.cache_data(ttl=60 * 60 * 4, show_spinner=False)
def fetch_batch_market_sentiment(areas: list, api_key: str) -> dict:
    result_map = {}
    api_key_clean = api_key.strip() if api_key else ""
    
    if not api_key_clean:
        for area in areas: result_map[area] = {"score": 50, "label": "Nötr (API Yok)", "ai_active": False, "ai_reason": "API Anahtarı Yok"}
        return result_map

    areas_text = "\n".join([f"- {a}" for a in areas])
    prompt = f"""Sen kıdemli bir fon yöneticisisin. Aşağıdaki alanlar için 0-100 arası duyarlılık puanı ve 6 kelimelik gerekçe üret:
{areas_text}
SADECE geçerli bir JSON objesi üret: {{"Alan Adı": {{"score": 75, "label": "Kısa gerekçe"}}}}"""

    last_err = ""
    # Model stabil sürüme çekildi (gemini-1.5-flash) 404 hatasını çözmek için
    target_model = 'gemini-1.5-flash'
    
    if HAS_GOOGLE_GENAI:
        try:
            client = genai.Client(api_key=api_key_clean)
            response = client.models.generate_content(
                model=target_model,
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2)
            )
            raw_text = response.text
            parsed_data = json.loads(raw_text.strip("```json\n").strip("```").strip())
            for area in areas:
                if area in parsed_data:
                    result_map[area] = {"score": int(clamp(safe_float(parsed_data[area].get("score", 50)), 0.0, 100.0)), "label": str(parsed_data[area].get("label", "Nötr")), "ai_active": True, "ai_reason": "google-genai Başarılı"}
                else:
                    result_map[area] = {"score": 50, "label": "Nötr", "ai_active": True, "ai_reason": "Alan bulunamadı"}
            return result_map
        except Exception as e:
            last_err = f"genai hatası: {str(e)[:40]}"
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{target_model}:generateContent?key={api_key_clean}"
    try:
        response = requests.post(url, headers={'Content-Type': 'application/json'}, json={"contents": [{"role": "user", "parts": [{"text": prompt}]}]}, timeout=15)
        if response.status_code == 200:
            raw_text = response.json().get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "{}")
            parsed_data = json.loads(raw_text.strip("```json\n").strip("
