# ============================================================
# KGDM-3 & KAZRİSK - SÜRÜM V16.1.5 (TAM VE HATASIZ MİMARİ)
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
import traceback
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
st.caption("TEFAS + İş Yatırım | Kararlı Çapraz AI + Anomali Koruması + ACİL SAT Alarmı | V16.1.5")

# ============================================================
# AYARLAR VE SABİTLER
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF", "KAT", "")
DEFAULT_FUND_KIND = "YAT"

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 5
MIN_REFERENCE_SAMPLE = 20

HTTP_TIMEOUT = 15
AI_API_TIMEOUT = 45
MAX_WORKERS = 3

OVERHEAT_Z_THRESHOLD = 2.0
OVERHEAT_PENALTY = 6.0

GITHUB_FALLBACK_URL = "https://github.com/tlgssk/kgdm3-fon-analiz/raw/refs/heads/main/Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"

DEFAULT_MOMENTUM_WEIGHTS = {"return": 0.30, "sharpe": 0.25, "cumulative": 0.25, "drawdown": 0.20}
MOMENTUM_WEIGHTS = DEFAULT_MOMENTUM_WEIGHTS.copy() # Hata Giderildi: Sabit buraya tanımlandı
SECURITY_SCALE = {"aum": 20.0, "investor": 20.0, "aum_flow": 8.0, "investor_change": 6.0}

Z_LIMIT = 2.5
STRONG_BUY = 75
WATCH_LIST = 50
CORRECTION = 35
BIST30_BONUS = 5.0
EMA_DECAY = 0.65

COLOR_NAVY, COLOR_GREEN, COLOR_RED, COLOR_YELLOW, COLOR_WHITE = "1F4E79", "008000", "FF0000", "B8860B", "FFFFFF"
COLOR_LIGHT_GREEN, COLOR_LIGHT_YELLOW, COLOR_LIGHT_RED = "E2F0D9", "FFF2CC", "FCE4D6"

AI_PROVIDERS = ["Gemini (Google)", "Groq (OpenAI OSS)", "OpenRouter (Ücretsiz Router)"]
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.6-flash")
GROQ_MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-120b")
GROQ_FALLBACK_MODEL = os.environ.get("GROQ_FALLBACK_MODEL", "llama-3.1-8b-instant")
OPENROUTER_MODEL = os.environ.get("OPENROUTER_MODEL", "openrouter/free")
AI_RETRY_COUNT = 2
AI_MAX_PROVIDER_DIFF = 25.0
AI_FUND_MAX_ITEMS = 10 
AI_FUND_RELIABLE_DIFF = 25.0

# ============================================================
# SIDEBAR VE KULLANICI PARAMETRELERİ
# ============================================================

st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

# Varsayılan (default) API anahtarları buraya doğrudan tanımlandı
DEFAULT_GEMINI_KEY = ""
DEFAULT_GROQ_KEY = ""

env_gemini_key = os.environ.get("GEMINI_API_KEY", "")
env_groq_key = os.environ.get("GROQ_API_KEY", "")
env_openrouter_key = os.environ.get("OPENROUTER_API_KEY", "")

try:
    if not env_gemini_key and "GEMINI_API_KEY" in st.secrets:
        env_gemini_key = st.secrets["GEMINI_API_KEY"]
    if not env_groq_key and "GROQ_API_KEY" in st.secrets:
        env_groq_key = st.secrets["GROQ_API_KEY"]
    if not env_openrouter_key and "OPENROUTER_API_KEY" in st.secrets:
        env_openrouter_key = st.secrets["OPENROUTER_API_KEY"]
except Exception: pass

# Eğer sistemde/secrets'ta kayıtlı özel bir anahtar yoksa, yukarıdaki default değerler otomatik yüklenecektir.
final_gemini_key = env_gemini_key if env_gemini_key else DEFAULT_GEMINI_KEY
final_groq_key = env_groq_key if env_groq_key else DEFAULT_GROQ_KEY

api_key_gemini = st.sidebar.text_input("🔑 Gemini API Key", value=final_gemini_key, type="password", help=f"Model: {GEMINI_MODEL}")
api_key_groq = st.sidebar.text_input("🔑 Groq API Key", value=final_groq_key, type="password", help=f"Model: {GROQ_MODEL}")
api_key_openrouter = st.sidebar.text_input("🔑 OpenRouter API Key", value=env_openrouter_key, type="password", help=f"Model: {OPENROUTER_MODEL}")

st.sidebar.markdown("---")
st.sidebar.markdown("🎯 **Yatırım Ufku (Vade Stratejisi)**")
yatirim_vadesi = st.sidebar.radio(
    "Vade Seçin",
    ["Kısa Vade (1-3 Ay - Agresif)", "Orta Vade (3-12 Ay - Dengeli)", "Uzun Vade (1 Yıl+ - Güvenli/Risk Yönetimi)"],
    index=1
)

if yatirim_vadesi.startswith("Kısa"):
    HYBRID_MOMENTUM_WEIGHT = 0.65
    HYBRID_SECURITY_WEIGHT = 0.25
    HYBRID_SENTIMENT_WEIGHT = 0.10
elif yatirim_vadesi.startswith("Uzun"):
    HYBRID_MOMENTUM_WEIGHT = 0.30
    HYBRID_SECURITY_WEIGHT = 0.55
    HYBRID_SENTIMENT_WEIGHT = 0.15
else:
    HYBRID_MOMENTUM_WEIGHT = 0.50
    HYBRID_SECURITY_WEIGHT = 0.35
    HYBRID_SENTIMENT_WEIGHT = 0.15

with st.sidebar.expander("⚖️ Gelişmiş Skor Ağırlıkları (Opsiyonel)"):
    w_return = st.slider("Getiri ağırlığı", 0.0, 1.0, MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe ağırlığı", 0.0, 1.0, MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_cumulative = st.slider("Kümülatif ağırlığı", 0.0, 1.0, MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_drawdown = st.slider("Drawdown ağırlığı", 0.0, 1.0, MOMENTUM_WEIGHTS["drawdown"], 0.05)
    total_m = w_return + w_sharpe + w_cumulative + w_drawdown
    total_m = 1.0 if total_m <= 0 else total_m
    MOMENTUM_WEIGHTS = {"return": w_return / total_m, "sharpe": w_sharpe / total_m, "cumulative": w_cumulative / total_m, "drawdown": w_drawdown / total_m}

SHOW_DIAGNOSTICS = st.sidebar.checkbox("Kaynak Tanılama & AI Loglarını Göster", value=True)
ANALYZE_ALL_FUNDS_AI = st.sidebar.checkbox("🤖 Tüm fonlarda fon-bazlı AI analizi", value=False, help="Kapalıyken sadece öne çıkan fonlar analiz edilir.")

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
    elif "." in text and re.match(r"^-?\d{1,3}(\.\d{3})+$", text): text = text.replace(".", "")
    try: return float(text)
    except: return None
def normalize_fund_code(value):
    code = str(value).strip().upper()
    return code[:-2] if code.endswith(".0") else code
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
def population_mean_std(values: list):
    valid = [v for v in values if v is not None]
    if len(valid) < 2: return 0.0, 1.0
    mean_v = sum(valid) / len(valid)
    std = (sum((x - mean_v) ** 2 for x in valid) / len(valid)) ** 0.5
    return mean_v, std if std > 1e-12 else 1.0
def zscore_against_population(val: float, pop_mean: float, pop_std: float) -> float:
    if val is None: return 0.0
    return clamp((val - pop_mean) / pop_std, -Z_LIMIT, Z_LIMIT)
def reference_sample_size(ref, kind): 
    return len(ref.get(kind, {}).get("mean_return", []))
def percentile_score(value, population, neutral=50.0) -> float:
    v = optional_float(value)
    vals = sorted([optional_float(x) for x in population if optional_float(x) is not None])
    if v is None or not vals: return neutral
    if len(vals) == 1: return neutral
    less = sum(x < v for x in vals)
    equal = sum(x == v for x in vals)
    pct = (less + 0.5 * equal) / len(vals)
    return clamp(pct * 100.0, 0.0, 100.0)

# ============================================================
# İSTATİSTİKSEL ANOMALİ TESPİTİ
# ============================================================
def detect_data_anomalies(prices: List[float], returns: List[float]) -> tuple:
    if not prices or len(prices) < 2: return False, ""
    if any(p <= 0 for p in prices if p is not None): 
        return True, "Negatif/sıfır fiyat"
    if any(r > 20.0 or r < -20.0 for r in returns if r is not None): 
        return True, "Günlük %20+ fiyat değişimi"
    identical_count = 0
    for i in range(1, len(prices)):
        if prices[i] == prices[i-1] and prices[i] is not None:
            identical_count += 1
            if identical_count >= 3: 
                return True, "Ardışık 4+ gün sabit fiyat"
        else:
            identical_count = 0
    return False, ""

# ============================================================
# ÇAPRAZ DOĞRULAMALI AI DUYARLILIK MOTORU
# ============================================================
def _extract_json_object(text: str) -> dict:
    if not text:
        raise ValueError("AI boş yanıt döndürdü.")
    text = str(text).strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.S)
        if not m:
            raise ValueError(f"Geçerli JSON bulunamadı. Gelen metin: {text[:200]}")
        return json.loads(m.group(0))

def _post_json_with_retry(url: str, headers: dict, data: dict) -> dict:
    last_error = None
    for attempt in range(AI_RETRY_COUNT):
        try:
            res = requests.post(url, headers=headers, json=data, timeout=AI_API_TIMEOUT)
            if res.status_code == 429:
                retry_after = res.headers.get("Retry-After")
                time.sleep(min(float(retry_after), 4.0) if retry_after else 2.0)
                last_error = requests.exceptions.HTTPError(f"429 RateLimitError: {res.text}")
                continue
            if 500 <= res.status_code < 600:
                time.sleep(1.0 * (attempt + 1))
                last_error = requests.exceptions.HTTPError(f"{res.status_code} ServerError: {res.text}")
                continue
            
            if res.status_code != 200:
                raise requests.exceptions.HTTPError(f"HTTP {res.status_code}: {res.text}")
                
            return res.json()
        except requests.exceptions.HTTPError as exc:
            last_error = exc
            if attempt < AI_RETRY_COUNT - 1:
                time.sleep(1.0 * (attempt + 1))
        except Exception as exc:
            last_error = exc
            if attempt < AI_RETRY_COUNT - 1:
                time.sleep(1.0 * (attempt + 1))
    raise last_error or RuntimeError("AI isteği başarısız.")

def call_gemini_direct(prompt: str, api_key: str, model: str = GEMINI_MODEL) -> dict:
    if not api_key or not api_key.strip():
        raise ValueError("Gemini API anahtarı yok.")
    url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + api_key.strip()
    headers = {"Content-Type": "application/json"}
    data = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1,
            "maxOutputTokens": 2048
        }
    }
    js = _post_json_with_retry(url, headers, data)
    candidates = js.get("candidates") or []
    if not candidates:
        raise ValueError(f"Gemini aday yanıt üretmedi: {js.get('promptFeedback', {})}")
    parts = candidates[0].get("content", {}).get("parts", [])
    text = "".join(str(x.get("text", "")) for x in parts if isinstance(x, dict))
    return _extract_json_object(text)

def call_groq_direct(prompt: str, api_key: str, model: str = GROQ_MODEL) -> dict:
    if not api_key or not api_key.strip():
        raise ValueError("Groq API anahtarı yok.")
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"}
    system = "You are a financial AI. You must strictly output ONLY valid JSON format."
    attempts = [model or GROQ_MODEL, GROQ_FALLBACK_MODEL]
    errors = []
    seen = set()
    
    for mdl in attempts:
        if not mdl or mdl in seen:
            continue
        seen.add(mdl)
        payload = {
            "model": mdl,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": prompt}
            ],
            "response_format": {"type": "json_object"},
            "temperature": 0.1,
            "max_completion_tokens": 1200
        }
        if mdl.startswith("openai/gpt-oss-"):
            payload["reasoning_effort"] = "low"
            
        try:
            js = _post_json_with_retry(url, headers, payload)
            text = js.get("choices", [{}])[0].get("message", {}).get("content", "")
            if not text:
                raise ValueError("Groq content boş döndü.")
            return _extract_json_object(text)
        except Exception as exc:
            if "reasoning_effort" in payload:
                payload.pop("reasoning_effort", None)
                try:
                    js = _post_json_with_retry(url, headers, payload)
                    text = js.get("choices", [{}])[0].get("message", {}).get("content", "")
                    if not text:
                        raise ValueError("Groq content boş döndü.")
                    return _extract_json_object(text)
                except Exception as exc2:
                    errors.append(f"{mdl} (fallback): {type(exc2).__name__}: {str(exc2)[:260]}")
                    continue
            else:
                errors.append(f"{mdl}: {type(exc).__name__}: {str(exc)[:260]}")
                continue
    raise RuntimeError("Groq tüm model denemeleri başarısız: " + " | ".join(errors))

def call_openrouter_direct(prompt: str, api_key: str, model: str = OPENROUTER_MODEL) -> dict:
    if not api_key or not api_key.strip():
        raise ValueError("OpenRouter API anahtarı yok.")
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
        "HTTP-Referer": "[https://github.com/tlgssk/kgdm3-fon-analiz](https://github.com/tlgssk/kgdm3-fon-analiz)",
        "X-Title": "KGDM-3 KAZRİSK"
    }
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a financial market sentiment classifier. Return ONLY valid JSON. Never invent data."},
            {"role": "user", "content": prompt}
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.1,
        "max_tokens": 2048
    }
    js = _post_json_with_retry(url, headers, data)
    text = js.get("choices", [{}])[0].get("message", {}).get("content", "")
    return _extract_json_object(text)

def _normalize_ai_result(raw: dict, areas: list) -> dict:
    out = {}
    if not isinstance(raw, dict): return out
    raw_map = {str(k).strip().casefold(): v for k, v in raw.items()}
    def norm_key(x): return re.sub(r"\s+", " ", str(x).strip()).casefold()
    for area in areas:
        item = raw.get(area)
        if item is None: item = raw_map.get(norm_key(area))
        if item is None and len(raw) == len(areas):
            try: item = list(raw.values())[list(areas).index(area)]
            except: item = None
        if not isinstance(item, dict): continue
        score = optional_float(item.get("score"))
        if score is None: continue
        out[area] = {
            "score": int(round(clamp(score, 0, 100))),
            "label": str(item.get("label", "Nötr"))[:120]
        }
    return out

@st.cache_data(ttl=60 * 60 * 4, show_spinner=False)
def fetch_batch_market_sentiment_cross_validated(areas: list, key_gemini: str, key_groq: str, key_openrouter: str) -> dict:
    result_map = {}
    areas = list(dict.fromkeys([str(a) for a in areas if str(a).strip()]))
    if not areas: return {}

    areas_text = "\n".join([f"- {a}" for a in areas])
    prompt = f"""Sen kıdemli bir portföy risk analistisin.
Aşağıdaki yatırım alanlarının mevcut piyasa duyarlılığını 0-100 arasında puanla.
SADECE verilen alanları kullan. SADECE şu JSON yapısını döndür:
{{"Alan Adı": {{"score": 75, "label": "Kısa gerekçe"}}}}

Alanlar:
{areas_text}"""

    providers = [
        ("Gemini", key_gemini, call_gemini_direct, GEMINI_MODEL),
        ("Groq", key_groq, call_groq_direct, GROQ_MODEL),
        ("OpenRouter", key_openrouter, call_openrouter_direct, OPENROUTER_MODEL),
    ]
    active = [(name, key, fn, model) for name, key, fn, model in providers if key and key.strip()]

    if not active:
        for area in areas:
            result_map[area] = {"score": 50, "label": "Nötr (AI Yok)", "ai_active": False, "reliable": False, "ai_reason": "API anahtarı yok.", "ai_providers": []}
        return result_map

    provider_results = {}
    provider_errors = {}

    def run_provider(item):
        name, key, fn, model = item
        try:
            return name, _normalize_ai_result(fn(prompt, key, model), areas), ""
        except Exception as exc:
            return name, {}, f"{type(exc).__name__}: {str(exc)[:260]}"

    with concurrent.futures.ThreadPoolExecutor(max_workers=min(3, len(active))) as pool:
        futures = [pool.submit(run_provider, item) for item in active]
        for fut in concurrent.futures.as_completed(futures):
            name, data, err = fut.result()
            if data: provider_results[name] = data
            if err: provider_errors[name] = err

    for area in areas:
        observations = []
        for name, data in provider_results.items():
            if area in data:
                observations.append((name, safe_float(data[area]["score"]), data[area]["label"]))

        if not observations:
            errors = " | ".join(f"{k}: {v}" for k, v in provider_errors.items())
            result_map[area] = {"score": 50, "label": "Nötr (Hata)", "ai_active": False, "reliable": False, "ai_reason": errors, "ai_providers": []}
            continue

        scores = [x[1] for x in observations]
        labels = [x[2] for x in observations]
        final_score = int(round(float(pd.Series(scores).median())))
        spread = max(scores) - min(scores) if len(scores) > 1 else 999.0
        reliable = len(observations) >= 2 and spread <= AI_MAX_PROVIDER_DIFF

        if len(observations) >= 2: reason = f"Çapraz doğrulama: {len(observations)} AI, yayılım {spread:.0f}"
        else: reason = f"Tek AI: {observations[0][0]}"
        if provider_errors: reason += " | Hata: " + "; ".join(provider_errors.keys())

        result_map[area] = {
            "score": clamp(final_score, 0, 100), "label": labels[0] if labels else "Nötr",
            "ai_active": True, "reliable": reliable, "ai_reason": reason,
            "ai_providers": [x[0] for x in observations], "ai_scores": {x[0]: x[1] for x in observations}
        }
    return result_map

# ============================================================
# MULTI-CRAWLER VERİ ÇEKME MOTORU
# ============================================================
@dataclass
class SourceStatus:
    source: str; attempted: bool = False; ok: bool = False; status_code: Optional[int] = None
    error_type: str = ""; message: str = ""; elapsed_ms: Optional[int] = None; retry_count: int = 0
    is_synthetic: bool = False

def fetch_tier1_tefascrawler(code: str, start, end):
    t0 = time.time()
    status = SourceStatus("1. Hat: tefas-crawler", attempted=True)
    if not HAS_TEFAS_CRAWLER:
        status.message = "tefas kütüphanesi yüklü değil"
        return None, status
    try:
        crawler = TefasCrawler()
        df_raw = crawler.fetch(start=start.strftime("%Y-%m-%d"), end=end.strftime("%Y-%m-%d"), name=code)
        if df_raw is not None and not df_raw.empty:
            df = df_raw.copy()
            date_col = next((c for c in ["date", "TARIH", "Tarih"] if c in df.columns), None)
            price_col = next((c for c in ["price", "FIYAT", "Fiyat"] if c in df.columns), None)
            if date_col and price_col:
                df["date"] = pd.to_datetime(df[date_col], errors="coerce")
                df["price"] = df[price_col].apply(parse_number)
                df["aum"] = df["market_cap"].apply(parse_number) if "market_cap" in df.columns else None
                df["investors"] = df["number_of_investors"].apply(parse_number) if "number_of_investors" in df.columns else None
                df = df.dropna(subset=["date", "price"])[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                if len(df) >= 2:
                    status.ok = True; status.status_code = 200; status.message = f"Başarılı ({len(df)} gün)"
                    status.elapsed_ms = int((time.time() - t0) * 1000)
                    return df, status
        status.message = "Boş liste"
    except Exception as e: 
        status.message = str(e)[:50]
    return None, status

def fetch_tier2_cloudscraper(code: str, start, end):
    t0 = time.time()
    status = SourceStatus("2. Hat: Cloudscraper WAF Bypass", attempted=True)
    try:
        scraper = cloudscraper.create_scraper(browser={'browser': 'chrome', 'platform': 'windows', 'mobile': False})
        url = "[https://www.tefas.gov.tr/api/DB/BindComparisonFundReturns](https://www.tefas.gov.tr/api/DB/BindComparisonFundReturns)"
        payload = {"calismatipi": "2", "bastarih": start.strftime("%d.%m.%Y"), "bittarih": end.strftime("%d.%m.%Y"), "fonkod": code}
        headers = {"X-Requested-With": "XMLHttpRequest", "Origin": "[https://www.tefas.gov.tr](https://www.tefas.gov.tr)"}
        res = scraper.post(url, data=payload, headers=headers, timeout=HTTP_TIMEOUT)
        status.status_code = res.status_code
        status.elapsed_ms = int((time.time() - t0) * 1000)
        if res.status_code == 200 and res.text.strip():
            raw_data = res.json().get("data", [])
            if raw_data:
                df = pd.DataFrame(raw_data)
                df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
                df["price"] = df["FIYAT"].apply(parse_number)
                df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number) if "PORTFOYBUYUKLUK" in df.columns else None
                df["investors"] = df["KISISAYISI"].apply(parse_number) if "KISISAYISI" in df.columns else None
                df = df.dropna(subset=["date", "price"])[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                if len(df) >= 2:
                    status.ok = True; status.message = f"Başarılı ({len(df)} gün)"
                    return df, status
        status.message = f"HTTP {res.status_code} veya Boş"
    except Exception as e: 
        status.message = str(e)[:50]
    return None, status

def fetch_tier3_isyatirim(code: str, start, end):
    t0 = time.time()
    status = SourceStatus("3. Hat: İş Yatırım", attempted=True)
    url = "[https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri](https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri)"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    try:
        session = requests.Session()
        res = session.get(url, params=params, headers={"Accept": "application/json", "X-Requested-With": "XMLHttpRequest"}, timeout=HTTP_TIMEOUT)
        status.status_code = res.status_code
        status.elapsed_ms = int((time.time() - t0) * 1000)
        if res.status_code == 200:
            df = pd.DataFrame(res.json().get("value", []))
            if not df.empty:
                df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
                df["price"] = df["Fiyat"].apply(parse_number)
                df["aum"], df["investors"] = None, None
                df = df.dropna(subset=["date", "price"])[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
                if len(df) >= 2:
                    status.ok = True; status.message = f"Başarılı ({len(df)} gün)"
                    return df, status
        status.message = f"HTTP {res.status_code}"
    except Exception as e: 
        status.message = str(e)[:50]
    return None, status

def fetch_tier4_fallback(code: str):
    status = SourceStatus("4. Hat: Smart Fallback", attempted=True, ok=False, status_code=200, message="SENTETİK VERİ KULLANILDI", is_synthetic=True)
    end = dt.datetime.now()
    dates = pd.bdate_range(end=end, periods=TARGET_TRADING_DAYS + 5)
    drift = 0.0015
    prices = [100.0]
    for _ in range(1, len(dates)): prices.append(prices[-1] * (1.0 + drift + random.uniform(-0.004, 0.004)))
    df = pd.DataFrame({"date": dates, "price": prices, "aum": [1_500_000_000]*len(dates), "investors": [12500]*len(dates)})
    return df, status

def get_fund_series(fund_code: str):
    code = normalize_fund_code(fund_code)
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    statuses = []

    df1, s1 = fetch_tier1_tefascrawler(code, start, end)
    statuses.append(s1)
    if df1 is not None: return df1, s1.source, statuses

    df2, s2 = fetch_tier2_cloudscraper(code, start, end)
    statuses.append(s2)
    if df2 is not None: return df2, s2.source, statuses

    df3, s3 = fetch_tier3_isyatirim(code, start, end)
    statuses.append(s3)
    if df3 is not None: return df3, s3.source, statuses

    df4, s4 = fetch_tier4_fallback(code)
    statuses.append(s4)
    return df4, s4.source, statuses

# ============================================================
# VERİ HAZIRLAMA VE HESAPLAMALAR
# ============================================================
@st.cache_data(show_spinner=False, ttl=60 * 60)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
        df = Crawler(timeout=30).fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is not None and not df.empty:
            df.rename(columns={"fund_code": "code", "fund_name": "title", "investor_count": "investors", "portfolio_size": "aum", "fund_type": "kind"}, inplace=True)
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
            df["price"] = df["price"].apply(parse_number)
            df["code"] = df["code"].astype(str).str.strip().str.upper()
            return df.dropna(subset=["date", "code", "price"])[df["price"] > 0].sort_values(["code", "date"]).drop_duplicates(subset=["code", "date"], keep="last").reset_index(drop=True)
    except Exception: pass
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
    latest = universe.sort_values("date").drop_duplicates(subset=["code"], keep="last")
    for _, row in latest.iterrows():
        k_str = str(row.get("kind", DEFAULT_FUND_KIND)).strip().upper()
        if k_str in ref:
            if safe_float(row.get("aum")) > 0: ref[k_str]["aum"].append(safe_float(row.get("aum")))
            if safe_float(row.get("investors")) > 0: ref[k_str]["investors"].append(safe_float(row.get("investors")))

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

def fetch_fund_structural_data(fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {"is_bist30": False, "investment_area": "-"}
    t_upper = (fund_title or "").upper()
    if "PARA PİYASASI" in t_upper or "PPF" in t_upper: structural["investment_area"] = "Para Piyasası"
    elif "ALTIN" in t_upper or "GÜMÜŞ" in t_upper or "KIYMETLİ" in t_upper: structural["investment_area"] = "Kıymetli Maden"
    elif "YABANCI TEKNOLOJİ" in t_upper: structural["investment_area"] = "Hisse Senedi (Yabancı Teknoloji)"
    elif "HİSSE" in t_upper: structural["investment_area"] = "Hisse Senedi"
    elif "BORÇLANMA" in t_upper: structural["investment_area"] = "Borçlanma Araçları"
    elif "DEĞİŞKEN" in t_upper or "KARMA" in t_upper: structural["investment_area"] = "Karma / Değişken"
    if "BIST 30" in t_upper or "BIST30" in t_upper: structural["is_bist30"] = True
    return structural

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
        if prices[i - 1] and prices[i - 1] > 0 and prices[i] and prices[i] > 0:
            rets.append((prices[i] / prices[i - 1] - 1.0) * 100.0)
            return_dates.append(date_keys_all[i])

    if not rets: return None

    has_anomaly, anomaly_reason = detect_data_anomalies(prices, rets)

    struct = fetch_fund_structural_data(fund_code, fund_kind, fund_title)
    aum_last = next((v for v in reversed(aums) if v is not None and v > 0), None)
    inv_last = next((v for v in reversed(invs) if v is not None and v >= 0), None)
    aum_first = next((v for v in aums if v is not None and v > 0), None)
    inv_first = next((v for v in invs if v is not None and v > 0), None)

    aum_change = ((aum_last / aum_first) - 1.0) * 100.0 if aum_last and aum_first else None
    inv_change = ((inv_last / inv_first) - 1.0) * 100.0 if inv_last is not None and inv_first else None
    price_cum = ((prices[-1] / prices[0]) - 1.0) * 100.0 if prices[0] > 0 else None
    aum_flow_proxy = (aum_change - price_cum) if aum_change is not None and price_cum is not None else None

    return {
        "code": fund_code, "dates": return_dates, "prices": prices, "price_dates": date_keys_all,
        "price_map": dict(zip(date_keys_all, prices)), "daily_returns": rets, "n_days": len(rets),
        "aum": aum_last, "investors": int(inv_last) if inv_last is not None else None,
        "aum_change": aum_change, "aum_flow_proxy": aum_flow_proxy, "inv_change": inv_change,
        "max_dd": calculate_max_drawdown(prices), "weekly_return": calculate_compounded_return(rets[-5:]),
        "fund_title": fund_title or "-", 
        "has_anomaly": has_anomaly, "anomaly_reason": anomaly_reason,
        **struct,
    }

def fetch_and_compute_one_fund(code: str, meta_map: dict):
    meta = meta_map.get(code, {})
    series, source, statuses = get_fund_series(code)
    metrics = compute_fund_metrics(series, code, meta.get("kind"), meta.get("title"))
    if metrics is None: return code, None, source
    metrics["source"] = source
    metrics["is_synthetic"] = any(s.is_synthetic for s in statuses if s.source == source)
    metrics["kind"] = meta.get("kind", DEFAULT_FUND_KIND)
    metrics["source_statuses"] = [asdict(x) for x in statuses]
    return code, metrics, source

def calculate_security_scores(funds: List[dict], reference: dict):
    by_kind = defaultdict(list)
    for idx, fund in enumerate(funds): by_kind[fund.get("kind", DEFAULT_FUND_KIND)].append(idx)

    for kind, indices in by_kind.items():
        subset = [funds[i] for i in indices]
        ref = reference.get(kind, {})

        aum_ref = [safe_float(x) for x in ref.get("aum", []) if optional_float(x) is not None and safe_float(x) > 0]
        inv_ref = [safe_float(x) for x in ref.get("investors", []) if optional_float(x) is not None and safe_float(x) > 0]

        flow_z = zscore([f.get("aum_flow_proxy") for f in subset])
        inv_c_z = zscore([f.get("inv_change") for f in subset])

        for local_i, f in enumerate(subset):
            aum, investors = optional_float(f.get("aum")), optional_float(f.get("investors"))
            aum_pop = [math.log1p(x) for x in aum_ref]
            inv_pop = [math.log1p(x) for x in inv_ref]
            aum_pct = percentile_score(math.log1p(aum) if aum and aum > 0 else None, aum_pop)
            inv_pct = percentile_score(math.log1p(investors) if investors and investors > 0 else None, inv_pop)

            s = 50.0 + (aum_pct - 50.0) * 0.22 + (inv_pct - 50.0) * 0.18
            if f.get("aum_flow_proxy") is not None and flow_z[local_i] is not None: s += SECURITY_SCALE["aum_flow"] * flow_z[local_i]
            if f.get("inv_change") is not None and inv_c_z[local_i] is not None: s += SECURITY_SCALE["investor_change"] * inv_c_z[local_i]
            if f.get("is_bist30", False): s += BIST30_BONUS
            
            f["security_score"] = int(round(clamp(s, 0.0, 100.0)))

def calculate_market_relative_momentum(funds: List[dict], reference, window: int):
    for f in funds:
        k = f.get("kind", DEFAULT_FUND_KIND)
        rets = f.get("daily_returns", [])[-window:]
        prc = f.get("prices", [])[-(window + 1):]
        if len(rets) < MIN_ROLLING_DAYS or len(prc) < MIN_ROLLING_DAYS + 1:
            f["market_momentum"] = None
            continue

        m_r = sum(rets) / len(rets)
        vol = (sum((x - m_r) ** 2 for x in rets) / len(rets)) ** 0.5
        cum = (prc[-1] / prc[0] - 1.0) * 100.0 if prc[0] > 0 else 0.0
        dd = calculate_max_drawdown(prc)

        f["_final_mean_return"], f["_final_sharpe"], f["_final_cumulative"], f["_final_max_dd"], f["volatility"] = m_r, m_r / vol if vol > 1e-12 else 0.0, cum, dd, vol

        if reference_sample_size(reference, k) >= MIN_REFERENCE_SAMPLE:
            ref = reference[k]
            mm, ms = population_mean_std(ref["mean_return"])
            sm, ss = population_mean_std(ref["sharpe"])
            cm, cs = population_mean_std(ref["cumulative"])
            dm, ds = population_mean_std(ref["max_dd_inv"])

            zm = zscore_against_population(m_r, mm, ms)
            zs = zscore_against_population(f["_final_sharpe"], sm, ss)
            zc = zscore_against_population(cum, cm, cs)
            zd = zscore_against_population(-dd, dm, ds)
            f["reference_scope"] = f"Piyasa Geneli ({k})"
        else:
            fb = [x for x in funds if x.get("kind") == k and x.get("_final_mean_return") is not None]
            idx = next((i for i, x in enumerate(fb) if x is f), 0)
            zm = zscore([x.get("_final_mean_return") for x in fb])[idx] if fb else 0.0
            zs = zscore([x.get("_final_sharpe") for x in fb])[idx] if fb else 0.0
            zc = zscore([x.get("_final_cumulative") for x in fb])[idx] if fb else 0.0
            zd = zscore([-safe_float(x.get("_final_max_dd")) for x in fb])[idx] if fb else 0.0
            f["reference_scope"] = "Liste-Bağıl (Yetersiz Evren)"

        wz = (MOMENTUM_WEIGHTS["return"] * zm + MOMENTUM_WEIGHTS["sharpe"] * zs + MOMENTUM_WEIGHTS["cumulative"] * zc + MOMENTUM_WEIGHTS["drawdown"] * zd)
        mom = clamp(50.0 + 20.0 * wz, 0.0, 100.0)

        last_d = rets[-1]
        last_2 = sum(rets[-2:]) / 2.0 if len(rets) >= 2 else last_d
        oh = zc >= OVERHEAT_Z_THRESHOLD and (last_d < 0 or last_2 < 0)
        f["overheat_flag"] = oh
        if oh: mom = clamp(mom - OVERHEAT_PENALTY, 0.0, 100.0)
        f["market_momentum"] = int(round(mom))

def calculate_trend_scores(funds: List[dict], batch_sentiments: dict) -> int:
    if not funds: return 0
    all_dates = set()
    for f in funds: all_dates.update(f.get("dates", []))
        
    master_dates = sorted(list(all_dates))
    if len(master_dates) < MIN_ROLLING_DAYS: return 0
    master_dates = master_dates[-TARGET_TRADING_DAYS:]
    
    for f in funds:
        ret_map = dict(zip(f.get("dates", []), f.get("daily_returns", [])))
        f["dates"] = master_dates
        f["daily_returns"] = [ret_map.get(d) for d in master_dates]
        f["n_days"] = len([r for r in f["daily_returns"] if r is not None])

        pmap = f.get("price_map", {})
        f["prices"] = [pmap.get(d) for d in master_dates]
        f["running_trend_momentum"] = []

    for end_idx, day in enumerate(master_dates):
        if end_idx + 1 < MIN_ROLLING_DAYS:
            for f in funds: f["running_trend_momentum"].append(None)
            continue

        cur = []
        window_start = end_idx + 1 - MIN_ROLLING_DAYS
        for f in funds:
            r_raw = f["daily_returns"][window_start:end_idx + 1]
            r = [x for x in r_raw if x is not None]
            if len(r) < MIN_ROLLING_DAYS: continue
                
            pmap = f.get("price_map", {})
            all_pd = sorted(pmap.keys())
            first_day = master_dates[window_start]
            
            try:
                pidx = all_pd.index(first_day)
                prev = all_pd[pidx - 1] if pidx > 0 else None
            except:
                prev_candidates = [k for k in all_pd if k < first_day]
                prev = prev_candidates[-1] if prev_candidates else None
                
            p_window_dates = ([prev] if prev else []) + master_dates[window_start:end_idx + 1]
            p = [pmap[d] for d in p_window_dates if d in pmap]
            
            if len(p) < MIN_ROLLING_DAYS + 1: continue
                
            mr = sum(r) / len(r)
            vol = (sum((x - mr) ** 2 for x in r) / len(r)) ** 0.5
            cur.append({
                "fund": f, "mr": mr, "sh": mr / vol if vol > 1e-12 else 0.0,
                "cm": calculate_compounded_return(r), "dd": calculate_max_drawdown(p),
            })

        if not cur:
            for f in funds: f["running_trend_momentum"].append(None)
            continue

        zm = zscore([x["mr"] for x in cur])
        zs = zscore([x["sh"] for x in cur])
        zc = zscore([x["cm"] for x in cur])
        zd = zscore([-x["dd"] for x in cur])

        score_by_id = {}
        for i, data in enumerate(cur):
            wz = (MOMENTUM_WEIGHTS["return"] * zm[i] + MOMENTUM_WEIGHTS["sharpe"] * zs[i] + MOMENTUM_WEIGHTS["cumulative"] * zc[i] + MOMENTUM_WEIGHTS["drawdown"] * zd[i])
            score_by_id[id(data["fund"])] = int(round(clamp(50.0 + 20.0 * wz, 0.0, 100.0)))

        for f in funds: f["running_trend_momentum"].append(score_by_id.get(id(f)))

    for f in funds:
        sec = safe_float(f.get("security_score"), 50.0)
        sent_data = batch_sentiments.get(f.get("investment_area", "-"), {"score": 50, "label": "Nötr"})
        sent = clamp(safe_float(sent_data.get("score"), 50.0), 0.0, 100.0)
        
        if not sent_data.get("ai_active"):
            eff_mom_weight = HYBRID_MOMENTUM_WEIGHT / (HYBRID_MOMENTUM_WEIGHT + HYBRID_SECURITY_WEIGHT)
            eff_sec_weight = HYBRID_SECURITY_WEIGHT / (HYBRID_MOMENTUM_WEIGHT + HYBRID_SECURITY_WEIGHT)
            eff_sent_weight = 0.0
        elif not sent_data.get("reliable", True):
            eff_sent_weight = HYBRID_SENTIMENT_WEIGHT / 2.0
            eff_mom_weight = HYBRID_MOMENTUM_WEIGHT + (HYBRID_SENTIMENT_WEIGHT / 4.0)
            eff_sec_weight = HYBRID_SECURITY_WEIGHT + (HYBRID_SENTIMENT_WEIGHT / 4.0)
        else:
            eff_mom_weight = HYBRID_MOMENTUM_WEIGHT
            eff_sec_weight = HYBRID_SECURITY_WEIGHT
            eff_sent_weight = HYBRID_SENTIMENT_WEIGHT

        run_h = []
        for m in f["running_trend_momentum"]:
            if m is None: run_h.append(None)
            else: run_h.append(int(round(clamp(m * eff_mom_weight + sec * eff_sec_weight + sent * eff_sent_weight, 0.0, 100.0))))

        f["running_trend_hybrid"] = run_h
        valid = [s for s in run_h if s is not None]
        val_l = valid[-5:]
        
        if val_l:
            weights = [EMA_DECAY ** (len(val_l) - 1 - i) for i in range(len(val_l))]
            f["trend_skor"] = int(round(sum(s * w for s, w in zip(val_l, weights)) / sum(weights)))
            mean_score = sum(val_l) / len(val_l)
            std_score = (sum((s - mean_score)**2 for s in val_l) / len(val_l))**0.5
            f["trend_std"] = round(std_score, 1)
        else: 
            f["trend_skor"] = None
            f["trend_std"] = None

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
        if f.get("is_synthetic", False):
            f["decision_score"], f["karar"] = None, "YETERSİZ VERİ (SENTETİK)"
            continue
            
        mom = f.get("market_momentum")
        sec = f.get("security_score")
        if mom is None or sec is None:
            f["decision_score"], f["karar"] = None, "YETERSİZ VERİ"
            continue

        sent_data = batch_sentiments.get(f.get("investment_area", "-"), {"score": 50, "label": "Nötr"})
        sent = sent_data["score"]
        f["sentiment_score"] = sent
        f["sentiment_label"] = sent_data["label"]
        f["sentiment_ai_active"] = sent_data.get("ai_active", False)
        f["sentiment_ai_reason"] = sent_data.get("ai_reason", "Bilinmiyor")
        f["sentiment_ai_providers"] = sent_data.get("ai_providers", [])
        f["sentiment_ai_reliable"] = bool(sent_data.get("reliable", False))
        f["sentiment_ai_scores"] = sent_data.get("ai_scores", {})

        if not sent_data.get("ai_active"):
            eff_mom_weight = HYBRID_MOMENTUM_WEIGHT / (HYBRID_MOMENTUM_WEIGHT + HYBRID_SECURITY_WEIGHT)
            eff_sec_weight = HYBRID_SECURITY_WEIGHT / (HYBRID_MOMENTUM_WEIGHT + HYBRID_SECURITY_WEIGHT)
            eff_sent_weight = 0.0
        elif not sent_data.get("reliable", True):
            eff_sent_weight = HYBRID_SENTIMENT_WEIGHT / 2.0
            eff_mom_weight = HYBRID_MOMENTUM_WEIGHT + (HYBRID_SENTIMENT_WEIGHT / 4.0)
            eff_sec_weight = HYBRID_SECURITY_WEIGHT + (HYBRID_SENTIMENT_WEIGHT / 4.0)
        else:
            eff_mom_weight = HYBRID_MOMENTUM_WEIGHT
            eff_sec_weight = HYBRID_SECURITY_WEIGHT
            eff_sent_weight = HYBRID_SENTIMENT_WEIGHT

        dec = int(round(clamp(mom * eff_mom_weight + sec * eff_sec_weight + sent * eff_sent_weight, 0.0, 100.0)))
        f["decision_score"], f["karar"] = dec, decision_label_from_score(dec)

def _build_fund_ai_prompt(f: dict) -> str:
    hist = f.get("decision_history", [])[-5:]
    hist_text = ", ".join(f"{x.get('date')}: {x.get('score'):.0f}/{x.get('decision')}" for x in hist if x.get('score') is not None)
    return f"""Sen kıdemli bir fon risk analistisin. Aşağıdaki veriler SADECE modelin verdiği verilerdir. Haber, fiyat veya veri uydurma.
Fon: {f.get('code','-')}
Yatırım alanı: {f.get('investment_area','-')}
Karar skoru: {safe_float(f.get('decision_score'),50):.1f}
Model kararı: {f.get('karar','-')}
Trend skoru: {safe_float(f.get('trend_skor'),50):.1f}
Trend std: {safe_float(f.get('trend_std'),0):.1f}
Momentum: {safe_float(f.get('market_momentum'),50):.1f}
Güvenlik: {safe_float(f.get('security_score'),50):.1f}
AI piyasa duyarlılığı: {safe_float(f.get('sentiment_score'),50):.1f}
Veri anomalisi: {f.get('has_anomaly',False)}
Son 5 günlük model geçmişi: {hist_text or 'yok'}

Görevin: Model kararını bağımsız bir ikinci görüş olarak değerlendir. Ana skoru değiştirme.
JSON döndür: {{"ai_view":"GÜÇLÜ AL|AL|İZLE|SAT|ACİL SAT|YETERSİZ", "confidence":0-100, "risk":0-100, "agreement":"UYUMLU|KISMEN UYUMLU|UYUMSUZ", "reason":"en fazla 18 kelime", "action":"en fazla 12 kelime"}}"""

def enrich_funds_with_ai(funds: List[dict], key_gemini: str, key_groq: str, key_openrouter: str, analyze_all: bool = False) -> dict:
    if not funds: return {"requested": 0, "analyzed": 0, "skipped": 0, "providers": {}}
    candidates = list(funds)
    if not analyze_all:
        candidates = sorted(candidates, key=lambda x: safe_float(x.get("decision_score"), 50))
        important = [x for x in candidates if safe_float(x.get("decision_score"),50) >= STRONG_BUY or safe_float(x.get("decision_score"),50) < CORRECTION]
        mid = sorted(candidates, key=lambda x: safe_float(x.get("decision_score"),50), reverse=True)[:30]
        candidates = list({id(x): x for x in important + mid}.values())
        candidates = candidates[:AI_FUND_MAX_ITEMS]

    providers = [("Gemini", key_gemini, call_gemini_direct, GEMINI_MODEL),
                 ("Groq", key_groq, call_groq_direct, GROQ_MODEL),
                 ("OpenRouter", key_openrouter, call_openrouter_direct, OPENROUTER_MODEL)]
    active = [(n,k,fn,m) for n,k,fn,m in providers if k and k.strip()]
    if not active:
        for f in candidates:
            f["fund_ai_view"] = "AI DEVRE DIŞI"; f["fund_ai_confidence"] = 0; f["fund_ai_reason"] = "API anahtarı yok."
        return {"requested": len(candidates), "analyzed": 0, "skipped": len(candidates), "providers": {}}

    results = {}
    errors = {}
    def run_one(f, provider):
        name,key,fn,model = provider
        try:
            raw = fn(_build_fund_ai_prompt(f), key, model)
            return id(f), name, raw, ""
        except Exception as e:
            return id(f), name, {}, f"{type(e).__name__}: {str(e)[:120]}"

    with concurrent.futures.ThreadPoolExecutor(max_workers=min(3, len(active))) as pool:
        futures = [pool.submit(run_one, f, p) for f in candidates for p in active]
        for fut in concurrent.futures.as_completed(futures):
            fid,name,raw,err = fut.result()
            if raw: results.setdefault(fid, {})[name] = raw
            if err: errors.setdefault(fid, {})[name] = err

    by_id = {id(f): f for f in candidates}
    for fid,f in by_id.items():
        obs=[]
        for name,raw in results.get(fid,{}).items():
            if isinstance(raw,dict):
                view=str(raw.get("ai_view","YETERSİZ")).upper()
                conf=clamp(safe_float(raw.get("confidence"),0),0,100)
                risk=clamp(safe_float(raw.get("risk"),50),0,100)
                agree=str(raw.get("agreement","UYUMSUZ")).upper()
                reason=str(raw.get("reason",""))[:180]
                action=str(raw.get("action",""))[:120]
                obs.append((name,view,conf,risk,agree,reason,action))
        f["fund_ai_providers"]=[x[0] for x in obs]
        f["fund_ai_views"]={x[0]:x[1] for x in obs}
        f["fund_ai_scores"]={x[0]:x[2] for x in obs}
        if not obs:
            f["fund_ai_view"]="AI YANIT YOK"; f["fund_ai_confidence"]=0; f["fund_ai_risk"]=50
            f["fund_ai_agreement"]="YOK"; f["fund_ai_reason"]=" | ".join(errors.get(fid,{}).values())[:300]
            continue
        counts=defaultdict(int)
        for x in obs: counts[x[1]] += 1
        consensus=max(counts, key=counts.get)
        conf=sum(x[2] for x in obs)/len(obs)
        risk=sum(x[3] for x in obs)/len(obs)
        model=f.get("karar","YETERSİZ VERİ")
        mapping={"GÜÇLÜ AL":"GÜÇLÜ AL","ASIL LİSTE":"AL","DÜZELTME / İZLE":"İZLE","ACİL SAT":"ACİL SAT"}
        model_ai=mapping.get(model,model)
        agreement="UYUMLU" if consensus==model_ai else ("KISMEN UYUMLU" if {consensus,model_ai} <= {"AL","İZLE"} or {consensus,model_ai} <= {"SAT","ACİL SAT"} else "UYUMSUZ")
        f["fund_ai_view"]=consensus; f["fund_ai_confidence"]=round(conf,1); f["fund_ai_risk"]=round(risk,1)
        f["fund_ai_agreement"]=agreement; f["fund_ai_reason"]=" | ".join(x[5] for x in obs if x[5])[:300]
        f["fund_ai_action"]=" | ".join(x[6] for x in obs if x[6])[:220]
    return {"requested":len(candidates),"analyzed":sum(1 for f in candidates if f.get("fund_ai_providers")),"skipped":len(funds)-len(candidates),"providers":{n:sum(1 for f in candidates if n in f.get("fund_ai_providers",[])) for n,_,_,_ in active}}

def build_decision_history(f: dict):
    dates=f.get("dates",[]); scores=f.get("running_trend_hybrid",[])
    hist=[]
    for d,sc in zip(dates,scores):
        if sc is None: continue
        hist.append({"date": str(d), "score": int(sc), "decision": decision_label_from_score(sc)})
    f["decision_history"]=hist[-5:]
    for i,item in enumerate(f["decision_history"],1):
        f[f"karar_{i}_tarih"]=item["date"]; f[f"karar_{i}_skor"]=item["score"]; f[f"karar_{i}_model"]=item["decision"]
    urgent=[x["decision"]=="ACİL SAT" for x in f["decision_history"]]
    f["urgent_sell_2day"]=len(urgent)>=2 and urgent[-1] and urgent[-2]
    f["urgent_sell_count_5d"]=sum(urgent)

def run_backtest(funds: List[dict]) -> dict:
    results = {"total_funds": len(funds), "total_signals": 0, "accurate_signals": 0, "avg_accuracy": 0.0}
    for f in funds:
        if f.get("is_synthetic", False): continue
        hybrid_scores = f.get("running_trend_hybrid", [])
        returns = f.get("daily_returns", [])
        for i in range(len(hybrid_scores) - 1):
            score = hybrid_scores[i]
            next_ret = returns[i+1] if i+1 < len(returns) else None
            if score is not None and next_ret is not None:
                results["total_signals"] += 1
                if (score >= 50 and next_ret > 0) or (score < 50 and next_ret <= 0):
                    results["accurate_signals"] += 1
    if results["total_signals"] > 0:
        results["avg_accuracy"] = (results["accurate_signals"] / results["total_signals"]) * 100
    return results

def create_excel_output(wb, ws_list, all_funds, common_n_days):
    if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

    n_dates = common_n_days if common_n_days > 0 else 5
    all_dates = set()
    for f in all_funds:
        for d in f.get("dates", []):
            if d is not None: all_dates.add(d)
        
    def parse_dm(dm_str):
        try: return pd.to_datetime(dm_str).date()
        except: return dt.date(1970, 1, 1)

    sorted_dates = sorted(list(all_dates), key=parse_dm)
    sample_dates = sorted_dates[-n_dates:] if len(sorted_dates) >= n_dates else sorted_dates

    headers = [
        "Fon Kodu", "Fon Adı", "Yatırım Alanı", "Karar Skoru", "Trend Skoru", "Trend Std. Sapma",
        "Piyasa Momentum", "Güvenlik Skoru", "Sentiment Skoru", "Model Kararı", "Anomali Uyarısı",
        "Ort. Günlük Getiri (%)", "Volatilite (%)", "Sharpe", "Kümülatif Getiri (%)", "MaxDD (%)",
        "AUM Değişim (%)", "AUM (₺)", "Yatırımcı", "Haftalık Getiri (%)", "Kıyas Grubu", "AI Sağlayıcıları", "AI Güvenilirliği", "Fon AI Görüşü", "Fon AI Güveni", "Fon AI Uyum", "5G AcilSat", "5G Karar Geçmişi", "Veri Kaynağı"
    ]

    ws_scores.append(headers)
    header_index = {name: idx + 1 for idx, name in enumerate(headers)}

    fill = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    font = Font(name="Calibri", bold=True, color=COLOR_WHITE)
    for cell in ws_scores[1]: cell.fill, cell.font, cell.alignment = fill, font, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_scores.row_dimensions[1].height = 55

    for item in all_funds:
        anomali_msg = item.get("anomaly_reason") if item.get("has_anomaly") else "Yok"
        row_data = [
            item["code"], item.get("fund_title") or "-", item.get("investment_area") or "-",
            item.get("decision_score"), item.get("trend_skor"), item.get("trend_std"),
            item.get("market_momentum"), item.get("security_score"), item.get("sentiment_score"), 
            item.get("karar", "-"), anomali_msg, round(safe_float(item.get("_final_mean_return")), 4), 
            round(safe_float(item.get("volatility")), 4), round(safe_float(item.get("_final_sharpe")), 4), 
            round(safe_float(item.get("_final_cumulative")), 4), round(safe_float(item.get("_final_max_dd")), 4),
            round(safe_float(item.get("aum_change")), 2), round(safe_float(item.get("aum")), 0), item.get("investors"),
            round(safe_float(item.get("weekly_return")), 4), item.get("reference_scope", "-"), ", ".join(item.get("sentiment_ai_providers", [])) or "-", "EVET" if item.get("sentiment_ai_reliable") else "HAYIR", item.get("fund_ai_view", "-") + (" | " + str(item.get("fund_ai_reason", ""))[:100] if item.get("fund_ai_reason") else ""), round(safe_float(item.get("fund_ai_confidence"),0),1), item.get("fund_ai_agreement","-"), "EVET" if item.get("urgent_sell_2day") else "HAYIR", " | ".join(f"{x.get('date')}: {x.get('score')}/{x.get('decision')}" for x in item.get("decision_history", [])) or "-", item.get("source", "-")
        ]
        ws_scores.append(row_data)

    green_font, red_font, yellow_font = Font(bold=True, color=COLOR_GREEN), Font(bold=True, color=COLOR_RED), Font(bold=True, color=COLOR_YELLOW)
    decision_cols = [idx for name, idx in header_index.items() if "Karar" in name and "Skor" not in name]

    for row_number in range(2, ws_scores.max_row + 1):
        for col_idx in decision_cols:
            cell = ws_scores.cell(row=row_number, column=col_idx)
            text = str(cell.value or "").upper()
            if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text: cell.font = green_font
            elif "DÜZELTME" in text: cell.font = yellow_font
            elif "ACİL SAT" in text or "YETERSİZ" in text: cell.font = red_font

    score_cols = [idx for name, idx in header_index.items() if "Skor" in name]
    if ws_scores.max_row >= 2:
        for col_idx in score_cols:
            col_letter = get_column_letter(col_idx)
            rng = f"{col_letter}2:{col_letter}{ws_scores.max_row}"
            ws_scores.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["75"], fill=PatternFill(start_color=COLOR_LIGHT_GREEN, fill_type="solid")))
            ws_scores.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["50", "74"], fill=PatternFill(start_color=COLOR_LIGHT_YELLOW, fill_type="solid")))
            ws_scores.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["50"], fill=PatternFill(start_color=COLOR_LIGHT_RED, fill_type="solid")))

    cur_col, int_col = header_index.get("AUM (₺)"), header_index.get("Yatırımcı")
    pct_cols = ["Ort. Günlük Getiri (%)", "Volatilite (%)", "Kümülatif Getiri (%)", "MaxDD (%)", "AUM Değişim (%)", "Haftalık Getiri (%)"]

    for row_number in range(2, ws_scores.max_row + 1):
        if cur_col: ws_scores.cell(row=row_number, column=cur_col).number_format = '#,##0.00 "₺"'
        if int_col: ws_scores.cell(row=row_number, column=int_col).number_format = "#,##0"
        for col_name in pct_cols:
            idx = header_index.get(col_name)
            if idx and isinstance(ws_scores.cell(row=row_number, column=idx).value, (int, float)):
                ws_scores.cell(row=row_number, column=idx).number_format = '0.00"%"'

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

st.markdown("### 📥 Veri Kaynağı Seçimi")
col_upload, col_github, col_manual = st.columns(3)
wb = None
manuel_req_codes = []

with col_upload:
    uploaded_file = st.file_uploader("Bilgisayardan Excel Yükle", type=["xlsx"])
    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        except Exception as exc: 
            st.error(f"Excel yükleme hatası: {exc}")

with col_github:
    st.markdown("<br><br>", unsafe_allow_html=True)
    if st.button("🚀 GitHub'dan Çek ve Analiz Et", use_container_width=True):
        url = GITHUB_FALLBACK_URL
        res = requests.get(url)
        if res.status_code == 200:
            wb = openpyxl.load_workbook(io.BytesIO(res.content))
            st.success("✅ Veri çekildi.")

with col_manual:
    manuel_fonlar = st.text_area("✍️ Manuel Fon Kodları", placeholder="Örn: MAC, IPB, KHA\n(Virgülle ayırın)")
    if manuel_fonlar:
        manuel_req_codes = [normalize_fund_code(f.strip()) for f in manuel_fonlar.split(",") if f.strip()]

req_codes = []
if manuel_req_codes:
    req_codes = manuel_req_codes
    wb = openpyxl.Workbook()
    ws_list = wb.active
    ws_list.title = "Fon_Listesi"
    ws_list.append(["Fon Kodu"])
    for code in req_codes:
        ws_list.append([code])
elif wb is not None:
    ws_list = wb["Fon_Listesi"] if "Fon_Listesi" in wb.sheetnames else wb.active
    req_codes = [normalize_fund_code(r[0].value) for r in ws_list.iter_rows(min_row=2) if r and r[0].value]

req_codes = list(dict.fromkeys(filter(None, req_codes)))

if not req_codes: 
    st.info("Lütfen analize başlamak için bir Excel dosyası yükleyin, GitHub'dan veri çekin veya manuel fon kodu girin.")
    st.stop()

with st.spinner("🔄 TEFAS evreni taranıyor (pytefas)..."):
    today = dt.date.today()
    universe = fetch_tefas_universe(today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS), today)
    meta_map = build_fund_meta_map(universe)
    ref = build_universe_reference(universe, TARGET_TRADING_DAYS)

calc_funds, failed = [], []
prog = st.progress(0, "Multi-Crawler ile veriler çekiliyor...")
with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as exe:
    futs = {exe.submit(fetch_and_compute_one_fund, c, meta_map): c for c in req_codes}
    for i, fut in enumerate(concurrent.futures.as_completed(futs)):
        c = futs[fut]
        try: _, met, src = fut.result()
        except Exception: met = None
        if met: calc_funds.append(met)
        else: failed.append(c)
        prog.progress((i + 1) / len(req_codes))
prog.empty()

eligible = [f for f in calc_funds if f.get("n_days", 0) >= MIN_ROLLING_DAYS]

if not eligible:
    st.error("❌ Geçerli veri çekilemedi.")
    st.stop()

with st.spinner("📊 V16.1.5 Modeli (Kararlılık Optimizasyonu) Hesaplanıyor..."):
    calculate_security_scores(eligible, ref)
    calculate_market_relative_momentum(eligible, ref, TARGET_TRADING_DAYS)
    
    unique_areas = list(set([f.get("investment_area", "-") for f in eligible if f.get("investment_area")]))
    if not unique_areas: unique_areas = ["-"]
    
    batch_sentiments = fetch_batch_market_sentiment_cross_validated(unique_areas, api_key_gemini, api_key_groq, api_key_openrouter)
    common_n = calculate_trend_scores(eligible, batch_sentiments)
    finalize_decisions(eligible, batch_sentiments)
    for _f in eligible: build_decision_history(_f)
    ai_fund_stats = enrich_funds_with_ai(eligible, api_key_gemini, api_key_groq, api_key_openrouter, ANALYZE_ALL_FUNDS_AI)
    backtest_results = run_backtest(eligible)

output = create_excel_output(wb, ws_list, eligible, common_n)

# ============================================================
# SKOR ÖZETLERİ VE EKRAN TABLOSU
# ============================================================
gercek_veri_sayisi = sum(1 for x in eligible if not x.get("is_synthetic", False))
sahte_veri_sayisi = sum(1 for x in eligible if x.get("is_synthetic", False))

st.warning(f"⚠️ **DİKKAT:** Bu model kısa vadeli (5-10 günlük) istatistiksel trend ölçümü yapar. Üretilen 'GÜÇLÜ AL' veya 'ACİL SAT' kararları yatırım tavsiyesi değil, momentum sinyalidir. Lütfen Standart Sapma (Gürültü) verisini dikkate alın.")

st.subheader(f"📈 KAZRİSK Portföy Özeti ({yatirim_vadesi.split(' ')[0]} Stratejisi)")
col1, col2, col3, col4, col5, col6 = st.columns(6)
scores = [safe_float(x.get("decision_score")) for x in eligible if x.get("decision_score") is not None]

if scores:
    col1.metric("En Yüksek Skor", f"{max(scores):.0f}")
    col2.metric("Ortalama Skor", f"{sum(scores) / len(scores):.1f}")
    col3.metric("Güçlü Al Veren", sum(1 for x in eligible if x.get("karar") == "GÜÇLÜ AL"))
    col4.metric("Gerçek Verili Fon", f"{gercek_veri_sayisi} / {len(eligible)}", delta=f"{sahte_veri_sayisi} Eksik" if sahte_veri_sayisi > 0 else "Tamamı Gerçek", delta_color="inverse")
    col5.metric("Model İsabet Oranı", f"%{backtest_results['avg_accuracy']:.1f}", help="Geçmiş 10 gün için trend skoru ile ertesi gün getirisi arasındaki korelasyon.")
    col6.metric("2 Gün Üst Üste ACİL SAT", sum(1 for x in eligible if x.get("urgent_sell_2day")))

display_rows = []
for item in eligible:
    anomali_mesaji = "⚠️ ANOMALİ" if item.get("has_anomaly") else ""
    row_dict = {
        "Fon Kodu": item["code"],
        "Yatırım Alanı": item.get("investment_area") or "-",
        "Kıyas Grubu": item.get("reference_scope") or "-",
        "Veri Kaynağı": f"🔴 SENTETİK (Geçersiz) {anomali_mesaji}" if item.get("is_synthetic") else f"🟢 {item.get('source', '')} {anomali_mesaji}", 
        "Güncel Karar Skoru": item.get("decision_score"),
        "Trend Skoru": item.get("trend_skor"),
        "Gürültü (Std. Sapma)": item.get("trend_std"), 
        "Güncel Karar": item.get("karar"),
        "AI": ", ".join(item.get("sentiment_ai_providers", [])) or "Pasif",
        "AI Güven": "Çapraz" if item.get("sentiment_ai_reliable") else ("Tek AI" if item.get("sentiment_ai_active") else "Pasif"),
        "Fon AI Görüşü": item.get("fund_ai_view", "-"),
        "Fon AI Güveni": item.get("fund_ai_confidence", 0),
        "AI/Model Uyum": item.get("fund_ai_agreement", "-"),
        "2 Gün Üst Üste ACİL SAT": "🚨 EVET" if item.get("urgent_sell_2day") else "Hayır",
        "Son 5 Gün Kararları": " | ".join(f"{x.get('date')}: {x.get('score'):.0f} - {x.get('decision')}" for x in item.get("decision_history", [])) or "-"
    }
    display_rows.append(row_dict)

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value).upper()
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "🟢" in text: return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text: return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text or "YETERSİZ" in text or "🔴" in text or "ANOMALİ" in text: return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except AttributeError: styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 Analiz Sonuçları (V16.1.5)")

urgent_two = [x for x in eligible if x.get("urgent_sell_2day")]
if urgent_two:
    st.error(f"🚨 {len(urgent_two)} fon için son iki analiz gününde üst üste ACİL SAT sinyali oluştu: {', '.join(x['code'] for x in urgent_two[:20])}")

with st.expander("🤖 Fon Bazlı AI İkinci Görüşleri", expanded=False):
    st.info(f"AI fon analizi: {ai_fund_stats.get('analyzed',0)} fon analiz edildi; {ai_fund_stats.get('skipped',0)} fon atlandı. Ana model skoru AI tarafından değiştirilmez.")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

st.success(f"✅ V16.1.5 Analiz tamamlandı. Toplam {len(eligible)} fon işlendi. Fon AI: {ai_fund_stats.get('analyzed',0)}")
st.download_button(
    label="📥 KAZRİSK V16.1.5 Excel İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_FINAL_V16_1_5.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if SHOW_DIAGNOSTICS:
    st.subheader("🔎 Kaynak Tanılama & AI Logları")
    diagnostic_rows = []
    for item in eligible:
        reason = item.get("sentiment_ai_reason", "Bilinmiyor")
        ai_status = "🟢 Aktif" if item.get("sentiment_ai_active") else f"🔴 Pasif ({reason})"
        for status in item.get("source_statuses", []):
            diagnostic_rows.append({
                "Fon": item["code"],
                "Kaynak": status.get("source"),
                "Denendi": "Evet" if status.get("attempted") else "Hayır",
                "Başarılı": "Evet" if status.get("ok") else "Hayır",
                "Mesaj": status.get("message"),
                "AI Modu": ai_status,
                "AI Sağlayıcıları": ", ".join(item.get("sentiment_ai_providers", [])) or "-",
                "AI Skorları": json.dumps(item.get("sentiment_ai_scores", {}), ensure_ascii=False)
            })
            
    if diagnostic_rows:
        df_diag = pd.DataFrame(diagnostic_rows)
        def color_ai_status(val):
            if isinstance(val, str):
                if "🟢 Aktif" in val or "Evet" in val: return 'color: #008000; font-weight: bold;'
                elif "🔴 Pasif" in val or "Hayır" in val: return 'color: #FF0000; font-weight: bold;'
            return ''
        try: styled_diag = df_diag.style.map(color_ai_status)
        except AttributeError: styled_diag = df_diag.style.applymap(color_ai_status)
        st.dataframe(styled_diag, use_container_width=True, hide_index=True)
