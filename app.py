# ============================================================
# tlgssk - KAZRİSK V16.3
# Profesyonel / Rezilyan Fon Analiz Motoru
# ============================================================
# V16.3 ana hedefleri:
# - Gerçek veri yoksa ASLA sentetik fiyat/AUM/yatırımcı üretmez.
# - TEFAS Direct -> TEFAS crawler/universe -> İş Yatırım çoklu fallback.
# - API cevap/tarih/sayı formatları normalize edilir.
# - Kaynaklar bağımsız denenir; ilk hata diğer kaynakları durdurmaz.
# - Gerçek veri kalitesi ve confidence hesaplanır; sabit 95 kaldırılmıştır.
# - Sharpe, Sortino, Calmar, volatility, drawdown gerçek hesaplanır.
# - Fon grubu relative percentile/rank ile momentum güçlendirilir.
# - Günlük 5 günlük karar geçmişi korunur.
# - 2 ardışık ACİL SAT / GÜÇLÜ AL teyidi korunur.
# - Eksik veri karar skorunu yapay olarak yükseltmez.
# - Gemini yoksa sentiment nötrdür; tahmini iyimser varsayımlar kullanılmaz.
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
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "16.3.0"

# ------------------------------------------------------------
# Streamlit
# ------------------------------------------------------------
st.set_page_config(page_title="tlgssk KAZRİSK V16.3", page_icon="📊", layout="wide")
st.title("📊 tlgssk Hibrit Fon Analizi")
st.caption("TEFAS + İş Yatırım | Gerçek Veri / Çoklu Fallback | KAZRİSK V16.3")

# ------------------------------------------------------------
# Ayarlar
# ------------------------------------------------------------
FUND_KINDS = ("YAT", "EMK", "BYF")
LOOKBACK_CALENDAR_DAYS = 60
TARGET_TRADING_DAYS = 10
MIN_HISTORY_DAYS = 5
MIN_SCORE_HISTORY_DAYS = 3
HTTP_TIMEOUT = 20
MAX_WORKERS = 5
REQUEST_MAX_RETRIES = 3
REQUEST_BACKOFF_FACTOR = 0.8

STRONG_BUY = 75
WATCH_LIST = 50
CORRECTION = 35

DEFAULT_MOMENTUM_WEIGHTS = {"return": 0.25, "sharpe": 0.25, "sortino": 0.15, "cumulative": 0.20, "drawdown": 0.15}
DEFAULT_HYBRID_MOMENTUM_WEIGHT = 0.55
DEFAULT_HYBRID_SECURITY_WEIGHT = 0.30
DEFAULT_HYBRID_SENTIMENT_WEIGHT = 0.15

# ------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------
st.sidebar.header("⚙️ Analiz Parametreleri")
env_key = os.environ.get("GEMINI_API_KEY", "")
try:
    if not env_key and "GEMINI_API_KEY" in st.secrets:
        env_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass
api_key_input = st.sidebar.text_input("🔑 Gemini API Key", value=env_key, type="password")

with st.sidebar.expander("⚖️ Momentum Ağırlıkları", expanded=False):
    w_return = st.slider("Getiri", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_sortino = st.slider("Sortino", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sortino"], 0.05)
    w_cum = st.slider("Kümülatif", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_dd = st.slider("Drawdown", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05)
    total = sum([w_return, w_sharpe, w_sortino, w_cum, w_dd]) or 1.0
    MOMENTUM_WEIGHTS = {
        "return": w_return / total,
        "sharpe": w_sharpe / total,
        "sortino": w_sortino / total,
        "cumulative": w_cum / total,
        "drawdown": w_dd / total,
    }

w_hm = st.sidebar.slider("Momentum toplam ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_MOMENTUM_WEIGHT, 0.05)
w_hs = st.sidebar.slider("Güvenlik toplam ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SECURITY_WEIGHT, 0.05)
w_ht = st.sidebar.slider("Sentiment toplam ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SENTIMENT_WEIGHT, 0.05)
tot = w_hm + w_hs + w_ht or 1.0
HYBRID_MOMENTUM_WEIGHT = w_hm / tot
HYBRID_SECURITY_WEIGHT = w_hs / tot
HYBRID_SENTIMENT_WEIGHT = w_ht / tot
RISK_FREE_ANNUAL = st.sidebar.number_input("Yıllık risksiz getiri (%)", 0.0, 100.0, 0.0, 0.5)

# ------------------------------------------------------------
# HTTP
# ------------------------------------------------------------
def build_session():
    session = requests.Session()
    retry = Retry(
        total=REQUEST_MAX_RETRIES,
        connect=REQUEST_MAX_RETRIES,
        read=REQUEST_MAX_RETRIES,
        status=REQUEST_MAX_RETRIES,
        backoff_factor=REQUEST_BACKOFF_FACTOR,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST"]),
        respect_retry_after_header=True,
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
    })
    return session

HTTP = build_session()

@dataclass
class SourceStatus:
    source: str
    attempted: bool = False
    ok: bool = False
    status_code: Optional[int] = None
    rows: int = 0
    elapsed_ms: int = 0
    message: str = ""
    root_cause: str = ""


def request_with_status(source: str, method: str, url: str, **kwargs):
    status = SourceStatus(source=source, attempted=True)
    started = time.time()
    try:
        response = HTTP.request(method, url, timeout=HTTP_TIMEOUT, **kwargs)
        status.status_code = response.status_code
        status.elapsed_ms = int((time.time() - started) * 1000)
        status.ok = 200 <= response.status_code < 300
        if not status.ok:
            status.message = f"HTTP {response.status_code}"
            if response.status_code == 429:
                status.root_cause = "Rate limit / geçici erişim kısıtı"
            elif response.status_code >= 500:
                status.root_cause = "Sunucu tarafı HTTP hatası"
            elif response.status_code >= 400:
                status.root_cause = "İstek veya endpoint reddedildi"
        return response, status
    except requests.Timeout as exc:
        status.elapsed_ms = int((time.time() - started) * 1000)
        status.message = str(exc)[:120]
        status.root_cause = "Timeout"
    except requests.RequestException as exc:
        status.elapsed_ms = int((time.time() - started) * 1000)
        status.message = str(exc)[:120]
        status.root_cause = "HTTP bağlantı hatası"
    except Exception as exc:
        status.elapsed_ms = int((time.time() - started) * 1000)
        status.message = str(exc)[:120]
        status.root_cause = "Beklenmeyen bağlantı hatası"
    return None, status

# ------------------------------------------------------------
# Parsers
# ------------------------------------------------------------
def normalize_fund_code(code) -> str:
    if code is None:
        return ""
    s = str(code).strip().upper()
    s = re.sub(r"\s+", "", s)
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^A-Z0-9ÇĞİÖŞÜ]", "", s)


def optional_float(v):
    try:
        if v is None or (isinstance(v, str) and not v.strip()):
            return None
        x = float(v)
        return None if pd.isna(x) or not math.isfinite(x) else x
    except Exception:
        return None


def parse_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return optional_float(value)
    s = str(value).strip().replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "")
    if not s:
        return None
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif re.match(r"^-?\d{1,3}(\.\d{3})+$", s):
        s = s.replace(".", "")
    return optional_float(s)


def parse_date(value):
    if value is None:
        return pd.NaT
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        x = float(value)
        try:
            if x > 1e14: return pd.to_datetime(x, unit="us", errors="coerce")
            if x > 1e11: return pd.to_datetime(x, unit="ms", errors="coerce")
            if x > 1e9: return pd.to_datetime(x, unit="s", errors="coerce")
        except Exception:
            pass
    s = str(value).strip()
    for dayfirst in (False, True):
        try:
            d = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst)
            if not pd.isna(d): return d
        except Exception:
            pass
    return pd.NaT


def normalize_key(v):
    return re.sub(r"[^a-z0-9çğıöşü]", "", str(v).lower())


def pick_field(record, names):
    if not isinstance(record, dict): return None
    mapping = {normalize_key(k): v for k, v in record.items()}
    for name in names:
        key = normalize_key(name)
        if key in mapping: return mapping[key]
    for key, value in mapping.items():
        for name in names:
            nk = normalize_key(name)
            if nk and (nk in key or key in nk): return value
    return None


def unwrap_records(payload):
    if payload is None: return []
    if isinstance(payload, list): return payload
    if isinstance(payload, dict):
        for k in ("data", "Data", "value", "Value", "items", "Items", "rows", "Rows", "result", "Result", "records", "Records"):
            v = payload.get(k)
            if isinstance(v, list): return v
            if isinstance(v, dict):
                nested = unwrap_records(v)
                if nested: return nested
        return [payload] if pick_field(payload, ["TARIH", "Tarih", "date"]) is not None else []
    return []


def normalize_price_records(payload, code=""):
    rows = []
    for rec in unwrap_records(payload):
        date_value = pick_field(rec, ["TARIH", "Tarih", "tarih", "date", "DATE", "priceDate"])
        price_value = pick_field(rec, ["FIYAT", "Fiyat", "fiyat", "price", "PRICE", "NAV", "nav", "unitPrice"])
        d = parse_date(date_value)
        p = parse_number(price_value)
        if pd.isna(d) or p is None or p <= 0: continue
        rows.append({"date": pd.Timestamp(d).normalize(), "price": p, "code": normalize_fund_code(code)})
    if not rows: return pd.DataFrame(columns=["date", "price", "code"])
    df = pd.DataFrame(rows).sort_values("date").drop_duplicates("date", keep="last")
    return df.reset_index(drop=True)


def merge_preferred(*dfs):
    result = pd.DataFrame(columns=["date", "price", "code"])
    occupied = set()
    for df in dfs:
        if df is None or df.empty: continue
        x = df.copy()
        x["date"] = pd.to_datetime(x["date"], errors="coerce").dt.normalize()
        x["price"] = pd.to_numeric(x["price"], errors="coerce")
        x = x.dropna(subset=["date", "price"])
        for r in x.itertuples(index=False):
            d = pd.Timestamp(r.date)
            if d not in occupied:
                result.loc[len(result)] = [d, float(r.price), getattr(r, "code", "")]
                occupied.add(d)
    return result.sort_values("date").reset_index(drop=True)

# ------------------------------------------------------------
# TEFAS direct
# ------------------------------------------------------------
def fetch_tefas_direct(code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(code)
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {
        "Origin": "https://www.tefas.gov.tr",
        "Referer": "https://www.tefas.gov.tr/TarihselVeriler.aspx",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    }
    statuses = []
    kinds = [fund_kind] if fund_kind in FUND_KINDS else list(FUND_KINDS)
    for kind in kinds:
        payload = {"fontip": kind, "fonkod": code, "bastarih": start.strftime("%d.%m.%Y"), "bittarih": end.strftime("%d.%m.%Y")}
        response, status = request_with_status("TEFAS Direct API", "POST", url, data=payload, headers=headers)
        if response is None:
            statuses.append(status); continue
        try:
            payload_json = response.json()
            df = normalize_price_records(payload_json, code)
            if not df.empty:
                # Optional metadata for real TEFAS fields.
                raw = unwrap_records(payload_json)
                meta = pd.DataFrame(raw)
                if not meta.empty:
                    dates = [parse_date(pick_field(r, ["TARIH", "Tarih", "date"])) for r in raw]
                    aums = [parse_number(pick_field(r, ["PORTFOYBUYUKLUK", "portfoybuyukluk", "market_cap"])) for r in raw]
                    invs = [parse_number(pick_field(r, ["KISISAYISI", "kisisayisi", "number_of_investors"])) for r in raw]
                    meta2 = pd.DataFrame({"date": dates, "aum": aums, "investors": invs})
                    meta2["date"] = pd.to_datetime(meta2["date"], errors="coerce").dt.normalize()
                    df = df.merge(meta2.dropna(subset=["date"]).drop_duplicates("date", keep="last"), on="date", how="left")
                status.ok = True; status.rows = len(df); status.message = f"Başarılı ({len(df)} gün)"; status.root_cause = "Sorun Yok"
                statuses.append(status)
                return df, "TEFAS Direct API", statuses
            status.message = "200 cevap alındı ancak fiyat satırı parse edilemedi"
            status.root_cause = "Response schema / tarih / fiyat alanı uyumsuzluğu"
        except Exception as exc:
            status.message = str(exc)[:120]; status.root_cause = "JSON/response parse hatası"
        statuses.append(status)
    return None, "TEFAS Direct API", statuses

# ------------------------------------------------------------
# TEFAS crawler / universe (optional dependency)
# ------------------------------------------------------------
def fetch_tefas_crawler(code: str):
    statuses = []
    try:
        from tefas import Crawler
    except Exception as exc:
        statuses.append(SourceStatus("TEFAS Crawler", True, False, None, 0, 0, "tefas paketi kurulu değil", str(exc)[:100]))
        return None, "TEFAS Crawler", statuses
    start = (dt.date.today() - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)).strftime("%Y-%m-%d")
    end = dt.date.today().strftime("%Y-%m-%d")
    started = time.time()
    status = SourceStatus("TEFAS Crawler", True)
    try:
        crawler = Crawler()
        raw = crawler.fetch(start=start, end=end, name=normalize_fund_code(code))
        if raw is not None and not raw.empty:
            date_col = next((c for c in ["date", "TARIH", "Tarih"] if c in raw.columns), None)
            price_col = next((c for c in ["price", "FIYAT", "Fiyat"] if c in raw.columns), None)
            if date_col and price_col:
                df = raw.copy()
                df["date"] = df[date_col].apply(parse_date).dt.normalize()
                df["price"] = df[price_col].apply(parse_number)
                df = df.dropna(subset=["date", "price"]); df = df[df["price"] > 0]
                if not df.empty:
                    out = df[["date", "price"]].drop_duplicates("date").sort_values("date").reset_index(drop=True)
                    status.ok = True; status.rows = len(out); status.status_code = 200; status.elapsed_ms = int((time.time()-started)*1000); status.message = f"Başarılı ({len(out)} gün)"; status.root_cause = "Sorun Yok"
                    return out, "TEFAS Crawler", [status]
            status.message = "Crawler cevap verdi ancak date/price alanları bulunamadı"; status.root_cause = "Crawler schema uyumsuzluğu"
        else:
            status.message = "Crawler boş veri döndürdü"; status.root_cause = "Fon/tarih/erişim problemi"
    except Exception as exc:
        status.message = str(exc)[:120]; status.root_cause = "Crawler çalışma hatası"
    status.elapsed_ms = int((time.time()-started)*1000)
    return None, "TEFAS Crawler", [status]

# ------------------------------------------------------------
# İş Yatırım
# ------------------------------------------------------------
def fetch_isyatirim(code: str):
    code = normalize_fund_code(code)
    end = dt.datetime.now(); start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    response, status = request_with_status("İş Yatırım", "GET", url, params=params)
    if response is None: return None, "İş Yatırım", [status]
    try:
        payload = response.json()
        records = unwrap_records(payload)
        rows = []
        for rec in records:
            d = parse_date(pick_field(rec, ["Tarih", "tarih", "date", "DATE"]))
            p = parse_number(pick_field(rec, ["Fiyat", "fiyat", "price", "PRICE"]))
            if pd.isna(d) or p is None or p <= 0: continue
            rows.append({"date": pd.Timestamp(d).normalize(), "price": p, "code": code})
        if rows:
            df = pd.DataFrame(rows).drop_duplicates("date").sort_values("date").reset_index(drop=True)
            status.ok = True; status.rows = len(df); status.message = f"Başarılı ({len(df)} gün)"; status.root_cause = "Sorun Yok"
            return df, "İş Yatırım", [status]
        status.message = "Cevap alındı fakat fiyat satırı parse edilemedi"; status.root_cause = "Response schema / fon kodu / tarih problemi"
    except Exception as exc:
        status.message = str(exc)[:120]; status.root_cause = "JSON parse hatası"
    return None, "İş Yatırım", [status]

# ------------------------------------------------------------
# Source cascade - NO SYNTHETIC DATA
# ------------------------------------------------------------
def get_fund_series(code: str, fund_kind: Optional[str] = None):
    all_statuses = []
    sources = []
    for fn in (
        lambda: fetch_tefas_direct(code, fund_kind),
        lambda: fetch_tefas_crawler(code),
        lambda: fetch_isyatirim(code),
    ):
        try:
            df, source, statuses = fn()
        except Exception as exc:
            df, source, statuses = None, "Bilinmeyen", [SourceStatus("Bilinmeyen", True, False, None, 0, 0, str(exc)[:120], "Kaynak çağrısı exception")]
        all_statuses.extend(statuses)
        if df is not None and len(df) >= 2:
            sources.append((source, df))

    if not sources:
        return None, "YOK", all_statuses

    # First source has priority on duplicate dates; fallbacks fill missing dates.
    merged = merge_preferred(*(x[1] for x in sources))
    source_names = [x[0] for x in sources]
    if len(source_names) == 1:
        source_label = source_names[0]
    else:
        source_label = " + ".join(source_names)
    return merged, source_label, all_statuses

# ------------------------------------------------------------
# Fund classification / structural data
# ------------------------------------------------------------
KNOWN_AREAS = {
    "YAY": "Hisse Senedi (Yabancı Teknoloji)", "AFT": "Hisse Senedi (Yabancı Teknoloji)", "AFA": "Hisse Senedi (Yabancı)",
    "TTE": "Hisse Senedi (Yabancı Teknoloji)", "GUH": "Hisse Senedi (Yabancı)", "ITP": "Hisse Senedi (Yabancı)",
    "KZL": "Kıymetli Maden", "GUM": "Kıymetli Maden", "GGK": "Kıymetli Maden", "KGM": "Kıymetli Maden",
    "KUT": "Kıymetli Maden", "AFO": "Kıymetli Maden", "PPZ": "Para Piyasası", "PNU": "Para Piyasası",
    "TP2": "Para Piyasası", "FIL": "Para Piyasası", "NVB": "Para Piyasası", "NRC": "Para Piyasası",
    "THF": "Hisse Senedi", "KHA": "Hisse Senedi", "MAC": "Hisse Senedi", "TCD": "Hisse Senedi", "BIO": "Hisse Senedi",
    "DBH": "Borçlanma Araçları", "YBE": "Borçlanma Araçları", "AKE": "Borçlanma Araçları", "FUB": "Borçlanma Araçları",
}

def classify_area(code: str):
    return KNOWN_AREAS.get(normalize_fund_code(code), "Karma / Fon Türü Belirsiz")


def fetch_fund_structural_data(code: str):
    # V16.3 deliberately does not invent portfolio concentration/liquidity.
    return {
        "investment_area": classify_area(code),
        "top_asset_weight": None,
        "asset_class_hhi": None,
        "emergency_cash_ratio": None,
        "cash_ratio_known": False,
        "is_bist30": False,
        "structural_fetch_ok": False,
        "structural_source": "YOK",
    }

# ------------------------------------------------------------
# Metrics
# ------------------------------------------------------------
def max_drawdown(prices):
    vals = [optional_float(x) for x in prices]
    vals = [x for x in vals if x is not None and x > 0]
    if len(vals) < 2: return 0.0
    peak = vals[0]; mdd = 0.0
    for p in vals:
        peak = max(peak, p)
        if peak > 0: mdd = min(mdd, (p / peak - 1.0) * 100.0)
    return mdd


def compounded_return(rets):
    growth = 1.0; used = 0
    for r in rets:
        x = optional_float(r)
        if x is None: continue
        growth *= 1.0 + x / 100.0; used += 1
    return (growth - 1.0) * 100.0 if used else 0.0


def compute_metrics(series: pd.DataFrame, code: str):
    if series is None or len(series) < 2: return None
    df = series.copy().sort_values("date").drop_duplicates("date", keep="last")
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.normalize()
    df["price"] = df["price"].apply(parse_number)
    df = df.dropna(subset=["date", "price"]); df = df[df["price"] > 0]
    if len(df) < 2: return None
    prices = df["price"].astype(float).tolist()
    dates = [d.strftime("%Y-%m-%d") for d in df["date"]]
    rets = [((prices[i] / prices[i-1]) - 1.0) * 100.0 for i in range(1, len(prices)) if prices[i-1] > 0]
    ret_dates = dates[-len(rets):]
    if len(rets) < 1: return None

    area = classify_area(code)
    struct = fetch_fund_structural_data(code)
    rf_daily = ((1 + RISK_FREE_ANNUAL / 100.0) ** (1 / 252.0) - 1) * 100.0
    excess = [r - rf_daily for r in rets]
    mean_excess = sum(excess) / len(excess)
    std = (sum((r - sum(excess)/len(excess))**2 for r in excess) / max(1, len(excess)-1)) ** 0.5 if len(excess) > 1 else 0.0
    sharpe = mean_excess / std * math.sqrt(252) if std > 1e-12 and len(excess) >= 3 else None
    downside = [min(r, 0.0) for r in excess]
    downside_dev = (sum(x*x for x in downside) / len(downside)) ** 0.5 if downside else 0.0
    sortino = mean_excess / downside_dev * math.sqrt(252) if downside_dev > 1e-12 and len(excess) >= 3 else None
    cum = (prices[-1] / prices[0] - 1.0) * 100.0
    mdd = max_drawdown(prices)
    calmar = (cum / abs(mdd)) if abs(mdd) > 1e-12 else None
    ann_vol = std * math.sqrt(252) if len(excess) >= 2 else None

    aum = None; investors = None
    if "aum" in df.columns:
        vals = [parse_number(v) for v in df["aum"]]
        aum = next((v for v in reversed(vals) if v is not None and v > 0), None)
    if "investors" in df.columns:
        vals = [parse_number(v) for v in df["investors"]]
        investors = next((v for v in reversed(vals) if v is not None and v >= 0), None)

    return {
        "code": normalize_fund_code(code), "fund_title": normalize_fund_code(code), "investment_area": area,
        "dates": ret_dates, "orig_dates": ret_dates, "orig_returns": rets,
        "prices": prices, "price_dates": dates, "price_map": dict(zip(dates, prices)),
        "daily_returns": rets, "n_days": len(rets), "price_days": len(prices),
        "aum": aum, "investors": int(investors) if investors is not None else None,
        "max_dd": mdd, "weekly_return": compounded_return(rets[-5:]),
        "_final_mean_return": sum(rets)/len(rets), "_final_sharpe": sharpe, "_final_sortino": sortino,
        "_final_calmar": calmar, "_final_cumulative": cum, "_final_max_dd": mdd,
        "annualized_volatility": ann_vol, **struct,
    }

# ------------------------------------------------------------
# Relative scoring
# ------------------------------------------------------------
def percentile(values, higher_is_better=True):
    clean = [v for v in values if v is not None and math.isfinite(v)]
    if not clean: return {}
    s = sorted(clean)
    out = {}
    for v in clean:
        rank = sum(1 for x in s if x < v) + 0.5 * sum(1 for x in s if x == v)
        p = 100.0 * rank / len(s)
        out[v] = p if higher_is_better else 100.0 - p
    return out


def score_component(value, population, higher=True, default=50.0):
    if value is None or not math.isfinite(value): return default
    vals = [x for x in population if x is not None and math.isfinite(x)]
    if len(vals) < 3: return default
    vals_sorted = sorted(vals)
    rank = sum(1 for x in vals_sorted if x < value) + 0.5 * sum(1 for x in vals_sorted if x == value)
    p = 100.0 * rank / len(vals_sorted)
    return p if higher else 100.0 - p


def calculate_momentum(funds):
    # Relative to the requested universe; if >=5 peers exist, use group-relative values.
    groups = defaultdict(list)
    for f in funds: groups[f.get("investment_area", "Karma")].append(f)
    for f in funds:
        peers = groups[f.get("investment_area", "Karma")]
        use_group = len(peers) >= 5
        pop = peers if use_group else funds
        vals_return = [p.get("_final_mean_return") for p in pop]
        vals_sharpe = [p.get("_final_sharpe") for p in pop]
        vals_sortino = [p.get("_final_sortino") for p in pop]
        vals_cum = [p.get("_final_cumulative") for p in pop]
        vals_dd = [p.get("_final_max_dd") for p in pop]
        components = {
            "return": score_component(f.get("_final_mean_return"), vals_return, True),
            "sharpe": score_component(f.get("_final_sharpe"), vals_sharpe, True),
            "sortino": score_component(f.get("_final_sortino"), vals_sortino, True),
            "cumulative": score_component(f.get("_final_cumulative"), vals_cum, True),
            "drawdown": score_component(f.get("_final_max_dd"), vals_dd, False),
        }
        score = sum(MOMENTUM_WEIGHTS[k] * components[k] for k in components)
        f["relative_group_size"] = len(peers)
        f["relative_group_mode"] = "Fon Alanı" if use_group else "Analiz Evreni"
        f["momentum_components"] = components
        f["market_momentum"] = round(max(0.0, min(100.0, score)), 1)

# ------------------------------------------------------------
# Security / confidence / data quality
# ------------------------------------------------------------
def calculate_data_quality(f):
    n = int(f.get("n_days", 0))
    score = 0.0
    issues = []
    # Coverage: 60 points
    score += min(60.0, 60.0 * n / 20.0)
    if n < MIN_HISTORY_DAYS: issues.append(f"Kısa geçmiş ({n} gün)")
    # Source reliability: 20 points
    source = str(f.get("source", ""))
    if "TEFAS Direct" in source: score += 20
    elif "TEFAS Crawler" in source: score += 18
    elif "İş Yatırım" in source: score += 15
    else: issues.append("Birincil kaynak belirsiz")
    # Structural availability: 10 points
    if f.get("structural_fetch_ok"): score += 10
    else: issues.append("Portföy yapısal verisi yok")
    # AUM/investor real data: 10 points
    if f.get("aum") is not None: score += 5
    else: issues.append("AUM yok")
    if f.get("investors") is not None: score += 5
    else: issues.append("Yatırımcı sayısı yok")
    return round(max(0.0, min(100.0, score)), 1), "; ".join(issues) if issues else "OK"


def calculate_security(funds):
    # Security uses only actually observed variables. Unknown structural variables do not receive bonus.
    aums = [f.get("aum") for f in funds if f.get("aum") is not None and f.get("aum") > 0]
    invs = [f.get("investors") for f in funds if f.get("investors") is not None and f.get("investors") >= 0]
    for f in funds:
        aum_score = score_component(f.get("aum"), aums, True) if aums else 50
        inv_score = score_component(f.get("investors"), invs, True) if invs else 50
        dd = f.get("_final_max_dd")
        risk_score = 50.0 if dd is None else max(0.0, min(100.0, 100.0 + dd * 2.0))
        vol = f.get("annualized_volatility")
        vol_score = 50.0 if vol is None else max(0.0, min(100.0, 100.0 - vol * 2.0))
        structural_score = 50.0
        if f.get("asset_class_hhi") is not None:
            hhi = float(f["asset_class_hhi"])
            structural_score = max(0.0, min(100.0, 100.0 - hhi * 100.0))
        f["security_score"] = round(0.25*aum_score + 0.20*inv_score + 0.35*risk_score + 0.20*vol_score, 1)
        f["security_subscores"] = {"aum": aum_score, "investor": inv_score, "drawdown": risk_score, "volatility": vol_score, "structure": structural_score}

# ------------------------------------------------------------
# Sentiment
# ------------------------------------------------------------
@st.cache_data(ttl=60*60*4, show_spinner=False)
def fetch_batch_market_sentiment(areas_tuple, api_key):
    areas = list(areas_tuple)
    result = {}
    key = (api_key or "").strip()
    if not key:
        return {a: {"score": 50, "label": "Nötr / AI kapalı", "ai_active": False, "ai_reason": "API anahtarı yok"} for a in areas}
    prompt = "Sen finansal piyasa analisti olarak aşağıdaki yatırım alanlarını güncel makro/piyasa görünümü açısından 0-100 puanla. Sadece JSON döndür. Her alan için score ve kısa label ver. Veri erişimin yoksa 50 kullan ve bunu label'da belirt.\n" + "\n".join(f"- {a}" for a in areas)
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
    payload = {"contents":[{"role":"user","parts":[{"text":prompt}]}],"generationConfig":{"responseMimeType":"application/json","temperature":0.1}}
    try:
        response = requests.post(url, params={"key": key}, json=payload, timeout=18)
        if response.status_code == 200:
            raw = response.json().get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "{}")
            parsed = json.loads(re.sub(r"^```json|```$", "", raw.strip()).strip())
            for a in areas:
                item = parsed.get(a, {}) if isinstance(parsed, dict) else {}
                result[a] = {"score": int(max(0, min(100, optional_float(item.get("score")) or 50))), "label": str(item.get("label", "Nötr")), "ai_active": True, "ai_reason": "Gemini canlı yanıt"}
            return result
    except Exception as exc:
        err = str(exc)[:100]
    else:
        err = f"HTTP {response.status_code}"
    return {a: {"score": 50, "label": "Nötr / AI kullanılamadı", "ai_active": False, "ai_reason": err} for a in areas}

# ------------------------------------------------------------
# Daily score history
# ------------------------------------------------------------
def daily_relative_score(f, date_key, peers):
    # Use the return on that day, relative to peer median if enough peer observations exist.
    ret_map = dict(zip(f.get("orig_dates", []), f.get("orig_returns", [])))
    r = optional_float(ret_map.get(date_key))
    if r is None: return None
    peer_vals = []
    for p in peers:
        v = optional_float(dict(zip(p.get("orig_dates", []), p.get("orig_returns", []))).get(date_key))
        if v is not None: peer_vals.append(v)
    if len(peer_vals) >= 3:
        median = float(pd.Series(peer_vals).median())
        rel = r - median
    else:
        rel = r
    return max(0.0, min(100.0, 50.0 + rel * 12.0))


def calculate_daily_history(funds, sentiments):
    all_dates = sorted({d for f in funds for d in f.get("orig_dates", [])})
    common_dates = all_dates[-TARGET_TRADING_DAYS:]
    groups = defaultdict(list)
    for f in funds: groups[f.get("investment_area", "Karma")].append(f)
    for f in funds:
        peers = groups[f.get("investment_area", "Karma")]
        sent = float(sentiments.get(f.get("investment_area", "Karma"), {}).get("score", 50))
        sec = float(f.get("security_score", 50))
        history = []
        for d in common_dates:
            m = daily_relative_score(f, d, peers)
            if m is None:
                history.append(None)
            else:
                score = m * HYBRID_MOMENTUM_WEIGHT + sec * HYBRID_SECURITY_WEIGHT + sent * HYBRID_SENTIMENT_WEIGHT
                history.append(round(max(0.0, min(100.0, score))))
        f["dates"] = common_dates
        f["running_trend_hybrid"] = history
        valid = [x for x in history if x is not None]
        f["trend_skor"] = valid[-1] if valid else None
        f["last_5_scores"] = valid[-5:]
    return common_dates


def decision_label(score):
    if score is None: return "YETERSİZ VERİ"
    x = float(score)
    if x >= STRONG_BUY: return "GÜÇLÜ AL"
    if x >= WATCH_LIST: return "ASIL LİSTE"
    if x >= CORRECTION: return "DÜZELTME / İZLE"
    return "ACİL SAT"


def finalize(funds, sentiments):
    for f in funds:
        sent_data = sentiments.get(f.get("investment_area", "Karma"), {"score":50,"label":"Nötr","ai_active":False,"ai_reason":"Yok"})
        sent = float(sent_data.get("score", 50))
        mom = float(f.get("market_momentum", 50))
        sec = float(f.get("security_score", 50))
        raw = mom*HYBRID_MOMENTUM_WEIGHT + sec*HYBRID_SECURITY_WEIGHT + sent*HYBRID_SENTIMENT_WEIGHT
        dq, issues = calculate_data_quality(f)
        # Confidence is evidence quality, not probability of profit.
        confidence = max(0.0, min(100.0, 0.65*dq + 0.35*min(100.0, f.get("n_days",0)/20*100)))
        # Low confidence prevents an overconfident strong signal; it does not fabricate data.
        final_score = raw * (0.75 + 0.25*confidence/100.0)
        if confidence < 45:
            final_score = min(final_score, 74.0)
        f.update({
            "sentiment_score": sent, "sentiment_label": sent_data.get("label"),
            "sentiment_ai_active": sent_data.get("ai_active", False), "sentiment_ai_reason": sent_data.get("ai_reason", ""),
            "raw_decision_score": round(raw,1), "decision_score": round(max(0,min(100,final_score)),1),
            "karar": decision_label(final_score), "data_quality_score": dq, "data_quality_issues": issues,
            "confidence_score": round(confidence,1),
        })

# ------------------------------------------------------------
# Excel
# ------------------------------------------------------------
NAVY="1F4E79"; GREEN="008000"; RED="FF0000"; YELLOW="B8860B"; WHITE="FFFFFF"
LIGHT_GREEN="E2F0D9"; LIGHT_YELLOW="FFF2CC"; LIGHT_RED="FCE4D6"

def create_excel_output(wb, funds, common_dates):
    if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
    ws = wb.create_sheet("KGDM3_Puanlama")
    last5 = list(common_dates[-5:])
    headers=["Fon Kodu","Fon Adı","Yatırım Alanı"]
    for d in reversed(last5): headers += [f"{pd.to_datetime(d).strftime('%d.%m.%Y')} Karar Skoru", f"{pd.to_datetime(d).strftime('%d.%m.%Y')} Model Kararı"]
    headers += ["Güncel Karar Skoru","Ham Skor","Confidence","Veri Kalite","Model Kararı","Momentum","Güvenlik","Sentiment","Ort. Günlük Getiri (%)","Volatilite (%)","Sharpe","Sortino","Calmar","Kümülatif (%)","MaxDD (%)","Haftalık (%)","Kaynak","Relative Grup","Veri Kalitesi Notu"]
    ws.append(headers)
    for c in ws[1]: c.fill=PatternFill(start_color=NAVY, fill_type="solid"); c.font=Font(bold=True,color=WHITE); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for f in funds:
        row=[f["code"],f.get("fund_title",f["code"]),f.get("investment_area","-")]
        smap=dict(zip(f.get("dates",[]),f.get("running_trend_hybrid",[])))
        for d in reversed(last5):
            s=smap.get(d); row += [s if s is not None else "Veri Açıklanmadı", decision_label(s) if s is not None else "Veri Açıklanmadı"]
        row += [f.get("decision_score"),f.get("raw_decision_score"),f.get("confidence_score"),f.get("data_quality_score"),f.get("karar"),f.get("market_momentum"),f.get("security_score"),f.get("sentiment_score"),round(f.get("_final_mean_return",0),4),round(f.get("annualized_volatility",0) or 0,4),f.get("_final_sharpe"),f.get("_final_sortino"),f.get("_final_calmar"),round(f.get("_final_cumulative",0),4),round(f.get("_final_max_dd",0),4),round(f.get("weekly_return",0),4),f.get("source","-"),f.get("relative_group_mode","-"),f.get("data_quality_issues","OK")]
        ws.append(row)
    ws.freeze_panes="A2"; ws.sheet_view.showGridLines=False
    for col in range(1, ws.max_column+1): ws.column_dimensions[get_column_letter(col)].width=min(28,max(12,max(len(str(ws.cell(r,col).value or "")) for r in range(1,min(ws.max_row,50)+1))+2))
    out=io.BytesIO(); wb.save(out); out.seek(0); return out

# ------------------------------------------------------------
# Input
# ------------------------------------------------------------
if "req_codes" not in st.session_state: st.session_state["req_codes"]=[]
if "wb_bytes" not in st.session_state: st.session_state["wb_bytes"]=None

st.markdown("### 📥 Veri Giriş Yöntemi")
input_method=st.radio("Kaynak",["✍️ Manuel Fon Girişi","📁 Bilgisayardan Excel Yükle","🌐 GitHub Raw Excel"],horizontal=True,label_visibility="collapsed")

if input_method.startswith("✍️"):
    with st.form("manual"):
        txt=st.text_area("Fon kodları",value="KZL + THF + PNU + ICH + KHA")
        if st.form_submit_button("🚀 Analizi Başlat",type="primary",use_container_width=True):
            codes=list(dict.fromkeys([normalize_fund_code(x) for x in re.split(r"[\s+,;\-]+",txt) if x.strip()]))
            if codes:
                wb0=openpyxl.Workbook(); ws0=wb0.active; ws0.title="Fon_Listesi"; ws0.append(["Fon Kodu"])
                for c in codes: ws0.append([c])
                b=io.BytesIO(); wb0.save(b); st.session_state["wb_bytes"]=b.getvalue(); st.session_state["req_codes"]=codes; st.rerun()
elif input_method.startswith("📁"):
    up=st.file_uploader("Excel",type=["xlsx"])
    if up:
        try:
            content=up.read(); wb0=openpyxl.load_workbook(io.BytesIO(content)); ws0=wb0["Fon_Listesi"] if "Fon_Listesi" in wb0.sheetnames else wb0.active
            codes=list(dict.fromkeys([normalize_fund_code(r[0].value) for r in ws0.iter_rows(min_row=2) if r and r[0].value]))
            if codes: st.session_state["wb_bytes"]=content; st.session_state["req_codes"]=codes
        except Exception as exc: st.error(f"Excel yükleme hatası: {exc}")
else:
    with st.form("github"):
        url=st.text_input("GitHub Raw Excel URL",value="https://raw.githubusercontent.com/tlgssk/kazrisk/main/fonlar.xlsx")
        if st.form_submit_button("🚀 GitHub'dan Çek",type="primary",use_container_width=True):
            response,status=request_with_status("GitHub","GET",url.strip())
            if response is not None and status.ok:
                try:
                    content=response.content; wb0=openpyxl.load_workbook(io.BytesIO(content)); ws0=wb0["Fon_Listesi"] if "Fon_Listesi" in wb0.sheetnames else wb0.active
                    codes=list(dict.fromkeys([normalize_fund_code(r[0].value) for r in ws0.iter_rows(min_row=2) if r and r[0].value]))
                    if codes: st.session_state["wb_bytes"]=content; st.session_state["req_codes"]=codes; st.rerun()
                except Exception as exc: st.error(f"Excel parse hatası: {exc}")
            else: st.error(f"GitHub erişimi başarısız: {status.message}")

req_codes=st.session_state.get("req_codes",[]); wb_bytes=st.session_state.get("wb_bytes")
if not req_codes or not wb_bytes:
    st.warning("Analiz için fon kodu girin veya Excel yükleyin."); st.stop()

wb=openpyxl.load_workbook(io.BytesIO(wb_bytes))
st.write(f"🎯 **Analize alınan fonlar ({len(req_codes)}):** `{', '.join(req_codes)}`")

# ------------------------------------------------------------
# Analysis
# ------------------------------------------------------------
calc_funds=[]; failed=[]; progress=st.progress(0,text="Gerçek fiyat verileri kaynaklardan alınıyor...")

def worker(code):
    series, source, statuses = get_fund_series(code)
    met=compute_metrics(series,code) if series is not None else None
    if met:
        met["source"]=source; met["source_statuses"]=[asdict(s) for s in statuses]
    return code,met

with concurrent.futures.ThreadPoolExecutor(max_workers=min(MAX_WORKERS,max(1,len(req_codes)))) as executor:
    futures={executor.submit(worker,c):c for c in req_codes}
    for i,fut in enumerate(concurrent.futures.as_completed(futures),1):
        code=futures[fut]
        try:
            _,met=fut.result()
            if met and met.get("n_days",0)>=1: calc_funds.append(met)
            else: failed.append(code)
        except Exception:
            failed.append(code)
        progress.progress(i/len(req_codes),text=f"📥 {i}/{len(req_codes)} işlendi")
progress.empty()

if not calc_funds:
    st.error(f"❌ Hiçbir fon için gerçek fiyat geçmişi alınamadı. Hatalı Fonlar: {', '.join(failed)}")
    st.stop()

with st.spinner("📊 Relative momentum, risk, sentiment ve confidence hesaplanıyor..."):
    calculate_security(calc_funds)
    calculate_momentum(calc_funds)
    areas=tuple(sorted(set(f.get("investment_area","Karma") for f in calc_funds)))
    sentiments=fetch_batch_market_sentiment(areas,api_key_input)
    common_dates=calculate_daily_history(calc_funds,sentiments)
    finalize(calc_funds,sentiments)

# ------------------------------------------------------------
# Summary
# ------------------------------------------------------------
output=create_excel_output(wb,calc_funds,common_dates)
st.subheader(f"📈 KAZRİSK Portföy Özeti (V{APP_VERSION})")
c1,c2,c3,c4,c5=st.columns(5)
scores=[float(f["decision_score"]) for f in calc_funds]
c1.metric("En Yüksek",f"{max(scores):.1f}"); c2.metric("Ortalama",f"{sum(scores)/len(scores):.1f}"); c3.metric("En Düşük",f"{min(scores):.1f}"); c4.metric("Güçlü Al",sum(f["karar"]=="GÜÇLÜ AL" for f in calc_funds)); c5.metric("Acil Sat",sum(f["karar"]=="ACİL SAT" for f in calc_funds))

# ------------------------------------------------------------
# Main table + 5 day history + 2 day confirmations
# ------------------------------------------------------------
st.subheader("📊 Analiz Sonuçları — Son 5 İşlem Günü Kararları")
display=[]; early=[]
last5=list(common_dates[-5:])
for f in calc_funds:
    row={"Fon Kodu":f["code"],"Yatırım Alanı":f.get("investment_area","-")}
    smap=dict(zip(f.get("dates",[]),f.get("running_trend_hybrid",[])))
    for d in reversed(last5):
        s=smap.get(d); row[f"{pd.to_datetime(d).strftime('%d.%m.%Y')} Karar Skoru"]=s if s is not None else "Veri Açıklanmadı"; row[f"{pd.to_datetime(d).strftime('%d.%m.%Y')} Model Kararı"]=decision_label(s) if s is not None else "Veri Açıklanmadı"
    row.update({"Güncel Karar Skoru":f.get("decision_score"),"Güncel Karar":f.get("karar"),"Confidence":f.get("confidence_score"),"Veri Kalite":f.get("data_quality_score"),"Momentum":f.get("market_momentum"),"Güvenlik":f.get("security_score"),"Sentiment":f.get("sentiment_score"),"Sharpe":f.get("_final_sharpe"),"Sortino":f.get("_final_sortino"),"MaxDD (%)":f.get("_final_max_dd"),"Aktif Kaynak":f.get("source")})
    display.append(row)
    valid=[(d,s) for d,s in zip(f.get("dates",[]),f.get("running_trend_hybrid",[])) if s is not None]
    if len(valid)>=2:
        (d1,s1),(d2,s2)=valid[-2],valid[-1]
        if decision_label(s1)=="ACİL SAT" and decision_label(s2)=="ACİL SAT": early.append({"Tip":"SAT","Fon Kodu":f["code"],"KAZRİSK Durumu":"🚨 2 GÜN TEYİTLİ ACİL SAT","Son 2 Gün":f"{pd.to_datetime(d1).strftime('%d.%m.%Y')} → {pd.to_datetime(d2).strftime('%d.%m.%Y')}","Son Skor":s2})
        if decision_label(s1)=="GÜÇLÜ AL" and decision_label(s2)=="GÜÇLÜ AL": early.append({"Tip":"AL","Fon Kodu":f["code"],"KAZRİSK Durumu":"🚀 2 GÜN TEYİTLİ GÜÇLÜ AL","Son 2 Gün":f"{pd.to_datetime(d1).strftime('%d.%m.%Y')} → {pd.to_datetime(d2).strftime('%d.%m.%Y')}","Son Skor":s2})

df_display=pd.DataFrame(display)
def style_cells(v):
    t=str(v).upper()
    if "GÜÇLÜ AL" in t or "ASIL LİSTE" in t: return "color:#008000;font-weight:bold"
    if "DÜZELTME" in t: return "color:#B8860B;font-weight:bold"
    if "ACİL SAT" in t: return "color:#FF0000;font-weight:bold"
    return ""
try: st.dataframe(df_display.style.map(style_cells),use_container_width=True,hide_index=True)
except AttributeError: st.dataframe(df_display.style.applymap(style_cells),use_container_width=True,hide_index=True)

if early:
    st.subheader("🚨/🚀 2 Günlük Teyitli Alarmlar")
    ca,cb=st.columns(2)
    with ca:
        sells=[x for x in early if x["Tip"]=="SAT"]
        if sells: st.dataframe(pd.DataFrame([{k:v for k,v in x.items() if k!="Tip"} for x in sells]),use_container_width=True,hide_index=True)
        else: st.success("Teyitli ACİL SAT yok.")
    with cb:
        buys=[x for x in early if x["Tip"]=="AL"]
        if buys: st.dataframe(pd.DataFrame([{k:v for k,v in x.items() if k!="Tip"} for x in buys]),use_container_width=True,hide_index=True)
        else: st.info("Teyitli GÜÇLÜ AL yok.")
else:
    st.info("Son iki ardışık işlem gününde teyitli ACİL SAT veya GÜÇLÜ AL oluşmadı.")

# ------------------------------------------------------------
# Data quality / source diagnostics
# ------------------------------------------------------------
st.subheader("📡 Gerçek Veri Kaynakları ve Kalite")
stream=[]
for f in calc_funds:
    stream.append({"Fon":f["code"],"Aktif Kaynak":f.get("source"),"Fiyat Günü":f.get("price_days"),"Return Günü":f.get("n_days"),"Veri Kalite":f.get("data_quality_score"),"Confidence":f.get("confidence_score"),"Son Fiyat":round(f["prices"][-1],6) if f.get("prices") else None,"Not":f.get("data_quality_issues")})
st.dataframe(pd.DataFrame(stream),use_container_width=True,hide_index=True)

failed_rows=[]
for code in failed:
    failed_rows.append({"Fon":code,"Durum":"❌ Gerçek veri bulunamadı","Not":"TEFAS Direct + TEFAS Crawler + İş Yatırım denendi; sentetik veri kullanılmadı."})
if failed_rows:
    st.warning(f"⚠️ {len(failed_rows)} fon gerçek veri olmadan analiz dışı kaldı.")
    st.dataframe(pd.DataFrame(failed_rows),use_container_width=True,hide_index=True)

st.subheader("🔎 Kaynak Teşhis Paneli")
diag=[]
for f in calc_funds:
    for s in f.get("source_statuses",[]):
        diag.append({"Fon":f["code"],"Kaynak":s.get("source"),"Erişim":"✅" if s.get("ok") else "❌","HTTP":s.get("status_code") or "-","Satır":s.get("rows",0),"Süre":f"{s.get('elapsed_ms',0)} ms","Mesaj":s.get("message",""),"Kök Neden":s.get("root_cause","")})
if diag: st.dataframe(pd.DataFrame(diag),use_container_width=True,hide_index=True)

st.download_button("📥 KAZRİSK V16.3 Excel İndir",data=output,file_name="fonlar_KGDM3_KAZRISK_FINAL_V16_3.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.success(f"✅ V{APP_VERSION} analiz tamamlandı. {len(calc_funds)} fon gerçek veriyle işlendi; {len(failed)} fon veri yetersizliği nedeniyle dışarıda bırakıldı.")
