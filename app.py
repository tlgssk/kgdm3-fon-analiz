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
# Excel dosyalarındaki bilinmeyen dolgu parametresini yoksayar.
# ============================================================
original_init = PatternFill.__init__
def new_init(self, *args, **kwargs):
    if 'extLst' in kwargs:
        del kwargs['extLst']
    original_init(self, *args, **kwargs)
PatternFill.__init__ = new_init
# ============================================================

# Google GenAI SDK (Opsiyonel / Hata fırlatmaz)
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


# ============================================================
# KGDM-3 & KAZRİSK - SÜRÜM V10.0 (CANLI GEMINI SENTIMENT & KALİTE)
# ============================================================

st.set_page_config(
    page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK Hibrit Fon Analizi")
st.caption(
    "TEFAS + TEFAS Direct API + İş Yatırım + Fintables | "
    "Gemini Canlı Sentiment + Evrensel Baseline + Kalite Denetimi | V10.0"
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

APP_VERSION = "10.0.0"

GITHUB_OWNER = "tlgssk"
GITHUB_REPO = "kgdm3-fon-analiz"
GITHUB_BRANCH = "main"
GITHUB_FALLBACK_URL = (
    "https://github.com/tlgssk/kgdm3-fon-analiz/"
    "raw/refs/heads/main/"
    "Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"
)

DEFAULT_MOMENTUM_WEIGHTS = {
    "return": 0.30,
    "sharpe": 0.25,
    "cumulative": 0.25,
    "drawdown": 0.20,
}

SECURITY_WEIGHTS = {
    "aum": 0.30,
    "investor": 0.25,
    "concentration": 0.25,
    "liquidity": 0.20,
}

SECURITY_SCALE = {
    "aum": 20.0,
    "investor": 20.0,
    "aum_flow": 8.0,
    "investor_change": 6.0,
    "concentration": 20.0,
}

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
# RENKLER
# ============================================================

COLOR_NAVY = "1F4E79"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"

COLOR_LIGHT_GREEN = "E2F0D9"
COLOR_LIGHT_YELLOW = "FFF2CC"
COLOR_LIGHT_RED = "FCE4D6"


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

# API Anahtarı Yönetimi (Önce secrets/env, yoksa sidebar)
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
    help="Google AI Studio'dan aldığınız API anahtarı. Boş bırakılırsa simüle edilmiş piyasa duyarlılığı kullanılır.",
)

ENABLE_FILTERS = st.sidebar.checkbox(
    "Filtreleri Etkinleştir",
    value=False,
)

TARGET_WEEKLY_RETURN = st.sidebar.slider(
    "Hedef Haftalık Getiri (%)",
    min_value=-5.00,
    max_value=10.00,
    value=0.00,
    step=0.10,
)

MIN_INVESTOR_COUNT = st.sidebar.slider(
    "Minimum Yatırımcı Sayısı",
    min_value=0,
    max_value=100000,
    value=0,
    step=500,
)

with st.sidebar.expander("⚖️ Skor Ağırlıkları (V10.0)"):
    w_return = st.slider("Getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe-benzeri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_cumulative = st.slider("Kümülatif getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_drawdown = st.slider("Drawdown ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05)

    total_m = w_return + w_sharpe + w_cumulative + w_drawdown
    total_m = 1.0 if total_m <= 0 else total_m

    MOMENTUM_WEIGHTS = {
        "return": w_return / total_m,
        "sharpe": w_sharpe / total_m,
        "cumulative": w_cumulative / total_m,
        "drawdown": w_drawdown / total_m,
    }

    st.markdown("---")
    st.markdown("**Hibrit Karar Dağılımı**")
    w_hybrid_mom = st.slider("Momentum Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_MOMENTUM_WEIGHT, 0.05)
    w_hybrid_sec = st.slider("Güvenlik Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SECURITY_WEIGHT, 0.05)
    w_hybrid_sent = st.slider("Sentiment (Duyarlılık) Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_SENTIMENT_WEIGHT, 0.05)

    tot_h = w_hybrid_mom + w_hybrid_sec + w_hybrid_sent
    if tot_h <= 0:
        tot_h = 1.0

    HYBRID_MOMENTUM_WEIGHT = w_hybrid_mom / tot_h
    HYBRID_SECURITY_WEIGHT = w_hybrid_sec / tot_h
    HYBRID_SENTIMENT_WEIGHT = w_hybrid_sent / tot_h

with st.sidebar.expander("🔧 Tanılama"):
    SHOW_DIAGNOSTICS = st.checkbox(
        "Kaynak tanılama bilgisini göster",
        value=True,
    )


# ============================================================
# CANLI GEMINI DUYARLILIK (MARKET SENTIMENT) MOTORU
# ============================================================

@st.cache_data(ttl=60 * 60 * 4, show_spinner=False)
def fetch_market_sentiment(investment_area: str, api_key: str) -> dict:
    area = str(investment_area).strip()

    if not api_key or not GENAI_AVAILABLE:
        # API Anahtarı girilmediyse kural tabanlı simülasyon
        area_upper = area.upper()
        if "YABANCI TEKNOLOJİ" in area_upper or "YABANCI" in area_upper:
            return {"score": 38, "label": "Negatif (Kâr Satışı/Volatilite)"}
        elif "ALTIN" in area_upper or "GÜMÜŞ" in area_upper or "KIYMETLİ" in area_upper:
            return {"score": 82, "label": "Güçlü Pozitif (Faiz İndirimi Beklentisi)"}
        elif "PARA PİYASASI" in area_upper or "BORÇLANMA" in area_upper:
            return {"score": 65, "label": "Pozitif (Yüksek Sabit Getiri)"}
        elif "HİSSE" in area_upper or "BIST" in area_upper:
            return {"score": 54, "label": "Dengeli / Pozitif Beklenti"}
        return {"score": 50, "label": "Nötr / Kural Tabanlı"}

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
2. En fazla 6 kelimelik kısa bir gerekçe etiketi üret (Örn: "Faiz indirimi beklentisiyle güçlü").

Sadece ve sadece aşağıdaki JSON şemasında çıktı ver:
{{"score": 75, "label": "Gerekçe etiketi"}}
"""
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.2,
            ),
        )
        data = json.loads(response.text)
        return {
            "score": int(clamp(safe_float(data.get("score", 50)), 0.0, 100.0)),
            "label": str(data.get("label", "Nötr")),
        }
    except Exception as exc:
        return {"score": 50, "label": f"Nötr (API Gecikmesi: {str(exc)[:20]})"}


# ============================================================
# VERİ KALİTESİ + MATEMATİKSEL TUTARLILIK DENETİMİ
# ============================================================

PRICE_CONSISTENCY_TOLERANCE = 0.0005

def safe_float(value, default=0.0) -> float:
    try:
        if value is None:
            return default
        number = float(value)
        if pd.isna(number):
            return default
        return number
    except Exception:
        return default


_THOUSANDS_ONLY_RE = re.compile(r"^-?\d{1,3}(\.\d{3})+$")


def parse_number(value) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        try:
            if pd.isna(value):
                return None
        except Exception:
            pass
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = (
        text.replace("₺", "")
        .replace("TL", "")
        .replace("%", "")
        .replace(" ", "")
        .strip()
    )

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    elif "." in text and _THOUSANDS_ONLY_RE.match(text):
        text = text.replace(".", "")

    try:
        return float(text)
    except (ValueError, TypeError):
        return None


def normalize_fund_code(value) -> str:
    if value is None:
        return ""
    code = str(value).strip().upper()
    if code.endswith(".0"):
        code = code[:-2]
    return code


def format_percent(value) -> str:
    number = parse_number(value)
    if number is None:
        return "-"
    if number > 0:
        return f"+%{number:.2f}"
    if number < 0:
        return f"-%{abs(number):.2f}"
    return "%0.00"


def clamp(value, low, high):
    return max(low, min(high, value))


def calculate_compounded_return(returns) -> float:
    clean = [parse_number(v) for v in returns]
    clean = [v for v in clean if v is not None]
    if not clean:
        return 0.0
    growth = 1.0
    for r in clean:
        growth *= 1.0 + r / 100.0
    return (growth - 1.0) * 100.0


def calculate_max_drawdown(prices) -> float:
    if not prices or len(prices) < 2:
        return 0.0
    peak = safe_float(prices[0])
    max_dd = 0.0
    for price in prices:
        price = safe_float(price)
        if price <= 0:
            continue
        if price > peak:
            peak = price
        if peak > 0:
            dd = (price / peak - 1.0) * 100.0
            if dd < max_dd:
                max_dd = dd
    return max_dd


def zscore(values) -> List[float]:
    clean = [safe_float(v) for v in values if v is not None and not pd.isna(safe_float(v))]
    if len(clean) < 2:
        return [0.0] * len(values)
    mean_value = sum(clean) / len(clean)
    variance = sum((x - mean_value) ** 2 for x in clean) / len(clean)
    std = variance ** 0.5
    if std <= 1e-12:
        return [0.0] * len(values)
    return [
        clamp((safe_float(v) - mean_value) / std, -Z_LIMIT, Z_LIMIT) if v is not None else 0.0
        for v in values
    ]


def population_mean_std(values: List[Optional[float]]) -> Tuple[float, float]:
    valid = [v for v in values if v is not None]
    if len(valid) < 2:
        return 0.0, 0.0
    mean_v = sum(valid) / len(valid)
    var = sum((v - mean_v) ** 2 for v in valid) / len(valid)
    return mean_v, var ** 0.5


def zscore_against_population(value: Optional[float], mean_v: float, std_v: float) -> float:
    if value is None or std_v <= 1e-12:
        return 0.0
    return clamp((value - mean_v) / std_v, -Z_LIMIT, Z_LIMIT)


def calculate_valor_penalty(excess_valor) -> float:
    excess_valor = safe_float(excess_valor)
    if excess_valor <= 0:
        return 0.0
    normalized = clamp(excess_valor / 3.0, 0.0, 1.0)
    return normalized * MAX_VALOR_PENALTY


def validate_price_series(fund: dict) -> Dict[str, Any]:
    dates = fund.get("dates") or []
    prices = fund.get("prices") or []
    returns = fund.get("daily_returns") or []
    issues = []

    if len(prices) < 2:
        issues.append("Yetersiz fiyat gözlemi")
    if dates and any(dates[i] >= dates[i + 1] for i in range(len(dates) - 1)):
        issues.append("Tarih sırası sorunu")
    if any((safe_float(p) is None or safe_float(p) <= 0) for p in prices):
        issues.append("Pozitif olmayan fiyat")
    if returns and len(returns) != max(0, len(prices) - 1):
        issues.append("Getiri-fiyat uzunluk uyumsuzluğu")

    if len(prices) >= 2 and len(returns) >= len(prices) - 1:
        for i in range(len(prices) - 1):
            p0 = safe_float(prices[i])
            p1 = safe_float(prices[i + 1])
            r = safe_float(returns[i])
            if p0 > 0 and p1 > 0 and r is not None:
                expected = (p1 / p0 - 1.0) * 100.0
                if abs(expected - r) > (PRICE_CONSISTENCY_TOLERANCE * 100.0):
                    issues.append("Getiri-fiyat tutarsızlığı")
                    break

    return {
        "ok": not issues,
        "issues": issues,
        "n_dates": len(dates),
        "n_prices": len(prices),
        "n_returns": len(returns),
    }


def validate_structural_data(fund: dict) -> Dict[str, Any]:
    issues = []
    top_weight = fund.get("top_asset_weight")
    weights = [safe_float(top_weight)] if top_weight else []
    hhi = fund.get("asset_class_hhi")

    if fund.get("structural_fetch_ok") and not weights:
        issues.append("Yapısal kaynak başarılı fakat dağılım yok")

    return {
        "ok": not issues,
        "issues": issues,
        "distribution_count": len(weights),
        "distribution_sum": sum(weights) if weights else None,
        "hhi": hhi,
    }


def audit_fund_data(fund: dict) -> dict:
    price = validate_price_series(fund)
    structural = validate_structural_data(fund)

    score = 100.0
    score -= 20 if not price["ok"] else 0
    score -= 15 if fund.get("n_days", 0) < TARGET_TRADING_DAYS else 0
    score -= 15 if not fund.get("structural_fetch_ok", False) else 0
    score -= 10 if fund.get("aum") is None and fund.get("investors") is None else 0
    score -= 10 if "Liste-bağıl" in str(fund.get("reference_scope", "")) else 0

    issues = price["issues"] + structural["issues"]
    if not fund.get("source"):
        issues.append("Fiyat kaynağı yok")
    if fund.get("structural_error"):
        issues.append(str(fund.get("structural_error")))

    fund["price_data_audit"] = price
    fund["structural_data_audit"] = structural
    fund["structural_hhi"] = structural["hhi"]
    fund["data_quality_score"] = int(round(clamp(score, 0, 100)))
    fund["data_quality_issues"] = " | ".join(dict.fromkeys(issues)) if issues else "OK"
    return fund


# ============================================================
# HTTP OTURUMU
# ============================================================

@dataclass
class SourceStatus:
    source: str
    attempted: bool = False
    ok: bool = False
    status_code: Optional[int] = None
    error_type: str = ""
    message: str = ""
    elapsed_ms: Optional[int] = None
    retry_count: int = 0


def new_status(source: str) -> SourceStatus:
    return SourceStatus(source=source)


def build_http_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=REQUEST_MAX_RETRIES,
        connect=REQUEST_MAX_RETRIES,
        read=REQUEST_MAX_RETRIES,
        status=REQUEST_MAX_RETRIES,
        backoff_factor=REQUEST_BACKOFF_FACTOR,
        status_forcelist=(429, 500, 502, 503, 504),
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=MAX_WORKERS,
        pool_maxsize=MAX_WORKERS,
    )
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": "KGDM3-Fon-Analiz/10.0 (+live-sentiment)",
        "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
    })
    return session


HTTP = build_http_session()


def request_with_status(
    source: str,
    method: str,
    url: str,
    *,
    params: Optional[dict] = None,
    data: Optional[dict] = None,
    headers: Optional[dict] = None,
    timeout: int = HTTP_TIMEOUT,
) -> Tuple[Optional[requests.Response], SourceStatus]:

    status = new_status(source)
    status.attempted = True
    started = time.perf_counter()

    try:
        response = HTTP.request(
            method=method,
            url=url,
            params=params,
            data=data,
            headers=headers,
            timeout=timeout,
        )
        status.status_code = response.status_code
        status.elapsed_ms = int((time.perf_counter() - started) * 1000)

        if response.status_code == 200:
            status.ok = True
            status.message = "OK"
        elif response.status_code == 403:
            status.error_type = "HTTP_403"
            status.message = "Erişim reddedildi"
        elif response.status_code == 429:
            status.error_type = "HTTP_429"
            status.message = "Rate limit"
        elif 500 <= response.status_code <= 599:
            status.error_type = f"HTTP_{response.status_code}"
            status.message = "Sunucu hatası"
        else:
            status.error_type = f"HTTP_{response.status_code}"
            status.message = f"HTTP {response.status_code}"

        return response, status

    except requests.Timeout:
        status.error_type = "TIMEOUT"
        status.message = "Zaman aşımı"
    except requests.ConnectionError:
        status.error_type = "CONNECTION_ERROR"
        status.message = "Bağlantı hatası"
    except Exception as exc:
        status.error_type = "ERROR"
        status.message = str(exc)[:200]

    status.elapsed_ms = int((time.perf_counter() - started) * 1000)
    return None, status


# ============================================================
# TEFAS API VE YARDIMCI SCRAPING
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
        crawler = Crawler(timeout=60, max_retry=3)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is None or df.empty:
            return pd.DataFrame()

        df.rename(
            columns={
                "fund_code": "code",
                "fund_name": "title",
                "investor_count": "investors",
                "portfolio_size": "aum",
                "fund_type": "kind",
            },
            inplace=True,
        )

        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["price"] = df["price"].apply(parse_number)
        df["aum"] = df["aum"].apply(parse_number).fillna(0.0) if "aum" in df.columns else 0.0
        df["investors"] = df["investors"].apply(parse_number).fillna(0.0) if "investors" in df.columns else 0.0
        df["code"] = df["code"].astype(str).str.strip().str.upper()

        df = df.dropna(subset=["date", "code", "price"])
        return df[df["price"] > 0].sort_values(["code", "date"]).drop_duplicates(subset=["code", "date"], keep="last").reset_index(drop=True)
    except Exception:
        return pd.DataFrame()


def build_fund_meta_map(universe: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    meta = {}
    if universe is not None and not universe.empty:
        latest = universe.sort_values("date").drop_duplicates(subset=["code"], keep="last")
        for _, row in latest.iterrows():
            code = str(row.get("code", "")).strip().upper()
            if code:
                meta[code] = {
                    "kind": str(row.get("kind", DEFAULT_FUND_KIND)),
                    "title": str(row.get("title", "")),
                }
    return meta


def build_universe_reference(universe: pd.DataFrame, window: int) -> Dict[str, Dict[str, List[float]]]:
    # Evrensel Baseline için AUM ve Yatırımcı eklendi
    ref = {
        k: {
            "mean_return": [],
            "sharpe": [],
            "cumulative": [],
            "max_dd_inv": [],
            "aum": [],
            "investors": [],
        }
        for k in FUND_KINDS
    }
    if universe is None or universe.empty or window < 2:
        return ref

    latest = universe.sort_values("date").drop_duplicates(subset=["code"], keep="last")
    for _, row in latest.iterrows():
        kind_str = str(row.get("kind", DEFAULT_FUND_KIND)).strip().upper()
        if kind_str in ref:
            if safe_float(row.get("aum")) > 0:
                ref[kind_str]["aum"].append(safe_float(row.get("aum")))
            if safe_float(row.get("investors")) > 0:
                ref[kind_str]["investors"].append(safe_float(row.get("investors")))

    for code, group in universe.groupby("code"):
        group = group.sort_values("date")
        kind = str(group["kind"].iloc[-1]).strip().upper()
        if kind not in FUND_KINDS:
            continue

        prices = group["price"].astype(float).tolist()
        if len(prices) < window + 1:
            continue

        w_prices = prices[-(window + 1):]
        rets = [
            0.0 if p0 <= 0 else (p1 / p0 - 1.0) * 100.0
            for p0, p1 in zip(w_prices[:-1], w_prices[1:])
        ]
        mean_r = sum(rets) / len(rets)
        vol = (sum((r - mean_r) ** 2 for r in rets) / len(rets)) ** 0.5
        ref[kind]["mean_return"].append(mean_r)
        ref[kind]["sharpe"].append(mean_r / vol if vol > 1e-12 else 0.0)
        ref[kind]["cumulative"].append((w_prices[-1] / w_prices[0] - 1.0) * 100.0)
        ref[kind]["max_dd_inv"].append(calculate_max_drawdown(w_prices))

    return ref


def reference_sample_size(ref, kind) -> int:
    return len(ref.get(kind, {}).get("mean_return", []))


def fetch_isyatirim_series(fund_code: str) -> Tuple[Optional[pd.DataFrame], SourceStatus]:
    code = normalize_fund_code(fund_code)
    status = new_status("İş Yatırım")
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}

    response, status = request_with_status("İş Yatırım", "GET", url, params=params, headers={"Accept": "application/json"})
    if response and status.ok:
        try:
            df = pd.DataFrame(response.json().get("value", []))
            df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
            df["price"] = df["Fiyat"].apply(parse_number)
            df["aum"], df["investors"] = 0.0, 0.0
            df = df.dropna(subset=["date", "price"])
            df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
            if len(df) >= 2:
                return df[["date", "price", "aum", "investors"]], status
        except Exception:
            pass
    return None, status


def fetch_tefas_direct_api(fund_code: str, fund_kind: Optional[str] = None) -> Tuple[Optional[pd.DataFrame], SourceStatus]:
    code = normalize_fund_code(fund_code)
    status = new_status("TEFAS Direct API")
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {"X-Requested-With": "XMLHttpRequest", "Origin": "https://www.tefas.gov.tr"}

    kinds = [fund_kind] if fund_kind in FUND_KINDS else list(FUND_KINDS)
    for kind in kinds:
        payload = {"fontip": kind, "fonkod": code, "bastarih": start.strftime("%d.%m.%Y"), "bittarih": end.strftime("%d.%m.%Y")}
        res, stat = request_with_status("TEFAS Direct API", "POST", url, data=payload, headers=headers)
        if res and stat.ok:
            try:
                df = pd.DataFrame(res.json().get("data", []))
                df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
                df["price"] = df["FIYAT"].apply(parse_number)
                df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number).fillna(0.0) if "PORTFOYBUYUKLUK" in df.columns else 0.0
                df["investors"] = df["KISISAYISI"].apply(parse_number).fillna(0.0) if "KISISAYISI" in df.columns else 0.0
                df = df.dropna(subset=["date", "price"])[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
                if len(df) >= 2:
                    return df, stat
            except Exception:
                pass
    return None, status


@st.cache_data(show_spinner=False, ttl=60 * 60 * 2)
def fetch_tefas_breakdown_snapshot(fund_kind: Optional[str], reference_date: Optional[str]) -> dict:
    kind = (fund_kind or "YAT").upper()
    try:
        ref = pd.to_datetime(reference_date).date() if reference_date else dt.date.today()
    except Exception:
        ref = dt.date.today()
    try:
        from pytefas import Crawler
    except ImportError as exc:
        return {"ok": False, "error": str(exc), "rows": {}}

    crawler = Crawler(timeout=60, max_retry=3)
    for offset in range(0, 8):
        q_date = ref - dt.timedelta(days=offset)
        try:
            df = crawler.fetch(start=q_date, end=q_date, columns="breakdown", kind=kind)
        except Exception:
            continue
        if df is not None and not df.empty:
            rows = {}
            for _, row in df.iterrows():
                c = normalize_fund_code(row.get("fund_code"))
                if c:
                    rows[c] = {
                        col: parse_number(row.get(col))
                        for col in df.columns
                        if col not in ("fund_code", "fund_name", "date", "kind")
                        and parse_number(row.get(col)) is not None
                    }
            return {"ok": True, "source": "TEFAS", "rows": rows}
    return {"ok": False, "error": "Veri Yok", "rows": {}}


def fetch_fund_structural_data(fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {
        "top_asset_weight": None,
        "asset_class_hhi": None,
        "is_bist30": False,
        "emergency_cash_ratio": None,
        "cash_ratio_known": False,
        "structural_fetch_ok": False,
        "structural_source": "YOK",
        "investment_area": "-",
    }

    t_upper = (fund_title or "").upper()
    if "PARA PİYASASI" in t_upper or "PPF" in t_upper:
        structural["investment_area"] = "Para Piyasası"
    elif "ALTIN" in t_upper or "GÜMÜŞ" in t_upper or "KIYMETLİ" in t_upper:
        structural["investment_area"] = "Kıymetli Maden"
    elif "YABANCI TEKNOLOJİ" in t_upper:
        structural["investment_area"] = "Hisse Senedi (Yabancı Teknoloji)"
    elif "HİSSE" in t_upper:
        structural["investment_area"] = "Hisse Senedi"
    elif "BORÇLANMA" in t_upper:
        structural["investment_area"] = "Borçlanma Araçları"
    elif "DEĞİŞKEN" in t_upper or "KARMA" in t_upper:
        structural["investment_area"] = "Karma / Değişken"

    if "BIST 30" in t_upper or "BIST30" in t_upper:
        structural["is_bist30"] = True

    snapshot = fetch_tefas_breakdown_snapshot(fund_kind, None)
    if snapshot.get("ok"):
        row = snapshot.get("rows", {}).get(code)
        if row:
            allocs = [v for k, v in row.items() if v is not None and v > 0]
            if allocs:
                total = sum(allocs)
                structural["top_asset_weight"] = max(allocs) / total * 100.0 if total > 0 else None
                structural["asset_class_hhi"] = sum((x / total) ** 2 for x in allocs) * 100 if total > 0 else None
                cash_keys = ["takasbank_money_market_pct", "repo_pct", "reverse_repo_pct", "term_deposit_pct"]
                cash_val = sum(safe_float(row.get(k)) for k in cash_keys if row.get(k) is not None)
                if cash_val > 0:
                    structural["emergency_cash_ratio"] = clamp(cash_val, 0.0, 100.0)
                    structural["cash_ratio_known"] = True
                structural["structural_fetch_ok"] = True
                structural["structural_source"] = "TEFAS"

    return structural


# ============================================================
# METRİK & HESAPLAMALAR
# ============================================================

def get_fund_series(universe: pd.DataFrame, fund_code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(fund_code)
    statuses = []
    if universe is not None and not universe.empty:
        rows = universe[universe["code"].eq(code)].copy()
        if len(rows) >= 2:
            statuses.append(new_status("TEFAS"))
            return rows.tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True), "TEFAS", statuses

    df_dir, stat_dir = fetch_tefas_direct_api(code, fund_kind)
    statuses.append(stat_dir)
    if df_dir is not None:
        return df_dir, "TEFAS Direct API", statuses

    df_is, stat_is = fetch_isyatirim_series(code)
    statuses.append(stat_is)
    if df_is is not None:
        return df_is, "İş Yatırım", statuses

    return None, "YOK", statuses


def compute_fund_metrics(series: pd.DataFrame, fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None):
    if series is None or len(series) < 2:
        return None
    df = series.copy().sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    prices = df["price"].tolist()
    dates = df["date"].dt.strftime("%d.%m").tolist()
    aums = df["aum"].tolist()
    invs = df["investors"].tolist()
    rets = [
        0.0 if prices[i - 1] <= 0 else (prices[i] / prices[i - 1] - 1.0) * 100.0
        for i in range(1, len(prices))
    ]
    if not rets:
        return None

    struct = fetch_fund_structural_data(fund_code, fund_kind, fund_title)

    return {
        "code": fund_code,
        "dates": dates[1:],
        "prices": prices,
        "daily_returns": rets,
        "n_days": len(rets),
        "aum": aums[-1],
        "investors": int(invs[-1]),
        "aum_change": (aums[-1] / aums[0] - 1.0) * 100.0 if aums[0] > 0 else None,
        "aum_flow_proxy": (
            ((aums[-1] / aums[0] - 1.0) * 100.0) - ((prices[-1] / prices[0] - 1.0) * 100.0)
        ) if aums[0] > 0 else None,
        "inv_change": (invs[-1] / invs[0] - 1.0) * 100.0 if invs[0] > 0 else None,
        "max_dd": calculate_max_drawdown(prices),
        "weekly_return": calculate_compounded_return(rets[-5:]),
        "fund_title": fund_title or "-",
        **struct,
    }


def fetch_and_compute_one_fund(code: str, universe: pd.DataFrame, meta_map: dict, valor_dict: dict):
    meta = meta_map.get(code, {})
    series, source, statuses = get_fund_series(universe, code, meta.get("kind"))
    metrics = compute_fund_metrics(series, code, meta.get("kind"), meta.get("title"))
    if metrics is None:
        return code, None, source
    metrics["valor"] = valor_dict.get(code, 0.0)
    metrics["source"] = source
    metrics["kind"] = meta.get("kind", DEFAULT_FUND_KIND)
    metrics["source_statuses"] = [asdict(x) for x in statuses]
    metrics["source_chain"] = " → ".join(x.source for x in statuses if x.attempted)
    return code, metrics, source


def calculate_security_scores(funds: List[dict], reference: dict):
    by_kind = defaultdict(list)
    for idx, fund in enumerate(funds):
        by_kind[fund.get("kind", DEFAULT_FUND_KIND)].append(idx)

    for kind, indices in by_kind.items():
        subset = [funds[i] for i in indices]

        # Evrensel (Global Baseline) Verilerini Çek
        ref = reference.get(kind, {})
        aum_m, aum_s = population_mean_std(ref.get("aum", []))
        inv_m, inv_s = population_mean_std(ref.get("investors", []))

        flow_z = zscore([f.get("aum_flow_proxy") for f in subset])
        inv_c_z = zscore([f.get("inv_change") for f in subset])

        for local_i, fund_idx in enumerate(indices):
            f = funds[fund_idx]
            
            # YENİ: Sadece yüklenen listeye değil, TÜM TEFAS'a göre Z-Skoru hesapla
            local_aum_z = zscore_against_population(f.get("aum"), aum_m, aum_s) if aum_s > 0 else zscore([x.get("aum") for x in subset])[local_i]
            local_inv_z = zscore_against_population(f.get("investors"), inv_m, inv_s) if inv_s > 0 else zscore([x.get("investors") for x in subset])[local_i]

            s = (
                50.0
                + SECURITY_SCALE["aum"] * SECURITY_WEIGHTS["aum"] * local_aum_z
                + SECURITY_SCALE["investor"] * SECURITY_WEIGHTS["investor"] * local_inv_z
            )

            if f.get("aum_flow_proxy") is not None:
                s += SECURITY_SCALE["aum_flow"] * flow_z[local_i]
            if f.get("inv_change") is not None:
                s += SECURITY_SCALE["investor_change"] * inv_c_z[local_i]

            hhi = f.get("structural_hhi")
            if hhi is not None and hhi > 25.0:
                s -= min((hhi - 25.0) * 0.35, MAX_CONCENTRATION_PENALTY)

            if f.get("is_bist30", False):
                s += BIST30_BONUS
            cash = f.get("emergency_cash_ratio")
            if f.get("cash_ratio_known", False) and cash is not None:
                if cash >= 15:
                    s += HIGH_LIQUIDITY_BONUS
                elif cash < 5:
                    s -= LOW_LIQUIDITY_PENALTY

            f["security_score"] = int(round(clamp(s, 0.0, 100.0)))


def calculate_market_relative_momentum(funds: List[dict], reference, window: int):
    for f in funds:
        k = f.get("kind", DEFAULT_FUND_KIND)
        rets = f["daily_returns"][-window:]
        prc = f["prices"][-(window + 1):]
        if not rets:
            f["market_momentum"] = None
            continue

        m_r = sum(rets) / len(rets)
        vol = (sum((x - m_r) ** 2 for x in rets) / len(rets)) ** 0.5
        cum = (prc[-1] / prc[0] - 1.0) * 100.0 if prc[0] > 0 else 0.0
        dd = calculate_max_drawdown(prc)

        f["_final_mean_return"] = m_r
        f["_final_sharpe"] = m_r / vol if vol > 1e-12 else 0.0
        f["_final_cumulative"] = cum
        f["_final_max_dd"] = dd
        f["volatility"] = vol

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
            f["reference_scope"] = f"Piyasa ({k})"
        else:
            fb = [x for x in funds if x.get("kind") == k]
            idx = fb.index(f)
            zm = zscore([x.get("_final_mean_return") for x in fb])[idx]
            zs = zscore([x.get("_final_sharpe") for x in fb])[idx]
            zc = zscore([x.get("_final_cumulative") for x in fb])[idx]
            zd = zscore([-safe_float(x.get("_final_max_dd")) for x in fb])[idx]
            f["reference_scope"] = "Liste-bağıl"

        wz = (
            MOMENTUM_WEIGHTS["return"] * zm
            + MOMENTUM_WEIGHTS["sharpe"] * zs
            + MOMENTUM_WEIGHTS["cumulative"] * zc
            + MOMENTUM_WEIGHTS["drawdown"] * zd
        )
        mom = clamp(50.0 + 20.0 * wz, 0.0, 100.0)

        last_d = rets[-1]
        last_2 = sum(rets[-2:]) / 2.0 if len(rets) >= 2 else last_d
        oh = zc >= OVERHEAT_Z_THRESHOLD and (last_d < 0 or last_2 < 0)
        f["overheat_flag"] = oh
        if oh:
            mom = clamp(mom - OVERHEAT_PENALTY, 0.0, 100.0)

        f["market_momentum"] = int(round(mom))


def calculate_trend_scores(funds: List[dict], api_key: str = "") -> int:
    if not funds:
        return 0
    n_days = min(f["n_days"] for f in funds)
    for f in funds:
        f["dates"] = f["dates"][-n_days:]
        f["daily_returns"] = f["daily_returns"][-n_days:]
        f["prices"] = f["prices"][-(n_days + 1):]
        f["running_trend_momentum"] = []

    for d in range(1, n_days + 1):
        if d < MIN_ROLLING_DAYS:
            for f in funds:
                f["running_trend_momentum"].append(None)
            continue
        cur = []
        for f in funds:
            r = f["daily_returns"][d - MIN_ROLLING_DAYS:d]
            p = f["prices"][d - MIN_ROLLING_DAYS:d + 1]
            if len(r) < MIN_ROLLING_DAYS:
                continue
            mr = sum(r) / len(r)
            vol = (sum((x - mr) ** 2 for x in r) / len(r)) ** 0.5
            cur.append({
                "fund": f,
                "mr": mr,
                "sh": mr / vol if vol > 1e-12 else 0.0,
                "cm": calculate_compounded_return(r),
                "dd": calculate_max_drawdown(p),
            })
        if not cur:
            continue

        zm = zscore([x["mr"] for x in cur])
        zs = zscore([x["sh"] for x in cur])
        zc = zscore([x["cm"] for x in cur])
        zd = zscore([-x["dd"] for x in cur])

        for i, data in enumerate(cur):
            wz = (
                MOMENTUM_WEIGHTS["return"] * zm[i]
                + MOMENTUM_WEIGHTS["sharpe"] * zs[i]
                + MOMENTUM_WEIGHTS["cumulative"] * zc[i]
                + MOMENTUM_WEIGHTS["drawdown"] * zd[i]
            )
            data["fund"]["running_trend_momentum"].append(int(round(clamp(50.0 + 20.0 * wz, 0.0, 100.0))))

    for f in funds:
        sec = safe_float(f.get("security_score"), 50.0)
        sent_data = fetch_market_sentiment(f.get("investment_area"), api_key)
        sent = sent_data["score"]

        run_h = []
        for m in f["running_trend_momentum"]:
            if m is None:
                run_h.append(None)
            else:
                hybrid_val = (
                    m * HYBRID_MOMENTUM_WEIGHT
                    + sec * HYBRID_SECURITY_WEIGHT
                    + sent * HYBRID_SENTIMENT_WEIGHT
                )
                run_h.append(int(round(clamp(hybrid_val, 0.0, 100.0))))
        f["running_trend_hybrid"] = run_h
        val_l = [s for s in run_h if s is not None][-5:]
        f["last_5_scores_str"] = " ➔ ".join(str(x) for x in val_l) if val_l else "-"
        if val_l:
            weights = [EMA_DECAY ** (len(val_l) - 1 - i) for i in range(len(val_l))]
            f["trend_skor"] = int(round(sum(s * w for s, w in zip(val_l, weights)) / sum(weights)))
        else:
            f["trend_skor"] = None
    return n_days


def decision_label_from_score(score) -> str:
    if score is None:
        return "YETERSİZ VERİ"
    score = safe_float(score)
    if score >= STRONG_BUY:
        return "GÜÇLÜ AL"
    if score >= WATCH_LIST:
        return "ASIL LİSTE"
    if score >= CORRECTION:
        return "DÜZELTME / İZLE"
    return "ACİL SAT"


def finalize_decisions(funds: List[dict], api_key: str = ""):
    for f in funds:
        mom = f.get("market_momentum")
        sec = f.get("security_score")
        if mom is None or sec is None:
            f["decision_score"] = None
            f["karar"] = "YETERSİZ VERİ"
            continue

        sent_data = fetch_market_sentiment(f.get("investment_area"), api_key)
        sent = sent_data["score"]
        f["sentiment_score"] = sent
        f["sentiment_label"] = sent_data["label"]

        dec = int(round(clamp(
            mom * HYBRID_MOMENTUM_WEIGHT
            + sec * HYBRID_SECURITY_WEIGHT
            + sent * HYBRID_SENTIMENT_WEIGHT,
            0.0,
            100.0,
        )))
        f["decision_score"] = dec
        f["karar"] = decision_label_from_score(dec)


def compute_confidence_label(fund: dict) -> str:
    score = calculate_confidence_score(fund)
    if score >= 80:
        return f"🟢 Yüksek ({score})"
    if score >= 60:
        return f"🟡 Orta ({score})"
    return f"🔴 Düşük ({score})"


def calculate_confidence_score(fund: dict) -> int:
    score = 0.0
    if fund.get("n_days", 0) >= TARGET_TRADING_DAYS:
        score += 20
    elif fund.get("n_days", 0) >= MIN_ROLLING_DAYS:
        score += 12
    if fund.get("source") == "TEFAS":
        score += 25
    if fund.get("structural_fetch_ok", False):
        score += 20
    if fund.get("aum") is not None and safe_float(fund.get("aum")) > 0:
        score += 5
    return int(round(clamp(score, 0, 100)))


# ============================================================
# EXCEL ÇIKTISI (D SÜTUNU BUGÜN TARİHLİ & DİNAMİK RENKLENDİRME)
# ============================================================

def create_excel_output(wb, ws_list, all_funds, common_n_days):
    if "KGDM3_Puanlama" in wb.sheetnames:
        del wb["KGDM3_Puanlama"]
    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

    headers = [
        "Fon Kodu", "Fon Adı", "Yatırım Alanı", "Valör", "Karar Skoru (Piyasa-Bağıl)", "Trend Skoru (Liste-Bağıl)",
        "Piyasa Momentum", "Güvenlik/Likidite Skoru", "Sentiment Skoru", "Duyarlılık Yönü", "Referans Kapsamı", "Veri Kalitesi", "Aşırı Isınma",
        "Son 5 Trend Skoru", "Model Kararı", "Ort. Günlük Getiri (%)", "Volatilite (%)", "Sharpe-benzeri",
        "Kümülatif Getiri (%)", "MaxDD (%)", "En Büyük Varlık (%)", "BIST30", "Net Likidite (%)", "KAZRİSK",
        "AUM Değişim (%)", "AUM Akış Proxy (%)", "Yatırımcı Değişim (%)", "AUM (₺)", "Yatırımcı",
        "Haftalık Bileşik (%)", "Veri Kaynağı", "Kalite Uyarıları"
    ]

    today_dt = dt.date.today()
    n_dates = common_n_days if common_n_days > 0 else 5
    sample_dates = [(today_dt - dt.timedelta(days=(n_dates - 1 - i))).strftime("%d.%m") for i in range(n_dates)]

    # D Sütunu için Bugün -> Dün -> Önceki Gün sıralaması
    last_5_dates = list(reversed(sample_dates[-5:]))

    daily_headers = []
    for day in last_5_dates:
        daily_headers.extend([f"{day} Karar Skoru", f"{day} Model Kararı"])
    headers[3:3] = daily_headers

    for day in sample_dates:
        headers.append(f"{day} Trend Skor")
    for day in sample_dates:
        headers.append(f"{day} Getiri")

    ws_scores.append(headers)
    header_index = {name: idx + 1 for idx, name in enumerate(headers)}

    fill = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    font = Font(name="Calibri", bold=True, color=COLOR_WHITE)
    for cell in ws_scores[1]:
        cell.fill, cell.font, cell.alignment = fill, font, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_scores.row_dimensions[1].height = 55

    for item in all_funds:
        top_asset = item.get("top_asset_weight")
        risk_label = "⚪ Veri Yok" if top_asset is None else ("⚠️ Yüksek" if top_asset > 30 else ("🟡 Orta" if top_asset > 15 else "🛡️ Dengeli"))

        row_data = [item["code"], item.get("fund_title") or "-", item.get("investment_area") or "-"]

        daily_scores = (item.get("running_trend_hybrid") or [])[-5:]
        daily_scores = [None] * max(0, len(last_5_dates) - len(daily_scores)) + daily_scores

        for s in reversed(daily_scores):
            row_data.extend([s if s is not None else "", decision_label_from_score(s) if s is not None else ""])

        row_data.extend([
            item.get("valor", 0), item.get("decision_score"), item.get("trend_skor"), item.get("market_momentum"),
            item.get("security_score"), item.get("sentiment_score"), item.get("sentiment_label"), item.get("reference_scope", "-"), compute_confidence_label(item),
            "🔥 Evet" if item.get("overheat_flag") else "-", item.get("last_5_scores_str", "-"), item.get("karar", "-"),
            round(safe_float(item.get("_final_mean_return")), 4), round(safe_float(item.get("volatility")), 4),
            round(safe_float(item.get("_final_sharpe")), 4), round(safe_float(item.get("_final_cumulative")), 4),
            round(safe_float(item.get("_final_max_dd")), 4), round(safe_float(top_asset), 2) if top_asset else None,
            "EVET" if item.get("is_bist30", False) else "HAYIR",
            f"%{safe_float(item.get('emergency_cash_ratio')):.2f}" if item.get("cash_ratio_known") else "Veri Yok",
            risk_label, round(safe_float(item.get("aum_change")), 2), round(safe_float(item.get("aum_flow_proxy")), 2),
            round(safe_float(item.get("inv_change")), 2), round(safe_float(item.get("aum")), 0), item.get("investors"),
            round(safe_float(item.get("weekly_return")), 4), item.get("source", "-"), item.get("data_quality_issues", "")
        ])

        own_scores = (item.get("running_trend_hybrid") or [])
        own_scores = [None] * (len(sample_dates) - len(own_scores)) + own_scores if len(own_scores) < len(sample_dates) else own_scores[-len(sample_dates):]
        row_data.extend([s if s is not None else "" for s in own_scores])

        own_rets = (item.get("daily_returns") or [])
        own_rets = [None] * (len(sample_dates) - len(own_rets)) + own_rets if len(own_rets) < len(sample_dates) else own_rets[-len(sample_dates):]
        row_data.extend([format_percent(x) if x is not None else "-" for x in own_rets])

        ws_scores.append(row_data)

    green_font, red_font, yellow_font = Font(bold=True, color=COLOR_GREEN), Font(bold=True, color=COLOR_RED), Font(bold=True, color=COLOR_YELLOW)
    decision_cols = [idx for name, idx in header_index.items() if "Karar" in name and "Skor" not in name]

    for row_number in range(2, ws_scores.max_row + 1):
        for col_idx in decision_cols:
            cell = ws_scores.cell(row=row_number, column=col_idx)
            text = str(cell.value or "").upper()
            if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text:
                cell.font = green_font
            elif "DÜZELTME" in text:
                cell.font = yellow_font
            elif "ACİL SAT" in text or "YETERSİZ" in text:
                cell.font = red_font

    score_cols = [idx for name, idx in header_index.items() if "Skor" in name]
    for col_idx in score_cols:
        col_letter = get_column_letter(col_idx)
        rng = f"{col_letter}2:{col_letter}{ws_scores.max_row}"
        ws_scores.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["75"], fill=PatternFill(start_color=COLOR_LIGHT_GREEN, fill_type="solid")))
        ws_scores.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["50", "74"], fill=PatternFill(start_color=COLOR_LIGHT_YELLOW, fill_type="solid")))
        ws_scores.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["50"], fill=PatternFill(start_color=COLOR_LIGHT_RED, fill_type="solid")))

    cur_col, int_col = header_index.get("AUM (₺)"), header_index.get("Yatırımcı")
    pct_cols = [
        "Ort. Günlük Getiri (%)", "Volatilite (%)", "Kümülatif Getiri (%)", "MaxDD (%)",
        "En Büyük Varlık (%)", "Net Likidite (%)", "AUM Değişim (%)", "AUM Akış Proxy (%)",
        "Yatırımcı Değişim (%)", "Haftalık Bileşik (%)"
    ]

    for row_number in range(2, ws_scores.max_row + 1):
        if cur_col:
            ws_scores.cell(row=row_number, column=cur_col).number_format = '#,##0.00 "₺"'
        if int_col:
            ws_scores.cell(row=row_number, column=int_col).number_format = "#,##0"
        for col_name in pct_cols:
            idx = header_index.get(col_name)
            if idx and isinstance(ws_scores.cell(row=row_number, column=idx).value, (int, float)):
                ws_scores.cell(row=row_number, column=idx).number_format = '0.00"%"'

    thin = Side(style="thin", color="D9E1F2")
    for row in ws_scores.iter_rows():
        for cell in row:
            cell.alignment, cell.border = Alignment(vertical="center"), Border(bottom=thin)

    ws_scores.freeze_panes = "A2"
    ws_scores.sheet_view.showGridLines = False

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# ANA ARAYÜZ (STREAMLIT)
# ============================================================

col_upload, col_github = st.columns(2)
wb = None

with col_upload:
    uploaded_file = st.file_uploader("Bilgisayardan Excel Yükle", type=["xlsx"])
    if uploaded_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_file)
        except Exception as exc:
            st.error(f"Excel yükleme hatası: {exc}")

with col_github:
    if st.button("🚀 GitHub'dan Çek ve Analiz Et", use_container_width=True):
        url = GITHUB_FALLBACK_URL
        res, stat = request_with_status("GitHub", "GET", url)
        if res and stat.ok:
            wb = openpyxl.load_workbook(io.BytesIO(res.content))
            st.success("✅ Veri çekildi.")

if wb is None:
    st.stop()

ws_list = wb["Fon_Listesi"] if "Fon_Listesi" in wb.sheetnames else wb.active
req_codes = [normalize_fund_code(r[0].value) for r in ws_list.iter_rows(min_row=2) if r and r[0].value]
req_codes = list(dict.fromkeys(filter(None, req_codes)))

if not req_codes:
    st.stop()

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
        try:
            _, met, src = fut.result()
        except Exception:
            met = None
        if met:
            calc_funds.append(met)
        else:
            failed.append(c)
        prog.progress((i + 1) / len(req_codes))
prog.empty()

eligible = [f for f in calc_funds if f.get("n_days", 0) >= MIN_ROLLING_DAYS]

with st.spinner("📊 V10 Modeli (Gemini Canlı Sentiment + Baseline) Hesaplanıyor..."):
    for f in eligible:
        audit_fund_data(f)
    calculate_security_scores(eligible, ref)
    calculate_market_relative_momentum(eligible, ref, TARGET_TRADING_DAYS)
    common_n = calculate_trend_scores(eligible, api_key_input)
    finalize_decisions(eligible, api_key_input)

output = create_excel_output(wb, ws_list, eligible, common_n)

# ============================================================
# SKOR ÖZETLERİ VE EKRAN TABLOSU
# ============================================================

st.subheader("📈 KAZRİSK Portföy Özeti (V10.0)")
col1, col2, col3, col4 = st.columns(4)
scores = [safe_float(x.get("decision_score")) for x in eligible if x.get("decision_score") is not None]
if scores:
    col1.metric("En Yüksek Skor", f"{max(scores):.0f}")
    col2.metric("Ortalama Skor", f"{sum(scores) / len(scores):.1f}")
    col3.metric("En Düşük Skor", f"{min(scores):.0f}")
    col4.metric("Güçlü Al Veren", sum(1 for x in eligible if x.get("karar") == "GÜÇLÜ AL"))

display_rows = []
early_alerts = []
today_dt = dt.date.today()

for item in eligible:
    top_asset = item.get("top_asset_weight")
    risk_label = "⚪ Veri Yok" if top_asset is None else ("⚠️ Yüksek Konsantrasyon" if top_asset > 30 else ("🟡 Orta Konsantrasyon" if top_asset > 15 else "🛡️ Dengeli"))

    row_dict = {
        "Fon Kodu": item["code"],
        "Fon Adı": item.get("fund_title") or "-",
        "Yatırım Alanı": item.get("investment_area") or "-",
    }

    own_scores = item.get("running_trend_hybrid") or []
    last_5_s = own_scores[-5:] if len(own_scores) >= 5 else own_scores
    last_5_dates_web = [(today_dt - dt.timedelta(days=i)).strftime("%d.%m") for i in range(len(last_5_s))]

    for day, score in zip(last_5_dates_web, reversed(last_5_s)):
        row_dict[f"{day} Karar Skoru"] = score if score is not None else ""
        row_dict[f"{day} Model Kararı"] = decision_label_from_score(score)

    row_dict.update({
        "Sentiment Skoru": item.get("sentiment_score"),
        "Duyarlılık Yönü": item.get("sentiment_label"),
        "Güncel Karar Skoru": item.get("decision_score"),
        "Trend Skoru": item.get("trend_skor"),
        "Güncel Karar": item.get("karar"),
        "Net Likidite (%)": f"%{safe_float(item.get('emergency_cash_ratio')):.2f}" if item.get("cash_ratio_known") else "Veri Yok",
        "KAZRİSK Konsantrasyon": risk_label,
        "Haftalık Getiri (%)": round(safe_float(item.get("weekly_return")), 2),
        "Veri Kalite Skoru": item.get("data_quality_score"),
    })
    display_rows.append(row_dict)

    if len(last_5_s) >= 2:
        lbls = [decision_label_from_score(s) for s in last_5_s]
        if lbls[-1] == "ACİL SAT" and lbls[-2] == "ACİL SAT":
            early_alerts.append({
                "Tip": "SAT", "Fon Kodu": item["code"], "Fon Adı": item.get("fund_title"), "Alan": item.get("investment_area"),
                "KAZRİSK Durumu": "🚨 2 GÜN TEYİTLİ ACİL SAT", "Son Skor": last_5_s[-1]
            })
        elif lbls[-1] == "GÜÇLÜ AL" and lbls[-2] == "GÜÇLÜ AL":
            early_alerts.append({
                "Tip": "AL", "Fon Kodu": item["code"], "Fon Adı": item.get("fund_title"), "Alan": item.get("investment_area"),
                "KAZRİSK Durumu": "🚀 2 GÜN TEYİTLİ GÜÇLÜ AL", "Son Skor": last_5_s[-1]
            })

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value).upper()
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "🟢" in text or "DENGELİ" in text:
        return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text or "🟡" in text or "ORTA KONSANTRASYON" in text:
        return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text or "YETERSİZ" in text or "🔴" in text or "YÜKSEK KONSANTRASYON" in text:
        return "color: #FF0000; font-weight: bold;"
    return ""

try:
    styled_df = df_display.style.map(color_cells)
except AttributeError:
    styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 Analiz Sonuçları (V10.0)")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

# ============================================================
# ALARM TABLOLARI (SATIŞ VE ALIM YAN YANA)
# ============================================================
sell_alerts = [{k: v for k, v in a.items() if k != "Tip"} for a in early_alerts if a["Tip"] == "SAT"]
buy_alerts = [{k: v for k, v in a.items() if k != "Tip"} for a in early_alerts if a["Tip"] == "AL"]

if sell_alerts or buy_alerts:
    st.subheader("🚨/🚀 KAZRİSK® 2 Günlük Teyitli Alarmlar")
    col_alert1, col_alert2 = st.columns(2)

    with col_alert1:
        st.markdown("### 🚨 Satış Alarmları")
        if sell_alerts:
            st.dataframe(pd.DataFrame(sell_alerts), use_container_width=True, hide_index=True)
        else:
            st.info("Şu an teyitli 'Acil Sat' sinyali veren fon yok.")

    with col_alert2:
        st.markdown("### 🚀 Fırsat Alarmları")
        if buy_alerts:
            st.dataframe(pd.DataFrame(buy_alerts), use_container_width=True, hide_index=True)
        else:
            st.success("Şu an teyitli 'Güçlü Al' fırsatı veren fon yok.")

st.success(f"✅ V10.0 Analiz tamamlandı. Toplam {len(eligible)} fon işlendi.")
st.download_button(
    label="📥 KAZRİSK V10.0 Excel İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_FINAL_V10_0.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if SHOW_DIAGNOSTICS:
    st.subheader("🔎 Veri Kaynağı Tanılaması")
    diagnostic_rows = []
    for item in eligible:
        for status in item.get("source_statuses", []):
            diagnostic_rows.append({
                "Fon": item["code"],
                "Kaynak": status.get("source"),
                "Denendi": "Evet" if status.get("attempted") else "Hayır",
                "Başarılı": "Evet" if status.get("ok") else "Hayır",
                "HTTP": status.get("status_code"),
                "Hata": status.get("error_type"),
                "Mesaj": status.get("message"),
                "Süre ms": status.get("elapsed_ms"),
            })
    if diagnostic_rows:
        st.dataframe(pd.DataFrame(diagnostic_rows), use_container_width=True, hide_index=True)
