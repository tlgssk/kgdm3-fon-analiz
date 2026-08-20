import concurrent.futures
import datetime as dt
import io
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

from openpyxl.styles.fills import PatternFill

# openpyxl 'extLst' hatasını çözmek için yama (Monkey Patch)
original_init = PatternFill.__init__
def new_init(self, *args, **kwargs):
    if 'extLst' in kwargs:
        del kwargs['extLst']
    original_init(self, *args, **kwargs)
PatternFill.__init__ = new_init

# ============================================================
# KGDM-3 & KAZRİSK - GELİŞTİRİLMİŞ SÜRÜM
# ============================================================
# Bu sürüm:
# - Kaynakları ayrı ayrı izler.
# - 403/429/5xx/timeout gibi durumları gizlemez.
# - Anti-bot / CAPTCHA / WAF aşmaya çalışmaz.
# - Başlık üzerinden NUMERİK portföy tahmini yapmaz.
# - Gerçek veri ile tahmini/sınıflandırılmış veriyi ayırır.
# - AUM akışını "proxy" olarak tutar; gerçek net para akışı olarak sunmaz.
# - Veri güvenini 0-100 arasında hesaplar.
# - İsteklerde sınırlı, kontrollü retry/backoff kullanır.
# - Excel çıktısına kaynak tanılamasını ekler.
# ============================================================


st.set_page_config(
    page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK Hibrit Fon Analizi")
st.caption(
    "TEFAS + TEFAS Direct API + İş Yatırım + Fintables | "
    "Kaynak Tanılama + Momentum + Risk + Likidite | V8.1"
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

# Aynı kaynağa agresif istek göndermemek için.
REQUEST_MAX_RETRIES = 2
REQUEST_BACKOFF_FACTOR = 1.5

MIN_REFERENCE_SAMPLE = 5
OVERHEAT_Z_THRESHOLD = 2.0
OVERHEAT_PENALTY = 6.0

APP_VERSION = "8.1.0"

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
    "aum_flow": 8.0,          # proxy etkisi azaltıldı
    "investor_change": 6.0,
    "concentration": 20.0,
}

DEFAULT_HYBRID_MOMENTUM_WEIGHT = 0.60

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

with st.sidebar.expander("⚖️ Skor Ağırlıkları"):
    st.caption(
        "Varsayılan ağırlıklar sezgiseldir; geçmiş performansla kalibre edilmiş "
        "bir yatırım tavsiyesi modeli değildir."
    )

    w_return = st.slider(
        "Getiri ağırlığı", 0.0, 1.0,
        DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05
    )
    w_sharpe = st.slider(
        "Sharpe-benzeri ağırlığı", 0.0, 1.0,
        DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05
    )
    w_cumulative = st.slider(
        "Kümülatif getiri ağırlığı", 0.0, 1.0,
        DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05
    )
    w_drawdown = st.slider(
        "Drawdown ağırlığı", 0.0, 1.0,
        DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05
    )

    total = w_return + w_sharpe + w_cumulative + w_drawdown
    if total <= 0:
        total = 1.0

    MOMENTUM_WEIGHTS = {
        "return": w_return / total,
        "sharpe": w_sharpe / total,
        "cumulative": w_cumulative / total,
        "drawdown": w_drawdown / total,
    }

    hybrid_momentum_w = st.slider(
        "Hibrit skorda Momentum ağırlığı",
        0.0, 1.0,
        DEFAULT_HYBRID_MOMENTUM_WEIGHT,
        0.05,
    )
    HYBRID_MOMENTUM_WEIGHT = hybrid_momentum_w
    HYBRID_SECURITY_WEIGHT = 1.0 - hybrid_momentum_w

with st.sidebar.expander("🔧 Tanılama"):
    SHOW_DIAGNOSTICS = st.checkbox(
        "Kaynak tanılama bilgisini göster",
        value=True,
    )


# ============================================================
# VERİ KAYNAĞI DURUMU
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


# ============================================================
# HTTP OTURUMU
# ============================================================

def build_http_session() -> requests.Session:
    """
    Kontrollü retry/backoff.
    Anti-bot/WAF/CAPTCHA aşma yapmaz.
    """
    session = requests.Session()

    retry = Retry(
        total=REQUEST_MAX_RETRIES,
        connect=REQUEST_MAX_RETRIES,
        read=REQUEST_MAX_RETRIES,
        status=REQUEST_MAX_RETRIES,
        backoff_factor=REQUEST_BACKOFF_FACTOR,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"}),
        respect_retry_after_header=True,
        raise_on_status=False,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=MAX_WORKERS,
        pool_maxsize=MAX_WORKERS,
    )

    session.mount("https://", adapter)
    session.mount("http://", adapter)

    session.headers.update({
        "User-Agent": "KGDM3-Fon-Analiz/8.0 (+source-diagnostics)",
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
    json_body: Optional[dict] = None,
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
            json=json_body,
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
            status.message = "Rate limit / çok fazla istek"
        elif 500 <= response.status_code <= 599:
            status.error_type = f"HTTP_{response.status_code}"
            status.message = "Sunucu hatası"
        else:
            status.error_type = f"HTTP_{response.status_code}"
            status.message = f"HTTP {response.status_code}"

        return response, status

    except requests.Timeout:
        status.error_type = "TIMEOUT"
        status.message = "İstek zaman aşımına uğradı"
    except requests.ConnectionError:
        status.error_type = "CONNECTION_ERROR"
        status.message = "Bağlantı hatası"
    except requests.RequestException as exc:
        status.error_type = "REQUEST_ERROR"
        status.message = str(exc)[:200]
    except Exception as exc:
        status.error_type = "UNEXPECTED_ERROR"
        status.message = str(exc)[:200]

    status.elapsed_ms = int((time.perf_counter() - started) * 1000)
    return None, status


# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

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
    clean = []

    for value in values:
        if value is None:
            clean.append(None)
            continue

        try:
            value = float(value)
            clean.append(None if pd.isna(value) else value)
        except Exception:
            clean.append(None)

    valid = [x for x in clean if x is not None]

    if len(valid) < 2:
        return [0.0 for _ in clean]

    mean_value = sum(valid) / len(valid)
    variance = sum((x - mean_value) ** 2 for x in valid) / len(valid)
    std = variance ** 0.5

    if std <= 1e-12:
        return [0.0 for _ in clean]

    result = []

    for value in clean:
        if value is None:
            result.append(0.0)
            continue

        result.append(
            clamp(
                (value - mean_value) / std,
                -Z_LIMIT,
                Z_LIMIT,
            )
        )

    return result


def population_mean_std(values: List[Optional[float]]) -> Tuple[float, float]:
    valid = [v for v in values if v is not None]

    if len(valid) < 2:
        return 0.0, 0.0

    mean_v = sum(valid) / len(valid)
    var = sum((v - mean_v) ** 2 for v in valid) / len(valid)

    return mean_v, var ** 0.5


def zscore_against_population(
    value: Optional[float],
    mean_v: float,
    std_v: float,
) -> float:

    if value is None or std_v <= 1e-12:
        return 0.0

    return clamp(
        (value - mean_v) / std_v,
        -Z_LIMIT,
        Z_LIMIT,
    )


def calculate_valor_penalty(excess_valor) -> float:
    excess_valor = safe_float(excess_valor)

    if excess_valor <= 0:
        return 0.0

    normalized = clamp(excess_valor / 3.0, 0.0, 1.0)

    return normalized * MAX_VALOR_PENALTY


# ============================================================
# GITHUB
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def resolve_latest_github_excel_url() -> Optional[str]:

    api_url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/contents/"
        f"?ref={GITHUB_BRANCH}"
    )

    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "KGDM3-Fon-Analiz/8.0",
    }

    response, status = request_with_status(
        "GitHub",
        "GET",
        api_url,
        headers=headers,
    )

    if response is None or not status.ok:
        return None

    try:
        items = response.json()
    except Exception:
        return None

    if not isinstance(items, list):
        return None

    xlsx_files = [
        item
        for item in items
        if (
            isinstance(item, dict)
            and str(item.get("name", "")).lower().endswith(".xlsx")
            and item.get("download_url")
        )
    ]

    if not xlsx_files:
        return None

    date_pattern = re.compile(r"(\d{4}-\d{2}-\d{2})")

    def sort_key(item):
        match = date_pattern.search(item.get("name", ""))
        return (
            match.group(1) if match else "",
            item.get("name", ""),
        )

    xlsx_files.sort(key=sort_key, reverse=True)

    return xlsx_files[0]["download_url"]


# ============================================================
# TEFAS UNIVERSE
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(
    start_date: dt.date,
    end_date: dt.date,
) -> pd.DataFrame:

    try:
        from pytefas import Crawler
    except ImportError:
        return pd.DataFrame()

    try:
        crawler = Crawler(
            timeout=60,
            max_retry=3,
        )

        df = crawler.fetch_many(
            start=start_date,
            end=end_date,
            kinds=FUND_KINDS,
            columns="info",
        )

        if df is None or df.empty:
            return pd.DataFrame()

        df = df.copy()

        rename_map = {
            "fund_code": "code",
            "fund_name": "title",
            "investor_count": "investors",
            "portfolio_size": "aum",
            "fund_type": "kind",
            "kind": "kind",
        }

        df.rename(
            columns={
                k: v
                for k, v in rename_map.items()
                if k in df.columns
            },
            inplace=True,
        )

        if "date" in df.columns:
            df["date"] = pd.to_datetime(
                df["date"],
                errors="coerce",
            )

        if "price" in df.columns:
            df["price"] = df["price"].apply(parse_number)

        if "aum" in df.columns:
            df["aum"] = (
                df["aum"]
                .apply(parse_number)
                .fillna(0.0)
            )
        else:
            df["aum"] = 0.0

        if "investors" in df.columns:
            df["investors"] = (
                df["investors"]
                .apply(parse_number)
                .fillna(0.0)
            )
        else:
            df["investors"] = 0.0

        if "kind" not in df.columns:
            df["kind"] = DEFAULT_FUND_KIND

        if "title" not in df.columns:
            df["title"] = ""

        df["code"] = (
            df["code"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df = df.dropna(
            subset=["date", "code", "price"]
        )

        df = df[df["price"] > 0]

        return (
            df.sort_values(["code", "date"])
            .drop_duplicates(
                subset=["code", "date"],
                keep="last",
            )
            .reset_index(drop=True)
        )

    except Exception:
        return pd.DataFrame()


def build_fund_meta_map(
    universe: pd.DataFrame,
) -> Dict[str, Dict[str, str]]:

    meta = {}

    if universe is None or universe.empty:
        return meta

    try:
        latest = (
            universe
            .sort_values("date")
            .drop_duplicates(
                subset=["code"],
                keep="last",
            )
        )

        for _, row in latest.iterrows():

            code = str(
                row.get("code", "")
            ).strip().upper()

            if not code:
                continue

            kind = str(
                row.get("kind", "")
            ).strip().upper()

            title = str(
                row.get("title", "") or ""
            ).strip()

            if title.lower() in (
                "",
                "nan",
                "none",
            ):
                title = ""

            meta[code] = {
                "kind": (
                    kind
                    if kind in FUND_KINDS
                    else DEFAULT_FUND_KIND
                ),
                "title": title,
            }

    except Exception:
        pass

    return meta


def build_universe_reference(
    universe: pd.DataFrame,
    window: int,
) -> Dict[str, Dict[str, List[float]]]:

    reference = {
        k: {
            "mean_return": [],
            "sharpe": [],
            "cumulative": [],
            "max_dd_inv": [],
        }
        for k in FUND_KINDS
    }

    if (
        universe is None
        or universe.empty
        or "kind" not in universe.columns
        or window < 2
    ):
        return reference

    for code, group in universe.groupby("code"):

        group = group.sort_values("date")

        kind = str(
            group["kind"].iloc[-1]
        ).strip().upper()

        if kind not in FUND_KINDS:
            continue

        prices = (
            group["price"]
            .astype(float)
            .tolist()
        )

        if len(prices) < window + 1:
            continue

        window_prices = prices[-(window + 1):]

        returns = [
            0.0 if p0 <= 0
            else (p1 / p0 - 1.0) * 100.0
            for p0, p1 in zip(
                window_prices[:-1],
                window_prices[1:],
            )
        ]

        mean_return = (
            sum(returns) / len(returns)
        )

        variance = (
            sum(
                (r - mean_return) ** 2
                for r in returns
            )
            / len(returns)
        )

        volatility = variance ** 0.5

        sharpe = (
            mean_return / volatility
            if volatility > 1e-12
            else 0.0
        )

        cumulative = (
            window_prices[-1]
            / window_prices[0]
            - 1.0
        ) * 100.0

        max_dd = calculate_max_drawdown(
            window_prices
        )

        reference[kind]["mean_return"].append(
            mean_return
        )
        reference[kind]["sharpe"].append(
            sharpe
        )
        reference[kind]["cumulative"].append(
            cumulative
        )
        reference[kind]["max_dd_inv"].append(
            -max_dd
        )

    return reference


def reference_sample_size(
    reference,
    kind,
) -> int:

    if kind not in reference:
        return 0

    return len(
        reference[kind].get(
            "mean_return",
            [],
        )
    )


# ============================================================
# İŞ YATIRIM
# ============================================================

def fetch_isyatirim_series(
    fund_code: str,
) -> Tuple[Optional[pd.DataFrame], SourceStatus]:

    code = normalize_fund_code(fund_code)
    status = new_status("İş Yatırım")

    if not code:
        return None, status

    end = dt.datetime.now()
    start = (
        end
        - dt.timedelta(
            days=LOOKBACK_CALENDAR_DAYS
        )
    )

    url = (
        "https://www.isyatirim.com.tr/"
        "_layouts/15/IsYatirim.Website/Common/"
        "Data.aspx/YatirimFonGecmisGetiri"
    )

    params = {
        "fonKod": code,
        "baslangic": start.strftime(
            "%d-%m-%Y"
        ),
        "bitis": end.strftime(
            "%d-%m-%Y"
        ),
    }

    response, status = request_with_status(
        "İş Yatırım",
        "GET",
        url,
        params=params,
        headers={
            "Accept": "application/json",
        },
    )

    if response is None or not status.ok:
        return None, status

    try:
        payload = response.json()
        values = payload.get("value")

        if not values:
            status.error_type = "EMPTY_DATA"
            status.message = "Boş veri"
            return None, status

        df = pd.DataFrame(values)

        if (
            "Tarih" not in df.columns
            or "Fiyat" not in df.columns
        ):
            status.error_type = "SCHEMA_ERROR"
            status.message = "Beklenen kolonlar yok"
            return None, status

        df["date"] = pd.to_datetime(
            df["Tarih"],
            dayfirst=True,
            errors="coerce",
        )

        df["price"] = df["Fiyat"].apply(
            parse_number
        )

        # İş Yatırım bu endpoint'te AUM/yatırımcı
        # sağlamıyorsa bunları 0 kabul ediyoruz.
        df["aum"] = 0.0
        df["investors"] = 0.0

        df = df.dropna(
            subset=["date", "price"]
        )

        df = df[df["price"] > 0]

        if len(df) < 2:
            status.error_type = "INSUFFICIENT_DATA"
            status.message = "En az 2 fiyat gözlemi gerekli"
            return None, status

        status.ok = True
        status.message = (
            f"{len(df)} ham gözlem"
        )

        return (
            df.sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .tail(
                TARGET_TRADING_DAYS + 1
            )
            .reset_index(drop=True)[
                [
                    "date",
                    "price",
                    "aum",
                    "investors",
                ]
            ],
            status,
        )

    except ValueError:
        status.error_type = "JSON_ERROR"
        status.message = "JSON çözümlenemedi"
    except Exception as exc:
        status.error_type = "PARSE_ERROR"
        status.message = str(exc)[:200]

    return None, status


# ============================================================
# TEFAS DIRECT API
# ============================================================

def fetch_tefas_direct_api(
    fund_code: str,
    fund_kind: Optional[str] = None,
) -> Tuple[Optional[pd.DataFrame], SourceStatus]:

    code = normalize_fund_code(fund_code)

    status = new_status("TEFAS Direct API")

    if not code:
        return None, status

    end = dt.datetime.now()
    start = (
        end
        - dt.timedelta(
            days=LOOKBACK_CALENDAR_DAYS
        )
    )

    url = (
        "https://www.tefas.gov.tr/"
        "api/DB/BindHistoryInfo"
    )

    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.tefas.gov.tr",
        "Referer": "https://www.tefas.gov.tr/",
    }

    kind_candidates = []

    if fund_kind in FUND_KINDS:
        kind_candidates.append(fund_kind)

    kind_candidates += [
        k for k in FUND_KINDS
        if k not in kind_candidates
    ]

    last_status = status

    for kind in kind_candidates:

        payload = {
            "fontip": kind,
            "fonkod": code,
            "bastarih": start.strftime(
                "%d.%m.%Y"
            ),
            "bittarih": end.strftime(
                "%d.%m.%Y"
            ),
        }

        response, attempt_status = request_with_status(
            "TEFAS Direct API",
            "POST",
            url,
            data=payload,
            headers=headers,
        )

        last_status = attempt_status

        if response is None or not attempt_status.ok:
            continue

        try:
            data = response.json().get(
                "data",
                [],
            )

            if not data:
                last_status.error_type = "EMPTY_DATA"
                last_status.message = (
                    f"{kind}: boş veri"
                )
                continue

            df = pd.DataFrame(data)

            if not all(
                c in df.columns
                for c in [
                    "TARIH",
                    "FIYAT",
                ]
            ):
                last_status.error_type = "SCHEMA_ERROR"
                last_status.message = (
                    f"{kind}: beklenen kolonlar yok"
                )
                continue

            df["date"] = pd.to_datetime(
                df["TARIH"],
                unit="ms",
                errors="coerce",
            )

            df["price"] = df["FIYAT"].apply(
                parse_number
            )

            if "PORTFOYBUYUKLUK" in df.columns:
                df["aum"] = (
                    df["PORTFOYBUYUKLUK"]
                    .apply(parse_number)
                    .fillna(0.0)
                )
            else:
                df["aum"] = 0.0

            if "KISISAYISI" in df.columns:
                df["investors"] = (
                    df["KISISAYISI"]
                    .apply(parse_number)
                    .fillna(0.0)
                )
            else:
                df["investors"] = 0.0

            df = df.dropna(
                subset=["date", "price"]
            )

            df = df[df["price"] > 0]

            if len(df) < 2:
                last_status.error_type = (
                    "INSUFFICIENT_DATA"
                )
                continue

            last_status.ok = True
            last_status.message = (
                f"{kind}: {len(df)} ham gözlem"
            )

            return (
                df.sort_values("date")
                .drop_duplicates(
                    subset=["date"],
                    keep="last",
                )
                .tail(
                    TARGET_TRADING_DAYS + 1
                )
                .reset_index(drop=True),
                last_status,
            )

        except Exception as exc:
            last_status.error_type = "PARSE_ERROR"
            last_status.message = str(exc)[:200]

    return None, last_status


# ============================================================
# FINTABLES - SADECE GERÇEK YAPISAL VERİ
# ============================================================

@st.cache_data(
    show_spinner=False,
    ttl=60 * 60 * 2,
)

@st.cache_data(
    show_spinner=False,
    ttl=60 * 60 * 2,
)
def fetch_tefas_breakdown_snapshot(
    fund_kind: Optional[str],
    reference_date: Optional[str],
) -> dict:
    """
    TEFAS'ın yeni resmi JSON portföy dağılımı endpoint'inden tek tarih /
    fon tipi için toplu snapshot alır.

    Aynı snapshot tüm fonlar tarafından paylaşıldığı için 670 fon için
    670 ayrı HTTP isteği yapılmaz; fon tipi başına cache'lenmiş tek istek
    kullanılır.
    """
    kind = (fund_kind or "YAT").upper()
    if kind not in ("YAT", "EMK", "BYF", "GYF", "GSYF"):
        kind = "YAT"

    if reference_date:
        try:
            ref = pd.to_datetime(reference_date).date()
        except Exception:
            ref = dt.date.today()
    else:
        ref = dt.date.today()

    # TEFAS resmi endpoint'i yaklaşık bir aylık aralık kabul ediyor.
    # Burada tek gün sorguladığımız için aralık problemi yok.
    url = "https://www.tefas.gov.tr/api/funds/dagilimSiraliGetirT"

    body = {
        "fonTipi": kind,
        "fonKodu": None,
        "aramaMetni": None,
        "fonTurKod": None,
        "fonGrubu": None,
        "sfonTurKod": None,
        "fonTurAciklama": None,
        "kurucuKod": None,
        "basTarih": ref.strftime("%Y%m%d"),
        "bitTarih": ref.strftime("%Y%m%d"),
        "basSira": 1,
        "bitSira": 100000,
        "dil": "TR",
        "sFonTurKod": "",
        "fonKod": "",
        "fonGrup": "",
        "fonUnvanTip": "",
    }

    response, status = request_with_status(
        "TEFAS Direct Structural API",
        "POST",
        url,
        json_body=body,
        headers={
            "Accept": "*/*",
            "Content-Type": "application/json",
            "Origin": "https://www.tefas.gov.tr",
            "Referer": "https://www.tefas.gov.tr/tr/fon-verileri",
        },
        timeout=max(HTTP_TIMEOUT, 30),
    )

    if response is None or not status.ok:
        return {
            "ok": False,
            "source": "TEFAS Direct Structural API",
            "date": ref.isoformat(),
            "kind": kind,
            "error": f"{status.error_type}: {status.message}",
            "rows": {},
        }

    try:
        payload = response.json()
    except Exception as exc:
        return {
            "ok": False,
            "source": "TEFAS Direct Structural API",
            "date": ref.isoformat(),
            "kind": kind,
            "error": f"JSON_PARSE_ERROR: {str(exc)[:180]}",
            "rows": {},
        }

    err_code = payload.get("errorCode")
    err_msg = payload.get("errorMessage")
    if err_code or err_msg:
        return {
            "ok": False,
            "source": "TEFAS Direct Structural API",
            "date": ref.isoformat(),
            "kind": kind,
            "error": f"TEFAS_API_ERROR: {err_msg or err_code}",
            "rows": {},
        }

    result_list = payload.get("resultList") or []
    rows = {}

    # TEFAS API kısa alanları -> anlamlı yüzde alanları.
    field_map = {
        "hs": "stock_pct",
        "dt": "government_bond_pct",
        "hb": "treasury_bill_pct",
        "fb": "financing_bill_pct",
        "ost": "private_sector_bond_pct",
        "bb": "bank_bill_pct",
        "vdm": "asset_backed_securities_pct",
        "eut": "eurobond_pct",
        "kibd": "government_external_debt_pct",
        "osdb": "private_sector_external_debt_pct",
        "kba": "fx_government_internal_debt_pct",
        "dot": "fx_payable_bill_pct",
        "db": "fx_payable_bond_pct",
        "tpp": "takasbank_money_market_pct",
        "bpp": "bist_money_market_pct",
        "btaa": "bist_committed_buy_pct",
        "btas": "bist_committed_sell_pct",
        "r": "repo_pct",
        "tr": "reverse_repo_pct",
        "vm": "term_deposit_pct",
        "vmtl": "deposit_tl_pct",
        "vmd": "deposit_fx_pct",
        "vmau": "deposit_gold_pct",
        "kh": "participation_account_pct",
        "khtl": "participation_account_tl_pct",
        "khd": "participation_account_fx_pct",
        "khau": "participation_account_gold_pct",
        "kks": "government_lease_certificate_pct",
        "kkstl": "government_lease_certificate_tl_pct",
        "kksd": "government_lease_certificate_fx_pct",
        "kksyd": "government_foreign_lease_certificate_pct",
        "osks": "private_sector_lease_certificate_pct",
        "oksyd": "private_sector_foreign_lease_certificate_pct",
        "km": "precious_metals_pct",
        "kmbyf": "precious_metals_etf_pct",
        "kmkba": "precious_metals_government_debt_pct",
        "kmkks": "precious_metals_lease_certificate_pct",
        "ymk": "foreign_security_pct",
        "yba": "foreign_debt_security_pct",
        "ybkb": "foreign_government_debt_pct",
        "ybosb": "foreign_private_sector_debt_pct",
        "yhs": "foreign_stock_pct",
        "ybyf": "foreign_etf_pct",
        "fkb": "fund_participation_certificate_pct",
        "yyf": "investment_fund_pct",
        "byf": "etf_pct",
        "gykb": "real_estate_fund_pct",
        "gyy": "real_estate_investment_pct",
        "gsykb": "venture_capital_fund_pct",
        "gsyy": "venture_capital_investment_pct",
        "t": "derivative_pct",
        "vint": "futures_cash_collateral_pct",
        "gas": "real_estate_certificate_pct",
        "d": "other_pct",
    }

    for row in result_list:
        code = normalize_fund_code(
            row.get("fonKodu")
            or row.get("fonKod")
        )
        if not code:
            continue

        parsed = {
            "fund_code": code,
            "fund_name": row.get("fonUnvan"),
            "date": row.get("tarih"),
        }

        for short, target in field_map.items():
            parsed[target] = parse_number(row.get(short))

        rows[code] = parsed

    return {
        "ok": bool(rows),
        "source": "TEFAS Direct Structural API",
        "date": ref.isoformat(),
        "kind": kind,
        "error": "" if rows else "NO_STRUCTURAL_ROWS",
        "rows": rows,
    }


@st.cache_data(
    show_spinner=False,
    ttl=60 * 60 * 2,
)
def fetch_fund_structural_data(
    fund_code: str,
    fund_kind: Optional[str] = None,
    fund_title: Optional[str] = None,
    reference_date: Optional[str] = None,
) -> dict:

    code = normalize_fund_code(fund_code)

    structural = {
        "top_asset_weight": None,
        "top_asset_weight_basis": None,
        "is_bist30": False,
        "is_bist30_known": False,
        "emergency_cash_ratio": None,
        "cash_ratio_known": False,
        "structural_fetch_ok": False,
        "structural_source": "YOK",
        "structural_estimated": False,
        "structural_error": "",
        "structural_date": reference_date,
        "fund_title": fund_title if fund_title else None,
        "investment_area": None,
        "investment_area_source": "YOK",
    }

    if not code:
        structural["structural_error"] = "Geçersiz fon kodu"
        return structural

    # --------------------------------------------------------
    # BAŞLIKTAN SADECE SINIFLANDIRMA
    # --------------------------------------------------------
    title_upper = (fund_title or "").upper()

    if (
        "PARA PİYASASI" in title_upper
        or "PPF" in title_upper
        or "LİKİT" in title_upper
    ):
        structural["investment_area"] = "Para Piyasası"
        structural["investment_area_source"] = "TEFAS başlığı / sınıflandırma"
    elif (
        "ALTIN" in title_upper
        or "KIYMETLİ MADEN" in title_upper
        or "GÜMÜŞ" in title_upper
    ):
        structural["investment_area"] = "Kıymetli Maden"
        structural["investment_area_source"] = "TEFAS başlığı / sınıflandırma"
    elif "BIST 30" in title_upper or "BIST30" in title_upper:
        structural["is_bist30"] = True
        structural["is_bist30_known"] = True
        structural["investment_area"] = "Hisse Senedi"
        structural["investment_area_source"] = "TEFAS başlığı / sınıflandırma"
    elif "HİSSE SENEDİ" in title_upper:
        structural["investment_area"] = "Hisse Senedi"
        structural["investment_area_source"] = "TEFAS başlığı / sınıflandırma"

    # --------------------------------------------------------
    # 1) TEFAS YENİ RESMİ JSON PORTFÖY DAĞILIMI
    #
    # Kritik düzeltme:
    # Eski V8.0 Fintables HTML'sini ana kaynak kabul ediyordu.
    # Yeni TEFAS API'si doğrudan 50+ varlık sınıfı yüzdesi döndürüyor.
    # Aynı tarih/fon tipi snapshot'ı cache'lendiği için 670 fon
    # için 670 ayrı istek yapılmaz.
    # --------------------------------------------------------
    snapshot = fetch_tefas_breakdown_snapshot(
        fund_kind or "YAT",
        reference_date,
    )

    if snapshot.get("ok"):
        row = snapshot.get("rows", {}).get(code)

        if row:
            pct_fields = [
                k for k, v in row.items()
                if k.endswith("_pct") and v is not None
            ]

            valid_allocations = [
                (field, safe_float(row.get(field)))
                for field in pct_fields
                if row.get(field) is not None
                and safe_float(row.get(field)) >= 0
            ]

            if valid_allocations:
                top_field, top_value = max(
                    valid_allocations,
                    key=lambda x: x[1],
                )

                structural["top_asset_weight"] = top_value
                structural["top_asset_weight_basis"] = (
                    f"TEFAS varlık sınıfı: {top_field}"
                )

            # TEFAS doğrudan "nakit" isimli tek alan vermiyor.
            # Likit / nakit-benzeri kalemler açık alanlardan toplanıyor.
            cash_fields = {
                "takasbank_money_market_pct",
                "bist_money_market_pct",
                "bist_committed_buy_pct",
                "bist_committed_sell_pct",
                "repo_pct",
                "reverse_repo_pct",
                "term_deposit_pct",
                "deposit_tl_pct",
                "deposit_fx_pct",
                "deposit_gold_pct",
                "participation_account_pct",
                "participation_account_tl_pct",
                "participation_account_fx_pct",
                "participation_account_gold_pct",
                "futures_cash_collateral_pct",
            }

            cash_values = [
                safe_float(row.get(field))
                for field in cash_fields
                if row.get(field) is not None
            ]

            if cash_values:
                structural["emergency_cash_ratio"] = clamp(
                    sum(cash_values),
                    0.0,
                    100.0,
                )
                structural["cash_ratio_known"] = True

            # TEFAS'tan gerçek yapısal veri geldi.
            structural["structural_fetch_ok"] = True
            structural["structural_source"] = "TEFAS Direct Structural API"
            structural["structural_estimated"] = False
            structural["structural_date"] = (
                row.get("date") or snapshot.get("date")
            )

            # Başlıktan gelen BIST30 bilgisi ayrı tutulur; API'nin
            # varlık sınıfı dağılımı BIST30'u tek başına doğrulamaz.
            if (
                structural["investment_area"] is None
                and structural["top_asset_weight_basis"]
            ):
                basis = structural["top_asset_weight_basis"]
                if "stock_pct" in basis:
                    structural["investment_area"] = "Hisse Senedi"
                    structural["investment_area_source"] = (
                        "TEFAS portföy dağılımı"
                    )
                elif "precious_metals" in basis:
                    structural["investment_area"] = "Kıymetli Maden"
                    structural["investment_area_source"] = (
                        "TEFAS portföy dağılımı"
                    )
                elif "foreign_stock" in basis:
                    structural["investment_area"] = (
                        "Hisse Senedi (Yabancı)"
                    )
                    structural["investment_area_source"] = (
                        "TEFAS portföy dağılımı"
                    )
                elif "government_bond" in basis or "bond" in basis:
                    structural["investment_area"] = (
                        "Borçlanma Araçları"
                    )
                    structural["investment_area_source"] = (
                        "TEFAS portföy dağılımı"
                    )

            # En büyük varlık sınıfı var ama tek tek hisse/menkul
            # kıymet yoğunlaşması yok. Bu ayrımı açıkça kaydet.
            structural["structural_error"] = (
                "TEFAS varlık sınıfı dağılımı bulundu; "
                "tekil menkul kıymet Top-N yoğunlaşması TEFAS public API'de yok."
            )
            return structural

        structural["structural_error"] = (
            f"TEFAS snapshot bulundu fakat {code} fonu "
            "için portföy satırı bulunamadı."
        )
    else:
        structural["structural_error"] = (
            snapshot.get("error")
            or "TEFAS yapısal veri alınamadı."
        )

    # --------------------------------------------------------
    # 2) FINTABLES FALLBACK
    #
    # TEFAS Direct başarısız olursa eski kaynak denenebilir.
    # Ancak başlıktan sayısal değer ÜRETİLMEZ.
    # --------------------------------------------------------
    fintables_url = (
        f"https://fintables.com/fonlar/{code.lower()}"
    )

    response, status = request_with_status(
        "Fintables",
        "GET",
        fintables_url,
        headers={
            "Accept": (
                "text/html,application/xhtml+xml;"
                "q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.7",
        },
    )

    if response is not None and status.ok:
        text = response.text
        text_lower = text.lower()

        challenge_markers = (
            "captcha",
            "cf-chl-",
            "challenge-platform",
            "verify you are human",
            "just a moment",
            "access denied",
        )

        if not any(
            marker in text_lower
            for marker in challenge_markers
        ):
            # Yalnızca açıkça etiketlenmiş gerçek sayısal alanları kabul et.
            match_top = re.search(
                r'En Büyük Pay["\s:]+'
                r'([0-9]+(?:[.,][0-9]+)?)',
                text,
                re.IGNORECASE,
            )
            if match_top:
                value = parse_number(match_top.group(1))
                if value is not None:
                    structural["top_asset_weight"] = value
                    structural["top_asset_weight_basis"] = (
                        "Fintables açık alanı"
                    )

            match_cash = re.search(
                r'(?:Nakit|Ters Repo|PPF)["\s:]+'
                r'([0-9]+(?:[.,][0-9]+)?)',
                text,
                re.IGNORECASE,
            )
            if match_cash:
                cash_val = parse_number(match_cash.group(1))
                if cash_val is not None:
                    structural["emergency_cash_ratio"] = cash_val
                    structural["cash_ratio_known"] = True

            if structural["top_asset_weight"] is not None or (
                structural["cash_ratio_known"]
            ):
                structural["structural_fetch_ok"] = True
                structural["structural_source"] = "Fintables"
                structural["structural_estimated"] = False
                structural["structural_error"] = (
                    "TEFAS Direct API kullanılamadı; "
                    "Fintables gerçek açık alanı kullanıldı."
                )
                return structural

    # Hiçbir kaynak gerçek sayısal yapısal veri vermediyse:
    # değerler None kalır; BAŞLIKTAN SAYISAL TAHMİN YAPILMAZ.
    structural["structural_fetch_ok"] = False
    structural["structural_source"] = "YOK"
    structural["structural_estimated"] = False

    if not structural["structural_error"]:
        structural["structural_error"] = (
            "TEFAS Direct ve Fintables yapısal veri sağlamadı."
        )

    return structural

# ============================================================
# FON SERİSİ
# ============================================================

def get_fund_series(
    universe: pd.DataFrame,
    fund_code: str,
    fund_kind: Optional[str] = None,
):

    code = normalize_fund_code(fund_code)

    source_statuses = []

    if not code:
        return None, "YOK", source_statuses

    # 1) Universe / pytefas
    if (
        universe is not None
        and not universe.empty
        and "code" in universe.columns
    ):

        rows = universe[
            universe["code"]
            .astype(str)
            .str.upper()
            .eq(code)
        ].copy()

        if not rows.empty:

            rows = (
                rows
                .sort_values("date")
                .drop_duplicates(
                    subset=["date"],
                    keep="last",
                )
            )

            if len(rows) >= 2:

                source_statuses.append(
                    SourceStatus(
                        source="TEFAS",
                        attempted=True,
                        ok=True,
                        message=(
                            f"Universe: "
                            f"{len(rows)} gözlem"
                        ),
                    )
                )

                return (
                    rows
                    .tail(
                        TARGET_TRADING_DAYS + 1
                    )
                    .reset_index(drop=True),
                    "TEFAS",
                    source_statuses,
                )

    # 2) TEFAS Direct API
    direct_df, direct_status = (
        fetch_tefas_direct_api(
            code,
            fund_kind,
        )
    )

    source_statuses.append(
        direct_status
    )

    if (
        direct_df is not None
        and len(direct_df) >= 2
    ):
        return (
            direct_df,
            "TEFAS Direct API",
            source_statuses,
        )

    # 3) İş Yatırım
    is_df, is_status = (
        fetch_isyatirim_series(code)
    )

    source_statuses.append(
        is_status
    )

    if (
        is_df is not None
        and len(is_df) >= 2
    ):
        return (
            is_df,
            "İş Yatırım",
            source_statuses,
        )

    return None, "YOK", source_statuses


# ============================================================
# FON METRİKLERİ
# ============================================================

def compute_fund_metrics(
    series: Optional[pd.DataFrame],
    fund_code: str,
    fund_kind: Optional[str] = None,
    fund_title: Optional[str] = None,
) -> Optional[dict]:

    if series is None or len(series) < 2:
        return None

    df = series.copy()

    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce",
    )

    df["price"] = df["price"].apply(
        parse_number
    )

    if "aum" not in df.columns:
        df["aum"] = 0.0

    if "investors" not in df.columns:
        df["investors"] = 0.0

    df["aum"] = (
        df["aum"]
        .apply(parse_number)
        .fillna(0.0)
    )

    df["investors"] = (
        df["investors"]
        .apply(parse_number)
        .fillna(0.0)
    )

    df = df.dropna(
        subset=["date", "price"]
    )

    df = df[df["price"] > 0]

    df = (
        df.sort_values("date")
        .drop_duplicates(
            subset=["date"],
            keep="last",
        )
        .reset_index(drop=True)
    )

    if len(df) < 2:
        return None

    prices = (
        df["price"]
        .astype(float)
        .tolist()
    )

    dates = (
        df["date"]
        .dt.strftime("%d.%m")
        .tolist()
    )

    aums = (
        df["aum"]
        .astype(float)
        .tolist()
    )

    investors = (
        df["investors"]
        .astype(float)
        .tolist()
    )

    daily_returns = [
        0.0 if previous <= 0
        else (current / previous - 1.0) * 100.0
        for previous, current in zip(
            prices[:-1],
            prices[1:],
        )
    ]

    if not daily_returns:
        return None

    max_dd = calculate_max_drawdown(
        prices
    )

    aum_change = (
        (
            aums[-1] / aums[0]
            - 1.0
        ) * 100.0
        if aums[0] > 0
        else None
    )

    investor_change = (
        (
            investors[-1] / investors[0]
            - 1.0
        ) * 100.0
        if investors[0] > 0
        else None
    )

    price_return_same_window = (
        (
            prices[-1] / prices[0]
            - 1.0
        ) * 100.0
        if prices[0] > 0
        else 0.0
    )

    # Bu GERÇEK NET PARA AKIŞI değildir.
    # Sadece AUM değişiminden fiyat etkisini çıkartan proxy'dir.
    if aum_change is not None:
        aum_flow_proxy = (
            aum_change
            - price_return_same_window
        )
    else:
        aum_flow_proxy = None

    recent_weekly_returns = (
        daily_returns[-5:]
        if len(daily_returns) >= 5
        else daily_returns
    )

    weekly_return = calculate_compounded_return(
        recent_weekly_returns
    )

    reference_date = (
        df["date"].iloc[-1].strftime("%Y-%m-%d")
        if not df.empty
        else None
    )

    structural = fetch_fund_structural_data(
        fund_code,
        fund_kind,
        fund_title,
        reference_date,
    )

    return {
        "dates": dates[1:],
        "prices": prices,
        "daily_returns": daily_returns,
        "n_days": len(daily_returns),

        "aum": aums[-1],
        "investors": (
            int(round(investors[-1]))
        ),

        "aum_change": aum_change,
        "aum_flow_proxy": aum_flow_proxy,
        "inv_change": investor_change,

        "max_dd": max_dd,
        "weekly_return": weekly_return,

        **structural,
    }


def fetch_and_compute_one_fund(
    code: str,
    universe: pd.DataFrame,
    meta_map: Dict[str, Dict[str, str]],
    valor_dict: Dict[str, float],
) -> Tuple[str, Optional[dict], str]:

    meta = meta_map.get(
        code,
        {},
    )

    fund_kind = meta.get("kind")
    fund_title_tefas = meta.get("title")

    series, source, statuses = (
        get_fund_series(
            universe,
            code,
            fund_kind,
        )
    )

    metrics = compute_fund_metrics(
        series,
        code,
        fund_kind,
        fund_title_tefas,
    )

    if metrics is None:
        return code, None, source

    metrics["code"] = code
    metrics["valor"] = (
        valor_dict.get(code, 0.0)
    )
    metrics["source"] = source
    metrics["kind"] = (
        fund_kind
        or DEFAULT_FUND_KIND
    )
    metrics["kind_known"] = (
        fund_kind is not None
    )

    metrics["source_statuses"] = [
        asdict(x)
        for x in statuses
    ]

    metrics["source_chain"] = " → ".join(
        x.source
        for x in statuses
        if x.attempted
    )

    # Fon adı: TEFAS öncelikli.
    title = (
        metrics.get("fund_title")
        or fund_title_tefas
        or ""
    ).strip()

    metrics["fund_title"] = (
        title if title else "-"
    )

    # Yatırım alanı:
    # Başlıktan sınıflandırma yapılabilir,
    # ancak numerik portföy tahmini yapılmaz.
    area = (
        metrics.get("investment_area")
        or ""
    ).strip()

    if not area and title:
        for pattern, label in [
            (
                r"Hisse\s*Senedi",
                "Hisse Senedi",
            ),
            (
                r"Teknoloji",
                "Hisse Senedi (Yabancı Teknoloji)",
            ),
            (
                r"De[gğ]i[sş]ken|Karma",
                "Karma / Değişken",
            ),
            (
                r"Alt[ıi]n",
                "Kıymetli Maden",
            ),
            (
                r"Kat[ıi]l[ıi]m",
                "Katılım",
            ),
            (
                r"Para\s*Piyasas",
                "Para Piyasası",
            ),
            (
                r"Serbest",
                "Serbest",
            ),
        ]:
            if re.search(
                pattern,
                title,
                re.IGNORECASE,
            ):
                area = label
                metrics[
                    "investment_area_source"
                ] = (
                    "TEFAS başlığı / sınıflandırma"
                )
                break

    metrics["investment_area"] = (
        area if area else "-"
    )

    return code, metrics, source


# ============================================================
# PENCERE METRİKLERİ
# ============================================================

def calculate_window_metrics(
    prices,
    returns,
    window,
) -> Optional[dict]:

    if (
        len(returns) < window
        or len(prices) < window + 1
    ):
        return None

    slice_returns = returns[-window:]
    slice_prices = prices[-(window + 1):]

    mean_return = (
        sum(slice_returns)
        / len(slice_returns)
    )

    variance = (
        sum(
            (r - mean_return) ** 2
            for r in slice_returns
        )
        / len(slice_returns)
    )

    volatility = variance ** 0.5

    sharpe_like = (
        mean_return / volatility
        if volatility > 1e-12
        else 0.0
    )

    cumulative_return = (
        slice_prices[-1]
        / slice_prices[0]
        - 1.0
    ) * 100.0

    max_dd = calculate_max_drawdown(
        slice_prices
    )

    return {
        "mean_return": mean_return,
        "volatility": volatility,
        "sharpe": sharpe_like,
        "cumulative": cumulative_return,
        "max_dd": max_dd,
    }


# ============================================================
# GÜVENLİK SKORU
# ============================================================

def calculate_security_scores(
    funds: List[dict],
) -> None:

    by_kind = defaultdict(list)

    for idx, fund in enumerate(funds):
        by_kind[
            fund.get(
                "kind",
                DEFAULT_FUND_KIND,
            )
        ].append(idx)

    kind_valor_median = {}

    for kind, indices in by_kind.items():

        valors = [
            safe_float(
                funds[i].get("valor")
            )
            for i in indices
        ]

        kind_valor_median[kind] = (
            statistics.median(valors)
            if valors
            else 0.0
        )

    for kind, indices in by_kind.items():

        subset = [
            funds[i]
            for i in indices
        ]

        aum_z = zscore([
            f.get("aum")
            for f in subset
        ])

        investor_z = zscore([
            f.get("investors")
            for f in subset
        ])

        aum_flow_z = zscore([
            f.get("aum_flow_proxy")
            for f in subset
        ])

        investor_change_z = zscore([
            f.get("inv_change")
            for f in subset
        ])

        concentration_z = zscore([
            f.get("top_asset_weight")
            for f in subset
        ])

        for local_i, fund_idx in enumerate(indices):

            fund = funds[fund_idx]

            score = 50.0

            score += (
                SECURITY_SCALE["aum"]
                * SECURITY_WEIGHTS["aum"]
                * aum_z[local_i]
            )

            score += (
                SECURITY_SCALE["investor"]
                * SECURITY_WEIGHTS["investor"]
                * investor_z[local_i]
            )

            # AUM-flow proxy etkisi bilinçli olarak azaltıldı.
            if fund.get("aum_flow_proxy") is not None:
                score += (
                    SECURITY_SCALE["aum_flow"]
                    * aum_flow_z[local_i]
                )

            if fund.get("inv_change") is not None:
                score += (
                    SECURITY_SCALE[
                        "investor_change"
                    ]
                    * investor_change_z[local_i]
                )

            top_asset = (
                fund.get("top_asset_weight")
            )

            if top_asset is not None:
                score -= (
                    SECURITY_SCALE[
                        "concentration"
                    ]
                    * SECURITY_WEIGHTS[
                        "concentration"
                    ]
                    * concentration_z[local_i]
                )

            if fund.get(
                "is_bist30",
                False,
            ):
                score += BIST30_BONUS

            cash_ratio = (
                fund.get(
                    "emergency_cash_ratio"
                )
            )

            if (
                fund.get(
                    "cash_ratio_known",
                    False,
                )
                and cash_ratio is not None
            ):
                if cash_ratio >= 15:
                    score += HIGH_LIQUIDITY_BONUS
                elif cash_ratio < 5:
                    score -= LOW_LIQUIDITY_PENALTY

            inv_change = fund.get(
                "inv_change"
            )

            if (
                inv_change is not None
                and safe_float(inv_change) > 0
            ):
                score += POSITIVE_INVESTOR_FLOW_BONUS

            if top_asset is not None:

                if top_asset > 30:
                    score -= min(
                        (
                            top_asset - 30
                        ) * 1.0,
                        MAX_CONCENTRATION_PENALTY,
                    )

                elif top_asset > 15:
                    score -= (
                        top_asset - 15
                    ) * 0.25

            median_valor = (
                kind_valor_median.get(
                    kind,
                    0.0,
                )
            )

            excess_valor = (
                safe_float(
                    fund.get("valor")
                )
                - median_valor
            )

            score -= calculate_valor_penalty(
                excess_valor
            )

            fund["security_score"] = int(
                round(
                    clamp(
                        score,
                        0.0,
                        100.0,
                    )
                )
            )


# ============================================================
# PİYASA-BAĞIL MOMENTUM
# ============================================================

def calculate_market_relative_momentum(
    funds: List[dict],
    reference,
    final_window: int,
) -> None:

    for fund in funds:

        kind = fund.get(
            "kind",
            DEFAULT_FUND_KIND,
        )

        sample_size = (
            reference_sample_size(
                reference,
                kind,
            )
        )

        window = min(
            final_window,
            fund["n_days"],
        )

        metrics = calculate_window_metrics(
            fund["prices"],
            fund["daily_returns"],
            window,
        )

        if metrics is None:

            fund["market_momentum"] = None
            fund["overheat_flag"] = False
            fund["reference_scope"] = (
                "Hesaplanamadı"
            )
            fund["volatility"] = None

            continue

        fund["_final_mean_return"] = (
            metrics["mean_return"]
        )
        fund["_final_sharpe"] = (
            metrics["sharpe"]
        )
        fund["_final_cumulative"] = (
            metrics["cumulative"]
        )
        fund["_final_max_dd"] = (
            metrics["max_dd"]
        )
        fund["volatility"] = (
            metrics["volatility"]
        )

        if sample_size >= MIN_REFERENCE_SAMPLE:

            ref = reference[kind]

            mean_m, mean_s = (
                population_mean_std(
                    ref["mean_return"]
                )
            )

            sharpe_m, sharpe_s = (
                population_mean_std(
                    ref["sharpe"]
                )
            )

            cum_m, cum_s = (
                population_mean_std(
                    ref["cumulative"]
                )
            )

            dd_m, dd_s = (
                population_mean_std(
                    ref["max_dd_inv"]
                )
            )

            z_mean = (
                zscore_against_population(
                    metrics["mean_return"],
                    mean_m,
                    mean_s,
                )
            )

            z_sharpe = (
                zscore_against_population(
                    metrics["sharpe"],
                    sharpe_m,
                    sharpe_s,
                )
            )

            z_cum = (
                zscore_against_population(
                    metrics["cumulative"],
                    cum_m,
                    cum_s,
                )
            )

            z_dd = (
                zscore_against_population(
                    -metrics["max_dd"],
                    dd_m,
                    dd_s,
                )
            )

            fund["reference_scope"] = (
                f"Piyasa ({kind}, n={sample_size})"
            )

        else:

            fallback_group = [
                f
                for f in funds
                if f.get("kind") == kind
            ]

            z_mean_all = zscore([
                f.get("_final_mean_return")
                for f in fallback_group
            ])

            z_sharpe_all = zscore([
                f.get("_final_sharpe")
                for f in fallback_group
            ])

            z_cum_all = zscore([
                f.get("_final_cumulative")
                for f in fallback_group
            ])

            z_dd_all = zscore([
                -safe_float(
                    f.get("_final_max_dd")
                )
                for f in fallback_group
            ])

            local_idx = fallback_group.index(
                fund
            )

            z_mean = z_mean_all[
                local_idx
            ]

            z_sharpe = z_sharpe_all[
                local_idx
            ]

            z_cum = z_cum_all[
                local_idx
            ]

            z_dd = z_dd_all[
                local_idx
            ]

            fund["reference_scope"] = (
                "Liste-bağıl "
                "(yetersiz evren verisi)"
            )

        weighted_z = (
            MOMENTUM_WEIGHTS["return"]
            * z_mean
            + MOMENTUM_WEIGHTS["sharpe"]
            * z_sharpe
            + MOMENTUM_WEIGHTS["cumulative"]
            * z_cum
            + MOMENTUM_WEIGHTS["drawdown"]
            * z_dd
        )

        momentum_score = clamp(
            50.0 + 20.0 * weighted_z,
            0.0,
            100.0,
        )

        daily_rets = (
            fund.get("daily_returns")
            or []
        )

        last_day_return = (
            daily_rets[-1]
            if daily_rets
            else 0.0
        )

        last_2_avg = (
            sum(daily_rets[-2:]) / 2.0
            if len(daily_rets) >= 2
            else last_day_return
        )

        overheat = (
            z_cum >= OVERHEAT_Z_THRESHOLD
            and (
                last_day_return < 0
                or last_2_avg < 0
            )
        )

        fund["overheat_flag"] = overheat

        if overheat:
            momentum_score = clamp(
                momentum_score
                - OVERHEAT_PENALTY,
                0.0,
                100.0,
            )

        fund["market_momentum"] = int(
            round(momentum_score)
        )


# ============================================================
# TREND SKORU
# ============================================================

def calculate_trend_scores(
    funds: List[dict],
) -> int:

    if not funds:
        return 0

    n_days = min(
        f["n_days"]
        for f in funds
    )

    for fund in funds:

        fund["dates"] = (
            fund["dates"][-n_days:]
        )

        fund["daily_returns"] = (
            fund["daily_returns"][-n_days:]
        )

        fund["prices"] = (
            fund["prices"][-(n_days + 1):]
        )

        fund[
            "running_trend_momentum"
        ] = []

    for d in range(
        1,
        n_days + 1,
    ):

        if d < MIN_ROLLING_DAYS:

            for fund in funds:
                fund[
                    "running_trend_momentum"
                ].append(None)

            continue

        current_metrics = []

        for fund in funds:

            returns_slice = (
                fund["daily_returns"][
                    d - MIN_ROLLING_DAYS:d
                ]
            )

            prices_slice = (
                fund["prices"][
                    d - MIN_ROLLING_DAYS:d + 1
                ]
            )

            if len(
                returns_slice
            ) < MIN_ROLLING_DAYS:
                continue

            mean_return = (
                sum(returns_slice)
                / len(returns_slice)
            )

            variance = (
                sum(
                    (r - mean_return) ** 2
                    for r in returns_slice
                )
                / len(returns_slice)
            )

            volatility = variance ** 0.5

            sharpe = (
                mean_return / volatility
                if volatility > 1e-12
                else 0.0
            )

            cumulative = (
                calculate_compounded_return(
                    returns_slice
                )
            )

            max_dd = calculate_max_drawdown(
                prices_slice
            )

            current_metrics.append({
                "fund": fund,
                "mean_return": mean_return,
                "sharpe": sharpe,
                "cumulative": cumulative,
                "max_dd": max_dd,
            })

        if not current_metrics:
            continue

        mean_z = zscore([
            x["mean_return"]
            for x in current_metrics
        ])

        sharpe_z = zscore([
            x["sharpe"]
            for x in current_metrics
        ])

        cumulative_z = zscore([
            x["cumulative"]
            for x in current_metrics
        ])

        drawdown_z = zscore([
            -x["max_dd"]
            for x in current_metrics
        ])

        for i, data in enumerate(
            current_metrics
        ):

            weighted_z = (
                MOMENTUM_WEIGHTS["return"]
                * mean_z[i]
                + MOMENTUM_WEIGHTS["sharpe"]
                * sharpe_z[i]
                + MOMENTUM_WEIGHTS["cumulative"]
                * cumulative_z[i]
                + MOMENTUM_WEIGHTS["drawdown"]
                * drawdown_z[i]
            )

            momentum_score = clamp(
                50.0 + 20.0 * weighted_z,
                0.0,
                100.0,
            )

            data["fund"][
                "running_trend_momentum"
            ].append(
                int(round(momentum_score))
            )

    for fund in funds:

        security_score = safe_float(
            fund.get(
                "security_score",
            ),
            50.0,
        )

        running_hybrid = []

        for momentum in fund[
            "running_trend_momentum"
        ]:

            if momentum is None:
                running_hybrid.append(None)
                continue

            hybrid = (
                momentum
                * HYBRID_MOMENTUM_WEIGHT
                + security_score
                * HYBRID_SECURITY_WEIGHT
            )

            running_hybrid.append(
                int(
                    round(
                        clamp(
                            hybrid,
                            0.0,
                            100.0,
                        )
                    )
                )
            )

        fund[
            "running_trend_hybrid"
        ] = running_hybrid

        valid_last = [
            s
            for s in running_hybrid
            if s is not None
        ][-5:]

        fund["last_5_scores"] = (
            valid_last
        )

        fund[
            "last_5_scores_str"
        ] = (
            " ➔ ".join(
                str(x)
                for x in valid_last
            )
            if valid_last
            else "-"
        )

        if valid_last:

            n = len(valid_last)

            weights = [
                EMA_DECAY ** (
                    n - 1 - i
                )
                for i in range(n)
            ]

            trend_score = (
                sum(
                    s * w
                    for s, w in zip(
                        valid_last,
                        weights,
                    )
                )
                / sum(weights)
            )

            fund["trend_skor"] = int(
                round(trend_score)
            )

        else:
            fund["trend_skor"] = None

    return n_days


# ============================================================
# KARAR
# ============================================================

def decision_label_from_score(score) -> str:
    """Karar skorunu mevcut model eşiklerine göre etikete çevirir."""
    if score is None:
        return "YETERSİZ VERİ"

    score = safe_float(score)

    if score >= STRONG_BUY:
        return "GÜÇLÜ AL"
    elif score >= WATCH_LIST:
        return "ASIL LİSTE"
    elif score >= CORRECTION:
        return "DÜZELTME / İZLE"
    else:
        return "ACİL SAT"


def calculate_rolling_decisions(
    funds: List[dict],
) -> None:
    """
    Günlük karar skorunu ve son 2/3 işlem gününün ortalama karar skorlarını
    hesaplar.

    Buradaki günlük skor, mevcut modelin günlük/rolling 5-günlük momentum
    hesabından üretilen running_trend_hybrid değeridir. Böylece ana
    'Karar Skoru' formülü değiştirilmez; sadece erken uyarı için kısa vadeli
    karar görünümü eklenir.
    """
    for fund in funds:
        running_scores = [
            safe_float(x)
            for x in (fund.get("running_trend_hybrid") or [])
            if x is not None
        ]

        if not running_scores:
            fund["daily_decision_score"] = None
            fund["daily_model_decision"] = "YETERSİZ VERİ"
            fund["decision_score_2d"] = None
            fund["model_decision_2d"] = "YETERSİZ VERİ"
            fund["decision_score_3d"] = None
            fund["model_decision_3d"] = "YETERSİZ VERİ"
            continue

        # En son geçerli işlem günü.
        daily_score = running_scores[-1]
        fund["daily_decision_score"] = int(round(daily_score))
        fund["daily_model_decision"] = decision_label_from_score(
            daily_score
        )

        # Son 2 işlem günü.
        if len(running_scores) >= 2:
            score_2d = sum(running_scores[-2:]) / 2.0
            fund["decision_score_2d"] = int(round(score_2d))
            fund["model_decision_2d"] = decision_label_from_score(
                score_2d
            )
        else:
            fund["decision_score_2d"] = None
            fund["model_decision_2d"] = "YETERSİZ VERİ"

        # Son 3 işlem günü.
        if len(running_scores) >= 3:
            score_3d = sum(running_scores[-3:]) / 3.0
            fund["decision_score_3d"] = int(round(score_3d))
            fund["model_decision_3d"] = decision_label_from_score(
                score_3d
            )
        else:
            fund["decision_score_3d"] = None
            fund["model_decision_3d"] = "YETERSİZ VERİ"


def finalize_decisions(
    funds: List[dict],
) -> None:

    for fund in funds:

        market_momentum = fund.get(
            "market_momentum"
        )

        security_score = fund.get(
            "security_score"
        )

        if (
            market_momentum is None
            or security_score is None
        ):
            fund["decision_score"] = None
            fund["karar"] = (
                "YETERSİZ VERİ"
            )
            continue

        decision_score = clamp(
            market_momentum
            * HYBRID_MOMENTUM_WEIGHT
            + security_score
            * HYBRID_SECURITY_WEIGHT,
            0.0,
            100.0,
        )

        fund["decision_score"] = int(
            round(decision_score)
        )

        score = fund[
            "decision_score"
        ]

        fund["karar"] = decision_label_from_score(score)


# ============================================================
# VERİ GÜVENİ 0-100
# ============================================================

def calculate_confidence_score(
    fund: dict,
) -> int:

    score = 0.0

    # 20 puan: yeterli fiyat geçmişi
    n_days = fund.get("n_days", 0)

    if n_days >= TARGET_TRADING_DAYS:
        score += 20
    elif n_days >= MIN_ROLLING_DAYS:
        score += 12
    elif n_days >= 2:
        score += 6

    # 25 puan: fiyat kaynağı
    source = fund.get("source")

    if source == "TEFAS":
        score += 25
    elif source == "TEFAS Direct API":
        score += 22
    elif source == "İş Yatırım":
        score += 17

    # 20 puan: yapısal veri
    if fund.get(
        "structural_fetch_ok",
        False,
    ):
        score += 20

    # 15 puan: piyasa referansı
    if str(
        fund.get(
            "reference_scope",
            ""
        )
    ).startswith("Piyasa"):
        score += 15
    elif "Liste-bağıl" in str(
        fund.get(
            "reference_scope",
            ""
        )
    ):
        score += 8

    # 10 puan: AUM/yatırımcı gibi ek verilerin
    # gerçekten dolu olması
    if (
        fund.get("aum") is not None
        and safe_float(
            fund.get("aum")
        ) > 0
    ):
        score += 5

    if (
        fund.get("investors") is not None
        and safe_float(
            fund.get("investors")
        ) > 0
    ):
        score += 5

    return int(
        round(
            clamp(
                score,
                0.0,
                100.0,
            )
        )
    )


def compute_confidence_label(
    fund: dict,
) -> str:

    score = calculate_confidence_score(
        fund
    )

    if score >= 80:
        return f"🟢 Yüksek ({score})"

    if score >= 60:
        return f"🟡 Orta ({score})"

    return f"🔴 Düşük ({score})"


# ============================================================
# EXCEL
# ============================================================

def style_excel_sheet(ws):

    thin_gray = Side(
        style="thin",
        color="D9E1F2",
    )

    for row in ws.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="center"
            )

            cell.border = Border(
                bottom=thin_gray
            )

    ws.freeze_panes = "A2"

    if ws.max_row >= 1:
        ws.auto_filter.ref = (
            ws.dimensions
        )

    ws.sheet_view.showGridLines = False


def auto_fit_columns(
    ws,
    min_width=10,
    max_width=45,
):

    for column_cells in ws.columns:

        column_index = (
            column_cells[0].column
        )

        max_length = 0

        for cell in column_cells:

            value = (
                ""
                if cell.value is None
                else str(cell.value)
            )

            max_length = max(
                max_length,
                len(value),
            )

        width = max(
            min_width,
            min(
                max_length + 3,
                max_width,
            ),
        )

        ws.column_dimensions[
            get_column_letter(
                column_index
            )
        ].width = width


PERCENT_COLUMNS = [
    "Ort. Günlük Getiri (%)",
    "Volatilite (%)",
    "Kümülatif Getiri (%)",
    "MaxDD (%)",
    "AUM Değişim - Ham (%)",
    "AUM Akış Proxy (%)",
    "Yatırımcı Değişim (%)",
    "Haftalık Bileşik Getiri (%)",
]


def create_excel_output(
    wb,
    ws_list,
    all_funds_for_output,
    common_n_days,
):

    if (
        "KGDM3_Puanlama"
        in wb.sheetnames
    ):
        del wb["KGDM3_Puanlama"]

    ws_scores = wb.create_sheet(
        title="KGDM3_Puanlama"
    )

    headers = [
        "Fon Kodu",
        "Fon Adı",
        "Yatırım Alanı",
        "Günlük Karar Skoru",
        "Günlük Model Kararı",
        "2 Günlük Karar Skoru",
        "2 Günlük Model Kararı",
        "3 Günlük Karar Skoru",
        "3 Günlük Model Kararı",
        "Valör",
        "Karar Skoru (Piyasa-Bağıl)",
        "Trend Skoru (Liste-Bağıl)",
        "Piyasa Momentum",
        "Güvenlik/Likidite Skoru",
        "Referans Kapsamı",
        "Veri Güveni",
        "Aşırı Isınma",
        "Son 5 Trend Skoru",
        "Model Kararı",
        "Ort. Günlük Getiri (%)",
        "Volatilite (%)",
        "Sharpe-benzeri",
        "Kümülatif Getiri (%)",
        "MaxDD (%)",
        "En Büyük Varlık (%)",
        "BIST30",
        "Nakit Verisi",
        "KAZRİSK",
        "AUM Değişim - Ham (%)",
        "AUM Akış Proxy (%)",
        "Yatırımcı Değişim (%)",
        "AUM (₺)",
        "Yatırımcı",
        "Haftalık Bileşik Getiri (%)",
        "Veri Kaynağı",
        "Kaynak Zinciri",
        "Yapısal Kaynak",
        "Yapısal Tahmin?",
        "Yapısal Hata",
    ]

    sample_dates = []

    for item in all_funds_for_output:

        if (
            item.get("dates")
            and len(item["dates"])
            >= common_n_days
        ):
            sample_dates = item[
                "dates"
            ][-common_n_days:]
            break

    if not sample_dates:

        for item in all_funds_for_output:

            if item.get("dates"):
                sample_dates = item[
                    "dates"
                ]
                break

    for day in sample_dates:
        headers.append(
            f"{day} Trend Hibrit Skor"
        )

    for day in sample_dates:
        headers.append(
            f"{day} Getiri"
        )

    ws_scores.append(headers)

    header_index = {
        name: idx + 1
        for idx, name in enumerate(
            headers
        )
    }

    header_fill = PatternFill(
        start_color=COLOR_NAVY,
        fill_type="solid",
    )

    header_font = Font(
        name="Calibri",
        bold=True,
        color=COLOR_WHITE,
    )

    for cell in ws_scores[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    ws_scores.row_dimensions[1].height = 55

    for item in all_funds_for_output:

        top_asset = item.get(
            "top_asset_weight"
        )

        if top_asset is None:
            risk_label = "⚪ Veri Yok"
        elif top_asset > 30:
            risk_label = (
                "⚠️ Yüksek Konsantrasyon"
            )
        elif top_asset > 15:
            risk_label = (
                "🟡 Orta Konsantrasyon"
            )
        else:
            risk_label = "🛡️ Dengeli"

        cash_known = item.get(
            "cash_ratio_known",
            False,
        )

        cash_label = (
            f"%{safe_float(item.get('emergency_cash_ratio')):.2f}"
            if cash_known
            else "Veri Yok"
        )

        row_data = [
            item["code"],
            item.get("fund_title") or "-",
            item.get(
                "investment_area"
            ) or "-",
            item.get("daily_decision_score"),
            item.get(
                "daily_model_decision",
                "YETERSİZ VERİ",
            ),
            item.get("decision_score_2d"),
            item.get(
                "model_decision_2d",
                "YETERSİZ VERİ",
            ),
            item.get("decision_score_3d"),
            item.get(
                "model_decision_3d",
                "YETERSİZ VERİ",
            ),
            item.get("valor", 0),
            item.get(
                "decision_score"
            ),
            item.get("trend_skor"),
            item.get("market_momentum"),
            item.get("security_score"),
            item.get(
                "reference_scope",
                "-",
            ),
            compute_confidence_label(item),
            (
                "🔥 Evet"
                if item.get(
                    "overheat_flag"
                )
                else "-"
            ),
            item.get(
                "last_5_scores_str",
                "-",
            ),
            item.get(
                "karar",
                "-",
            ),
            round(
                safe_float(
                    item.get(
                        "_final_mean_return"
                    )
                ),
                4,
            ),
            (
                round(
                    safe_float(
                        item.get(
                            "volatility"
                        )
                    ),
                    4,
                )
                if item.get(
                    "volatility"
                )
                is not None
                else ""
            ),
            round(
                safe_float(
                    item.get(
                        "_final_sharpe"
                    )
                ),
                4,
            ),
            round(
                safe_float(
                    item.get(
                        "_final_cumulative"
                    )
                ),
                4,
            ),
            round(
                safe_float(
                    item.get(
                        "_final_max_dd"
                    )
                ),
                4,
            ),
            (
                round(
                    safe_float(
                        top_asset
                    ),
                    2,
                )
                if top_asset is not None
                else None
            ),
            (
                "EVET"
                if item.get(
                    "is_bist30",
                    False,
                )
                else "HAYIR / YOK"
            ),
            cash_label,
            risk_label,
            (
                round(
                    safe_float(
                        item.get(
                            "aum_change"
                        )
                    ),
                    2,
                )
                if item.get(
                    "aum_change"
                )
                is not None
                else None
            ),
            (
                round(
                    safe_float(
                        item.get(
                            "aum_flow_proxy"
                        )
                    ),
                    2,
                )
                if item.get(
                    "aum_flow_proxy"
                )
                is not None
                else None
            ),
            (
                round(
                    safe_float(
                        item.get(
                            "inv_change"
                        )
                    ),
                    2,
                )
                if item.get(
                    "inv_change"
                )
                is not None
                else None
            ),
            (
                round(
                    safe_float(
                        item.get("aum")
                    ),
                    2,
                )
                if item.get(
                    "aum"
                )
                is not None
                else None
            ),
            (
                int(
                    item.get(
                        "investors"
                    )
                )
                if item.get(
                    "investors"
                )
                is not None
                else None
            ),
            round(
                safe_float(
                    item.get(
                        "weekly_return"
                    )
                ),
                4,
            ),
            item.get(
                "source",
                "-",
            ),
            item.get(
                "source_chain",
                "-",
            ),
            item.get(
                "structural_source",
                "YOK",
            ),
            (
                "EVET"
                if item.get(
                    "structural_estimated",
                    False,
                )
                else "HAYIR"
            ),
            item.get(
                "structural_error",
                "",
            ),
        ]

        n_dates = len(sample_dates)

        own_scores = (
            item.get(
                "running_trend_hybrid"
            )
            or []
        )

        if len(own_scores) < n_dates:
            own_scores = (
                [None] * (
                    n_dates
                    - len(own_scores)
                )
                + own_scores
            )
        else:
            own_scores = own_scores[
                -n_dates:
            ]

        row_data.extend([
            s
            if s is not None
            else ""
            for s in own_scores
        ])

        own_returns = (
            item.get(
                "daily_returns"
            )
            or []
        )

        if len(own_returns) < n_dates:
            own_returns = (
                [None] * (
                    n_dates
                    - len(own_returns)
                )
                + own_returns
            )
        else:
            own_returns = own_returns[
                -n_dates:
            ]

        row_data.extend([
            (
                format_percent(x)
                if x is not None
                else "-"
            )
            for x in own_returns
        ])

        ws_scores.append(
            row_data
        )

    green_font = Font(
        bold=True,
        color=COLOR_GREEN,
    )

    red_font = Font(
        bold=True,
        color=COLOR_RED,
    )

    yellow_font = Font(
        bold=True,
        color=COLOR_YELLOW,
    )

    decision_columns = [
        "Günlük Model Kararı",
        "2 Günlük Model Kararı",
        "3 Günlük Model Kararı",
        "Model Kararı",
    ]

    for row_number in range(
        2,
        ws_scores.max_row + 1,
    ):

        for decision_name in decision_columns:
            decision_col = header_index.get(
                decision_name
            )

            if not decision_col:
                continue

            decision_cell = (
                ws_scores.cell(
                    row=row_number,
                    column=decision_col,
                )
            )

            decision_text = str(
                decision_cell.value or ""
            )

            if (
                "GÜÇLÜ AL"
                in decision_text
                or "ASIL LİSTE"
                in decision_text
            ):
                decision_cell.font = (
                    green_font
                )

            elif (
                "DÜZELTME"
                in decision_text
            ):
                decision_cell.font = (
                    yellow_font
                )

            elif (
                "ACİL SAT"
                in decision_text
            ):
                decision_cell.font = (
                    red_font
                )

    score_columns = [
        "Günlük Karar Skoru",
        "2 Günlük Karar Skoru",
        "3 Günlük Karar Skoru",
        "Karar Skoru (Piyasa-Bağıl)",
    ]

    for score_name in score_columns:
        score_col = header_index.get(
            score_name
        )

        if not score_col:
            continue

        score_col_letter = get_column_letter(
            score_col
        )

        score_range = (
            f"{score_col_letter}2:"
            f"{score_col_letter}"
            f"{ws_scores.max_row}"
        )

        ws_scores.conditional_formatting.add(
            score_range,
            CellIsRule(
                operator="greaterThanOrEqual",
                formula=["75"],
                fill=PatternFill(
                    start_color=COLOR_LIGHT_GREEN,
                    fill_type="solid",
                ),
            ),
        )

        ws_scores.conditional_formatting.add(
            score_range,
            CellIsRule(
                operator="between",
                formula=["50", "74"],
                fill=PatternFill(
                    start_color=COLOR_LIGHT_YELLOW,
                    fill_type="solid",
                ),
            ),
        )

        ws_scores.conditional_formatting.add(
            score_range,
            CellIsRule(
                operator="lessThan",
                formula=["50"],
                fill=PatternFill(
                    start_color=COLOR_LIGHT_RED,
                    fill_type="solid",
                ),
            ),
        )

    currency_col = header_index.get(
        "AUM (₺)"
    )

    integer_col = header_index.get(
        "Yatırımcı"
    )

    for row_number in range(
        2,
        ws_scores.max_row + 1,
    ):

        if currency_col:
            ws_scores.cell(
                row=row_number,
                column=currency_col,
            ).number_format = (
                '#,##0.00 "₺"'
            )

        if integer_col:
            ws_scores.cell(
                row=row_number,
                column=integer_col,
            ).number_format = "#,##0"

        for col_name in (
            PERCENT_COLUMNS
        ):

            col_idx = header_index.get(
                col_name
            )

            if col_idx:
                ws_scores.cell(
                    row=row_number,
                    column=col_idx,
                ).number_format = (
                    '0.00"%"'
                )

    style_excel_sheet(
        ws_scores
    )

    auto_fit_columns(
        ws_scores
    )

    style_excel_sheet(
        ws_list
    )

    auto_fit_columns(
        ws_list
    )

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# ANA ARAYÜZ
# ============================================================

st.subheader("📂 Portföy Excel Listesi")

col_upload, col_github = st.columns(2)

wb = None

with col_upload:

    uploaded_file = st.file_uploader(
        "Bilgisayardan Excel Yükle",
        type=["xlsx"],
    )

    if uploaded_file is not None:

        try:
            wb = openpyxl.load_workbook(
                uploaded_file
            )

        except Exception as exc:
            st.error(
                f"Excel yükleme hatası: {exc}"
            )


with col_github:

    st.write(
        "Veya GitHub'daki listeyi kullanın:"
    )

    if st.button(
        "🚀 GitHub'dan Çek ve Analiz Et",
        use_container_width=True,
    ):

        resolved_url = (
            resolve_latest_github_excel_url()
        )

        target_url = (
            resolved_url
            or GITHUB_FALLBACK_URL
        )

        response, status = (
            request_with_status(
                "GitHub Excel",
                "GET",
                target_url,
            )
        )

        if (
            response is not None
            and status.ok
        ):

            try:
                wb = openpyxl.load_workbook(
                    io.BytesIO(
                        response.content
                    )
                )

                if resolved_url:
                    st.success(
                        "✅ En güncel Excel "
                        "dosyası indirildi."
                    )
                else:
                    st.warning(
                        "⚠️ Güncel dosya bulunamadı; "
                        "sabit yedek URL kullanıldı."
                    )

            except Exception as exc:
                st.error(
                    f"Excel ayrıştırma hatası: {exc}"
                )

        else:
            st.error(
                "GitHub dosyası alınamadı: "
                f"{status.error_type} "
                f"{status.message}"
            )


if wb is None:

    st.info(
        "Analize başlamak için Excel "
        "dosyanızı yükleyin."
    )

    st.stop()


# ============================================================
# FON LİSTESİ
# ============================================================

ws_list = (
    wb["Fon_Listesi"]
    if "Fon_Listesi"
    in wb.sheetnames
    else wb.active
)

requested_codes = []
excel_valor_dict = {}

for row in ws_list.iter_rows(
    min_row=2,
    values_only=False,
):

    if (
        not row
        or row[0].value is None
    ):
        continue

    code = normalize_fund_code(
        row[0].value
    )

    if not code:
        continue

    requested_codes.append(
        code
    )

    try:

        if len(row) > 3:

            valor = parse_number(
                row[3].value
            )

            excel_valor_dict[
                code
            ] = (
                valor
                if valor is not None
                else 0.0
            )

        else:
            excel_valor_dict[
                code
            ] = 0.0

    except Exception:
        excel_valor_dict[
            code
        ] = 0.0


requested_codes = list(
    dict.fromkeys(
        requested_codes
    )
)


if not requested_codes:

    st.error(
        "Fon_Listesi sayfasında "
        "fon kodu bulunamadı."
    )

    st.stop()


# ============================================================
# TARİH & TEFAS
# ============================================================

today = dt.date.today()

start_date = (
    today
    - dt.timedelta(
        days=LOOKBACK_CALENDAR_DAYS
    )
)

with st.spinner(
    "🔄 TEFAS verileri alınıyor..."
):

    universe = fetch_tefas_universe(
        start_date,
        today,
    )

    fund_meta_map = (
        build_fund_meta_map(
            universe
        )
    )

    universe_reference = (
        build_universe_reference(
            universe,
            window=TARGET_TRADING_DAYS,
        )
    )


if SHOW_DIAGNOSTICS:

    ref_counts = {
        k: reference_sample_size(
            universe_reference,
            k,
        )
        for k in FUND_KINDS
    }

    weak_kinds = [
        k
        for k, n in ref_counts.items()
        if 0 < n < MIN_REFERENCE_SAMPLE
    ]

    if all(
        n == 0
        for n in ref_counts.values()
    ):

        st.warning(
            "⚠️ TEFAS evren referansı "
            "oluşturulamadı. Piyasa-bağıl "
            "skor yerine liste-bağıl "
            "z-score kullanılabilir."
        )

    elif weak_kinds:

        st.info(
            "ℹ️ Zayıf evren referansı: "
            + ", ".join(weak_kinds)
        )


# ============================================================
# FONLARIN HESAPLANMASI
# ============================================================

calculated_funds = []
failed_codes = []

structural_fetch_failures = 0

progress = st.progress(
    0,
    text="Fonlar analiz ediliyor...",
)

total_funds = len(
    requested_codes
)

completed = 0

with concurrent.futures.ThreadPoolExecutor(
    max_workers=MAX_WORKERS
) as executor:

    future_to_code = {
        executor.submit(
            fetch_and_compute_one_fund,
            code,
            universe,
            fund_meta_map,
            excel_valor_dict,
        ): code
        for code in requested_codes
    }

    for future in (
        concurrent.futures.as_completed(
            future_to_code
        )
    ):

        code = future_to_code[
            future
        ]

        completed += 1

        try:

            _, metrics, source = (
                future.result()
            )

        except Exception as exc:

            metrics = None
            source = "YOK"

            if SHOW_DIAGNOSTICS:
                st.warning(
                    f"{code}: "
                    f"beklenmeyen hata: "
                    f"{exc}"
                )

        if metrics is None:

            failed_codes.append(
                code
            )

            progress.progress(
                completed / total_funds,
                text=f"{code}: veri yok",
            )

            continue

        if not metrics.get(
            "structural_fetch_ok",
            False,
        ):
            structural_fetch_failures += 1

        if ENABLE_FILTERS:

            investor_ok = (
                safe_float(
                    metrics.get(
                        "investors"
                    )
                )
                >= MIN_INVESTOR_COUNT
            )

            weekly_ok = (
                safe_float(
                    metrics.get(
                        "weekly_return"
                    )
                )
                >= TARGET_WEEKLY_RETURN
            )

            if not (
                investor_ok
                and weekly_ok
            ):

                progress.progress(
                    completed
                    / total_funds,
                    text=(
                        f"{code}: "
                        "filtre dışı"
                    ),
                )

                continue

        calculated_funds.append(
            metrics
        )

        progress.progress(
            completed / total_funds,
            text=(
                f"{code}: "
                f"{source}"
            ),
        )

progress.empty()


if failed_codes:

    st.warning(
        "Veri bulunamayan fonlar: "
        + ", ".join(
            sorted(
                failed_codes
            )
        )
    )


if (
    SHOW_DIAGNOSTICS
    and structural_fetch_failures > 0
):

    st.info(
        f"ℹ️ {structural_fetch_failures} "
        "fon için yapısal veri eksik veya "
        "kaynak tarafından sunulmadı. "
        "Bu alanlar artık başlıktan sayısal "
        "olarak tahmin edilmiyor."
    )


if not calculated_funds:

    st.error(
        "Hesaplanabilecek geçerli "
        "fon verisi bulunamadı."
    )

    st.stop()


# ============================================================
# YETERLİ / YETERSİZ
# ============================================================

eligible_funds = [
    f
    for f in calculated_funds
    if f.get(
        "n_days",
        0,
    ) >= MIN_ROLLING_DAYS
]

insufficient_funds = [
    f
    for f in calculated_funds
    if f.get(
        "n_days",
        0,
    ) < MIN_ROLLING_DAYS
]


for f in insufficient_funds:

    f["security_score"] = None
    f["market_momentum"] = None
    f["decision_score"] = None
    f["trend_skor"] = None
    f["karar"] = "YETERSİZ VERİ"

    f[
        "running_trend_hybrid"
    ] = []

    f[
        "last_5_scores_str"
    ] = "-"

    f["daily_decision_score"] = None
    f["daily_model_decision"] = "YETERSİZ VERİ"
    f["decision_score_2d"] = None
    f["model_decision_2d"] = "YETERSİZ VERİ"
    f["decision_score_3d"] = None
    f["model_decision_3d"] = "YETERSİZ VERİ"

    f[
        "reference_scope"
    ] = "-"

    f[
        "overheat_flag"
    ] = False

    f[
        "_final_mean_return"
    ] = 0.0

    f[
        "_final_sharpe"
    ] = 0.0

    f[
        "_final_cumulative"
    ] = 0.0

    f[
        "_final_max_dd"
    ] = 0.0

    f["volatility"] = 0.0

    if not f.get(
        "fund_title"
    ):
        f["fund_title"] = "-"

    if not f.get(
        "investment_area"
    ):
        f["investment_area"] = "-"


if insufficient_funds:

    st.info(
        f"ℹ️ {len(insufficient_funds)} "
        f"fon, en az {MIN_ROLLING_DAYS} "
        "işlem günü geçmişine sahip "
        "olmadığı için skorlanmadı."
    )


if not eligible_funds:

    st.error(
        "Fonlarda yeterli tarihsel veri "
        f"bulunmuyor (en az "
        f"{MIN_ROLLING_DAYS} gün gerekli)."
    )

    st.stop()


# ============================================================
# SKORLAMA
# ============================================================

with st.spinner(
    "📊 Hibrit skor hesaplanıyor..."
):

    calculate_security_scores(
        eligible_funds
    )

    calculate_market_relative_momentum(
        eligible_funds,
        universe_reference,
        final_window=TARGET_TRADING_DAYS,
    )

    common_n_days = (
        calculate_trend_scores(
            eligible_funds
        )
    )

    finalize_decisions(
        eligible_funds
    )

    # Kısa vadeli erken uyarı kararları:
    # Günlük + son 2 işlem günü + son 3 işlem günü.
    calculate_rolling_decisions(
        eligible_funds
    )


all_funds_for_output = (
    eligible_funds
    + insufficient_funds
)


all_funds_for_output.sort(
    key=lambda x: (
        -safe_float(
            x.get(
                "decision_score"
            )
        ),
        -safe_float(
            x.get(
                "_final_cumulative"
            )
        ),
    )
)


# ============================================================
# SONUÇ TABLOSU
# ============================================================

display_rows = []

for item in all_funds_for_output:

    top_asset = item.get(
        "top_asset_weight"
    )

    if top_asset is None:
        risk_status = "⚪ Veri Yok"
    elif top_asset > 30:
        risk_status = (
            "⚠️ Yüksek Konsantrasyon"
        )
    elif top_asset > 15:
        risk_status = (
            "🟡 Orta Konsantrasyon"
        )
    else:
        risk_status = "🛡️ Dengeli"

    display_rows.append({

        "Fon Kodu":
            item["code"],

        "Fon Adı":
            item.get(
                "fund_title"
            ) or "-",

        "Yatırım Alanı":
            item.get(
                "investment_area"
            ) or "-",

        "Günlük Karar Skoru":
            item.get(
                "daily_decision_score"
            ),

        "Günlük Model Kararı":
            item.get(
                "daily_model_decision",
                "YETERSİZ VERİ",
            ),

        "2 Günlük Karar Skoru":
            item.get(
                "decision_score_2d"
            ),

        "2 Günlük Model Kararı":
            item.get(
                "model_decision_2d",
                "YETERSİZ VERİ",
            ),

        "3 Günlük Karar Skoru":
            item.get(
                "decision_score_3d"
            ),

        "3 Günlük Model Kararı":
            item.get(
                "model_decision_3d",
                "YETERSİZ VERİ",
            ),

        "Valör":
            item.get(
                "valor",
                0,
            ),

        "Karar Skoru":
            item.get(
                "decision_score"
            ),

        "Piyasa Momentum":
            item.get(
                "market_momentum"
            ),

        "Güvenlik/Likidite":
            item.get(
                "security_score"
            ),

        "Trend Skoru":
            item.get(
                "trend_skor"
            ),

        "Model Kararı":
            item.get(
                "karar"
            ),

        "🔥 Isınma":
            (
                "Evet"
                if item.get(
                    "overheat_flag"
                )
                else "-"
            ),

        "Veri Güveni":
            compute_confidence_label(
                item
            ),

        "Referans":
            item.get(
                "reference_scope",
                "-",
            ),

        "Kümülatif Getiri %":
            round(
                safe_float(
                    item.get(
                        "_final_cumulative"
                    )
                ),
                3,
            ),

        "MaxDD %":
            round(
                safe_float(
                    item.get(
                        "_final_max_dd"
                    )
                ),
                3,
            ),

        "En Büyük Varlık %":
            (
                round(
                    safe_float(
                        top_asset
                    ),
                    2,
                )
                if top_asset is not None
                else None
            ),

        "KAZRİSK":
            risk_status,

        "Haftalık Bileşik %":
            round(
                safe_float(
                    item.get(
                        "weekly_return"
                    )
                ),
                3,
            ),

        "AUM ₺":
            (
                round(
                    safe_float(
                        item.get(
                            "aum"
                        )
                    ),
                    0,
                )
                if item.get(
                    "aum"
                ) is not None
                else None
            ),

        "Yatırımcı":
            item.get(
                "investors"
            ),

        "Kaynak":
            item.get(
                "source"
            ),

        "Kaynak Zinciri":
            item.get(
                "source_chain",
                "-",
            ),

        "Yapısal Kaynak":
            item.get(
                "structural_source",
                "YOK",
            ),

        "Yapısal Hata":
            item.get(
                "structural_error",
                "",
            ),
    })


df_display = pd.DataFrame(
    display_rows
)


def color_cells(value):

    text = str(value)

    if (
        "GÜÇLÜ AL" in text
        or "ASIL LİSTE" in text
        or "Dengeli" in text
        or "🟢" in text
    ):
        return (
            "color: #008000; "
            "font-weight: bold;"
        )

    if (
        "DÜZELTME" in text
        or "Orta" in text
        or "🟡" in text
    ):
        return (
            "color: #B8860B; "
            "font-weight: bold;"
        )

    if (
        "ACİL SAT" in text
        or "Yüksek Konsantrasyon"
        in text
        or "YETERSİZ" in text
        or "Düşük" in text
        or "🔴" in text
    ):
        return (
            "color: #FF0000; "
            "font-weight: bold;"
        )

    return ""


try:
    styled_df = (
        df_display.style.map(
            color_cells
        )
    )
except AttributeError:
    styled_df = (
        df_display.style.applymap(
            color_cells
        )
    )


st.subheader(
    "📊 Analiz Sonuçları"
)

st.caption(
    "**Karar Skoru** piyasa-bağıl ana skordur. "
    "**Günlük / 2 Günlük / 3 Günlük Karar Skoru** "
    "mevcut günlük rolling hibrit skorların sırasıyla "
    "son 1, 2 ve 3 işlem günündeki değerlerinden hesaplanır. "
    "Bu alanlar özellikle kısa vadeli **ACİL SAT** sinyalini "
    "erken fark etmek için eklenmiştir. "
    "AUM Akış Proxy gerçek net para girişi/çıkışı değildir."
)

st.dataframe(
    styled_df,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# KISA VADELİ ACİL SAT UYARISI
# ============================================================

early_sell_alerts = []

for item in all_funds_for_output:
    alerts = []

    if item.get("daily_model_decision") == "ACİL SAT":
        alerts.append("Günlük")

    if item.get("model_decision_2d") == "ACİL SAT":
        alerts.append("2 Günlük")

    if item.get("model_decision_3d") == "ACİL SAT":
        alerts.append("3 Günlük")

    if alerts:
        early_sell_alerts.append({
            "Fon Kodu": item.get("code"),
            "Fon Adı": item.get("fund_title") or "-",
            "Yatırım Alanı": item.get("investment_area") or "-",
            "Uyarı": " + ".join(alerts),
            "Günlük Skor": item.get("daily_decision_score"),
            "2 Günlük Skor": item.get("decision_score_2d"),
            "3 Günlük Skor": item.get("decision_score_3d"),
            "Ana Karar Skoru": item.get("decision_score"),
            "Ana Model Kararı": item.get("karar"),
        })

if early_sell_alerts:
    st.error(
        f"🚨 {len(early_sell_alerts)} fon için kısa vadeli "
        '"ACİL SAT" uyarısı oluştu.'
    )
    st.dataframe(
        pd.DataFrame(early_sell_alerts),
        use_container_width=True,
        hide_index=True,
    )
else:
    st.success(
        '✅ Günlük, 2 günlük ve 3 günlük kısa vadeli '
        '"ACİL SAT" uyarısı oluşmadı.'
    )


# ============================================================
# KAYNAK TANILAMA
# ============================================================

if SHOW_DIAGNOSTICS:

    st.subheader(
        "🔎 Veri Kaynağı Tanılaması"
    )

    diagnostic_rows = []

    for item in all_funds_for_output:

        statuses = (
            item.get(
                "source_statuses",
                [],
            )
        )

        for status in statuses:

            diagnostic_rows.append({
                "Fon":
                    item["code"],
                "Kaynak":
                    status.get(
                        "source"
                    ),
                "Denendi":
                    "Evet"
                    if status.get(
                        "attempted"
                    )
                    else "Hayır",
                "Başarılı":
                    "Evet"
                    if status.get(
                        "ok"
                    )
                    else "Hayır",
                "HTTP":
                    status.get(
                        "status_code"
                    ),
                "Hata":
                    status.get(
                        "error_type"
                    ),
                "Mesaj":
                    status.get(
                        "message"
                    ),
                "Süre ms":
                    status.get(
                        "elapsed_ms"
                    ),
            })

    if diagnostic_rows:

        st.dataframe(
            pd.DataFrame(
                diagnostic_rows
            ),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# SKOR ÖZETİ
# ============================================================

st.subheader(
    "📈 Skor Özeti"
)

col1, col2, col3, col4 = (
    st.columns(4)
)

scores = [
    safe_float(
        x.get(
            "decision_score"
        )
    )
    for x in all_funds_for_output
    if x.get(
        "decision_score"
    )
    is not None
]

if scores:

    with col1:
        st.metric(
            "En Yüksek Skor",
            f"{max(scores):.0f}",
        )

    with col2:
        st.metric(
            "Ortalama Skor",
            f"{sum(scores) / len(scores):.1f}",
        )

    with col3:
        st.metric(
            "En Düşük Skor",
            f"{min(scores):.0f}",
        )

    with col4:

        strong_count = sum(
            1
            for x
            in all_funds_for_output
            if x.get(
                "karar"
            ) == "GÜÇLÜ AL"
        )

        st.metric(
            "Güçlü Al",
            strong_count,
        )


overheat_count = sum(
    1
    for x
    in all_funds_for_output
    if x.get(
        "overheat_flag"
    )
)

if overheat_count:

    st.info(
        f"🔥 {overheat_count} fon "
        "'aşırı ısınma' uyarısı taşıyor."
    )


# ============================================================
# VERİ KALİTESİ ÖZETİ
# ============================================================

confidence_scores = [
    calculate_confidence_score(
        x
    )
    for x in all_funds_for_output
]

if confidence_scores:

    st.metric(
        "Ortalama Veri Güveni",
        f"{sum(confidence_scores) / len(confidence_scores):.1f}/100",
    )


# ============================================================
# EXCEL ÇIKTISI
# ============================================================

output = create_excel_output(
    wb=wb,
    ws_list=ws_list,
    all_funds_for_output=all_funds_for_output,
    common_n_days=common_n_days,
)


st.success(
    f"✅ Analiz tamamlandı. "
    f"{len(all_funds_for_output)} fon işlendi "
    f"({len(eligible_funds)} skorlandı, "
    f"{len(insufficient_funds)} yetersiz veri). "
    f"Model sürümü: {APP_VERSION}"
)


st.download_button(
    label="📥 Güncellenmiş Hibrit Excel'i İndir",
    data=output,
    file_name=(
        "fonlar_KGDM3_KAZRISK_V8_1.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-"
        "officedocument.spreadsheetml.sheet"
    ),
)
