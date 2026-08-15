import concurrent.futures
import datetime as dt
import io
import re
import statistics
from collections import defaultdict
from typing import Optional, List, Dict, Any, Tuple

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# SAYFA AYARLARI
# ============================================================

st.set_page_config(
    page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK® Hibrit Fon Analiz ve Excel Otomasyonu")
st.caption(
    "TEFAS + İş Yatırım + TEFAS Direct API + Fintables/KAP | "
    "Momentum + Risk + Likidite Hibrit Skor Motoru V7.2 (Piyasa-Bağıl)"
)


# ============================================================
# GENEL AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")
DEFAULT_FUND_KIND = "YAT"

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 5
HTTP_TIMEOUT = 12
MAX_WORKERS = 8

MIN_REFERENCE_SAMPLE = 5
OVERHEAT_Z_THRESHOLD = 2.0
OVERHEAT_PENALTY = 6.0

APP_VERSION = "7.2.0"

GITHUB_OWNER = "tlgssk"
GITHUB_REPO = "kgdm3-fon-analiz"
GITHUB_BRANCH = "main"
GITHUB_FALLBACK_URL = (
    "https://github.com/tlgssk/kgdm3-fon-analiz/"
    "raw/refs/heads/main/"
    "Menkul_Kiymet_Yatirim_Fonlari_EXCEL_Tum_Veri_2026-08-14.xlsx"
)

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
# SKOR PARAMETRELERİ
# ============================================================

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
    "aum_flow": 12.0,
    "investor_change": 8.0,
    "concentration": 20.0,
}

DEFAULT_HYBRID_MOMENTUM_WEIGHT = 0.60
DEFAULT_HYBRID_SECURITY_WEIGHT = 0.40

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
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

ENABLE_FILTERS = st.sidebar.checkbox(
    "Filtreleri Etkinleştir",
    value=False,
    help=(
        "Kapalı olduğunda listedeki tüm fonlar hesaplanır. "
        "Açık olduğunda yatırımcı ve haftalık getiri kriterleri uygulanır."
    ),
)

TARGET_WEEKLY_RETURN = st.sidebar.slider(
    "Hedef Haftalık Getiri (%)", min_value=-5.00, max_value=10.00, value=0.00, step=0.10,
)

MIN_INVESTOR_COUNT = st.sidebar.slider(
    "Minimum Yatırımcı Sayısı", min_value=0, max_value=100000, value=0, step=500,
)

with st.sidebar.expander("⚖️ Skor Ağırlıkları (deneysel / kalibre edilmemiş)"):
    st.caption(
        "Bu katsayılar geçmiş performansla test edilmemiş sezgisel "
        "varsayılanlardır. Değiştirerek duyarlılık analizi yapabilirsiniz."
    )
    w_return = st.slider("Getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe-benzeri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_cumulative = st.slider("Kümülatif getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_drawdown = st.slider("Drawdown ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05)
    _w_total = w_return + w_sharpe + w_cumulative + w_drawdown
    if _w_total <= 0:
        _w_total = 1.0
    MOMENTUM_WEIGHTS = {
        "return": w_return / _w_total,
        "sharpe": w_sharpe / _w_total,
        "cumulative": w_cumulative / _w_total,
        "drawdown": w_drawdown / _w_total,
    }

    hybrid_momentum_w = st.slider(
        "Hibrit skorda Momentum ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_MOMENTUM_WEIGHT, 0.05
    )
    HYBRID_MOMENTUM_WEIGHT = hybrid_momentum_w
    HYBRID_SECURITY_WEIGHT = 1.0 - hybrid_momentum_w

with st.sidebar.expander("🔧 Tanılama"):
    SHOW_DIAGNOSTICS = st.checkbox("Veri kaynağı tanılama bilgisi göster", value=True)


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
        text.replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
    )
    if not text:
        return None

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
    clean = [safe_float(v, None) for v in returns]
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
    peak = prices[0]
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
    if not values:
        return []
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
        z = clamp((value - mean_value) / std, -Z_LIMIT, Z_LIMIT)
        result.append(z)
    return result


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


# ============================================================
# GITHUB
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def resolve_latest_github_excel_url() -> Optional[str]:
    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/?ref={GITHUB_BRANCH}"
    )
    headers = {"Accept": "application/vnd.github+json", "User-Agent": "kgdm3-fon-analiz-app"}
    try:
        response = requests.get(api_url, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200:
            return None
        items = response.json()
        if not isinstance(items, list):
            return None
        xlsx_files = [
            item for item in items
            if isinstance(item, dict)
            and str(item.get("name", "")).lower().endswith(".xlsx")
            and item.get("download_url")
        ]
        if not xlsx_files:
            return None
        date_pattern = re.compile(r"(\d{4}-\d{2}-\d{2})")

        def sort_key(item):
            match = date_pattern.search(item.get("name", ""))
            return (match.group(1) if match else "", item.get("name", ""))

        xlsx_files.sort(key=sort_key, reverse=True)
        return xlsx_files[0]["download_url"]
    except Exception:
        return None


# ============================================================
# TEFAS UNIVERSE
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
    except ImportError:
        return pd.DataFrame()

    try:
        crawler = Crawler(timeout=60, max_retry=3)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")

        if df is None or df.empty:
            return pd.DataFrame()

        df = df.copy()
        rename_map = {
            "fund_code": "code", "fund_name": "title",
            "investor_count": "investors", "portfolio_size": "aum",
            "fund_type": "kind", "kind": "kind",
        }
        df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

        if "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
        if "price" in df.columns:
            df["price"] = df["price"].apply(parse_number)

        if "aum" in df.columns:
            df["aum"] = df["aum"].apply(parse_number).fillna(0.0)
        else:
            df["aum"] = 0.0

        if "investors" in df.columns:
            df["investors"] = df["investors"].apply(parse_number).fillna(0.0)
        else:
            df["investors"] = 0.0

        if "kind" not in df.columns:
            df["kind"] = DEFAULT_FUND_KIND
        if "title" not in df.columns:
            df["title"] = ""

        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df = df.dropna(subset=["date", "code", "price"])
        df = df[df["price"] > 0]

        return (
            df.sort_values(["code", "date"])
            .drop_duplicates(subset=["code", "date"], keep="last")
            .reset_index(drop=True)
        )
    except Exception:
        return pd.DataFrame()


def build_fund_meta_map(universe: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """code -> {kind, title}"""
    meta: Dict[str, Dict[str, str]] = {}
    if universe is None or universe.empty:
        return meta
    try:
        latest = universe.sort_values("date").drop_duplicates(subset=["code"], keep="last")
        for _, row in latest.iterrows():
            code = str(row.get("code", "")).strip().upper()
            if not code:
                continue
            kind = str(row.get("kind", "")).strip().upper()
            title = str(row.get("title", "") or "").strip()
            if title.lower() in ("", "nan", "none"):
                title = ""
            meta[code] = {
                "kind": kind if kind in FUND_KINDS else DEFAULT_FUND_KIND,
                "title": title,
            }
    except Exception:
        pass
    return meta


def build_universe_reference(universe: pd.DataFrame, window: int) -> Dict[str, Dict[str, List[float]]]:
    reference: Dict[str, Dict[str, List[float]]] = {
        k: {"mean_return": [], "sharpe": [], "cumulative": [], "max_dd_inv": []} for k in FUND_KINDS
    }
    if universe is None or universe.empty or "kind" not in universe.columns or window < 2:
        return reference

    for code, group in universe.groupby("code"):
        group = group.sort_values("date")
        kind = str(group["kind"].iloc[-1]).strip().upper()
        if kind not in FUND_KINDS:
            continue
        prices = group["price"].astype(float).tolist()
        if len(prices) < window + 1:
            continue

        window_prices = prices[-(window + 1):]
        returns = []
        for p0, p1 in zip(window_prices[:-1], window_prices[1:]):
            returns.append(0.0 if p0 <= 0 else (p1 / p0 - 1.0) * 100.0)

        mean_return = sum(returns) / len(returns)
        variance = sum((r - mean_return) ** 2 for r in returns) / len(returns)
        volatility = variance ** 0.5
        sharpe = (mean_return / volatility) if volatility > 1e-12 else 0.0
        cumulative = (window_prices[-1] / window_prices[0] - 1.0) * 100.0
        max_dd = calculate_max_drawdown(window_prices)

        reference[kind]["mean_return"].append(mean_return)
        reference[kind]["sharpe"].append(sharpe)
        reference[kind]["cumulative"].append(cumulative)
        reference[kind]["max_dd_inv"].append(-max_dd)

    return reference


def reference_sample_size(reference: Dict[str, Dict[str, List[float]]], kind: str) -> int:
    if kind not in reference:
        return 0
    return len(reference[kind].get("mean_return", []))


# ============================================================
# İŞ YATIRIM
# ============================================================

def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code:
        return None

    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = (
        "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/"
        "Data.aspx/YatirimFonGecmisGetiri"
    )
    params = {
        "fonKod": code,
        "baslangic": start.strftime("%d-%m-%Y"),
        "bitis": end.strftime("%d-%m-%Y"),
    }
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200:
            return None
        payload = response.json()
        values = payload.get("value")
        if not values:
            return None

        df = pd.DataFrame(values)
        if "Tarih" not in df.columns or "Fiyat" not in df.columns:
            return None

        df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
        df["price"] = df["Fiyat"].apply(parse_number)
        df["aum"] = 0.0
        df["investors"] = 0.0

        df = df.dropna(subset=["date", "price"])
        df = df[df["price"] > 0]
        if len(df) < 2:
            return None

        return (
            df.sort_values("date")
            .drop_duplicates(subset=["date"], keep="last")
            .tail(TARGET_TRADING_DAYS + 1)
            .reset_index(drop=True)[["date", "price", "aum", "investors"]]
        )
    except Exception:
        return None


# ============================================================
# TEFAS DIRECT API
# ============================================================

def fetch_tefas_direct_api(fund_code: str, fund_kind: Optional[str] = None) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code:
        return None

    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.tefas.gov.tr",
    }

    kind_candidates = [fund_kind] if fund_kind in FUND_KINDS else []
    kind_candidates += [k for k in FUND_KINDS if k not in kind_candidates]

    for kind in kind_candidates:
        payload = {
            "fontip": kind, "fonkod": code,
            "bastarih": start.strftime("%d.%m.%Y"), "bittarih": end.strftime("%d.%m.%Y"),
        }
        try:
            response = requests.post(url, data=payload, headers=headers, timeout=HTTP_TIMEOUT)
            if response.status_code != 200:
                continue
            data = response.json().get("data", [])
            if not data:
                continue

            df = pd.DataFrame(data)
            if not all(c in df.columns for c in ["TARIH", "FIYAT"]):
                continue

            df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
            df["price"] = df["FIYAT"].apply(parse_number)
            df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number).fillna(0.0) if "PORTFOYBUYUKLUK" in df.columns else 0.0
            df["investors"] = df["KISISAYISI"].apply(parse_number).fillna(0.0) if "KISISAYISI" in df.columns else 0.0

            df = df.dropna(subset=["date", "price"])
            df = df[df["price"] > 0]
            if len(df) < 2:
                continue

            return (
                df.sort_values("date")
                .drop_duplicates(subset=["date"], keep="last")
                .tail(TARGET_TRADING_DAYS + 1)
                .reset_index(drop=True)
            )
        except Exception:
            continue

    return None


# ============================================================
# FİNTABLES / YAPISAL VERİ + FON ADI + YATIRIM ALANI
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 60 * 6)
def fetch_fund_structural_data(fund_code: str) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {
        "top_asset_weight": None,
        "is_bist30": False,
        "is_bist30_known": False,
        "emergency_cash_ratio": None,
        "cash_ratio_known": False,
        "structural_fetch_ok": False,
        "fund_title": None,
        "investment_area": None,
    }
    if not code:
        return structural

    try:
        fintables_url = f"https://fintables.com/fonlar/{code.lower()}"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(fintables_url, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200:
            return structural

        text = response.text

        # En büyük varlık
        match_top = re.search(r'En Büyük Pay["\s:]+([0-9]+(?:[.,][0-9]+)?)', text, re.IGNORECASE)
        if match_top:
            structural["top_asset_weight"] = parse_number(match_top.group(1))

        # BIST30
        if re.search(r"BIST\s*30", text, re.IGNORECASE):
            structural["is_bist30"] = True
            structural["is_bist30_known"] = True

        # Nakit
        match_cash = re.search(
            r'(?:Nakit|Ters Repo|PPF)["\s:]+([0-9]+(?:[.,][0-9]+)?)', text, re.IGNORECASE
        )
        if match_cash:
            cash_value = parse_number(match_cash.group(1))
            if cash_value is not None:
                structural["emergency_cash_ratio"] = cash_value
                structural["cash_ratio_known"] = True

        # Fon adı
        match_title = re.search(
            r'<h1[^>]*>([^<]{5,150})</h1>|<title>([^<]{5,150}?)\s*[-|–|·|]',
            text,
            re.IGNORECASE | re.DOTALL,
        )
        if match_title:
            raw = (match_title.group(1) or match_title.group(2) or "").strip()
            raw = re.sub(r"\s+", " ", raw)
            if raw:
                structural["fund_title"] = raw[:120]

        # Yatırım alanı
        area_patterns = [
            (r"Hisse\s*Senedi\s*Yo[gğ]un|Hisse\s*Senedi\s*Fon|Hisse\s*Senedi", "Hisse Senedi"),
            (r"Yabanc[ıi]\s*Teknoloji|Teknoloji\s*Sekt[oö]r|Yeni\s*Teknoloj", "Hisse Senedi (Yabancı Teknoloji)"),
            (r"Yabanc[ıi]\s*Hisse", "Hisse Senedi (Yabancı)"),
            (r"Bor[cç]lanma\s*Ara[cç]|Tahvil|Bono", "Borçlanma Araçları"),
            (r"Para\s*Piyasas", "Para Piyasası"),
            (r"De[gğ]i[sş]ken\s*Fon|Karma\s*Fon|De[gğ]i[sş]ken", "Karma / Değişken"),
            (r"Alt[ıi]n\s*Kat[ıi]l[ıi]m|Kat[ıi]l[ıi]m.*Alt[ıi]n", "Kıymetli Maden (Altın Katılım)"),
            (r"K[ıi]ymetli\s*Maden|Alt[ıi]n\s*Fon|G[uü]m[uü][sş]", "Kıymetli Maden"),
            (r"Kat[ıi]l[ıi]m\s*Fon|Faizsiz", "Katılım"),
            (r"Fon\s*Sepeti", "Fon Sepeti"),
            (r"Serbest\s*Fon|Serbest", "Serbest"),
            (r"Koruma\s*Ama[cç]l[ıi]|Anapara\s*Koruma", "Koruma Amaçlı"),
            (r"BYF|ETF|Borsa\s*Yat[ıi]r[ıi]m", "BYF / ETF"),
            (r"Emeklilik", "Emeklilik"),
        ]
        for pat, label in area_patterns:
            if re.search(pat, text, re.IGNORECASE):
                structural["investment_area"] = label
                break

        # İsimden de tahmin
        if not structural["investment_area"] and structural.get("fund_title"):
            for pat, label in area_patterns:
                if re.search(pat, structural["fund_title"], re.IGNORECASE):
                    structural["investment_area"] = label
                    break

        structural["structural_fetch_ok"] = any([
            structural["top_asset_weight"] is not None,
            structural["is_bist30_known"],
            structural["cash_ratio_known"],
            structural["fund_title"] is not None,
            structural["investment_area"] is not None,
        ])
    except Exception:
        pass

    return structural


# ============================================================
# FON SERİSİ
# ============================================================

def get_fund_series(universe: pd.DataFrame, fund_code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(fund_code)
    if not code:
        return None, "YOK"

    if universe is not None and not universe.empty and "code" in universe.columns:
        rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
        if not rows.empty:
            rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
            if len(rows) >= 2:
                return rows.tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True), "TEFAS"

    direct_df = fetch_tefas_direct_api(code, fund_kind)
    if direct_df is not None and len(direct_df) >= 2:
        return direct_df, "TEFAS Direct API"

    is_df = fetch_isyatirim_series(code)
    if is_df is not None:
        return is_df, "İş Yatırım"

    return None, "YOK"


# ============================================================
# FON METRİKLERİ
# ============================================================

def compute_fund_metrics(series: Optional[pd.DataFrame], fund_code: str) -> Optional[dict]:
    if series is None or len(series) < 2:
        return None

    df = series.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["price"] = df["price"].apply(parse_number)
    if "aum" not in df.columns:
        df["aum"] = 0.0
    if "investors" not in df.columns:
        df["investors"] = 0.0
    df["aum"] = df["aum"].apply(parse_number).fillna(0.0)
    df["investors"] = df["investors"].apply(parse_number).fillna(0.0)

    df = df.dropna(subset=["date", "price"])
    df = df[df["price"] > 0]
    df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)

    if len(df) < 2:
        return None

    prices = df["price"].astype(float).tolist()
    dates = df["date"].dt.strftime("%d.%m").tolist()
    aums = df["aum"].astype(float).tolist()
    investors = df["investors"].astype(float).tolist()

    daily_returns = []
    for previous, current in zip(prices[:-1], prices[1:]):
        daily_returns.append(0.0 if previous <= 0 else (current / previous - 1.0) * 100.0)

    if not daily_returns:
        return None

    max_dd = calculate_max_drawdown(prices)
    aum_change = ((aums[-1] / aums[0] - 1.0) * 100.0) if aums[0] > 0 else 0.0
    investor_change = ((investors[-1] / investors[0] - 1.0) * 100.0) if investors[0] > 0 else 0.0

    price_return_same_window = ((prices[-1] / prices[0] - 1.0) * 100.0) if prices[0] > 0 else 0.0
    aum_flow_proxy = aum_change - price_return_same_window

    recent_weekly_returns = daily_returns[-5:] if len(daily_returns) >= 5 else daily_returns
    weekly_return = calculate_compounded_return(recent_weekly_returns)

    structural = fetch_fund_structural_data(fund_code)

    return {
        "dates": dates[1:],
        "prices": prices,
        "daily_returns": daily_returns,
        "n_days": len(daily_returns),
        "aum": aums[-1],
        "investors": int(round(investors[-1])),
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
    meta = meta_map.get(code, {})
    fund_kind = meta.get("kind")
    series, source = get_fund_series(universe, code, fund_kind)
    metrics = compute_fund_metrics(series, code)
    if metrics is None:
        return code, None, source

    metrics["code"] = code
    metrics["valor"] = valor_dict.get(code, 0.0)
    metrics["source"] = source
    metrics["kind"] = fund_kind or DEFAULT_FUND_KIND
    metrics["kind_known"] = fund_kind is not None

    # Fon adı: TEFAS öncelikli, yoksa Fintables
    title = (meta.get("title") or metrics.get("fund_title") or "").strip()
    metrics["fund_title"] = title if title else "-"

    # Yatırım alanı
    area = (metrics.get("investment_area") or "").strip()
    if not area and title:
        # İsimden son bir tahmin
        for pat, label in [
            (r"Hisse\s*Senedi", "Hisse Senedi"),
            (r"Teknoloji", "Hisse Senedi (Yabancı Teknoloji)"),
            (r"De[gğ]i[sş]ken|Karma", "Karma / Değişken"),
            (r"Alt[ıi]n", "Kıymetli Maden"),
            (r"Kat[ıi]l[ıi]m", "Katılım"),
            (r"Para\s*Piyasas", "Para Piyasası"),
            (r"Serbest", "Serbest"),
        ]:
            if re.search(pat, title, re.IGNORECASE):
                area = label
                break
    metrics["investment_area"] = area if area else "-"

    return code, metrics, source


# ============================================================
# PENCERE METRİKLERİ
# ============================================================

def calculate_window_metrics(prices, returns, window) -> Optional[dict]:
    if len(returns) < window or len(prices) < window + 1:
        return None
    slice_returns = returns[-window:]
    slice_prices = prices[-(window + 1):]
    mean_return = sum(slice_returns) / len(slice_returns)
    variance = sum((r - mean_return) ** 2 for r in slice_returns) / len(slice_returns)
    volatility = variance ** 0.5
    sharpe_like = (mean_return / volatility) if volatility > 1e-12 else 0.0
    cumulative_return = (slice_prices[-1] / slice_prices[0] - 1.0) * 100.0
    max_dd = calculate_max_drawdown(slice_prices)
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

def calculate_security_scores(funds: List[dict]) -> None:
    by_kind: Dict[str, List[int]] = defaultdict(list)
    for idx, fund in enumerate(funds):
        by_kind[fund.get("kind", DEFAULT_FUND_KIND)].append(idx)

    kind_valor_median: Dict[str, float] = {}
    for kind, indices in by_kind.items():
        valors = [safe_float(funds[i].get("valor")) for i in indices]
        kind_valor_median[kind] = statistics.median(valors) if valors else 0.0

    for kind, indices in by_kind.items():
        subset = [funds[i] for i in indices]
        aum_z = zscore([safe_float(f.get("aum")) for f in subset])
        investor_z = zscore([safe_float(f.get("investors")) for f in subset])
        aum_flow_z = zscore([safe_float(f.get("aum_flow_proxy")) for f in subset])
        investor_change_z = zscore([safe_float(f.get("inv_change")) for f in subset])
        concentration_z = zscore([f.get("top_asset_weight") for f in subset])

        for local_i, fund_idx in enumerate(indices):
            fund = funds[fund_idx]
            score = 50.0
            score += SECURITY_SCALE["aum"] * SECURITY_WEIGHTS["aum"] * aum_z[local_i]
            score += SECURITY_SCALE["investor"] * SECURITY_WEIGHTS["investor"] * investor_z[local_i]
            score += SECURITY_SCALE["aum_flow"] * aum_flow_z[local_i]
            score += SECURITY_SCALE["investor_change"] * investor_change_z[local_i]

            if fund.get("top_asset_weight") is not None:
                score -= SECURITY_SCALE["concentration"] * SECURITY_WEIGHTS["concentration"] * concentration_z[local_i]

            if fund.get("is_bist30", False):
                score += BIST30_BONUS

            cash_ratio = fund.get("emergency_cash_ratio")
            if fund.get("cash_ratio_known", False) and cash_ratio is not None:
                if cash_ratio >= 15:
                    score += HIGH_LIQUIDITY_BONUS
                elif cash_ratio < 5:
                    score -= LOW_LIQUIDITY_PENALTY

            if safe_float(fund.get("inv_change")) > 0:
                score += POSITIVE_INVESTOR_FLOW_BONUS

            top_asset = fund.get("top_asset
