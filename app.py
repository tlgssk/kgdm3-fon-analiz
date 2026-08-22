import concurrent.futures
import datetime as dt
import io
import re
import time
from collections import defaultdict
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Tuple

import openpyxl
from openpyxl.styles.fills import PatternFill

# openpyxl 'extLst' Excel biçimlendirme hatası yaması (Monkey Patch)
original_init = PatternFill.__init__
def new_init(self, *args, **kwargs):
    if 'extLst' in kwargs:
        del kwargs['extLst']
    original_init(self, *args, **kwargs)
PatternFill.__init__ = new_init

import pandas as pd
import requests
import streamlit as st

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ============================================================
# KGDM-3 & KAZRİSK - SÜRÜM V8.6 (KAZRİSK UYUMLU)
# ============================================================

st.set_page_config(
    page_title="KGDM-3 & KAZRİSK Hibrit Fon Analizi",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK Hibrit Fon Analizi")
st.caption(
    "TEFAS + TEFAS Direct API + İş Yatırım | "
    "Net Likidite + Serbest Fon Filtresi + 2 Mum Teyit Kuralı | V8.6"
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

APP_VERSION = "8.5.0"

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
FREE_FUND_PENALTY = 7.0  # Serbest Fon Cezası (-7 Puan)

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

ENABLE_FILTERS = st.sidebar.checkbox("Filtreleri Etkinleştir", value=False)

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
    st.caption("KAZRİSK & KGDM-3 Skorlama Ağırlıkları")

    w_return = st.slider("Getiri ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["return"], 0.05)
    w_sharpe = st.slider("Sharpe ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["sharpe"], 0.05)
    w_cumulative = st.slider("Kümülatif ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["cumulative"], 0.05)
    w_drawdown = st.slider("Drawdown ağırlığı", 0.0, 1.0, DEFAULT_MOMENTUM_WEIGHTS["drawdown"], 0.05)

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
        "Hibrit Momentum Ağırlığı", 0.0, 1.0, DEFAULT_HYBRID_MOMENTUM_WEIGHT, 0.05
    )
    HYBRID_MOMENTUM_WEIGHT = hybrid_momentum_w
    HYBRID_SECURITY_WEIGHT = 1.0 - hybrid_momentum_w

with st.sidebar.expander("🔧 Tanılama"):
    SHOW_DIAGNOSTICS = st.checkbox("Kaynak tanılama bilgisini göster", value=True)


# ============================================================
# HTTP & YARDIMCI FONKSİYONLAR
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
        allowed_methods=frozenset({"GET", "POST"}),
        respect_retry_after_header=True,
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=MAX_WORKERS, pool_maxsize=MAX_WORKERS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": "KGDM3-Fon-Analiz/8.5 (+kazrisk-engine)",
        "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
    })
    return session

HTTP = build_http_session()

def request_with_status(source: str, method: str, url: str, *, params=None, data=None, json_body=None, headers=None, timeout=HTTP_TIMEOUT):
    status = new_status(source)
    status.attempted = True
    started = time.perf_counter()
    try:
        response = HTTP.request(method=method, url=url, params=params, data=data, json=json_body, headers=headers, timeout=timeout)
        status.status_code = response.status_code
        status.elapsed_ms = int((time.perf_counter() - started) * 1000)
        if response.status_code == 200:
            status.ok = True
            status.message = "OK"
        else:
            status.error_type = f"HTTP_{response.status_code}"
            status.message = f"HTTP {response.status_code}"
        return response, status
    except Exception as exc:
        status.error_type = "ERROR"
        status.message = str(exc)[:200]
        status.elapsed_ms = int((time.perf_counter() - started) * 1000)
        return None, status

def safe_float(value, default=0.0) -> float:
    try:
        if value is None: return default
        n = float(value)
        return default if pd.isna(n) else n
    except Exception:
        return default

_THOUSANDS_ONLY_RE = re.compile(r"^-?\d{1,3}(\.\d{3})+$")

def parse_number(value) -> Optional[float]:
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)):
        return None if pd.isna(value) else float(value)
    text = str(value).replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
    if not text: return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    elif "." in text and _THOUSANDS_ONLY_RE.match(text):
        text = text.replace(".", "")
    try:
        return float(text)
    except Exception:
        return None

def normalize_fund_code(value) -> str:
    if value is None: return ""
    code = str(value).strip().upper()
    return code[:-2] if code.endswith(".0") else code

def format_percent(value) -> str:
    n = parse_number(value)
    if n is None: return "-"
    return f"+%{n:.2f}" if n > 0 else (f"-%{abs(n):.2f}" if n < 0 else "%0.00")

def clamp(value, low, high):
    return max(low, min(high, value))

def calculate_compounded_return(returns) -> float:
    clean = [parse_number(v) for v in returns if parse_number(v) is not None]
    if not clean: return 0.0
    growth = 1.0
    for r in clean: growth *= (1.0 + r / 100.0)
    return (growth - 1.0) * 100.0

def calculate_max_drawdown(prices) -> float:
    if not prices or len(prices) < 2: return 0.0
    peak = safe_float(prices[0])
    max_dd = 0.0
    for p in prices:
        p = safe_float(p)
        if p <= 0: continue
        if p > peak: peak = p
        if peak > 0:
            dd = (p / peak - 1.0) * 100.0
            if dd < max_dd: max_dd = dd
    return max_dd

def zscore(values) -> List[float]:
    valid = [safe_float(v) for v in values if v is not None and not pd.isna(v)]
    if len(valid) < 2: return [0.0 for _ in values]
    mean_v = sum(valid) / len(valid)
    var = sum((x - mean_v) ** 2 for x in valid) / len(valid)
    std = var ** 0.5
    if std <= 1e-12: return [0.0 for _ in values]
    return [clamp((safe_float(v) - mean_v) / std, -Z_LIMIT, Z_LIMIT) if v is not None else 0.0 for v in values]

def population_mean_std(values: List[Optional[float]]) -> Tuple[float, float]:
    valid = [v for v in values if v is not None]
    if len(valid) < 2: return 0.0, 0.0
    mean_v = sum(valid) / len(valid)
    var = sum((v - mean_v) ** 2 for v in valid) / len(valid)
    return mean_v, var ** 0.5

def zscore_against_population(value: Optional[float], mean_v: float, std_v: float) -> float:
    if value is None or std_v <= 1e-12: return 0.0
    return clamp((value - mean_v) / std_v, -Z_LIMIT, Z_LIMIT)

def calculate_valor_penalty(excess_valor) -> float:
    excess_valor = safe_float(excess_valor)
    if excess_valor <= 0: return 0.0
    return clamp(excess_valor / 3.0, 0.0, 1.0) * MAX_VALOR_PENALTY

def decision_label_from_score(score) -> str:
    if score is None or score == "": return "YETERSİZ VERİ"
    try: s = float(score)
    except: return "-"
    if s >= STRONG_BUY: return "GÜÇLÜ AL"
    if s >= WATCH_LIST: return "ASIL LİSTE"
    if s >= CORRECTION: return "DÜZELTME / İZLE"
    return "ACİL SAT"


# ============================================================
# GITHUB & TEFAS VERİ ÇEKİMİ
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def resolve_latest_github_excel_url() -> Optional[str]:
    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/?ref={GITHUB_BRANCH}"
    response, status = request_with_status("GitHub", "GET", api_url, headers={"Accept": "application/vnd.github+json"})
    if response is None or not status.ok: return None
    try:
        items = response.json()
        xlsx_files = [item for item in items if isinstance(item, dict) and str(item.get("name", "")).lower().endswith(".xlsx") and item.get("download_url")]
        if not xlsx_files: return None
        xlsx_files.sort(key=lambda x: x.get("name", ""), reverse=True)
        return xlsx_files[0]["download_url"]
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
        crawler = Crawler(timeout=60, max_retry=3)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
        if df is None or df.empty: return pd.DataFrame()
        df.rename(columns={"fund_code": "code", "fund_name": "title", "investor_count": "investors", "portfolio_size": "aum", "fund_type": "kind"}, inplace=True)
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
    if universe is None or universe.empty: return meta
    try:
        latest = universe.sort_values("date").drop_duplicates(subset=["code"], keep="last")
        for _, row in latest.iterrows():
            code = str(row.get("code", "")).strip().upper()
            if not code: continue
            meta[code] = {
                "kind": str(row.get("kind", DEFAULT_FUND_KIND)).strip().upper(),
                "title": str(row.get("title", "") or "").strip(),
            }
    except Exception:
        pass
    return meta

def build_universe_reference(universe: pd.DataFrame, window: int) -> Dict[str, Dict[str, List[float]]]:
    reference = {k: {"mean_return": [], "sharpe": [], "cumulative": [], "max_dd_inv": []} for k in FUND_KINDS}
    if universe is None or universe.empty or window < 2: return reference
    for code, group in universe.groupby("code"):
        group = group.sort_values("date")
        kind = str(group["kind"].iloc[-1]).strip().upper()
        if kind not in FUND_KINDS: continue
        prices = group["price"].astype(float).tolist()
        if len(prices) < window + 1: continue
        window_prices = prices[-(window + 1):]
        returns = [0.0 if p0 <= 0 else (p1 / p0 - 1.0) * 100.0 for p0, p1 in zip(window_prices[:-1], window_prices[1:])]
        mean_r = sum(returns) / len(returns)
        var = sum((r - mean_r) ** 2 for r in returns) / len(returns)
        vol = var ** 0.5
        sharpe = mean_r / vol if vol > 1e-12 else 0.0
        cum = (window_prices[-1] / window_prices[0] - 1.0) * 100.0
        max_dd = calculate_max_drawdown(window_prices)
        reference[kind]["mean_return"].append(mean_r)
        reference[kind]["sharpe"].append(sharpe)
        reference[kind]["cumulative"].append(cum)
        reference[kind]["max_dd_inv"].append(-max_dd)
    return reference

def reference_sample_size(reference, kind) -> int:
    return len(reference.get(kind, {}).get("mean_return", []))


# ============================================================
# TEFAS PORTFÖY DAĞILIMI & NET LİKİDİTE (KAZRİSK UYUMLU)
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 60 * 2)
def fetch_tefas_breakdown_snapshot(fund_kind: Optional[str], reference_date: Optional[str]) -> dict:
    kind = (fund_kind or "YAT").upper()
    ref = pd.to_datetime(reference_date).date() if reference_date else dt.date.today()
    url = "https://www.tefas.gov.tr/api/funds/dagilimSiraliGetirT"
    body = {
        "fonTipi": kind, "fonKodu": None, "aramaMetni": None, "fonTurKod": None,
        "fonGrubu": None, "sfonTurKod": None, "fonTurAciklama": None, "kurucuKod": None,
        "basTarih": ref.strftime("%Y%m%d"), "bitTarih": ref.strftime("%Y%m%d"),
        "basSira": 1, "bitSira": 100000, "dil": "TR", "sFonTurKod": "", "fonKod": "",
        "fonGrup": "", "fonUnvanTip": "",
    }
    response, status = request_with_status("TEFAS Direct Structural API", "POST", url, json_body=body, headers={"Accept": "*/*", "Content-Type": "application/json", "Origin": "https://www.tefas.gov.tr", "Referer": "https://www.tefas.gov.tr/tr/fon-verileri"})
    if response is None or not status.ok:
        return {"ok": False, "rows": {}}
    try:
        result_list = response.json().get("resultList") or []
        rows = {}
        field_map = {
            "hs": "stock_pct", "dt": "government_bond_pct", "hb": "treasury_bill_pct", "fb": "financing_bill_pct",
            "ost": "private_sector_bond_pct", "bb": "bank_bill_pct", "vdm": "asset_backed_securities_pct",
            "eut": "eurobond_pct", "kibd": "government_external_debt_pct", "osdb": "private_sector_external_debt_pct",
            "kba": "fx_government_internal_debt_pct", "tpp": "takasbank_money_market_pct", "bpp": "bist_money_market_pct",
            "r": "repo_pct", "tr": "reverse_repo_pct", "vm": "term_deposit_pct", "vmtl": "deposit_tl_pct", "vmd": "deposit_fx_pct",
            "km": "precious_metals_pct", "ymk": "foreign_security_pct", "yhs": "foreign_stock_pct", "fkb": "fund_participation_pct", "d": "other_pct"
        }
        for row in result_list:
            code = normalize_fund_code(row.get("fonKodu") or row.get("fonKod"))
            if not code: continue
            parsed = {"fund_code": code, "fund_name": row.get("fonUnvan"), "date": row.get("tarih")}
            for short, target in field_map.items():
                parsed[target] = parse_number(row.get(short))
            rows[code] = parsed
        return {"ok": bool(rows), "rows": rows}
    except Exception:
        return {"ok": False, "rows": {}}


@st.cache_data(show_spinner=False, ttl=60 * 60 * 2)
def fetch_fund_structural_data(fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None, reference_date: Optional[str] = None) -> dict:
    code = normalize_fund_code(fund_code)
    structural = {
        "top_asset_weight": None, "top_asset_weight_basis": None, "is_bist30": False,
        "is_free_fund": False,  # Serbest fon bayrağı
        "emergency_cash_ratio": None, "cash_ratio_known": False, "structural_fetch_ok": False,
        "structural_source": "YOK", "fund_title": fund_title if fund_title else None,
        "investment_area": None,
    }
    if not code: return structural

    title_upper = (fund_title or "").upper()
    if "SERBEST" in title_upper or fund_kind == "BYF":
        structural["is_free_fund"] = True
        structural["investment_area"] = "Serbest"
    elif "PARA PİYASASI" in title_upper or "PPF" in title_upper or "LİKİT" in title_upper:
        structural["investment_area"] = "Para Piyasası"
    elif "ALTIN" in title_upper or "KIYMETLİ MADEN" in title_upper or "GÜMÜŞ" in title_upper:
        structural["investment_area"] = "Kıymetli Maden"
    elif "BIST 30" in title_upper or "BIST30" in title_upper:
        structural["is_bist30"] = True
        structural["investment_area"] = "Hisse Senedi (BIST 30)"
    elif "HİSSE SENEDİ" in title_upper:
        structural["investment_area"] = "Hisse Senedi"

    snapshot = fetch_tefas_breakdown_snapshot(fund_kind or "YAT", reference_date)
    if snapshot.get("ok"):
        row = snapshot.get("rows", {}).get(code)
        if row:
            # En büyük varlık ağırlığı
            pct_fields = [k for k, v in row.items() if k.endswith("_pct") and v is not None]
            valid_alloc = [(f, safe_float(row.get(f))) for f in pct_fields if row.get(f) is not None]
            if valid_alloc:
                top_field, top_val = max(valid_alloc, key=lambda x: x[1])
                structural["top_asset_weight"] = top_val
                structural["top_asset_weight_basis"] = top_field

            # KAZRİSK KURALI: Net Likidite = (Ters Repo + Mevduat + Takasbank/BIST Para Piyasası) - REPO (Eksi Borç)
            positive_cash_fields = ["reverse_repo_pct", "term_deposit_pct", "deposit_tl_pct", "takasbank_money_market_pct", "bist_money_market_pct"]
            positive_cash = sum(safe_float(row.get(f)) for f in positive_cash_fields if row.get(f) is not None)
            repo_liability = safe_float(row.get("repo_pct", 0.0))  # Repo borçluluktur
            
            # Net acil durum likiditesi
            net_liquidity = max(0.0, positive_cash - repo_liability)
            structural["emergency_cash_ratio"] = clamp(net_liquidity, 0.0, 100.0)
            structural["cash_ratio_known"] = True
            structural["structural_fetch_ok"] = True
            structural["structural_source"] = "TEFAS Direct Structural API"
            return structural

    return structural


# ============================================================
# METRİK & GÜVENLİK SKORLAMA (KAZRİSK)
# ============================================================

def fetch_isyatirim_series(fund_code: str) -> Tuple[Optional[pd.DataFrame], SourceStatus]:
    code = normalize_fund_code(fund_code)
    status = new_status("İş Yatırım")
    if not code: return None, status
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    response, status = request_with_status("İş Yatırım", "GET", url, params=params, headers={"Accept": "application/json"})
    if response is None or not status.ok: return None, status
    try:
        values = response.json().get("value", [])
        if not values: return None, status
        df = pd.DataFrame(values)
        df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
        df["price"] = df["Fiyat"].apply(parse_number)
        df["aum"], df["investors"] = 0.0, 0.0
        df = df.dropna(subset=["date", "price"])
        df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
        return df[["date", "price", "aum", "investors"]], status
    except Exception:
        return None, status

def fetch_tefas_direct_api(fund_code: str, fund_kind: Optional[str] = None) -> Tuple[Optional[pd.DataFrame], SourceStatus]:
    code = normalize_fund_code(fund_code)
    status = new_status("TEFAS Direct API")
    if not code: return None, status
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.tefas.gov.tr/api/DB/BindHistoryInfo"
    headers = {"X-Requested-With": "XMLHttpRequest", "Origin": "https://www.tefas.gov.tr", "Referer": "https://www.tefas.gov.tr/"}
    kind_candidates = [fund_kind] if fund_kind in FUND_KINDS else list(FUND_KINDS)
    for kind in kind_candidates:
        payload = {"fontip": kind, "fonkod": code, "bastarih": start.strftime("%d.%m.%Y"), "bittarih": end.strftime("%d.%m.%Y")}
        response, attempt_status = request_with_status("TEFAS Direct API", "POST", url, data=payload, headers=headers)
        if response is None or not attempt_status.ok: continue
        try:
            data = response.json().get("data", [])
            if not data: continue
            df = pd.DataFrame(data)
            df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
            df["price"] = df["FIYAT"].apply(parse_number)
            df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number).fillna(0.0) if "PORTFOYBUYUKLUK" in df.columns else 0.0
            df["investors"] = df["KISISAYISI"].apply(parse_number).fillna(0.0) if "KISISAYISI" in df.columns else 0.0
            df = df.dropna(subset=["date", "price"])
            df = df[df["price"] > 0].sort_values("date").drop_duplicates(subset=["date"], keep="last").tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
            attempt_status.ok = True
            return df, attempt_status
        except Exception:
            continue
    return None, status

def get_fund_series(universe: pd.DataFrame, fund_code: str, fund_kind: Optional[str] = None):
    code = normalize_fund_code(fund_code)
    source_statuses = []
    if not code: return None, "YOK", source_statuses
    if universe is not None and not universe.empty and "code" in universe.columns:
        rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
        if not rows.empty:
            rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
            if len(rows) >= 2:
                source_statuses.append(SourceStatus(source="TEFAS", attempted=True, ok=True, message=f"{len(rows)} gözlem"))
                return rows.tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True), "TEFAS", source_statuses
    direct_df, direct_status = fetch_tefas_direct_api(code, fund_kind)
    source_statuses.append(direct_status)
    if direct_df is not None and len(direct_df) >= 2:
        return direct_df, "TEFAS Direct API", source_statuses
    is_df, is_status = fetch_isyatirim_series(code)
    source_statuses.append(is_status)
    if is_df is not None and len(is_df) >= 2:
        return is_df, "İş Yatırım", source_statuses
    return None, "YOK", source_statuses

def compute_fund_metrics(series: Optional[pd.DataFrame], fund_code: str, fund_kind: Optional[str] = None, fund_title: Optional[str] = None) -> Optional[dict]:
    if series is None or len(series) < 2: return None
    df = series.copy().sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    prices = df["price"].astype(float).tolist()
    dates = pd.to_datetime(df["date"]).dt.strftime("%d.%m").tolist()
    aums = df["aum"].astype(float).tolist()
    investors = df["investors"].astype(float).tolist()
    daily_returns = [0.0 if prices[i-1] <= 0 else (prices[i] / prices[i-1] - 1.0) * 100.0 for i in range(1, len(prices))]
    if not daily_returns: return None
    max_dd = calculate_max_drawdown(prices)
    aum_change = ((aums[-1] / aums[0] - 1.0) * 100.0 if aums[0] > 0 else None)
    inv_change = ((investors[-1] / investors[0] - 1.0) * 100.0 if investors[0] > 0 else None)
    price_ret = ((prices[-1] / prices[0] - 1.0) * 100.0 if prices[0] > 0 else 0.0)
    aum_flow_proxy = (aum_change - price_ret) if aum_change is not None else None
    recent_weekly = daily_returns[-5:] if len(daily_returns) >= 5 else daily_returns
    weekly_return = calculate_compounded_return(recent_weekly)
    ref_date = pd.to_datetime(df["date"].iloc[-1]).strftime("%Y-%m-%d") if not df.empty else None
    structural = fetch_fund_structural_data(fund_code, fund_kind, fund_title, ref_date)

    return {
        "dates": dates[1:], "prices": prices, "daily_returns": daily_returns, "n_days": len(daily_returns),
        "aum": aums[-1], "investors": int(round(investors[-1])), "aum_change": aum_change,
        "aum_flow_proxy": aum_flow_proxy, "inv_change": inv_change, "max_dd": max_dd,
        "weekly_return": weekly_return, **structural,
    }

def fetch_and_compute_one_fund(code: str, universe: pd.DataFrame, meta_map: Dict[str, Dict[str, str]], valor_dict: Dict[str, float]) -> Tuple[str, Optional[dict], str]:
    meta = meta_map.get(code, {})
    series, source, statuses = get_fund_series(universe, code, meta.get("kind"))
    metrics = compute_fund_metrics(series, code, meta.get("kind"), meta.get("title"))
    if metrics is None: return code, None, source
    metrics["code"] = code
    metrics["valor"] = valor_dict.get(code, 0.0)
    metrics["source"] = source
    metrics["kind"] = meta.get("kind", DEFAULT_FUND_KIND)
    metrics["source_statuses"] = [asdict(x) for x in statuses]
    metrics["source_chain"] = " → ".join(x.source for x in statuses if x.attempted)
    title = (metrics.get("fund_title") or meta.get("title") or "").strip()
    metrics["fund_title"] = title if title else "-"
    return code, metrics, source


# ============================================================
# KAZRİSK SKORLAMA MOTORU
# ============================================================

def calculate_security_scores(funds: List[dict]) -> None:
    by_kind = defaultdict(list)
    for idx, fund in enumerate(funds):
        by_kind[fund.get("kind", DEFAULT_FUND_KIND)].append(idx)

    for kind, indices in by_kind.items():
        subset = [funds[i] for i in indices]
        aum_z = zscore([f.get("aum") for f in subset])
        inv_z = zscore([f.get("investors") for f in subset])
        aum_flow_z = zscore([f.get("aum_flow_proxy") for f in subset])
        inv_change_z = zscore([f.get("inv_change") for f in subset])
        conc_z = zscore([f.get("top_asset_weight") for f in subset])

        for local_i, fund_idx in enumerate(indices):
            fund = funds[fund_idx]
            score = 50.0

            score += SECURITY_SCALE["aum"] * SECURITY_WEIGHTS["aum"] * aum_z[local_i]
            score += SECURITY_SCALE["investor"] * SECURITY_WEIGHTS["investor"] * inv_z[local_i]
            if fund.get("aum_flow_proxy") is not None: score += SECURITY_SCALE["aum_flow"] * aum_flow_z[local_i]
            if fund.get("inv_change") is not None: score += SECURITY_SCALE["investor_change"] * inv_change_z[local_i]

            top_asset = fund.get("top_asset_weight")
            if top_asset is not None:
                score -= SECURITY_SCALE["concentration"] * SECURITY_WEIGHTS["concentration"] * conc_z[local_i]
                if top_asset > 30: score -= min((top_asset - 30) * 1.0, MAX_CONCENTRATION_PENALTY)
                elif top_asset > 15: score -= (top_asset - 15) * 0.25

            # KAZRİSK Kuralı: BIST 30 Derinlik Bonusu
            if fund.get("is_bist30", False): score += BIST30_BONUS

            # KAZRİSK Kuralı: Serbest Fon Cezası (-7 Puan)
            if fund.get("is_free_fund", False): score -= FREE_FUND_PENALTY

            # Net Likidite Bonusu / Cezası
            cash_ratio = fund.get("emergency_cash_ratio")
            if fund.get("cash_ratio_known", False) and cash_ratio is not None:
                if cash_ratio >= 15: score += HIGH_LIQUIDITY_BONUS
                elif cash_ratio < 5: score -= LOW_LIQUIDITY_PENALTY

            if fund.get("inv_change") is not None and safe_float(fund.get("inv_change")) > 0:
                score += POSITIVE_INVESTOR_FLOW_BONUS

            # DÜZELTME: Valör (takas) süresi cezası artık gerçekten uygulanıyor.
            # T+1 baz kabul edilir; fazla her gün likidite riskini artırır.
            valor_days = safe_float(fund.get("valor"), 0.0)
            if valor_days > 1:
                score -= calculate_valor_penalty(valor_days - 1.0)

            fund["security_score"] = int(round(clamp(score, 0.0, 100.0)))

def calculate_market_relative_momentum(funds: List[dict], reference, final_window: int) -> None:
    for fund in funds:
        kind = fund.get("kind", DEFAULT_FUND_KIND)
        sample_size = reference_sample_size(reference, kind)
        window = min(final_window, fund["n_days"])
        
        # Pencere metrikleri
        slice_r = fund["daily_returns"][-window:]
        slice_p = fund["prices"][-(window + 1):]
        mean_r = sum(slice_r) / len(slice_r) if slice_r else 0.0
        var = sum((r - mean_r) ** 2 for r in slice_r) / len(slice_r) if slice_r else 0.0
        vol = var ** 0.5
        sharpe = mean_r / vol if vol > 1e-12 else 0.0
        cum = (slice_p[-1] / slice_p[0] - 1.0) * 100.0 if len(slice_p) > 1 and slice_p[0] > 0 else 0.0
        max_dd = calculate_max_drawdown(slice_p)

        fund["_final_mean_return"] = mean_r
        fund["_final_sharpe"] = sharpe
        fund["_final_cumulative"] = cum
        fund["_final_max_dd"] = max_dd
        fund["volatility"] = vol

        if sample_size >= MIN_REFERENCE_SAMPLE:
            ref = reference[kind]
            mean_m, mean_s = population_mean_std(ref["mean_return"])
            sharpe_m, sharpe_s = population_mean_std(ref["sharpe"])
            cum_m, cum_s = population_mean_std(ref["cumulative"])
            dd_m, dd_s = population_mean_std(ref["max_dd_inv"])

            z_mean = zscore_against_population(mean_r, mean_m, mean_s)
            z_sharpe = zscore_against_population(sharpe, sharpe_m, sharpe_s)
            z_cum = zscore_against_population(cum, cum_m, cum_s)
            z_dd = zscore_against_population(-max_dd, dd_m, dd_s)
            fund["reference_scope"] = f"Piyasa ({kind}, n={sample_size})"
        else:
            fallback = [f for f in funds if f.get("kind") == kind]
            local_i = fallback.index(fund)
            z_mean = zscore([f.get("_final_mean_return") for f in fallback])[local_i]
            z_sharpe = zscore([f.get("_final_sharpe") for f in fallback])[local_i]
            z_cum = zscore([f.get("_final_cumulative") for f in fallback])[local_i]
            z_dd = zscore([-safe_float(f.get("_final_max_dd")) for f in fallback])[local_i]
            fund["reference_scope"] = "Liste-bağıl"

        weighted_z = (
            MOMENTUM_WEIGHTS["return"] * z_mean
            + MOMENTUM_WEIGHTS["sharpe"] * z_sharpe
            + MOMENTUM_WEIGHTS["cumulative"] * z_cum
            + MOMENTUM_WEIGHTS["drawdown"] * z_dd
        )

        momentum_score = 50.0 + 20.0 * weighted_z

        # KAZRİSK Kuralı: Yeni Trend Mutlak Getiri Bonusu (Haftalık >= %5 ise +10 Puan)
        weekly_ret = safe_float(fund.get("weekly_return"))
        if weekly_ret >= 5.0: momentum_score += 10.0
        elif weekly_ret <= -5.0: momentum_score -= 10.0

        daily_rets = fund.get("daily_returns") or []
        last_day = daily_rets[-1] if daily_rets else 0.0
        last_2_avg = sum(daily_rets[-2:]) / 2.0 if len(daily_rets) >= 2 else last_day
        overheat = (z_cum >= OVERHEAT_Z_THRESHOLD and (last_day < 0 or last_2_avg < 0))
        fund["overheat_flag"] = overheat
        if overheat: momentum_score -= OVERHEAT_PENALTY

        fund["market_momentum"] = int(round(clamp(momentum_score, 0.0, 100.0)))

def calculate_trend_scores(funds: List[dict]) -> int:
    if not funds: return 0
    n_days = min(f["n_days"] for f in funds)
    for fund in funds:
        fund["dates"] = fund["dates"][-n_days:]
        fund["daily_returns"] = fund["daily_returns"][-n_days:]
        fund["prices"] = fund["prices"][-(n_days + 1):]
        fund["running_trend_momentum"] = []

    for d in range(1, n_days + 1):
        if d < MIN_ROLLING_DAYS:
            for fund in funds: fund["running_trend_momentum"].append(None)
            continue
        current_metrics = []
        for fund in funds:
            r_slice = fund["daily_returns"][d - MIN_ROLLING_DAYS:d]
            p_slice = fund["prices"][d - MIN_ROLLING_DAYS:d + 1]
            if len(r_slice) < MIN_ROLLING_DAYS: continue
            mean_r = sum(r_slice) / len(r_slice)
            var = sum((r - mean_r) ** 2 for r in r_slice) / len(r_slice)
            vol = var ** 0.5
            sharpe = mean_r / vol if vol > 1e-12 else 0.0
            cum = calculate_compounded_return(r_slice)
            max_dd = calculate_max_drawdown(p_slice)
            current_metrics.append({"fund": fund, "mean_r": mean_r, "sharpe": sharpe, "cum": cum, "max_dd": max_dd})

        if not current_metrics: continue
        mean_z = zscore([x["mean_r"] for x in current_metrics])
        sharpe_z = zscore([x["sharpe"] for x in current_metrics])
        cum_z = zscore([x["cum"] for x in current_metrics])
        dd_z = zscore([-x["max_dd"] for x in current_metrics])

        for i, data in enumerate(current_metrics):
            w_z = (MOMENTUM_WEIGHTS["return"] * mean_z[i] + MOMENTUM_WEIGHTS["sharpe"] * sharpe_z[i] + MOMENTUM_WEIGHTS["cumulative"] * cum_z[i] + MOMENTUM_WEIGHTS["drawdown"] * dd_z[i])
            mom = clamp(50.0 + 20.0 * w_z, 0.0, 100.0)
            data["fund"]["running_trend_momentum"].append(int(round(mom)))

    for fund in funds:
        sec_score = safe_float(fund.get("security_score"), 50.0)
        running_hybrid = []
        for mom in fund["running_trend_momentum"]:
            if mom is None: running_hybrid.append(None)
            else:
                hyb = mom * HYBRID_MOMENTUM_WEIGHT + sec_score * HYBRID_SECURITY_WEIGHT
                running_hybrid.append(int(round(clamp(hyb, 0.0, 100.0))))
        fund["running_trend_hybrid"] = running_hybrid
        valid_last = [s for s in running_hybrid if s is not None][-5:]
        fund["last_5_scores"] = valid_last
        fund["last_5_scores_str"] = " ➔ ".join(str(x) for x in valid_last) if valid_last else "-"
        if valid_last:
            n = len(valid_last)
            weights = [EMA_DECAY ** (n - 1 - i) for i in range(n)]
            trend = sum(s * w for s, w in zip(valid_last, weights)) / sum(weights)
            fund["trend_skor"] = int(round(trend))
        else:
            fund["trend_skor"] = None
    return n_days

def finalize_decisions(funds: List[dict]) -> None:
    for fund in funds:
        mom = fund.get("market_momentum")
        sec = fund.get("security_score")
        if mom is None or sec is None:
            fund["decision_score"] = None
            fund["karar"] = "YETERSİZ VERİ"
            continue
        dec_score = clamp(mom * HYBRID_MOMENTUM_WEIGHT + sec * HYBRID_SECURITY_WEIGHT, 0.0, 100.0)
        fund["decision_score"] = int(round(dec_score))
        fund["karar"] = decision_label_from_score(fund["decision_score"])

def compute_confidence_label(fund: dict) -> str:
    score = 0
    if fund.get("n_days", 0) >= TARGET_TRADING_DAYS: score += 30
    if fund.get("source") == "TEFAS": score += 30
    if fund.get("structural_fetch_ok", False): score += 20
    if fund.get("aum", 0) > 0: score += 20
    return f"🟢 Yüksek ({score})" if score >= 80 else (f"🟡 Orta ({score})" if score >= 60 else f"🔴 Düşük ({score})")


# ============================================================
# EXCEL ÇIKTISI (SON 5 GÜNLÜK VERİLERLE)
# ============================================================

def create_excel_output(wb, ws_list, all_funds_for_output, common_n_days):
    if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

    sample_dates = []
    for item in all_funds_for_output:
        if item.get("dates") and len(item["dates"]) >= common_n_days:
            sample_dates = item["dates"][-common_n_days:]
            break
    if not sample_dates:
        for item in all_funds_for_output:
            if item.get("dates"):
                sample_dates = item["dates"]
                break

    last_5_asc = sample_dates[-5:] if len(sample_dates) >= 5 else sample_dates
    # DÜZELTME: D sütunundan itibaren EN GÜNCEL gün önce gelecek şekilde
    # (geriye doğru kronolojik) sıralanıyor.
    last_5_dates = list(reversed(last_5_asc))

    headers = ["Fon Kodu", "Fon Adı", "Yatırım Alanı"]
    for day in last_5_dates:
        headers.append(f"{day} Karar Skoru")
        headers.append(f"{day} Model Kararı")

    headers.extend([
        "Valör", "Karar Skoru (Piyasa-Bağıl)", "Trend Skoru (Liste-Bağıl)", "Piyasa Momentum",
        "Güvenlik/Likidite Skoru", "Referans Kapsamı", "Veri Güveni", "Aşırı Isınma",
        "Son 5 Trend Skoru", "Model Kararı", "Ort. Günlük Getiri (%)", "Volatilite (%)",
        "Sharpe-benzeri", "Kümülatif Getiri (%)", "MaxDD (%)", "En Büyük Varlık (%)",
        "BIST30", "Net Likidite (%)", "KAZRİSK Durumu", "AUM Değişim (%)", "AUM Akış Proxy (%)",
        "Yatırımcı Değişim (%)", "AUM (₺)", "Yatırımcı", "Haftalık Bileşik (%)",
        "Veri Kaynağı", "Kaynak Zinciri", "Serbest Fon mu?", "Yapısal Kaynak"
    ])

    sample_dates_desc = list(reversed(sample_dates))
    for day in sample_dates_desc: headers.append(f"{day} Trend Hibrit Skor")
    for day in sample_dates_desc: headers.append(f"{day} Getiri")
    ws_scores.append(headers)

    header_index = {name: idx + 1 for idx, name in enumerate(headers)}
    header_fill = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color=COLOR_WHITE)

    for cell in ws_scores[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_scores.row_dimensions[1].height = 55

    for item in all_funds_for_output:
        top_asset = item.get("top_asset_weight")
        risk_label = "⚪ Veri Yok" if top_asset is None else ("⚠️ Yüksek Konsantrasyon" if top_asset > 30 else ("🟡 Orta Konsantrasyon" if top_asset > 15 else "🛡️ Dengeli"))
        cash_label = f"%{safe_float(item.get('emergency_cash_ratio')):.2f}" if item.get("cash_ratio_known", False) else "Veri Yok"

        row_data = [item["code"], item.get("fund_title") or "-", item.get("investment_area") or "-"]
        own_scores = item.get("running_trend_hybrid") or []
        last_5_s_asc = own_scores[-5:] if len(own_scores) >= 5 else own_scores
        pad_len = len(last_5_dates) - len(last_5_s_asc)
        last_5_padded_asc = [None] * pad_len + last_5_s_asc
        # Tarih sütunlarıyla aynı sırada: en güncel gün önce
        last_5_padded = list(reversed(last_5_padded_asc))

        for score in last_5_padded:
            row_data.append(score if score is not None else "")
            row_data.append(decision_label_from_score(score))

        row_data.extend([
            item.get("valor", 0), item.get("decision_score"), item.get("trend_skor"),
            item.get("market_momentum"), item.get("security_score"), item.get("reference_scope", "-"),
            compute_confidence_label(item), ("🔥 Evet" if item.get("overheat_flag") else "-"),
            item.get("last_5_scores_str", "-"), item.get("karar", "-"),
            round(safe_float(item.get("_final_mean_return")), 4),
            (round(safe_float(item.get("volatility")), 4) if item.get("volatility") is not None else ""),
            round(safe_float(item.get("_final_sharpe")), 4), round(safe_float(item.get("_final_cumulative")), 4),
            round(safe_float(item.get("_final_max_dd")), 4),
            (round(safe_float(top_asset), 2) if top_asset is not None else None),
            ("EVET" if item.get("is_bist30", False) else "HAYIR"), cash_label, risk_label,
            (round(safe_float(item.get("aum_change")), 2) if item.get("aum_change") is not None else None),
            (round(safe_float(item.get("aum_flow_proxy")), 2) if item.get("aum_flow_proxy") is not None else None),
            (round(safe_float(item.get("inv_change")), 2) if item.get("inv_change") is not None else None),
            (round(safe_float(item.get("aum")), 2) if item.get("aum") is not None else None),
            (int(item.get("investors")) if item.get("investors") is not None else None),
            round(safe_float(item.get("weekly_return")), 4), item.get("source", "-"), item.get("source_chain", "-"),
            ("EVET" if item.get("is_free_fund") else "HAYIR"), item.get("structural_source", "YOK")
        ])

        n_dates = len(sample_dates)
        own_s_tail = ([None] * (n_dates - len(own_scores)) + own_scores) if len(own_scores) < n_dates else own_scores[-n_dates:]
        row_data.extend([s if s is not None else "" for s in reversed(own_s_tail)])

        own_rets = item.get("daily_returns") or []
        own_r_tail = ([None] * (n_dates - len(own_rets)) + own_rets) if len(own_rets) < n_dates else own_rets[-n_dates:]
        row_data.extend([(format_percent(x) if x is not None else "-") for x in reversed(own_r_tail)])
        ws_scores.append(row_data)

    # Renklendirme ve Biçimlendirme
    green_font = Font(bold=True, color=COLOR_GREEN)
    red_font = Font(bold=True, color=COLOR_RED)
    yellow_font = Font(bold=True, color=COLOR_YELLOW)
    dec_cols = [header_index[f"{day} Model Kararı"] for day in last_5_dates] + [header_index["Model Kararı"]]

    for r in range(2, ws_scores.max_row + 1):
        for c in dec_cols:
            cell = ws_scores.cell(row=r, column=c)
            txt = str(cell.value or "")
            if "GÜÇLÜ AL" in txt or "ASIL LİSTE" in txt: cell.font = green_font
            elif "DÜZELTME" in txt: cell.font = yellow_font
            elif "ACİL SAT" in txt: cell.font = red_font

    score_cols = [header_index[f"{day} Karar Skoru"] for day in last_5_dates] + [header_index["Karar Skoru (Piyasa-Bağıl)"]]
    for sc in score_cols:
        col_let = get_column_letter(sc)
        s_range = f"{col_let}2:{col_let}{ws_scores.max_row}"
        ws_scores.conditional_formatting.add(s_range, CellIsRule(operator="greaterThanOrEqual", formula=["75"], fill=PatternFill(start_color=COLOR_LIGHT_GREEN, end_color=COLOR_LIGHT_GREEN, fill_type="solid")))
        ws_scores.conditional_formatting.add(s_range, CellIsRule(operator="between", formula=["50", "74"], fill=PatternFill(start_color=COLOR_LIGHT_YELLOW, end_color=COLOR_LIGHT_YELLOW, fill_type="solid")))
        ws_scores.conditional_formatting.add(s_range, CellIsRule(operator="lessThan", formula=["50"], fill=PatternFill(start_color=COLOR_LIGHT_RED, end_color=COLOR_LIGHT_RED, fill_type="solid")))

    thin_gray = Side(style="thin", color="D9E1F2")
    for r in ws_scores.iter_rows():
        for cell in r:
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin_gray)
    ws_scores.freeze_panes = "A2"
    ws_scores.sheet_view.showGridLines = False

    for col in ws_scores.columns:
        max_l = max(len(str(cell.value or "")) for cell in col)
        ws_scores.column_dimensions[get_column_letter(col[0].column)].width = max(10, min(max_l + 3, 45))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, last_5_dates


# ============================================================
# ANA ARAYÜZ
# ============================================================

st.subheader("📂 Portföy Excel Listesi")
col_upload, col_github = st.columns(2)
wb = None

with col_upload:
    uploaded_file = st.file_uploader("Bilgisayardan Excel Yükle", type=["xlsx"])
    if uploaded_file is not None:
        try: wb = openpyxl.load_workbook(uploaded_file)
        except Exception as exc: st.error(f"Excel yükleme hatası: {exc}")

with col_github:
    st.write("Veya GitHub'daki listeyi kullanın:")
    if st.button("🚀 GitHub'dan Çek ve Analiz Et", use_container_width=True):
        resolved_url = resolve_latest_github_excel_url() or GITHUB_FALLBACK_URL
        response, status = request_with_status("GitHub Excel", "GET", resolved_url)
        if response is not None and status.ok:
            try:
                wb = openpyxl.load_workbook(io.BytesIO(response.content))
                st.success("✅ En güncel Excel dosyası indirildi.")
            except Exception as exc:
                st.error(f"Excel ayrıştırma hatası: {exc}")
        else:
            st.error(f"GitHub dosyası alınamadı: {status.error_type} {status.message}")

if wb is None:
    st.info("Analize başlamak için Excel dosyanızı yükleyin.")
    st.stop()

ws_list = wb["Fon_Listesi"] if "Fon_Listesi" in wb.sheetnames else wb.active
requested_codes, excel_valor_dict = [], {}

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if not row or row[0].value is None: continue
    code = normalize_fund_code(row[0].value)
    if not code: continue
    requested_codes.append(code)
    try: excel_valor_dict[code] = parse_number(row[3].value) if len(row) > 3 and parse_number(row[3].value) is not None else 0.0
    except: excel_valor_dict[code] = 0.0

requested_codes = list(dict.fromkeys(requested_codes))
if not requested_codes:
    st.error("Fon_Listesi sayfasında geçerli fon kodu bulunamadı.")
    st.stop()

today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

with st.spinner("🔄 TEFAS verileri alınıyor..."):
    universe = fetch_tefas_universe(start_date, today)
    fund_meta_map = build_fund_meta_map(universe)
    universe_reference = build_universe_reference(universe, window=TARGET_TRADING_DAYS)

calculated_funds, failed_codes = [], []
progress = st.progress(0, text="Fonlar analiz ediliyor...")
total_funds = len(requested_codes)

with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    future_to_code = {
        executor.submit(fetch_and_compute_one_fund, code, universe, fund_meta_map, excel_valor_dict): code
        for code in requested_codes
    }
    completed = 0
    for future in concurrent.futures.as_completed(future_to_code):
        code = future_to_code[future]
        completed += 1
        try: _, metrics, source = future.result()
        except Exception: metrics, source = None, "YOK"
        if metrics is None:
            failed_codes.append(code)
            progress.progress(completed / total_funds, text=f"{code}: veri yok")
            continue
        if ENABLE_FILTERS:
            inv_ok = safe_float(metrics.get("investors")) >= MIN_INVESTOR_COUNT
            wk_ok = safe_float(metrics.get("weekly_return")) >= TARGET_WEEKLY_RETURN
            if not (inv_ok and wk_ok): continue
        calculated_funds.append(metrics)
        progress.progress(completed / total_funds, text=f"{code}: {source}")

progress.empty()

eligible_funds = [f for f in calculated_funds if f.get("n_days", 0) >= MIN_ROLLING_DAYS]
insufficient_funds = [f for f in calculated_funds if f.get("n_days", 0) < MIN_ROLLING_DAYS]

for f in insufficient_funds:
    f["security_score"] = None
    f["market_momentum"] = None
    f["decision_score"] = None
    f["trend_skor"] = None
    f["karar"] = "YETERSİZ VERİ"
    f["running_trend_hybrid"] = []
    f["last_5_scores_str"] = "-"

if not eligible_funds:
    st.error(f"En az {MIN_ROLLING_DAYS} gün verisi olan geçerli fon bulunamadı.")
    st.stop()

with st.spinner("📊 KAZRİSK & KGDM-3 hibrit skorları hesaplanıyor..."):
    calculate_security_scores(eligible_funds)
    calculate_market_relative_momentum(eligible_funds, universe_reference, final_window=TARGET_TRADING_DAYS)
    common_n_days = calculate_trend_scores(eligible_funds)
    finalize_decisions(eligible_funds)

all_funds_for_output = eligible_funds + insufficient_funds
all_funds_for_output.sort(key=lambda x: (-safe_float(x.get("decision_score")), -safe_float(x.get("_final_cumulative"))))

output, last_5_dates_for_ui = create_excel_output(wb, ws_list, all_funds_for_output, common_n_days)


# ============================================================
# EKRAN TABLOSU & KAZRİSK 2 GÜN TEYİT ALARMLARI
# ============================================================

display_rows = []
early_alerts = []

for item in all_funds_for_output:
    row_dict = {
        "Fon Kodu": item["code"],
        "Fon Adı": item.get("fund_title") or "-",
        "Yatırım Alanı": item.get("investment_area") or "-",
    }
    own_scores = item.get("running_trend_hybrid") or []
    # last_5_s_asc: kronolojik (eski -> yeni) sıra; 2 gün teyit alarmı bunu kullanır.
    last_5_s_asc = own_scores[-5:] if len(own_scores) >= 5 else own_scores
    pad_len = len(last_5_dates_for_ui) - len(last_5_s_asc)
    last_5_padded_asc = [None] * pad_len + last_5_s_asc
    # last_5_dates_for_ui artık en güncel gün önce sıralı (create_excel_output içinde
    # ters çevrildi) -> ekran tablosu da aynı sırada gösterilsin.
    last_5_padded_desc = list(reversed(last_5_padded_asc))

    for day, score in zip(last_5_dates_for_ui, last_5_padded_desc):
        row_dict[f"{day} Karar Skoru"] = score if score is not None else ""
        row_dict[f"{day} Model Kararı"] = decision_label_from_score(score)

    row_dict.update({
        "Karar Skoru": item.get("decision_score"),
        "Güvenlik/Likidite": item.get("security_score"),
        "Trend Skoru": item.get("trend_skor"),
        "Model Kararı": item.get("karar"),
        "Net Likidite (%)": f"%{safe_float(item.get('emergency_cash_ratio')):.2f}",
        "Haftalık (%)": round(safe_float(item.get("weekly_return")), 2),
        "AUM ₺": round(safe_float(item.get("aum")), 0) if item.get("aum") is not None else None,
        "Yatırımcı": item.get("investors"),
    })
    display_rows.append(row_dict)

    # 2 Gün / 2 Mum KAZRİSK Teyit Algoritması (kronolojik sırayla hesaplanır)
    if len(last_5_s_asc) >= 2:
        lbls = [decision_label_from_score(s) for s in last_5_s_asc]
        if lbls[-1] == "ACİL SAT" and lbls[-2] == "ACİL SAT":
            early_alerts.append({
                "Fon Kodu": item["code"], "Fon Adı": item.get("fund_title"), "Alan": item.get("investment_area"),
                "KAZRİSK Durumu": "🚨 2 GÜNDÜR TEYİTLİ ACİL SAT", "Son Skor": last_5_s_asc[-1]
            })
        elif lbls[-1] == "GÜÇLÜ AL" and lbls[-2] == "GÜÇLÜ AL":
            early_alerts.append({
                "Fon Kodu": item["code"], "Fon Adı": item.get("fund_title"), "Alan": item.get("investment_area"),
                "KAZRİSK Durumu": "🚀 2 GÜNDÜR TEYİTLİ GÜÇLÜ AL", "Son Skor": last_5_s_asc[-1]
            })

df_display = pd.DataFrame(display_rows)

def color_cells(value):
    text = str(value)
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "🟢" in text: return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text or "🟡" in text: return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text or "YETERSİZ" in text or "🔴" in text: return "color: #FF0000; font-weight: bold;"
    return ""

try: styled_df = df_display.style.map(color_cells)
except: styled_df = df_display.style.applymap(color_cells)

st.subheader("📊 Analiz Sonuçları (V8.5)")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

if early_alerts:
    st.subheader("🚨 KAZRİSK® 2 Günlük Teyitli Alarmlar")
    st.dataframe(pd.DataFrame(early_alerts), use_container_width=True, hide_index=True)

if SHOW_DIAGNOSTICS:
    with st.expander("🔧 Kaynak Tanılama Bilgisi"):
        if failed_codes:
            st.warning(f"Veri alınamayan {len(failed_codes)} fon: {', '.join(failed_codes)}")
        diag_rows = [{
            "Fon Kodu": item["code"],
            "Veri Kaynağı": item.get("source", "-"),
            "Kaynak Zinciri": item.get("source_chain", "-"),
            "Veri Güveni": compute_confidence_label(item),
            "Gözlem Sayısı (gün)": item.get("n_days", 0),
            "Yapısal Veri Kaynağı": item.get("structural_source", "YOK"),
        } for item in all_funds_for_output]
        st.dataframe(pd.DataFrame(diag_rows), use_container_width=True, hide_index=True)

st.download_button(
    label="📥 Güncellenmiş KAZRİSK Excel'i İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_V8_5.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
