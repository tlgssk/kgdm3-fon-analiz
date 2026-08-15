import concurrent.futures
import datetime as dt
import io
import re
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
    "Momentum + Risk + Likidite Hibrit Skor Motoru V6.3"
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
MAX_WORKERS = 8  # paralel fon çekme iş parçacığı sayısı

APP_VERSION = "6.3.0"

# GitHub'da tarih damgalı dosya adları kullanıldığından, sabit bir dosya
# adına güvenmek yerine repodaki en güncel .xlsx dosyasını GitHub API
# üzerinden otomatik buluyoruz. Bu bulunamazsa aşağıdaki sabit URL'e
# (son bilinen dosya) düşüyoruz.
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

# Momentum tarafı toplam 100 puan
MOMENTUM_WEIGHTS = {
    "return": 0.30,
    "sharpe": 0.25,
    "cumulative": 0.25,
    "drawdown": 0.20,
}

# Güvenlik / likidite tarafı — TÜM katsayılar burada tek yerde toplanır.
# "*_scale" değerleri o bileşenin z-score'unun skor üzerindeki maksimum
# etkisini belirler (0..100 skalasında puan birimi).
SECURITY_WEIGHTS = {
    "aum": 0.30,
    "investor": 0.25,
    "concentration": 0.25,
    "liquidity": 0.20,
}

SECURITY_SCALE = {
    "aum": 20.0,
    "investor": 20.0,
    "aum_change": 12.0,
    "investor_change": 8.0,
    "concentration": 20.0,
}

# Nihai hibrit
HYBRID_MOMENTUM_WEIGHT = 0.60
HYBRID_SECURITY_WEIGHT = 0.40

# Z-score sınırı
Z_LIMIT = 2.5

# Karar seviyeleri
STRONG_BUY = 75
WATCH_LIST = 50
CORRECTION = 35

# Valör cezası (valör genelde 0-3 arası)
MAX_VALOR_PENALTY = 8.0

# Yapısal risk cezaları / bonuslar
MAX_CONCENTRATION_PENALTY = 20.0
BIST30_BONUS = 5.0
HIGH_LIQUIDITY_BONUS = 5.0
LOW_LIQUIDITY_PENALTY = 3.0
POSITIVE_INVESTOR_FLOW_BONUS = 3.0


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
    """
    Sayısal metinleri (TL / % işaretli, Türkçe/İngilizce ondalık ayraçlı)
    float'a çevirir. Belirsiz durumlarda (tek nokta, tek grup) ondalık
    nokta olarak yorumlanır; sadece "1.234" / "12.345.678" gibi saf
    binlik gruplama örüntüsünde (birden fazla 3'lü grup VEYA tek 3'lü
    grup sonrasında ondalık kısım yoksa) binlik ayraç olarak kabul edilir.
    """
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

    if not text:
        return None

    if "," in text and "." in text:
        # Hem nokta hem virgül varsa, en sağdaki ayraç ondalık kabul edilir.
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    elif "." in text and _THOUSANDS_ONLY_RE.match(text):
        # Örn. "1.234" veya "12.345.678" -> saf binlik gruplama
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


# ============================================================
# BİLEŞİK GETİRİ
# ============================================================

def calculate_compounded_return(returns) -> float:
    clean = []
    for value in returns:
        value = safe_float(value, None)
        if value is None:
            continue
        clean.append(value)

    if not clean:
        return 0.0

    growth = 1.0
    for daily_return in clean:
        growth *= 1.0 + daily_return / 100.0

    return (growth - 1.0) * 100.0


# ============================================================
# MAX DRAWDOWN
# ============================================================

def calculate_max_drawdown(prices) -> float:
    """Negatif döner. Örn. -7.5 = %7.5 drawdown"""
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


# ============================================================
# Z-SCORE
# ============================================================

def zscore(values) -> List[float]:
    """
    None girişler ortalama/std hesabına dahil edilmez; bu girişler için
    çıktı olarak 0.0 (nötr) döner. Çağıran taraf, None olan girdilerin
    "nötr" mü yoksa "gerçekten bilinmeyen" mi olduğuna karar vermeli —
    bilinmeyen değerleri zscore'a 0.0 olarak BESLEMEK yerine None
    geçmek, dağılımın ortalama/std'ini bozmadan nötr sonuç almanızı sağlar.
    """
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
        z = (value - mean_value) / std
        z = clamp(z, -Z_LIMIT, Z_LIMIT)
        result.append(z)

    return result


# ============================================================
# VALÖR NORMALİZASYONU
# ============================================================

def calculate_valor_penalty(valor) -> float:
    """
    Valör genelde 0, 1, 2, 3 şeklindedir.
    0 → ceza yok
    3 → maksimum ceza
    """
    valor = safe_float(valor)
    if valor <= 0:
        return 0.0

    normalized = clamp(valor / 3.0, 0.0, 1.0)
    return normalized * MAX_VALOR_PENALTY


# ============================================================
# GITHUB: EN GÜNCEL EXCEL DOSYASINI BUL
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def resolve_latest_github_excel_url() -> Optional[str]:
    """
    Repo kökündeki .xlsx dosyalarını GitHub Contents API üzerinden listeler
    ve dosya adına gömülü tarihe (varsa) ya da alfabetik sıraya göre en
    güncel olanı seçer. Böylece sabit kodlanmış tarihli bir dosya adı
    repoda güncellendiğinde link kırılmaz.
    """
    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/?ref={GITHUB_BRANCH}"
    )
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "kgdm3-fon-analiz-app",
    }

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
# 1. TEFAS UNIVERSE
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
    except ImportError:
        return pd.DataFrame()

    try:
        crawler = Crawler(timeout=60, max_retry=3)
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


def build_fund_kind_map(universe: pd.DataFrame) -> Dict[str, str]:
    """Fon kodu -> TEFAS fon tipi (YAT/EMK/BYF) eşlemesi üretir."""
    kind_map: Dict[str, str] = {}
    if universe is None or universe.empty or "kind" not in universe.columns:
        return kind_map
    try:
        latest = (
            universe.sort_values("date")
            .drop_duplicates(subset=["code"], keep="last")
        )
        for _, row in latest.iterrows():
            code = str(row.get("code", "")).strip().upper()
            kind = str(row.get("kind", "")).strip().upper()
            if code and kind in FUND_KINDS:
                kind_map[code] = kind
    except Exception:
        pass
    return kind_map


# ============================================================
# 2. İŞ YATIRIM
# ============================================================

def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code:
        return None

    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

    url = (
        "https://www.isyatirim.com.tr/"
        "_layouts/15/IsYatirim.Website/Common/"
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
# 3. TEFAS DIRECT API
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

    # Fon tipi bilinmiyorsa YAT/EMK/BYF sırasıyla denenir; bilinen tip
    # varsa önce o denenir, gerekirse diğerlerine düşülür.
    kind_candidates = [fund_kind] if fund_kind in FUND_KINDS else []
    kind_candidates += [k for k in FUND_KINDS if k not in kind_candidates]

    for kind in kind_candidates:
        payload = {
            "fontip": kind,
            "fonkod": code,
            "bastarih": start.strftime("%d.%m.%Y"),
            "bittarih": end.strftime("%d.%m.%Y"),
        }
        try:
            response = requests.post(url, data=payload, headers=headers, timeout=HTTP_TIMEOUT)
            if response.status_code != 200:
                continue

            data = response.json().get("data", [])
            if not data:
                continue

            df = pd.DataFrame(data)

            required = ["TARIH", "FIYAT"]
            if not all(column in df.columns for column in required):
                continue

            df["date"] = pd.to_datetime(df["TARIH"], unit="ms", errors="coerce")
            df["price"] = df["FIYAT"].apply(parse_number)

            if "PORTFOYBUYUKLUK" in df.columns:
                df["aum"] = df["PORTFOYBUYUKLUK"].apply(parse_number).fillna(0.0)
            else:
                df["aum"] = 0.0

            if "KISISAYISI" in df.columns:
                df["investors"] = df["KISISAYISI"].apply(parse_number).fillna(0.0)
            else:
                df["investors"] = 0.0

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
# 4. FİNTABLES / YAPISAL VERİ (CACHE'Lİ)
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 60 * 6)
def fetch_fund_structural_data(fund_code: str) -> dict:
    """
    NOT: Fintables sayfası istemci tarafında (JS ile) render ediliyor
    olabilir; bu durumda ham HTML üzerinden regex ile veri çekmek
    başarısız olur ve tüm alanlar "bilinmiyor" (None/False) döner.
    Bu fonksiyon best-effort'tur: veri bulunamazsa bunu açıkça
    `*_known=False` bayraklarıyla işaretler; çağıran taraf bu bayrakları
    kullanarak eksik veriyi "düşük risk" gibi yorumlamamalıdır.
    """
    code = normalize_fund_code(fund_code)

    structural = {
        "top_asset_weight": None,
        "is_bist30": False,
        "is_bist30_known": False,
        "emergency_cash_ratio": None,
        "cash_ratio_known": False,
        "structural_fetch_ok": False,
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
        match_top = re.search(
            r'En Büyük Pay["\s:]+([0-9]+(?:[.,][0-9]+)?)',
            text,
            re.IGNORECASE,
        )
        if match_top:
            structural["top_asset_weight"] = parse_number(match_top.group(1))

        # BIST30 – sadece açıkça geçiyorsa
        if re.search(r"BIST\s*30", text, re.IGNORECASE):
            structural["is_bist30"] = True
            structural["is_bist30_known"] = True

        # Nakit / Ters Repo / PPF
        match_cash = re.search(
            r'(?:Nakit|Ters Repo|PPF)["\s:]+([0-9]+(?:[.,][0-9]+)?)',
            text,
            re.IGNORECASE,
        )
        if match_cash:
            cash_value = parse_number(match_cash.group(1))
            if cash_value is not None:
                structural["emergency_cash_ratio"] = cash_value
                structural["cash_ratio_known"] = True

        structural["structural_fetch_ok"] = any(
            [
                structural["top_asset_weight"] is not None,
                structural["is_bist30_known"],
                structural["cash_ratio_known"],
            ]
        )

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

    # 1. TEFAS Universe
    if universe is not None and not universe.empty and "code" in universe.columns:
        rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
        if not rows.empty:
            rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
            if len(rows) >= 2:
                return rows.tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True), "TEFAS"

    # 2. TEFAS Direct API
    direct_df = fetch_tefas_direct_api(code, fund_kind)
    if direct_df is not None and len(direct_df) >= 2:
        return direct_df, "TEFAS Direct API"

    # 3. İş Yatırım
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
        if previous <= 0:
            daily_returns.append(0.0)
        else:
            daily_returns.append((current / previous - 1.0) * 100.0)

    if not daily_returns:
        return None

    max_dd = calculate_max_drawdown(prices)

    aum_change = ((aums[-1] / aums[0] - 1.0) * 100.0) if aums[0] > 0 else 0.0
    investor_change = ((investors[-1] / investors[0] - 1.0) * 100.0) if investors[0] > 0 else 0.0

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
        "inv_change": investor_change,
        "max_dd": max_dd,
        "weekly_return": weekly_return,
        **structural,
    }


def fetch_and_compute_one_fund(
    code: str,
    universe: pd.DataFrame,
    kind_map: Dict[str, str],
    valor_dict: Dict[str, float],
) -> Tuple[str, Optional[dict], str]:
    """Tek bir fon için seri çekme + metrik hesaplama — paralel çalıştırılabilir."""
    fund_kind = kind_map.get(code)
    series, source = get_fund_series(universe, code, fund_kind)
    metrics = compute_fund_metrics(series, code)

    if metrics is None:
        return code, None, source

    metrics["code"] = code
    metrics["valor"] = valor_dict.get(code, 0.0)
    metrics["source"] = source
    return code, metrics, source


# ============================================================
# ROLLING METRİKLER
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
# GÜVENLİK / LİKİDİTE SKORU
# ============================================================

def calculate_security_scores(funds: List[dict]) -> None:
    aum_values = [safe_float(x.get("aum")) for x in funds]
    investor_values = [safe_float(x.get("investors")) for x in funds]
    aum_change_values = [safe_float(x.get("aum_change")) for x in funds]
    investor_change_values = [safe_float(x.get("inv_change")) for x in funds]

    aum_z = zscore(aum_values)
    investor_z = zscore(investor_values)
    aum_change_z = zscore(aum_change_values)
    investor_change_z = zscore(investor_change_values)

    # Konsantrasyon (yüksek = kötü). Bilinmeyen (None) değerler zscore
    # dağılımına dahil EDİLMEZ — böylece veri eksikliği, dağılımın
    # ortalama/std'sini bozarak dolaylı "düşük risk" avantajı yaratmaz.
    concentration_raw = [fund.get("top_asset_weight") for fund in funds]
    concentration_z = zscore(concentration_raw)

    for i, fund in enumerate(funds):
        score = 50.0

        # AUM & Yatırımcı (mutlak seviye)
        score += SECURITY_SCALE["aum"] * SECURITY_WEIGHTS["aum"] * aum_z[i]
        score += SECURITY_SCALE["investor"] * SECURITY_WEIGHTS["investor"] * investor_z[i]

        # Değişimler (trend)
        score += SECURITY_SCALE["aum_change"] * aum_change_z[i]
        score += SECURITY_SCALE["investor_change"] * investor_change_z[i]

        # Konsantrasyon (yüksek z = yüksek risk → negatif). Sadece veri
        # bilinen fonlar için uygulanır.
        if fund.get("top_asset_weight") is not None:
            score -= SECURITY_SCALE["concentration"] * SECURITY_WEIGHTS["concentration"] * concentration_z[i]

        # BIST30 bonus
        if fund.get("is_bist30", False):
            score += BIST30_BONUS

        # Nakit / likidite — sadece veri bilinen fonlar için
        cash_ratio = fund.get("emergency_cash_ratio")
        if fund.get("cash_ratio_known", False) and cash_ratio is not None:
            if cash_ratio >= 15:
                score += HIGH_LIQUIDITY_BONUS
            elif cash_ratio < 5:
                score -= LOW_LIQUIDITY_PENALTY

        # Pozitif yatırımcı akışı
        if safe_float(fund.get("inv_change")) > 0:
            score += POSITIVE_INVESTOR_FLOW_BONUS

        # Doğrudan konsantrasyon cezası (sadece bilinen değerler için)
        top_asset = fund.get("top_asset_weight")
        if top_asset is not None:
            if top_asset > 30:
                penalty = min((top_asset - 30) * 1.0, MAX_CONCENTRATION_PENALTY)
                score -= penalty
            elif top_asset > 15:
                score -= (top_asset - 15) * 0.25

        # Valör cezası
        score -= calculate_valor_penalty(fund.get("valor"))

        fund["security_score"] = int(round(clamp(score, 0.0, 100.0)))


# ============================================================
# ANA HİBRİT SKOR MOTORU
# ============================================================

def calculate_hybrid_scores(funds: List[dict]) -> Tuple[int, List[dict], List[dict]]:
    """
    Hibrit skoru hesaplar. Yetersiz geçmişe sahip (n_days < MIN_ROLLING_DAYS)
    fonlar, ortak pencereyi kısaltmasın diye momentum hesaplamasından
    HARİÇ tutulur ve "YETERSİZ VERİ" olarak işaretlenir; yine de güvenlik
    skoru hesaplanır ve çıktı listesinde gösterilir.

    Döndürür: (n_days, hesaplanan_fonlar, yetersiz_veri_fonları)
    """
    if not funds:
        return 0, [], []

    eligible = [f for f in funds if f.get("n_days", 0) >= MIN_ROLLING_DAYS]
    insufficient = [f for f in funds if f.get("n_days", 0) < MIN_ROLLING_DAYS]

    for fund in insufficient:
        fund["security_score"] = None
        fund["kgdm_skor"] = None
        fund["karar"] = "YETERSİZ VERİ"
        fund["running_scores"] = []
        fund["running_momentum"] = []
        fund["last_5_scores"] = []
        fund["last_5_scores_str"] = "-"

    if not eligible:
        return 0, [], insufficient

    n_days = min(x["n_days"] for x in eligible)

    for fund in eligible:
        fund["dates"] = fund["dates"][-n_days:]
        fund["daily_returns"] = fund["daily_returns"][-n_days:]
        fund["prices"] = fund["prices"][-(n_days + 1):]
        fund["running_scores"] = []
        fund["running_momentum"] = []

    # Güvenlik skoru (yapısal, bir kez) — sadece uygun (eligible) fonlar
    # üzerinde hesaplanır ki dağılım yetersiz veri fonlarınca bozulmasın.
    calculate_security_scores(eligible)

    for d in range(1, n_days + 1):
        if d < MIN_ROLLING_DAYS:
            for fund in eligible:
                fund["running_momentum"].append(None)
            continue

        current_metrics = []

        for fund in eligible:
            returns_slice = fund["daily_returns"][d - MIN_ROLLING_DAYS: d]
            prices_slice = fund["prices"][d - MIN_ROLLING_DAYS: d + 1]

            if len(returns_slice) < MIN_ROLLING_DAYS:
                continue

            mean_return = sum(returns_slice) / len(returns_slice)
            variance = sum((r - mean_return) ** 2 for r in returns_slice) / len(returns_slice)
            volatility = variance ** 0.5
            sharpe = (mean_return / volatility) if volatility > 1e-12 else 0.0
            cumulative = calculate_compounded_return(returns_slice)
            max_dd = calculate_max_drawdown(prices_slice)

            current_metrics.append({
                "fund": fund,
                "mean_return": mean_return,
                "volatility": volatility,
                "sharpe": sharpe,
                "cumulative": cumulative,
                "max_dd": max_dd,
            })

        if not current_metrics:
            continue

        mean_z = zscore([x["mean_return"] for x in current_metrics])
        sharpe_z = zscore([x["sharpe"] for x in current_metrics])
        cumulative_z = zscore([x["cumulative"] for x in current_metrics])
        drawdown_z = zscore([-x["max_dd"] for x in current_metrics])

        for i, data in enumerate(current_metrics):
            weighted_z = (
                MOMENTUM_WEIGHTS["return"] * mean_z[i]
                + MOMENTUM_WEIGHTS["sharpe"] * sharpe_z[i]
                + MOMENTUM_WEIGHTS["cumulative"] * cumulative_z[i]
                + MOMENTUM_WEIGHTS["drawdown"] * drawdown_z[i]
            )

            momentum_score = clamp(50.0 + 20.0 * weighted_z, 0.0, 100.0)
            data["fund"]["running_momentum"].append(int(round(momentum_score)))

    for fund in eligible:
        security_score = safe_float(fund.get("security_score"), 50.0)

        running_hybrid = []
        for momentum in fund["running_momentum"]:
            if momentum is None:
                running_hybrid.append(None)
                continue
            hybrid = (
                momentum * HYBRID_MOMENTUM_WEIGHT
                + security_score * HYBRID_SECURITY_WEIGHT
            )
            running_hybrid.append(int(round(clamp(hybrid, 0.0, 100.0))))

        fund["running_scores"] = running_hybrid

        valid_last = [s for s in running_hybrid if s is not None][-5:]
        fund["last_5_scores"] = valid_last
        fund["last_5_scores_str"] = " ➔ ".join(str(x) for x in valid_last) if valid_last else "-"

        if valid_last:
            weights = list(range(1, len(valid_last) + 1))
            weighted_average = sum(s * w for s, w in zip(valid_last, weights)) / sum(weights)
            fund["kgdm_skor"] = int(round(weighted_average))
        else:
            fund["kgdm_skor"] = None

        score = fund.get("kgdm_skor")
        if score is None:
            fund["karar"] = "YETERSİZ VERİ"
        elif score >= STRONG_BUY:
            fund["karar"] = "GÜÇLÜ AL"
        elif score >= WATCH_LIST:
            fund["karar"] = "ASIL LİSTE"
        elif score >= CORRECTION:
            fund["karar"] = "DÜZELTME / İZLE"
        else:
            fund["karar"] = "ACİL SAT"

    return n_days, eligible, insufficient


# ============================================================
# EXCEL STİL
# ============================================================

def style_excel_sheet(ws):
    thin_gray = Side(style="thin", color="D9E1F2")

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin_gray)

    ws.freeze_panes = "A2"
    if ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False


def auto_fit_columns(ws, min_width=10, max_width=45):
    for column_cells in ws.columns:
        column_index = column_cells[0].column
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        width = max(min_width, min(max_length + 3, max_width))
        ws.column_dimensions[get_column_letter(column_index)].width = width


# ============================================================
# EXCEL ÇIKTISI
# ============================================================

PERCENT_COLUMNS = [
    "Ort. Günlük Getiri (%)", "Volatilite (%)",
    "Kümülatif Getiri (%)", "MaxDD (%)",
    "AUM Değişim (%)", "Yatırımcı Değişim (%)",
    "Haftalık Bileşik Getiri (%)",
]


def create_excel_output(wb, ws_list, calculated_funds, all_funds_for_output, n_days):
    if "KGDM3_Puanlama" in wb.sheetnames:
        del wb["KGDM3_Puanlama"]

    ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

    headers = [
        "Fon Kodu", "Valör",
        "Hibrit Skor", "Momentum Skor", "Güvenlik Skor",
        "Son 5 Skor", "Model Kararı",
        "Ort. Günlük Getiri (%)", "Volatilite (%)", "Sharpe",
        "Kümülatif Getiri (%)", "MaxDD (%)",
        "En Büyük Varlık (%)", "BIST30", "Nakit Verisi", "KAZRİSK",
        "AUM Değişim (%)", "Yatırımcı Değişim (%)",
        "AUM (₺)", "Yatırımcı",
        "Haftalık Bileşik Getiri (%)", "Veri Kaynağı",
    ]

    # Tarih bazlı kolonlar (en uzun geçmişe sahip fonun tarihleri baz alınır)
    sample_dates = []
    for item in calculated_funds:
        if item.get("dates"):
            sample_dates = item["dates"]
            break

    for day in sample_dates:
        headers.append(f"{day} Hibrit Skor")
    for day in sample_dates:
        headers.append(f"{day} Getiri")

    ws_scores.append(headers)

    # Sütun adı -> 1-tabanlı Excel kolon indeksi (hardcode yerine dinamik)
    header_index = {name: idx + 1 for idx, name in enumerate(headers)}

    header_fill = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color=COLOR_WHITE)

    for cell in ws_scores[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_scores.row_dimensions[1].height = 42

    for item in all_funds_for_output:
        top_asset = item.get("top_asset_weight")

        if top_asset is None:
            risk_label = "⚪ Veri Yok"
        elif top_asset > 30:
            risk_label = "⚠️ Yüksek Konsantrasyon"
        elif top_asset > 15:
            risk_label = "🟡 Orta Konsantrasyon"
        else:
            risk_label = "🛡️ Dengeli"

        cash_known = item.get("cash_ratio_known", False)
        cash_label = (
            f"%{safe_float(item.get('emergency_cash_ratio')):.2f}"
            if cash_known else "Veri Yok"
        )

        last_momentum = None
        if item.get("running_momentum"):
            valid_m = [m for m in item["running_momentum"] if m is not None]
            last_momentum = valid_m[-1] if valid_m else None

        row_data = [
            item["code"],
            item.get("valor", 0),
            item.get("kgdm_skor"),
            last_momentum,
            item.get("security_score"),
            item.get("last_5_scores_str", "-"),
            item.get("karar", "-"),
            round(safe_float(item.get("mean_return")), 4),
            round(safe_float(item.get("volatility")), 4),
            round(safe_float(item.get("sharpe_like")), 4),
            round(safe_float(item.get("cumulative_return")), 4),
            round(safe_float(item.get("max_dd")), 4),
            round(safe_float(top_asset), 2) if top_asset is not None else None,
            "EVET" if item.get("is_bist30", False) else "HAYIR / YOK",
            cash_label,
            risk_label,
            round(safe_float(item.get("aum_change")), 2),
            round(safe_float(item.get("inv_change")), 2),
            round(safe_float(item.get("aum")), 2) if item.get("aum") is not None else None,
            int(item.get("investors")) if item.get("investors") is not None else None,
            round(safe_float(item.get("weekly_return")), 4),
            item.get("source", "-"),
        ]

        # Hibrit skorlar — bu fonun kendi tarih dizisine göre; uzunluk
        # sample_dates ile farklıysa sağa hizalı doldurulur.
        own_scores = item.get("running_scores", [])
        padded_scores = [None] * (len(sample_dates) - len(own_scores)) + own_scores
        row_data.extend([s if s is not None else "" for s in padded_scores[-len(sample_dates):]] if sample_dates else [])

        own_returns = item.get("daily_returns", [])
        padded_returns = [None] * (len(sample_dates) - len(own_returns)) + own_returns
        row_data.extend(
            [format_percent(x) if x is not None else "-" for x in padded_returns[-len(sample_dates):]]
            if sample_dates else []
        )

        ws_scores.append(row_data)

    green_font = Font(bold=True, color=COLOR_GREEN)
    red_font = Font(bold=True, color=COLOR_RED)
    yellow_font = Font(bold=True, color=COLOR_YELLOW)

    decision_col = header_index["Model Kararı"]
    for row_number in range(2, ws_scores.max_row + 1):
        decision_cell = ws_scores.cell(row=row_number, column=decision_col)
        decision_text = str(decision_cell.value or "")

        if "GÜÇLÜ AL" in decision_text or "ASIL LİSTE" in decision_text:
            decision_cell.font = green_font
        elif "DÜZELTME" in decision_text:
            decision_cell.font = yellow_font
        elif "ACİL SAT" in decision_text:
            decision_cell.font = red_font

    hybrid_col_letter = get_column_letter(header_index["Hibrit Skor"])
    score_range = f"{hybrid_col_letter}2:{hybrid_col_letter}{ws_scores.max_row}"
    ws_scores.conditional_formatting.add(
        score_range,
        CellIsRule(
            operator="greaterThanOrEqual",
            formula=["75"],
            fill=PatternFill(start_color=COLOR_LIGHT_GREEN, fill_type="solid"),
        ),
    )
    ws_scores.conditional_formatting.add(
        score_range,
        CellIsRule(
            operator="between",
            formula=["50", "74"],
            fill=PatternFill(start_color=COLOR_LIGHT_YELLOW, fill_type="solid"),
        ),
    )
    ws_scores.conditional_formatting.add(
        score_range,
        CellIsRule(
            operator="lessThan",
            formula=["50"],
            fill=PatternFill(start_color=COLOR_LIGHT_RED, fill_type="solid"),
        ),
    )

    # Sayı formatları — sütun adına göre dinamik olarak uygulanır.
    currency_col = header_index.get("AUM (₺)")
    integer_col = header_index.get("Yatırımcı")

    for row_number in range(2, ws_scores.max_row + 1):
        if currency_col:
            ws_scores.cell(row=row_number, column=currency_col).number_format = '#,##0.00 "₺"'
        if integer_col:
            ws_scores.cell(row=row_number, column=integer_col).number_format = "#,##0"

        for col_name in PERCENT_COLUMNS:
            col_idx = header_index.get(col_name)
            if col_idx:
                ws_scores.cell(row=row_number, column=col_idx).number_format = '0.00"%"'

    style_excel_sheet(ws_scores)
    auto_fit_columns(ws_scores)
    style_excel_sheet(ws_list)
    auto_fit_columns(ws_list)

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
source_mode = None

with col_upload:
    uploaded_file = st.file_uploader("Bilgisayardan Excel Yükle", type=["xlsx"])
    if uploaded_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_file)
            source_mode = "upload"
        except Exception as exc:
            st.error(f"Excel yükleme hatası: {exc}")

with col_github:
    st.write("Veya GitHub'daki listeyi kullanın:")
    if st.button("🚀 GitHub'dan Çek ve Analiz Et", use_container_width=True):
        resolved_url = resolve_latest_github_excel_url()
        target_url = resolved_url or GITHUB_FALLBACK_URL
        try:
            response = requests.get(target_url, timeout=HTTP_TIMEOUT)
            response.raise_for_status()
            wb = openpyxl.load_workbook(io.BytesIO(response.content))
            source_mode = "github"
            if resolved_url:
                st.success(f"✅ En güncel Excel dosyası indirildi: {resolved_url.split('/')[-1]}")
            else:
                st.warning(
                    "⚠️ GitHub API üzerinden en güncel dosya bulunamadı, "
                    "sabit yedek URL kullanıldı. Dosya güncel olmayabilir."
                )
        except Exception as exc:
            st.error(f"GitHub bağlantı hatası: {exc}")

if wb is None:
    st.info("Analize başlamak için Excel dosyanızı yükleyin.")
    st.stop()


# ============================================================
# FON LİSTESİ
# ============================================================

if "Fon_Listesi" in wb.sheetnames:
    ws_list = wb["Fon_Listesi"]
else:
    ws_list = wb.active

requested_codes = []
excel_valor_dict = {}

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if not row or row[0].value is None:
        continue

    code = normalize_fund_code(row[0].value)
    if not code:
        continue

    requested_codes.append(code)

    try:
        if len(row) > 3:
            valor = parse_number(row[3].value)
            excel_valor_dict[code] = valor if valor is not None else 0.0
        else:
            excel_valor_dict[code] = 0.0
    except Exception:
        excel_valor_dict[code] = 0.0

requested_codes = list(dict.fromkeys(requested_codes))

if not requested_codes:
    st.error("Fon_Listesi sayfasında fon kodu bulunamadı.")
    st.stop()


# ============================================================
# TARİH & TEFAS UNIVERSE
# ============================================================

today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

with st.spinner("🔄 TEFAS verileri alınıyor..."):
    universe = fetch_tefas_universe(start_date, today)
    fund_kind_map = build_fund_kind_map(universe)


# ============================================================
# FONLARIN HESAPLANMASI (PARALEL)
# ============================================================

calculated_funds = []
failed_codes = []
structural_fetch_failures = 0

progress = st.progress(0, text="Fonlar analiz ediliyor...")
total_funds = len(requested_codes)
completed = 0

with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    future_to_code = {
        executor.submit(
            fetch_and_compute_one_fund, code, universe, fund_kind_map, excel_valor_dict
        ): code
        for code in requested_codes
    }

    for future in concurrent.futures.as_completed(future_to_code):
        code = future_to_code[future]
        completed += 1
        try:
            _, metrics, source = future.result()
        except Exception:
            metrics, source = None, "YOK"

        if metrics is None:
            failed_codes.append(code)
            progress.progress(completed / total_funds, text=f"{code}: veri yok")
            continue

        if not metrics.get("structural_fetch_ok", False):
            structural_fetch_failures += 1

        if ENABLE_FILTERS:
            investor_ok = metrics["investors"] >= MIN_INVESTOR_COUNT
            weekly_ok = metrics["weekly_return"] >= TARGET_WEEKLY_RETURN
            if not (investor_ok and weekly_ok):
                progress.progress(completed / total_funds, text=f"{code}: filtre dışı")
                continue

        calculated_funds.append(metrics)
        progress.progress(completed / total_funds, text=f"{code}: analiz edildi")

progress.empty()

if failed_codes:
    st.warning("Veri bulunamayan fonlar: " + ", ".join(sorted(failed_codes)))

if SHOW_DIAGNOSTICS and structural_fetch_failures > 0:
    st.warning(
        f"⚠️ {structural_fetch_failures} fon için yapısal veri "
        f"(en büyük varlık / BIST30 / nakit oranı) alınamadı. "
        "Bu fonlar güvenlik skorunda nötr (bilinmeyen) olarak işaretlendi; "
        "sonucu bu fonlar lehine ya da aleyhine sistematik olarak "
        "çarpıtmaz, sadece o bileşenler için 'Veri Yok' gösterilir."
    )

if not calculated_funds:
    st.error("Hesaplanabilecek geçerli fon verisi bulunamadı.")
    st.stop()


# ============================================================
# ORTAK GÜN SAYISI & HİBRİT SKOR
# ============================================================

with st.spinner("📊 KGDM-3 + KAZRİSK hibrit skoru hesaplanıyor..."):
    n_days, eligible_funds, insufficient_funds = calculate_hybrid_scores(calculated_funds)

if insufficient_funds:
    st.info(
        f"ℹ️ {len(insufficient_funds)} fon, en az {MIN_ROLLING_DAYS} işlem günü "
        "geçmişine sahip olmadığı için momentum/hibrit skor hesabına dahil "
        "edilmedi (\"YETERSİZ VERİ\" olarak işaretlendi): "
        + ", ".join(sorted(f["code"] for f in insufficient_funds))
    )

if not eligible_funds:
    st.error(f"Fonlarda yeterli tarihsel veri bulunmuyor (en az {MIN_ROLLING_DAYS} gün gerekli).")
    st.stop()

all_funds_for_output = eligible_funds + insufficient_funds


# ============================================================
# SON DÖNEM METRİKLERİ
# ============================================================

for item in all_funds_for_output:
    metrics = calculate_window_metrics(
        item["prices"],
        item["daily_returns"],
        min(MIN_ROLLING_DAYS, item["n_days"]),
    )

    if metrics is None:
        item["mean_return"] = 0.0
        item["volatility"] = 0.0
        item["sharpe_like"] = 0.0
        item["cumulative_return"] = 0.0
    else:
        item["mean_return"] = metrics["mean_return"]
        item["volatility"] = metrics["volatility"]
        item["sharpe_like"] = metrics["sharpe"]
        item["cumulative_return"] = metrics["cumulative"]


# ============================================================
# SIRALAMA
# ============================================================

all_funds_for_output.sort(
    key=lambda x: (
        -safe_float(x.get("kgdm_skor")),
        -safe_float(x.get("cumulative_return")),
    )
)


# ============================================================
# SONUÇ TABLOSU
# ============================================================

display_rows = []

for item in all_funds_for_output:
    top_asset = item.get("top_asset_weight")

    if top_asset is None:
        risk_status = "⚪ Veri Yok"
    elif top_asset > 30:
        risk_status = "⚠️ Yüksek Konsantrasyon"
    elif top_asset > 15:
        risk_status = "🟡 Orta Konsantrasyon"
    else:
        risk_status = "🛡️ Dengeli"

    last_momentum = None
    if item.get("running_momentum"):
        valid_m = [m for m in item["running_momentum"] if m is not None]
        last_momentum = valid_m[-1] if valid_m else None

    display_rows.append({
        "Fon Kodu": item["code"],
        "Hibrit Skor": item.get("kgdm_skor"),
        "Momentum": last_momentum,
        "Güvenlik": item.get("security_score"),
        "Model Kararı": item.get("karar"),
        "Ort. Günlük %": round(safe_float(item.get("mean_return")), 3),
        "Sharpe": round(safe_float(item.get("sharpe_like")), 3),
        "Kümülatif Getiri %": round(safe_float(item.get("cumulative_return")), 3),
        "MaxDD %": round(safe_float(item.get("max_dd")), 3),
        "En Büyük Varlık %": round(safe_float(top_asset), 2) if top_asset is not None else None,
        "KAZRİSK": risk_status,
        "Haftalık Bileşik %": round(safe_float(item.get("weekly_return")), 3),
        "AUM ₺": round(safe_float(item.get("aum")), 0) if item.get("aum") is not None else None,
        "Yatırımcı": item.get("investors"),
        "Kaynak": item.get("source"),
    })

df_display = pd.DataFrame(display_rows)


def color_cells(value):
    text = str(value)
    if "GÜÇLÜ AL" in text or "ASIL LİSTE" in text or "Dengeli" in text:
        return "color: #008000; font-weight: bold;"
    if "DÜZELTME" in text or "Orta" in text:
        return "color: #B8860B; font-weight: bold;"
    if "ACİL SAT" in text or "Yüksek" in text or "YETERSİZ" in text:
        return "color: #FF0000; font-weight: bold;"
    return ""


try:
    styled_df = df_display.style.map(color_cells)
except AttributeError:
    styled_df = df_display.style.applymap(color_cells)


# ============================================================
# EKRAN
# ============================================================

st.subheader("📊 Analiz Sonuçları")
st.dataframe(styled_df, use_container_width=True, hide_index=True)

st.subheader("📈 Skor Özeti")

col1, col2, col3, col4 = st.columns(4)

scores = [
    safe_float(x.get("kgdm_skor"))
    for x in all_funds_for_output
    if x.get("kgdm_skor") is not None
]

if scores:
    with col1:
        st.metric("En Yüksek Skor", f"{max(scores):.0f}")
    with col2:
        st.metric("Ortalama Skor", f"{sum(scores) / len(scores):.1f}")
    with col3:
        st.metric("En Düşük Skor", f"{min(scores):.0f}")
    with col4:
        strong_count = sum(1 for x in all_funds_for_output if x.get("karar") == "GÜÇLÜ AL")
        st.metric("Güçlü Al", strong_count)


# ============================================================
# EXCEL ÇIKTISI
# ============================================================

output = create_excel_output(
    wb=wb,
    ws_list=ws_list,
    calculated_funds=eligible_funds,
    all_funds_for_output=all_funds_for_output,
    n_days=n_days,
)

st.success(
    f"✅ Analiz tamamlandı. "
    f"{len(all_funds_for_output)} fon işlendi "
    f"({len(eligible_funds)} skorlandı, {len(insufficient_funds)} yetersiz veri). "
    f"Model sürümü: {APP_VERSION}"
)

st.download_button(
    label="📥 Güncellenmiş Hibrit Excel'i İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_V6_3.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
