import datetime as dt
import io
import re
from typing import Optional

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
    "Momentum + Risk + Likidite Hibrit Skor Motoru V6.1"
)


# ============================================================
# GENEL AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
MIN_ROLLING_DAYS = 5
HTTP_TIMEOUT = 12

APP_VERSION = "6.1.0"

GITHUB_EXCEL_URL = (
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

# Güvenlik / likidite tarafı toplam 100 puan
SECURITY_WEIGHTS = {
    "aum": 0.30,
    "investor": 0.25,
    "concentration": 0.25,
    "liquidity": 0.20,
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

# Valör cezası
# Artık doğrudan "valor * 0.5" kullanılmıyor.
# Valör 0-100 aralığında normalize edilerek maksimum sınırlı ceza uygulanıyor.
MAX_VALOR_PENALTY = 8.0

# Yapısal risk cezaları
MAX_CONCENTRATION_PENALTY = 20.0
BIST30_BONUS = 5.0
HIGH_LIQUIDITY_BONUS = 5.0
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

    if not text:
        return None

    if "," in text and "." in text:

        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")

    elif "," in text:

        text = text.replace(",", ".")

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

def calculate_compounded_return(returns):

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

def calculate_max_drawdown(prices):

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

    # Negatif olarak döndürülür.
    # Örn. -7.5 = %7.5 drawdown
    return max_dd


# ============================================================
# Z-SCORE
# ============================================================

def zscore(values):

    if not values:
        return []

    clean = []

    for value in values:

        try:

            value = float(value)

            if pd.isna(value):
                clean.append(None)
            else:
                clean.append(value)

        except Exception:

            clean.append(None)

    valid = [x for x in clean if x is not None]

    if len(valid) < 2:
        return [0.0 for _ in clean]

    mean_value = sum(valid) / len(valid)

    variance = sum(
        (x - mean_value) ** 2
        for x in valid
    ) / len(valid)

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

def calculate_valor_penalty(valor):

    """
    Valör artık ham TL veya nominal sayı olarak skoru ezmez.

    Varsayım:
    - 0 veya negatif -> ceza yok
    - 1-100 arası -> doğrusal normalize
    - 100 üzeri -> maksimum ceza
    """

    valor = safe_float(valor)

    if valor <= 0:
        return 0.0

    normalized = clamp(valor / 100.0, 0.0, 1.0)

    return normalized * MAX_VALOR_PENALTY


# ============================================================
# 1. TEFAS UNIVERSE
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
        }

        df.rename(
            columns=rename_map,
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
            df
            .sort_values(["code", "date"])
            .drop_duplicates(
                subset=["code", "date"],
                keep="last",
            )
            .reset_index(drop=True)
        )

    except Exception:

        return pd.DataFrame()


# ============================================================
# 2. İŞ YATIRIM
# ============================================================

def fetch_isyatirim_series(
    fund_code: str,
) -> Optional[pd.DataFrame]:

    code = normalize_fund_code(fund_code)

    if not code:
        return None

    end = dt.datetime.now()

    start = (
        end -
        dt.timedelta(
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
        "baslangic": start.strftime("%d-%m-%Y"),
        "bitis": end.strftime("%d-%m-%Y"),
    }

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:

        response = requests.get(
            url,
            params=params,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        if response.status_code != 200:
            return None

        payload = response.json()

        values = payload.get("value")

        if not values:
            return None

        df = pd.DataFrame(values)

        if (
            "Tarih" not in df.columns
            or "Fiyat" not in df.columns
        ):
            return None

        df["date"] = pd.to_datetime(
            df["Tarih"],
            dayfirst=True,
            errors="coerce",
        )

        df["price"] = (
            df["Fiyat"]
            .apply(parse_number)
        )

        df["aum"] = 0.0
        df["investors"] = 0.0

        df = df.dropna(
            subset=["date", "price"]
        )

        df = df[df["price"] > 0]

        if len(df) < 2:
            return None

        df = (
            df
            .sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .tail(TARGET_TRADING_DAYS + 1)
            .reset_index(drop=True)
        )

        return df[
            [
                "date",
                "price",
                "aum",
                "investors",
            ]
        ]

    except Exception:

        return None


# ============================================================
# 3. TEFAS DIRECT API
# ============================================================

def fetch_tefas_direct_api(
    fund_code: str,
) -> Optional[pd.DataFrame]:

    code = normalize_fund_code(fund_code)

    if not code:
        return None

    end = dt.datetime.now()

    start = (
        end -
        dt.timedelta(
            days=LOOKBACK_CALENDAR_DAYS
        )
    )

    url = (
        "https://www.tefas.gov.tr/"
        "api/DB/BindHistoryInfo"
    )

    payload = {
        "fontip": "YAT",
        "fonkod": code,
        "bastarih": start.strftime("%d.%m.%Y"),
        "bittarih": end.strftime("%d.%m.%Y"),
    }

    headers = {
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.tefas.gov.tr",
    }

    try:

        response = requests.post(
            url,
            data=payload,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        if response.status_code != 200:
            return None

        data = response.json().get(
            "data",
            [],
        )

        if not data:
            return None

        df = pd.DataFrame(data)

        required = [
            "TARIH",
            "FIYAT",
        ]

        if not all(
            column in df.columns
            for column in required
        ):
            return None

        df["date"] = pd.to_datetime(
            df["TARIH"],
            unit="ms",
            errors="coerce",
        )

        df["price"] = (
            df["FIYAT"]
            .apply(parse_number)
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
            return None

        return (
            df
            .sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .tail(TARGET_TRADING_DAYS + 1)
            .reset_index(drop=True)
        )

    except Exception:

        return None


# ============================================================
# 4. FİNTABLES / YAPISAL VERİ
# ============================================================

def fetch_fund_structural_data(
    fund_code: str,
) -> dict:

    code = normalize_fund_code(fund_code)

    structural = {
        "top_asset_weight": None,
        "is_bist30": False,
        "is_bist30_known": False,
        "emergency_cash_ratio": None,
        "cash_ratio_known": False,
    }

    try:

        fintables_url = (
            f"https://fintables.com/fonlar/"
            f"{code.lower()}"
        )

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        response = requests.get(
            fintables_url,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        if response.status_code != 200:
            return structural

        text = response.text

        # ----------------------------------------------------
        # En büyük varlık
        # ----------------------------------------------------

        match_top = re.search(
            r'En Büyük Pay["\s:]+'
            r'([0-9]+(?:[.,][0-9]+)?)',
            text,
            re.IGNORECASE,
        )

        if match_top:

            structural["top_asset_weight"] = (
                parse_number(
                    match_top.group(1)
                )
            )

        # ----------------------------------------------------
        # BIST30
        # ----------------------------------------------------

        # Önceki hatalı mantık:
        # "BIST 100" veya "Ters Repo" görülünce BIST30=True
        #
        # Yeni mantık:
        # Sadece BIST 30 açıkça bulunuyorsa True.
        # ----------------------------------------------------

        bist30_match = re.search(
            r"BIST\s*30",
            text,
            re.IGNORECASE,
        )

        if bist30_match:

            structural["is_bist30"] = True
            structural["is_bist30_known"] = True

        # ----------------------------------------------------
        # Nakit / Ters Repo / PPF
        # ----------------------------------------------------

        match_cash = re.search(
            r'(?:Nakit|Ters Repo|PPF)'
            r'["\s:]+'
            r'([0-9]+(?:[.,][0-9]+)?)',
            text,
            re.IGNORECASE,
        )

        if match_cash:

            cash_value = parse_number(
                match_cash.group(1)
            )

            if cash_value is not None:

                structural[
                    "emergency_cash_ratio"
                ] = cash_value

                structural[
                    "cash_ratio_known"
                ] = True

    except Exception:

        pass

    return structural


# ============================================================
# FON SERİSİ
# ============================================================

def get_fund_series(
    universe: pd.DataFrame,
    fund_code: str,
):

    code = normalize_fund_code(fund_code)

    if not code:
        return None, "YOK"

    # --------------------------------------------------------
    # 1. TEFAS Universe
    # --------------------------------------------------------

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

                return (
                    rows
                    .tail(TARGET_TRADING_DAYS + 1)
                    .reset_index(drop=True),
                    "TEFAS",
                )

    # --------------------------------------------------------
    # 2. TEFAS Direct API
    # --------------------------------------------------------

    direct_df = fetch_tefas_direct_api(code)

    if (
        direct_df is not None
        and len(direct_df) >= 2
    ):

        return direct_df, "TEFAS Direct API"

    # --------------------------------------------------------
    # 3. İş Yatırım
    # --------------------------------------------------------

    is_df = fetch_isyatirim_series(code)

    if is_df is not None:

        return is_df, "İş Yatırım"

    return None, "YOK"


# ============================================================
# FON METRİKLERİ
# ============================================================

def compute_fund_metrics(
    series: Optional[pd.DataFrame],
    fund_code: str,
) -> Optional[dict]:

    if series is None or len(series) < 2:
        return None

    df = series.copy()

    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce",
    )

    df["price"] = (
        df["price"]
        .apply(parse_number)
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
        df
        .sort_values("date")
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

    # --------------------------------------------------------
    # Günlük getiriler
    # --------------------------------------------------------

    daily_returns = []

    for previous, current in zip(
        prices[:-1],
        prices[1:],
    ):

        if previous <= 0:

            daily_returns.append(0.0)

        else:

            daily_returns.append(
                (current / previous - 1.0)
                * 100.0
            )

    if not daily_returns:
        return None

    # --------------------------------------------------------
    # Max DD
    # --------------------------------------------------------

    max_dd = calculate_max_drawdown(
        prices
    )

    # --------------------------------------------------------
    # AUM değişimi
    # --------------------------------------------------------

    if aums[0] > 0:

        aum_change = (
            (aums[-1] / aums[0] - 1.0)
            * 100.0
        )

    else:

        aum_change = 0.0

    # --------------------------------------------------------
    # Yatırımcı değişimi
    # --------------------------------------------------------

    if investors[0] > 0:

        investor_change = (
            (investors[-1] / investors[0] - 1.0)
            * 100.0
        )

    else:

        investor_change = 0.0

    # --------------------------------------------------------
    # DOĞRU HAFTALIK GETİRİ
    #
    # Eski:
    # sum(daily_returns[-5:])
    #
    # Yeni:
    # bileşik getiri
    # --------------------------------------------------------

    recent_weekly_returns = (
        daily_returns[-5:]
        if len(daily_returns) >= 5
        else daily_returns
    )

    weekly_return = calculate_compounded_return(
        recent_weekly_returns
    )

    # --------------------------------------------------------
    # Yapısal veriler
    # --------------------------------------------------------

    structural = fetch_fund_structural_data(
        fund_code
    )

    return {
        "dates": dates[1:],
        "prices": prices,
        "daily_returns": daily_returns,
        "n_days": len(daily_returns),

        "aum": aums[-1],
        "investors": int(
            round(investors[-1])
        ),

        "aum_change": aum_change,
        "inv_change": investor_change,

        # Negatif tutuluyor:
        # -5 = %5 drawdown
        "max_dd": max_dd,

        "weekly_return": weekly_return,

        **structural,
    }


# ============================================================
# ROLLING METRİKLER
# ============================================================

def calculate_window_metrics(
    prices,
    returns,
    window,
):

    if len(returns) < window:
        return None

    if len(prices) < window + 1:
        return None

    slice_returns = returns[-window:]

    slice_prices = prices[-(window + 1):]

    mean_return = (
        sum(slice_returns)
        / len(slice_returns)
    )

    variance = sum(
        (r - mean_return) ** 2
        for r in slice_returns
    ) / len(slice_returns)

    volatility = variance ** 0.5

    if volatility > 1e-12:

        sharpe_like = (
            mean_return
            / volatility
        )

    else:

        sharpe_like = 0.0

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
# MOMENTUM SKORU
# ============================================================

def calculate_momentum_score(
    metrics,
    all_metrics,
):

    mean_z = zscore(
        [
            x["mean_return"]
            for x in all_metrics
        ]
    )

    sharpe_z = zscore(
        [
            x["sharpe"]
            for x in all_metrics
        ]
    )

    cumulative_z = zscore(
        [
            x["cumulative"]
            for x in all_metrics
        ]
    )

    # Drawdown zaten negatif:
    # -1 iyi değil, -10 daha kötü.
    #
    # Bu nedenle z'yi tersine çeviriyoruz.
    drawdown_z = zscore(
        [
            x["max_dd"]
            for x in all_metrics
        ]
    )

    current_index = all_metrics.index(
        metrics
    )

    weighted_z = (

        MOMENTUM_WEIGHTS["return"]
        * mean_z[current_index]

        +

        MOMENTUM_WEIGHTS["sharpe"]
        * sharpe_z[current_index]

        +

        MOMENTUM_WEIGHTS["cumulative"]
        * cumulative_z[current_index]

        -

        MOMENTUM_WEIGHTS["drawdown"]
        * drawdown_z[current_index]
    )

    # Z yaklaşık [-2.5, +2.5]
    # 50 merkez, 20 ölçek
    score = 50.0 + 20.0 * weighted_z

    return clamp(
        score,
        0.0,
        100.0,
    )


# ============================================================
# GÜVENLİK / LİKİDİTE SKORU
# ============================================================

def calculate_security_scores(
    funds,
):

    # --------------------------------------------------------
    # AUM Z
    # --------------------------------------------------------

    aum_values = [
        safe_float(
            x.get("aum")
        )
        for x in funds
    ]

    investor_values = [
        safe_float(
            x.get("investors")
        )
        for x in funds
    ]

    aum_change_values = [
        safe_float(
            x.get("aum_change")
        )
        for x in funds
    ]

    investor_change_values = [
        safe_float(
            x.get("inv_change")
        )
        for x in funds
    ]

    aum_z = zscore(
        aum_values
    )

    investor_z = zscore(
        investor_values
    )

    aum_change_z = zscore(
        aum_change_values
    )

    investor_change_z = zscore(
        investor_change_values
    )

    # --------------------------------------------------------
    # Konsantrasyon
    #
    # En büyük varlık yüksekse risk artar.
    # --------------------------------------------------------

    concentration_values = []

    for fund in funds:

        value = fund.get(
            "top_asset_weight"
        )

        if value is None:
            # Veri yoksa ortalama risk kabulü
            concentration_values.append(
                0.0
            )
        else:
            concentration_values.append(
                safe_float(value)
            )

    concentration_z = zscore(
        concentration_values
    )

    for i, fund in enumerate(funds):

        score = 50.0

        # AUM
        score += (
            15.0
            * aum_z[i]
        )

        # Yatırımcı
        score += (
            12.5
            * investor_z[i]
        )

        # AUM büyümesi
        score += (
            10.0
            * aum_change_z[i]
        )

        # Yatırımcı büyümesi
        score += (
            7.5
            * investor_change_z[i]
        )

        # ----------------------------------------------------
        # Konsantrasyon
        #
        # Z yüksek = yoğunlaşma yüksek = negatif
        # ----------------------------------------------------

        if fund.get(
            "top_asset_weight"
        ) is not None:

            score -= (
                12.5
                * concentration_z[i]
            )

        # ----------------------------------------------------
        # BIST30 bonus
        # ----------------------------------------------------

        if fund.get(
            "is_bist30",
            False,
        ):

            score += BIST30_BONUS

        # ----------------------------------------------------
        # Gerçek nakit verisi varsa
        # ----------------------------------------------------

        cash_ratio = fund.get(
            "emergency_cash_ratio"
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

                score -= 3.0

        # ----------------------------------------------------
        # Pozitif yatırımcı akışı
        # ----------------------------------------------------

        if safe_float(
            fund.get("inv_change")
        ) > 0:

            score += POSITIVE_INVESTOR_FLOW_BONUS

        # ----------------------------------------------------
        # Konsantrasyon doğrudan ceza
        # ----------------------------------------------------

        top_asset = fund.get(
            "top_asset_weight"
        )

        if top_asset is not None:

            if top_asset > 30:

                penalty = (
                    top_asset - 30
                ) * 1.0

                penalty = min(
                    penalty,
                    MAX_CONCENTRATION_PENALTY,
                )

                score -= penalty

            elif top_asset > 15:

                # Orta düzey konsantrasyon
                score -= (
                    top_asset - 15
                ) * 0.25

        # ----------------------------------------------------
        # Valör
        # ----------------------------------------------------

        score -= calculate_valor_penalty(
            fund.get("valor")
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
# ANA HİBRİT SKOR MOTORU
# ============================================================

def calculate_hybrid_scores(
    funds,
):

    if not funds:
        return

    # Tüm fonların ortak gün sayısı
    n_days = min(
        x["n_days"]
        for x in funds
    )

    # --------------------------------------------------------
    # Her fonun serisini hizala
    # --------------------------------------------------------

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

        fund["running_scores"] = []

        fund["running_momentum"] = []

    # --------------------------------------------------------
    # Yapısal güvenlik skoru
    # Günlük olarak değişmediği için bir kez hesaplanır.
    # --------------------------------------------------------

    calculate_security_scores(
        funds
    )

    # --------------------------------------------------------
    # Günlük momentum
    #
    # Son 5 günlük rolling pencere.
    #
    # Böylece:
    # 1. gün skoru ile 10. gün skoru farklı
    # pencerelerden gelen anlamsız skorlar olmaz.
    # --------------------------------------------------------

    for d in range(
        1,
        n_days + 1,
    ):

        current_metrics = []

        for fund in funds:

            available_window = min(
                MIN_ROLLING_DAYS,
                d,
            )

            returns_slice = (
                fund["daily_returns"][
                    d - available_window:d
                ]
            )

            prices_slice = (
                fund["prices"][
                    d - available_window:
                    d + 1
                ]
            )

            if len(
                returns_slice
            ) < 1:

                continue

            mean_return = (
                sum(returns_slice)
                / len(returns_slice)
            )

            variance = sum(
                (
                    r - mean_return
                ) ** 2
                for r in returns_slice
            ) / len(returns_slice)

            volatility = (
                variance ** 0.5
            )

            if volatility > 1e-12:

                sharpe = (
                    mean_return
                    / volatility
                )

            else:

                sharpe = 0.0

            cumulative = (
                calculate_compounded_return(
                    returns_slice
                )
            )

            max_dd = calculate_max_drawdown(
                prices_slice
            )

            current_metrics.append(
                {
                    "fund": fund,
                    "mean_return": mean_return,
                    "volatility": volatility,
                    "sharpe": sharpe,
                    "cumulative": cumulative,
                    "max_dd": max_dd,
                }
            )

        if not current_metrics:
            continue

        # ----------------------------------------------------
        # Z-score
        # ----------------------------------------------------

        mean_z = zscore(
            [
                x["mean_return"]
                for x in current_metrics
            ]
        )

        sharpe_z = zscore(
            [
                x["sharpe"]
                for x in current_metrics
            ]
        )

        cumulative_z = zscore(
            [
                x["cumulative"]
                for x in current_metrics
            ]
        )

        drawdown_z = zscore(
            [
                x["max_dd"]
                for x in current_metrics
            ]
        )

        # ----------------------------------------------------
        # Fon bazında momentum
        # ----------------------------------------------------

        for i, data in enumerate(
            current_metrics
        ):

            # Büyük negatif DD kötü.
            # DD zaten negatif olduğu için
            # yüksek negatif z'yi ödüllendirmemek adına
            # ters yön kullanıyoruz.

            weighted_z = (

                MOMENTUM_WEIGHTS["return"]
                * mean_z[i]

                +

                MOMENTUM_WEIGHTS["sharpe"]
                * sharpe_z[i]

                +

                MOMENTUM_WEIGHTS["cumulative"]
                * cumulative_z[i]

                -

                MOMENTUM_WEIGHTS["drawdown"]
                * drawdown_z[i]
            )

            momentum_score = (
                50.0
                + 20.0 * weighted_z
            )

            momentum_score = clamp(
                momentum_score,
                0.0,
                100.0,
            )

            fund = data["fund"]

            fund["running_momentum"].append(
                int(round(momentum_score))
            )

    # --------------------------------------------------------
    # Hibrit final
    # --------------------------------------------------------

    for fund in funds:

        momentum_scores = (
            fund["running_momentum"]
        )

        security_score = safe_float(
            fund.get(
                "security_score"
            ),
            50.0,
        )

        running_hybrid = []

        for momentum in momentum_scores:

            hybrid = (
                momentum
                * HYBRID_MOMENTUM_WEIGHT
                +
                security_score
                * HYBRID_SECURITY_WEIGHT
            )

            hybrid = clamp(
                hybrid,
                0.0,
                100.0,
            )

            running_hybrid.append(
                int(round(hybrid))
            )

        fund["running_scores"] = (
            running_hybrid
        )

        # ----------------------------------------------------
        # Son 5 gün
        # ----------------------------------------------------

        last_5 = (
            running_hybrid[-5:]
            if len(running_hybrid) >= 5
            else running_hybrid
        )

        fund["last_5_scores"] = last_5

        fund["last_5_scores_str"] = (
            " ➔ ".join(
                str(x)
                for x in last_5
            )
        )

        if last_5:

            # Son gün daha önemli
            weights = list(
                range(
                    1,
                    len(last_5) + 1,
                )
            )

            weighted_average = (
                sum(
                    score * weight
                    for score, weight
                    in zip(
                        last_5,
                        weights,
                    )
                )
                /
                sum(weights)
            )

            fund["kgdm_skor"] = int(
                round(
                    weighted_average
                )
            )

        else:

            fund["kgdm_skor"] = None

        # ----------------------------------------------------
        # Karar
        # ----------------------------------------------------

        score = fund.get(
            "kgdm_skor"
        )

        if score is None:

            fund["karar"] = (
                "YETERSİZ VERİ"
            )

        elif score >= STRONG_BUY:

            fund["karar"] = (
                "GÜÇLÜ AL"
            )

        elif score >= WATCH_LIST:

            fund["karar"] = (
                "ASIL LİSTE"
            )

        elif score >= CORRECTION:

            fund["karar"] = (
                "DÜZELTME / İZLE"
            )

        else:

            fund["karar"] = (
                "ACİL SAT"
            )

    return n_days


# ============================================================
# EXCEL STİL
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


# ============================================================
# EXCEL ÇIKTISI
# ============================================================

def create_excel_output(
    wb,
    ws_list,
    calculated_funds,
    n_days,
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
        "Valör",

        "Hibrit Skor",
        "Momentum Skor",
        "Güvenlik Skor",

        "Son 5 Skor",
        "Model Kararı",

        "Ort. Günlük Getiri (%)",
        "Volatilite (%)",
        "Sharpe",

        "Kümülatif Getiri (%)",
        "MaxDD (%)",

        "En Büyük Varlık (%)",
        "BIST30",
        "Nakit Verisi",
        "KAZRİSK",

        "AUM Değişim (%)",
        "Yatırımcı Değişim (%)",

        "AUM (₺)",
        "Yatırımcı",

        "Haftalık Bileşik Getiri (%)",
        "Veri Kaynağı",
    ]

    for day in calculated_funds[0]["dates"]:

        headers.append(
            f"{day} Hibrit Skor"
        )

    for day in calculated_funds[0]["dates"]:

        headers.append(
            f"{day} Getiri"
        )

    ws_scores.append(headers)

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

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

    ws_scores.row_dimensions[1].height = 42

    # --------------------------------------------------------
    # Satırlar
    # --------------------------------------------------------

    for item in calculated_funds:

        top_asset = item.get(
            "top_asset_weight"
        )

        if top_asset is None:

            risk_label = (
                "⚪ Veri Yok"
            )

        elif top_asset > 30:

            risk_label = (
                "⚠️ Yüksek Konsantrasyon"
            )

        elif top_asset > 15:

            risk_label = (
                "🟡 Orta Konsantrasyon"
            )

        else:

            risk_label = (
                "🛡️ Dengeli"
            )

        cash_known = item.get(
            "cash_ratio_known",
            False,
        )

        if cash_known:

            cash_label = (
                f"%{safe_float(item.get('emergency_cash_ratio')):.2f}"
            )

        else:

            cash_label = (
                "Veri Yok"
            )

        row_data = [

            item["code"],

            item.get(
                "valor",
                0,
            ),

            item.get(
                "kgdm_skor"
            ),

            (
                item.get(
                    "running_momentum",
                    []
                )[-1]
                if item.get(
                    "running_momentum"
                )
                else None
            ),

            item.get(
                "security_score"
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
                        "mean_return"
                    )
                ),
                4,
            ),

            round(
                safe_float(
                    item.get(
                        "volatility"
                    )
                ),
                4,
            ),

            round(
                safe_float(
                    item.get(
                        "sharpe_like"
                    )
                ),
                4,
            ),

            round(
                safe_float(
                    item.get(
                        "cumulative_return"
                    )
                ),
                4,
            ),

            round(
                safe_float(
                    item.get("max_dd")
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

            round(
                safe_float(
                    item.get(
                        "aum_change"
                    )
                ),
                2,
            ),

            round(
                safe_float(
                    item.get(
                        "inv_change"
                    )
                ),
                2,
            ),

            (
                round(
                    safe_float(
                        item.get("aum")
                    ),
                    2,
                )
                if item.get("aum")
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
        ]

        row_data.extend(
            item.get(
                "running_scores",
                [],
            )
        )

        row_data.extend(
            [
                format_percent(x)
                for x in item.get(
                    "daily_returns",
                    [],
                )
            ]
        )

        ws_scores.append(
            row_data
        )

    # --------------------------------------------------------
    # Renklendirme
    # --------------------------------------------------------

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

    for row_number in range(
        2,
        ws_scores.max_row + 1,
    ):

        decision_cell = (
            ws_scores.cell(
                row=row_number,
                column=7,
            )
        )

        decision_text = str(
            decision_cell.value
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

    # --------------------------------------------------------
    # Hibrit skor conditional formatting
    # --------------------------------------------------------

    score_range = (
        f"C2:C{ws_scores.max_row}"
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

    # --------------------------------------------------------
    # Formatlar
    # --------------------------------------------------------

    for row_number in range(
        2,
        ws_scores.max_row + 1,
    ):

        ws_scores.cell(
            row=row_number,
            column=19,
        ).number_format = (
            '#,##0.00 "₺"'
        )

        ws_scores.cell(
            row=row_number,
            column=20,
        ).number_format = (
            '#,##0'
        )

        for col in [
            8,
            9,
            11,
            12,
            13,
            17,
            18,
            21,
        ]:

            ws_scores.cell(
                row=row_number,
                column=col,
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

    wb.save(
        output
    )

    output.seek(0)

    return output


# ============================================================
# ANA ARAYÜZ
# ============================================================

st.subheader(
    "📂 Portföy Excel Listesi"
)

col_upload, col_github = st.columns(
    2
)

wb = None
source_mode = None

# ------------------------------------------------------------
# Upload
# ------------------------------------------------------------

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

            source_mode = "upload"

        except Exception as exc:

            st.error(
                f"Excel yükleme hatası: {exc}"
            )


# ------------------------------------------------------------
# GitHub
# ------------------------------------------------------------

with col_github:

    st.write(
        "Veya GitHub'daki listeyi kullanın:"
    )

    if st.button(
        "🚀 GitHub'dan Çek ve Analiz Et",
        use_container_width=True,
    ):

        try:

            response = requests.get(
                GITHUB_EXCEL_URL,
                timeout=HTTP_TIMEOUT,
            )

            response.raise_for_status()

            wb = openpyxl.load_workbook(
                io.BytesIO(
                    response.content
                )
            )

            source_mode = "github"

            st.success(
                "✅ Excel dosyası başarıyla indirildi."
            )

        except Exception as exc:

            st.error(
                f"GitHub bağlantı hatası: {exc}"
            )


if wb is None:

    st.info(
        "Analize başlamak için Excel dosyanızı yükleyin."
    )

    st.stop()


# ============================================================
# FON LİSTESİ
# ============================================================

if "Fon_Listesi" in wb.sheetnames:

    ws_list = wb[
        "Fon_Listesi"
    ]

else:

    ws_list = wb.active

requested_codes = []

excel_valor_dict = {}


for row in ws_list.iter_rows(
    min_row=2,
    values_only=False,
):

    if not row:
        continue

    if row[0].value is None:
        continue

    code = normalize_fund_code(
        row[0].value
    )

    if not code:
        continue

    requested_codes.append(
        code
    )

    # Valör 4. kolon
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
        "Fon_Listesi sayfasında fon kodu bulunamadı."
    )

    st.stop()


# ============================================================
# TARİH
# ============================================================

today = dt.date.today()

start_date = (
    today -
    dt.timedelta(
        days=LOOKBACK_CALENDAR_DAYS
    )
)


# ============================================================
# TEFAS UNIVERSE
# ============================================================

with st.spinner(
    "🔄 TEFAS verileri alınıyor..."
):

    universe = fetch_tefas_universe(
        start_date,
        today,
    )


# ============================================================
# FONLARIN HESAPLANMASI
# ============================================================

calculated_funds = []

failed_codes = []

progress = st.progress(
    0,
    text="Fonlar analiz ediliyor..."
)

total_funds = len(
    requested_codes
)


for index, code in enumerate(
    requested_codes,
    start=1,
):

    series, source = get_fund_series(
        universe,
        code,
    )

    metrics = compute_fund_metrics(
        series,
        code,
    )

    if metrics is None:

        failed_codes.append(
            code
        )

        progress.progress(
            index / total_funds,
            text=f"{code}: veri yok",
        )

        continue

    metrics["code"] = code

    metrics["valor"] = (
        excel_valor_dict.get(
            code,
            0.0,
        )
    )

    metrics["source"] = source

    # --------------------------------------------------------
    # Filtre
    # --------------------------------------------------------

    if ENABLE_FILTERS:

        investor_ok = (
            metrics["investors"]
            >= MIN_INVESTOR_COUNT
        )

        weekly_ok = (
            metrics["weekly_return"]
            >= TARGET_WEEKLY_RETURN
        )

        if not (
            investor_ok
            and weekly_ok
        ):

            progress.progress(
                index / total_funds,
                text=f"{code}: filtre dışı",
            )

            continue

    calculated_funds.append(
        metrics
    )

    progress.progress(
        index / total_funds,
        text=f"{code}: analiz edildi",
    )


progress.empty()


if failed_codes:

    st.warning(
        "Veri bulunamayan fonlar: "
        + ", ".join(
            failed_codes
        )
    )


if not calculated_funds:

    st.error(
        "Hesaplanabilecek geçerli fon verisi bulunamadı."
    )

    st.stop()


# ============================================================
# ORTAK GÜN SAYISI
# ============================================================

n_days = min(
    item["n_days"]
    for item in calculated_funds
)

if n_days < 2:

    st.error(
        "Fonlarda yeterli tarihsel veri bulunmuyor."
    )

    st.stop()


# ============================================================
# HİBRİT SKOR MOTORU
# ============================================================

with st.spinner(
    "📊 KGDM-3 + KAZRİSK hibrit skoru hesaplanıyor..."
):

    n_days = calculate_hybrid_scores(
        calculated_funds
    )


# ============================================================
# SON DÖNEM METRİKLERİ
# ============================================================

for item in calculated_funds:

    metrics = calculate_window_metrics(
        item["prices"],
        item["daily_returns"],
        min(
            MIN_ROLLING_DAYS,
            item["n_days"],
        ),
    )

    if metrics is None:

        item["mean_return"] = 0.0
        item["volatility"] = 0.0
        item["sharpe_like"] = 0.0
        item["cumulative_return"] = 0.0

    else:

        item["mean_return"] = (
            metrics["mean_return"]
        )

        item["volatility"] = (
            metrics["volatility"]
        )

        item["sharpe_like"] = (
            metrics["sharpe"]
        )

        item["cumulative_return"] = (
            metrics["cumulative"]
        )


# ============================================================
# SIRALAMA
# ============================================================

calculated_funds.sort(
    key=lambda x: (
        -safe_float(
            x.get("kgdm_skor")
        ),
        -safe_float(
            x.get("cumulative_return")
        ),
    )
)


# ============================================================
# SONUÇ TABLOSU
# ============================================================

display_rows = []

for item in calculated_funds:

    top_asset = item.get(
        "top_asset_weight"
    )

    if top_asset is None:

        risk_status = (
            "⚪ Veri Yok"
        )

    elif top_asset > 30:

        risk_status = (
            "⚠️ Yüksek Konsantrasyon"
        )

    elif top_asset > 15:

        risk_status = (
            "🟡 Orta Konsantrasyon"
        )

    else:

        risk_status = (
            "🛡️ Dengeli"
        )

    display_rows.append(
        {
            "Fon Kodu":
                item["code"],

            "Hibrit Skor":
                item.get(
                    "kgdm_skor"
                ),

            "Momentum":
                (
                    item.get(
                        "running_momentum",
                        []
                    )[-1]
                    if item.get(
                        "running_momentum"
                    )
                    else None
                ),

            "Güvenlik":
                item.get(
                    "security_score"
                ),

            "Model Kararı":
                item.get(
                    "karar"
                ),

            "Ort. Günlük %":
                round(
                    safe_float(
                        item.get(
                            "mean_return"
                        )
                    ),
                    3,
                ),

            "Sharpe":
                round(
                    safe_float(
                        item.get(
                            "sharpe_like"
                        )
                    ),
                    3,
                ),

            "Kümülatif Getiri %":
                round(
                    safe_float(
                        item.get(
                            "cumulative_return"
                        )
                    ),
                    3,
                ),

            "MaxDD %":
                round(
                    safe_float(
                        item.get(
                            "max_dd"
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
                round(
                    safe_float(
                        item.get(
                            "aum"
                        )
                    ),
                    0,
                ),

            "Yatırımcı":
                item.get(
                    "investors"
                ),

            "Kaynak":
                item.get(
                    "source"
                ),
        }
    )


df_display = pd.DataFrame(
    display_rows
)


# ============================================================
# DATAFRAME RENKLENDİRME
# ============================================================

def color_cells(value):

    text = str(value)

    if (
        "GÜÇLÜ AL"
        in text
        or "ASIL LİSTE"
        in text
        or "Dengeli"
        in text
    ):

        return (
            "color: #008000; "
            "font-weight: bold;"
        )

    if (
        "DÜZELTME"
        in text
        or "Orta"
        in text
    ):

        return (
            "color: #B8860B; "
            "font-weight: bold;"
        )

    if (
        "ACİL SAT"
        in text
        or "Yüksek"
        in text
    ):

        return (
            "color: #FF0000; "
            "font-weight: bold;"
        )

    return ""


try:

    styled_df = (
        df_display
        .style
        .map(color_cells)
    )

except AttributeError:

    styled_df = (
        df_display
        .style
        .applymap(color_cells)
    )


# ============================================================
# EKRAN
# ============================================================

st.subheader(
    "📊 Analiz Sonuçları"
)

st.dataframe(
    styled_df,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# SKOR DAĞILIMI
# ============================================================

st.subheader(
    "📈 Skor Özeti"
)

col1, col2, col3, col4 = st.columns(
    4
)

scores = [
    safe_float(
        x.get("kgdm_skor")
    )
    for x in calculated_funds
    if x.get("kgdm_skor") is not None
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
            for x in calculated_funds
            if x.get("karar")
            == "GÜÇLÜ AL"
        )

        st.metric(
            "Güçlü Al",
            strong_count,
        )


# ============================================================
# EXCEL ÇIKTISI
# ============================================================

output = create_excel_output(
    wb=wb,
    ws_list=ws_list,
    calculated_funds=calculated_funds,
    n_days=n_days,
)


st.success(
    f"✅ Analiz tamamlandı. "
    f"{len(calculated_funds)} fon hesaplandı. "
    f"Model sürümü: {APP_VERSION}"
)


st.download_button(
    label="📥 Güncellenmiş Hibrit Excel'i İndir",
    data=output,
    file_name="fonlar_KGDM3_KAZRISK_V6_1.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)
