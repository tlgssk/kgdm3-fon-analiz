import datetime as dt
import io
import math
from typing import Optional, List

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# SAYFA AYARLARI
# ============================================================

st.set_page_config(
    page_title="Multi-Vade Fon Analizi V7",
    page_icon="📈",
    layout="wide",
)

st.title("📈 Multi-Vade Fon Analizi V7")
st.caption(
    "Gelişmiş Z-Skor Motoru + Kategori/Global Harmanlama + "
    "Normalize Valor + Tanh Benchmark + Asimetrik Sortino"
)


# ============================================================
# GENEL AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")

LOOKBACK_CALENDAR_DAYS = 400
HTTP_TIMEOUT = 15

PERIODS = {
    "Kısa Vade (1 Hafta)": 5,
    "Orta Vade (1 Ay)": 21,
    "Uzun Vade (3 Ay)": 63,
    "Çok Uzun Vade (1 Yıl)": 252,
}

MAX_DAYS = max(PERIODS.values())

FINAL_WEIGHTS = {
    5: 0.10,
    21: 0.20,
    63: 0.30,
    252: 0.40,
}

# Nihai skor üretmek için minimum kaç vade gerekir?
MIN_FINAL_PERIODS = 3

# V7 metrik ağırlıkları
W_CUM = 0.25
W_SHP = 0.25
W_SRT = 0.25
W_MDD = 0.25

# Z-score sınırı
Z_LIMIT = 2.5

# Valor cezası:
# Ceza hiçbir durumda 5 puanı aşmaz.
VALOR_PENALTY_MAX = 5.0
VALOR_SCALE = 5.0

# KUT benchmark ağırlıkları
KUT_GOLD_WEIGHT = 0.45
KUT_SILVER_WEIGHT = 0.45
KUT_CASH_WEIGHT = 0.10

COLOR_NAVY = "1F4E79"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"
COLOR_LIGHT_GREEN = "E2F0D9"
COLOR_LIGHT_RED = "FCE4D6"
COLOR_LIGHT_YELLOW = "FFF2CC"


# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

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


def safe_float(value, default=0.0):
    try:
        if value is None:
            return default

        value = float(value)

        if math.isnan(value) or math.isinf(value):
            return default

        return value

    except Exception:
        return default


def infer_category(title: str) -> str:
    if not title:
        return "Diğer"

    t = title.upper()

    if any(
        k in t
        for k in [
            "ALTIN",
            "GOLD",
            "KIYMETLİ MADEN",
            "GÜMÜŞ",
            "SILVER",
            "KITA",
        ]
    ):
        return "Kıymetli Maden"

    if any(
        k in t
        for k in [
            "HİSSE",
            "EQUITY",
            "HİSSE SENEDİ",
            "HİSSE YOĞUN",
        ]
    ):
        return "Hisse Senedi"

    if any(
        k in t
        for k in [
            "PARA PİYASASI",
            "LIKIT",
            "LİKİT",
            "PARA PIYASASI",
        ]
    ):
        return "Para Piyasası"

    if any(
        k in t
        for k in [
            "BORÇLANMA",
            "TAHVİL",
            "BONO",
            "BORÇLANMA ARAÇLARI",
            "KİRA SERTİFİKASI",
        ]
    ):
        return "Borçlanma"

    if any(
        k in t
        for k in [
            "KARMA",
            "DEĞİŞKEN",
            "DENGELİ",
            "ÇOKLU VARLIK",
            "FON SEPETİ",
        ]
    ):
        return "Karma / Değişken"

    if any(
        k in t
        for k in [
            "YABANCI",
            "EUROBOND",
            "DIŞ BORÇ",
            "USD",
            "EUR",
            "DÖVİZ",
        ]
    ):
        return "Yabancı / Döviz"

    if "SERBEST" in t:
        return "Serbest"

    if "KATILIM" in t:
        return "Katılım"

    return "Diğer"


# ============================================================
# TEFAS VERİSİ
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
        crawler = Crawler(timeout=60, max_retry=5)

        df = crawler.fetch_many(
            start=start_date,
            end=end_date,
            kinds=FUND_KINDS,
            columns="info",
        )

        if df is None or df.empty:
            return pd.DataFrame()

        df = df.copy()

        api_category_col = next(
            (
                col
                for col in [
                    "fund_umbrella_title",
                    "umbrella_fund_type",
                    "fon_turu",
                ]
                if col in df.columns
            ),
            None,
        )

        df["spk_category"] = (
            df[api_category_col].astype(str)
            if api_category_col
            else ""
        )

        df.rename(
            columns={
                "fund_code": "code",
                "fund_name": "title",
                "investor_count": "investors",
                "portfolio_size": "aum",
            },
            inplace=True,
        )

        required = ["date", "code", "price"]

        if not all(c in df.columns for c in required):
            return pd.DataFrame()

        df["code"] = (
            df["code"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df["date"] = pd.to_datetime(
            df["date"],
            errors="coerce",
        )

        df["price"] = df["price"].apply(parse_number)

        if "aum" in df.columns:
            df["aum"] = df["aum"].apply(parse_number)
        else:
            df["aum"] = 0.0

        if "investors" in df.columns:
            df["investors"] = df["investors"].apply(parse_number)
        else:
            df["investors"] = 0.0

        if "title" not in df.columns:
            df["title"] = ""

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


def get_fund_series(
    universe: pd.DataFrame,
    fund_code: str,
) -> Optional[pd.DataFrame]:

    if universe is None or universe.empty:
        return None

    code = normalize_fund_code(fund_code)

    rows = universe[
        universe["code"]
        .astype(str)
        .str.upper()
        .eq(code)
    ].copy()

    if rows.empty:
        return None

    rows = (
        rows.sort_values("date")
        .drop_duplicates(
            subset=["date"],
            keep="last",
        )
    )

    if len(rows) < 2:
        return None

    if len(rows) > MAX_DAYS + 1:
        rows = rows.tail(MAX_DAYS + 1)

    return rows.reset_index(drop=True)


def fetch_isyatirim_series(
    fund_code: str,
) -> Optional[pd.DataFrame]:

    code = normalize_fund_code(fund_code)

    if not code:
        return None

    end = dt.datetime.now()

    start = (
        end
        - dt.timedelta(
            days=LOOKBACK_CALENDAR_DAYS
        )
    )

    url = (
        "https://www.isyatirim.com.tr/"
        "_layouts/15/IsYatirim.Website/"
        "Common/Data.aspx/"
        "YatirimFonGecmisGetiri"
    )

    params = {
        "fonKod": code,
        "baslangic": start.strftime("%d-%m-%Y"),
        "bitis": end.strftime("%d-%m-%Y"),
    }

    try:
        response = requests.get(
            url,
            params=params,
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=HTTP_TIMEOUT,
        )

        response.raise_for_status()

        values = response.json().get("value")

        if not values:
            return None

        df = pd.DataFrame(values)

        if "Tarih" not in df.columns:
            return None

        if "Fiyat" not in df.columns:
            return None

        df["date"] = pd.to_datetime(
            df["Tarih"],
            dayfirst=True,
            errors="coerce",
        )

        df["price"] = df["Fiyat"].apply(parse_number)

        df = df.dropna(
            subset=["date", "price"]
        )

        df = df[df["price"] > 0]

        if len(df) < 2:
            return None

        return (
            df.sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .tail(MAX_DAYS + 1)
            .reset_index(drop=True)[
                ["date", "price"]
            ]
        )

    except Exception:
        return None


# ============================================================
# MDD
# ============================================================

def compute_max_drawdown(
    prices: List[float],
) -> float:

    if not prices or len(prices) < 2:
        return 0.0

    peak = prices[0]
    max_dd = 0.0

    for p in prices:

        if p > peak:
            peak = p

        if peak > 0:

            dd = (
                p / peak - 1.0
            ) * 100.0

            if dd < max_dd:
                max_dd = dd

    return max_dd


# ============================================================
# V7 SORTINO
# ============================================================

def compute_sortino(
    returns: List[float],
    daily_rf: float = 0.0,
    max_sortino: float = 10.0,
) -> float:
    """
    V7 Sortino.

    Downside deviation bütün gözlem sayısına
    göre hesaplanır.

    Böylece negatif gözlem sayısı değiştiğinde
    metodoloji değişmez.
    """

    if not returns:
        return 0.0

    clean_returns = []

    for r in returns:
        try:
            r = float(r)

            if math.isfinite(r):
                clean_returns.append(r)

        except Exception:
            continue

    if not clean_returns:
        return 0.0

    excess = [
        r - daily_rf
        for r in clean_returns
    ]

    mean_excess = (
        sum(excess) / len(excess)
    )

    downside_squared = [
        min(0.0, x) ** 2
        for x in excess
    ]

    downside_deviation = math.sqrt(
        sum(downside_squared)
        / len(excess)
    )

    if downside_deviation <= 1e-12:

        if mean_excess > 0:
            return max_sortino

        return 0.0

    sortino = (
        mean_excess
        / downside_deviation
    )

    return max(
        -max_sortino,
        min(max_sortino, sortino),
    )


# ============================================================
# FON METRİKLERİ
# ============================================================

def compute_fund_metrics(
    series: Optional[pd.DataFrame],
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

    daily_returns = []

    for i in range(1, len(prices)):

        if prices[i - 1] > 0:

            ret = (
                prices[i]
                / prices[i - 1]
                - 1
            ) * 100

        else:
            ret = 0.0

        daily_returns.append(ret)

    title = (
        str(df["title"].iloc[-1] or "")
        if "title" in df.columns
        else ""
    )

    spk_cat = (
        str(df["spk_category"].iloc[-1]).strip()
        if "spk_category" in df.columns
        else ""
    )

    final_category = (
        spk_cat
        if spk_cat
        and spk_cat.lower()
        not in ["nan", "none", "null"]
        else infer_category(title)
    )

    return {
        "prices": prices,
        "daily_returns": daily_returns,
        "title": title,
        "aum": (
            safe_float(df["aum"].iloc[-1])
            if "aum" in df.columns
            else 0.0
        ),
        "investors": (
            safe_float(df["investors"].iloc[-1])
            if "investors" in df.columns
            else 0.0
        ),
        "category": final_category,
    }


# ============================================================
# V7 VALOR CEZASI
# ============================================================

def calculate_valor_penalty(valor) -> float:
    """
    Valor etkisini -5 / +5 aralığında sınırlar.

    Pozitif valor = negatif puan etkisi.
    """

    value = safe_float(valor, 0.0)

    normalized = math.tanh(
        value / VALOR_SCALE
    )

    penalty = (
        VALOR_PENALTY_MAX
        * normalized
    )

    return penalty


# ============================================================
# V7 VADE SKORU
# ============================================================

def calculate_period_scores(
    funds: List[dict],
    days: int,
    daily_rf: float,
    use_category: bool,
):
    """
    V7 Vade Skoru

    Metrikler:
        Kümülatif Getiri : %25
        Sharpe           : %25
        Sortino          : %25
        MDD              : %25

    ÖNEMLİ:
    Valor burada uygulanmaz.

    Önce istatistiksel skor hesaplanır.
    Daha sonra kategori/global harmanlama yapılır.
    Valor en son bir kez uygulanır.
    """

    # --------------------------------------------------------
    # 1. HAM METRİKLER
    # --------------------------------------------------------

    for item in funds:

        returns = item.get(
            "daily_returns",
            [],
        )

        prices = item.get(
            "prices",
            [],
        )

        if (
            len(returns) < days
            or len(prices) < days + 1
        ):

            item[f"score_{days}"] = None
            item[f"karar_{days}"] = (
                "YETERSİZ VERİ"
            )

            for m in [
                "mean",
                "vol",
                "shp",
                "cum",
                "mdd",
                "sortino",
            ]:
                item[f"{m}_{days}"] = None

            continue

        slice_ret = returns[-days:]

        slice_prices = prices[
            -(days + 1):
        ]

        # Ortalama günlük getiri
        mean_ret = (
            sum(slice_ret)
            / len(slice_ret)
        )

        # Günlük volatilite
        variance = sum(
            (r - mean_ret) ** 2
            for r in slice_ret
        ) / len(slice_ret)

        vol = math.sqrt(
            max(0.0, variance)
        )

        # Yıllıklandırılmış Sharpe
        if vol > 1e-12:

            sharpe_daily = (
                mean_ret - daily_rf
            ) / vol

            sharpe = (
                sharpe_daily
                * math.sqrt(252)
            )

        else:
            sharpe = 0.0

        start_price = slice_prices[0]
        end_price = slice_prices[-1]

        if start_price > 0:

            cumulative = (
                end_price
                / start_price
                - 1.0
            ) * 100.0

        else:
            cumulative = None

        mdd = compute_max_drawdown(
            slice_prices
        )

        sortino = compute_sortino(
            slice_ret,
            daily_rf=daily_rf,
        )

        item[f"mean_{days}"] = mean_ret
        item[f"vol_{days}"] = vol
        item[f"shp_{days}"] = sharpe
        item[f"cum_{days}"] = cumulative
        item[f"mdd_{days}"] = mdd
        item[f"sortino_{days}"] = sortino

    # --------------------------------------------------------
    # 2. GEÇERLİ FONLAR
    # --------------------------------------------------------

    valid_indices = [
        i
        for i, f in enumerate(funds)
        if (
            not f.get("filtered_out", False)
            and f.get(f"cum_{days}") is not None
            and f.get(f"shp_{days}") is not None
            and f.get(f"sortino_{days}") is not None
            and f.get(f"mdd_{days}") is not None
        )
    ]

    if not valid_indices:

        for item in funds:

            item[f"score_{days}"] = None

            if item.get("filtered_out"):
                item[f"karar_{days}"] = (
                    "FİLTRE DIŞI"
                )
            else:
                item[f"karar_{days}"] = (
                    "YETERSİZ VERİ"
                )

        return

    # --------------------------------------------------------
    # 3. KATEGORİLER
    # --------------------------------------------------------

    categories = {}

    if use_category:

        for idx in valid_indices:

            category = (
                funds[idx].get("category")
                or "Diğer"
            )

            categories.setdefault(
                category,
                [],
            ).append(idx)

    else:

        categories = {
            "ALL": valid_indices
        }

    # --------------------------------------------------------
    # 4. İSTATİSTİK
    # --------------------------------------------------------

    def clean_values(
        indices,
        field,
    ):

        values = []

        for i in indices:

            value = funds[i].get(
                f"{field}_{days}"
            )

            if value is None:
                continue

            try:
                value = float(value)

                if math.isfinite(value):
                    values.append(value)

            except Exception:
                continue

        return values

    def get_stats(values):

        if not values:
            return 0.0, 0.0

        mean_value = (
            sum(values)
            / len(values)
        )

        variance = sum(
            (x - mean_value) ** 2
            for x in values
        ) / len(values)

        return (
            mean_value,
            math.sqrt(
                max(0.0, variance)
            ),
        )

    def z_score(
        value,
        mean_value,
        std,
    ):

        if value is None:
            return 0.0

        if std <= 1e-12:
            return 0.0

        z = (
            float(value)
            - mean_value
        ) / std

        return max(
            -Z_LIMIT,
            min(Z_LIMIT, z),
        )

    # --------------------------------------------------------
    # 5. SADECE İSTATİSTİKSEL SKOR
    # --------------------------------------------------------

    def calculate_statistical_score(
        eval_indices,
        target_indices,
    ):

        if len(eval_indices) < 2:

            for idx in target_indices:

                funds[idx][
                    f"stat_score_{days}"
                ] = None

            return

        mean_cum = get_stats(
            clean_values(
                eval_indices,
                "cum",
            )
        )

        mean_shp = get_stats(
            clean_values(
                eval_indices,
                "shp",
            )
        )

        mean_srt = get_stats(
            clean_values(
                eval_indices,
                "sortino",
            )
        )

        mean_mdd = get_stats(
            clean_values(
                eval_indices,
                "mdd",
            )
        )

        for idx in target_indices:

            item = funds[idx]

            z_cum = z_score(
                item.get(
                    f"cum_{days}"
                ),
                mean_cum[0],
                mean_cum[1],
            )

            z_shp = z_score(
                item.get(
                    f"shp_{days}"
                ),
                mean_shp[0],
                mean_shp[1],
            )

            z_srt = z_score(
                item.get(
                    f"sortino_{days}"
                ),
                mean_srt[0],
                mean_srt[1],
            )

            z_mdd = z_score(
                item.get(
                    f"mdd_{days}"
                ),
                mean_mdd[0],
                mean_mdd[1],
            )

            weighted_z = (
                W_CUM * z_cum
                + W_SHP * z_shp
                + W_SRT * z_srt
                + W_MDD * z_mdd
            )

            stat_score = (
                50.0
                + 20.0 * weighted_z
            )

            item[
                f"stat_score_{days}"
            ] = int(
                round(
                    max(
                        0.0,
                        min(
                            100.0,
                            stat_score,
                        ),
                    )
                )
            )

    # --------------------------------------------------------
    # 6. GLOBAL SKOR
    # --------------------------------------------------------

    calculate_statistical_score(
        valid_indices,
        valid_indices,
    )

    global_scores = {
        idx: funds[idx].get(
            f"stat_score_{days}"
        )
        for idx in valid_indices
    }

    # --------------------------------------------------------
    # 7. KATEGORİ / GLOBAL HARMANLAMA
    # --------------------------------------------------------

    for category, indices in categories.items():

        n = len(indices)

        # --------------------------------------------
        # 10+ fon:
        # tamamen kategori
        # --------------------------------------------

        if n >= 10:

            calculate_statistical_score(
                indices,
                indices,
            )

            for idx in indices:

                cat_score = funds[idx].get(
                    f"stat_score_{days}"
                )

                funds[idx][
                    f"score_{days}"
                ] = cat_score

        # --------------------------------------------
        # 5-9 fon:
        # %60 kategori + %40 global
        # --------------------------------------------

        elif n >= 5:

            calculate_statistical_score(
                indices,
                indices,
            )

            for idx in indices:

                cat_score = funds[idx].get(
                    f"stat_score_{days}"
                )

                glob_score = global_scores.get(
                    idx
                )

                if (
                    cat_score is not None
                    and glob_score is not None
                ):

                    blended = (
                        0.60 * cat_score
                        + 0.40 * glob_score
                    )

                elif cat_score is not None:
                    blended = cat_score

                else:
                    blended = glob_score

                funds[idx][
                    f"score_{days}"
                ] = (
                    int(round(blended))
                    if blended is not None
                    else None
                )

        # --------------------------------------------
        # <5 fon:
        # global
        # --------------------------------------------

        else:

            for idx in indices:

                funds[idx][
                    f"score_{days}"
                ] = global_scores.get(
                    idx
                )

    # --------------------------------------------------------
    # 8. VALOR CEZASI
    # --------------------------------------------------------

    for item in funds:

        if item.get("filtered_out", False):

            item[f"score_{days}"] = None
            item[f"karar_{days}"] = (
                "FİLTRE DIŞI"
            )
            continue

        base_score = item.get(
            f"score_{days}"
        )

        if base_score is None:

            item[f"karar_{days}"] = (
                "YETERSİZ VERİ"
            )

            continue

        valor_penalty = (
            calculate_valor_penalty(
                item.get("valor", 0)
            )
        )

        item[
            f"valor_penalty_{days}"
        ] = valor_penalty

        final_period_score = (
            float(base_score)
            - valor_penalty
        )

        final_period_score = max(
            0.0,
            min(
                100.0,
                final_period_score,
            ),
        )

        item[
            f"score_{days}"
        ] = int(
            round(
                final_period_score
            )
        )

        score = item[
            f"score_{days}"
        ]

        if score >= 60:
            item[
                f"karar_{days}"
            ] = "GÜÇLÜ AL"

        elif score >= 40:
            item[
                f"karar_{days}"
            ] = "ASIL LİSTE"

        elif score >= 25:
            item[
                f"karar_{days}"
            ] = "NÖTR"

        else:
            item[
                f"karar_{days}"
            ] = "ACİL SAT"


# ============================================================
# YAHOO
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_yahoo_series(
    symbol: str,
    start_date: dt.date,
    end_date: dt.date,
) -> Optional[pd.DataFrame]:

    try:

        p1 = int(
            dt.datetime.combine(
                start_date,
                dt.time.min,
            ).timestamp()
        )

        p2 = int(
            dt.datetime.combine(
                end_date
                + dt.timedelta(days=1),
                dt.time.min,
            ).timestamp()
        )

        url = (
            "https://query1.finance.yahoo.com/"
            f"v8/finance/chart/{symbol}"
        )

        res = requests.get(
            url,
            params={
                "period1": p1,
                "period2": p2,
                "interval": "1d",
                "events": "history",
                "includeAdjustedClose": "true",
            },
            headers={
                "User-Agent": "Mozilla/5.0"
            },
            timeout=HTTP_TIMEOUT,
        )

        res.raise_for_status()

        result = (
            res.json()
            .get("chart", {})
            .get("result", [{}])[0]
        )

        timestamps = result.get(
            "timestamp"
        )

        closes = (
            result
            .get("indicators", {})
            .get("quote", [{}])[0]
            .get("close")
        )

        if not timestamps or not closes:
            return None

        rows = []

        for t, c in zip(
            timestamps,
            closes,
        ):

            if c is None:
                continue

            try:
                c = float(c)

                if c <= 0:
                    continue

                rows.append(
                    {
                        "date": pd.Timestamp(
                            dt.datetime
                            .fromtimestamp(t)
                            .date()
                        ),
                        "price": c,
                    }
                )

            except Exception:
                continue

        if len(rows) < 2:
            return None

        return (
            pd.DataFrame(rows)
            .sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .reset_index(drop=True)
        )

    except Exception:
        return None


# ============================================================
# KUT BENCHMARK
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_kut_benchmark(
    start_date: dt.date,
    end_date: dt.date,
) -> Optional[pd.DataFrame]:

    gold = fetch_yahoo_series(
        "GC=F",
        start_date,
        end_date,
    )

    silver = fetch_yahoo_series(
        "SI=F",
        start_date,
        end_date,
    )

    usdtry = fetch_yahoo_series(
        "USDTRY=X",
        start_date,
        end_date,
    )

    if (
        gold is None
        or silver is None
        or usdtry is None
    ):
        return None

    df = (
        gold.rename(
            columns={"price": "gold"}
        )
        .merge(
            silver.rename(
                columns={"price": "silver"}
            ),
            on="date",
            how="inner",
        )
        .merge(
            usdtry.rename(
                columns={"price": "usd"}
            ),
            on="date",
            how="inner",
        )
    )

    if df.empty:
        return None

    g_try = (
        df["gold"] * df["usd"]
    )

    s_try = (
        df["silver"] * df["usd"]
    )

    df["benchmark"] = (
        KUT_GOLD_WEIGHT
        * (
            g_try
            / g_try.iloc[0]
        )
        +
        KUT_SILVER_WEIGHT
        * (
            s_try
            / s_try.iloc[0]
        )
        +
        KUT_CASH_WEIGHT
    )

    return df[
        ["date", "benchmark"]
    ].reset_index(drop=True)


# ============================================================
# BENCHMARK SKORU
# ============================================================

def benchmark_score_from_diff(
    diff: float,
) -> Optional[int]:

    if diff is None:
        return None

    normalized = math.tanh(
        diff / 5.0
    )

    score = (
        50.0
        + 40.0 * normalized
    )

    return int(
        round(
            max(
                0.0,
                min(100.0, score),
            )
        )
    )


def calculate_kut_benchmark_metrics(
    funds,
    benchmark_df,
):

    bench_returns = {}

    for d in PERIODS.values():

        if (
            benchmark_df is not None
            and not benchmark_df.empty
            and len(benchmark_df) > d
        ):

            start_value = (
                benchmark_df[
                    "benchmark"
                ].iloc[-(d + 1)]
            )

            end_value = (
                benchmark_df[
                    "benchmark"
                ].iloc[-1]
            )

            if start_value > 0:

                bench_returns[d] = (
                    (
                        end_value
                        / start_value
                        - 1
                    )
                    * 100
                )

            else:
                bench_returns[d] = None

        else:
            bench_returns[d] = None

    for item in funds:

        if item["code"] != "KUT":

            item["benchmark_active"] = False

            continue

        item["benchmark_active"] = True

        for days in PERIODS.values():

            fr = item.get(
                f"cum_{days}"
            )

            br = bench_returns.get(
                days
            )

            if fr is None or br is None:

                item[
                    f"benchmark_{days}"
                ] = None

                item[
                    f"benchmark_diff_{days}"
                ] = None

                item[
                    f"benchmark_score_{days}"
                ] = None

            else:

                diff = fr - br

                item[
                    f"benchmark_{days}"
                ] = br

                item[
                    f"benchmark_diff_{days}"
                ] = diff

                item[
                    f"benchmark_score_{days}"
                ] = benchmark_score_from_diff(
                    diff
                )


# ============================================================
# TREND
# ============================================================

def calculate_trends(funds):

    for item in funds:

        s63 = item.get("score_63")
        s252 = item.get("score_252")
        s21 = item.get("score_21")

        if s63 is None or s21 is None:

            item["trend_3a"] = (
                "VERİ YOK"
            )

        else:

            diff = s21 - s63

            if diff >= 10:
                item["trend_3a"] = (
                    "YUKARI ↑"
                )

            elif diff <= -10:
                item["trend_3a"] = (
                    "AŞAĞI ↓"
                )

            else:
                item["trend_3a"] = (
                    "YATAY →"
                )

        if s252 is None or s63 is None:

            item["trend_1y"] = (
                "VERİ YOK"
            )

        else:

            diff = s63 - s252

            if diff >= 10:
                item["trend_1y"] = (
                    "YUKARI ↑"
                )

            elif diff <= -10:
                item["trend_1y"] = (
                    "AŞAĞI ↓"
                )

            else:
                item["trend_1y"] = (
                    "YATAY →"
                )


# ============================================================
# GÜVEN SEVİYESİ V7
# ============================================================

def calculate_confidence(item):

    scores = []

    for days in PERIODS.values():

        score = item.get(
            f"score_{days}"
        )

        if score is None:
            continue

        try:

            score = float(score)

            if math.isfinite(score):
                scores.append(score)

        except Exception:
            continue

    period_count = len(scores)

    if period_count == 0:
        return "DÜŞÜK"

    data_ratio = (
        period_count
        / len(PERIODS)
    )

    mean_score = (
        sum(scores)
        / len(scores)
    )

    dispersion = math.sqrt(
        sum(
            (x - mean_score) ** 2
            for x in scores
        )
        / len(scores)
    )

    # 4/4 veri
    if (
        period_count == 4
        and dispersion <= 8
    ):
        return "ÇOK YÜKSEK"

    if (
        period_count == 4
        and dispersion <= 15
    ):
        return "YÜKSEK"

    # En az 3 vade
    if (
        period_count >= 3
        and dispersion <= 15
    ):
        return "YÜKSEK"

    if (
        period_count >= 3
        and dispersion <= 25
    ):
        return "ORTA"

    if data_ratio >= 0.50:
        return "ORTA"

    return "DÜŞÜK"


# ============================================================
# NİHAİ SKOR V7
# ============================================================

def calculate_final_scores(funds):

    for item in funds:

        available_scores = []

        for days, weight in FINAL_WEIGHTS.items():

            score = item.get(
                f"score_{days}"
            )

            if score is None:
                continue

            try:

                score = float(score)

                if math.isfinite(score):

                    available_scores.append(
                        (
                            days,
                            weight,
                            score,
                        )
                    )

            except Exception:
                continue

        # ----------------------------------------------------
        # V7: Minimum 3 vade
        # ----------------------------------------------------

        if len(available_scores) < MIN_FINAL_PERIODS:

            item["final_score"] = None

            item["final_decision"] = (
                "YETERSİZ VERİ"
            )

            item["confidence"] = (
                calculate_confidence(item)
            )

            continue

        # ----------------------------------------------------
        # Ağırlıklı skor
        # ----------------------------------------------------

        total_weight = sum(
            weight
            for _, weight, _
            in available_scores
        )

        weighted_sum = sum(
            score * weight
            for _, weight, score
            in available_scores
        )

        base_final = (
            weighted_sum
            / total_weight
            if total_weight > 0
            else None
        )

        if base_final is None:

            item["final_score"] = None

            item["final_decision"] = (
                "YETERSİZ VERİ"
            )

            item["confidence"] = "DÜŞÜK"

            continue

        final_score = base_final

        # ----------------------------------------------------
        # KUT Benchmark
        # ----------------------------------------------------

        if item.get(
            "benchmark_active"
        ):

            benchmark_scores = []

            for days in PERIODS.values():

                b_score = item.get(
                    f"benchmark_score_{days}"
                )

                if b_score is None:
                    continue

                try:

                    b_score = float(
                        b_score
                    )

                    if math.isfinite(
                        b_score
                    ):
                        benchmark_scores.append(
                            b_score
                        )

                except Exception:
                    continue

            if benchmark_scores:

                benchmark_average = (
                    sum(benchmark_scores)
                    / len(benchmark_scores)
                )

                item[
                    "benchmark_average_score"
                ] = benchmark_average

                final_score = (
                    base_final * 0.80
                    + benchmark_average
                    * 0.20
                )

            else:

                item[
                    "benchmark_average_score"
                ] = None

        else:

            item[
                "benchmark_average_score"
            ] = None

        # ----------------------------------------------------
        # 0-100
        # ----------------------------------------------------

        final_score = int(
            round(
                max(
                    0.0,
                    min(
                        100.0,
                        final_score,
                    ),
                )
            )
        )

        item["final_score"] = (
            final_score
        )

        # ----------------------------------------------------
        # Nihai karar
        # ----------------------------------------------------

        if final_score >= 70:

            item["final_decision"] = (
                "GÜÇLÜ AL"
            )

        elif final_score >= 60:

            item["final_decision"] = (
                "AL"
            )

        elif final_score >= 45:

            item["final_decision"] = (
                "ASIL LİSTE"
            )

        elif final_score >= 30:

            item["final_decision"] = (
                "NÖTR"
            )

        else:

            item["final_decision"] = (
                "ACİL SAT"
            )

        item["confidence"] = (
            calculate_confidence(item)
        )


# ============================================================
# ARAYÜZ
# ============================================================

with st.sidebar:

    st.header(
        "⚙️ Filtreler & Ayarlar"
    )

    annual_rf_rate = st.number_input(
        "Yıllık Risksiz Getiri Oranı (%)",
        min_value=0.0,
        max_value=150.0,
        value=50.0,
        step=5.0,
    )

    min_aum = st.number_input(
        "Minimum Portföy Büyüklüğü (TL)",
        min_value=0.0,
        value=50_000_000.0,
        step=10_000_000.0,
    )

    min_investors = st.number_input(
        "Minimum Yatırımcı Sayısı",
        min_value=0,
        value=100,
        step=50,
    )

    use_category_scoring = st.checkbox(
        "Kategori bazlı skorlama kullan",
        value=True,
    )

    st.divider()

    st.caption(
        "V7 Skor Ağırlıkları"
    )

    st.write(
        "Getiri: %25  |  Sharpe: %25"
    )

    st.write(
        "Sortino: %25 | MDD: %25"
    )

    st.caption(
        f"Minimum nihai vade: "
        f"{MIN_FINAL_PERIODS}/4"
    )

    st.caption(
        f"Valor maksimum ceza: "
        f"{VALOR_PENALTY_MAX:.1f} puan"
    )


# ============================================================
# RISKSİZ GETİRİ
# ============================================================

daily_rf = (
    (
        1
        + annual_rf_rate / 100.0
    )
    ** (1 / 252.0)
    - 1
) * 100.0


# ============================================================
# EXCEL YÜKLEME
# ============================================================

uploaded_file = st.file_uploader(
    "Excel Dosyanızı Yükleyin "
    "(Fon_Listesi içeren):",
    type=["xlsx"],
)

if not uploaded_file:
    st.stop()


wb = openpyxl.load_workbook(
    uploaded_file
)

if "Fon_Listesi" not in wb.sheetnames:

    st.error(
        "Dosyada 'Fon_Listesi' "
        "sayfası yok!"
    )

    st.stop()


# ============================================================
# FON LİSTESİ
# ============================================================

requested_codes = []
valor_map = {}

for row in wb[
    "Fon_Listesi"
].iter_rows(
    min_row=2,
    values_only=False,
):

    if not row:
        continue

    if not row[0].value:
        continue

    code = normalize_fund_code(
        row[0].value
    )

    if not code:
        continue

    requested_codes.append(code)

    valor_map[code] = (
        parse_number(
            row[1].value
        )
        if len(row) > 1
        and row[1].value is not None
        else 0.0
    )


requested_codes = list(
    dict.fromkeys(
        requested_codes
    )
)


# ============================================================
# TARİHLER
# ============================================================

today = dt.date.today()

start_date = (
    today
    - dt.timedelta(
        days=LOOKBACK_CALENDAR_DAYS
    )
)


# ============================================================
# VERİLERİ İŞLE
# ============================================================

with st.spinner("Veriler işleniyor (Bu işlem fon sayısına göre birkaç dakika sürebilir)..."):
    
    # TEFAS'ın tüm piyasayı indirmesi kilitlenmeye yol açtığı için bu satırı KAPATIYORUZ:
    # universe = fetch_tefas_universe(start_date, today)
    
    # Yerine boş bir veri seti tanımlıyoruz. Böylece sistem doğrudan İş Yatırım'a geçecek.
    universe = pd.DataFrame() 

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_funds = len(requested_codes)

    calculated_funds, failed_codes = [], []

    for i, code in enumerate(requested_codes):
        status_text.text(f"İndiriliyor ve Analiz Ediliyor: {code} ({i+1}/{total_funds})")
        
        series = None
        source = "Bulunamadı"

        if not universe.empty:
            series = get_fund_series(universe, code)
            if series is not None:
                source = "TEFAS"

        if series is None:
            series = fetch_isyatirim_series(code)
            if series is not None:
                source = "İş Yatırım"

        metrics = compute_fund_metrics(series)
        if metrics:
            metrics.update({"code": code, "source": source, "valor": valor_map.get(code, 0.0)})
            metrics["filtered_out"] = not ((metrics["aum"] >= min_aum if metrics["aum"] > 0 else True) and (metrics["investors"] >= min_investors if metrics["investors"] > 0 else True))
            calculated_funds.append(metrics)
        else: 
            failed_codes.append(code)
            
        progress_bar.progress((i + 1) / total_funds)
        
    status_text.empty()
    progress_bar.empty()

    universe = fetch_tefas_universe(
        start_date,
        today,
    )

    calculated_funds = []
    failed_codes = []

    for code in requested_codes:

        if not universe.empty:

            series = get_fund_series(
                universe,
                code,
            )

            source = (
                "TEFAS"
                if series is not None
                else "Bulunamadı"
            )

        else:

            series = None
            source = "Bulunamadı"

        # ----------------------------------------------------
        # İş Yatırım fallback
        # ----------------------------------------------------

        if series is None:

            series = (
                fetch_isyatirim_series(
                    code
                )
            )

            if series is not None:
                source = "İş Yatırım"

        metrics = compute_fund_metrics(
            series
        )

        if metrics:

            metrics.update(
                {
                    "code": code,
                    "source": source,
                    "valor": valor_map.get(
                        code,
                        0.0,
                    ),
                }
            )

            aum_ok = (
                metrics["aum"]
                >= min_aum
                if metrics["aum"] > 0
                else True
            )

            investor_ok = (
                metrics["investors"]
                >= min_investors
                if metrics["investors"] > 0
                else True
            )

            metrics[
                "filtered_out"
            ] = not (
                aum_ok
                and investor_ok
            )

            calculated_funds.append(
                metrics
            )

        else:

            failed_codes.append(code)


# ============================================================
# UYARILAR
# ============================================================

if failed_codes:

    st.warning(
        "Veri bulunamayan fonlar: "
        + ", ".join(failed_codes)
    )


# ============================================================
# VADE SKORLARI
# ============================================================

for period_days in PERIODS.values():

    calculate_period_scores(
        calculated_funds,
        period_days,
        daily_rf,
        use_category_scoring,
    )


# ============================================================
# KUT BENCHMARK
# ============================================================

if any(
    item["code"] == "KUT"
    for item in calculated_funds
):

    kut_benchmark = (
        fetch_kut_benchmark(
            start_date,
            today,
        )
    )

    calculate_kut_benchmark_metrics(
        calculated_funds,
        kut_benchmark,
    )


# ============================================================
# TREND + NİHAİ SKOR
# ============================================================

calculate_trends(
    calculated_funds
)

calculate_final_scores(
    calculated_funds
)


# ============================================================
# SIRALAMA
# ============================================================

calculated_funds.sort(
    key=lambda x: (
        x.get("final_score")
        if x.get("final_score")
        is not None
        else -1
    ),
    reverse=True,
)


# ============================================================
# EXCEL ÇIKTISI
# ============================================================

if "Vade_Analizi" in wb.sheetnames:

    del wb["Vade_Analizi"]


ws_out = wb.create_sheet(
    "Vade_Analizi",
    0,
)


headers = [
    "Fon Kodu",
    "Kategori",
    "AUM (TL)",
    "Yatırımcı",

    "1H Skor",
    "1H Karar",
    "1H Küm %",
    "1H MDD %",
    "1H Valor Ceza",

    "1A Skor",
    "1A Karar",
    "1A Küm %",
    "1A MDD %",
    "1A Valor Ceza",

    "3A Skor",
    "3A Karar",
    "3A Küm %",
    "3A MDD %",
    "3A Valor Ceza",

    "1Y Skor",
    "1Y Karar",
    "1Y Küm %",
    "1Y MDD %",
    "1Y Valor Ceza",

    "1Y Trend",
    "3A Trend",
    "Güven Seviyesi",

    "Nihai Skor",
    "Nihai Karar",

    "Benchmark 1H",
    "KUT-Bench 1H",

    "Benchmark 1A",
    "KUT-Bench 1A",

    "Benchmark 3A",
    "KUT-Bench 3A",

    "Benchmark 1Y",
    "KUT-Bench 1Y",

    "Benchmark Ort. Skor",

    "Kaynak",
]


ws_out.append(headers)


# ============================================================
# EXCEL BAŞLIK
# ============================================================

for cell in ws_out[1]:

    cell.fill = PatternFill(
        start_color=COLOR_NAVY,
        fill_type="solid",
    )

    cell.font = Font(
        color=COLOR_WHITE,
        bold=True,
    )

    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )


def fmt(val):

    return (
        val
        if val is not None
        else "-"
    )


# ============================================================
# EXCEL SATIRLARI
# ============================================================

for item in calculated_funds:

    b_scores = [
        item.get(
            f"benchmark_score_{d}"
        )
        for d in PERIODS.values()
        if item.get(
            f"benchmark_score_{d}"
        ) is not None
    ]

    benchmark_avg = (
        round(
            sum(b_scores)
            / len(b_scores),
            1,
        )
        if b_scores
        else "-"
    )

    ws_out.append(
        [
            item["code"],
            item.get(
                "category",
                "-",
            ),
            fmt(
                item.get("aum")
            ),
            fmt(
                item.get(
                    "investors"
                )
            ),

            fmt(
                item.get("score_5")
            ),
            item.get(
                "karar_5",
                "-",
            ),
            fmt(
                item.get("cum_5")
            ),
            fmt(
                item.get("mdd_5")
            ),
            fmt(
                item.get(
                    "valor_penalty_5"
                )
            ),

            fmt(
                item.get("score_21")
            ),
            item.get(
                "karar_21",
                "-",
            ),
            fmt(
                item.get("cum_21")
            ),
            fmt(
                item.get("mdd_21")
            ),
            fmt(
                item.get(
                    "valor_penalty_21"
                )
            ),

            fmt(
                item.get("score_63")
            ),
            item.get(
                "karar_63",
                "-",
            ),
            fmt(
                item.get("cum_63")
            ),
            fmt(
                item.get("mdd_63")
            ),
            fmt(
                item.get(
                    "valor_penalty_63"
                )
            ),

            fmt(
                item.get("score_252")
            ),
            item.get(
                "karar_252",
                "-",
            ),
            fmt(
                item.get("cum_252")
            ),
            fmt(
                item.get("mdd_252")
            ),
            fmt(
                item.get(
                    "valor_penalty_252"
                )
            ),

            item.get(
                "trend_1y",
                "-",
            ),

            item.get(
                "trend_3a",
                "-",
            ),

            item.get(
                "confidence",
                "-",
            ),

            fmt(
                item.get(
                    "final_score"
                )
            ),

            item.get(
                "final_decision",
                "-",
            ),

            fmt(
                item.get(
                    "benchmark_5"
                )
            ),

            fmt(
                item.get(
                    "benchmark_diff_5"
                )
            ),

            fmt(
                item.get(
                    "benchmark_21"
                )
            ),

            fmt(
                item.get(
                    "benchmark_diff_21"
                )
            ),

            fmt(
                item.get(
                    "benchmark_63"
                )
            ),

            fmt(
                item.get(
                    "benchmark_diff_63"
                )
            ),

            fmt(
                item.get(
                    "benchmark_252"
                )
            ),

            fmt(
                item.get(
                    "benchmark_diff_252"
                )
            ),

            benchmark_avg,

            item.get(
                "source",
                "-",
            ),
        ]
    )


# ============================================================
# EXCEL RENKLERİ
# ============================================================

green = Font(
    color=COLOR_GREEN,
    bold=True,
)

red = Font(
    color=COLOR_RED,
    bold=True,
)

yellow = Font(
    color=COLOR_YELLOW,
    bold=True,
)


decision_columns = [
    6,
    11,
    16,
    21,
    29,
]


for row in range(
    2,
    ws_out.max_row + 1,
):

    # Kararlar
    for col in decision_columns:

        v = str(
            ws_out.cell(
                row=row,
                column=col,
            ).value
        )

        if any(
            x in v
            for x in [
                "GÜÇLÜ AL",
                "AL",
                "LİSTE",
            ]
        ):

            ws_out.cell(
                row=row,
                column=col,
            ).font = green

        elif "NÖTR" in v:

            ws_out.cell(
                row=row,
                column=col,
            ).font = yellow

        elif any(
            x in v
            for x in [
                "SAT",
                "ACİL",
                "FİLTRE",
            ]
        ):

            ws_out.cell(
                row=row,
                column=col,
            ).font = red

    # Trend
    for col in [25, 26]:

        v = str(
            ws_out.cell(
                row=row,
                column=col,
            ).value
        )

        if "YUKARI" in v:

            ws_out.cell(
                row=row,
                column=col,
            ).font = green

        elif "AŞAĞI" in v:

            ws_out.cell(
                row=row,
                column=col,
            ).font = red

        elif "YATAY" in v:

            ws_out.cell(
                row=row,
                column=col,
            ).font = yellow

    # Güven
    c_conf = ws_out.cell(
        row=row,
        column=27,
    )

    if c_conf.value in [
        "ÇOK YÜKSEK",
        "YÜKSEK",
    ]:

        c_conf.font = green

    elif c_conf.value == "ORTA":

        c_conf.font = yellow

    else:

        c_conf.font = red

    # Nihai skor
    c_fin = ws_out.cell(
        row=row,
        column=28,
    )

    if isinstance(
        c_fin.value,
        (int, float),
    ):

        if c_fin.value >= 70:

            c_fin.fill = PatternFill(
                start_color=COLOR_LIGHT_GREEN,
                fill_type="solid",
            )

        elif c_fin.value >= 45:

            c_fin.fill = PatternFill(
                start_color=COLOR_LIGHT_YELLOW,
                fill_type="solid",
            )

        else:

            c_fin.fill = PatternFill(
                start_color=COLOR_LIGHT_RED,
                fill_type="solid",
            )


# ============================================================
# EXCEL SAYI FORMATLARI
# ============================================================

percentage_columns = [
    7,
    8,
    9,

    12,
    13,
    14,

    17,
    18,
    19,

    22,
    23,
    24,

    30,
    31,
    32,
    33,
    34,
    35,
    36,
    37,
]


for row in range(
    2,
    ws_out.max_row + 1,
):

    for col in percentage_columns:

        cell = ws_out.cell(
            row=row,
            column=col,
        )

        if isinstance(
            cell.value,
            (int, float),
        ):

            cell.number_format = (
                '0.00"%"'
            )


# ============================================================
# KOLON GENİŞLİKLERİ
# ============================================================

for col in ws_out.columns:

    max_len = 0

    for cell in col:

        try:
            max_len = max(
                max_len,
                len(str(cell.value)),
            )
        except Exception:
            pass

    letter = get_column_letter(
        col[0].column
    )

    ws_out.column_dimensions[
        letter
    ].width = min(
        max(max_len + 2, 11),
        26,
    )


ws_out.freeze_panes = "B2"

ws_out.auto_filter.ref = (
    ws_out.dimensions
)


# ============================================================
# DOSYAYI OLUŞTUR
# ============================================================

output = io.BytesIO()

wb.save(output)

output.seek(0)


# ============================================================
# STREAMLIT SONUÇ
# ============================================================

st.success(
    "✅ Multi-Vade V7 analizi "
    "tamamlandı!"
)

st.download_button(
    label="📥 V7 Excel Çıktısını İndir",
    data=output,
    file_name=(
        "fon_vade_analizi_V7.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-"
        "officedocument.spreadsheetml.sheet"
    ),
)


# ============================================================
# ÖZET İSTATİSTİKLER
# ============================================================

total_funds = len(
    calculated_funds
)

final_scored = sum(
    1
    for f in calculated_funds
    if f.get("final_score")
    is not None
)

strong_buy = sum(
    1
    for f in calculated_funds
    if f.get("final_decision")
    == "GÜÇLÜ AL"
)

buy = sum(
    1
    for f in calculated_funds
    if f.get("final_decision")
    == "AL"
)

insufficient = sum(
    1
    for f in calculated_funds
    if f.get("final_score")
    is None
)


col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Toplam Fon",
        total_funds,
    )

with col2:
    st.metric(
        "Nihai Skor Alan",
        final_scored,
    )

with col3:
    st.metric(
        "Güçlü AL",
        strong_buy,
    )

with col4:
    st.metric(
        "Yetersiz Veri",
        insufficient,
    )


# ============================================================
# ÖNİZLEME
# ============================================================

st.subheader(
    "📊 Fon Sıralaması Önizleme"
)


preview_rows = []

for f in calculated_funds:

    preview_rows.append(
        {
            "Fon": f["code"],
            "Kategori": f.get(
                "category",
                "-",
            ),
            "1H": fmt(
                f.get("score_5")
            ),
            "1A": fmt(
                f.get("score_21")
            ),
            "3A": fmt(
                f.get("score_63")
            ),
            "1Y": fmt(
                f.get("score_252")
            ),
            "Nihai Skor": fmt(
                f.get("final_score")
            ),
            "Nihai Karar": f.get(
                "final_decision",
                "-",
            ),
            "Güven": f.get(
                "confidence",
                "-",
            ),
            "Valor": fmt(
                f.get("valor")
            ),
            "AUM": fmt(
                f.get("aum")
            ),
        }
    )


df_preview = pd.DataFrame(
    preview_rows
)


st.dataframe(
    df_preview,
    use_container_width=True,
    hide_index=True,
)
