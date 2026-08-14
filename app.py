import datetime as dt
import io
import math
from typing import Optional

import openpyxl
import pandas as pd
import requests
import streamlit as st

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# AYARLAR
# ============================================================

st.set_page_config(
    page_title="Multi-Vade Fon Analizi V2",
    page_icon="📈",
    layout="wide",
)

st.title("📈 Multi-Vade Fon Analizi V2")
st.caption(
    "5 Gün / 21 Gün / 63 Gün / 252 Gün + Nihai Skor + Trend + Güven + KUT Benchmark Analizi"
)


# ============================================================
# GENEL AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")

# 252 işlem gününü rahatça karşılamak için 400 takvim günü
LOOKBACK_CALENDAR_DAYS = 400

HTTP_TIMEOUT = 15

PERIODS = {
    "Kısa Vade (1 Hafta)": 5,
    "Orta Vade (1 Ay)": 21,
    "Uzun Vade (3 Ay)": 63,
    "Çok Uzun Vade (1 Yıl)": 252,
}

PERIOD_LABELS = {
    "Kısa Vade (1 Hafta)": "1H",
    "Orta Vade (1 Ay)": "1A",
    "Uzun Vade (3 Ay)": "3A",
    "Çok Uzun Vade (1 Yıl)": "1Y",
}

MAX_DAYS = max(PERIODS.values())


# ============================================================
# NİHAİ SKOR AĞIRLIKLARI
# ============================================================

FINAL_WEIGHTS = {
    5: 0.10,    # 1 Hafta
    21: 0.20,   # 1 Ay
    63: 0.30,   # 3 Ay
    252: 0.40,  # 1 Yıl
}


# ============================================================
# KUT BENCHMARK AĞIRLIKLARI
# ============================================================

# KUT için:
# %45 Altın
# %45 Gümüş
# %10 nötr/likit bileşen
#
# %10'luk bölüm fiyat serisi gerektirmediği için benchmark
# getirisine doğrudan 0 katkı yapmaktadır.

KUT_GOLD_WEIGHT = 0.45
KUT_SILVER_WEIGHT = 0.45
KUT_CASH_WEIGHT = 0.10


# ============================================================
# RENKLER
# ============================================================

COLOR_NAVY = "1F4E79"
COLOR_BLUE = "5B9BD5"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"

COLOR_LIGHT_GREEN = "E2F0D9"
COLOR_LIGHT_RED = "FCE4D6"
COLOR_LIGHT_YELLOW = "FFF2CC"
COLOR_LIGHT_BLUE = "DDEBF7"
COLOR_GRAY = "D9E1F2"


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


# ============================================================
# Z-SKOR
# ============================================================

def zscore(values):
    if not values:
        return []

    clean = [
        float(v)
        for v in values
        if v is not None and pd.notna(v)
    ]

    if not clean:
        return [0.0] * len(values)

    mean_val = sum(clean) / len(clean)

    variance = sum(
        (v - mean_val) ** 2
        for v in clean
    ) / len(clean)

    std = variance ** 0.5

    if std < 1e-12:
        return [0.0] * len(values)

    return [
        (v - mean_val) / std
        for v in values
    ]


# ============================================================
# TEFAS VERİSİ
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(
    start_date: dt.date,
    end_date: dt.date
) -> pd.DataFrame:

    try:
        from pytefas import Crawler

    except ImportError:
        return pd.DataFrame()

    try:

        crawler = Crawler(
            timeout=60,
            max_retry=5
        )

        df = crawler.fetch_many(
            start=start_date,
            end=end_date,
            kinds=FUND_KINDS,
            columns="info"
        )

        if df is None or df.empty:
            return pd.DataFrame()

        df = df.copy()

        df.rename(
            columns={
                "fund_code": "code",
                "fund_name": "title",
                "investor_count": "investors",
                "portfolio_size": "aum",
            },
            inplace=True
        )

        required = [
            "date",
            "code",
            "price"
        ]

        if not all(
            c in df.columns
            for c in required
        ):
            return pd.DataFrame()

        df["code"] = (
            df["code"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df["date"] = pd.to_datetime(
            df["date"],
            errors="coerce"
        )

        df["price"] = df["price"].apply(
            parse_number
        )

        if "aum" not in df.columns:
            df["aum"] = 0.0
        else:
            df["aum"] = df["aum"].apply(
                parse_number
            )

        if "investors" not in df.columns:
            df["investors"] = 0.0
        else:
            df["investors"] = df["investors"].apply(
                parse_number
            )

        df = df.dropna(
            subset=[
                "date",
                "code",
                "price"
            ]
        )

        df = df[
            df["price"] > 0
        ]

        df = (
            df
            .sort_values(
                ["code", "date"]
            )
            .drop_duplicates(
                subset=[
                    "code",
                    "date"
                ],
                keep="last"
            )
            .reset_index(drop=True)
        )

        return df

    except Exception:
        return pd.DataFrame()


# ============================================================
# FON SERİSİ
# ============================================================

def get_fund_series(
    universe: pd.DataFrame,
    fund_code: str
) -> Optional[pd.DataFrame]:

    if universe is None or universe.empty:
        return None

    code = normalize_fund_code(
        fund_code
    )

    rows = universe[
        universe["code"]
        .astype(str)
        .str.upper()
        .eq(code)
    ].copy()

    if rows.empty:
        return None

    rows = (
        rows
        .sort_values("date")
        .drop_duplicates(
            subset=["date"],
            keep="last"
        )
    )

    if len(rows) < 2:
        return None

    # 252 gün için yeterli geçmiş
    if len(rows) > MAX_DAYS + 1:
        rows = rows.tail(
            MAX_DAYS + 1
        )

    return rows.reset_index(
        drop=True
    )


# ============================================================
# İŞ YATIRIM FALLBACK
# ============================================================

def fetch_isyatirim_series(
    fund_code: str
) -> Optional[pd.DataFrame]:

    code = normalize_fund_code(
        fund_code
    )

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
        "baslangic": start.strftime(
            "%d-%m-%Y"
        ),
        "bitis": end.strftime(
            "%d-%m-%Y"
        ),
    }

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:

        response = requests.get(
            url,
            params=params,
            headers=headers,
            timeout=HTTP_TIMEOUT
        )

        response.raise_for_status()

        values = response.json().get(
            "value"
        )

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
            errors="coerce"
        )

        df["price"] = (
            df["Fiyat"]
            .apply(parse_number)
        )

        df = df.dropna(
            subset=[
                "date",
                "price"
            ]
        )

        df = df[
            df["price"] > 0
        ]

        if len(df) < 2:
            return None

        df = (
            df
            .sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last"
            )
            .tail(MAX_DAYS + 1)
            .reset_index(drop=True)
        )

        return df[
            [
                "date",
                "price"
            ]
        ]

    except Exception:
        return None


# ============================================================
# FON METRİKLERİ
# ============================================================

def compute_fund_metrics(
    series: Optional[pd.DataFrame]
) -> Optional[dict]:

    if series is None or len(series) < 2:
        return None

    df = series.copy()

    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce"
    )

    df["price"] = (
        df["price"]
        .apply(parse_number)
    )

    df = df.dropna(
        subset=[
            "date",
            "price"
        ]
    )

    df = df[
        df["price"] > 0
    ]

    df = (
        df
        .sort_values("date")
        .drop_duplicates(
            subset=["date"],
            keep="last"
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

    daily_returns = [
        (
            current / previous - 1
        ) * 100
        if previous > 0
        else 0.0
        for previous, current
        in zip(
            prices[:-1],
            prices[1:]
        )
    ]

    return {
        "dates": dates[1:],
        "prices": prices,
        "daily_returns": daily_returns,
        "n_days": len(daily_returns),
    }


# ============================================================
# VADE SKORU
# ============================================================

def calculate_period_scores(
    funds,
    days
):

    means = []
    sharpes = []
    cums = []

    valid_indices = []

    for idx, item in enumerate(funds):

        if len(
            item["daily_returns"]
        ) < days:

            item[
                f"score_{days}"
            ] = 0

            item[
                f"karar_{days}"
            ] = "Yetersiz Veri"

            continue

        slice_ret = (
            item["daily_returns"][-days:]
        )

        slice_prices = (
            item["prices"][-(days + 1):]
        )

        mean_ret = (
            sum(slice_ret)
            / len(slice_ret)
        )

        variance = sum(
            (r - mean_ret) ** 2
            for r in slice_ret
        ) / len(slice_ret)

        vol = variance ** 0.5

        sharpe = (
            mean_ret / vol
            if vol > 1e-12
            else 0.0
        )

        cumulative = (
            slice_prices[-1]
            / slice_prices[0]
            - 1
        ) * 100

        item[
            f"mean_{days}"
        ] = mean_ret

        item[
            f"vol_{days}"
        ] = vol

        item[
            f"shp_{days}"
        ] = sharpe

        item[
            f"cum_{days}"
        ] = cumulative

        means.append(mean_ret)
        sharpes.append(sharpe)
        cums.append(cumulative)

        valid_indices.append(idx)

    z_mean = zscore(means)
    z_sharpe = zscore(sharpes)
    z_cum = zscore(cums)

    for zi, idx in enumerate(
        valid_indices
    ):

        item = funds[idx]

        valor = safe_float(
            item.get(
                "valor",
                0
            )
        )

        # Valör cezası
        valor_penalty = (
            valor * 0.5
            if valor > 0
            else 0
        )

        raw_score = (
            50
            + 15 * z_mean[zi]
            + 20 * z_sharpe[zi]
            + 15 * z_cum[zi]
            - valor_penalty
        )

        score = int(
            round(
                max(
                    0,
                    min(
                        100,
                        raw_score
                    )
                )
            )
        )

        item[
            f"score_{days}"
        ] = score

        if score >= 60:
            decision = "GÜÇLÜ AL"

        elif score >= 40:
            decision = "ASIL LİSTE"

        elif score >= 25:
            decision = "NÖTR"

        else:
            decision = "ACİL SAT"

        item[
            f"karar_{days}"
        ] = decision


# ============================================================
# YAHOO FİNANCE TARİHSEL VERİSİ
# ============================================================

@st.cache_data(
    show_spinner=False,
    ttl=60 * 30
)
def fetch_yahoo_series(
    symbol: str,
    start_date: dt.date,
    end_date: dt.date
) -> Optional[pd.DataFrame]:

    try:

        period1 = int(
            dt.datetime.combine(
                start_date,
                dt.time.min
            ).timestamp()
        )

        period2 = int(
            dt.datetime.combine(
                end_date + dt.timedelta(days=1),
                dt.time.min
            ).timestamp()
        )

        url = (
            "https://query1.finance.yahoo.com/"
            "v8/finance/chart/"
            + symbol
        )

        params = {
            "period1": period1,
            "period2": period2,
            "interval": "1d",
            "events": "history",
            "includeAdjustedClose": "true",
        }

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        response = requests.get(
            url,
            params=params,
            headers=headers,
            timeout=HTTP_TIMEOUT
        )

        response.raise_for_status()

        data = response.json()

        result = (
            data
            .get("chart", {})
            .get("result")
        )

        if not result:
            return None

        result = result[0]

        timestamps = result.get(
            "timestamp"
        )

        indicators = (
            result
            .get("indicators", {})
            .get("quote", [])
        )

        if not timestamps or not indicators:
            return None

        closes = indicators[0].get(
            "close"
        )

        if not closes:
            return None

        rows = []

        for timestamp, close in zip(
            timestamps,
            closes
        ):

            if close is None:
                continue

            try:

                date = dt.datetime.fromtimestamp(
                    timestamp
                ).date()

                price = float(close)

                if price <= 0:
                    continue

                rows.append(
                    {
                        "date": pd.Timestamp(
                            date
                        ),
                        "price": price,
                    }
                )

            except Exception:
                continue

        if len(rows) < 2:
            return None

        df = pd.DataFrame(rows)

        df = (
            df
            .sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last"
            )
            .reset_index(drop=True)
        )

        return df

    except Exception:
        return None


# ============================================================
# KUT BENCHMARK SERİSİ
# ============================================================

@st.cache_data(
    show_spinner=False,
    ttl=60 * 30
)
def fetch_kut_benchmark(
    start_date: dt.date,
    end_date: dt.date
) -> Optional[pd.DataFrame]:

    gold = fetch_yahoo_series(
        "GC=F",
        start_date,
        end_date
    )

    silver = fetch_yahoo_series(
        "SI=F",
        start_date,
        end_date
    )

    usdtry = fetch_yahoo_series(
        "USDTRY=X",
        start_date,
        end_date
    )

    if (
        gold is None
        or silver is None
        or usdtry is None
    ):
        return None

    gold = gold.rename(
        columns={
            "price": "gold_usd"
        }
    )

    silver = silver.rename(
        columns={
            "price": "silver_usd"
        }
    )

    usdtry = usdtry.rename(
        columns={
            "price": "usdtry"
        }
    )

    df = gold.merge(
        silver,
        on="date",
        how="inner"
    )

    df = df.merge(
        usdtry,
        on="date",
        how="inner"
    )

    if df.empty:
        return None

    # USD/ons -> TL/ons
    df["gold_try"] = (
        df["gold_usd"]
        * df["usdtry"]
    )

    df["silver_try"] = (
        df["silver_usd"]
        * df["usdtry"]
    )

    # %45 Altın + %45 Gümüş
    # %10 nakit/nötr
    #
    # Benchmark'ın seviyesi önemli değil.
    # Getiri serisinin doğru olması önemli.
    #
    # Ağırlıklı endeks oluşturuyoruz.

    gold_base = (
        df["gold_try"]
        / df["gold_try"].iloc[0]
    )

    silver_base = (
        df["silver_try"]
        / df["silver_try"].iloc[0]
    )

    df["benchmark"] = (
        KUT_GOLD_WEIGHT * gold_base
        + KUT_SILVER_WEIGHT * silver_base
        + KUT_CASH_WEIGHT * 1.0
    )

    df = df[
        [
            "date",
            "benchmark",
            "gold_try",
            "silver_try",
            "usdtry",
        ]
    ]

    return df.reset_index(
        drop=True
    )


# ============================================================
# BENCHMARK DÖNEM GETİRİSİ
# ============================================================

def benchmark_period_return(
    benchmark_df,
    days
):

    if (
        benchmark_df is None
        or benchmark_df.empty
    ):
        return None

    if len(benchmark_df) <= days:
        return None

    values = (
        benchmark_df["benchmark"]
        .astype(float)
        .tolist()
    )

    start = values[-(days + 1)]
    end = values[-1]

    if start <= 0:
        return None

    return (
        end / start - 1
    ) * 100


# ============================================================
# KUT BENCHMARK ANALİZİ
# ============================================================

def calculate_kut_benchmark_metrics(
    funds,
    benchmark_df
):

    benchmark_returns = {}

    for days in PERIODS.values():

        benchmark_returns[
            days
        ] = benchmark_period_return(
            benchmark_df,
            days
        )

    for item in funds:

        if item["code"] != "KUT":
            item["benchmark_active"] = False
            continue

        item["benchmark_active"] = True

        for days in PERIODS.values():

            fund_return = item.get(
                f"cum_{days}"
            )

            bench_return = benchmark_returns.get(
                days
            )

            if (
                fund_return is None
                or bench_return is None
            ):
                item[
                    f"benchmark_{days}"
                ] = None

                item[
                    f"benchmark_diff_{days}"
                ] = None

                item[
                    f"benchmark_score_{days}"
                ] = 50

                continue

            diff = (
                fund_return
                - bench_return
            )

            # Relative benchmark score
            #
            # +10 puan fark -> yaklaşık 100
            #  0 puan fark   -> 50
            # -10 puan fark -> yaklaşık 0

            relative_score = (
                50 + diff * 5
            )

            relative_score = max(
                0,
                min(
                    100,
                    relative_score
                )
            )

            item[
                f"benchmark_{days}"
            ] = bench_return

            item[
                f"benchmark_diff_{days}"
            ] = diff

            item[
                f"benchmark_score_{days}"
            ] = int(
                round(
                    relative_score
                )
            )


# ============================================================
# TREND HESAPLAMA
# ============================================================

def calculate_trends(funds):

    for item in funds:

        score_63 = item.get(
            "score_63"
        )

        score_252 = item.get(
            "score_252"
        )

        score_21 = item.get(
            "score_21"
        )

        # ----------------------------------------------------
        # 3A TREND
        # ----------------------------------------------------

        if (
            score_63 is None
            or score_63 == 0
            or score_21 is None
        ):

            trend_3a = "VERİ YOK"

        else:

            difference = (
                score_21
                - score_63
            )

            if difference >= 10:
                trend_3a = "YUKARI ↑"

            elif difference <= -10:
                trend_3a = "AŞAĞI ↓"

            else:
                trend_3a = "YATAY →"

        # ----------------------------------------------------
        # 1Y TREND
        # ----------------------------------------------------

        if (
            score_252 is None
            or score_252 == 0
            or score_63 is None
        ):

            trend_1y = "VERİ YOK"

        else:

            difference = (
                score_63
                - score_252
            )

            if difference >= 10:
                trend_1y = "YUKARI ↑"

            elif difference <= -10:
                trend_1y = "AŞAĞI ↓"

            else:
                trend_1y = "YATAY →"

        item["trend_3a"] = trend_3a
        item["trend_1y"] = trend_1y


# ============================================================
# GÜVEN SEVİYESİ
# ============================================================

def calculate_confidence(
    item
):

    available = 0

    total = 0

    for days in PERIODS.values():

        total += 1

        score = item.get(
            f"score_{days}"
        )

        if (
            score is not None
            and score != 0
        ):
            available += 1

    if total == 0:
        return "DÜŞÜK"

    data_ratio = (
        available / total
    )

    scores = []

    for days in PERIODS.values():

        score = item.get(
            f"score_{days}"
        )

        if score is not None and score != 0:
            scores.append(
                float(score)
            )

    if not scores:
        return "DÜŞÜK"

    mean_score = (
        sum(scores)
        / len(scores)
    )

    dispersion = (
        sum(
            (x - mean_score) ** 2
            for x in scores
        )
        / len(scores)
    ) ** 0.5

    # Düşük dağılım = daha yüksek güven
    if (
        data_ratio >= 1.0
        and dispersion <= 10
    ):
        return "ÇOK YÜKSEK"

    if (
        data_ratio >= 0.75
        and dispersion <= 15
    ):
        return "YÜKSEK"

    if data_ratio >= 0.50:
        return "ORTA"

    return "DÜŞÜK"


# ============================================================
# NİHAİ SKOR
# ============================================================

def calculate_final_scores(
    funds
):

    for item in funds:

        weighted_sum = 0.0
        weight_sum = 0.0

        for days, weight in FINAL_WEIGHTS.items():

            score = item.get(
                f"score_{days}"
            )

            if (
                score is None
                or score == 0
            ):
                continue

            weighted_sum += (
                score * weight
            )

            weight_sum += weight

        if weight_sum <= 0:

            item["final_score"] = 0
            item["final_decision"] = (
                "YETERSİZ VERİ"
            )

            item["confidence"] = "DÜŞÜK"

            continue

        base_final_score = (
            weighted_sum
            / weight_sum
        )

        # ----------------------------------------------------
        # KUT için benchmark düzeltmesi
        # ----------------------------------------------------

        if item.get(
            "benchmark_active",
            False
        ):

            benchmark_scores = []

            for days in PERIODS.values():

                b_score = item.get(
                    f"benchmark_score_{days}"
                )

                if b_score is not None:
                    benchmark_scores.append(
                        b_score
                    )

            if benchmark_scores:

                benchmark_avg = (
                    sum(benchmark_scores)
                    / len(benchmark_scores)
                )

                # Benchmark'ın nihai skora etkisi %20
                final_score = (
                    base_final_score * 0.80
                    + benchmark_avg * 0.20
                )

            else:

                final_score = (
                    base_final_score
                )

        else:

            final_score = (
                base_final_score
            )

        final_score = int(
            round(
                max(
                    0,
                    min(
                        100,
                        final_score
                    )
                )
            )
        )

        item["final_score"] = final_score

        # ----------------------------------------------------
        # Nihai karar
        # ----------------------------------------------------

        if final_score >= 70:
            decision = "GÜÇLÜ AL"

        elif final_score >= 60:
            decision = "AL"

        elif final_score >= 45:
            decision = "ASIL LİSTE"

        elif final_score >= 30:
            decision = "NÖTR"

        else:
            decision = "ACİL SAT"

        item[
            "final_decision"
        ] = decision

        item[
            "confidence"
        ] = calculate_confidence(
            item
        )


# ============================================================
# EXCEL
# ============================================================

uploaded_file = st.file_uploader(
    "Excel Dosyanızı Yükleyin (Fon_Listesi içeren):",
    type=["xlsx"]
)

if not uploaded_file:
    st.stop()


# ============================================================
# EXCEL AÇ
# ============================================================

wb = openpyxl.load_workbook(
    uploaded_file
)

if "Fon_Listesi" not in wb.sheetnames:

    st.error(
        "Dosyada 'Fon_Listesi' sayfası yok!"
    )

    st.stop()


ws_list = wb["Fon_Listesi"]


# ============================================================
# FONLARI OKU
# ============================================================

requested_codes = []

valor_map = {}


for row in ws_list.iter_rows(
    min_row=2,
    values_only=False
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

    requested_codes.append(
        code
    )

    # B sütunu = Valör
    valor_value = 0.0

    if len(row) > 1:

        if row[1].value is not None:

            parsed = parse_number(
                row[1].value
            )

            if parsed is not None:
                valor_value = parsed

    valor_map[code] = valor_value


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
# VERİLERİ ÇEK
# ============================================================

with st.spinner(
    "TEFAS / İş Yatırım verileri çekiliyor..."
):

    universe = fetch_tefas_universe(
        start_date,
        today
    )

    calculated_funds = []

    for code in requested_codes:

        series = None

        source = "Bulunamadı"

        # Önce TEFAS
        if not universe.empty:

            series = get_fund_series(
                universe,
                code
            )

            if series is not None:
                source = "TEFAS"

        # TEFAS yoksa İş Yatırım
        if series is None:

            series = fetch_isyatirim_series(
                code
            )

            if series is not None:
                source = "İş Yatırım"

        metrics = compute_fund_metrics(
            series
        )

        if metrics:

            metrics["code"] = code

            metrics["source"] = source

            metrics["valor"] = valor_map.get(
                code,
                0.0
            )

            calculated_funds.append(
                metrics
            )


# ============================================================
# VADE SKORLARI
# ============================================================

for period_name, period_days in PERIODS.items():

    calculate_period_scores(
        calculated_funds,
        period_days
    )


# ============================================================
# KUT BENCHMARK
# ============================================================

kut_exists = any(
    item["code"] == "KUT"
    for item in calculated_funds
)


benchmark_df = None


if kut_exists:

    with st.spinner(
        "KUT benchmark verisi çekiliyor..."
    ):

        benchmark_df = fetch_kut_benchmark(
            start_date,
            today
        )

    calculate_kut_benchmark_metrics(
        calculated_funds,
        benchmark_df
    )


# ============================================================
# TREND
# ============================================================

calculate_trends(
    calculated_funds
)


# ============================================================
# NİHAİ SKOR
# ============================================================

calculate_final_scores(
    calculated_funds
)


# ============================================================
# SONUÇLARI NİHAİ SKORA GÖRE SIRALA
# ============================================================

calculated_funds = sorted(
    calculated_funds,
    key=lambda x: x.get(
        "final_score",
        0
    ),
    reverse=True
)


# ============================================================
# EXCEL SAYFASI
# ============================================================

if "Vade_Analizi" in wb.sheetnames:

    del wb[
        "Vade_Analizi"
    ]


ws_out = wb.create_sheet(
    "Vade_Analizi",
    0
)


# ============================================================
# EXCEL BAŞLIKLARI
# ============================================================

headers = [
    "Fon Kodu",

    "1H Skor",
    "1H Karar",
    "1H Kümülatif %",

    "1A Skor",
    "1A Karar",
    "1A Kümülatif %",

    "3A Skor",
    "3A Karar",
    "3A Kümülatif %",

    "1Y Skor",
    "1Y Karar",
    "1Y Kümülatif %",

    "1Y Trend",
    "3A Trend",
    "Güven Seviyesi",

    "Nihai Skor",
    "Nihai Karar",

    "Benchmark 1H",
    "KUT-Benchmark 1H",

    "Benchmark 1A",
    "KUT-Benchmark 1A",

    "Benchmark 3A",
    "KUT-Benchmark 3A",

    "Benchmark 1Y",
    "KUT-Benchmark 1Y",

    "Benchmark Ortalama Skor",
]


ws_out.append(
    headers
)


# ============================================================
# BAŞLIK FORMAT
# ============================================================

for cell in ws_out[1]:

    cell.fill = PatternFill(
        start_color=COLOR_NAVY,
        fill_type="solid"
    )

    cell.font = Font(
        color=COLOR_WHITE,
        bold=True
    )

    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )


# ============================================================
# EXCEL VERİLERİ
# ============================================================

for item in calculated_funds:

    benchmark_scores = []

    for days in PERIODS.values():

        value = item.get(
            f"benchmark_score_{days}"
        )

        if value is not None:
            benchmark_scores.append(
                value
            )

    if benchmark_scores:

        benchmark_avg_score = round(
            sum(benchmark_scores)
            / len(benchmark_scores),
            1
        )

    else:

        benchmark_avg_score = "-"

    row = [
        item["code"],

        item.get("score_5", 0),
        item.get("karar_5", "-"),
        item.get("cum_5", 0),

        item.get("score_21", 0),
        item.get("karar_21", "-"),
        item.get("cum_21", 0),

        item.get("score_63", 0),
        item.get("karar_63", "-"),
        item.get("cum_63", 0),

        item.get("score_252", 0),
        item.get("karar_252", "-"),
        item.get("cum_252", 0),

        item.get(
            "trend_1y",
            "-"
        ),

        item.get(
            "trend_3a",
            "-"
        ),

        item.get(
            "confidence",
            "-"
        ),

        item.get(
            "final_score",
            0
        ),

        item.get(
            "final_decision",
            "-"
        ),

        # Benchmark 1H
        item.get(
            "benchmark_5",
            "-"
        ),

        item.get(
            "benchmark_diff_5",
            "-"
        ),

        # Benchmark 1A
        item.get(
            "benchmark_21",
            "-"
        ),

        item.get(
            "benchmark_diff_21",
            "-"
        ),

        # Benchmark 3A
        item.get(
            "benchmark_63",
            "-"
        ),

        item.get(
            "benchmark_diff_63",
            "-"
        ),

        # Benchmark 1Y
        item.get(
            "benchmark_252",
            "-"
        ),

        item.get(
            "benchmark_diff_252",
            "-"
        ),

        benchmark_avg_score,
    ]

    ws_out.append(
        row
    )


# ============================================================
# EXCEL KOŞULLU RENKLER
# ============================================================

green_font = Font(
    color=COLOR_GREEN,
    bold=True
)

red_font = Font(
    color=COLOR_RED,
    bold=True
)

yellow_font = Font(
    color=COLOR_YELLOW,
    bold=True
)

blue_font = Font(
    color="0000FF",
    bold=True
)


for row in range(
    2,
    ws_out.max_row + 1
):

    # --------------------------------------------------------
    # Karar sütunları
    # --------------------------------------------------------

    decision_columns = [
        3,
        6,
        9,
        12,
        18,
    ]

    for col in decision_columns:

        cell = ws_out.cell(
            row=row,
            column=col
        )

        value = str(
            cell.value
        )

        if (
            "GÜÇLÜ AL"
            in value
            or value == "AL"
            or "LİSTE" in value
        ):

            cell.font = green_font

        elif "NÖTR" in value:

            cell.font = yellow_font

        elif (
            "SAT" in value
            or "ACİL" in value
        ):

            cell.font = red_font


    # --------------------------------------------------------
    # Trend
    # --------------------------------------------------------

    trend_columns = [
        14,
        15,
    ]

    for col in trend_columns:

        cell = ws_out.cell(
            row=row,
            column=col
        )

        value = str(
            cell.value
        )

        if "YUKARI" in value:

            cell.font = green_font

        elif "AŞAĞI" in value:

            cell.font = red_font

        elif "YATAY" in value:

            cell.font = yellow_font


    # --------------------------------------------------------
    # Güven
    # --------------------------------------------------------

    confidence_cell = ws_out.cell(
        row=row,
        column=16
    )

    confidence = str(
        confidence_cell.value
    )

    if confidence in [
        "ÇOK YÜKSEK",
        "YÜKSEK"
    ]:

        confidence_cell.font = green_font

    elif confidence == "ORTA":

        confidence_cell.font = yellow_font

    else:

        confidence_cell.font = red_font


    # --------------------------------------------------------
    # Nihai skor
    # --------------------------------------------------------

    final_score_cell = ws_out.cell(
        row=row,
        column=17
    )

    final_score = safe_float(
        final_score_cell.value
    )

    if final_score >= 70:

        final_score_cell.fill = PatternFill(
            start_color=COLOR_LIGHT_GREEN,
            fill_type="solid"
        )

    elif final_score >= 45:

        final_score_cell.fill = PatternFill(
            start_color=COLOR_LIGHT_YELLOW,
            fill_type="solid"
        )

    else:

        final_score_cell.fill = PatternFill(
            start_color=COLOR_LIGHT_RED,
            fill_type="solid"
        )


# ============================================================
# YÜZDE FORMATLARI
# ============================================================

percent_columns = [
    4,
    7,
    10,
    13,
    19,
    20,
    21,
    22,
    23,
    24,
    25,
    26,
]


for row in range(
    2,
    ws_out.max_row + 1
):

    for col in percent_columns:

        ws_out.cell(
            row=row,
            column=col
        ).number_format = (
            '0.00"%"'
        )


# ============================================================
# SÜTUN GENİŞLİKLERİ
# ============================================================

for column_cells in ws_out.columns:

    column_letter = get_column_letter(
        column_cells[0].column
    )

    max_length = 0

    for cell in column_cells:

        try:

            length = len(
                str(cell.value)
            )

            if length > max_length:
                max_length = length

        except Exception:
            pass

    ws_out.column_dimensions[
        column_letter
    ].width = min(
        max(
            max_length + 2,
            12
        ),
        28
    )


# ============================================================
# DONDURMA
# ============================================================

ws_out.freeze_panes = "B2"


# ============================================================
# OTOMATİK FİLTRE
# ============================================================

ws_out.auto_filter.ref = (
    ws_out.dimensions
)


# ============================================================
# EXCEL DOSYASINI OLUŞTUR
# ============================================================

output = io.BytesIO()

wb.save(
    output
)

output.seek(0)


# ============================================================
# STREAMLIT SONUÇ
# ============================================================

st.success(
    "✅ Multi-Vade V2 analizi tamamlandı!"
)


st.download_button(
    label="📥 V2 Excel Çıktısını İndir",
    data=output,
    file_name="fon_vade_analizi_V2.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)


# ============================================================
# KUT ÖZEL ÖZET
# ============================================================

kut_items = [
    item
    for item in calculated_funds
    if item["code"] == "KUT"
]


if kut_items:

    kut = kut_items[0]

    st.subheader(
        "🥇 KUT Özel Analiz"
    )

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric(
        "Nihai Skor",
        kut.get(
            "final_score",
            0
        )
    )

    col2.metric(
        "Nihai Karar",
        kut.get(
            "final_decision",
            "-"
        )
    )

    col3.metric(
        "1Y Trend",
        kut.get(
            "trend_1y",
            "-"
        )
    )

    col4.metric(
        "3A Trend",
        kut.get(
            "trend_3a",
            "-"
        )
    )

    col5.metric(
        "Güven",
        kut.get(
            "confidence",
            "-"
        )
    )

    st.markdown(
        "### KUT vs Benchmark"
    )

    benchmark_display = pd.DataFrame(
        {
            "Dönem": [
                "1 Hafta",
                "1 Ay",
                "3 Ay",
                "1 Yıl",
            ],
            "KUT Getiri %": [
                kut.get(
                    "cum_5",
                    None
                ),
                kut.get(
                    "cum_21",
                    None
                ),
                kut.get(
                    "cum_63",
                    None
                ),
                kut.get(
                    "cum_252",
                    None
                ),
            ],
            "Benchmark %": [
                kut.get(
                    "benchmark_5",
                    None
                ),
                kut.get(
                    "benchmark_21",
                    None
                ),
                kut.get(
                    "benchmark_63",
                    None
                ),
                kut.get(
                    "benchmark_252",
                    None
                ),
            ],
            "KUT - Benchmark": [
                kut.get(
                    "benchmark_diff_5",
                    None
                ),
                kut.get(
                    "benchmark_diff_21",
                    None
                ),
                kut.get(
                    "benchmark_diff_63",
                    None
                ),
                kut.get(
                    "benchmark_diff_252",
                    None
                ),
            ],
        }
    )

    st.dataframe(
        benchmark_display,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# TÜM FONLAR ÖNİZLEME
# ============================================================

st.subheader(
    "📊 Fon Sıralaması"
)


preview_rows = []


for item in calculated_funds:

    preview_rows.append(
        {
            "Fon": item["code"],

            "1H": item.get(
                "score_5",
                0
            ),

            "1A": item.get(
                "score_21",
                0
            ),

            "3A": item.get(
                "score_63",
                0
            ),

            "1Y": item.get(
                "score_252",
                0
            ),

            "3A Trend": item.get(
                "trend_3a",
                "-"
            ),

            "1Y Trend": item.get(
                "trend_1y",
                "-"
            ),

            "Nihai Skor": item.get(
                "final_score",
                0
            ),

            "Nihai Karar": item.get(
                "final_decision",
                "-"
            ),

            "Güven": item.get(
                "confidence",
                "-"
            ),

            "Kaynak": item.get(
                "source",
                "-"
            ),
        }
    )


df_preview = pd.DataFrame(
    preview_rows
)


st.dataframe(
    df_preview,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# BENCHMARK BİLGİSİ
# ============================================================

if kut_exists:

    if benchmark_df is not None:

        st.info(
            "KUT benchmarkı aktif: "
            "%45 Altın + %45 Gümüş + %10 nötr bileşen. "
            "Altın ve gümüş serileri USD bazından USD/TRY ile "
            "TL bazına dönüştürülerek hesaplanmaktadır."
        )

    else:

        st.warning(
            "KUT bulundu ancak benchmark verisi alınamadı. "
            "Bu durumda KUT'un nihai skoru yalnızca çok-vadeli "
            "fon skorlarından hesaplanmıştır."
        )
