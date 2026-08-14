import datetime as dt
import io
import math
from typing import Optional, List, Dict, Any

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
    page_title="Multi-Vade Fon Analizi V3",
    page_icon="📈",
    layout="wide",
)

st.title("📈 Multi-Vade Fon Analizi V3")
st.caption(
    "5 / 21 / 63 / 252 Gün + MDD + Sortino + Kategori Skoru + Nihai Skor + Trend + Güven + KUT Benchmark"
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

PERIOD_LABELS = {
    "Kısa Vade (1 Hafta)": "1H",
    "Orta Vade (1 Ay)": "1A",
    "Uzun Vade (3 Ay)": "3A",
    "Çok Uzun Vade (1 Yıl)": "1Y",
}

MAX_DAYS = max(PERIODS.values())

FINAL_WEIGHTS = {
    5: 0.10,
    21: 0.20,
    63: 0.30,
    252: 0.40,
}

# KUT Benchmark
KUT_GOLD_WEIGHT = 0.45
KUT_SILVER_WEIGHT = 0.45
KUT_CASH_WEIGHT = 0.10

# Renkler
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


def zscore(values: List[float]) -> List[float]:
    if not values:
        return []
    clean = [float(v) for v in values if v is not None and pd.notna(v)]
    if not clean:
        return [0.0] * len(values)

    mean_val = sum(clean) / len(clean)
    variance = sum((v - mean_val) ** 2 for v in clean) / len(clean)
    std = variance ** 0.5
    if std < 1e-12:
        return [0.0] * len(values)
    return [(v - mean_val) / std for v in values]


def infer_category(title: str) -> str:
    """Fon adına göre basit kategori çıkarımı."""
    if not title:
        return "Diğer"
    t = title.upper()

    if any(k in t for k in ["ALTIN", "GOLD", "KIYMETLİ MADEN", "GÜMÜŞ", "SILVER", "KITA"]):
        return "Kıymetli Maden"
    if any(k in t for k in ["HİSSE", "EQUITY", "HİSSE SENEDİ", "HİSSE YOĞUN"]):
        return "Hisse Senedi"
    if any(k in t for k in ["PARA PİYASASI", "LIKIT", "LİKİT", "PARA PIYASASI"]):
        return "Para Piyasası"
    if any(k in t for k in ["BORÇLANMA", "TAHVİL", "BONO", "BORÇLANMA ARAÇLARI", "KİRA SERTİFİKASI"]):
        return "Borçlanma"
    if any(k in t for k in ["KARMA", "DEĞİŞKEN", "DENGELİ", "ÇOKLU VARLIK", "FON SEPETİ"]):
        return "Karma / Değişken"
    if any(k in t for k in ["YABANCI", "EUROBOND", "DIŞ BORÇ", "USD", "EUR", "DÖVİZ"]):
        return "Yabancı / Döviz"
    if any(k in t for k in ["SERBEST"]):
        return "Serbest"
    if any(k in t for k in ["KATILIM"]):
        return "Katılım"
    return "Diğer"


# ============================================================
# TEFAS VERİSİ
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
    except ImportError:
        st.error("`pytefas` kütüphanesi yüklü değil. `pip install pytefas` çalıştırın.")
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

        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["price"] = df["price"].apply(parse_number)

        if "aum" not in df.columns:
            df["aum"] = 0.0
        else:
            df["aum"] = df["aum"].apply(parse_number)

        if "investors" not in df.columns:
            df["investors"] = 0.0
        else:
            df["investors"] = df["investors"].apply(parse_number)

        if "title" not in df.columns:
            df["title"] = ""
        if "kind" not in df.columns:
            df["kind"] = ""

        df = df.dropna(subset=["date", "code", "price"])
        df = df[df["price"] > 0]
        df = (
            df.sort_values(["code", "date"])
            .drop_duplicates(subset=["code", "date"], keep="last")
            .reset_index(drop=True)
        )
        return df
    except Exception as e:
        st.warning(f"TEFAS verisi çekilirken hata: {e}")
        return pd.DataFrame()


def get_fund_series(universe: pd.DataFrame, fund_code: str) -> Optional[pd.DataFrame]:
    if universe is None or universe.empty:
        return None

    code = normalize_fund_code(fund_code)
    rows = universe[universe["code"].astype(str).str.upper().eq(code)].copy()
    if rows.empty:
        return None

    rows = rows.sort_values("date").drop_duplicates(subset=["date"], keep="last")
    if len(rows) < 2:
        return None

    if len(rows) > MAX_DAYS + 1:
        rows = rows.tail(MAX_DAYS + 1)

    return rows.reset_index(drop=True)


# ============================================================
# İŞ YATIRIM FALLBACK
# ============================================================

def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code:
        return None

    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

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
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
        values = response.json().get("value")
        if not values:
            return None

        df = pd.DataFrame(values)
        if "Tarih" not in df.columns or "Fiyat" not in df.columns:
            return None

        df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
        df["price"] = df["Fiyat"].apply(parse_number)
        df = df.dropna(subset=["date", "price"])
        df = df[df["price"] > 0]

        if len(df) < 2:
            return None

        df = (
            df.sort_values("date")
            .drop_duplicates(subset=["date"], keep="last")
            .tail(MAX_DAYS + 1)
            .reset_index(drop=True)
        )
        return df[["date", "price"]]
    except Exception:
        return None


# ============================================================
# FON METRİKLERİ + MDD + SORTINO
# ============================================================

def compute_max_drawdown(prices: List[float]) -> float:
    """Maksimum Drawdown (%) – negatif değer döner."""
    if not prices or len(prices) < 2:
        return 0.0
    peak = prices[0]
    max_dd = 0.0
    for p in prices:
        if p > peak:
            peak = p
        if peak > 0:
            dd = (p / peak - 1.0) * 100.0
            if dd < max_dd:
                max_dd = dd
    return max_dd  # örn. -18.5


def compute_sortino(returns: List[float]) -> float:
    """Sortino oranı (risksiz faiz = 0 varsayımı)."""
    if not returns:
        return 0.0
    mean_ret = sum(returns) / len(returns)
    downside = [r for r in returns if r < 0]
    if not downside:
        return mean_ret * 10 if mean_ret > 0 else 0.0  # downside yoksa yüksek puan
    downside_var = sum(r ** 2 for r in downside) / len(downside)
    downside_vol = math.sqrt(downside_var)
    if downside_vol < 1e-12:
        return 0.0
    return mean_ret / downside_vol


def compute_fund_metrics(series: Optional[pd.DataFrame]) -> Optional[dict]:
    if series is None or len(series) < 2:
        return None

    df = series.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["price"] = df["price"].apply(parse_number)
    df = df.dropna(subset=["date", "price"])
    df = df[df["price"] > 0]
    df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)

    if len(df) < 2:
        return None

    prices = df["price"].astype(float).tolist()
    dates = df["date"].dt.strftime("%d.%m").tolist()

    daily_returns = [
        (current / previous - 1) * 100 if previous > 0 else 0.0
        for previous, current in zip(prices[:-1], prices[1:])
    ]

    # Son AUM / yatırımcı (varsa)
    latest_aum = 0.0
    latest_investors = 0.0
    title = ""
    kind = ""

    if "aum" in df.columns:
        latest_aum = safe_float(df["aum"].iloc[-1])
    if "investors" in df.columns:
        latest_investors = safe_float(df["investors"].iloc[-1])
    if "title" in df.columns:
        title = str(df["title"].iloc[-1] or "")
    if "kind" in df.columns:
        kind = str(df["kind"].iloc[-1] or "")

    return {
        "dates": dates[1:],
        "prices": prices,
        "daily_returns": daily_returns,
        "n_days": len(daily_returns),
        "title": title,
        "kind": kind,
        "aum": latest_aum,
        "investors": latest_investors,
        "category": infer_category(title),
    }


# ============================================================
# VADE SKORU (MDD + Sortino dahil)
# ============================================================

def calculate_period_scores(funds: List[dict], days: int, use_category: bool = True):
    # Önce tüm geçerli fonlar için metrikleri hesapla
    for item in funds:
        if len(item["daily_returns"]) < days:
            item[f"score_{days}"] = 0
            item[f"karar_{days}"] = "Yetersiz Veri"
            item[f"mean_{days}"] = None
            item[f"vol_{days}"] = None
            item[f"shp_{days}"] = None
            item[f"cum_{days}"] = None
            item[f"mdd_{days}"] = None
            item[f"sortino_{days}"] = None
            continue

        slice_ret = item["daily_returns"][-days:]
        slice_prices = item["prices"][-(days + 1):]

        mean_ret = sum(slice_ret) / len(slice_ret)
        variance = sum((r - mean_ret) ** 2 for r in slice_ret) / len(slice_ret)
        vol = variance ** 0.5
        sharpe = mean_ret / vol if vol > 1e-12 else 0.0
        cumulative = (slice_prices[-1] / slice_prices[0] - 1) * 100
        mdd = compute_max_drawdown(slice_prices)
        sortino = compute_sortino(slice_ret)

        item[f"mean_{days}"] = mean_ret
        item[f"vol_{days}"] = vol
        item[f"shp_{days}"] = sharpe
        item[f"cum_{days}"] = cumulative
        item[f"mdd_{days}"] = mdd
        item[f"sortino_{days}"] = sortino

    # Kategori bazlı veya genel z-skor
    if use_category:
        categories = {}
        for idx, item in enumerate(funds):
            if item.get(f"cum_{days}") is None:
                continue
            cat = item.get("category", "Diğer")
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(idx)
    else:
        categories = {"ALL": [i for i, item in enumerate(funds) if item.get(f"cum_{days}") is not None]}

    for cat, indices in categories.items():
        if len(indices) < 2:
            # Tek fon varsa nötr skor ver
            for idx in indices:
                item = funds[idx]
                valor = safe_float(item.get("valor", 0))
                valor_penalty = valor * 1.2
                raw = 50 - valor_penalty
                score = int(round(max(0, min(100, raw))))
                item[f"score_{days}"] = score
                item[f"karar_{days}"] = _decision_from_score(score)
            continue

        means = [funds[i][f"mean_{days}"] for i in indices]
        sharpes = [funds[i][f"shp_{days}"] for i in indices]
        cums = [funds[i][f"cum_{days}"] for i in indices]
        sortinos = [funds[i][f"sortino_{days}"] for i in indices]
        mdds = [funds[i][f"mdd_{days}"] for i in indices]  # negatif değerler

        z_mean = zscore(means)
        z_sharpe = zscore(sharpes)
        z_cum = zscore(cums)
        z_sortino = zscore(sortinos)
        # MDD daha yüksek (daha az negatif) = daha iyi → z-skoru tersine çevirmeye gerek yok,
        # çünkü daha az düşüş = daha yüksek (az negatif) değer.
        z_mdd = zscore(mdds)

        for zi, idx in enumerate(indices):
            item = funds[idx]
            valor = safe_float(item.get("valor", 0))
            valor_penalty = valor * 1.2  # güçlendirilmiş ceza

            raw_score = (
                50
                + 12 * z_mean[zi]
                + 15 * z_sharpe[zi]
                + 12 * z_cum[zi]
                + 12 * z_sortino[zi]
                + 10 * z_mdd[zi]      # daha az drawdown = daha yüksek skor
                - valor_penalty
            )

            score = int(round(max(0, min(100, raw_score))))
            item[f"score_{days}"] = score
            item[f"karar_{days}"] = _decision_from_score(score)


def _decision_from_score(score: int) -> str:
    if score >= 60:
        return "GÜÇLÜ AL"
    elif score >= 40:
        return "ASIL LİSTE"
    elif score >= 25:
        return "NÖTR"
    else:
        return "ACİL SAT"


# ============================================================
# YAHOO + KUT BENCHMARK
# ============================================================

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_yahoo_series(symbol: str, start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    try:
        period1 = int(dt.datetime.combine(start_date, dt.time.min).timestamp())
        period2 = int(dt.datetime.combine(end_date + dt.timedelta(days=1), dt.time.min).timestamp())

        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
        params = {
            "period1": period1,
            "period2": period2,
            "interval": "1d",
            "events": "history",
            "includeAdjustedClose": "true",
        }
        headers = {"User-Agent": "Mozilla/5.0"}

        response = requests.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        result = data.get("chart", {}).get("result")
        if not result:
            return None
        result = result[0]
        timestamps = result.get("timestamp")
        indicators = result.get("indicators", {}).get("quote", [])
        if not timestamps or not indicators:
            return None
        closes = indicators[0].get("close")
        if not closes:
            return None

        rows = []
        for timestamp, close in zip(timestamps, closes):
            if close is None:
                continue
            try:
                date = dt.datetime.fromtimestamp(timestamp).date()
                price = float(close)
                if price <= 0:
                    continue
                rows.append({"date": pd.Timestamp(date), "price": price})
            except Exception:
                continue

        if len(rows) < 2:
            return None
        df = pd.DataFrame(rows)
        df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
        return df
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_kut_benchmark(start_date: dt.date, end_date: dt.date) -> Optional[pd.DataFrame]:
    gold = fetch_yahoo_series("GC=F", start_date, end_date)
    silver = fetch_yahoo_series("SI=F", start_date, end_date)
    usdtry = fetch_yahoo_series("USDTRY=X", start_date, end_date)

    if gold is None or silver is None or usdtry is None:
        return None

    gold = gold.rename(columns={"price": "gold_usd"})
    silver = silver.rename(columns={"price": "silver_usd"})
    usdtry = usdtry.rename(columns={"price": "usdtry"})

    df = gold.merge(silver, on="date", how="inner").merge(usdtry, on="date", how="inner")
    if df.empty:
        return None

    df["gold_try"] = df["gold_usd"] * df["usdtry"]
    df["silver_try"] = df["silver_usd"] * df["usdtry"]

    gold_base = df["gold_try"] / df["gold_try"].iloc[0]
    silver_base = df["silver_try"] / df["silver_try"].iloc[0]

    df["benchmark"] = (
        KUT_GOLD_WEIGHT * gold_base
        + KUT_SILVER_WEIGHT * silver_base
        + KUT_CASH_WEIGHT * 1.0
    )
    return df[["date", "benchmark", "gold_try", "silver_try", "usdtry"]].reset_index(drop=True)


def benchmark_period_return(benchmark_df, days):
    if benchmark_df is None or benchmark_df.empty or len(benchmark_df) <= days:
        return None
    values = benchmark_df["benchmark"].astype(float).tolist()
    start = values[-(days + 1)]
    end = values[-1]
    if start <= 0:
        return None
    return (end / start - 1) * 100


def calculate_kut_benchmark_metrics(funds, benchmark_df):
    benchmark_returns = {days: benchmark_period_return(benchmark_df, days) for days in PERIODS.values()}

    for item in funds:
        if item["code"] != "KUT":
            item["benchmark_active"] = False
            continue

        item["benchmark_active"] = True
        for days in PERIODS.values():
            fund_return = item.get(f"cum_{days}")
            bench_return = benchmark_returns.get(days)

            if fund_return is None or bench_return is None:
                item[f"benchmark_{days}"] = None
                item[f"benchmark_diff_{days}"] = None
                item[f"benchmark_score_{days}"] = 50
                continue

            diff = fund_return - bench_return
            relative_score = max(0, min(100, 50 + diff * 5))
            item[f"benchmark_{days}"] = bench_return
            item[f"benchmark_diff_{days}"] = diff
            item[f"benchmark_score_{days}"] = int(round(relative_score))


# ============================================================
# TREND + GÜVEN + NİHAİ SKOR
# ============================================================

def calculate_trends(funds):
    for item in funds:
        score_63 = item.get("score_63")
        score_252 = item.get("score_252")
        score_21 = item.get("score_21")

        if score_63 is None or score_63 == 0 or score_21 is None:
            trend_3a = "VERİ YOK"
        else:
            difference = score_21 - score_63
            if difference >= 10:
                trend_3a = "YUKARI ↑"
            elif difference <= -10:
                trend_3a = "AŞAĞI ↓"
            else:
                trend_3a = "YATAY →"

        if score_252 is None or score_252 == 0 or score_63 is None:
            trend_1y = "VERİ YOK"
        else:
            difference = score_63 - score_252
            if difference >= 10:
                trend_1y = "YUKARI ↑"
            elif difference <= -10:
                trend_1y = "AŞAĞI ↓"
            else:
                trend_1y = "YATAY →"

        item["trend_3a"] = trend_3a
        item["trend_1y"] = trend_1y


def calculate_confidence(item):
    available = 0
    total = 0
    scores = []
    for days in PERIODS.values():
        total += 1
        score = item.get(f"score_{days}")
        if score is not None and score != 0:
            available += 1
            scores.append(float(score))

    if total == 0 or not scores:
        return "DÜŞÜK"

    data_ratio = available / total
    mean_score = sum(scores) / len(scores)
    dispersion = (sum((x - mean_score) ** 2 for x in scores) / len(scores)) ** 0.5

    if data_ratio >= 1.0 and dispersion <= 10:
        return "ÇOK YÜKSEK"
    if data_ratio >= 0.75 and dispersion <= 15:
        return "YÜKSEK"
    if data_ratio >= 0.50:
        return "ORTA"
    return "DÜŞÜK"


def calculate_final_scores(funds):
    for item in funds:
        weighted_sum = 0.0
        weight_sum = 0.0

        for days, weight in FINAL_WEIGHTS.items():
            score = item.get(f"score_{days}")
            if score is None or score == 0:
                continue
            weighted_sum += score * weight
            weight_sum += weight

        if weight_sum <= 0:
            item["final_score"] = 0
            item["final_decision"] = "YETERSİZ VERİ"
            item["confidence"] = "DÜŞÜK"
            continue

        base_final_score = weighted_sum / weight_sum

        if item.get("benchmark_active", False):
            benchmark_scores = [
                item.get(f"benchmark_score_{days}")
                for days in PERIODS.values()
                if item.get(f"benchmark_score_{days}") is not None
            ]
            if benchmark_scores:
                benchmark_avg = sum(benchmark_scores) / len(benchmark_scores)
                final_score = base_final_score * 0.80 + benchmark_avg * 0.20
            else:
                final_score = base_final_score
        else:
            final_score = base_final_score

        final_score = int(round(max(0, min(100, final_score))))
        item["final_score"] = final_score

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

        item["final_decision"] = decision
        item["confidence"] = calculate_confidence(item)


# ============================================================
# SIDEBAR AYARLARI
# ============================================================

with st.sidebar:
    st.header("⚙️ Filtreler & Ayarlar")
    min_aum = st.number_input(
        "Minimum Portföy Büyüklüğü (TL)",
        min_value=0.0,
        value=50_000_000.0,
        step=10_000_000.0,
        help="Bu değerin altındaki fonlar skorlamaya dahil edilmez (likidite riski).",
    )
    min_investors = st.number_input(
        "Minimum Yatırımcı Sayısı",
        min_value=0,
        value=100,
        step=50,
        help="Çok az yatırımcısı olan fonlar elenir.",
    )
    use_category_scoring = st.checkbox(
        "Kategori bazlı skorlama kullan",
        value=True,
        help="Açıkken z-skorlar aynı kategori içindeki fonlara göre hesaplanır (önerilir).",
    )
    st.markdown("---")
    st.info(
        "Bu araç **göreceli sıralama** yapar. "
        "Tek başına alım-satım kararı vermek için yeterli değildir. "
        "Mutlaka kendi risk analizinizi de yapın."
    )


# ============================================================
# EXCEL YÜKLEME
# ============================================================

uploaded_file = st.file_uploader(
    "Excel Dosyanızı Yükleyin (Fon_Listesi içeren):",
    type=["xlsx"],
)

if not uploaded_file:
    st.stop()

wb = openpyxl.load_workbook(uploaded_file)

if "Fon_Listesi" not in wb.sheetnames:
    st.error("Dosyada 'Fon_Listesi' sayfası yok!")
    st.stop()

ws_list = wb["Fon_Listesi"]

requested_codes = []
valor_map = {}

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if not row or not row[0].value:
        continue
    code = normalize_fund_code(row[0].value)
    if not code:
        continue
    requested_codes.append(code)

    valor_value = 0.0
    if len(row) > 1 and row[1].value is not None:
        parsed = parse_number(row[1].value)
        if parsed is not None:
            valor_value = parsed
    valor_map[code] = valor_value

requested_codes = list(dict.fromkeys(requested_codes))

# ============================================================
# VERİ ÇEKME
# ============================================================

today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

status_messages = []

with st.spinner("TEFAS / İş Yatırım verileri çekiliyor..."):
    universe = fetch_tefas_universe(start_date, today)
    if not universe.empty:
        status_messages.append(f"TEFAS’tan {universe['code'].nunique()} farklı fon verisi alındı.")
    else:
        status_messages.append("TEFAS verisi alınamadı, yalnızca İş Yatırım denenecek.")

    calculated_funds = []
    failed_codes = []

    for code in requested_codes:
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
            metrics["code"] = code
            metrics["source"] = source
            metrics["valor"] = valor_map.get(code, 0.0)

            # Filtre kontrolü
            aum_ok = metrics["aum"] >= min_aum if metrics["aum"] > 0 else True  # AUM bilinmiyorsa geçir
            inv_ok = metrics["investors"] >= min_investors if metrics["investors"] > 0 else True
            metrics["filtered_out"] = not (aum_ok and inv_ok)

            calculated_funds.append(metrics)
        else:
            failed_codes.append(code)

if failed_codes:
    st.warning(f"Veri bulunamayan fonlar: {', '.join(failed_codes)}")

for msg in status_messages:
    st.info(msg)

# Skorlanacak fonlar (filtreyi geçenler)
scoring_funds = [f for f in calculated_funds if not f.get("filtered_out", False)]
filtered_out_funds = [f for f in calculated_funds if f.get("filtered_out", False)]

if filtered_out_funds:
    st.warning(
        f"{len(filtered_out_funds)} fon AUM / yatırımcı filtresinden elendi "
        f"(skorlamaya dahil edilmedi)."
    )

# ============================================================
# SKOR HESAPLAMALARI
# ============================================================

for period_name, period_days in PERIODS.items():
    calculate_period_scores(scoring_funds, period_days, use_category=use_category_scoring)

# Filtrelenenlere düşük skor ver
for item in filtered_out_funds:
    for days in PERIODS.values():
        item[f"score_{days}"] = 0
        item[f"karar_{days}"] = "FİLTRE DIŞI"
        item[f"cum_{days}"] = item.get(f"cum_{days}")  # varsa koru

all_funds = scoring_funds + filtered_out_funds

kut_exists = any(item["code"] == "KUT" for item in all_funds)
benchmark_df = None

if kut_exists:
    with st.spinner("KUT benchmark verisi çekiliyor..."):
        benchmark_df = fetch_kut_benchmark(start_date, today)
    calculate_kut_benchmark_metrics(all_funds, benchmark_df)

calculate_trends(all_funds)
calculate_final_scores(all_funds)

# Nihai skora göre sırala
all_funds = sorted(all_funds, key=lambda x: x.get("final_score", 0), reverse=True)

# ============================================================
# EXCEL ÇIKTISI
# ============================================================

if "Vade_Analizi" in wb.sheetnames:
    del wb["Vade_Analizi"]

ws_out = wb.create_sheet("Vade_Analizi", 0)

headers = [
    "Fon Kodu", "Kategori", "AUM (TL)", "Yatırımcı",
    "1H Skor", "1H Karar", "1H Küm %", "1H MDD %",
    "1A Skor", "1A Karar", "1A Küm %", "1A MDD %",
    "3A Skor", "3A Karar", "3A Küm %", "3A MDD %",
    "1Y Skor", "1Y Karar", "1Y Küm %", "1Y MDD %",
    "1Y Trend", "3A Trend", "Güven Seviyesi",
    "Nihai Skor", "Nihai Karar",
    "Benchmark 1H", "KUT-Bench 1H",
    "Benchmark 1A", "KUT-Bench 1A",
    "Benchmark 3A", "KUT-Bench 3A",
    "Benchmark 1Y", "KUT-Bench 1Y",
    "Benchmark Ort. Skor", "Kaynak",
]

ws_out.append(headers)

for cell in ws_out[1]:
    cell.fill = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    cell.font = Font(color=COLOR_WHITE, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for item in all_funds:
    benchmark_scores = [
        item.get(f"benchmark_score_{d}")
        for d in PERIODS.values()
        if item.get(f"benchmark_score_{d}") is not None
    ]
    benchmark_avg_score = (
        round(sum(benchmark_scores) / len(benchmark_scores), 1) if benchmark_scores else "-"
    )

    row = [
        item["code"],
        item.get("category", "-"),
        item.get("aum", 0),
        item.get("investors", 0),
        item.get("score_5", 0),
        item.get("karar_5", "-"),
        item.get("cum_5", 0),
        item.get("mdd_5", 0),
        item.get("score_21", 0),
        item.get("karar_21", "-"),
        item.get("cum_21", 0),
        item.get("mdd_21", 0),
        item.get("score_63", 0),
        item.get("karar_63", "-"),
        item.get("cum_63", 0),
        item.get("mdd_63", 0),
        item.get("score_252", 0),
        item.get("karar_252", "-"),
        item.get("cum_252", 0),
        item.get("mdd_252", 0),
        item.get("trend_1y", "-"),
        item.get("trend_3a", "-"),
        item.get("confidence", "-"),
        item.get("final_score", 0),
        item.get("final_decision", "-"),
        item.get("benchmark_5", "-"),
        item.get("benchmark_diff_5", "-"),
        item.get("benchmark_21", "-"),
        item.get("benchmark_diff_21", "-"),
        item.get("benchmark_63", "-"),
        item.get("benchmark_diff_63", "-"),
        item.get("benchmark_252", "-"),
        item.get("benchmark_diff_252", "-"),
        benchmark_avg_score,
        item.get("source", "-"),
    ]
    ws_out.append(row)

# Koşullu formatlama
green_font = Font(color=COLOR_GREEN, bold=True)
red_font = Font(color=COLOR_RED, bold=True)
yellow_font = Font(color=COLOR_YELLOW, bold=True)

for row in range(2, ws_out.max_row + 1):
    # Karar sütunları
    for col in [6, 10, 14, 18, 25]:
        cell = ws_out.cell(row=row, column=col)
        value = str(cell.value)
        if "GÜÇLÜ AL" in value or value == "AL" or "LİSTE" in value:
            cell.font = green_font
        elif "NÖTR" in value:
            cell.font = yellow_font
        elif "SAT" in value or "ACİL" in value or "FİLTRE" in value:
            cell.font = red_font

    # Trend
    for col in [21, 22]:
        cell = ws_out.cell(row=row, column=col)
        value = str(cell.value)
        if "YUKARI" in value:
            cell.font = green_font
        elif "AŞAĞI" in value:
            cell.font = red_font
        elif "YATAY" in value:
            cell.font = yellow_font

    # Güven
    conf_cell = ws_out.cell(row=row, column=23)
    conf = str(conf_cell.value)
    if conf in ["ÇOK YÜKSEK", "YÜKSEK"]:
        conf_cell.font = green_font
    elif conf == "ORTA":
        conf_cell.font = yellow_font
    else:
        conf_cell.font = red_font

    # Nihai skor
    final_cell = ws_out.cell(row=row, column=24)
    final_score = safe_float(final_cell.value)
    if final_score >= 70:
        final_cell.fill = PatternFill(start_color=COLOR_LIGHT_GREEN, fill_type="solid")
    elif final_score >= 45:
        final_cell.fill = PatternFill(start_color=COLOR_LIGHT_YELLOW, fill_type="solid")
    else:
        final_cell.fill = PatternFill(start_color=COLOR_LIGHT_RED, fill_type="solid")

# Yüzde formatı
percent_cols = [7, 8, 11, 12, 15, 16, 19, 20, 26, 27, 28, 29, 30, 31, 32, 33]
for row in range(2, ws_out.max_row + 1):
    for col in percent_cols:
        ws_out.cell(row=row, column=col).number_format = '0.00"%"'

# Sütun genişlikleri
for column_cells in ws_out.columns:
    column_letter = get_column_letter(column_cells[0].column)
    max_length = 0
    for cell in column_cells:
        try:
            length = len(str(cell.value))
            if length > max_length:
                max_length = length
        except Exception:
            pass
    ws_out.column_dimensions[column_letter].width = min(max(max_length + 2, 11), 26)

ws_out.freeze_panes = "B2"
ws_out.auto_filter.ref = ws_out.dimensions

output = io.BytesIO()
wb.save(output)
output.seek(0)

# ============================================================
# STREAMLIT SONUÇLARI
# ============================================================

st.success("✅ Multi-Vade V3 analizi tamamlandı!")

st.download_button(
    label="📥 V3 Excel Çıktısını İndir",
    data=output,
    file_name="fon_vade_analizi_V3.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# Metodoloji
with st.expander("📘 Metodoloji ve Uyarılar (Okumanız Önerilir)"):
    st.markdown("""
**Skor nasıl hesaplanıyor?**
- Her vade için: Ortalama getiri + Sharpe + Kümülatif getiri + **Sortino** + **MDD** (ters)  
- Bu metrikler z-skor ile normalize edilir (kategori bazlı veya genel).
- Valör cezası uygulanır.
- Nihai skor: 1H %10 + 1A %20 + 3A %30 + 1Y %40 ağırlıklı ortalama.

**Önemli Uyarılar**
- Skorlar **görecelidir**. Tüm fonlar kötü olsa bile bazıları “GÜÇLÜ AL” alabilir.
- Kısa vade (5 gün) çok gürültülüdür.
- KUT benchmark’ı Yahoo Finance vadeli kontratlarına dayanır; fiziksel altın/gümüşten sapma gösterebilir.
- Bu araç **karar destek** aracıdır, tek başına alım-satım tavsiyesi değildir.
    """)

# KUT Özeti
kut_items = [item for item in all_funds if item["code"] == "KUT"]
if kut_items:
    kut = kut_items[0]
    st.subheader("🥇 KUT Özel Analiz")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Nihai Skor", kut.get("final_score", 0))
    col2.metric("Nihai Karar", kut.get("final_decision", "-"))
    col3.metric("1Y Trend", kut.get("trend_1y", "-"))
    col4.metric("3A Trend", kut.get("trend_3a", "-"))
    col5.metric("Güven", kut.get("confidence", "-"))

    st.markdown("### KUT vs Benchmark")
    benchmark_display = pd.DataFrame({
        "Dönem": ["1 Hafta", "1 Ay", "3 Ay", "1 Yıl"],
        "KUT Getiri %": [
            kut.get("cum_5"), kut.get("cum_21"),
            kut.get("cum_63"), kut.get("cum_252"),
        ],
        "Benchmark %": [
            kut.get("benchmark_5"), kut.get("benchmark_21"),
            kut.get("benchmark_63"), kut.get("benchmark_252"),
        ],
        "KUT - Benchmark": [
            kut.get("benchmark_diff_5"), kut.get("benchmark_diff_21"),
            kut.get("benchmark_diff_63"), kut.get("benchmark_diff_252"),
        ],
        "MDD %": [
            kut.get("mdd_5"), kut.get("mdd_21"),
            kut.get("mdd_63"), kut.get("mdd_252"),
        ],
    })
    st.dataframe(benchmark_display, use_container_width=True, hide_index=True)

    if benchmark_df is not None:
        st.info(
            "KUT benchmarkı: %45 Altın + %45 Gümüş + %10 nötr. "
            "Veriler Yahoo Finance vadeli kontratlarından (GC=F, SI=F) alınmış ve USD/TRY ile TL’ye çevrilmiştir. "
            "Fiziksel altın fonlarından sapma gösterebilir."
        )
    else:
        st.warning("KUT bulundu ancak benchmark verisi alınamadı.")

# Tüm fonlar önizleme
st.subheader("📊 Fon Sıralaması")

preview_rows = []
for item in all_funds:
    preview_rows.append({
        "Fon": item["code"],
        "Kategori": item.get("category", "-"),
        "1H": item.get("score_5", 0),
        "1A": item.get("score_21", 0),
        "3A": item.get("score_63", 0),
        "1Y": item.get("score_252", 0),
        "3A Trend": item.get("trend_3a", "-"),
        "1Y Trend": item.get("trend_1y", "-"),
        "Nihai Skor": item.get("final_score", 0),
        "Nihai Karar": item.get("final_decision", "-"),
        "Güven": item.get("confidence", "-"),
        "AUM": item.get("aum", 0),
        "Kaynak": item.get("source", "-"),
    })

df_preview = pd.DataFrame(preview_rows)
st.dataframe(df_preview, use_container_width=True, hide_index=True)
