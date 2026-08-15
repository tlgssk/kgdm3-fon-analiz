ChatGPT:
Evet. Aşağıda verdiğin son kodu baştan sona revize edilmiş tek parça halinde veriyorum.

Yaptığım başlıca düzeltmeler:

Valör * 0.5 şeklindeki aşırı ceza kaldırıldı.
Haftalık getiri bileşik getiri olarak hesaplanıyor.
KGDM skorunda 0–100 normalize edilmiş alt skorlar kullanılıyor.
Sharpe + Sortino benzeri risk-düzeltilmiş skor eklendi.
Drawdown skoru ters yönde doğru çalışıyor.
Konsantrasyon riski ayrı skorlanıyor.
KAZRİSK; AUM + yatırımcı trendi + konsantrasyon + likidite bileşenlerine ayrıldı.
BIST30 tespiti ile ters repo tespiti birbirinden ayrıldı.
Günlük skor hesaplanırken AUM/yatırımcı değişimi yanlış şekilde her güne kopyalanmıyor; statik yapısal metrikler ayrı tutuluyor.
KGDM %55 + KAZRİSK %30 + Kalıcılık %15 hibrit yapısı kullanılıyor.
4 kademeli karar yapısı 5 kademeye çıkarıldı: GÜÇLÜ AL / AL / ASIL LİSTE / DÜZELTME-İZLE / ACİL SAT.
Excel'e KGDM ve KAZRİSK bileşenleri ayrıca yazılıyor.
math eksikliği giderildi.
TEFAS'ta aum veya investors kolonları bulunmazsa hata oluşması engellendi.
Kodun “4 farklı veri kaynağı” şeklindeki yanıltıcı ifadesi düzeltilip mevcut gerçek kaynaklara göre yazıldı.
n_days sıfır olduğunda hata oluşması engellendi.
total_funds == 0 progress-bar problemi giderildi.
İş Yatırım fallback'ında AUM/yatırımcı bilgilerinin sıfırlanması korunuyor.
weekly_return ile filtre artık gerçekten bileşik haftalık getiriyi kullanıyor.
Aşağıdaki kodu mevcut kodunun tamamının yerine koyabilirsin:

import datetime as dt
import io
import math
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
    page_title="KGDM-3 & KAZRİSK® Hibrit Fon Analiz Otomasyonu",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 & KAZRİSK® Hibrit Fon Analiz ve Excel Otomasyonu")

st.caption(
    "TEFAS + İş Yatırım + Fintables veri motoru | "
    "Normalize KGDM Momentum + KAZRİSK® Güvenlik/Likidite "
    "hibrit skorlama"
)


# ============================================================
# GENEL AYARLAR
# ============================================================

FUND_KINDS = ("YAT", "EMK", "BYF")

LOOKBACK_CALENDAR_DAYS = 45
TARGET_TRADING_DAYS = 10
HTTP_TIMEOUT = 12

APP_VERSION = "7.0.0"


# ============================================================
# GITHUB EXCEL
# ============================================================

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
# SKOR AĞIRLIKLARI
# ============================================================

# KGDM toplam ağırlığı = %55
KGDM_WEIGHT = 0.55

# KAZRİSK toplam ağırlığı = %30
KAZRISK_WEIGHT = 0.30

# Son dönem kalıcılığı = %15
PERSISTENCE_WEIGHT = 0.15


# KGDM iç ağırlıkları
KGDM_RETURN_WEIGHT = 0.30
KGDM_RISK_ADJUSTED_WEIGHT = 0.25
KGDM_VOLATILITY_WEIGHT = 0.15
KGDM_DRAWDOWN_WEIGHT = 0.15
KGDM_MOMENTUM_WEIGHT = 0.15


# KAZRİSK iç ağırlıkları
KAZRISK_AUM_WEIGHT = 0.35
KAZRISK_INVESTOR_WEIGHT = 0.25
KAZRISK_CONCENTRATION_WEIGHT = 0.20
KAZRISK_LIQUIDITY_WEIGHT = 0.20


# ============================================================
# KULLANICI KONTROL PANELİ
# ============================================================

st.sidebar.header("⚙️ Analiz & Filtre Kriterleri")

ENABLE_FILTERS = st.sidebar.checkbox(
    "Filtreleri Etkinleştir",
    value=False,
    help=(
        "Kapatılırsa listedeki bütün fonlar hesaplanır. "
        "Açılırsa yatırımcı sayısı ve haftalık getiri filtreleri uygulanır."
    ),
)

TARGET_WEEKLY_RETURN = st.sidebar.slider(
    "Hedef Haftalık Getiri (%)",
    min_value=-5.00,
    max_value=10.00,
    value=0.00,
    step=0.10,
    help=(
        "Filtre aktifse minimum bileşik 5 işlem günlük getiri "
        "beklentisi olarak kullanılır."
    ),
)

MIN_INVESTOR_COUNT = st.sidebar.slider(
    "Minimum Yatırımcı Sayısı",
    min_value=0,
    max_value=100000,
    value=0,
    step=500,
    help=(
        "Filtre aktifse yatırımcı sayısı bu değerin altında "
        "olan fonlar değerlendirme dışı bırakılır."
    ),
)


# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

def parse_number(value) -> Optional[float]:
    """
    Türkçe / İngilizce sayı formatlarını mümkün olduğunca
    güvenli şekilde float'a çevirir.
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
        text
        .replace("₺", "")
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

        if not math.isfinite(value):
            return default

        return value

    except Exception:
        return default


def clamp(value, low=0.0, high=100.0):
    """
    Değeri belirlenen aralıkta tutar.
    """

    try:
        value = float(value)
    except Exception:
        return low

    if not math.isfinite(value):
        return low

    return max(low, min(high, value))


def format_percent(value) -> str:
    number = parse_number(value)

    if number is None:
        return "-"

    if number > 0:
        return f"+%{number:.2f}"

    if number < 0:
        return f"-%{abs(number):.2f}"

    return "%0.00"


def compound_return(returns):
    """
    Günlük getirileri bileşik olarak birleştirir.

    Örnek:
    +%2 ve -%2
    = %0 değil,
    yaklaşık -%0.04.
    """

    if not returns:
        return 0.0

    wealth = 1.0

    for r in returns:
        try:
            r = float(r)

            if not math.isfinite(r):
                continue

            wealth *= (1.0 + r / 100.0)

        except Exception:
            continue

    return (wealth - 1.0) * 100.0


# ============================================================
# YÜZDELİK SKOR
# ============================================================

def percentile_score(
    value,
    values,
    higher_is_better=True,
):
    """
    Fonlar arası göreli performansı 0-100 aralığına taşır.

    Böylece:
        0   = en zayıf bölge
        50  = orta
        100 = en güçlü bölge

    Aykırı Z-score değerlerinin doğrudan skoru ezmesi engellenir.
    """

    clean = []

    for v in values:
        try:
            v = float(v)

            if math.isfinite(v):
                clean.append(v)

        except Exception:
            pass

    if not clean:
        return 50.0

    try:
        value = float(value)

        if not math.isfinite(value):
            return 50.0

    except Exception:
        return 50.0

    if len(clean) == 1:
        return 50.0

    below = sum(v < value for v in clean)
    equal = sum(v == value for v in clean)

    percentile = (
        (below + 0.5 * equal)
        / len(clean)
        * 100.0
    )

    if not higher_is_better:
        percentile = 100.0 - percentile

    return clamp(percentile)


# ============================================================
# DOWNSIDE VOLATİLİTE
# ============================================================

def downside_volatility(returns):
    """
    Negatif getirilerin RMS tipi downside volatilitesini hesaplar.
    """

    downside = []

    for r in returns:

        try:
            r = float(r)

            if math.isfinite(r) and r < 0:
                downside.append(r)

        except Exception:
            pass

    if not downside:
        return 0.0

    return math.sqrt(
        sum(r ** 2 for r in downside)
        / len(downside)
    )


# ============================================================
# RİSK DÜZELTİLMİŞ SKOR
# ============================================================

def risk_adjusted_score(returns):
    """
    Sharpe benzeri + Sortino benzeri skor.

    Sonuç:
        0-100

    Sortino daha yüksek ağırlıklı tutulur çünkü fon analizinde
    aşağı yönlü riskin cezalandırılması daha anlamlıdır.
    """

    if not returns:
        return 50.0

    clean = []

    for r in returns:
        try:
            r = float(r)

            if math.isfinite(r):
                clean.append(r)

        except Exception:
            pass

    if len(clean) < 2:
        return 50.0

    mean_ret = sum(clean) / len(clean)

    variance = sum(
        (r - mean_ret) ** 2
        for r in clean
    ) / len(clean)

    volatility = math.sqrt(
        max(0.0, variance)
    )

    if volatility > 1e-12:
        sharpe = mean_ret / volatility
    else:
        sharpe = 3.0 if mean_ret > 0 else 0.0

    downside = downside_volatility(clean)

    if downside > 1e-12:
        sortino = mean_ret / downside
    else:
        sortino = 3.0 if mean_ret > 0 else 0.0

    sharpe = max(-3.0, min(3.0, sharpe))
    sortino = max(-3.0, min(3.0, sortino))

    sharpe_score = (
        50.0 + (sharpe / 3.0) * 50.0
    )

    sortino_score = (
        50.0 + (sortino / 3.0) * 50.0
    )

    return clamp(
        sharpe_score * 0.45
        + sortino_score * 0.55
    )


# ============================================================
# MAX DRAWDOWN
# ============================================================

def calculate_drawdown(prices):
    """
    Maksimum drawdown değerini pozitif risk büyüklüğü olarak döndürür.

    Örneğin:
        -%5 drawdown -> 5
        -%20 drawdown -> 20
    """

    if not prices or len(prices) < 2:
        return 0.0

    peak = prices[0]
    max_dd = 0.0

    for price in prices:

        try:
            price = float(price)
        except Exception:
            continue

        if price <= 0:
            continue

        if price > peak:
            peak = price

        if peak > 0:
            dd = (
                price / peak - 1.0
            ) * 100.0

            max_dd = min(
                max_dd,
                dd,
            )

    return abs(max_dd)


def drawdown_score(max_dd):
    """
    Drawdown'u 0-100 güvenlik skoruna çevirir.

    0% DD -> 100
    10% DD -> yaklaşık 37
    Daha yüksek DD -> hızla düşer.
    """

    dd = max(
        0.0,
        safe_float(max_dd),
    )

    score = 100.0 * math.exp(
        -dd / 10.0
    )

    return clamp(score)


# ============================================================
# KONSANTRASYON SKORU
# ============================================================

def concentration_score(top_asset_weight):
    """
    En büyük varlık ağırlığından konsantrasyon riski çıkarır.

    <= %10  -> 100
    %30      -> 50
    >= %50   -> 0
    """

    weight = max(
        0.0,
        safe_float(top_asset_weight),
    )

    if weight <= 10:
        return 100.0

    if weight >= 50:
        return 0.0

    score = (
        100.0
        - (
            (weight - 10.0)
            / 40.0
            * 100.0
        )
    )

    return clamp(score)


# ============================================================
# LİKİDİTE SKORU
# ============================================================

def liquidity_score(item):
    """
    Likidite / güvenlik skoru.

    BIST30 bilgisi tek başına güvenlik garantisi olarak
    kullanılmaz; yalnızca sınırlı bir bonus sağlar.
    """

    score = 50.0

    cash_ratio = safe_float(
        item.get(
            "emergency_cash_ratio",
            5.0,
        )
    )

    if cash_ratio >= 20:
        score += 25

    elif cash_ratio >= 10:
        score += 15

    elif cash_ratio >= 5:
        score += 5

    else:
        score -= 10

    if item.get("is_bist30", False):
        score += 10

    top_weight = safe_float(
        item.get(
            "top_asset_weight",
            0.0,
        )
    )

    if top_weight > 40:
        score -= 20

    elif top_weight > 30:
        score -= 10

    return clamp(score)


# ============================================================
# KAZRİSK SKORU
# ============================================================

def calculate_kazrisk_score(item):
    """
    KAZRİSK® 0-100 güvenlik / dayanıklılık skoru.

    %35 AUM
    %25 yatırımcı trendi
    %20 konsantrasyon
    %20 likidite
    """

    # --------------------------------------------------------
    # AUM SKORU
    # --------------------------------------------------------

    aum = max(
        safe_float(item.get("aum", 0.0)),
        1.0,
    )

    aum_million = aum / 1_000_000.0

    if aum_million <= 0:
        aum_score = 0.0

    else:
        aum_score = clamp(
            50.0
            + math.log10(
                max(aum_million, 0.1)
            ) * 20.0
        )

    # --------------------------------------------------------
    # YATIRIMCI TRENDİ
    # --------------------------------------------------------

    inv_change = safe_float(
        item.get("inv_change", 0.0)
    )

    investor_score = clamp(
        50.0 + inv_change * 2.0
    )

    # --------------------------------------------------------
    # KONSANTRASYON
    # --------------------------------------------------------

    concentration = concentration_score(
        item.get(
            "top_asset_weight",
            0.0,
        )
    )

    # --------------------------------------------------------
    # LİKİDİTE
    # --------------------------------------------------------

    liquidity = liquidity_score(item)

    # --------------------------------------------------------
    # KAZRİSK
    # --------------------------------------------------------

    kazrisk = (
        aum_score
        * KAZRISK_AUM_WEIGHT

        + investor_score
        * KAZRISK_INVESTOR_WEIGHT

        + concentration
        * KAZRISK_CONCENTRATION_WEIGHT

        + liquidity
        * KAZRISK_LIQUIDITY_WEIGHT
    )

    return clamp(kazrisk)


# ============================================================
# KGDM SKORU
# ============================================================

def calculate_kgdm_score(
    item,
    peer_items,
):
    """
    KGDM Momentum skoru.

    %30 Kümülatif Getiri
    %25 Risk Düzeltilmiş Getiri
    %15 Volatilite
    %15 Drawdown
    %15 Son 3 Gün Momentum
    """

    returns = item.get(
        "daily_returns",
        [],
    )

    prices = item.get(
        "prices",
        [],
    )

    if not returns or len(prices) < 2:
        return 50.0

    # --------------------------------------------------------
    # FON METRİKLERİ
    # --------------------------------------------------------

    mean_ret = (
        sum(returns) / len(returns)
    )

    if len(returns) >= 2:

        variance = sum(
            (r - mean_ret) ** 2
            for r in returns
        ) / len(returns)

        volatility = math.sqrt(
            max(0.0, variance)
        )

    else:
        volatility = 0.0

    cumulative = compound_return(
        returns
    )

    max_dd = calculate_drawdown(
        prices
    )

    risk_adj = risk_adjusted_score(
        returns
    )

    recent_returns = returns[-3:]

    recent_momentum = compound_return(
        recent_returns
    )

    # --------------------------------------------------------
    # AKRAN HAVUZLARI
    # --------------------------------------------------------

    peer_cum = []
    peer_mean = []
    peer_vol = []
    peer_momentum = []

    for peer in peer_items:

        peer_returns = peer.get(
            "daily_returns",
            [],
        )

        if not peer_returns:
            continue

        peer_cum.append(
            compound_return(
                peer_returns
            )
        )

        peer_mean.append(
            sum(peer_returns)
            / len(peer_returns)
        )

        if len(peer_returns) >= 2:

            peer_mean_value = (
                sum(peer_returns)
                / len(peer_returns)
            )

            peer_variance = sum(
                (
                    r - peer_mean_value
                ) ** 2
                for r in peer_returns
            ) / len(peer_returns)

            peer_vol.append(
                math.sqrt(
                    max(
                        0.0,
                        peer_variance,
                    )
                )
            )

        else:
            peer_vol.append(0.0)

        peer_momentum.append(
            compound_return(
                peer_returns[-3:]
            )
        )

    # --------------------------------------------------------
    # ALT SKORLAR
    # --------------------------------------------------------

    score_return = percentile_score(
        cumulative,
        peer_cum,
        higher_is_better=True,
    )

    score_mean = percentile_score(
        mean_ret,
        peer_mean,
        higher_is_better=True,
    )

    score_vol = percentile_score(
        volatility,
        peer_vol,
        higher_is_better=False,
    )

    score_momentum = percentile_score(
        recent_momentum,
        peer_momentum,
        higher_is_better=True,
    )

    score_risk_adj = clamp(
        risk_adj
    )

    score_dd = drawdown_score(
        max_dd
    )

    # --------------------------------------------------------
    # KGDM
    # --------------------------------------------------------

    kgdm = (
        score_return
        * KGDM_RETURN_WEIGHT

        + score_risk_adj
        * KGDM_RISK_ADJUSTED_WEIGHT

        + score_vol
        * KGDM_VOLATILITY_WEIGHT

        + score_dd
        * KGDM_DRAWDOWN_WEIGHT

        + score_momentum
        * KGDM_MOMENTUM_WEIGHT
    )

    return clamp(kgdm)


# ============================================================
# EXCEL / VERİ FONKSİYONLARI
# ============================================================

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
            df["price"] = df[
                "price"
            ].apply(parse_number)

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
            subset=[
                "date",
                "code",
                "price",
            ]
        )

        df = df[
            df["price"] > 0
        ]

        return (
            df
            .sort_values(
                ["code", "date"]
            )
            .drop_duplicates(
                subset=[
                    "code",
                    "date",
                ],
                keep="last",
            )
            .reset_index(drop=True)
        )

    except Exception:

        return pd.DataFrame()


# ============================================================
# İŞ YATIRIM
# ============================================================

def fetch_isyatirim_series(
    fund_code: str,
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
        "_layouts/15/"
        "IsYatirim.Website/"
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
            timeout=HTTP_TIMEOUT,
        )

        if response.status_code != 200:
            return None

        payload = response.json()

        values = payload.get(
            "value"
        )

        if not values:
            return None

        df = pd.DataFrame(
            values
        )

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
            subset=[
                "date",
                "price",
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
                keep="last",
            )
            .tail(
                TARGET_TRADING_DAYS + 1
            )
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
# TEFAS DOĞRUDAN API
# ============================================================

def fetch_tefas_direct_api(
    fund_code: str,
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
        "https://www.tefas.gov.tr/"
        "api/DB/BindHistoryInfo"
    )

    payload = {
        "fontip": "YAT",
        "fonkod": code,
        "bastarih": start.strftime(
            "%d.%m.%Y"
        ),
        "bittarih": end.strftime(
            "%d.%m.%Y"
        ),
    }

    headers = {
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://www.tefas.gov.tr",
    }

    try:

        res = requests.post(
            url,
            data=payload,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        if res.status_code != 200:
            return None

        data = res.json().get(
            "data",
            [],
        )

        if not data:
            return None

        df = pd.DataFrame(data)

        if (
            "TARIH" not in df.columns
            or "FIYAT" not in df.columns
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
            subset=[
                "date",
                "price",
            ]
        )

        df = df[
            df["price"] > 0
        ]

        if len(df) < 2:
            return None

        return (
            df
            .sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .tail(
                TARGET_TRADING_DAYS + 1
            )
            .reset_index(drop=True)
        )

    except Exception:

        return None


# ============================================================
# FİNTABLES YAPISAL VERİ
# ============================================================

def fetch_fund_structural_data(
    fund_code: str,
) -> dict:

    code = normalize_fund_code(
        fund_code
    )

    structural = {
        "top_asset_weight": 0.0,
        "is_bist30": False,
        "emergency_cash_ratio": 5.0,
    }

    try:

        fintables_url = (
            f"https://fintables.com/"
            f"fonlar/{code.lower()}"
        )

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        fin_res = requests.get(
            fintables_url,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        if fin_res.status_code != 200:
            return structural

        text = fin_res.text

        # ----------------------------------------------------
        # EN BÜYÜK VARLIK
        # ----------------------------------------------------

        match_top = re.search(
            r'En Büyük Pay["\s:]+'
            r'([0-9]+(?:\.[0-9]+)?)',
            text,
            re.IGNORECASE,
        )

        if match_top:

            structural[
                "top_asset_weight"
            ] = safe_float(
                match_top.group(1)
            )

        # ----------------------------------------------------
        # BIST30
        # ----------------------------------------------------
        # Ters Repo görülmesi artık BIST30 olarak işaretlenmiyor.

        if (
            "BIST 30" in text
            or "BIST30" in text
        ):

            structural[
                "is_bist30"
            ] = True

        # ----------------------------------------------------
        # NAKİT / TERS REPO
        # ----------------------------------------------------

        match_cash = re.search(
            r'(?:Nakit|Ters Repo|PPF)'
            r'["\s:]+'
            r'([0-9]+(?:\.[0-9]+)?)',
            text,
            re.IGNORECASE,
        )

        if match_cash:

            structural[
                "emergency_cash_ratio"
            ] = safe_float(
                match_cash.group(1),
                5.0,
            )

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
    code = normalize_fund_code(
        fund_code
    )

    if not code:
        return None, "Bulunamadı"

    # --------------------------------------------------------
    # 1. TEFAS UNIVERSE
    # --------------------------------------------------------

    if (
        universe is not None
        and not universe.empty
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
                    .tail(
                        TARGET_TRADING_DAYS + 1
                    )
                    .reset_index(drop=True),
                    "TEFAS",
                )

    # --------------------------------------------------------
    # 2. TEFAS DIRECT API
    # --------------------------------------------------------

    direct_df = (
        fetch_tefas_direct_api(
            code
        )
    )

    if (
        direct_df is not None
        and len(direct_df) >= 2
    ):

        return (
            direct_df,
            "TEFAS Direct API",
        )

    # --------------------------------------------------------
    # 3. İŞ YATIRIM
    # --------------------------------------------------------

    is_df = (
        fetch_isyatirim_series(
            code
        )
    )

    if (
        is_df is not None
        and len(is_df) >= 2
    ):

        return (
            is_df,
            "İş Yatırım",
        )

    return None, "Bulunamadı"


# ============================================================
# FON METRİKLERİ
# ============================================================

def compute_fund_metrics(
    series: Optional[pd.DataFrame],
    fund_code: str,
) -> Optional[dict]:

    if (
        series is None
        or len(series) < 2
    ):
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

    # TEFAS / API kolonları eksik olabilir
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
        subset=[
            "date",
            "price",
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
    # GÜNLÜK GETİRİLER
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
                (
                    current
                    / previous
                    - 1.0
                )
                * 100.0
            )

    if not daily_returns:
        return None

    # --------------------------------------------------------
    # MAX DD
    # --------------------------------------------------------

    max_dd = calculate_drawdown(
        prices
    )

    # --------------------------------------------------------
    # AUM DEĞİŞİMİ
    # --------------------------------------------------------

    if aums[0] > 0:

        aum_change = (
            (
                aums[-1]
                / aums[0]
            )
            - 1.0
        ) * 100.0

    else:

        aum_change = 0.0

    # --------------------------------------------------------
    # YATIRIMCI DEĞİŞİMİ
    # --------------------------------------------------------

    if investors[0] > 0:

        inv_change = (
            (
                investors[-1]
                / investors[0]
            )
            - 1.0
        ) * 100.0

    else:

        inv_change = 0.0

    # --------------------------------------------------------
    # BİLEŞİK 5 GÜNLÜK GETİRİ
    # --------------------------------------------------------

    weekly_returns = (
        daily_returns[-5:]
        if len(daily_returns) >= 5
        else daily_returns
    )

    recent_weekly_ret = compound_return(
        weekly_returns
    )

    # --------------------------------------------------------
    # YAPISAL VERİ
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
            round(
                investors[-1]
            )
        ),

        "aum_change": aum_change,
        "inv_change": inv_change,

        "max_dd": max_dd,

        "weekly_return": recent_weekly_ret,

        **structural,
    }


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
    min_width: int = 10,
    max_width: int = 45,
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

    if "KGDM3_Puanlama" in wb.sheetnames:
        del wb["KGDM3_Puanlama"]

    ws_scores = wb.create_sheet(
        title="KGDM3_Puanlama"
    )

    headers = [
        "Fon Kodu",
        "Valör",

        "Hibrit Skor",
        "KGDM Skor",
        "KAZRİSK Skor",
        "Kalıcılık Skoru",

        "Son 5 Hibrit Skor",
        "Son Skor",
        "Skor Trend",
        "Model Kararı",

        "Ort. Günlük Getiri (%)",
        "Volatilite (%)",
        "Risk Düzeltilmiş Skor",
        "Kümülatif Getiri (%)",

        "MaxDD (%)",

        "En Büyük Varlık (%)",
        "KAZRİSK Likidite Skoru",

        "AUM Değişim (%)",
        "Yatırımcı Değişim (%)",

        "Fon Büyüklüğü (AUM ₺)",
        "Yatırımcı Sayısı",

        "Haftalık Bileşik Getiri (%)",
    ]

    for day in range(1, n_days + 1):

        headers.append(
            f"Gün {day} Hibrit Skor"
        )

    for day in range(1, n_days + 1):

        headers.append(
            f"Gün {day} Getiri"
        )

    ws_scores.append(
        headers
    )

    # --------------------------------------------------------
    # HEADER
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

    ws_scores.row_dimensions[
        1
    ].height = 45

    # --------------------------------------------------------
    # SATIRLAR
    # --------------------------------------------------------

    for item in calculated_funds:

        risk_adj = risk_adjusted_score(
            item["daily_returns"]
        )

        cumulative = compound_return(
            item["daily_returns"]
        )

        liquidity = liquidity_score(
            item
        )

        persistence_score = (
            clamp(
                50.0
                + item.get(
                    "score_trend",
                    0.0
                ) * 5.0
            )
        )

        row_data = [
            item["code"],
            item["valor"],

            item["kgdm_skor"],
            item.get(
                "kgdm_component",
                0.0,
            ),
            item.get(
                "kazrisk_component",
                0.0,
            ),
            round(
                persistence_score,
                1,
            ),

            item["last_5_scores_str"],
            item.get(
                "last_score",
                0,
            ),
            round(
                item.get(
                    "score_trend",
                    0.0,
                ),
                1,
            ),
            item["karar"],

            round(
                sum(
                    item["daily_returns"]
                )
                / len(
                    item["daily_returns"]
                ),
                4,
            ),

            round(
                (
                    sum(
                        (
                            r
                            - (
                                sum(
                                    item[
                                        "daily_returns"
                                    ]
                                )
                                / len(
                                    item[
                                        "daily_returns"
                                    ]
                                )
                            )
                        ) ** 2
                        for r in item[
                            "daily_returns"
                        ]
                    )
                    / len(
                        item[
                            "daily_returns"
                        ]
                    )
                )
                ** 0.5,
                4,
            ),

            round(
                risk_adj,
                2,
            ),

            round(
                cumulative,
                4,
            ),

            round(
                item["max_dd"],
                2,
            ),

            round(
                item[
                    "top_asset_weight"
                ],
                2,
            ),

            round(
                liquidity,
                2,
            ),

            round(
                item["aum_change"],
                2,
            ),

            round(
                item["inv_change"],
                2,
            ),

            (
                round(
                    item["aum"],
                    2,
                )
                if item["aum"]
                else None
            ),

            (
                int(
                    item["investors"]
                )
                if item["investors"]
                else None
            ),

            round(
                item[
                    "weekly_return"
                ],
                4,
            ),
        ]

        row_data.extend(
            item["running_scores"]
        )

        row_data.extend(
            [
                format_percent(value)
                for value in item[
                    "daily_returns"
                ]
            ]
        )

        ws_scores.append(
            row_data
        )

    # --------------------------------------------------------
    # KARAR RENKLERİ
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
                column=10,
            )
        )

        decision_text = str(
            decision_cell.value
        )

        if (
            "GÜÇLÜ AL"
            in decision_text
            or decision_text == "AL"
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
    # HİBRİT SKOR KOŞULLU FORMAT
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
            formula=["60", "74.99"],
            fill=PatternFill(
                start_color="EAF4E3",
                fill_type="solid",
            ),
        ),
    )

    ws_scores.conditional_formatting.add(
        score_range,
        CellIsRule(
            operator="between",
            formula=["45", "59.99"],
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
            formula=["45"],
            fill=PatternFill(
                start_color=COLOR_LIGHT_RED,
                fill_type="solid",
            ),
        ),
    )

    # --------------------------------------------------------
    # SAYFA STİL
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # ÇIKTI
    # --------------------------------------------------------

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
    "📂 Portföy Excel Listesi Seçimi"
)

col_upload, col_github = st.columns(
    2
)

wb = None
source_mode = None


# ============================================================
# DOSYA YÜKLEME
# ============================================================

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
                f"Hata: {exc}"
            )


# ============================================================
# GITHUB
# ============================================================

with col_github:

    st.write(
        "Veya GitHub'daki listeyi kullanın:"
    )

    if st.button(
        "🚀 GitHub'dan Çek ve Analiz Et",
        use_container_width=True,
    ):

        try:

            gh_response = requests.get(
                GITHUB_EXCEL_URL,
                timeout=HTTP_TIMEOUT,
            )

            gh_response.raise_for_status()

            wb = openpyxl.load_workbook(
                io.BytesIO(
                    gh_response.content
                )
            )

            source_mode = "github"

            st.success(
                "✅ Excel dosyası başarıyla indirildi!"
            )

        except Exception as exc:

            st.error(
                f"Bağlantı hatası: {exc}"
            )


if wb is None:
    st.stop()


# ============================================================
# FON LİSTESİ SAYFASI
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

        valor_cell = (
            row[3].value
            if len(row) > 3
            else 0
        )

        parsed_valor = parse_number(
            valor_cell
        )

        excel_valor_dict[
            code
        ] = (
            int(
                round(
                    parsed_valor
                )
            )
            if parsed_valor is not None
            else 0
        )

    except Exception:

        excel_valor_dict[
            code
        ] = 0


requested_codes = list(
    dict.fromkeys(
        requested_codes
    )
)


if not requested_codes:

    st.error(
        "Fon_Listesi sayfasında geçerli fon kodu bulunamadı."
    )

    st.stop()


# ============================================================
# TARİH
# ============================================================

today = dt.date.today()

start_date = (
    today
    - dt.timedelta(
        days=LOOKBACK_CALENDAR_DAYS
    )
)


# ============================================================
# TEFAS UNIVERSE
# ============================================================

with st.spinner(
    "🔄 TEFAS fon evreni alınıyor..."
):

    universe = fetch_tefas_universe(
        start_date,
        today,
    )


# ============================================================
# FON VERİLERİ
# ============================================================

calculated_funds = []

failed_codes = []

progress = st.progress(
    0,
    text="Fonlar analiz ediliyor...",
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

    else:

        weekly_ok = (
            metrics["weekly_return"]
            >= TARGET_WEEKLY_RETURN
        )

        investor_ok = (
            metrics["investors"]
            >= MIN_INVESTOR_COUNT
        )

        if (
            not ENABLE_FILTERS
            or (
                weekly_ok
                and investor_ok
            )
        ):

            calculated_funds.append(
                {
                    "code": code,
                    "valor": excel_valor_dict.get(
                        code,
                        0,
                    ),
                    "source": source,
                    **metrics,
                }
            )

    progress.progress(
        index / total_funds,
        text=f"{code} analiz edildi...",
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
        "Fonların ortak analiz edilebileceği yeterli veri günü yok."
    )

    st.stop()


for item in calculated_funds:

    item["dates"] = (
        item["dates"][-n_days:]
    )

    item["daily_returns"] = (
        item["daily_returns"][-n_days:]
    )

    item["prices"] = (
        item["prices"][
            -(n_days + 1):
        ]
    )

    item["running_scores"] = []
    item["running_kgdm"] = []
    item["running_kazrisk"] = []


# ============================================================
# KGDM + KAZRİSK V2 HİBRİT SKORLAMA
# ============================================================

for d in range(
    1,
    n_days + 1,
):

    # --------------------------------------------------------
    # O GÜNE KADAR OLAN VERİDEN GEÇİCİ HAVUZ
    # --------------------------------------------------------

    daily_pool = []

    for original in calculated_funds:

        temp = dict(
            original
        )

        temp["daily_returns"] = (
            original[
                "daily_returns"
            ][:d]
        )

        temp["prices"] = (
            original[
                "prices"
            ][:d + 1]
        )

        daily_pool.append(
            temp
        )

    # --------------------------------------------------------
    # HER FON
    # --------------------------------------------------------

    for i, item in enumerate(
        calculated_funds
    ):

        current = daily_pool[i]

        # KGDM
        kgdm = calculate_kgdm_score(
            current,
            daily_pool,
        )

        # KAZRİSK
        kazrisk = (
            calculate_kazrisk_score(
                current
            )
        )

        # ----------------------------------------------------
        # KALICILIK
        # ----------------------------------------------------
        # Günlük aşırı sıçramaların nihai skoru tek başına
        # değiştirmemesi için 50 merkezli bir kalıcılık
        # komponenti kullanılır.
        #
        # İlk günlerde yeterli geçmiş olmadığı için 50'ye
        # yakın tutulur.

        existing_scores = (
            item["running_scores"]
        )

        if len(existing_scores) >= 3:

            recent_mean = (
                sum(
                    existing_scores[-3:]
                )
                / 3.0
            )

            persistence = clamp(
                recent_mean
            )

        elif existing_scores:

            persistence = clamp(
                sum(existing_scores)
                / len(existing_scores)
            )

        else:

            persistence = 50.0

        # ----------------------------------------------------
        # HİBRİT SKOR
        # ----------------------------------------------------

        hybrid = (
            kgdm
            * KGDM_WEIGHT

            + kazrisk
            * KAZRISK_WEIGHT

            + persistence
            * PERSISTENCE_WEIGHT
        )

        hybrid = int(
            round(
                clamp(
                    hybrid,
                    0.0,
                    100.0,
                )
            )
        )

        item[
            "running_kgdm"
        ].append(
            round(
                kgdm,
                1,
            )
        )

        item[
            "running_kazrisk"
        ].append(
            round(
                kazrisk,
                1,
            )
        )

        item[
            "running_scores"
        ].append(
            hybrid
        )


# ============================================================
# NİHAİ FON SKORLARI
# ============================================================

for item in calculated_funds:

    scores = (
        item["running_scores"]
    )

    kgdm_scores = (
        item["running_kgdm"]
    )

    kazrisk_scores = (
        item["running_kazrisk"]
    )

    # --------------------------------------------------------
    # SON 5
    # --------------------------------------------------------

    last_5 = (
        scores[-5:]
        if len(scores) >= 5
        else scores
    )

    last_kgdm = (
        kgdm_scores[-5:]
        if len(kgdm_scores) >= 5
        else kgdm_scores
    )

    last_kazrisk = (
        kazrisk_scores[-5:]
        if len(kazrisk_scores) >= 5
        else kazrisk_scores
    )

    # --------------------------------------------------------
    # ORTALAMALAR
    # --------------------------------------------------------

    item[
        "kgdm_skor"
    ] = int(
        round(
            sum(last_5)
            / len(last_5)
        )
    )

    item[
        "kgdm_component"
    ] = round(
        sum(last_kgdm)
        / len(last_kgdm),
        1,
    )

    item[
        "kazrisk_component"
    ] = round(
        sum(last_kazrisk)
        / len(last_kazrisk),
        1,
    )

    item[
        "last_score"
    ] = scores[-1]

    # --------------------------------------------------------
    # TREND
    # --------------------------------------------------------

    item[
        "score_trend"
    ] = (
        item["last_score"]
        - item["kgdm_skor"]
    )

    item[
        "last_5_scores_str"
    ] = " ➔ ".join(
        str(s)
        for s in last_5
    )

    # --------------------------------------------------------
    # TEMEL METRİKLER
    # --------------------------------------------------------

    returns = item[
        "daily_returns"
    ]

    item[
        "mean_return"
    ] = (
        sum(returns)
        / len(returns)
        if returns
        else 0.0
    )

    if len(returns) >= 2:

        variance = sum(
            (
                r
                - item[
                    "mean_return"
                ]
            ) ** 2
            for r in returns
        ) / len(returns)

        item[
            "volatility"
        ] = math.sqrt(
            max(
                0.0,
                variance,
            )
        )

    else:

        item[
            "volatility"
        ] = 0.0

    item[
        "sharpe_like"
    ] = risk_adjusted_score(
        returns
    )

    item[
        "cumulative_return"
    ] = compound_return(
        returns
    )

    # --------------------------------------------------------
    # KALICILIK
    # --------------------------------------------------------

    item[
        "persistence_score"
    ] = clamp(
        50.0
        + item["score_trend"]
        * 5.0
    )

    # --------------------------------------------------------
    # KARAR
    # --------------------------------------------------------

    score = item[
        "kgdm_skor"
    ]

    if score >= 75:

        item[
            "karar"
        ] = "GÜÇLÜ AL"

    elif score >= 60:

        item[
            "karar"
        ] = "AL"

    elif score >= 45:

        item[
            "karar"
        ] = "ASIL LİSTE"

    elif score >= 30:

        item[
            "karar"
        ] = "DÜZELTME / İZLE"

    else:

        item[
            "karar"
        ] = "ACİL SAT"


# ============================================================
# SIRALAMA
# ============================================================

calculated_funds.sort(
    key=lambda x: (
        -x["kgdm_skor"],
        -x.get(
            "cumulative_return",
            0.0,
        ),
        -x.get(
            "kazrisk_component",
            0.0,
        ),
    )
)


# ============================================================
# EKRAN SONUÇLARI
# ============================================================

display_rows = []

for item in calculated_funds:

    top_weight = safe_float(
        item.get(
            "top_asset_weight",
            0.0,
        )
    )

    if top_weight > 30.0:

        risk_status = (
            "⚠️ Sığ / Yüksek Konsantrasyon"
        )

    elif top_weight > 0:

        risk_status = (
            "🟡 Kontrollü Konsantrasyon"
        )

    else:

        risk_status = (
            "✅ Normal"
        )

    display_rows.append(
        {
            "Fon Kodu": item[
                "code"
            ],

            "Hibrit Skor": item[
                "kgdm_skor"
            ],

            "KGDM": item.get(
                "kgdm_component",
                0,
            ),

            "KAZRİSK": item.get(
                "kazrisk_component",
                0,
            ),

            "Kalıcılık": item.get(
                "persistence_score",
                50,
            ),

            "Son Skor": item.get(
                "last_score",
                0,
            ),

            "Trend": round(
                item.get(
                    "score_trend",
                    0,
                ),
                1,
            ),

            "Model Kararı": item[
                "karar"
            ],

            "Ort. Günlük %": round(
                item[
                    "mean_return"
                ],
                3,
            ),

            "Risk Düzeltilmiş": round(
                item[
                    "sharpe_like"
                ],
                2,
            ),

            "Kümülatif Getiri %": round(
                item[
                    "cumulative_return"
                ],
                3,
            ),

            "MaxDD %": round(
                item[
                    "max_dd"
                ],
                2,
            ),

            "En Büyük Varlık %": round(
                item[
                    "top_asset_weight"
                ],
                2,
            ),

            "KAZRİSK Güvenlik": risk_status,

            "AUM ₺": round(
                item[
                    "aum"
                ],
                0,
            ),

            "Yatırımcı": item[
                "investors"
            ],
        }
    )


df_display = pd.DataFrame(
    display_rows
)


# ============================================================
# EKRAN RENKLENDİRME
# ============================================================

def color_cells(value):

    text = str(value)

    if (
        "GÜÇLÜ AL"
        in text
        or text == "AL"
        or "ASIL LİSTE"
        in text
        or "Normal"
        in text
    ):

        return (
            "color: #008000; "
            "font-weight: bold;"
        )

    if (
        "DÜZELTME"
        in text
        or "Kontrollü"
        in text
    ):

        return (
            "color: #B8860B; "
            "font-weight: bold;"
        )

    if (
        "ACİL SAT"
        in text
        or "Sığ"
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


st.subheader(
    "📊 KGDM + KAZRİSK V2 Analiz Sonuçları"
)

st.dataframe(
    styled_df,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# MODEL AÇIKLAMASI
# ============================================================

with st.expander(
    "ℹ️ Skor Mantığını Gör"
):

    st.markdown(
        """
### Hibrit Skor

**Nihai Hibrit Skor =**

- **%55 KGDM Momentum**
- **%30 KAZRİSK Güvenlik/Likidite**
- **%15 Kalıcılık**

### KGDM

- %30 Kümülatif getiri
- %25 Risk-düzeltilmiş getiri
- %15 Düşük volatilite avantajı
- %15 Drawdown avantajı
- %15 Son 3 günlük momentum

### KAZRİSK

- %35 Fon büyüklüğü / AUM
- %25 Yatırımcı trendi
- %20 Portföy konsantrasyonu
- %20 Likidite / güvenlik

### Karar Matrisi

| Skor | Karar |
|---:|---|
| 75–100 | 🟢 GÜÇLÜ AL |
| 60–74 | 🟢 AL |
| 45–59 | 🟡 ASIL LİSTE |
| 30–44 | 🟠 DÜZELTME / İZLE |
| 0–29 | 🔴 ACİL SAT |

**Not:** Skor göreli bir sıralama modelidir; yatırım tavsiyesi veya gelecekteki getirinin garantisi değildir.
"""
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


# ============================================================
# SONUÇ
# ============================================================

st.success(
    "✅ Analiz tamamlandı. "
    "KGDM + KAZRİSK V2 normalize hibrit skor motoru uygulandı."
)

st.download_button(
    label="📥 Güncellenmiş Hibrit Excel'i İndir",
    data=output,
    file_name=(
        "fonlar_KGDM_KAZRISK_V7.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)
