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

st.set_page_config(page_title="Multi-Vade Fon Analizi", page_icon="📈", layout="wide")

st.title("📈 Multi-Vade (Kısa-Orta-Uzun-ÇokUzun) Fon Analiz Otomasyonu")
st.caption("Fonları 1 Hafta (5 gün), 1 Ay (21 gün), 3 Ay (63 gün) ve 1 Yıl (252 gün) pencerelerinde bağımsız Z-Skor ile yarıştırır.")

FUND_KINDS = ("YAT", "EMK", "BYF")
# 252 işlem günü (1 yıl) verisi toplayabilmek için takvim gününü genişletiyoruz
# (~365 takvim gününde ~252 işlem günü olur; tatil/veri boşluklarına karşı pay bırakıyoruz)
LOOKBACK_CALENDAR_DAYS = 400
HTTP_TIMEOUT = 8

# Vade İşlem Günü Tanımlamaları
PERIODS = {
    "Kısa Vade (1 Hafta)": 5,
    "Orta Vade (1 Ay)": 21,
    "Uzun Vade (3 Ay)": 63,
    "Çok Uzun Vade (1 Yıl)": 252,
}
PERIOD_LABELS = {
    "Kısa Vade (1 Hafta)": "Kısa Vade (1H)",
    "Orta Vade (1 Ay)": "Orta Vade (1A)",
    "Uzun Vade (3 Ay)": "Uzun Vade (3A)",
    "Çok Uzun Vade (1 Yıl)": "Çok Uzun Vade (1Y)",
}
MAX_DAYS = max(PERIODS.values())

COLOR_NAVY = "1F4E79"
COLOR_GREEN = "008000"
COLOR_RED = "FF0000"
COLOR_YELLOW = "B8860B"
COLOR_WHITE = "FFFFFF"
COLOR_LIGHT_GREEN = "E2F0D9"
COLOR_LIGHT_RED = "FCE4D6"
COLOR_LIGHT_YELLOW = "FFF2CC"


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
    text = text.replace("₺", "").replace("TL", "").replace("%", "").replace(" ", "").strip()
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


@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: dt.date, end_date: dt.date) -> pd.DataFrame:
    try:
        from pytefas import Crawler
    except ImportError:
        return pd.DataFrame()
    try:
        crawler = Crawler(timeout=60, max_retry=5)
        df = crawler.fetch_many(start=start_date, end=end_date, kinds=FUND_KINDS, columns="info")
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
        if not all(c in df.columns for c in ["date", "code", "price"]):
            return pd.DataFrame()
        df["code"] = df["code"].astype(str).str.strip().str.upper()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["price"] = df["price"].apply(parse_number)

        # FIX (bug #4): df.get("aum", 0.0) döndürdüğünde sütun yoksa bir float
        # döner ve .apply() çağrısı hata verirdi. Sütun eksikse önce 0.0 ile
        # dolu bir Series/sütun oluşturup öyle apply ediyoruz.
        if "aum" not in df.columns:
            df["aum"] = 0.0
        else:
            df["aum"] = df["aum"].apply(parse_number)

        if "investors" not in df.columns:
            df["investors"] = 0.0
        else:
            df["investors"] = df["investors"].apply(parse_number)

        df = df.dropna(subset=["date", "code", "price"])
        df = df[df["price"] > 0]
        return (
            df.sort_values(["code", "date"])
            .drop_duplicates(subset=["code", "date"], keep="last")
            .reset_index(drop=True)
        )
    except Exception:
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


def fetch_isyatirim_series(fund_code: str) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)
    if not code:
        return None
    end = dt.datetime.now()
    start = end - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
    params = {"fonKod": code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        response = requests.get(url, params=params, headers=headers, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
        values = response.json().get("value")
        if not values:
            return None
        df = pd.DataFrame(values)
        df["date"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
        df["price"] = df["Fiyat"].apply(parse_number)
        df = df.dropna(subset=["date", "price"])[df["price"] > 0]
        if len(df) < 2:
            return None
        return (
            df.sort_values("date")
            .drop_duplicates(subset=["date"], keep="last")
            .tail(MAX_DAYS + 1)
            .reset_index(drop=True)[["date", "price"]]
        )
    except Exception:
        return None


def compute_fund_metrics(series: Optional[pd.DataFrame]) -> Optional[dict]:
    if series is None or len(series) < 2:
        return None
    df = series.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["price"] = df["price"].apply(parse_number)
    df = df.dropna(subset=["date", "price"])[df["price"] > 0]
    df = df.sort_values("date").drop_duplicates(subset=["date"], keep="last").reset_index(drop=True)
    if len(df) < 2:
        return None

    prices = df["price"].astype(float).tolist()
    dates = df["date"].dt.strftime("%d.%m").tolist()
    daily_returns = [
        (current / previous - 1) * 100 if previous > 0 else 0.0
        for previous, current in zip(prices[:-1], prices[1:])
    ]

    return {"dates": dates[1:], "prices": prices, "daily_returns": daily_returns, "n_days": len(daily_returns)}


def zscore(values: list[float]) -> list[float]:
    clean = [v for v in values if v is not None and pd.notna(v)]
    if not clean:
        return [0.0] * len(values)
    mean_val = sum(clean) / len(clean)
    std = (sum((v - mean_val) ** 2 for v in clean) / len(clean)) ** 0.5
    if std < 1e-12:
        return [0.0] * len(values)
    return [(v - mean_val) / std for v in values]


def calculate_period_scores(funds, days):
    """Belirli bir gün sayısı için fonları kendi aralarında yarıştırır.

    FIX (bug #2): Eskiden yetersiz veriye sahip fonlar için 0.0 değerleri
    means/sharpes/cums listelerine ekleniyordu; bu sahte sıfırlar geçerli
    fonların z-skor ortalama/std hesabını bozuyordu. Artık sadece yeterli
    veriye sahip fonlar z-skor havuzuna giriyor; diğerleri baştan
    "Yetersiz Veri" olarak işaretlenip havuzun dışında tutuluyor.
    """
    means, sharpes, cums = [], [], []
    valid_indices = []

    for idx, item in enumerate(funds):
        if len(item["daily_returns"]) < days:
            item[f"score_{days}"] = 0
            item[f"karar_{days}"] = "Yetersiz Veri"
            continue

        slice_ret = item["daily_returns"][-days:]
        slice_prices = item["prices"][-(days + 1):]

        m_ret = sum(slice_ret) / len(slice_ret)
        vol = (sum((r - m_ret) ** 2 for r in slice_ret) / len(slice_ret)) ** 0.5
        shp = (m_ret / vol) if vol > 1e-12 else 0.0
        cum = (slice_prices[-1] / slice_prices[0] - 1) * 100

        item[f"mean_{days}"] = m_ret
        item[f"vol_{days}"] = vol
        item[f"shp_{days}"] = shp
        item[f"cum_{days}"] = cum

        means.append(m_ret)
        sharpes.append(shp)
        cums.append(cum)
        valid_indices.append(idx)

    z_m = zscore(means)
    z_s = zscore(sharpes)
    z_c = zscore(cums)

    for zi, idx in enumerate(valid_indices):
        item = funds[idx]

        # FIX (bug #1): "valor" artık compute_fund_metrics çıktısında değil,
        # doğrudan Excel'deki Fon_Listesi sayfasının B sütunundan (varsa)
        # okunup her fon kaydına ekleniyor (bkz. valor_map). B sütunu boşsa
        # veya sayfa tek sütunluysa ceza 0 olarak kalır.
        val = item.get("valor", 0) or 0
        v_pen = (val * 0.5) if val else 0.0

        raw_score = 50 + 15 * z_m[zi] + 20 * z_s[zi] + 15 * z_c[zi] - v_pen
        score = int(round(max(0.0, min(100.0, raw_score))))

        item[f"score_{days}"] = score
        if score >= 60:
            item[f"karar_{days}"] = "GÜÇLÜ AL"
        elif score >= 40:
            item[f"karar_{days}"] = "ASIL LİSTE"
        elif score >= 25:
            item[f"karar_{days}"] = "NÖTR"
        else:
            item[f"karar_{days}"] = "ACİL SAT"


# --- ARAYÜZ VE ANA ÇALIŞMA MANTIĞI ---

uploaded_file = st.file_uploader("Excel Dosyanızı Yükleyin (Fon_Listesi içeren):", type=["xlsx"])
if not uploaded_file:
    st.stop()

wb = openpyxl.load_workbook(uploaded_file)
if "Fon_Listesi" not in wb.sheetnames:
    st.error("Dosyada 'Fon_Listesi' sayfası yok!")
    st.stop()

ws_list = wb["Fon_Listesi"]
requested_codes = []
valor_map = {}  # FIX (bug #1): fon kodu -> valör (gün) eşlemesi

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if row and row[0].value:
        code = normalize_fund_code(row[0].value)
        requested_codes.append(code)
        # B sütununda valör (gün sayısı) varsa oku; yoksa 0 kabul et.
        valor_val = None
        if len(row) > 1 and row[1].value is not None:
            valor_val = parse_number(row[1].value)
        valor_map[code] = valor_val if valor_val is not None else 0.0

requested_codes = list(dict.fromkeys(requested_codes))
today = dt.date.today()
start_date = today - dt.timedelta(days=LOOKBACK_CALENDAR_DAYS)

with st.spinner("Veriler çekiliyor ve hesaplanıyor..."):
    universe = fetch_tefas_universe(start_date, today)
    calculated_funds = []

    for code in requested_codes:
        series = get_fund_series(universe, code) if not universe.empty else None
        source = "TEFAS" if series is not None else "Bulunamadı"
        if series is None:
            series = fetch_isyatirim_series(code)
            if series is not None:
                source = "İş Yatırım"

        metrics = compute_fund_metrics(series)
        if metrics:
            metrics["code"] = code
            metrics["source"] = source
            metrics["valor"] = valor_map.get(code, 0.0)  # FIX (bug #1)
            calculated_funds.append(metrics)

# Vade Skorlarını Hesapla
for period_name, period_days in PERIODS.items():
    calculate_period_scores(calculated_funds, period_days)

# Excel Oluşturma
output = io.BytesIO()
if "Vade_Analizi" in wb.sheetnames:
    del wb["Vade_Analizi"]
ws_out = wb.create_sheet("Vade_Analizi", 0)

headers = ["Fon Kodu"]
for period_name in PERIODS:
    label = PERIOD_LABELS[period_name]
    headers += [f"{label} Skor", f"{label} Karar", f"{label} Kümülatif %"]
ws_out.append(headers)

for cell in ws_out[1]:
    cell.fill = PatternFill(start_color=COLOR_NAVY, fill_type="solid")
    cell.font = Font(color=COLOR_WHITE, bold=True)
    cell.alignment = Alignment(horizontal="center")

# Verileri Ekle
for item in calculated_funds:
    row = [item["code"]]
    for period_name, period_days in PERIODS.items():
        row += [
            item.get(f"score_{period_days}", 0),
            item.get(f"karar_{period_days}", "-"),
            item.get(f"cum_{period_days}", 0),
        ]
    ws_out.append(row)

# Koşullu Biçimlendirme (Karar / Kümülatif % Sütunları İçin)
green_font = Font(color=COLOR_GREEN, bold=True)
red_font = Font(color=COLOR_RED, bold=True)
yellow_font = Font(color=COLOR_YELLOW, bold=True)

n_periods = len(PERIODS)
# Sütun düzeni: 1=Fon Kodu, sonra her vade için (Skor, Karar, Kümülatif%) üçlüsü
decision_cols = [3 + i * 3 for i in range(n_periods)]  # 3, 6, 9, 12 ...
percent_cols = [4 + i * 3 for i in range(n_periods)]  # 4, 7, 10, 13 ...

for row in range(2, ws_out.max_row + 1):
    for col in decision_cols:
        val = str(ws_out.cell(row=row, column=col).value)
        if "AL" in val or "LİSTE" in val:
            ws_out.cell(row=row, column=col).font = green_font
        elif "NÖTR" in val:
            ws_out.cell(row=row, column=col).font = yellow_font
        elif "SAT" in val:
            ws_out.cell(row=row, column=col).font = red_font

    for col in percent_cols:
        ws_out.cell(row=row, column=col).number_format = '0.00"%"'

# Sütun Genişlikleri
for col in ws_out.columns:
    ws_out.column_dimensions[col[0].column_letter].width = 18

wb.save(output)
output.seek(0)

# Arayüz Gösterimi
st.success("✅ Multi-Vade Analizi Tamamlandı!")
st.download_button(
    "📥 Kapsamlı Excel Çıktısını İndir",
    data=output,
    file_name="fon_vade_analizi.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# Tablo Önizleme
df_preview = pd.DataFrame(ws_out.values)
df_preview.columns = df_preview.iloc[0]
df_preview = df_preview[1:]
st.dataframe(df_preview, use_container_width=True, hide_index=True)
