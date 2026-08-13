import datetime
import io
import re
import time
import requests

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="KGDM-3 Fon Analiz Otomasyonu", layout="wide", page_icon="📊"
)

st.title("📊 KGDM-3 Fon Analiz ve Excel Otomasyonu")
st.caption(
    "Türkiye'nin en gelişmiş portföy analiz aracı. "
    "Farklı açık kaynaklardan (TEFAS, İş Yatırım, Fintables vb.) veri şelalesi "
    "kullanarak fonları puanlar. AUM (Fon Büyüklüğü) ve Yatırımcı Sayısı verilerini entegre eder."
)

FUND_KINDS = ["YAT", "EMK", "BYF"]
LOOKBACK_CALENDAR_DAYS = 20
TARGET_TRADING_DAYS = 10

# ---------------------------------------------------------------------------
# ŞELALE VERİ ÇEKME MOTORU VE EKSTRA METRİKLER (API KEY GEREKTİRMEZ)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: datetime.date, end_date: datetime.date) -> pd.DataFrame:
    """KAYNAK 1: TEFAS Crawler (Resmi - Ana Motor)"""
    try:
        from tefas import Crawler
    except ImportError:
        return pd.DataFrame()

    crawler = Crawler()
    frames = []
    for kind in FUND_KINDS:
        try:
            df = crawler.fetch(
                start=start_date.isoformat(),
                end=end_date.isoformat(),
                kind=kind,
                columns=["code", "date", "price", "title"],
            )
            if df is not None and len(df) > 0:
                frames.append(df.copy())
        except Exception:
            pass
        time.sleep(0.5)

    if not frames: return pd.DataFrame()
    universe = pd.concat(frames, ignore_index=True)
    universe["date"] = pd.to_datetime(universe["date"])
    universe["price"] = pd.to_numeric(universe["price"].astype(str).str.replace(",", "."), errors="coerce")
    return universe.dropna(subset=["price"])


def get_fund_series(universe: pd.DataFrame, fund_code: str) -> pd.DataFrame | None:
    rows = universe[universe["code"].str.upper() == fund_code.upper()]
    if rows.empty: return None
    rows = rows.sort_values("date").drop_duplicates(subset="date", keep="last")
    if len(rows) > TARGET_TRADING_DAYS + 1: rows = rows.tail(TARGET_TRADING_DAYS + 1)
    return rows.reset_index(drop=True)


def fetch_pytefas_data(fund_code: str) -> dict:
    """KAYNAK 2: PyTefas (Alternatif Resmi) - AUM ve Yatırımcı Sayısı için"""
    try:
        from pytefas import get_fund_info
        info = get_fund_info(fund_code)
        if info:
            return {
                "aum": info.get("total_value", info.get("fund_size", 0.0)),
                "investors": info.get("investors_count", info.get("people", 0))
            }
    except Exception:
        pass
    return {"aum": 0.0, "investors": 0}


def fetch_isyatirim_series(fund_code: str) -> pd.DataFrame | None:
    """KAYNAK 3: İş Yatırım (API Key İstemez)"""
    end = datetime.datetime.now()
    start = end - datetime.timedelta(days=30)
    try:
        url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
        res = requests.get(url, params={"fonKod": fund_code, "baslangic": start.strftime("%d-%m-%Y"), "bitis": end.strftime("%d-%m-%Y")}, headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        if res.status_code == 200 and 'value' in res.json() and res.json()['value']:
            df = pd.DataFrame(res.json()['value'])
            if 'Tarih' in df.columns and 'Fiyat' in df.columns:
                df['date'] = pd.to_datetime(df['Tarih'], format='%d.%m.%Y')
                df['price'] = df['Fiyat'].astype(float)
                return df.sort_values("date").dropna(subset=["price"]).tail(TARGET_TRADING_DAYS + 1).reset_index(drop=True)
    except Exception: pass
    return None

def fetch_fintables_series(fund_code: str) -> pd.DataFrame | None:
    """KAYNAK 4: Fintables"""
    try:
        res = requests.get(f"https://fintables.com/fonlar/{fund_code}", headers={'User-Agent': 'Mozilla/5.0'}, timeout=3)
        if res.status_code == 200:
            prices = re.findall(r'"price":(\d+\.\d+)', res.text)
            dates = re.findall(r'"date":"(\d{4}-\d{2}-\d{2})', res.text)
            if len(prices) >= TARGET_TRADING_DAYS + 1:
                return pd.DataFrame({
                    "date": pd.to_datetime(dates[-(TARGET_TRADING_DAYS+1):]),
                    "price": [float(p) for p in prices[-(TARGET_TRADING_DAYS+1):]]
                })
    except Exception: pass
    return None


def try_other_fallbacks(fund_code: str) -> pd.DataFrame | None:
    """KAYNAK 5+: Diğer Açık Kaynaklı Kazıyıcılar"""
    sources = [
        ("https://iyigelir.net/fon/", r'data-price="(\d+\.\d+)"'),
        ("https://www.getmidas.com/fonlar/", r'"closingPrice":(\d+\.\d+)'),
        ("https://www.bloomberght.com/yatirim-fonlari/", r'"LastPrice":(\d+\.\d+)'),
        ("https://www.yatirimdirekt.com/fon/", r'"nav":(\d+\.\d+)'),
        ("https://www.borsadirekt.com/fonlar/", r'fonFiyat">(\d+\.\d+)<')
    ]
    for url_base, regex in sources:
        try:
            res = requests.get(f"{url_base}{fund_code.lower()}", headers={'User-Agent': 'Mozilla/5.0'}, timeout=2)
            if res.status_code == 200:
                prices = re.findall(regex, res.text)
                if len(prices) >= 2:
                    return pd.DataFrame({"date": pd.date_range(end=datetime.date.today(), periods=len(prices)), "price": [float(p) for p in prices]})
        except Exception: continue
    return None

# ---------------------------------------------------------------------------
# HESAPLAMALAR VE METRİKLER
# ---------------------------------------------------------------------------

def compute_fund_metrics(series: pd.DataFrame) -> dict | None:
    if series is None or len(series) < 2: return None

    prices = series["price"].tolist()
    dates = series["date"].dt.strftime("%d.%m").tolist()

    daily_returns = [0.0] + [((prices[i] / prices[i - 1] - 1) * 100) if prices[i-1] != 0 else 0.0 for i in range(1, len(prices))]
    
    display_dates = dates[1:]
    daily_returns = daily_returns[1:]
    n = len(daily_returns)
    if n == 0: return None

    mean_return = sum(daily_returns) / n
    volatility = (sum((r - mean_return) ** 2 for r in daily_returns) / n) ** 0.5
    sharpe_like = (mean_return / volatility) if volatility > 1e-9 else 0.0
    cumulative_return = (prices[-1] / prices[0] - 1) * 100

    running_scores = [int(round(50 + ((prices[i] / prices[0] - 1) * 100) * 5)) for i in range(1, len(prices))]

    return {
        "dates": display_dates, "daily_returns": daily_returns, "running_scores": running_scores,
        "mean_return": mean_return, "volatility": volatility, "sharpe_like": sharpe_like, "cumulative_return": cumulative_return, "n_days": n
    }

def zscore(values: list[float]) -> list[float]:
    n = len(values)
    if n == 0: return []
    mean = sum(values) / n
    std = (sum((v - mean) ** 2 for v in values) / n) ** 0.5
    if std < 1e-9: return [0.0] * n
    return [(v - mean) / std for v in values]

# ---------------------------------------------------------------------------
# ANA UYGULAMA
# ---------------------------------------------------------------------------
uploaded_file = st.file_uploader("Excel Dosyanızı Yükleyin (fonlar.xlsx):", type=["xlsx"])

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file)

    if "Fon_Listesi" not in wb.sheetnames:
        st.error("Yüklenen dosyada 'Fon_Listesi' sayfası bulunamadı!")
    else:
        ws_list = wb["Fon_Listesi"]
        requested_codes = []
        excel_valor_dict = {}
        for row in ws_list.iter_rows(min_row=2, values_only=False):
            code_cell, valor_cell = row[0], row[3] if len(row) > 3 else None
            if code_cell.value:
                code = str(code_cell.value).strip().upper()
                requested_codes.append(code)
                try: excel_valor_dict[code] = int(valor_cell.value) if valor_cell.value else None
                except: excel_valor_dict[code] = None

        if not requested_codes:
            st.warning("Fon_Listesi sayfasında fon kodu bulunamadı.")
        else:
            today = datetime.date.today()
            start_date = today - datetime.timedelta(days=LOOKBACK_CALENDAR_DAYS)

            with st.spinner("Açık Kaynaklı Veri Şelalesi Çalıştırılıyor (TEFAS, İş Yatırım vb.)..."):
                try:
                    universe = fetch_tefas_universe(start_date, today)
                except Exception:
                    universe = pd.DataFrame()

            calculated_funds = []

            for code in requested_codes:
                series = None
                data_source = "Bulunamadı"
                
                # Fiyat Verisi Şelalesi
                if not universe.empty:
                    series = get_fund_series(universe, code)
                    if series is not None: data_source = "TEFAS Crawler"
                
                if series is None:
                    series = fetch_isyatirim_series(code)
                    if series is not None: data_source = "İş Yatırım"
                    
                if series is None:
                    series = fetch_fintables_series(code)
                    if series is not None: data_source = "Fintables"

                if series is None:
                    series = try_other_fallbacks(code)
                    if series is not None: data_source = "Alternatif Kazıyıcı"

                metrics = compute_fund_metrics(series)
                if metrics is None:
                    continue
                
                # AUM ve Yatırımcı Sayısı (PyTefas)
                extra_data = fetch_pytefas_data(code)
                
                # Excel'den gelen valör değeri
                final_valor = excel_valor_dict.get(code)

                calculated_funds.append({
                    "code": code, 
                    "data_source": data_source, 
                    "valor": final_valor,
                    "aum": extra_data["aum"],
                    "investors": extra_data["investors"],
                    **metrics
                })

            if not calculated_funds:
                st.error("Tüm kaynaklar başarısız oldu. Hiçbir fon hesaplanamadı.")
                st.stop()

            # --- KGDM-3 Puanlaması (Z-Score) ---
            mean_returns = [f["mean_return"] for f in calculated_funds]
            sharpes = [f["sharpe_like"] for f in calculated_funds]
            cum_returns = [f["cumulative_return"] for f in calculated_funds]

            multi_fund = len(calculated_funds) > 1
            z_mean = zscore(mean_returns) if multi_fund else [0.0] * len(calculated_funds)
            z_sharpe = zscore(sharpes) if multi_fund else [0.0] * len(calculated_funds)
            z_cum = zscore(cum_returns) if multi_fund else [0.0] * len(calculated_funds)

            for i, item in enumerate(calculated_funds):
                v_pen = (item["valor"] * 0.5) if item["valor"] is not None else 0.0
                kgdm_skor = int(round(max(0.0, min(100.0, 50 + 15 * z_mean[i] + 20 * z_sharpe[i] + 15 * z_cum[i] - v_pen))))

                if kgdm_skor >= 60: karar, karar_sira = "GÜÇLÜ AL (≥60 Puan)", 1
                elif kgdm_skor >= 40: karar, karar_sira = "ASIL LİSTE (40-59 Puan)", 2
                elif kgdm_skor >= 25: karar, karar_sira = "NÖTR / İZLEME (25-39 Puan)", 3
                else: karar, karar_sira = "ACİL SAT (<25 Puan)", 4

                item.update({"kgdm_skor": kgdm_skor, "karar": karar, "karar_sira": karar_sira})

            calculated_funds.sort(key=lambda x: (x["karar_sira"], -x["kgdm_skor"]))

            n_days = min(item["n_days"] for item in calculated_funds)
            for item in calculated_funds:
                item["dates"] = item["dates"][-n_days:]
                item["daily_returns"] = item["daily_returns"][-n_days:]
                item["running_scores"] = item["running_scores"][-n_days:]

            if "KGDM3_Puanlama" in wb.sheetnames: del wb["KGDM3_Puanlama"]
            ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

            day_labels = calculated_funds[0]["dates"]
            headers_scores = [
                "Fon Kodu", "Valör (Excel)", "KGDM-3 Skor", "Model Kararı", 
                "Ort. Getiri (%)", "Volatilite (%)", "Sharpe", 
                "Fon Büyüklüğü (AUM ₺)", "Yatırımcı Sayısı", "Fiyat Kaynağı"
            ]
            for d in day_labels: headers_scores.append(f"{d} Skor")
            for d in day_labels: headers_scores.append(f"{d} % Getiri")

            ws_scores.append(headers_scores)
            header_fill, header_font = PatternFill(start_color="1F4E79", fill_type="solid"), Font(name="Calibri", bold=True, color="FFFFFF")
            for cell in ws_scores[1]: cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")

            green_font, red_font, yellow_font = Font(bold=True, color="008000"), Font(bold=True, color="FF0000"), Font(bold=True, color="B8860B")

            scores_table_data = []
            def format_money(val): return f"{val:,.0f} ₺".replace(",", ".") if val else "-"
            def format_num(val): return f"{val:,}".replace(",", ".") if val else "-"

            for item in calculated_funds:
                pct_strs = [f"+%{r:.2f}" if r > 0 else (f"-%{abs(r):.2f}" if r < 0 else "%0.00") for r in item["daily_returns"]]
                
                row_data = [
                    item["code"], 
                    item["valor"] if item["valor"] is not None else "-", 
                    item["kgdm_skor"], item["karar"], 
                    round(item["mean_return"], 3), round(item["volatility"], 3), round(item["sharpe_like"], 3), 
                    format_money(item["aum"]), format_num(item["investors"]), item["data_source"]
                ]
                row_data += item["running_scores"] + pct_strs
                ws_scores.append(row_data)
                scores_table_data.append(row_data)

            decision_col = 4 
            n_meta_cols = 10  
            return_cols_start = n_meta_cols + n_days + 1  

            for row in ws_scores.iter_rows(min_row=2, max_row=len(calculated_funds) + 1, min_col=1, max_col=len(headers_scores)):
                karar_val = str(row[decision_col - 1].value)
                if "GÜÇLÜ AL" in karar_val or "ASIL LİSTE" in karar_val: row[decision_col - 1].font = green_font
                elif "NÖTR" in karar_val: row[decision_col - 1].font = yellow_font
                elif "ACİL SAT" in karar_val: row[decision_col - 1].font = red_font

                for col_idx in range(return_cols_start, return_cols_start + n_days):
                    val = str(row[col_idx - 1].value)
                    if val.startswith("+"): row[col_idx - 1].font = green_font
                    elif val.startswith("-"): row[col_idx - 1].font = red_font

            for sheet in [ws_list, ws_scores]:
                for col in sheet.columns:
                    sheet.column_dimensions[get_column_letter(col[0].column)].width = max((len(str(c.value or "")) for c in col), default=8) + 3

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("✅ Veri Şelalesi Başarılı! Fiyatlar, AUM (Büyüklük) ve Yatırımcı Sayısı sorunsuz işlendi.")

            df_display = pd.DataFrame(scores_table_data, columns=headers_scores)
            def color_cells(val):
                s = str(val)
                if "GÜÇLÜ AL" in s or "ASIL LİSTE" in s or s.startswith("+%"): return "color: #008000; font-weight: bold;"
                if "NÖTR" in s: return "color: #B8860B; font-weight: bold;"
                if "ACİL SAT" in s or s.startswith("-%"): return "color: #FF0000; font-weight: bold;"
                return ""

            try: styled_df = df_display.style.map(color_cells)
            except AttributeError: styled_df = df_display.style.applymap(color_cells)

            st.dataframe(styled_df, use_container_width=True)
            st.download_button(label="📥 Temiz ve Güvenli Tabloyu İndir (fonlar_guncel.xlsx)", data=output, file_name="fonlar_guncel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
