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
    "Excel dosyasındaki fonları analiz eder. TEFAS engellerine karşı "
    "**4 farklı kaynaktan** (TEFAS, Fonoloji, Fintables, İş Yatırım) "
    "canlı veri toplayarak skorları ve % getirileri hesaplar."
)

FUND_KINDS = ["YAT", "EMK", "BYF"]
LOOKBACK_CALENDAR_DAYS = 20
TARGET_TRADING_DAYS = 10
FONOLOJI_BASE_URL = "https://fonoloji.com/v1"

with st.sidebar:
    st.subheader("⚙️ Ayarlar")
    fonoloji_api_key = st.text_input(
        "Fonoloji API Anahtarı (isteğe bağlı, yedek kaynak için)",
        type="password",
        help=(
            "TEFAS'tan veri çekilemeyen fonlar için 1. Yedek kaynak olarak "
            "kullanılır. Ücretsiz anahtar: https://fonoloji.com/kayit"
        ),
    )

# ---------------------------------------------------------------------------
# 4 KATMANLI VERİ ÇEKME MOTORU
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: datetime.date, end_date: datetime.date) -> pd.DataFrame:
    """KAYNAK 1: TEFAS (Ana Kaynak) - Tüm evreni tek seferde çeker."""
    try:
        from tefas import Crawler
    except ImportError as exc:
        raise RuntimeError("tefas-crawler kütüphanesi kurulu değil.") from exc

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
                df = df.copy()
                df["kind"] = kind
                frames.append(df)
        except Exception as exc:
            st.warning(f"TEFAS'tan '{kind}' tipi fonlar çekilirken hata: {exc}")
        time.sleep(1)

    if not frames:
        raise RuntimeError("TEFAS'tan hiçbir veri alınamadı. Yedek kaynaklar devreye girecek.")

    universe = pd.concat(frames, ignore_index=True)
    universe["date"] = pd.to_datetime(universe["date"])
    universe["price"] = pd.to_numeric(
        universe["price"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    return universe.dropna(subset=["price"])

def get_fund_series(universe: pd.DataFrame, fund_code: str) -> pd.DataFrame | None:
    rows = universe[universe["code"].str.upper() == fund_code.upper()]
    if rows.empty:
        return None
    rows = rows.sort_values("date").drop_duplicates(subset="date", keep="last")
    if len(rows) > TARGET_TRADING_DAYS + 1:
        rows = rows.tail(TARGET_TRADING_DAYS + 1)
    return rows.reset_index(drop=True)


def fetch_fonoloji_series(fund_code: str, api_key: str) -> pd.DataFrame | None:
    """KAYNAK 2: Fonoloji API (1. Yedek)"""
    headers = {"X-API-Key": api_key}
    try:
        detail_resp = requests.get(f"{FONOLOJI_BASE_URL}/funds/{fund_code}", headers=headers, timeout=5)
        if detail_resp.status_code != 200:
            return None
        
        fund_name = detail_resp.json().get("fund", {}).get("name", fund_code)
        
        hist_resp = requests.get(
            f"{FONOLOJI_BASE_URL}/funds/{fund_code}/history",
            headers=headers, params={"period": "1m"}, timeout=5
        )
        if hist_resp.status_code != 200:
            return None
            
        points = hist_resp.json().get("points", [])
        if not points:
            return None

        df = pd.DataFrame(points)
        df["date"] = pd.to_datetime(df["date"])
        df["price"] = pd.to_numeric(df["price"], errors="coerce")
        df = df.dropna(subset=["price"]).sort_values("date")
        df["title"] = fund_name
        df["code"] = fund_code

        if len(df) > TARGET_TRADING_DAYS + 1:
            df = df.tail(TARGET_TRADING_DAYS + 1)
        return df.reset_index(drop=True)
    except Exception:
        return None


def fetch_fintables_series(fund_code: str) -> pd.DataFrame | None:
    """KAYNAK 3: Fintables Web Analizi (2. Yedek)"""
    try:
        url = f"https://fintables.com/fonlar/{fund_code}"
        res = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=4)
        if res.status_code == 200:
            # HTML içindeki JSON verilerinden fiyat dizisini bulur
            prices = re.findall(r'"price":(\d+\.\d+)', res.text)
            dates = re.findall(r'"date":"(\d{4}-\d{2}-\d{2})', res.text)
            
            if len(prices) >= TARGET_TRADING_DAYS + 1:
                df = pd.DataFrame({
                    "date": pd.to_datetime(dates[-(TARGET_TRADING_DAYS+1):]),
                    "price": [float(p) for p in prices[-(TARGET_TRADING_DAYS+1):]],
                    "code": fund_code,
                    "title": f"{fund_code} Yatırım Fonu"
                })
                return df
    except Exception:
        pass
    return None


def fetch_isyatirim_series(fund_code: str) -> pd.DataFrame | None:
    """KAYNAK 4: İş Yatırım Alternatif Fon Verisi (3. Yedek)"""
    try:
        # Son 30 günün tarihlerini hesapla
        end = datetime.datetime.now()
        start = end - datetime.timedelta(days=30)
        
        url = f"https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/YatirimFonGecmisGetiri"
        params = {
            "fonKod": fund_code,
            "baslangic": start.strftime("%d-%m-%Y"),
            "bitis": end.strftime("%d-%m-%Y")
        }
        res = requests.get(url, params=params, headers={'User-Agent': 'Mozilla/5.0'}, timeout=4)
        if res.status_code == 200:
            data = res.json()
            if 'value' in data and len(data['value']) > 0:
                df = pd.DataFrame(data['value'])
                # Sütun isimleri İş Yatırım yapısına göre eşleştirilir
                if 'Tarih' in df.columns and 'Fiyat' in df.columns:
                    df['date'] = pd.to_datetime(df['Tarih'], format='%d.%m.%Y')
                    df['price'] = df['Fiyat'].astype(float)
                    df['code'] = fund_code
                    df['title'] = f"{fund_code} Yatırım Fonu"
                    
                    df = df.sort_values("date").dropna(subset=["price"])
                    if len(df) > TARGET_TRADING_DAYS + 1:
                        df = df.tail(TARGET_TRADING_DAYS + 1)
                    return df.reset_index(drop=True)
    except Exception:
        pass
    return None

# ---------------------------------------------------------------------------
# HESAPLAMALAR VE METRİKLER
# ---------------------------------------------------------------------------

def compute_fund_metrics(series: pd.DataFrame) -> dict | None:
    if series is None or len(series) < 2:
        return None

    prices = series["price"].tolist()
    dates = series["date"].dt.strftime("%d.%m").tolist()

    daily_returns = []
    for i in range(1, len(prices)):
        prev, curr = prices[i - 1], prices[i]
        if prev == 0:
            daily_returns.append(0.0)
        else:
            daily_returns.append((curr / prev - 1) * 100)

    display_dates = dates[1:]
    n = len(daily_returns)
    if n == 0:
        return None

    mean_return = sum(daily_returns) / n
    variance = sum((r - mean_return) ** 2 for r in daily_returns) / n
    volatility = variance ** 0.5
    sharpe_like = (mean_return / volatility) if volatility > 1e-9 else 0.0
    cumulative_return = (prices[-1] / prices[0] - 1) * 100

    running_scores = []
    for i in range(1, len(prices)):
        cum = (prices[i] / prices[0] - 1) * 100
        running_scores.append(int(round(50 + cum * 5))) 

    return {
        "dates": display_dates,
        "daily_returns": daily_returns,
        "running_scores": running_scores,
        "mean_return": mean_return,
        "volatility": volatility,
        "sharpe_like": sharpe_like,
        "cumulative_return": cumulative_return,
        "fund_name": series["title"].iloc[-1] if "title" in series.columns else None,
        "n_days": n,
    }


def zscore(values: list[float]) -> list[float]:
    n = len(values)
    if n == 0:
        return []
    mean = sum(values) / n
    variance = sum((v - mean) ** 2 for v in values) / n
    std = variance ** 0.5
    if std < 1e-9:
        return [0.0] * n
    return [(v - mean) / std for v in values]


# ---------------------------------------------------------------------------
# ANA UYGULAMA (EXCEL İŞLEME)
# ---------------------------------------------------------------------------
uploaded_file = st.file_uploader("Excel Dosyanızı Yükleyin (fonlar.xlsx):", type=["xlsx"])

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file)

    if "Fon_Listesi" not in wb.sheetnames:
        st.error("Yüklenen dosyada 'Fon_Listesi' sayfası bulunamadı!")
    else:
        ws_list = wb["Fon_Listesi"]

        requested_codes = []
        valor_by_code = {}
        for row in ws_list.iter_rows(min_row=2, values_only=False):
            code_cell, valor_cell = row[0], row[3] if len(row) > 3 else None
            if code_cell.value:
                code = str(code_cell.value).strip().upper()
                requested_codes.append(code)
                if valor_cell is not None and valor_cell.value is not None:
                    try:
                        valor_by_code[code] = int(valor_cell.value)
                    except (TypeError, ValueError):
                        valor_by_code[code] = None
                else:
                    valor_by_code[code] = None

        if not requested_codes:
            st.warning("Fon_Listesi sayfasında fon kodu bulunamadı.")
        else:
            today = datetime.date.today()
            start_date = today - datetime.timedelta(days=LOOKBACK_CALENDAR_DAYS)

            with st.spinner("Ana Kaynaktan (TEFAS) gerçek fiyat verisi çekiliyor..."):
                try:
                    universe = fetch_tefas_universe(start_date, today)
                except RuntimeError as exc:
                    st.error(str(exc))
                    universe = pd.DataFrame() # Boş dataframe atayarak yedeklere geç

            calculated_funds = []
            not_found = []
            missing_valor = []
            source_tracker = {"TEFAS": [], "Fonoloji": [], "Fintables": [], "İş Yatırım": []}

            # VERİ KAYNAKLARI ZİNCİRİ
            for code in requested_codes:
                series = None
                data_source = "Bilinmiyor"
                
                # 1. Kaynak: TEFAS Evreni
                if not universe.empty:
                    series = get_fund_series(universe, code)
                    if series is not None: data_source = "TEFAS"
                
                # 2. Kaynak: Fonoloji
                if series is None and fonoloji_api_key:
                    series = fetch_fonoloji_series(code, fonoloji_api_key)
                    if series is not None: data_source = "Fonoloji"
                
                # 3. Kaynak: Fintables
                if series is None:
                    series = fetch_fintables_series(code)
                    if series is not None: data_source = "Fintables"
                    
                # 4. Kaynak: İş Yatırım
                if series is None:
                    series = fetch_isyatirim_series(code)
                    if series is not None: data_source = "İş Yatırım"

                # Metrikleri hesapla
                metrics = compute_fund_metrics(series)

                if metrics is None:
                    not_found.append(code)
                    continue
                
                source_tracker[data_source].append(code)

                if valor_by_code.get(code) is None:
                    missing_valor.append(code)

                calculated_funds.append({
                    "code": code,
                    "data_source": data_source,
                    **metrics,
                })

            if not calculated_funds:
                st.error("4 farklı kaynak taranmasına rağmen hiçbir veri bulunamadı.")
                st.stop()
                
            # Kullanılan kaynakları raporla
            for source, f_list in source_tracker.items():
                if f_list:
                    st.info(f"**{source}** kullanılarak çekilen fonlar: {', '.join(f_list)}")

            # --- Göreli (z-skor tabanlı) KGDM-3 skoru ---
            mean_returns = [f["mean_return"] for f in calculated_funds]
            sharpes = [f["sharpe_like"] for f in calculated_funds]
            cum_returns = [f["cumulative_return"] for f in calculated_funds]

            multi_fund = len(calculated_funds) > 1
            z_mean = zscore(mean_returns) if multi_fund else [0.0] * len(calculated_funds)
            z_sharpe = zscore(sharpes) if multi_fund else [0.0] * len(calculated_funds)
            z_cum = zscore(cum_returns) if multi_fund else [0.0] * len(calculated_funds)

            if not multi_fund:
                st.info("Sadece 1 fon olduğu için göreli skorlama hesaplanamadı.")

            for i, item in enumerate(calculated_funds):
                valor = valor_by_code.get(item["code"])
                valor_penalty = (valor * 0.5) if valor is not None else 0.0

                kgdm_skor = 50 + 15 * z_mean[i] + 20 * z_sharpe[i] + 15 * z_cum[i] - valor_penalty
                kgdm_skor = int(round(max(0.0, min(100.0, kgdm_skor))))

                if kgdm_skor >= 60:
                    karar, karar_sira = "GÜÇLÜ AL (≥60 Puan)", 1
                elif kgdm_skor >= 40:
                    karar, karar_sira = "ASIL LİSTE (40-59 Puan)", 2
                elif kgdm_skor >= 25:
                    karar, karar_sira = "NÖTR / İZLEME (25-39 Puan)", 3
                else:
                    karar, karar_sira = "ACİL SAT (<25 Puan)", 4

                item.update({
                    "valor": valor if valor is not None else "-",
                    "kgdm_skor": kgdm_skor,
                    "karar": karar,
                    "karar_sira": karar_sira,
                })

            calculated_funds.sort(key=lambda x: (x["karar_sira"], -x["kgdm_skor"]))

            # Farklı kaynaklardan gelen gün sayılarını hizalama
            n_days = min(item["n_days"] for item in calculated_funds)
            for item in calculated_funds:
                item["dates"] = item["dates"][-n_days:]
                item["daily_returns"] = item["daily_returns"][-n_days:]
                item["running_scores"] = item["running_scores"][-n_days:]

            # --- Excel sayfası oluşturma ---
            if "KGDM3_Puanlama" in wb.sheetnames:
                del wb["KGDM3_Puanlama"]
            ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

            day_labels = calculated_funds[0]["dates"]

            headers_scores = [
                "Fon Kodu", "Fon Adı", "Valör", "KGDM-3 Skor", "Model Kararı",
                "Ort. Günlük Getiri (%)", "Volatilite (%)", "Sharpe-benzeri Oran", "Veri Kaynağı"
            ]
            for d in day_labels:
                headers_scores.append(f"{d} Skor")
            for d in day_labels:
                headers_scores.append(f"{d} % Getiri")

            ws_scores.append(headers_scores)
            header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            for cell in ws_scores[1]:
                cell.fill, cell.font = header_fill, header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            green_font = Font(name="Calibri", bold=True, color="008000")
            red_font = Font(name="Calibri", bold=True, color="FF0000")
            yellow_font = Font(name="Calibri", bold=True, color="B8860B")

            scores_table_data = []
            for item in calculated_funds:
                pct_strs = [
                    f"+%{r:.2f}" if r > 0 else (f"-%{abs(r):.2f}" if r < 0 else "%0.00")
                    for r in item["daily_returns"]
                ]
                row_data = [
                    item["code"], item["fund_name"] or item["code"], item["valor"],
                    item["kgdm_skor"], item["karar"],
                    round(item["mean_return"], 3), round(item["volatility"], 3),
                    round(item["sharpe_like"], 3), item["data_source"]
                ]
                row_data += item["running_scores"]
                row_data += pct_strs
                ws_scores.append(row_data)
                scores_table_data.append(row_data)

            decision_col = 5 
            n_meta_cols = 9  
            return_cols_start = n_meta_cols + n_days + 1  

            for row in ws_scores.iter_rows(
                min_row=2, max_row=len(calculated_funds) + 1,
                min_col=1, max_col=len(headers_scores),
            ):
                karar_val = str(row[decision_col - 1].value)
                if "GÜÇLÜ AL" in karar_val or "ASIL LİSTE" in karar_val:
                    row[decision_col - 1].font = green_font
                elif "NÖTR" in karar_val:
                    row[decision_col - 1].font = yellow_font
                elif "ACİL SAT" in karar_val:
                    row[decision_col - 1].font = red_font

                for col_idx in range(return_cols_start, return_cols_start + n_days):
                    cell = row[col_idx - 1]
                    val = str(cell.value)
                    if val.startswith("+"):
                        cell.font = green_font
                    elif val.startswith("-"):
                        cell.font = red_font

            for sheet in [ws_list, ws_scores]:
                for col in sheet.columns:
                    width = max((len(str(c.value or "")) for c in col), default=8) + 3
                    sheet.column_dimensions[get_column_letter(col[0].column)].width = max(width, 12)

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success(
                f"✅ {len(calculated_funds)} fon **4 Farklı Güvenlik Kaynağı** (TEFAS, Fonoloji, Fintables, İş Yatırım) "
                f"kullanılarak başarılı şekilde tarandı ve analiz edildi."
            )

            df_display = pd.DataFrame(scores_table_data, columns=headers_scores)

            def color_cells(val):
                s = str(val)
                if "GÜÇLÜ AL" in s or "ASIL LİSTE" in s or s.startswith("+%"):
                    return "color: #008000; font-weight: bold;"
                if "NÖTR" in s:
                    return "color: #B8860B; font-weight: bold;"
                if "ACİL SAT" in s or s.startswith("-%"):
                    return "color: #FF0000; font-weight: bold;"
                return ""

            try:
                styled_df = df_display.style.map(color_cells)
            except AttributeError:
                styled_df = df_display.style.applymap(color_cells)

            st.dataframe(styled_df, use_container_width=True)
            st.download_button(
                label="📥 Güvenli Tabloyu İndir (fonlar_guncel.xlsx)",
                data=output,
                file_name="fonlar_guncel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
