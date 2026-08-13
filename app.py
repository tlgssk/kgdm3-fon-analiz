import datetime
import io
import time

import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="KGDM-3 Fon Analiz Otomasyonu", layout="wide", page_icon="📊"
)

st.title("📊 KGDM-3 Fon Analiz ve Excel Otomasyonu")
st.caption(
    "Excel dosyasındaki fonları gerçek TEFAS verisiyle (tefas-crawler) "
    "analiz eder; risk-ayarlı, göreli bir skorla gruplar. TEFAS'ta "
    "bulunamayan/başarısız olan fonlar için isteğe bağlı Fonoloji API "
    "yedeği kullanılır."
)

FUND_KINDS = ["YAT", "EMK", "BYF"]  # Menkul Kıymet Yatırım Fonu, Emeklilik, Borsa Yatırım Fonu
LOOKBACK_CALENDAR_DAYS = 20          # ~10 iş günü elde etmek için tampon
TARGET_TRADING_DAYS = 10

FONOLOJI_BASE_URL = "https://fonoloji.com/v1"

with st.sidebar:
    st.subheader("⚙️ Ayarlar")
    fonoloji_api_key = st.text_input(
        "Fonoloji API Anahtarı (isteğe bağlı, yedek kaynak için)",
        type="password",
        help=(
            "TEFAS'tan veri çekilemeyen fonlar için yedek kaynak olarak "
            "kullanılır. Ücretsiz anahtar: https://fonoloji.com/kayit — "
            "boş bırakılırsa TEFAS'ta bulunamayan fonlar analiz dışı kalır."
        ),
    )


@st.cache_data(show_spinner=False, ttl=60 * 30)
def fetch_tefas_universe(start_date: datetime.date, end_date: datetime.date) -> pd.DataFrame:
    """Belirtilen tarih aralığındaki TÜM fonları (her tip için tek istek) çeker.

    Fon kodu bazlı ayrı ayrı istek atmak yerine tek seferde tüm evreni çekip
    yerelde filtrelemek, TEFAS'ın dakikalık istek sınırına takılma riskini
    azaltır.
    """
    try:
        from tefas import Crawler
    except ImportError as exc:
        raise RuntimeError(
            "tefas-crawler kütüphanesi kurulu değil. Kurmak için: "
            "pip install tefas-crawler"
        ) from exc

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
        except Exception as exc:  # noqa: BLE001 - TEFAS tarafı hata verirse diğer tiplere devam et
            st.warning(f"TEFAS'tan '{kind}' tipi fonlar çekilirken hata oluştu: {exc}")
        time.sleep(1)  # nazik davran, hız sınırına takılma

    if not frames:
        raise RuntimeError(
            "TEFAS'tan hiçbir veri alınamadı. İnternet bağlantınızı veya "
            "tefas-crawler kütüphanesinin güncel olup olmadığını kontrol edin."
        )

    universe = pd.concat(frames, ignore_index=True)
    universe["date"] = pd.to_datetime(universe["date"])
    universe["price"] = pd.to_numeric(
        universe["price"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    universe = universe.dropna(subset=["price"])
    return universe


def get_fund_series(universe: pd.DataFrame, fund_code: str) -> pd.DataFrame | None:
    """Bir fon koduna ait fiyat serisini tarihe göre sıralı döndürür."""
    rows = universe[universe["code"].str.upper() == fund_code.upper()]
    if rows.empty:
        return None
    rows = rows.sort_values("date").drop_duplicates(subset="date", keep="last")
    if len(rows) > TARGET_TRADING_DAYS + 1:
        rows = rows.tail(TARGET_TRADING_DAYS + 1)
    return rows.reset_index(drop=True)


def compute_fund_metrics(series: pd.DataFrame) -> dict | None:
    """Gerçek fiyat serisinden günlük getiri, ortalama getiri, volatilite ve
    sharpe-benzeri oranı hesaplar. Look-ahead YOKTUR: her gün sadece o güne
    kadarki veriden hesaplanır.
    """
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

    # Günlük getiri listesi (len = len(prices)-1); tabloda hizalamak için
    # başa bir "veri yok" değeri koyuyoruz (ilk gün için önceki fiyat yok).
    display_dates = dates[1:]
    n = len(daily_returns)
    if n == 0:
        return None

    mean_return = sum(daily_returns) / n
    variance = sum((r - mean_return) ** 2 for r in daily_returns) / n
    volatility = variance ** 0.5
    sharpe_like = (mean_return / volatility) if volatility > 1e-9 else 0.0
    cumulative_return = (prices[-1] / prices[0] - 1) * 100

    # İleri yönlü (look-ahead içermeyen), gerçek kümülatif getiriye dayalı
    # günlük skor eğrisi: her gün için, dönem başından o güne kadarki
    # kümülatif getiri baz alınıyor.
    running_scores = []
    for i in range(1, len(prices)):
        cum = (prices[i] / prices[0] - 1) * 100
        running_scores.append(round(50 + cum * 5, 1))  # 50 = nötr baz çizgi

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


def fetch_fonoloji_series(fund_code: str, api_key: str) -> pd.DataFrame | None:
    """TEFAS'ta bulunamayan/başarısız olan fonlar için Fonoloji API'sinden
    (https://fonoloji.com/api-docs) NAV geçmişini çeker. TEFAS'la aynı
    işlem mantığına (compute_fund_metrics) girecek şekilde bir DataFrame
    döndürür: date, price, title. Ayrıca resmi risk değeri gibi ekstra
    alanları da fonksiyon dışına döndürmek için (df, extra_info) tuple'ı
    kullanılır.
    """
    headers = {"X-API-Key": api_key}

    try:
        detail_resp = requests.get(
            f"{FONOLOJI_BASE_URL}/funds/{fund_code}", headers=headers, timeout=10
        )
    except requests.RequestException as exc:
        st.warning(f"Fonoloji'ye bağlanılamadı ({fund_code}): {exc}")
        return None, {}

    if detail_resp.status_code == 401:
        st.error("Fonoloji API anahtarı geçersiz.")
        return None, {}
    if detail_resp.status_code == 429:
        st.warning(f"Fonoloji kota sınırına takıldı ({fund_code}), atlanıyor.")
        return None, {}
    if detail_resp.status_code == 404:
        return None, {}
    if detail_resp.status_code != 200:
        st.warning(
            f"Fonoloji'den {fund_code} için beklenmeyen yanıt: {detail_resp.status_code}"
        )
        return None, {}

    detail = detail_resp.json().get("fund", {})
    fund_name = detail.get("name")
    risk_score = detail.get("risk_score")

    try:
        hist_resp = requests.get(
            f"{FONOLOJI_BASE_URL}/funds/{fund_code}/history",
            headers=headers,
            params={"period": "1m"},
            timeout=10,
        )
    except requests.RequestException as exc:
        st.warning(f"Fonoloji geçmiş verisi alınamadı ({fund_code}): {exc}")
        return None, {}

    if hist_resp.status_code != 200:
        st.warning(
            f"Fonoloji geçmiş verisi alınamadı ({fund_code}): {hist_resp.status_code}"
        )
        return None, {}

    points = hist_resp.json().get("points", [])
    if not points:
        return None, {}

    df = pd.DataFrame(points)
    if "date" not in df.columns or "price" not in df.columns:
        return None, {}

    df["date"] = pd.to_datetime(df["date"])
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df = df.dropna(subset=["price"]).sort_values("date")
    df["title"] = fund_name or fund_code
    df["code"] = fund_code

    if len(df) > TARGET_TRADING_DAYS + 1:
        df = df.tail(TARGET_TRADING_DAYS + 1)

    return df.reset_index(drop=True), {"risk_score": risk_score, "fund_name": fund_name}


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
# Excel yükleme ve işleme
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

            with st.spinner("TEFAS'tan gerçek fiyat verisi çekiliyor..."):
                try:
                    universe = fetch_tefas_universe(start_date, today)
                except RuntimeError as exc:
                    st.error(str(exc))
                    st.stop()

            calculated_funds = []
            not_found = []
            missing_valor = []
            fallback_used = []

            for code in requested_codes:
                series = get_fund_series(universe, code)
                metrics = compute_fund_metrics(series)
                data_source = "TEFAS"
                extra_info = {}

                if metrics is None and fonoloji_api_key:
                    fb_series, extra_info = fetch_fonoloji_series(code, fonoloji_api_key)
                    metrics = compute_fund_metrics(fb_series)
                    if metrics is not None:
                        data_source = "Fonoloji (yedek)"
                        fallback_used.append(code)

                if metrics is None:
                    not_found.append(code)
                    continue

                if valor_by_code.get(code) is None:
                    missing_valor.append(code)

                calculated_funds.append(
                    {
                        "code": code,
                        "data_source": data_source,
                        "risk_score": extra_info.get("risk_score", "-"),
                        **metrics,
                    }
                )

            if fallback_used:
                st.info(
                    "TEFAS'ta bulunamadığı için Fonoloji API yedeği kullanıldı: "
                    f"{', '.join(fallback_used)}"
                )
            if not_found:
                extra_hint = "" if fonoloji_api_key else " (Fonoloji API anahtarı girilirse otomatik yedek denenir)"
                st.warning(
                    "Hiçbir kaynakta bulunamadı veya yeterli fiyat verisi yok "
                    f"(analiz dışı bırakıldı){extra_hint}: {', '.join(not_found)}"
                )
            if missing_valor:
                st.info(
                    "Valör bilgisi girilmemiş fonlar için valör cezası "
                    f"uygulanmadı: {', '.join(missing_valor)}"
                )

            if not calculated_funds:
                st.error("Analiz edilebilecek hiçbir fon bulunamadı.")
                st.stop()

            # --- Göreli (z-skor tabanlı) KGDM-3 skoru ---
            mean_returns = [f["mean_return"] for f in calculated_funds]
            sharpes = [f["sharpe_like"] for f in calculated_funds]
            cum_returns = [f["cumulative_return"] for f in calculated_funds]

            multi_fund = len(calculated_funds) > 1
            z_mean = zscore(mean_returns) if multi_fund else [0.0] * len(calculated_funds)
            z_sharpe = zscore(sharpes) if multi_fund else [0.0] * len(calculated_funds)
            z_cum = zscore(cum_returns) if multi_fund else [0.0] * len(calculated_funds)

            if not multi_fund:
                st.info(
                    "Sadece 1 fon yüklendiği için göreli (peer) karşılaştırma "
                    "yapılamıyor; skor yalnızca bu fonun kendi getiri/volatilite "
                    "değerlerine dayanıyor."
                )

            for i, item in enumerate(calculated_funds):
                valor = valor_by_code.get(item["code"])
                valor_penalty = (valor * 0.5) if valor is not None else 0.0

                kgdm_skor = 50 + 15 * z_mean[i] + 20 * z_sharpe[i] + 15 * z_cum[i] - valor_penalty
                kgdm_skor = round(max(0.0, min(100.0, kgdm_skor)), 1)

                if kgdm_skor >= 60:
                    karar, karar_sira = "GÜÇLÜ AL (≥60 Puan)", 1
                elif kgdm_skor >= 40:
                    karar, karar_sira = "ASIL LİSTE (40-59 Puan)", 2
                elif kgdm_skor >= 25:
                    karar, karar_sira = "NÖTR / İZLEME (25-39 Puan)", 3
                else:
                    karar, karar_sira = "ACİL SAT (<25 Puan)", 4

                item.update(
                    {
                        "valor": valor if valor is not None else "-",
                        "kgdm_skor": kgdm_skor,
                        "karar": karar,
                        "karar_sira": karar_sira,
                    }
                )

            calculated_funds.sort(key=lambda x: (x["karar_sira"], -x["kgdm_skor"]))

            # TEFAS ve Fonoloji'den gelen fonların gün sayısı farklı olabilir
            # (örn. tatil günleri, kısmi veri). Excel satırlarının kaymaması
            # için hepsini ortak (en kısa) gün sayısına hizalıyoruz.
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
                "Fon Kodu", "Fon Adı", "Valör", "KGDM-3 Skor (göreli)", "Model Kararı",
                "Ort. Günlük Getiri (%)", "Volatilite (%)", "Sharpe-benzeri Oran",
                "Veri Kaynağı", "Resmi Risk Değeri (1-7)",
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
                    round(item["sharpe_like"], 3),
                    item["data_source"], item["risk_score"],
                ]
                row_data += item["running_scores"]
                row_data += pct_strs
                ws_scores.append(row_data)
                scores_table_data.append(row_data)

            decision_col = 5  # 1-indexed: "Model Kararı"
            n_meta_cols = 10  # Fon Kodu .. Resmi Risk Değeri arasındaki sabit sütun sayısı
            return_cols_start = n_meta_cols + n_days + 1  # % Getiri sütunlarının başladığı yer

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
                    width = max(
                        (len(str(c.value or "")) for c in col), default=8
                    ) + 3
                    sheet.column_dimensions[get_column_letter(col[0].column)].width = max(width, 12)

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success(
                f"✅ {len(calculated_funds)} fon gerçek TEFAS verisiyle analiz edildi "
                f"({n_days} işlem günü, {start_date} - {today} aralığından)."
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
                label="📥 Tam Tabloyu İndir (fonlar_guncel.xlsx)",
                data=output,
                file_name="fonlar_guncel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
