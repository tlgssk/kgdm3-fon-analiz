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


st.set_page_config(
    page_title="KGDM-3 Fon Analiz Otomasyonu",
    page_icon="📊",
    layout="wide",
)

st.title("📊 KGDM-3 Fon Analiz ve Excel Otomasyonu")
st.caption(
    "TEFAS merkezli veri şelalesi kullanarak fonları analiz eder, "
    "KGDM-3 puanı üretir ve güncel Excel dosyası oluşturur."
)


FUND_KINDS = ("YAT", "EMK", "BYF")
LOOKBACK_CALENDAR_DAYS = 35
TARGET_TRADING_DAYS = 10
HTTP_TIMEOUT = 8
APP_VERSION = "3.0.0"

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
            text = text.replace(".", "")
            text = text.replace(",", ".")
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


def format_money(value) -> str:
    number = parse_number(value)

    if number is None:
        return "-"

    return f"{number:,.0f} ₺".replace(",", ".")


def format_integer(value) -> str:
    number = parse_number(value)

    if number is None:
        return "-"

    return f"{int(round(number)):,}".replace(",", ".")


def format_percent(value) -> str:
    number = parse_number(value)

    if number is None:
        return "-"

    if number > 0:
        return f"+%{number:.2f}"

    if number < 0:
        return f"-%{abs(number):.2f}"

    return "%0.00"


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
            max_retry=5,
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

        required_columns = [
            "date",
            "code",
            "price",
        ]

        for column in required_columns:
            if column not in df.columns:
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
            df["investors"] = 0

        if "kind" not in df.columns:
            df["kind"] = ""

        if "title" not in df.columns:
            df["title"] = ""

        df = df.dropna(
            subset=[
                "date",
                "code",
                "price",
            ]
        )

        df = df[df["price"] > 0]

        df = (
            df.sort_values(["code", "date"])
            .drop_duplicates(
                subset=["code", "date"],
                keep="last",
            )
            .reset_index(drop=True)
        )

        return df

    except Exception:
        return pd.DataFrame()


def get_fund_series(
    universe: pd.DataFrame,
    fund_code: str,
) -> Optional[pd.DataFrame]:
    if universe is None or universe.empty:
        return None

    code = normalize_fund_code(fund_code)

    if not code:
        return None

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

    if len(rows) > TARGET_TRADING_DAYS + 1:
        rows = rows.tail(
            TARGET_TRADING_DAYS + 1
        )

    return rows.reset_index(drop=True)


def get_latest_fund_info(
    universe: pd.DataFrame,
    fund_code: str,
) -> dict:
    result = {
        "aum": 0.0,
        "investors": 0,
    }

    if universe is None or universe.empty:
        return result

    code = normalize_fund_code(fund_code)

    rows = universe[
        universe["code"]
        .astype(str)
        .str.upper()
        .eq(code)
    ].copy()

    if rows.empty:
        return result

    rows = rows.sort_values("date")
    latest = rows.iloc[-1]

    aum = parse_number(latest.get("aum"))
    investors = parse_number(latest.get("investors"))

    result["aum"] = aum if aum is not None else 0.0

    result["investors"] = int(
        round(
            investors
            if investors is not None
            else 0
        )
    )

    return result


def fetch_isyatirim_series(
    fund_code: str,
) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)

    if not code:
        return None

    end = dt.datetime.now()
    start = end - dt.timedelta(days=45)

    url = (
        "https://www.isyatirim.com.tr/"
        "_layouts/15/IsYatirim.Website/Common/Data.aspx/"
        "YatirimFonGecmisGetiri"
    )

    params = {
        "fonKod": code,
        "baslangic": start.strftime("%d-%m-%Y"),
        "bitis": end.strftime("%d-%m-%Y"),
    }

    headers = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/128 Safari/537.36"
        ),
        "Accept": (
            "application/json,"
            "text/plain,"
            "*/*"
        ),
    }

    try:
        response = requests.get(
            url,
            params=params,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        response.raise_for_status()

        payload = response.json()
        values = payload.get("value")

        if not isinstance(values, list) or not values:
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

        df["price"] = df["Fiyat"].apply(parse_number)

        df = df.dropna(
            subset=[
                "date",
                "price",
            ]
        )

        df = df[df["price"] > 0]

        if len(df) < 2:
            return None

        df = (
            df.sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
            .tail(TARGET_TRADING_DAYS + 1)
            .reset_index(drop=True)
        )

        return df[["date", "price"]]

    except (
        requests.RequestException,
        ValueError,
        TypeError,
    ):
        return None


def fetch_fintables_series(
    fund_code: str,
) -> Optional[pd.DataFrame]:
    code = normalize_fund_code(fund_code)

    if not code:
        return None

    url = (
        "https://fintables.com/fonlar/"
        f"{code.lower()}"
    )

    headers = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/128 Safari/537.36"
        )
    }

    try:
        response = requests.get(
            url,
            headers=headers,
            timeout=HTTP_TIMEOUT,
        )

        if response.status_code != 200:
            return None

        html = response.text

        pattern = re.compile(
            r'"date"\s*:\s*"'
            r'(\d{4}-\d{2}-\d{2})'
            r'"[^{}]{0,500}?'
            r'"price"\s*:\s*'
            r'([0-9]+(?:\.[0-9]+)?)',
            re.IGNORECASE,
        )

        matches = pattern.findall(html)

        if not matches:
            return None

        rows = []

        for date_text, price_text in matches:
            date_value = pd.to_datetime(
                date_text,
                errors="coerce",
            )

            price_value = parse_number(
                price_text
            )

            if (
                pd.notna(date_value)
                and price_value is not None
                and price_value > 0
            ):
                rows.append(
                    {
                        "date": date_value,
                        "price": price_value,
                    }
                )

        if not rows:
            return None

        df = pd.DataFrame(rows)

        df = (
            df.sort_values("date")
            .drop_duplicates(
                subset=["date"],
                keep="last",
            )
        )

        if len(df) < 2:
            return None

        return (
            df.tail(TARGET_TRADING_DAYS + 1)
            .reset_index(drop=True)
        )

    except (
        requests.RequestException,
        ValueError,
        TypeError,
        re.error,
    ):
        return None


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
        subset=[
            "date",
            "price",
        ]
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

    dates = (
        df["date"]
        .dt.strftime("%d.%m")
        .tolist()
    )

    daily_returns = []

    for previous, current in zip(
        prices[:-1],
        prices[1:],
    ):
        if previous <= 0:
            daily_returns.append(0.0)
        else:
            daily_returns.append(
                (current / previous - 1) * 100
            )

    if not daily_returns:
        return None

    mean_return = (
        sum(daily_returns)
        / len(daily_returns)
    )

    variance = (
        sum(
            (value - mean_return) ** 2
            for value in daily_returns
        )
        / len(daily_returns)
    )

    volatility = variance ** 0.5

    if volatility > 1e-12:
        sharpe_like = (
            mean_return / volatility
        )
    else:
        sharpe_like = 0.0

    cumulative_return = (
        (prices[-1] / prices[0] - 1)
        * 100
    )

    running_scores = []

    for price in prices[1:]:
        raw_score = (
            50
            + (
                (price / prices[0] - 1)
                * 100
            )
            * 5
        )

        score = int(
            round(
                max(
                    0.0,
                    min(
                        100.0,
                        raw_score,
                    ),
                )
            )
        )

        running_scores.append(score)

    return {
        "dates": dates[1:],
        "daily_returns": daily_returns,
        "running_scores": running_scores,
        "mean_return": mean_return,
        "volatility": volatility,
        "sharpe_like": sharpe_like,
        "cumulative_return": cumulative_return,
        "n_days": len(daily_returns),
    }


def zscore(
    values: list[float],
) -> list[float]:
    if not values:
        return []

    clean_values = [
        float(value)
        for value in values
        if value is not None
        and pd.notna(value)
    ]

    if not clean_values:
        return [0.0 for _ in values]

    mean_value = (
        sum(clean_values)
        / len(clean_values)
    )

    variance = (
        sum(
            (value - mean_value) ** 2
            for value in clean_values
        )
        / len(clean_values)
    )

    std = variance ** 0.5

    if std < 1e-12:
        return [0.0 for _ in values]

    return [
        (
            float(value) - mean_value
        ) / std
        for value in values
    ]


def calculate_kgdm_score(
    mean_z: float,
    sharpe_z: float,
    cumulative_z: float,
    valor: Optional[int],
) -> int:
    valor_penalty = (
        valor * 0.5
        if valor is not None
        else 0.0
    )

    raw_score = (
        50
        + 15 * mean_z
        + 20 * sharpe_z
        + 15 * cumulative_z
        - valor_penalty
    )

    return int(
        round(
            max(
                0.0,
                min(
                    100.0,
                    raw_score,
                ),
            )
        )
    )


def get_decision(
    score: int,
) -> tuple[str, int]:
    if score >= 60:
        return (
            "GÜÇLÜ AL (≥60 Puan)",
            1,
        )

    if score >= 40:
        return (
            "ASIL LİSTE (40-59 Puan)",
            2,
        )

    if score >= 25:
        return (
            "NÖTR / İZLEME (25-39 Puan)",
            3,
        )

    return (
        "ACİL SAT (<25 Puan)",
        4,
    )


def style_excel_sheet(ws):
    thin_gray = Side(
        style="thin",
        color="D9E1F2",
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
            )
            cell.border = Border(
                bottom=thin_gray
            )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False


def auto_fit_columns(
    ws,
    min_width: int = 10,
    max_width: int = 45,
):
    for column_cells in ws.columns:
        column_index = column_cells[0].column
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
            get_column_letter(column_index)
        ].width = width


def create_excel_output(
    wb,
    ws_list,
    calculated_funds,
) -> io.BytesIO:

    if "KGDM3_Puanlama" in wb.sheetnames:
        del wb["KGDM3_Puanlama"]

    ws_scores = wb.create_sheet(
        title="KGDM3_Puanlama"
    )

    n_days = min(
        item["n_days"]
        for item in calculated_funds
    )

    for item in calculated_funds:
        item["dates"] = item["dates"][-n_days:]
        item["daily_returns"] = item[
            "daily_returns"
        ][-n_days:]
        item["running_scores"] = item[
            "running_scores"
        ][-n_days:]

    day_labels = calculated_funds[0]["dates"]

    headers = [
        "Fon Kodu",
        "Valör (Excel)",
        "KGDM-3 Skor",
        "Model Kararı",
        "Ort. Getiri (%)",
        "Volatilite (%)",
        "Sharpe",
        "Kümülatif Getiri (%)",
        "Fon Büyüklüğü (AUM ₺)",
        "Yatırımcı Sayısı",
        "Fiyat Kaynağı",
    ]

    for day in day_labels:
        headers.append(
            f"{day} Skor"
        )

    for day in day_labels:
        headers.append(
            f"{day} % Getiri"
        )

    ws_scores.append(headers)

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

    ws_scores.row_dimensions[1].height = 34

    for item in calculated_funds:
        row_data = [
            item["code"],
            (
                item["valor"]
                if item["valor"] is not None
                else None
            ),
            item["kgdm_skor"],
            item["karar"],
            round(item["mean_return"], 4),
            round(item["volatility"], 4),
            round(item["sharpe_like"], 4),
            round(item["cumulative_return"], 4),
            (
                round(item["aum"], 2)
                if item["aum"]
                else None
            ),
            (
                int(item["investors"])
                if item["investors"]
                else None
            ),
            item["data_source"],
        ]

        row_data.extend(
            item["running_scores"]
        )

        row_data.extend(
            [
                format_percent(value)
                for value in item["daily_returns"]
            ]
        )

        ws_scores.append(row_data)

    COL_VALOR = 2
    COL_SCORE = 3
    COL_DECISION = 4
    COL_MEAN = 5
    COL_VOL = 6
    COL_SHARPE = 7
    COL_CUM = 8
    COL_AUM = 9
    COL_INVESTORS = 10

    SCORE_START = 12
    RETURN_START = SCORE_START + n_days

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
        ws_scores.cell(
            row=row_number,
            column=COL_AUM,
        ).number_format = '#,##0.00 "₺"'

        ws_scores.cell(
            row=row_number,
            column=COL_INVESTORS,
        ).number_format = '#,##0'

        ws_scores.cell(
            row=row_number,
            column=COL_VALOR,
        ).number_format = '0'

        for col in [
            COL_MEAN,
            COL_VOL,
            COL_SHARPE,
            COL_CUM,
        ]:
            ws_scores.cell(
                row=row_number,
                column=col,
            ).number_format = '0.0000'

        ws_scores.cell(
            row=row_number,
            column=COL_SCORE,
        ).number_format = '0'

        decision_cell = ws_scores.cell(
            row=row_number,
            column=COL_DECISION,
        )

        decision_text = str(
            decision_cell.value
        )

        if (
            "GÜÇLÜ AL" in decision_text
            or "ASIL LİSTE" in decision_text
        ):
            decision_cell.font = green_font

        elif "NÖTR" in decision_text:
            decision_cell.font = yellow_font

        elif "ACİL SAT" in decision_text:
            decision_cell.font = red_font

        for col in range(
            RETURN_START,
            RETURN_START + n_days,
        ):
            cell = ws_scores.cell(
                row=row_number,
                column=col,
            )

            value = str(cell.value)

            if value.startswith("+"):
                cell.font = green_font
            elif value.startswith("-"):
                cell.font = red_font

    score_range = (
        f"C2:C{ws_scores.max_row}"
    )

    ws_scores.conditional_formatting.add(
        score_range,
        CellIsRule(
            operator="greaterThanOrEqual",
            formula=["60"],
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
            formula=["40", "59"],
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
            formula=["25", "39"],
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
            formula=["25"],
            fill=PatternFill(
                start_color=COLOR_LIGHT_RED,
                fill_type="solid",
            ),
        ),
    )

    style_excel_sheet(ws_scores)
    auto_fit_columns(ws_scores)

    for col in range(
        SCORE_START,
        ws_scores.max_column + 1,
    ):
        ws_scores.column_dimensions[
            get_column_letter(col)
        ].width = 13

    style_excel_sheet(ws_list)
    auto_fit_columns(ws_list)

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ---------------------------------------------------------------------------
# ANA UYGULAMA
# ---------------------------------------------------------------------------

uploaded_file = st.file_uploader(
    "Excel Dosyanızı Yükleyin (fonlar.xlsx):",
    type=["xlsx"],
)

if uploaded_file is None:
    st.info("Başlamak için Fon_Listesi sayfasını içeren Excel dosyanızı yükleyin.")
    st.markdown(
        "### Excel formatı\n\n"
        "`Fon_Listesi` sayfasında:\n\n"
        "| A sütunu | D sütunu |\n"
        "|---|---|\n"
        "| Fon Kodu | Valör |\n\n"
        "Örnek:\n\n"
        "```text\n"
        "Fon Kodu    ...    ...    Valör\n"
        "AFT                      1\n"
        "MAC                      2\n"
        "TCD                      3\n"
        "```"
    )
    st.stop()


try:
    wb = openpyxl.load_workbook(uploaded_file)
except Exception as exc:
    st.error(f"Excel dosyası okunamadı: {exc}")
    st.stop()

if "Fon_Listesi" not in wb.sheetnames:
    st.error("Yüklenen dosyada 'Fon_Listesi' sayfası bulunamadı!")
    st.stop()

ws_list = wb["Fon_Listesi"]

requested_codes = []
excel_valor_dict = {}

for row in ws_list.iter_rows(min_row=2, values_only=False):
    if not row:
        continue

    code_cell = row[0]
    valor_cell = row[3] if len(row) > 3 else None

    code = normalize_fund_code(code_cell.value)

    if not code:
        continue

    requested_codes.append(code)

    valor = None
    if (
        valor_cell is not None
        and valor_cell.value is not None
        and str(valor_cell.value).strip()
    ):
        parsed_valor = parse_number(valor_cell.value)
        if parsed_valor is not None:
            valor = int(round(
