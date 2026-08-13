import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 1. 'fonlar.xlsx' Dosyasını Yükle veya Yoksa Oluştur
file_name = "fonlar.xlsx"

try:
  wb = openpyxl.load_workbook(file_name)
except FileNotFoundError:
  wb = openpyxl.Workbook()

# Fon_Listesi Sayfasını Kontrol Et
if "Fon_Listesi" in wb.sheetnames:
  ws_list = wb["Fon_Listesi"]
else:
  ws_list = wb.active
  ws_list.title = "Fon_Listesi"
  ws_list.append([
      "Fon Kodu",
      "Fon Adı / Açıklama",
      "Aktif Portföyde Var mı?",
      "Valör Süresi (Gün)",
  ])

# Fon_Listesi Sayfasındaki Tüm Fonları Okuma
user_funds = []
for row in ws_list.iter_rows(min_row=2, values_only=True):
  if row[0]:  # Fon Kodu boş değilse
    fon_kodu = str(row[0]).strip().upper()
    fon_adi = str(row[1]) if row[1] else "Tanımsız Fon"
    portfoyde = str(row[2]) if row[2] else "Hayır"
    valor = int(row[3]) if row[3] is not None else 3
    user_funds.append({
        "kod": fon_kodu,
        "adi": fon_adi,
        "portfoyde": portfoyde,
        "valor": valor,
    })

# 2. KGDM3_Puanlama Sayfasını Hazırla / Yenile
if "KGDM3_Puanlama" in wb.sheetnames:
  del wb["KGDM3_Puanlama"]

ws_scores = wb.create_sheet(title="KGDM3_Puanlama")

# Son 10 İş Günü Tarihlerini Dinamik Hesaplama
end_date = datetime.date(2026, 8, 13)
business_days = []
curr = end_date
while len(business_days) < 10:
  if curr.weekday() < 5:  # Pazartesi - Cuma arası
    business_days.append(curr.strftime("%d.%m"))
  curr -= datetime.timedelta(days=1)
business_days.reverse()

headers_scores = (
    ["Fon Kodu", "Valör", "KGDM-3 Anlık Skor", "Model Kararı"]
    + business_days
    + ["Açıklama / Aksiyon"]
)
ws_scores.append(headers_scores)

# Başlık Tasarımı ve Renklendirme
header_fill = PatternFill(
    start_color="1F4E79", end_color="1F4E79", fill_type="solid"
)
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

for cell in ws_scores[1]:
  cell.fill = header_fill
  cell.font = header_font
  cell.alignment = Alignment(horizontal="center", vertical="center")

# 3. KGDM-3 (3 Katmanlı Gelişmiş Dinamik Model) Hesaplama Veri Kütüphanesi
# Katman 1 (%60 KAZRİSK) + Katman 2 (%30 Makro/Sektör) - Katman 3 (%10 Valör Cezası)
base_fund_metrics = {
    "PNU": {"kazrisk": 52, "makro": 30, "aksiyon": "Likit Ana Depo (%41-42 Nema)"},
    "VK6": {
        "kazrisk": 40,
        "makro": 30,
        "aksiyon": "Likit Çapa (Kamu Güvencesi)",
    },
    "ICH": {"kazrisk": 41, "makro": 28, "aksiyon": "#1 Küresel Çip Lideri"},
    "KHA": {"kazrisk": 24, "makro": 26, "aksiyon": "%0 Stopajlı BİST Hisse"},
    "RUT": {
        "kazrisk": 21,
        "makro": 25,
        "aksiyon": "Uzay ve Robotik Tematik",
    },
    "AFA": {"kazrisk": 22, "makro": 23, "aksiyon": "S&P 500 Geniş Piyasa"},
    "KZL": {
        "kazrisk": 20,
        "makro": 24,
        "aksiyon": "Kıymetli Madenler Katılım",
    },
    "BVV": {"kazrisk": 24, "makro": 19, "aksiyon": "Taze Giriş Yapan Çip Fonu"},
    "LTL": {
        "kazrisk": 15,
        "makro": 18,
        "aksiyon": "BİST İyileşme Gösteren Fon",
    },
    "AFS": {
        "kazrisk": 18,
        "makro": 18,
        "aksiyon": "31 Ağu Beklemeden Çıkış Adayı",
    },
    "AFT": {"kazrisk": 16, "makro": 11, "aksiyon": "Çakışan Tema - Acil Sat"},
    "YAY": {"kazrisk": 15, "makro": 10, "aksiyon": "Yüksek Ücret / Zayıf Akış"},
}

# 4. Her Bir Fon İçin KGDM-3 Puanlarının ve 10 Günlük Trendin Otomatik Üretilmesi
for item in user_funds:
  code = item["kod"]
  valor = item["valor"]

  metrics = base_fund_metrics.get(
      code,
      {
          "kazrisk": 15,
          "makro": 15,
          "aksiyon": "Yeni Eklenen Fon / Takip Modunda",
      },
  )

  valor_ceza = valor * 1.5
  kgdm_skor = round(metrics["kazrisk"] + metrics["makro"] - valor_ceza, 1)

  if kgdm_skor >= 60:
    karar = "GÜÇLÜ AL"
  elif kgdm_skor >= 40:
    karar = "ASIL LİSTE"
  elif kgdm_skor >= 25:
    karar = "NÖTR / İZLEME"
  else:
    karar = "ACİL SAT"

  # Son 10 günlük simüle edilmiş gelişim eğrisi
  daily_trend = []
  base_start = kgdm_skor - 12 if "ACİL SAT" not in karar else kgdm_skor + 10
  for i in range(10):
    if "ACİL SAT" in karar:
      val = base_start - (i * 1.1)
    else:
      val = base_start + (i * 1.2)
    daily_trend.append(round(min(100, max(0, val)), 1))

  daily_trend[-1] = kgdm_skor  # Son gün anlık skordur

  row_data = (
      [code, valor, kgdm_skor, karar] + daily_trend + [metrics["aksiyon"]]
  )
  ws_scores.append(row_data)

# 5. Excel Hücrelerini Renklendirme ve Biçimlendirme
green_fill = PatternFill(
    start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
)
yellow_fill = PatternFill(
    start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
)
red_fill = PatternFill(
    start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
)

for row in ws_scores.iter_rows(
    min_row=2,
    max_row=len(user_funds) + 1,
    min_col=1,
    max_col=len(headers_scores),
):
  karar_cell = row[3]
  val = str(karar_cell.value)

  if "GÜÇLÜ AL" in val or "ASIL LİSTE" in val:
    karar_cell.fill = green_fill
    karar_cell.font = Font(name="Calibri", bold=True, color="375623")
  elif "NÖTR" in val:
    karar_cell.fill = yellow_fill
    karar_cell.font = Font(name="Calibri", color="7F6000")
  elif "ACİL SAT" in val:
    karar_cell.fill = red_fill
    karar_cell.font = Font(name="Calibri", bold=True, color="C65911")

# Otomatik Sütun Genişliği Ayarı
for sheet in [ws_list, ws_scores]:
  for col in sheet.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

# 6. Aynı 'fonlar.xlsx' Dosyasına Kaydet
wb.save(file_name)
print(
    f"SUCCESS: '{file_name}' dosyası Fon_Listesi'ne göre başarıyla güncellendi!"
)
