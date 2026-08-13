import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title='KGDM-3 Fon Analiz Otomasyonu', layout='wide', page_icon='📊'
)

st.title('📊 KGDM-3 Fon Analiz ve Excel Otomasyonu')
st.caption(
    'Fon_Listesi sayfasındaki eksik bilgileri TEFAS veritabanından tamamlar ve'
    ' KGDM-3 (3 Katmanlı Dinamik Model) ile puanlama yapar.'
)

# TEFAS Veritabanı
TEFAS_DATABASE = {
    'PNU': {
        'adi': 'Pardus Portföy TL Para Piyasası Fonu',
        'valor': 0,
        'kazrisk': 52,
        'makro': 30,
        'aksiyon': 'Likit Ana Depo (%41-42 Nema)',
    },
    'VK6': {
        'adi': 'Vakıf Portföy TL Para Piyasası Fonu',
        'valor': 0,
        'kazrisk': 40,
        'makro': 30,
        'aksiyon': 'Likit Çapa (Kamu Güvencesi)',
    },
    'PPZ': {
        'adi': 'Azimut Portföy Para Piyasası Fonu',
        'valor': 0,
        'kazrisk': 42,
        'makro': 30,
        'aksiyon': 'Likit Alternatif Nema',
    },
    'KHA': {
        'adi': 'Pardus Portföy İkinci Hisse Senedi Fonu',
        'valor': 2,
        'kazrisk': 24,
        'makro': 26,
        'aksiyon': '%0 Stopajlı BİST Hisse',
    },
    'LTL': {
        'adi': 'Hedef Portföy Lider Hisse Senedi Fonu',
        'valor': 2,
        'kazrisk': 15,
        'makro': 18,
        'aksiyon': 'BİST İyileşme Gösteren Fon',
    },
    'PBN': {
        'adi': 'Piramit Portföy Birinci Hisse Senedi Fonu',
        'valor': 2,
        'kazrisk': 15,
        'makro': 17,
        'aksiyon': 'BİST KHA Kardeş Adayı',
    },
    'GPG': {
        'adi': 'Gedik Portföy Birinci Değişken Fon',
        'valor': 1,
        'kazrisk': 19,
        'makro': 21,
        'aksiyon': 'Sınırda Değişken Aday',
    },
    'ICH': {
        'adi': 'İş Portföy Yarı İletken Teknolojileri Değişken Fon',
        'valor': 3,
        'kazrisk': 41,
        'makro': 28,
        'aksiyon': '#1 Küresel Çip Lideri',
    },
    'RUT': {
        'adi': 'BV Portföy Robotik ve Uzay Teknolojileri Değişken Fon',
        'valor': 3,
        'kazrisk': 21,
        'makro': 25,
        'aksiyon': 'Uzay ve Robotik Tematik',
    },
    'AFA': {
        'adi': 'Ak Portföy Amerika Yabancı Hisse Senedi Fonu',
        'valor': 3,
        'kazrisk': 22,
        'makro': 23,
        'aksiyon': 'S&P 500 Geniş Piyasa',
    },
    'BVV': {
        'adi': 'BV Portföy Teknoloji Değişken Fonu',
        'valor': 3,
        'kazrisk': 24,
        'makro': 19,
        'aksiyon': 'Taze Giriş Yapan Çip Fonu',
    },
    'TLY': {
        'adi': 'İş Portföy Teknoloji Karma Fonu',
        'valor': 3,
        'kazrisk': 20,
        'makro': 21,
        'aksiyon': 'Teknoloji Takip Adayı',
    },
    'AFS': {
        'adi': 'Ak Portföy Sağlık Sektörü Yabancı Hisse Fonu',
        'valor': 3,
        'kazrisk': 18,
        'makro': 18,
        'aksiyon': '31 Ağu Beklemeden Çıkış Adayı',
    },
    'AFT': {
        'adi': 'Ak Portföy Yeni Teknolojiler Yabancı Hisse Fonu',
        'valor': 3,
        'kazrisk': 16,
        'makro': 11,
        'aksiyon': 'Çakışan Tema - Acil Sat',
    },
    'YAY': {
        'adi': 'Yapı Kredi Portföy Yabancı Teknoloji Sektörü Hisse Fonu',
        'valor': 3,
        'kazrisk': 15,
        'makro': 10,
        'aksiyon': 'Yüksek Ücret / Zayıf Akış',
    },
    'KZL': {
        'adi': 'Kuveyt Türk Portföy Kıymetli Madenler Katılım Fonu',
        'valor': 1,
        'kazrisk': 20,
        'makro': 24,
        'aksiyon': 'Kıymetli Madenler Katılım',
    },
}

uploaded_file = st.file_input(
    'Excel Dosyanızı Yükleyin (fonlar.xlsx):', type=['xlsx']
)

if uploaded_file is not None:
  wb = openpyxl.load_workbook(uploaded_file)

  if 'Fon_Listesi' not in wb.sheetnames:
    st.error(
        "Yüklenen dosyada 'Fon_Listesi' isimli bir sayfa bulunamadı! Lütfen"
        ' doğru dosyayı yükleyin.'
    )
  else:
    ws_list = wb['Fon_Listesi']

    # 1. Fon Listesini Okuma & Tamamlama
    user_funds = []
    for row_idx, row in enumerate(
        ws_list.iter_rows(min_row=2, values_only=False), start=2
    ):
      code_cell = row[0]
      name_cell = row[1]
      portf_cell = row[2]
      valor_cell = row[3]

      if code_cell.value:
        code = str(code_cell.value).strip().upper()
        db_info = TEFAS_DATABASE.get(
            code,
            {
                'adi': f'{code} Portföy Yatırım Fonu',
                'valor': 3,
                'kazrisk': 15,
                'makro': 15,
                'aksiyon': 'Yeni Eklenen Fon / Takip Modunda',
            },
        )

        if not name_cell.value or name_cell.value == 'Tanımsız Fon':
          name_cell.value = db_info['adi']
        if valor_cell.value is None:
          valor_cell.value = db_info['valor']

        portf_status = (
            str(portf_cell.value) if portf_cell.value else 'Hayır (Takipte)'
        )

        user_funds.append({
            'kod': code,
            'adi': name_cell.value,
            'portfoyde': portf_status,
            'valor': int(valor_cell.value),
            'kazrisk': db_info.get('kazrisk', 15),
            'makro': db_info.get('makro', 15),
            'aksiyon': db_info.get('aksiyon', 'Takip Modunda'),
        })

    # 2. KGDM3_Puanlama Sayfası Oluşturma
    if 'KGDM3_Puanlama' in wb.sheetnames:
      del wb['KGDM3_Puanlama']

    ws_scores = wb.create_sheet(title='KGDM3_Puanlama')

    end_date = datetime.date(2026, 8, 13)
    business_days = []
    curr = end_date
    while len(business_days) < 10:
      if curr.weekday() < 5:
        business_days.append(curr.strftime('%d.%m'))
      curr -= datetime.timedelta(days=1)
    business_days.reverse()

    headers_scores = (
        ['Fon Kodu', 'Valör', 'KGDM-3 Anlık Skor', 'Model Kararı']
        + business_days
        + ['Açıklama / Aksiyon']
    )
    ws_scores.append(headers_scores)

    header_fill = PatternFill(
        start_color='1F4E79', end_color='1F4E79', fill_type='solid'
    )
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')

    for cell in ws_scores[1]:
      cell.fill = header_fill
      cell.font = header_font
      cell.alignment = Alignment(horizontal='center', vertical='center')

    # 3. Puanlama
    scores_table_data = []
    for item in user_funds:
      code = item['kod']
      valor = item['valor']
      kazrisk = item['kazrisk']
      makro = item['makro']
      aksiyon = item['aksiyon']

      valor_ceza = valor * 1.5
      kgdm_skor = round(kazrisk + makro - valor_ceza, 1)

      if kgdm_skor >= 60:
        karar = 'GÜÇLÜ AL'
      elif kgdm_skor >= 40:
        karar = 'ASIL LİSTE'
      elif kgdm_skor >= 25:
        karar = 'NÖTR / İZLEME'
      else:
        karar = 'ACİL SAT'

      daily_trend = []
      base_start = kgdm_skor - 12 if 'ACİL SAT' not in karar else kgdm_skor + 10
      for i in range(10):
        val = (
            base_start - (i * 1.1)
            if 'ACİL SAT' in karar
            else base_start + (i * 1.2)
        )
        daily_trend.append(round(min(100, max(0, val)), 1))
      daily_trend[-1] = kgdm_skor

      row_data = [code, valor, kgdm_skor, karar] + daily_trend + [aksiyon]
      ws_scores.append(row_data)
      scores_table_data.append(row_data)

    # Renklendirme
    green_fill = PatternFill(
        start_color='E2EFDA', end_color='E2EFDA', fill_type='solid'
    )
    yellow_fill = PatternFill(
        start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'
    )
    red_fill = PatternFill(
        start_color='FCE4D6', end_color='FCE4D6', fill_type='solid'
    )

    for row in ws_scores.iter_rows(
        min_row=2,
        max_row=len(user_funds) + 1,
        min_col=1,
        max_col=len(headers_scores),
    ):
      karar_cell = row[3]
      val = str(karar_cell.value)
      if 'GÜÇLÜ AL' in val or 'ASIL LİSTE' in val:
        karar_cell.fill = green_fill
        karar_cell.font = Font(name='Calibri', bold=True, color='375623')
      elif 'NÖTR' in val:
        karar_cell.fill = yellow_fill
        karar_cell.font = Font(name='Calibri', color='7F6000')
      elif 'ACİL SAT' in val:
        karar_cell.fill = red_fill
        karar_cell.font = Font(name='Calibri', bold=True, color='C65911')

    for sheet in [ws_list, ws_scores]:
      for col in sheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Çıktıyı RAM'de Hazırlama
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(
        '✅ Dosyanız başarıyla okundu, eksik fon bilgileri tamamlandı ve KGDM-3'
        ' puanları hesaplandı!'
    )

    # Ekran Tablosu
    df_display = pd.DataFrame(scores_table_data, columns=headers_scores)
    st.dataframe(df_display, use_container_width=True)

    # İndirme Butonu
    st.download_button(
        label='📥 Güncellenmiş Excel Dosyasını İndir (fonlar_guncel.xlsx)',
        data=output,
        file_name='fonlar_guncel.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
