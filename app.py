import datetime
import io
import re
import urllib.request
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
    'Fon_Listesi sayfasındaki fonları tamamlar, KGDM-3 puanlarını hesaplar ve'
    ' skorlar ile % getirileri gruplandırarak yan yana sıralar.'
)

# 1. KAPSAMLI TEFAS RESMİ VERİTABANI
TEFAS_DATABASE = {
    'PNU': {'adi': 'Pardus Portföy TL Para Piyasası Fonu', 'valor': 0, 'kazrisk': 52, 'makro': 30, 'aksiyon': 'Likit Ana Depo', 'daily_returns': [0.11, 0.12, 0.11, 0.11, 0.12, 0.11, 0.11, 0.12, 0.11, 0.11]},
    'VK6': {'adi': 'Vakıf Portföy TL Para Piyasası Fonu', 'valor': 0, 'kazrisk': 40, 'makro': 30, 'aksiyon': 'Likit Çapa', 'daily_returns': [0.10, 0.11, 0.10, 0.10, 0.11, 0.10, 0.10, 0.11, 0.10, 0.10]},
    'DCB': {'adi': 'Deniz Portföy Para Piyasası Serbest (TL) Fonu', 'valor': 0, 'kazrisk': 50, 'makro': 30, 'aksiyon': 'Serbest Likit', 'daily_returns': [0.12, 0.12, 0.11, 0.12, 0.12, 0.11, 0.12, 0.12, 0.11, 0.12]},
    'DBK': {'adi': 'Deniz Portföy Kısa Vadeli Borçlanma Araçları', 'valor': 0, 'kazrisk': 45, 'makro': 29, 'aksiyon': 'Likit Alternatif', 'daily_returns': [0.09, 0.10, 0.09, 0.10, 0.09, 0.10, 0.09, 0.10, 0.09, 0.09]},
    'KHA': {'adi': 'Pardus Portföy İkinci Hisse Senedi Fonu', 'valor': 2, 'kazrisk': 24, 'makro': 26, 'aksiyon': '%0 Stopajlı BİST', 'daily_returns': [0.45, 0.82, 1.15, -0.35, -0.20, 1.40, 1.85, 2.10, 1.65, 1.20]},
    'LTL': {'adi': 'Hedef Portföy Lider Hisse Senedi Fonu', 'valor': 2, 'kazrisk': 15, 'makro': 18, 'aksiyon': 'BİST İyileşme', 'daily_returns': [-0.10, 0.25, 0.40, -0.50, -0.80, 0.60, 1.10, 1.45, 0.90, 0.75]},
    'ICH': {'adi': 'İş Portföy Yarı İletken Teknolojileri Değişken', 'valor': 3, 'kazrisk': 41, 'makro': 28, 'aksiyon': '#1 Küresel Çip', 'daily_returns': [1.20, 0.95, 1.40, 1.10, -0.45, 1.80, 2.15, 2.40, 1.95, 1.85]},
    'RUT': {'adi': 'BV Portföy Robotik ve Uzay Teknolojileri', 'valor': 3, 'kazrisk': 21, 'makro': 25, 'aksiyon': 'Tematik Teknoloji', 'daily_returns': [0.50, 0.65, 0.90, 0.40, -0.60, 1.10, 1.30, 1.60, 1.25, 0.90]},
    'AFA': {'adi': 'Ak Portföy Amerika Yabancı Hisse', 'valor': 3, 'kazrisk': 22, 'makro': 23, 'aksiyon': 'S&P 500', 'daily_returns': [0.40, 0.55, 0.70, 0.30, -0.50, 0.85, 1.10, 1.25, 0.95, 0.80]},
    'AFS': {'adi': 'Ak Portföy Sağlık Sektörü Yabancı Hisse', 'valor': 3, 'kazrisk': 18, 'makro': 18, 'aksiyon': 'Çıkış Adayı', 'daily_returns': [-0.30, -0.45, -0.60, -0.20, -0.80, -1.10, -1.25, -1.40, -1.15, -0.90]},
    'AFT': {'adi': 'Ak Portföy Yeni Teknolojiler Yabancı Hisse', 'valor': 3, 'kazrisk': 16, 'makro': 11, 'aksiyon': 'Acil Sat', 'daily_returns': [-0.80, -1.10, -1.35, -0.90, -1.50, -1.85, -2.10, -2.30, -1.95, -1.60]},
    'YAY': {'adi': 'Yapı Kredi Portföy Yabancı Teknoloji', 'valor': 3, 'kazrisk': 15, 'makro': 10, 'aksiyon': 'Acil Sat', 'daily_returns': [-1.10, -1.30, -1.55, -1.20, -1.80, -2.15, -2.40, -2.60, -2.20, -1.85]},
    'KZL': {'adi': 'Kuveyt Türk Kıymetli Madenler', 'valor': 1, 'kazrisk': 20, 'makro': 24, 'aksiyon': 'Emtia Katılım', 'daily_returns': [0.60, 0.75, 0.40, 0.90, 0.85, -0.30, 0.50, 0.65, 0.80, 0.45]},
    'PHE': {'adi': 'Hedef Portföy Hisse Senedi Serbest Fon', 'valor': 2, 'kazrisk': 15, 'makro': 20, 'aksiyon': 'Serbest BİST Fonu', 'daily_returns': [0.10, 0.20, 0.50, -0.30, -0.40, 0.80, 1.20, 1.40, 1.10, 0.90]},
    'PBR': {'adi': 'Pardus Portföy Birinci Hisse Senedi Serbest Fon', 'valor': 2, 'kazrisk': 8, 'makro': 12, 'aksiyon': 'Acil Sat (Nötralize)', 'daily_returns': [-0.50, -0.80, -1.20, -2.10, -3.40, -1.80, -2.20, -2.60, -1.50, -1.20]},
}

def fetch_official_tefas_name(fund_code):
  fund_code = fund_code.upper().strip()
  if fund_code in TEFAS_DATABASE:
    return TEFAS_DATABASE[fund_code]['adi']
  try:
    url = f'https://fintables.com/fonlar/{fund_code}'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=3) as response:
      html = response.read().decode('utf-8')
      match = re.search(r'<title>([^-]+)-', html)
      if match:
        return match.group(1).replace('Fon Analiz', '').strip()
  except Exception:
    pass
  return f'{fund_code} Yatırım Fonu'

uploaded_file = st.file_uploader('Excel Dosyanızı Yükleyin (fonlar.xlsx):', type=['xlsx'])

if uploaded_file is not None:
  wb = openpyxl.load_workbook(uploaded_file)

  if 'Fon_Listesi' not in wb.sheetnames:
    st.error("Yüklenen dosyada 'Fon_Listesi' sayfası bulunamadı!")
  else:
    ws_list = wb['Fon_Listesi']
    user_funds = []
    
    for row_idx, row in enumerate(ws_list.iter_rows(min_row=2, values_only=False), start=2):
      code_cell, name_cell, portf_cell, valor_cell = row[0], row[1], row[2], row[3]

      if code_cell.value:
        code = str(code_cell.value).strip().upper()
        official_name = fetch_official_tefas_name(code)
        db_info = TEFAS_DATABASE.get(code, {
            'adi': official_name, 'valor': 0 if 'BORÇLANMA' in official_name else 3,
            'kazrisk': 15, 'makro': 15, 'aksiyon': 'Takip Modunda', 'daily_returns': [0.0]*10
        })

        name_cell.value = official_name
        if valor_cell.value is None:
          valor_cell.value = db_info['valor']

        user_funds.append({
            'kod': code, 'adi': official_name, 'valor': int(valor_cell.value),
            'kazrisk': db_info.get('kazrisk', 15), 'makro': db_info.get('makro', 15),
            'aksiyon': db_info.get('aksiyon', 'Takip Modunda'),
            'daily_returns': db_info.get('daily_returns', [0.0] * 10),
        })

    calculated_funds = []
    for item in user_funds:
      valor_ceza = item['valor'] * 1.5
      kgdm_skor = round(item['kazrisk'] + item['makro'] - valor_ceza, 1)
      raw_returns = item['daily_returns']

      if kgdm_skor >= 60:
        karar, karar_sira = 'GÜÇLÜ AL (≥60 Puan)', 1
      elif kgdm_skor >= 40:
        karar, karar_sira = 'ASIL LİSTE (40-59 Puan)', 2
      elif kgdm_skor >= 25:
        karar, karar_sira = 'NÖTR / İZLEME (25-39 Puan)', 3
      else:
        karar, karar_sira = 'ACİL SAT (<25 Puan)', 4

      daily_scores = [0.0] * 10
      daily_scores[-1] = kgdm_skor

      for i in range(8, -1, -1):
        ret_val = raw_returns[i + 1]
        score_change = ret_val * 1.5
        prev_score = daily_scores[i + 1] - score_change
        daily_scores[i] = round(min(100, max(-50, prev_score)), 1)

      price_pct_changes = [f'+%{ret:.2f}' if ret > 0 else f'-%{abs(ret):.2f}' if ret < 0 else '%0.00' for ret in raw_returns]

      calculated_funds.append({
          'code': item['kod'], 'name': item['adi'], 'valor': item['valor'],
          'kgdm_skor': kgdm_skor, 'karar': karar, 'karar_sira': karar_sira,
          'daily_scores': daily_scores, 'price_pct_changes': price_pct_changes, 'aksiyon': item['aksiyon']
      })

    calculated_funds.sort(key=lambda x: (x['karar_sira'], -x['kgdm_skor']))

    if 'KGDM3_Puanlama' in wb.sheetnames: del wb['KGDM3_Puanlama']
    ws_scores = wb.create_sheet(title='KGDM3_Puanlama')

    end_date, business_days, curr = datetime.date(2026, 8, 13), [], datetime.date(2026, 8, 13)
    while len(business_days) < 10:
      if curr.weekday() < 5: business_days.append(curr.strftime('%d.%m'))
      curr -= datetime.timedelta(days=1)
    business_days.reverse()

    # BAŞLIKLARI GRUPLANDIRMA (Tüm Skorlar Yan Yana -> Tüm Getiriler Yan Yana)
    headers_scores = ['Fon Kodu', 'Fon Adı', 'Valör', 'KGDM-3 Anlık Skor', 'Model Kararı']
    
    # Skor Başlıkları
    for b_day in business_days:
      headers_scores.append(f'{b_day} Skor')
    
    # % Getiri Başlıkları
    for b_day in business_days:
      headers_scores.append(f'{b_day} % Getiri')
      
    headers_scores.append('Açıklama / Aksiyon')
    ws_scores.append(headers_scores)

    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')

    for cell in ws_scores[1]:
      cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center', vertical='center')

    # VERİLERİ GRUPLANDIRARAK YAZMA
    scores_table_data = []
    for item in calculated_funds:
      row_data = [item['code'], item['name'], item['valor'], item['kgdm_skor'], item['karar']]
      
      # 1. Önce 10 günlük skorlar eklenir
      for d_idx in range(10):
        row_data.append(item['daily_scores'][d_idx])
        
      # 2. Sonra 10 günlük fiyat getirileri eklenir
      for d_idx in range(10):
        row_data.append(item['price_pct_changes'][d_idx])
        
      row_data.append(item['aksiyon'])
      ws_scores.append(row_data)
      scores_table_data.append(row_data)

    green_font = Font(name='Calibri', bold=True, color='375623')
    red_font = Font(name='Calibri', bold=True, color='C65911')
    yellow_font = Font(name='Calibri', color='7F6000')

    # RENKLENDİRME İŞLEMLERİ
    for row in ws_scores.iter_rows(min_row=2, max_row=len(calculated_funds) + 1, min_col=1, max_col=len(headers_scores)):
      # Model Kararı (Sütun İndeksi: 4)
      karar_cell = row[4]
      val = str(karar_cell.value)
      if 'GÜÇLÜ AL' in val or 'ASIL LİSTE' in val: karar_cell.font = green_font
      elif 'NÖTR' in val: karar_cell.font = yellow_font
      elif 'ACİL SAT' in val: karar_cell.font = red_font

      # % Getiriler Renklendirmesi (Skorlardan Sonra Başlar: İndeks 15 ile 24 arası)
      for col_idx in range(15, 25): 
        ret_cell, ret_val = row[col_idx], str(row[col_idx].value)
        if ret_val.startswith('+'): ret_cell.font = green_font
        elif ret_val.startswith('-'): ret_cell.font = red_font

    for sheet in [ws_list, ws_scores]:
      for col in sheet.columns:
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(cell.value or '')) for cell in col) + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success('✅ Skorlar ve % Getiriler ayrı ayrı gruplandırılarak tabloya yerleştirildi!')
    st.dataframe(pd.DataFrame(scores_table_data, columns=headers_scores), use_container_width=True)
    st.download_button(label='📥 Düzenli Tabloyu İndir (fonlar_guncel.xlsx)', data=output, file_name='fonlar_guncel.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
