# compute_fund_metrics içindeki allocation kısmını şu şekilde güncelliyoruz:

def fetch_fintables_asset_allocation(fund_code: str) -> dict:
    code = normalize_fund_code(fund_code)
    url = f"https://fintables.com/fonlar/{code.lower()}"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    # Varsayılan: Hisse oranı %50, En büyük hisse ağırlığı %10 (Dengeli)
    result = {"stock_ratio": 50.0, "top_stock_weight": 10.0} 
    try:
        response = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        if response.status_code != 200: return result
        
        # En büyük hisse ağırlığını yakalayan regex (Fintables yapısına göre)
        html = response.text
        match_top = re.search(r'En Büyük Pay["\s:]+([0-9]+(?:\.[0-9]+)?)', html, re.IGNORECASE)
        if match_top:
            result["top_stock_weight"] = float(match_top.group(1))
        return result
    except Exception:
        return result

# Puanlama döngüsü içindeki raw_score formülünü şu şekilde değiştiriyoruz:

# Z_concentration (top_stock_weight üzerinden ceza puanı)
# Eğer top_stock_weight > 20 ise ceza katsayısı devreye girer
z_conc = zscore([item["top_stock_weight"] for item in calculated_funds]) 

for i, item in enumerate(calculated_funds):
    val = item["valor"]
    v_pen = (val * 0.5) if val is not None else 0.0
    
    # YOĞUNLAŞMA CEZASI: (z_conc * 20) kadar puan düşer. 
    # Yani tek hissede yoğunlaşan fonlar hem göreceli hem de mutlak olarak cezalandırılır.
    raw_score = 50 + 10 * z_m[i] + 15 * z_s[i] + 10 * z_c[i] + 10 * z_a[i] + 10 * z_i[i] - 20 * z_conc[i] - v_pen
    
    score = int(round(max(5.0, min(95.0, raw_score))))
    item["running_scores"].append(score)
