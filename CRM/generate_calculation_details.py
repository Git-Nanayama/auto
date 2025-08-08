import csv
import math
from datetime import datetime

# --- CONFIGURATION ---
# Input files
crm_dashboard_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\CRM_Dashboard.csv"
master_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\master.csv"

# Output file
output_report_path = r"C:\Users\nanayama\Downloads\auto\CRM\calculation_details_report.csv"

# --- DATA LOADING ---
# COGS: Load master data for cost lookup
product_costs = {}
def normalize_name(name):
    return name.strip().lower().replace(' ', '').replace('　', '')

with open(master_csv_path, 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile)
    for row in reader:
        jp_name = row.get("品名")
        cn_name = row.get("产品名称")
        cost_str = row.get("包装単位\n'@薬価\n（JPY）")
        if cost_str:
            try:
                cost = float(cost_str.replace(',', ''))
                if jp_name: product_costs[normalize_name(jp_name)] = cost
                if cn_name: product_costs[normalize_name(cn_name)] = cost
            except ValueError:
                continue # Skip if cost is not a valid number

# COGS: Override for specific products as requested
cogs_override_products = {
    normalize_name("救心丸 60粒"): 1000,
    normalize_name("救心丸 30粒"): 1000,
    normalize_name("レポフ口キサシン錠500mg"): 1000,
    normalize_name("塩野義製薬 シナール配合錠 100錠 左旋维他命c Cinal vc美白淡斑维生素片 100錠"): 1000,
}

# Logistics: Cost per box mapping
logistics_cost_per_box = {
    "EMS": 1308, "神洲": 606, "要町": 1304, "FedEx転送": 12190,
    "JD(madme)": 1355, "SF CN": 406, "ヤマト": 0, "神戸": 248,
    "神戸物流": 248, "神户物流": 248, "大阪神洲物流": 606, "要町SF": 1304,
    "JD": 1355, "SF": 406,
}

# --- HELPER FUNCTIONS ---
def infer_provider(row):
    remarks = row.get("備考", "")
    if "要町" in remarks: return "要町", "備考から推測"
    if "神戸" in remarks: return "神戸物流", "備考から推測"
    if "神洲" in remarks: return "神洲", "備考から推測"
    if "EMS" in remarks: return "EMS", "備考から推測"
    if "SF" in remarks: return "SF CN", "備考から推測"
    tracking_url = row.get("追跡URL", "")
    if "post.japanpost.jp" in tracking_url: return "EMS", "追跡URLから推測"
    if "sf-express.com" in tracking_url: return "SF CN", "追跡URLから推測"
    if "日本輸送" in row.get("物流方法", ""): return "ヤマト", "物流方法から推測"
    return "不明", "推測不能"

# --- MAIN PROCESSING ---
output_rows = []
total_sales = 0
total_cogs = 0
total_logistics_cost = 0

with open(crm_dashboard_csv_path, 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile)
    for i, row in enumerate(reader, 2):
        try:
            shipping_date_str = row.get('出荷日', '')
            if not shipping_date_str: continue
            shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')
            if not (shipping_date.month == 7 and shipping_date.year == 2025):
                continue

            # --- Data Extraction ---
            product_name = row.get('商品名', '')
            norm_product_name = normalize_name(product_name)
            quantity_str = row.get('数量', '0')
            quantity = float(quantity_str.replace(',', '')) if quantity_str else 0
            sales_str = row.get('合計金額', '0')
            sales = float(sales_str.replace(',', '')) if sales_str else 0

            # --- Calculations ---
            # 1. Sales
            calculated_sales = sales
            total_sales += calculated_sales

            # 2. COGS
            unit_cost, cogs_notes = (None, "不明")
            if norm_product_name in cogs_override_products:
                unit_cost = cogs_override_products[norm_product_name]
                cogs_notes = "固定単価(1000円)"
            elif norm_product_name in product_costs:
                unit_cost = product_costs[norm_product_name]
                cogs_notes = "マスター参照"
            calculated_cogs = (quantity * unit_cost) if unit_cost is not None else 0
            total_cogs += calculated_cogs

            # 3. Logistics
            provider = row.get('物流業者', '').strip()
            provider_reason = "元データ"
            if not provider or provider == '-':
                provider, provider_reason = infer_provider(row)
            
            box_cost, log_notes, boxes, box_rule = (0, "不明", 0, "")
            if provider in logistics_cost_per_box:
                box_cost = logistics_cost_per_box[provider]
                if "mounjaro" in product_name.lower():
                    boxes = math.ceil(quantity / 2)
                    box_rule = f"マンジャロルール (数量/2)"
                else:
                    boxes = quantity
                    box_rule = "通常ルール (数量)"
                log_notes = "計算済"
            calculated_logistics = boxes * box_cost
            total_logistics_cost += calculated_logistics

            # --- Append Row ---
            output_rows.append([
                row.get('Unique ID'), row.get('出荷日'), product_name, quantity, row.get('合計金額'),
                calculated_sales, # Sales
                unit_cost if unit_cost is not None else '', cogs_notes, calculated_cogs, # COGS
                provider, provider_reason, boxes, box_rule, box_cost, calculated_logistics, # Logistics
            ])

        except (ValueError, KeyError) as e:
            # Log errors if any
            continue

# --- WRITE OUTPUT ---
with open(output_report_path, 'w', newline='', encoding='utf-8-sig') as outfile:
    writer = csv.writer(outfile)
    # Header
    header = [
        'Unique ID', '出荷日', '商品名', '数量', '元データ合計金額',
        '売上高', '原価単価', '原価計算根拠', '売上原価',
        '物流業者(推測後)', '業者推測根拠', '箱数', '箱数計算ルール', '箱単価', '物流費'
    ]
    writer.writerow(header)
    writer.writerows(output_rows)
    # Footer with totals
    writer.writerow([])
    writer.writerow(['', '', '', '合計', '', f'{total_sales:,.0f}', '', '', f'{total_cogs:,.0f}', '', '', '', '', '', f'{total_logistics_cost:,.0f}'])

print(f"詳細な計算レポートが {output_report_path} に出力されました。")
