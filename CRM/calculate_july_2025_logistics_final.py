import csv
from datetime import datetime
import collections
import math

# --- Configuration ---
# File paths
crm_dashboard_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\CRM_Dashboard.csv"
output_logistics_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\monthly_logistics_report_final.csv"
log_file_path = r"C:\Users\nanayama\Downloads\auto\CRM\logistics_log_final.txt"

# Final cost mapping based on analysis and user feedback
cost_per_box = {
    "EMS": 1308,
    "神洲": 606,
    "要町": 1304,
    "FedEx転送": 12190,
    "JD(madme)": 1355,
    "SF CN": 406,
    "ヤマト": 0,
    "神戸": 248,
    "神戸物流": 248, "神户物流": 248,
    "大阪神洲物流": 606,
    "要町SF": 1304,
    "JD": 1355,
    "SF": 406,
}

# Column names
crm_product_name_col = "商品名"
crm_quantity_col = "数量"
crm_shipping_date_col = "出荷日"
crm_logistics_provider_col = "物流業者"
crm_remarks_col = "備考"
crm_tracking_url_col = "追跡URL"
crm_logistics_method_col = "物流方法"

# Target period
target_month = 7
target_year = 2025

# --- Main Logic ---
total_logistics_cost = 0.0
cost_by_provider = collections.defaultdict(float)
shipments_by_provider = collections.defaultdict(int)
unhandled_shipments = []

with open(log_file_path, 'w', encoding='utf-8') as log_file:
    log_file.write("Final Logistics Calculation Log\n=================================\n")

def infer_provider(row):
    remarks = row.get(crm_remarks_col, "")
    if "要町" in remarks: return "要町", "Inferred from Remarks"
    if "神戸" in remarks: return "神戸物流", "Inferred from Remarks"
    if "神洲" in remarks: return "神洲", "Inferred from Remarks"
    if "EMS" in remarks: return "EMS", "Inferred from Remarks"
    if "SF" in remarks: return "SF CN", "Inferred from Remarks"

    tracking_url = row.get(crm_tracking_url_col, "")
    if "post.japanpost.jp" in tracking_url: return "EMS", "Inferred from Tracking URL"
    if "sf-express.com" in tracking_url: return "SF CN", "Inferred from Tracking URL"

    logistics_method = row.get(crm_logistics_method_col, "")
    if "日本輸送" in logistics_method: return "ヤマト", "Inferred from Logistics Method"

    return "Unknown", "Could not infer"

try:
    with open(crm_dashboard_csv_path, 'r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        for i, row in enumerate(reader, 2):
            try:
                shipping_date_str = row.get(crm_shipping_date_col, '')
                if not shipping_date_str: continue

                shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')

                if shipping_date.month == target_month and shipping_date.year == target_year:
                    provider_name = row.get(crm_logistics_provider_col, "").strip()
                    inferred_reason = "Provided"

                    if not provider_name or provider_name == '-':
                        provider_name, inferred_reason = infer_provider(row)

                    if provider_name in cost_per_box:
                        product_name = row.get(crm_product_name_col, "").lower()
                        quantity_str = row.get(crm_quantity_col, "0")
                        quantity = float(quantity_str.replace(',', '')) if quantity_str else 0

                        num_boxes = 0
                        if "mounjaro" in product_name:
                            num_boxes = math.ceil(quantity / 2)
                            box_calc_reason = f"Mounjaro rule (ceil({quantity}/2))"
                        else:
                            num_boxes = quantity
                            box_calc_reason = f"Default rule (quantity)"

                        cost_per_unit_box = cost_per_box[provider_name]
                        cost_for_row = num_boxes * cost_per_unit_box

                        total_logistics_cost += cost_for_row
                        cost_by_provider[provider_name] += cost_for_row
                        shipments_by_provider[provider_name] += num_boxes
                        
                        with open(log_file_path, 'a', encoding='utf-8') as log:
                            log.write(f"L{i}: Provider '{provider_name}' ({inferred_reason}). Product: '{row[crm_product_name_col][:20]}'. Qty: {quantity}. Boxes: {num_boxes} ({box_calc_reason}). Cost: {cost_for_row}\n")

                    else:
                        unhandled_shipments.append((i, provider_name))
                        with open(log_file_path, 'a', encoding='utf-8') as log:
                            log.write(f"L{i}: UNHANDLED. Provider: '{provider_name}' ({inferred_reason}). Cost not found.\n")

            except (ValueError, KeyError) as e:
                with open(log_file_path, 'a', encoding='utf-8') as log:
                    log.write(f"Skipping row {i} due to error: {e}.\n")

except FileNotFoundError:
    print(f"Error: {crm_dashboard_csv_path} not found.")
    exit()

# --- Output ---
with open(output_logistics_csv_path, 'w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(["2025年7月 物流費レポート (最終版)"])
    writer.writerow([])
    writer.writerow(["総物流費", f"{total_logistics_cost:,.2f} JPY"])
    writer.writerow([])
    writer.writerow(["業者別内訳"])
    writer.writerow(["物流業者", "合計箱数", "合計費用 (JPY)"])
    for provider, total_boxes in sorted(shipments_by_provider.items()):
        total_cost = cost_by_provider[provider]
        writer.writerow([provider, int(total_boxes), f"{total_cost:,.2f}"])
    writer.writerow([])
    if unhandled_shipments:
        writer.writerow(["単価不明・推測不能の業者（費用計算対象外）"])
        writer.writerow(["行番号", "業者名"])
        for line_num, provider in unhandled_shipments:
            writer.writerow([line_num, provider])

print("最終版の物流費計算が完了しました。")
print(f"総物流費: {total_logistics_cost:,.2f} JPY")
print(f"レポート: {output_logistics_csv_path}")
if unhandled_shipments:
    print(f"注意: {len(unhandled_shipments)}件の処理不能項目あり。詳細はログを確認: {log_file_path}")
else:
    print("すべての項目が正常に処理されました。")