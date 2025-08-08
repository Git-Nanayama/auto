

import csv
from datetime import datetime
import collections

# --- Configuration ---
# File paths
crm_dashboard_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\CRM_Dashboard.csv"
output_logistics_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\monthly_logistics_report_revised.csv"
log_file_path = r"C:\Users\nanayama\Downloads\auto\CRM\logistics_log_revised.txt"

# Revised cost mapping based on analysis
cost_per_box = {
    # Direct Mappings from cost summary
    "EMS": 1308,
    "神洲": 606,
    "要町": 1304,
    "FedEx転送": 12190,
    "JD(madme)": 1355,
    "SF CN": 406,
    "ヤマト": 0, # Domestic shipping
    "神戸": 248,

    # Inferred Aliases
    "神戸物流": 248,    # Alias for 神戸
    "神户物流": 248,    # Alias for 神戸
    "大阪神洲物流": 606, # Alias for 神洲
    "要町SF": 1304,     # Alias for 要町
    "JD": 1355,         # Alias for JD(madme)
    "SF": 406,          # Alias for SF CN
}

# Column names
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

# Clear previous log file
with open(log_file_path, 'w', encoding='utf-8') as log_file:
    log_file.write("Revised Logistics Calculation Log\n=================================\n")

def infer_provider(row):
    # Rule 1: Check "備考" (Remarks) column for keywords
    remarks = row.get(crm_remarks_col, "")
    if "要町" in remarks:
        return "要町", "Inferred from Remarks"
    if "神戸" in remarks:
        return "神戸物流", "Inferred from Remarks"
    if "神洲" in remarks:
        return "神洲", "Inferred from Remarks"
    if "EMS" in remarks:
        return "EMS", "Inferred from Remarks"
    if "SF" in remarks:
        return "SF CN", "Inferred from Remarks"

    # Rule 2: Check "追跡URL" (Tracking URL) for domains
    tracking_url = row.get(crm_tracking_url_col, "")
    if "post.japanpost.jp" in tracking_url:
        return "EMS", "Inferred from Tracking URL"
    if "sf-express.com" in tracking_url:
        return "SF CN", "Inferred from Tracking URL"

    # Rule 3: Check "物流方法" (Logistics Method)
    logistics_method = row.get(crm_logistics_method_col, "")
    if "日本輸送" in logistics_method:
        return "ヤマト", "Inferred from Logistics Method"

    return "Unknown", "Could not infer"

try:
    with open(crm_dashboard_csv_path, 'r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        for i, row in enumerate(reader, 2): # Start from line 2 for logging
            try:
                shipping_date_str = row.get(crm_shipping_date_col, '')
                if not shipping_date_str:
                    continue

                shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')

                if shipping_date.month == target_month and shipping_date.year == target_year:
                    provider_name = row.get(crm_logistics_provider_col, "").strip()
                    inferred_reason = "Provided in data"

                    if not provider_name or provider_name == '-':
                        provider_name, inferred_reason = infer_provider(row)

                    if provider_name in cost_per_box:
                        cost = cost_per_box[provider_name]
                        total_logistics_cost += cost
                        cost_by_provider[provider_name] += cost
                        shipments_by_provider[provider_name] += 1
                        with open(log_file_path, 'a', encoding='utf-8') as log_file:
                            log_file.write(f"Line {i}: Processed. Provider: '{provider_name}', Cost: {cost}, Reason: {inferred_reason}\n")
                    else:
                        unhandled_shipments.append((i, provider_name, row))
                        with open(log_file_path, 'a', encoding='utf-8') as log_file:
                            log_file.write(f"Line {i}: UNHANDLED. Provider: '{provider_name}'. Cost not found.\n")

            except (ValueError, KeyError) as e:
                with open(log_file_path, 'a', encoding='utf-8') as log_file:
                    log_file.write(f"Skipping row {i} due to error: {e}.\n")
                continue

except FileNotFoundError:
    print(f"Error: The file {crm_dashboard_csv_path} was not found.")
    exit()

# --- Output ---
with open(output_logistics_csv_path, 'w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(["2025年7月 物流費レポート (改訂版)"])
    writer.writerow([])
    writer.writerow(["総物流費", f"{total_logistics_cost:,.2f} JPY"])
    writer.writerow([])
    writer.writerow(["業者別内訳"])
    writer.writerow(["物流業者", "箱数", "合計費用 (JPY)"])
    for provider, shipments in sorted(shipments_by_provider.items()):
        cost = cost_by_provider[provider]
        writer.writerow([provider, shipments, f"{cost:,.2f}"])
    writer.writerow([])
    if unhandled_shipments:
        writer.writerow(["単価不明・推測不能の業者（費用計算対象外）"])
        writer.writerow(["行番号", "業者名"])
        for line_num, provider, _ in unhandled_shipments:
            writer.writerow([line_num, provider])

print("改訂版の物流費計算が完了しました。")
print(f"総物流費: {total_logistics_cost:,.2f} JPY")
print(f"レポートが {output_logistics_csv_path} に出力されました。")
if unhandled_shipments:
    print(f"注意: {len(unhandled_shipments)}件の処理不能な項目がありました。詳細は {log_file_path} を確認してください。")
else:
    print("すべての項目が正常に処理されました。")

