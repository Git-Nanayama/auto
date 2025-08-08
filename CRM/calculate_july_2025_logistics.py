

import csv
from datetime import datetime
import collections

# --- Configuration ---
# File paths
crm_dashboard_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\CRM_Dashboard.csv"
output_logistics_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\monthly_logistics_report.csv"
log_file_path = r"C:\Users\nanayama\Downloads\auto\CRM\logistics_log.txt"

# Logistics cost per box (based on the provided summary file)
# Manually extracted from 物流費単価20250729(1).csv
cost_per_box = {
    "EMS": 1308,
    "神戸": 248,  # Assuming 神户 is 神戸
    "神洲": 606,
    "要町": 1304,
    "FedEx転送": 12190,
    "JD(madme)": 1355,
    "SF CN": 406,
    "ヤマト": 0  # Assuming 'なし' means no cost or handled differently
}

# Column names for CRM_Dashboard.csv
crm_shipping_date_col = "出荷日"
crm_logistics_provider_col = "物流業者"

# Target month and year
target_month = 7
target_year = 2025

# --- Calculation ---
total_logistics_cost = 0.0
cost_by_provider = collections.defaultdict(float)
shipments_by_provider = collections.defaultdict(int)
unhandled_providers = collections.defaultdict(int)

# Clear previous log file
with open(log_file_path, 'w', encoding='utf-8') as log_file:
    log_file.write("Logistics Calculation Log\n")
    log_file.write("="*30 + "\n")

try:
    with open(crm_dashboard_csv_path, 'r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        for row in reader:
            try:
                shipping_date_str = row.get(crm_shipping_date_col, '')
                if not shipping_date_str:
                    continue

                shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')

                if shipping_date.month == target_month and shipping_date.year == target_year:
                    provider_name = row.get(crm_logistics_provider_col, "Unknown").strip()

                    if provider_name in cost_per_box:
                        cost = cost_per_box[provider_name]
                        total_logistics_cost += cost
                        cost_by_provider[provider_name] += cost
                        shipments_by_provider[provider_name] += 1
                    else:
                        unhandled_providers[provider_name] += 1
                        with open(log_file_path, 'a', encoding='utf-8') as log_file:
                            log_file.write(f"Warning: Cost not found for provider '{provider_name}'. Assuming 0 cost for this shipment.\n")

            except (ValueError, KeyError) as e:
                with open(log_file_path, 'a', encoding='utf-8') as log_file:
                    log_file.write(f"Skipping row due to error: {e}. Row data: {row}\n")
                continue

except FileNotFoundError:
    print(f"Error: The file {crm_dashboard_csv_path} was not found.")
    exit()


# --- Output ---
# 1. Write the detailed report to a new CSV file
try:
    with open(output_logistics_csv_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile)
        
        # Header
        writer.writerow(["2025年7月 物流費レポート"])
        writer.writerow([])
        
        # Summary
        writer.writerow(["総物流費", f"{total_logistics_cost:,.2f} JPY"])
        writer.writerow([])
        
        # Breakdown by provider
        writer.writerow(["業者別内訳"])
        writer.writerow(["物流業者", "箱数", "合計費用 (JPY)"])
        for provider, shipments in sorted(shipments_by_provider.items()):
            cost = cost_by_provider[provider]
            writer.writerow([provider, shipments, f"{cost:,.2f}"])
        
        writer.writerow([])
        
        # Unhandled providers
        if unhandled_providers:
            writer.writerow(["単価不明の業者（費用計算対象外）"])
            writer.writerow(["物流業者", "件数"])
            for provider, count in sorted(unhandled_providers.items()):
                writer.writerow([provider, count])

except IOError:
    print(f"Error: Could not write to the file {output_logistics_csv_path}.")

# 2. Print summary to console
print("物流費の計算が完了しました。")
print(f"総物流費: {total_logistics_cost:,.2f} JPY")
print(f"レポートが {output_logistics_csv_path} に出力されました。")
if unhandled_providers:
    print(f"注意: {len(unhandled_providers)}件の単価不明な業者がありました。詳細は {log_file_path} を確認してください。")

