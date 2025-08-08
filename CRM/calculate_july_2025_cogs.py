import csv
from datetime import datetime

# File paths
master_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\master.csv"
crm_dashboard_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\CRM_Dashboard.csv"
output_cogs_csv_path = r"C:\Users\nanayama\Downloads\auto\CRM\monthly_cogs_report.csv"
log_file_path = r"C:\Users\nanayama\Downloads\auto\CRM\cogs_log.txt"

# Column names for master.csv
master_product_name_jp_col = "品名"
master_product_name_cn_col = "产品名称"
master_cost_col = "包装単位\n'@薬価\n（JPY）"

# Column names for CRM_Dashboard.csv
crm_product_name_col = "商品名"
crm_quantity_col = "数量"
crm_shipping_date_col = "出荷日"

# Target month and year
target_month = 7
target_year = 2025

# Function to normalize product names
def normalize_name(name):
    return name.strip().lower().replace(' ', '').replace('　', '') # Remove spaces and full-width spaces

# 1. Load master data (product name to cost mapping)
product_costs = {}
with open(master_csv_path, 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile)
    for row in reader:
        jp_name = row.get(master_product_name_jp_col)
        cn_name = row.get(master_product_name_cn_col)
        cost_str = row.get(master_cost_col)

        if cost_str:
            try:
                cost = float(cost_str.replace(',', ''))
                if jp_name:
                    product_costs[normalize_name(jp_name)] = cost
                if cn_name:
                    product_costs[normalize_name(cn_name)] = cost
            except ValueError:
                with open(log_file_path, 'a', encoding='utf-8') as log_file:
                    log_file.write(f"Warning: Could not convert cost '{cost_str}' for product (JP: {jp_name}, CN: {cn_name}) to float. Skipping.\n")


# 2. Process CRM Dashboard data to calculate COGS
total_cogs = 0.0

with open(crm_dashboard_csv_path, 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile)
    for row in reader:
        try:
            shipping_date_str = row.get(crm_shipping_date_col, '')
            if not shipping_date_str:
                with open(log_file_path, 'a', encoding='utf-8') as log_file:
                    log_file.write(f"Skipping row due to empty shipping date.\n")
                continue

            shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')

            if shipping_date.month == target_month and shipping_date.year == target_year:
                product_name_crm = row.get(crm_product_name_col)
                quantity_str = row.get(crm_quantity_col)

                if product_name_crm and quantity_str:
                    normalized_crm_name = normalize_name(product_name_crm)
                    quantity = float(quantity_str.replace(',', ''))

                    cost_per_unit = product_costs.get(normalized_crm_name)

                    if cost_per_unit is not None:
                        total_cogs += quantity * cost_per_unit
                    else:
                        with open(log_file_path, 'a', encoding='utf-8') as log_file:
                            log_file.write(f"Warning: Cost not found for product '{product_name_crm}' in master data. Skipping COGS calculation for this item.\n")
                else:
                    with open(log_file_path, 'a', encoding='utf-8') as log_file:
                        log_file.write(f"Warning: Missing product name or quantity in row. Skipping COGS calculation for this item.\n")

        except (ValueError, KeyError) as e:
            with open(log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(f"Skipping row due to error: {e}.\n")
            continue

# 3. Write the calculated total COGS to a new CSV file
with open(output_cogs_csv_path, 'w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(["当月の売上原価"])
    writer.writerow([total_cogs])

print(f"Total Cost of Goods Sold for July 2025: {total_cogs}")
print(f"COGS report written to {output_cogs_csv_path}")
print(f"Detailed logs written to {log_file_path}")
