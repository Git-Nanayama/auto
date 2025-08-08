

import csv

crm_dashboard_csv_path = r"C:\Users\nanayama\Downloads\accounting-auto\CRM\CRM_Dashboard.csv"
output_file_path = r"C:\Users\nanayama\Downloads\accounting-auto\CRM\unique_crm_product_names.txt"
crm_product_name_col = "商品名"

unique_product_names = set()

with open(crm_dashboard_csv_path, 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile)
    for row in reader:
        product_name = row.get(crm_product_name_col)
        if product_name:
            unique_product_names.add(product_name)

with open(output_file_path, 'w', encoding='utf-8') as outfile:
    for name in sorted(list(unique_product_names)):
        outfile.write(name + '\n')

print(f"Unique product names written to {output_file_path}")

