import csv
from datetime import datetime

input_csv_path = r"C:\Users\nanayama\Downloads\accounting-auto\CRM\CRM_Dashboard.csv"
output_csv_path = r"C:\Users\nanayama\Downloads\monthly_sales_report.csv"
total_amount_col = "合計金額"
shipping_date_col = "出荷日"
target_month = 7
target_year = 2025

total_sales = 0.0

with open(input_csv_path, 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile)
    for row in reader:
        try:
            shipping_date_str = row.get(shipping_date_col, '')
            if not shipping_date_str:
                print(f"Skipping row due to empty shipping date.")
                continue

            # Parse shipping date
            shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')

            # Check if the date is in July 2025
            if shipping_date.month == target_month and shipping_date.year == target_year:
                # Clean and convert total amount
                amount_str = row[total_amount_col].replace(',', '')
                amount = float(amount_str)
                total_sales += amount
        except (ValueError, KeyError) as e:
            # Handle cases where date or amount parsing fails, or column not found
            print(f"Skipping row due to error: {e}.")
            continue

with open(output_csv_path, 'w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(["当月の売上高"])
    writer.writerow([total_sales])

print(f"Total sales for July 2025: {total_sales}")
print(f"Report written to {output_csv_path}")