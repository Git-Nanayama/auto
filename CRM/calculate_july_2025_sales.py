import csv
from datetime import datetime

def calculate_total_sales(file_path):
    total_sales = 0
    with open(file_path, 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader)  # Skip header row
        
        # Find the index of the '合計金額' and '出荷日' columns
        try:
            total_amount_index = header.index('合計金額')
            shipping_date_index = header.index('出荷日')
        except ValueError as e:
            print(f"Error: {e} column not found in the CSV file.")
            return None

        for row in reader:
            if len(row) > total_amount_index and len(row) > shipping_date_index:
                shipping_date_str = row[shipping_date_index]
                try:
                    shipping_date = datetime.strptime(shipping_date_str, '%Y/%m/%d')
                    if shipping_date.year == 2025 and shipping_date.month == 7:
                        total_amount_str = row[total_amount_index]
                        # Remove commas and convert to float
                        try:
                            total_amount = float(total_amount_str.replace(',', ''))
                            total_sales += total_amount
                        except ValueError:
                            # Handle cases where conversion to float fails (e.g., empty string, non-numeric)
                            continue
                except ValueError:
                    # Handle cases where the date format is incorrect
                    continue
    return total_sales

def output_sales_to_csv(total_sales, output_file_path):
    with open(output_file_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(['2025年7月の売上高'])
        writer.writerow([total_sales])

if __name__ == "__main__":
    input_csv_file = r"C:\Users\nanayama\Downloads\auto\CRM\CRM_Dashboard.csv"
    output_csv_file = r"C:\Users\nanayama\Downloads\auto\CRM\monthly_sales_report.csv"

    sales = calculate_total_sales(input_csv_file)
    if sales is not None:
        output_sales_to_csv(sales, output_csv_file)
        print(f"2025年7月の売上高が {output_csv_file} に出力されました。")