import csv

def calculate_total_sales(file_path):
    total_sales = 0
    with open(file_path, 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader)  # Skip header row
        
        # Find the index of the '合計金額' column
        try:
            total_amount_index = header.index('合計金額')
        except ValueError:
            print("Error: '合計金額' column not found in the CSV file.")
            return None

        for row in reader:
            if len(row) > total_amount_index:
                total_amount_str = row[total_amount_index]
                # Remove commas and convert to float
                try:
                    total_amount = float(total_amount_str.replace(',', ''))
                    total_sales += total_amount
                except ValueError:
                    # Handle cases where conversion to float fails (e.g., empty string, non-numeric)
                    continue
    return total_sales

def output_sales_to_csv(total_sales, output_file_path):
    with open(output_file_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(['当月の売上高'])
        writer.writerow([total_sales])

if __name__ == "__main__":
    input_csv_file = r"C:\Users\nanayama\Downloads\CRM_Dashboard - 7月.csv"
    output_csv_file = r"C:\Users\nanayama\Downloads\monthly_sales_report.csv"

    sales = calculate_total_sales(input_csv_file)
    if sales is not None:
        output_sales_to_csv(sales, output_csv_file)
        print(f"当月の売上高が {output_csv_file} に出力されました。")