import pandas as pd
import os
import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))

current_year_month = datetime.datetime.now().strftime('%Y%m')

csv_file_filename = f"freee_import_sales_data_{current_year_month}.csv" # Added YYYYMM suffix
csv_file_path = os.path.join(script_dir, csv_file_filename)

excel_output_filename = f"freee_import_sales_data_{current_year_month}.xlsx"
excel_output_path = os.path.join(script_dir, excel_output_filename)

try:
    # Read the CSV file
    df = pd.read_csv(csv_file_path, encoding='utf-8')
    
    # Write to Excel file
    df.to_excel(excel_output_path, index=False)
    
    print(f"CSVファイルがExcelファイルに変換され、{excel_output_path} に保存されました。")
except Exception as e:
    print(f"ファイル変換中にエラーが発生しました: {e}")
    exit()
