import pandas as pd
import datetime
import glob
import os

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define file paths relative to the script directory
account_master_path = os.path.join(script_dir, "Account_Master_all.csv")

# Find the latest sales slip file to process in the script's directory
sales_files = glob.glob(os.path.join(script_dir, "売上伝票_*.csv"))
if not sales_files:
    print("エラー: 売上伝票ファイルが見つかりません。")
    exit()

latest_sales_file = max(sales_files, key=os.path.getctime)

# Read the CSV files using utf-8 encoding (determined during debugging)
try:
    account_master_df = pd.read_csv(account_master_path, encoding='utf-8')
    sales_df = pd.read_csv(latest_sales_file, encoding='utf-8')
except Exception as e:
    print(f"ファイル読み込み中にエラーが発生しました: {e}")
    exit()

# Strip whitespace from key columns to ensure proper matching
account_master_df['取引先名'] = account_master_df['取引先名'].str.strip()
sales_df['得意先名'] = sales_df['得意先名'].str.strip()

# Create a mapping from "取引先名" to "部門"
department_map = account_master_df.set_index('取引先名')['部門'].to_dict()

# Add the "部門" column to the sales data
sales_df['部門'] = sales_df['得意先名'].map(department_map)

# Generate the output filename for the current month
current_year_month = datetime.datetime.now().strftime('%Y%m')
output_filename = f"Sales_Data_{current_year_month}.csv"
output_path = os.path.join(script_dir, output_filename)

# Save the updated dataframe to a new CSV file using utf-8 encoding
try:
    sales_df.to_csv(output_path, index=False, encoding='utf-8')
except Exception as e:
    print(f"ファイル保存中にエラーが発生しました: {e}")
    exit()

print(f"処理が完了しました。出力ファイル: {output_path}")
