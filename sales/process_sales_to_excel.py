import pandas as pd
import os

script_dir = os.path.dirname(os.path.abspath(__file__))

fuzzy_matched_path = os.path.join(script_dir, "fuzzy_matched_results.csv")
sales_data_path = os.path.join(script_dir, "Sales_Data_202507.csv")
excel_template_path = os.path.join(script_dir, "エクセルインポートサンプル（振替伝票データ）.xlsx")
output_excel_path = os.path.join(script_dir, "会計データ_振替伝票.xlsx")

# Read CSVs with utf-8 encoding
try:
    fuzzy_matched_df = pd.read_csv(fuzzy_matched_path, encoding='utf-8')
    sales_df = pd.read_csv(sales_data_path, encoding='utf-8')
except Exception as e:
    print(f"CSVファイル読み込み中にエラーが発生しました: {e}")
    exit()

# Create a mapping dictionary from Tenkura Name to Matched Freee Name
name_mapping = dict(zip(fuzzy_matched_df['Tenkura Name'], fuzzy_matched_df['Matched Freee Name']))

# Apply the mapping to the '得意先名' column in sales_df
sales_df['得意先名（会計システム用）'] = sales_df['得意先名'].map(name_mapping)

# Load the Excel template
try:
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w') as writer:
        # Read existing sheets from the template
        template_excel = pd.ExcelFile(excel_template_path)
        for sheet_name in template_excel.sheet_names:
            # Copy existing sheets to the new workbook
            existing_df = pd.read_excel(excel_template_path, sheet_name=sheet_name)
            existing_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Write the transformed sales data to 'Sheet1' (or a new sheet if preferred)
        # For now, we'll append it to the end of the existing data in Sheet1
        # If you want to overwrite or write to a specific starting row, please specify.
        # Here, we'll write to a new sheet named 'Sales Data'
        sales_df.to_excel(writer, sheet_name='Sales Data', index=False)

except Exception as e:
    print(f"Excelファイル処理中にエラーが発生しました: {e}")
    exit()

print(f"会計データがExcelテンプレートに挿入され、{output_excel_path} に保存されました。")
