import pandas as pd
import os
import datetime
import glob

script_dir = os.path.dirname(os.path.abspath(__file__))

current_year_month = datetime.datetime.now().strftime('%Y%m')

sales_data_filename = f"Sales_Data_{current_year_month}.csv"
sales_data_path = os.path.join(script_dir, sales_data_filename)

fuzzy_matched_filename = "fuzzy_matched_results.csv" # Removed YYYYMM suffix
fuzzy_matched_path = os.path.join(script_dir, fuzzy_matched_filename)

freee_import_format_path = os.path.join(script_dir, "Freee_import_format.csv")

output_freee_csv_filename = f"freee_import_sales_data_{current_year_month}.csv"
output_freee_csv_path = os.path.join(script_dir, output_freee_csv_filename)

# Read CSVs with utf-8 encoding
try:
    sales_df = pd.read_csv(sales_data_path, encoding='utf-8')
    fuzzy_matched_df = pd.read_csv(fuzzy_matched_path, encoding='utf-8')
    freee_format_df = pd.read_csv(freee_import_format_path, encoding='utf-8')
except Exception as e:
    print(f"CSVファイル読み込み中にエラーが発生しました: {e}")
    exit()

# Create a mapping dictionary from Tenkura Name to Matched Freee Name
name_mapping = dict(zip(fuzzy_matched_df['Tenkura Name'], fuzzy_matched_df['Matched Freee Name']))

# Prepare an empty DataFrame for freee import format
freee_data = pd.DataFrame(columns=freee_format_df.columns)

# Iterate through sales data and populate freee_data
for index, row in sales_df.iterrows():
    # Map '得意先名' to freee's '借方取引先' using the fuzzy matching results
    freee_customer_name = name_mapping.get(row['得意先名'], row['得意先名']) # Use original if no match

    # Create a new row for freee import format
    new_row = pd.Series(index=freee_data.columns)

    # Direct mappings
    new_row['日付'] = pd.to_datetime(str(row['売上日'])).strftime('%Y/%m/%d') # Format date
    new_row['伝票番号'] = '' # 空欄にする
    new_row['摘要'] = row['摘要'] if pd.notna(row['摘要']) else '' # 元データにあれば転記、なければ空欄

    # Determine debit and credit accounts based on the sign of the sales amount
    amount = abs(row['売上金額合計（税込）'])
    tax_amount = abs(row['消費税額合計'])

    if row['売上金額合計（税込）'] >= 0:
        # Normal sales: Debit 売掛金, Credit 売上高
        debit_account = '売掛金'
        credit_account = '売上高'
    else:
        # Negative sales (returns/discounts): Debit 売上高, Credit 売掛金
        debit_account = '売上高'
        credit_account = '売掛金'

    # Debit side
    new_row['借方勘定科目'] = debit_account
    new_row['借方科目コード'] = '' 
    new_row['借方補助科目'] = ''
    new_row['借方取引先'] = freee_customer_name
    new_row['借方取引先コード'] = '' 
    new_row['借方部門'] = row['部門'] if pd.notna(row['部門']) else ''
    new_row['借方品目'] = ''
    new_row['借方メモタグ'] = ''
    new_row['借方セグメント1'] = ''
    new_row['借方セグメント2'] = ''
    new_row['借方セグメント3'] = ''
    new_row['借方金額'] = amount
    new_row['借方税区分'] = '課税売上10%' # Assuming 10% tax, adjust as needed
    new_row['借方税額'] = tax_amount

    # Credit side
    new_row['貸方勘定科目'] = credit_account
    new_row['貸方科目コード'] = ''
    new_row['貸方補助科目'] = ''
    new_row['貸方取引先'] = freee_customer_name
    new_row['貸方取引先コード'] = ''
    new_row['貸方部門'] = row['部門'] if pd.notna(row['部門']) else ''
    new_row['貸方品目'] = ''
    new_row['貸方メモタグ'] = ''
    new_row['貸方セグメント1'] = ''
    new_row['貸方セグメント2'] = ''
    new_row['貸方セグメント3'] = ''
    new_row['貸方金額'] = amount
    new_row['貸方税区分'] = '課税売上10%' # Assuming 10% tax, adjust as needed
    new_row['貸方税額'] = tax_amount

    # Other fields
    new_row['決算整理仕訳'] = '' # 空欄にする

    freee_data = pd.concat([freee_data, pd.DataFrame([new_row])], ignore_index=True)

# Save the transformed data to a new CSV file
try:
    freee_data.to_csv(output_freee_csv_path, index=False, encoding='utf-8')
except Exception as e:
    print(f"freeeインポート用CSVファイル保存中にエラーが発生しました: {e}")
    exit()

print(f"Sales_Data_202507.csv のデータが freee インポートフォーマットに変換され、{output_freee_csv_path} に保存されました。")