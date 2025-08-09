import sys
import pandas as pd
from pathlib import Path
from datetime import datetime

def convert_to_freee_format(input_path, output_path):
    """
    Converts processed sales data to the freee import format.
    """
    try:
        df = pd.read_csv(input_path, encoding='utf-8')

        if df.empty:
            print("Info: Processed sales data is empty. Creating an empty freee import file.")
            # Create empty file with headers
            pd.DataFrame(columns=[
                '売上まとめ番号', '売上日', '顧客名称', '行の種類', '商品名称',
                '数量', '単価', '消費税', '消費税額の調整', '明細取引タイプ'
            ]).to_csv(output_path, index=False, encoding='sjis')
            return

        # Create a new DataFrame for the freee format
        freee_df = pd.DataFrame()

        # --- Column Mapping (with assumptions) ---
        # If these columns don't exist, the script will fail.
        # This is better than guessing incorrectly.
        date_col = '伝票日付'
        customer_col = '得意先名'
        item_name_col = '品名'
        amount_col = '今回売上額'
        tax_col = '消費税'

        # Check if required columns exist
        required_cols = [date_col, customer_col, item_name_col, amount_col, tax_col]
        for col in required_cols:
            if col not in df.columns:
                print(f"Error: Required column '{col}' not found in {input_path}", file=sys.stderr)
                # Try to provide suggestions based on file columns
                print(f"Available columns: {df.columns.tolist()}", file=sys.stderr)
                sys.exit(1)


        freee_df['売上日'] = pd.to_datetime(df[date_col]).dt.strftime('%Y-%m-%d')
        freee_df['顧客名称'] = df[customer_col]
        freee_df['商品名称'] = df[item_name_col]
        freee_df['単価'] = df[amount_col]
        freee_df['消費税'] = df[tax_col]

        # --- Static and Default Values ---
        freee_df['売上まとめ番号'] = range(1, len(df) + 1) # Unique ID for each row
        freee_df['行の種類'] = '通常'
        freee_df['数量'] = 1
        freee_df['消費税額の調整'] = 1 # Use the value from the '消費税' column
        freee_df['明細取引タイプ'] = '売上' # Assuming 'Sales' as the transaction type

        # Reorder columns to a logical format
        final_columns = [
            '売上まとめ番号', '売上日', '顧客名称', '行の種類', '商品名称',
            '数量', '単価', '消費税', '消費税額の調整', '明細取引タイプ'
        ]
        freee_df = freee_df[final_columns]

        # Save to CSV in Shift-JIS format for compatibility with Japanese systems
        freee_df.to_csv(output_path, index=False, encoding='sjis')
        print(f"Successfully converted data to freee format at: {output_path}")

    except FileNotFoundError:
        print(f"Error: Input file not found at {input_path}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred during conversion to freee format: {e}", file=sys.stderr)
        sys.exit(1)

def main():
    print("--- Running convert_to_freee_format.py ---")

    input_csv_path = Path("sales") / "processed_sales.csv"

    timestamp = datetime.now().strftime("%Y%m%d")
    output_filename = f"freee_import_sales_data_{timestamp}.csv"
    output_csv_path = Path("sales") / output_filename

    convert_to_freee_format(input_csv_path, output_csv_path)

    print("--- convert_to_freee_format.py finished ---")

if __name__ == "__main__":
    main()
