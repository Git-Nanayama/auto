import sys
import os
import pandas as pd
from pathlib import Path
import glob

def find_latest_sales_file():
    """
    Find the most recent sales data CSV file in the ~/Documents directory.
    The filename is assumed to contain '売上伝票'.
    """
    documents_path = Path.home() / "Documents"
    # The filename from the download is "売上伝票_{YYYYMMDDHHMMSS}.csv"
    search_pattern = str(documents_path / "売上伝票_*.csv")
    files = glob.glob(search_pattern)
    if not files:
        return None
    latest_file = max(files, key=os.path.getctime)
    return latest_file

def process_sales_data(input_path, output_path):
    """
    Reads sales data, calculates consumption tax, and saves the processed data.
    """
    try:
        # Read the CSV file. Assuming the file is Shift-JIS encoded.
        df = pd.read_csv(input_path, encoding='sjis')

        # The amount column is likely '今回売上額'. If not, we have a problem.
        if '今回売上額' not in df.columns:
            print("Error: '今回売上額' column not found in the sales data CSV.")
            # Try to find a fallback, but this is a guess.
            potential_cols = [col for col in df.columns if '金額' in col or '売上' in col]
            if not potential_cols:
                 raise ValueError("Could not find a suitable amount column.")
            amount_col = potential_cols[0]
            print(f"Warning: Using fallback amount column '{amount_col}'.")
        else:
            amount_col = '今回売上額'

        # Calculate consumption tax (10%). Assuming the amount is pre-tax.
        # Let's create a new column for the tax.
        df['消費税'] = (df[amount_col] * 0.1).astype(int)

        # Save the processed data
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(output_path, index=False, encoding='utf-8')
        print(f"Processed data saved to {output_path}")

    except FileNotFoundError:
        print(f"Error: Input file not found at {input_path}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred during data processing: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    print("--- Running process_sales_data.py ---")

    sales_file = find_latest_sales_file()

    if not sales_file:
        print("Info: No sales data CSV file found in ~/Documents.")
        print("--- process_sales_data.py finished (no file to process) ---")
        # Create an empty processed file so the next step doesn't fail
        output_csv_path = Path("sales") / "processed_sales.csv"
        output_csv_path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame().to_csv(output_csv_path, index=False)
        print(f"Created empty processed file at: {output_csv_path}")
        sys.exit(0)

    print(f"Found sales data file: {sales_file}")

    output_csv_path = Path("sales") / "processed_sales.csv"

    process_sales_data(sales_file, output_csv_path)

    print("--- process_sales_data.py finished ---")


if __name__ == "__main__":
    main()
