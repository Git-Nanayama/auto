import csv
import sys
from datetime import datetime
from collections import defaultdict

def validate_journal(filepath='journal.csv'):
    """
    Validates a journal CSV file based on a set of accounting rules.
    """
    errors = []
    balances = defaultdict(lambda: {'debit': 0, 'credit': 0})

    try:
        with open(filepath, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader) # Skip header

            for i, row in enumerate(reader):
                line_num = i + 2 # 1-based index + header

                # Unpack row safely
                if len(row) != 6:
                    errors.append(f"Line {line_num}: Invalid number of columns. Expected 6, got {len(row)}.")
                    continue

                entry_id, date_str, account, debit_str, credit_str, desc = row

                # 1. Validate Date Format
                try:
                    datetime.strptime(date_str, '%Y/%m/%d')
                except ValueError:
                    errors.append(f"Line {line_num} (ID: {entry_id}): Invalid date format for '{date_str}'. Expected YYYY/MM/DD.")

                # 2. Check for required fields
                if not account:
                    errors.append(f"Line {line_num} (ID: {entry_id}): Account name is missing.")

                # 3. Validate and accumulate amounts
                try:
                    debit = float(debit_str)
                    credit = float(credit_str)

                    if debit < 0 or credit < 0:
                        errors.append(f"Line {line_num} (ID: {entry_id}): Negative amount found (Debit: {debit}, Credit: {credit}).")

                    balances[entry_id]['debit'] += debit
                    balances[entry_id]['credit'] += credit

                except ValueError:
                    errors.append(f"Line {line_num} (ID: {entry_id}): Invalid number format for debit/credit ('{debit_str}', '{credit_str}').")

    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

    # 4. Check for balanced entries
    for entry_id, amounts in balances.items():
        if round(amounts['debit'], 2) != round(amounts['credit'], 2):
            errors.append(f"Entry ID {entry_id}: Debits ({amounts['debit']}) do not equal Credits ({amounts['credit']}).")

    return errors

if __name__ == "__main__":
    validation_errors = validate_journal()

    if validation_errors:
        print("--- Validation Failed ---")
        for error in validation_errors:
            print(error)
        print("-------------------------")
        sys.exit(1)
    else:
        print("✅ Validation successful. All checks passed.")
        sys.exit(0)
