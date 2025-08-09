import csv
import os

# Define the data, including some intentional errors
journal_entries = [
    # Good entry
    ['1', '2025/07/31', 'Cash', '1000', '0', 'Initial deposit'],
    ['1', '2025/07/31', 'Capital', '0', '1000', 'Initial deposit'],

    # Bad entry: Unbalanced
    ['2', '2025/08/01', 'Office Supplies', '500', '0', 'Bought supplies'],
    ['2', '2025/08/01', 'Cash', '0', '400', 'Paid for supplies - ERROR: UNBALANCED'], # Debit 500, Credit 400

    # Bad entry: Bad date format
    ['3', '2025-08-02', 'Rent Expense', '1200', '0', 'Paid rent - ERROR: BAD DATE'],
    ['3', '2025-08-02', 'Cash', '0', '1200', 'Paid rent'],

    # Bad entry: Negative amount
    ['4', '2025/08/03', 'Sales Revenue', '0', '2000', 'Customer payment'],
    ['4', '2025/08/03', 'Cash', '-2000', '0', 'Customer payment - ERROR: NEGATIVE AMOUNT'], # Should be positive debit

    # Bad entry: Missing account
    ['5', '2025/08/04', '', '300', '0', 'Misc expense - ERROR: MISSING ACCOUNT'],
    ['5', '2025/08/04', 'Cash', '0', '300', 'Misc expense'],

    # Good entry
    ['6', '2025/08/05', 'Utilities', '150', '0', 'Electricity bill'],
    ['6', '2025/08/05', 'Cash', '0', '150', 'Electricity bill'],
]

# Define the header for the CSV file
header = ['entry_id', 'date', 'account', 'debit', 'credit', 'description']

# Define the output file path
output_file = 'journal.csv'

# Create and write to the CSV file
try:
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(journal_entries)
    print(f"Successfully created '{output_file}'")

except IOError as e:
    print(f"Error writing to file: {e}")
