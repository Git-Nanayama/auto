# Sales and Cost of Goods Sold (COGS) Calculation Directive

## 1. Purpose
This directive outlines the process for calculating monthly sales and Cost of Goods Sold (COGS) from provided CSV files. The goal is to automate the generation of two reports: "monthly_sales_report.csv" and "monthly_cogs_report.csv" for a specified month and year.

## 2. Input Files

### 2.1. Sales Data Input
*   **File Path:** `C:\Users\nanayama\Downloads\accounting-auto\CRM\CRM_Dashboard.csv`
*   **Overview:** Contains detailed sales transaction records.
*   **Key Columns:**
    *   `合計金額` (Total Amount): Contains the total sales amount for each transaction. Values may include commas (e.g., "10,000").
    *   `出荷日` (Shipping Date): Contains the date of shipment. Format is `YYYY/MM/DD`.
    *   `商品名` (Product Name): Contains the product name for each transaction.

### 2.2. Master Data Input
*   **File Path:** `C:\Users\nanayama\Downloads\accounting-auto\CRM\master.csv`
*   **Overview:** Contains master data for products, including their cost information.
*   **Key Columns:**
    *   `产品名称` (Product Name - Chinese): Contains the Chinese product name, used for matching with `CRM_Dashboard.csv`.
    *   `品名` (Product Name - Japanese): Contains the Japanese product name, also used for matching with `CRM_Dashboard.csv`.
    *   `包装単位\n'@薬価\n（JPY）` (Cost per Package Unit): Contains the cost information for each product. Values may include commas.

## 3. Output Requirements

### 3.1. Monthly Sales Report
*   **File Path:** `C:\Users\nanayama\Downloads\monthly_sales_report.csv`
*   **Content:** A single value representing the total sales for the target month.
*   **Header:** `当月の売上高`

### 3.2. Monthly COGS Report
*   **File Path:** `C:\Users\nanayama\Downloads\monthly_cogs_report.csv`
*   **Content:** A single value representing the total Cost of Goods Sold for the target month.
*   **Header:** `当月の売上原価`

### 3.3. Log File
*   **File Path:** `C:\Users\nanayama\Downloads\cogs_log.txt`
*   **Content:** Records warnings and errors encountered during the COGS calculation, such as unconvertible costs, missing shipping dates, or product names not found in master data.

## 4. Processing Logic

### 4.1. General
*   **Target Period:** Calculations are performed for July 2025 (Month: 7, Year: 2025).
*   **Encoding:** All CSV files are assumed to be UTF-8 encoded.

### 4.2. Sales Calculation (`calculate_sales.py`)
1.  Read `CRM_Dashboard.csv`.
2.  Iterate through each row:
    *   Parse `出荷日` (Shipping Date) using `YYYY/MM/DD` format.
    *   Filter rows where `出荷日` falls within July 2025.
    *   For matching rows, clean `合計金額` by removing commas (`,`) and convert to a floating-point number.
    *   Accumulate the cleaned `合計金額` to a running total.
3.  Write the final total sales to `monthly_sales_report.csv` with the header `当月の売上高`.

### 4.3. COGS Calculation (`calculate_cogs.py`)
1.  **Load Master Data:**
    *   Read `master.csv`.
    *   Create a dictionary mapping product names to their costs.
    *   Use both `产品名称` (Chinese) and `品名` (Japanese) from `master.csv` as potential keys for product lookup.
    *   Normalize product names (convert to lowercase, remove spaces and full-width spaces) from both master data and CRM data before storing/looking up.
    *   Clean `包装単位\n'@薬価\n（JPY）` by removing commas and convert to a floating-point number.
    *   Log warnings for any cost values that cannot be converted.
2.  **Process CRM Dashboard Data:**
    *   Read `CRM_Dashboard.csv`.
    *   Iterate through each row:
        *   Parse `出荷日` (Shipping Date) using `YYYY/MM/DD` format.
        *   Filter rows where `出荷日` falls within July 2025.
        *   For matching rows:
            *   Retrieve `商品名` and `数量`.
            *   Normalize `商品名` (lowercase, remove spaces).
            *   Clean `数量` by removing commas and convert to a floating-point number.
            *   Look up the cost per unit using the normalized product name in the `product_costs` dictionary.
            *   If a cost is found, calculate `quantity * cost_per_unit` and add to `total_cogs`.
            *   Log warnings if a product name from `CRM_Dashboard.csv` is not found in the master data or if quantity is missing.
            *   Log warnings for empty shipping dates.
3.  Write the final `total_cogs` to `monthly_cogs_report.csv` with the header `当月の売上原価`.
4.  All warnings and errors during COGS calculation are logged to `cogs_log.txt`.

## 5. Special Notes
*   **Date Format:** Ensure `出荷日` in `CRM_Dashboard.csv` strictly adheres to `YYYY/MM/DD`.
*   **Data Cleaning:** Numeric columns (`合計金額`, `包装単位\n'@薬価\n（JPY）`, `数量`) must have commas removed before conversion to numbers.
*   **Product Name Matching:** Product names are normalized (lowercase, spaces removed) for robust matching between `CRM_Dashboard.csv` and `master.csv`.
*   **Error Handling:** The scripts include basic error handling for data conversion and missing columns, logging issues to `cogs_log.txt` to avoid script termination.
*   **Output Directory:** All output files (`monthly_sales_report.csv`, `monthly_cogs_report.csv`, `cogs_log.txt`) will be generated in the `C:\Users\nanayama\Downloads\` directory.
