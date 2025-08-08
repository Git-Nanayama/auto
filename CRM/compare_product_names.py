import csv

crm_product_names_file = r"C:\Users\nanayama\Downloads\accounting-auto\CRM\unique_crm_product_names.txt"
master_product_names_file = r"C:\Users\nanayama\Downloads\accounting-auto\CRM\unique_master_product_names.txt"
mismatched_output_file = r"C:\Users\nanayama\Downloads\accounting-auto\CRM\mismatched_product_names.txt"

crm_names = set()
master_names = set()

with open(crm_product_names_file, 'r', encoding='utf-8') as f:
    for line in f:
        crm_names.add(line.strip())

with open(master_product_names_file, 'r', encoding='utf-8') as f:
    for line in f:
        master_names.add(line.strip())


mismatched_names = crm_names - master_names

with open(mismatched_output_file, 'w', encoding='utf-8') as outfile:
    if mismatched_names:
        outfile.write("Product names in CRM_Dashboard.csv but not in master.csv:\n")
        for name in sorted(list(mismatched_names)):
            outfile.write(name + '\n')
    else:
        outfile.write("All product names in CRM_Dashboard.csv are present in master.csv.\n")

print(f"Mismatched product names written to {mismatched_output_file}")