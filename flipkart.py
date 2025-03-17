import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string

# --- Step 1: Read the CSV data ---
csv_file_path = "Untitled spreadsheet - BL__Products__default_CSV_2025-02-12_18_43.csv"
df = pd.read_csv(csv_file_path)

# --- Step 2: Open the Excel template ---
template_path = "C_sling-bag_fd927b15e6244645_1703-2438FK_REQH2ILIQXHAH.xlsx"
wb = openpyxl.load_workbook(template_path)
sheet = wb["sling_bag"]

# --- Step 3: Define the mapping ---
# For each template column (specified as a letter), we set a tuple:
# (is_hardcoded, value). If is_hardcoded is True, the value is used directly;
# otherwise, the value is used as the CSV column header.
mapping = {
    "G":  (False, "Seller SKU"),
    "J":  (False, "Maximum Retail Price (Sell on Amazon)"),
    "K":  (False, "Your Price INR (Sell on Amazon, IN)"),
    "L":  (True,  "seller"),
    "P":  (True,  "FLIPKART"),
    "X":  (False, "External product information"),
    "Z":  (True,  "IN"),
    "AA": (False, "Manufacturer"),
    "AB": (False, "Manufacturer"),
    "AD": (True,  "12%"),
    "AF": (False, "Brand Name"),
    "AG": (False, "Model Name"),
    "AH": (False, "Color"),
    "AJ": (False, "Model Number"),
    "AK": (False, "Item Type"),
    "AL": (True,  "casual wear"),
    "AM": (True,  "casual"),
    "AN": (False, "Material"),
    "AO": (True,  "1"),
    "AP": (False, "Item height"),
    "AQ": (True,  "[cm]"),
    "AR": (False, "Item width"),
    "AS": (True,  "[cm]"),
    "AT": (True, "test")
}

# --- Step 4: Fill the template ---
# Start inserting data from row 5 onward (rows 1-4 remain unchanged)
start_row = 5

for i, csv_row in df.iterrows():
    template_row = start_row + i
    for col_letter, (is_hard, value_spec) in mapping.items():
        # Convert the column letter (e.g., "G") to a column index
        col_index = column_index_from_string(col_letter)
        if is_hard:
            # For hard coded fields, use the specified value directly.
            value = value_spec
        else:
            # For CSV-based fields, get the value from the CSV row.
            value = csv_row.get(value_spec)
        sheet.cell(row=template_row, column=col_index).value = value

# --- Step 5: Save the updated workbook ---
output_path = "C_sling-bag_filled.xlsx"
wb.save(output_path)
print("Data has been filled into the template starting from row 5 in the sling_bag tab.")
