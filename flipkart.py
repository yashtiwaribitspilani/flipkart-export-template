import openpyxl
import pandas as pd
import re
from urllib.parse import urlparse
from openpyxl.utils import column_index_from_string

# --- Helper validation functions ---

def is_positive_integer(value):
    try:
        num = float(value)
        return num.is_integer() and num > 0
    except (ValueError, TypeError):
        return False

def is_number(value):
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False

def is_valid_decimal_or_int(value):
    return is_number(value)

def is_valid_unit(value):
    return str(value) in {"cm", "mm", "inch"}

def is_valid_country(value):
    return bool(re.match(r"^[A-Z][a-zA-Z\s]*$", str(value).strip()))

def is_valid_url(value):
    try:
        result = urlparse(str(value).strip())
        return all([result.scheme, result.netloc])
    except Exception:
        return False

# --- Allowed sets for specific columns ---
allowed_AD = {"GST_0", "GST_12", "GST_18", "GST_3", "GST_5", "GST_APPAREL"}
allowed_AI = {"Beige", "Black", "Blue", "Brown", "Clear", "Gold", "Green", "Grey", 
              "Khaki", "Maroon", "Multicolor", "Orange", "Pink", "Purple", "Red", 
              "Silver", "Tan", "White", "Yellow"}
allowed_AK = {"Clutch", "Hand-held Bag", "Hobo", "Messenger Bag", "Satchel", 
              "Shoulder Bag", "Sling Bag", "Tote"}
allowed_AL = {"Boys", "Boys & Girls", "Girls", "Men", "Men & Women", "Women"}
allowed_AM = {"Casual", "Evening/Party", "Formal", "Sports"}
allowed_AN = {"Acrylic", "Beads", "Brocade", "Canvas", "Cotton", "Denim", "Fabric", 
              "Flex", "Genuine Leather", "Juco", "Jute", "Leatherette", "Metal", 
              "Natural Fibre", "PU", "Plastic", "Polyester", "Rexine", "Satin", 
              "Silicon", "Silk", "Synthetic Leather", "Tyvek", "Velvet", "Wood", "Wool"}

# --- Step 1: Read the XLSX data (ignoring header names) ---
xlsx_file_path = "sample_data (1).xlsx"  # Replace with your XLSX file path
df = pd.read_excel(xlsx_file_path, header=None)  

# --- Step 2: Open the Excel template ---
template_path = "C_sling-bag_fd927b15e6244645_1703-2438FK_REQH2ILIQXHAH.xlsx"
wb = openpyxl.load_workbook(template_path)
sheet = wb["sling_bag"]

# --- Step 3: Define the mapping ---
# Mapping: target template column letter -> (is_required, source Excel column letter)
mapping = {
    "G":  (True, "G"),
    "J":  (True, "J"),  # Must be positive integer
    "K":  (True, "K"),  # Must be positive integer
    "L":  (True, "L"),  # Must be exactly "Seller"
    "N":  (True, "N"),  # Must be positive integer
    "O":  (True, "O"),  # Must be positive integer
    "P":  (True, "P"),  # Must be exactly "Flipkart"
    "Q":  (True, "Q"),  # Must be positive integer
    "R":  (True, "R"),  # Must be positive integer
    "S":  (True, "S"),  # Must be positive integer
    "T":  (True, "T"),  # Can be int or decimal
    "U":  (True, "U"),  # Can be int or decimal
    "V":  (True, "V"),  # Can be int or decimal
    "W":  (True, "W"),  # Can be int or decimal
    "Z":  (True, "Z"),  # Valid country (first letter capital)
    "AA": (True, "AA"),
    "AB": (True, "AB"),
    "AD": (True, "AD"),  # Only allowed GST values
    "AF": (True, "AF"),
    "AG": (True, "AG"),  # AF and AG must not be the same
    "AH": (True, "AH"),
    "AI": (True, "AI"),  # Must be one of allowed_AI
    "AJ": (True, "AJ"),
    "AK": (True, "AK"),  # Must be one of allowed_AK
    "AL": (True, "AL"),  # Must be one of allowed_AL
    "AM": (True, "AM"),  # Must be one of allowed_AM
    "AN": (True, "AN"),  # Must be one of allowed_AN
    "AO": (True, "AO"),  # Must be a number
    "AP": (True, "AP"),  # Can be int or decimal
    "AQ": (True, "AQ"),  # Must be one of allowed units ("cm", "mm", "inch")
    "AR": (True, "AR"),  # Can be int or decimal
    "AS": (True, "AS"),  # Must be one of allowed units ("cm", "mm", "inch")
    "AT": (True, "AT")   # Must be a valid URL
}

# --- Step 4: Prepare for invalid data tracking and valid row counter ---
invalid_data_rows = []  # Rows that fail validation
start_row = 5  # Row in the template where data will be written
valid_row_counter = start_row  # Counter for valid rows in the template

# --- Step 5: Process each row from the input file ---
# Since we read without header, DataFrame columns are integer indexes (0, 1, 2, …)
for i, row in df.iterrows():
    error_list = []
    # Dictionary to hold values for writing and cross-field validation
    row_values = {}
    
    # Validate each mapped field
    for target_col, (is_required, source_letter) in mapping.items():
        # Convert target letter (for template) to column index (1-based)
        target_index = column_index_from_string(target_col)
        # Convert source letter to zero-based index for df (Excel: A=1 so subtract 1)
        source_index = column_index_from_string(source_letter) - 1
        
        try:
            value = row[source_index]
        except IndexError:
            value = None

        row_values[source_letter] = value

        # Check required fields
        if is_required and (pd.isnull(value) or str(value).strip() == "" or str(value).lower() == "nan"):
            error_list.append(f"Missing required value in column {source_letter}")
        
        # Column-specific validations:
        if source_letter == "J":
            if not is_positive_integer(value):
                error_list.append(f"Column J must be a positive integer; got '{value}'")
        if source_letter == "K":
            if not is_positive_integer(value):
                error_list.append(f"Column K must be a positive integer; got '{value}'")
        if source_letter == "L":
            if str(value).strip() != "Seller":
                error_list.append(f"Column L must be 'Seller'; got '{value}'")
        if source_letter == "N":
            if not is_positive_integer(value):
                error_list.append(f"Column N must be a positive integer; got '{value}'")
        if source_letter == "O":
            if not is_positive_integer(value):
                error_list.append(f"Column O must be a positive integer; got '{value}'")
        if source_letter == "P":
            if str(value).strip() != "Flipkart":
                error_list.append(f"Column P must be 'Flipkart'; got '{value}'")
        if source_letter in {"Q", "R", "S"}:
            if not is_positive_integer(value):
                error_list.append(f"Column {source_letter} must be a positive integer; got '{value}'")
        if source_letter in {"T", "U", "V", "W"}:
            if not is_valid_decimal_or_int(value):
                error_list.append(f"Column {source_letter} must be a number (int or decimal); got '{value}'")
        if source_letter == "Z":
            if not is_valid_country(value):
                error_list.append(f"Column Z must be a valid country name (first letter capital); got '{value}'")
        if source_letter == "AD":
            if str(value).strip() not in allowed_AD:
                error_list.append(f"Column AD must be one of {allowed_AD}; got '{value}'")
        if source_letter == "AI":
            if str(value).strip() not in allowed_AI:
                error_list.append(f"Column AI must be one of {allowed_AI}; got '{value}'")
        if source_letter == "AK":
            if str(value).strip() not in allowed_AK:
                error_list.append(f"Column AK must be one of {allowed_AK}; got '{value}'")
        if source_letter == "AL":
            if str(value).strip() not in allowed_AL:
                error_list.append(f"Column AL must be one of {allowed_AL}; got '{value}'")
        if source_letter == "AM":
            if str(value).strip() not in allowed_AM:
                error_list.append(f"Column AM must be one of {allowed_AM}; got '{value}'")
        if source_letter == "AN":
            if str(value).strip() not in allowed_AN:
                error_list.append(f"Column AN must be one of {allowed_AN}; got '{value}'")
        if source_letter == "AO":
            if not is_number(value):
                error_list.append(f"Column AO must be a number; got '{value}'")
        if source_letter in {"AP", "AR"}:
            if not is_valid_decimal_or_int(value):
                error_list.append(f"Column {source_letter} must be a number (int or decimal); got '{value}'")
        if source_letter in {"AQ", "AS"}:
            if not is_valid_unit(value):
                error_list.append(f"Column {source_letter} must be one of 'cm', 'mm', 'inch'; got '{value}'")
        if source_letter == "AT":
            if not is_valid_url(value):
                error_list.append(f"Column AT must be a valid URL; got '{value}'")
    
    # Cross-field validation: AF and AG must not have the same value.
    af_val = row_values.get("AF")
    ag_val = row_values.get("AG")
    if af_val is not None and ag_val is not None:
        if str(af_val).strip() == str(ag_val).strip():
            error_list.append("Columns AF and AG cannot have the same value")
    
    # If row is valid, write it to the template; otherwise, add it to invalid data report.
    if not error_list:
        # Write each mapped field into the template in the row indicated by valid_row_counter
        for target_col, (_, source_letter) in mapping.items():
            target_index = column_index_from_string(target_col)
            source_index = column_index_from_string(source_letter) - 1
            try:
                value = row[source_index]
            except IndexError:
                value = None
            sheet.cell(row=valid_row_counter, column=target_index).value = value
        valid_row_counter += 1  # Increment valid row counter
    else:
        # Create a dict for the row using Excel letters for columns (A=first column, etc.)
        row_dict = {chr(65 + j): row[j] for j in range(len(row))}
        row_dict["Validation Errors"] = ", ".join(error_list)
        invalid_data_rows.append(row_dict)

# --- Step 6: Save the updated workbook with only valid rows ---
output_path = "C_sling-bag_filled.xlsx"
wb.save(output_path)
print("✅ Valid rows have been filled into the template starting from row 5 in the sling_bag tab.")

# --- Step 7: Create a report for invalid rows (if any) ---
if invalid_data_rows:
    invalid_df = pd.DataFrame(invalid_data_rows)
    invalid_report_path = "invalid_data_report.xlsx"
    invalid_df.to_excel(invalid_report_path, index=False)
    print(f"⚠️ Invalid data report generated: {invalid_report_path}")
else:
    print("✅ All rows passed validation. No invalid data report generated.")
