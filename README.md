This Python script reads data from a CSV file and fills an Excel template (specifically the sling_bag sheet) using a predefined mapping. Some columns are filled with data extracted from the CSV, while others are hard coded.

Overview
The script performs the following tasks:

Reads data from a CSV file using pandas.
Opens an Excel template using openpyxl.
Uses a mapping to fill specific columns in the template with either CSV data or hard-coded values.
Starts populating the template from row 5, leaving the first four rows (e.g., headers) unchanged.
Saves the modified template as a new Excel file.
Mapping Details
The mapping defines which template columns correspond to which CSV columns or hard-coded values. For example:

Column G: Filled with Seller SKU (from the CSV)
Column J: Filled with Maximum Retail Price (Sell on Amazon) (from the CSV)
Column L: Filled with the hard-coded value "seller"
Column AT: Filled with data from test(coded) (from the CSV)
Refer to the code comments for the full mapping details.

Prerequisites
Make sure you have Python 3 installed. The script requires the following Python packages:

pandas
openpyxl
You can install them via pip:

bash
Copy
Edit
pip install pandas openpyxl
Files
flipkart.py – The Python script containing the code.
Untitled spreadsheet - BL__Products__default_CSV_2025-02-12_18_43.csv – The CSV data file.
C_sling-bag_fd927b15e6244645_1703-2438FK_REQH2ILIQXHAH.xlsx – The Excel template file.
C_sling-bag_filled.xlsx – The output file created after the script is run.
How to Run
Place the CSV file and the Excel template file in the same directory as the Python script, or update the file paths in the script accordingly.

Open a terminal (or command prompt) in the directory containing these files.

Run the script with:

bash
Copy
Edit
python flipkart.py
After execution, the script will create a new Excel file named C_sling-bag_filled.xlsx with the data filled in the sling_bag sheet, starting from row 5.

Code Explanation
CSV Reading:
The script uses pandas to load data from the CSV file into a DataFrame.

Excel Template Loading:
It uses openpyxl to open the Excel template and select the sling_bag sheet.

Mapping and Data Insertion:
A mapping dictionary defines which template column (using letters) gets its value from a specific CSV column or is hard coded. The script then loops over each row in the CSV DataFrame and populates the corresponding cells in the template, starting from row 5.

Saving the Workbook:
Finally, the modified workbook is saved under a new filename.

Customization
Mapping:
If you need to change the mapping (i.e., which CSV column corresponds to which template column), modify the mapping dictionary in the script accordingly.

Start Row:
The script is currently set to start filling data from row 5. If you want to change this, update the start_row variable.

Troubleshooting
Permission Errors:
Ensure that the output file is not open in Excel when you run the script.
File Paths:
If the CSV or template file is in a different directory, update the file paths in the script accordingly.
This README should help you understand the script, its requirements, and how to run it. Adjust the details as needed for your specific use case.
