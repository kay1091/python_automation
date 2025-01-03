import pandas as pd
import openpyxl

# Define file paths
SFID_FILE = "input/SFID_file.xlsx"
SFDC_DUMP_FILE = "input/SFDC_dump.xlsx"
TEMPLATE_FILE = "input/Weekly_Template.xlsx"
OUTPUT_FILE = "output/Updated_Template.xlsx"

# Step 1: Load SFID File and Normalize SFIDs
try:
    sfid_df = pd.read_excel(SFID_FILE)
    if sfid_df.empty:
        print("Error: SFID file is empty.")
        exit()
    if "SFID" not in sfid_df.columns:
        print(f"Error: Column 'SFID' not found in {SFID_FILE}. Please check the column header.")
        exit()
    sfids = [str(sfid).strip().upper() for sfid in sfid_df["SFID"]]
    print("Normalized SFIDs from Excel:", sfids)
except FileNotFoundError:
    print(f"Error: SFID file not found at {SFID_FILE}")
    exit()
except Exception as e:
    print(f"Error reading SFID file: {e}")
    exit()

# Step 2: Load SFDC Dump and Normalize Data
try:
    sfdc_dump_df = pd.read_excel(SFDC_DUMP_FILE)
    if sfdc_dump_df.empty:
        print("Error: SFDC dump file is empty.")
        exit()
    sfdc_dump_df["SFID"] = sfdc_dump_df["SFID"].astype(str).str.strip().str.upper()
    print("Normalized SFIDs from SFDC Dump:", sfdc_dump_df["SFID"].tolist())
except FileNotFoundError:
    print(f"Error: SFDC dump file not found at {SFDC_DUMP_FILE}")
    exit()
except Exception as e:
    print(f"Error reading SFDC dump file: {e}")
    exit()

# Step 3: Filter Data Based on SFIDs
filtered_df = sfdc_dump_df[sfdc_dump_df["SFID"].isin(sfids)]
print("Filtered Data:")
print(filtered_df)

# Step 4: Load Template Workbook
try:
    template_wb = openpyxl.load_workbook(TEMPLATE_FILE)
    print(f"Available sheets: {template_wb.sheetnames}")
    template_ws = template_wb["SFDC Data"]
except FileNotFoundError:
    print(f"Error: Template file not found at {TEMPLATE_FILE}")
    exit()
except KeyError:
    print(f"Error: Worksheet 'SFDC Data' not found in template file.")
    exit()
except Exception as e:
    print(f"Error loading template file: {e}")
    exit()

# Step 5: Write Filtered Data to Template
row_num = 2  # Start from the second row (assuming the first row is headers)
for index, row in filtered_df.iterrows():
    template_ws.cell(row=row_num, column=1, value=row["SFID"])
    template_ws.cell(row=row_num, column=2, value=row["Customer Name"])
    template_ws.cell(row=row_num, column=3, value=row["Opportunity Amount"])
    template_ws.cell(row=row_num, column=4, value=row["Stage"])
    template_ws.cell(row=row_num, column=5, value=row["Owner Name"])
    template_ws.cell(row=row_num, column=6, value=row["Region"])
    template_ws.cell(row=row_num, column=7, value=row["Created Date"])
    template_ws.cell(row=row_num, column=8, value=row["Last Modified Date"])
    row_num += 1

# Step 6: Save Updated Template
try:
    template_wb.save(OUTPUT_FILE)
    print(f"Updated template saved successfully to {OUTPUT_FILE}!")
except Exception as e:
    print(f"Error saving updated template: {e}")
