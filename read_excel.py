import pandas as pd

# Read Excel file
xls_path = "excels/XXX"

# Read only the relevant sheet
table1 = pd.read_excel(xls_path, sheet_name="AppBasicDet")

# Clean column names
table1.columns = table1.columns.str.strip()

# Extract A_ReferenceNo column and remove duplicates / NaNs
ref_nos = table1["A_ReferenceNo"].dropna().drop_duplicates().astype(int)

# Save to CSV with header 'ReferNo'
ref_nos.to_csv("reference_numbers.csv", index=False, header=["ReferNo"])

print("✅ CSV file saved as /content/A_ReferenceNos.csv")
