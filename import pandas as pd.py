import pandas as pd

# Load Excel file with multiple sheets
file_path = "C:\Users\Manisha\OneDrive\Documents"  # Update with your actual file path

# Load specific sheets
sheet1_df = pd.read_excel(file_path, sheet_name="Sheet1")  # First sheet
sheet2_df = pd.read_excel(file_path, sheet_name="Sheet2")  # Second sheet

# Define the common ID column (adjust as needed)
common_id = "PO"  # Replace with your actual column name

# Merge both sheets on the common ID
merged_df = sheet1_df.merge(sheet2_df, on=common_id, how="left")

# Save the merged data if needed
merged_df.to_excel("merged_output.xlsx", index=False)

# Display a preview of the result
print(merged_df.head())
