import pandas as pd

# Correct file path
file_path = r"C:\Users\Manisha\OneDrive\Documents\Jiffy lube test trial.xlsx"

# Load Statement and Estimate sheets
statement_df = pd.read_excel(file_path, sheet_name="Statement")  # Ensure sheet name is correct
estimate_df = pd.read_excel(file_path, sheet_name="Estimate")  # Ensure sheet name is correct

# Define the common ID column (modify if needed)
common_id = "PO"

# Perform a VLOOKUP-like merge
vlookup_result = statement_df[[common_id]].merge(
    estimate_df[[common_id]], 
    on=common_id, 
    how="left", 
    indicator=True
)

# Replace unmatched POs with "N/A"
vlookup_result["Vlookup from Estimates(PO)"] = vlookup_result.apply(
    lambda row: row[common_id] if row["_merge"] == "both" else "N/A", axis=1
)

# Keep only the renamed column
vlookup_result = vlookup_result[["Vlookup from Estimates(PO)"]]

# Save the result to a new Excel file
output_file = r"C:\Users\Manisha\OneDrive\Documents\matched_and_unmatched_pos.xlsx"
vlookup_result.to_excel(output_file, index=False)

# Display a preview
print(vlookup_result.head())

print(f"✅ Processing complete. Check the output file: {output_file}")
