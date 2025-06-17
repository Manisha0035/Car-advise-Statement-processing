import streamlit as st
import pandas as pd
import io  # For handling in-memory file downloads

# Title
st.title("📊 Statement & Estimate, Query File Matcher")

# Upload files
statement_file = st.file_uploader("📂 Upload Statement File (Excel)", type=["xls", "xlsx"])
estimate_file = st.file_uploader("📂 Upload Estimate File (Excel)", type=["xls", "xlsx"])
qr_file = st.file_uploader("📂 Upload Query Results File (Excel, Optional)", type=["xls", "xlsx"])

# Ensure both files are uploaded before proceeding
if statement_file is None or estimate_file is None:
    st.warning("⚠️ Please upload both Statement and Estimate files to continue.")
    st.stop()

# Load the uploaded files safely
try:
    statement_df = pd.read_excel(statement_file)
    estimate_df = pd.read_excel(estimate_file)
    st.success("✅ Files loaded successfully!")
except Exception as e:
    st.error(f"⚠️ Error reading files: {e}")
    st.stop()

# Identify the common key
required_keys = ["PO", "ROID"]
statement_cols = set(statement_df.columns)
estimate_cols = set(estimate_df.columns)

common_key = next((key for key in required_keys if key in statement_cols and key in estimate_cols), None)

if common_key is None:
    st.error("❌ No common key (PO or ROID) found in both files. Please check your file format.")
    st.stop()

# 🔹 Rename estimate_df columns (Avoid Duplicate Columns)
estimate_df = estimate_df.rename(columns={col: col + "_EST" for col in estimate_df.columns if col in statement_cols and col != common_key})

# 🔹 Merge without dropping columns
merged_df = pd.merge(statement_df, estimate_df, on=common_key, how="left", indicator=True)
merged_df["Match Status"] = merged_df["_merge"].map({
    "both": "Matched with Estimates ✅",
    "left_only": "Unmatched with Estimates ❌"
})
merged_df.drop(columns=["_merge"], inplace=True)

# 🔹 Process Query Results file if uploaded
if qr_file:
    try:
        qr_df = pd.read_excel(qr_file)
        st.success("✅ Query Results file uploaded!")

        if common_key in qr_df.columns:
            # Rename duplicate columns in qr_df
            qr_df = qr_df.rename(columns={col: col + "_QR" for col in qr_df.columns if col in merged_df.columns and col != common_key})

            merged_df = merged_df.merge(qr_df, on=common_key, how="left")

            qr_col_name = next((col for col in merged_df.columns if "_QR" in col), None)
            if qr_col_name:
                merged_df["Match Status"] = merged_df.apply(
                    lambda row: "Matched with Query Results ✅"
                    if row["Match Status"] == "Unmatched with Estimates ❌" and not pd.isna(row[qr_col_name])
                    else row["Match Status"], axis=1
                )
    except Exception as e:
        st.error(f"⚠️ Error loading Query Results file: {e}")

# 🔹 Drop all Statement columns except common key
columns_to_keep = [common_key] + [col for col in merged_df.columns if col.endswith("_EST") or col.endswith("_QR") or col == "Match Status"]
processed_df = merged_df[columns_to_keep]

# Display results
st.write("### Matched & Unmatched Records")
st.dataframe(processed_df)

# Provide a download option
if not processed_df.empty:
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            processed_df.to_excel(writer, sheet_name="Matched_Results", index=False)
            writer.book.close()  # Ensure data is written

        output.seek(0)  # Reset buffer position

        st.download_button(
            label="💽 Download Matched Results",
            data=output,
            file_name="matched_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"⚠️ Error saving Excel file: {e}")
else:
    st.warning("⚠️ No data available for download.")
