import streamlit as st
import pandas as pd
import io  # For handling in-memory file downloads

# -------------------------------------
# 🏷️ Title
# -------------------------------------
st.title("📊 Statement & Estimate, Query File Matcher")

# -------------------------------------
# 📥 File Uploads
# -------------------------------------
statement_file = st.file_uploader("📤 Upload Statement File (Excel)", type=["xls", "xlsx"])
estimate_file = st.file_uploader("📤 Upload Estimate File (Excel)", type=["xls", "xlsx"])
qr_file = st.file_uploader("📤 Upload Query Results File (Excel, Optional)", type=["xls", "xlsx"])

# Required columns from Estimate file
required_columns = [
    "Appointment Date", "Appointment Month", "Appointment Year", "Vendor Name",
    "Country", "ROID", "PO", "Invoice", "VIN", "Sub Total", "Tax Total",
    "Payable Amount", "Rebate", "Amount to Pay", "Transfee", "Merchfee",
    "Status in API", "AP Status", "Record Source"
]

# -------------------------------------
# 🛠️ Processing Uploaded Files
# -------------------------------------
if statement_file and estimate_file:
    # Load the uploaded files
    statement_df = pd.read_excel(statement_file)
    estimate_df = pd.read_excel(estimate_file)

    # Determine the common key (PO or ROID)
    if "PO" in statement_df.columns and "PO" in estimate_df.columns:
        common_key = "PO"
    elif "ROID" in statement_df.columns and "ROID" in estimate_df.columns:
        common_key = "ROID"
    else:
        st.error("🚨 No common columns (PO or ROID) found in both files.")
        common_key = None

    if common_key:
        # Filter estimate_df to only required columns
        missing_cols = [col for col in required_columns if col not in estimate_df.columns]
        if missing_cols:
            st.warning(f"⚠️ The following required columns are missing in Estimate file: {missing_cols}")
            estimate_df = estimate_df[[col for col in required_columns if col in estimate_df.columns]]
        else:
            estimate_df = estimate_df[required_columns]

        # Perform LEFT JOIN to keep all Statement records
        merged_df = pd.merge(statement_df, estimate_df, on=common_key, how="left", indicator=True)

        # Add new column for match status
        merged_df["Match Status"] = merged_df["_merge"].map({
            "both": "Matched with Estimates",
            "left_only": "Unmatched with Estimates (N/A)"
        })

        # Drop the merge indicator column
        merged_df.drop(columns=["_merge"], inplace=True)

        # -------------------------------------
        # 📁 Process Query Results file if uploaded
        # -------------------------------------
        if qr_file:
            qr_df = pd.read_excel(qr_file)

            # Debugging: Show columns in QR file
            st.write("Query Results File Columns:", qr_df.columns.tolist())

            if common_key in qr_df.columns:
                # Merge unmatched records with Query Results
                merged_df = merged_df.merge(qr_df, on=common_key, how="left", suffixes=("", "_QR"))

                # Dynamically find columns from Query Results to fill
                qr_cols = [col for col in qr_df.columns if col != common_key]

                # Identify unmatched rows (Unmatched with Estimates)
                unmatched_mask = merged_df["Match Status"] == "Unmatched with Estimates (N/A)"

                # Update blank values (NaN) in unmatched rows with Query Results data
                for col in qr_cols:
                    if col in merged_df.columns:
                        merged_df.loc[unmatched_mask, col] = merged_df.loc[unmatched_mask, col].fillna(
                            merged_df.loc[unmatched_mask, f"{col}_QR"]
                        )
                        # Drop the "_QR" suffix column after filling
                        merged_df.drop(columns=[f"{col}_QR"], inplace=True, errors='ignore')

                # Update match status for rows where data was filled
                filled_mask = merged_df.loc[unmatched_mask].notna().any(axis=1)
                merged_df.loc[filled_mask.index, "Match Status"] = "Updated from Query Results ✅"

            else:
                st.error(f"🚨 '{common_key}' column not found in Query Results file.")

        # -------------------------------------
        # 🏷️ Reorder Columns for Better Readability
        # -------------------------------------
        # Move "Match Status" before "Case" column (if "Case" exists)
        if "Case" in merged_df.columns:
            cols = merged_df.columns.tolist()
            cols.insert(cols.index("Case"), cols.pop(cols.index("Match Status")))
            merged_df = merged_df[cols]

        # -------------------------------------
        # 📤 Display Results
        # -------------------------------------
        st.write("### ✅ Matched & Unmatched Records (After Filling)")
        st.dataframe(merged_df)

        # Provide a download option using BytesIO
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Final_Results")
        output.seek(0)

        st.download_button(
            label="📥 Download Final Results",
            data=output,
            file_name="final_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No matching records found.")
