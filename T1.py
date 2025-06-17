import streamlit as st
import pandas as pd
import io

st.title("Statement & Estimate, Query File Matcher")

# Upload files
statement_file = st.file_uploader("Upload Statement File (Excel)", type=["xls", "xlsx"])
estimate_file = st.file_uploader("Upload Estimate File (Excel)", type=["xls", "xlsx"])
qr_file = st.file_uploader("Upload Query Results File (Excel, Optional)", type=["xls", "xlsx"])

# Final columns to export if matched with Estimates
EXPORT_COLUMNS = [
    "Appointment date", "Appointment month", "Appointment year", "Vendor Name", "ROID", "Invoice no", "VIN",
    "Sub Total", "Tax Total", "AI trans Fee", "FMC Rebate", "Payable Amount", "Rebate AI",
    "Amount to pay", "Trans fee", "Merch fee", "Status in api", "AP status"
]

if statement_file and estimate_file:
    statement_df = pd.read_excel(statement_file)
    estimate_df = pd.read_excel(estimate_file)

    if "PO" in statement_df.columns and "PO" in estimate_df.columns:
        common_key = "PO"
    elif "ROID" in statement_df.columns and "ROID" in estimate_df.columns:
        common_key = "ROID"
    else:
        st.error("No common columns (PO or ROID) found in both files.")
        common_key = None

    if common_key:
        duplicate_keys = estimate_df[estimate_df.duplicated(subset=[common_key], keep=False)][common_key].unique()
        estimate_unique = estimate_df[~estimate_df[common_key].isin(duplicate_keys)]
        merged_df = pd.merge(statement_df, estimate_unique, on=common_key, how="left", indicator=True)

        merged_df["Match Status"] = merged_df["_merge"].map({
            "both": "Matched with Estimates",
            "left_only": "Unmatched with Estimates (N/A)"
        })

        merged_df.loc[merged_df[common_key].isin(duplicate_keys), "Match Status"] = "Duplicating in Estimates"
        merged_df.drop(columns=["_merge"], inplace=True)

        if qr_file:
            qr_df = pd.read_excel(qr_file)
            if common_key in qr_df.columns:
                merged_df = merged_df.merge(qr_df, on=common_key, how="left", suffixes=("", "_QR"))
                qr_col_name = [col for col in merged_df.columns if "_QR" in col]
                if qr_col_name:
                    merged_df["Match Status"] = merged_df.apply(
                        lambda row: "Matched with Query Results ✅"
                        if row["Match Status"] == "Unmatched with Estimates (N/A)" and not pd.isna(row[qr_col_name[0]])
                        else row["Match Status"],
                        axis=1
                    )
                else:
                    st.warning("No '_QR' column found after merging. Please check file structure.")
            else:
                st.error(f"'{common_key}' column not found in Query Results file.")

        # Output: all records
        st.write("### Matched & Unmatched Records")
        st.dataframe(merged_df)

        # Summary
        st.markdown("#### 📊 Summary")
        st.write(f"✅ Matched with Estimates: {(merged_df['Match Status'] == 'Matched with Estimates').sum()}")
        st.write(f"🔍 Matched via Query Results: {(merged_df['Match Status'] == 'Matched with Query Results ✅').sum()}")
        st.write(f"⚠️ Duplicating in Estimates: {(merged_df['Match Status'] == 'Duplicating in Estimates').sum()}")
        st.write(f"❌ Unmatched: {(merged_df['Match Status'] == 'Unmatched with Estimates (N/A)').sum()}")

        # Filter Matched only and export required columns
        matched_export_df = merged_df[merged_df["Match Status"] == "Matched with Estimates"]

        # Check for column existence before filtering
        missing_cols = [col for col in EXPORT_COLUMNS if col not in matched_export_df.columns]
        if missing_cols:
            st.error(f"Missing columns in matched data: {missing_cols}")
        else:
            matched_export_df = matched_export_df[EXPORT_COLUMNS]

            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                matched_export_df.to_excel(writer, index=False, sheet_name="Matched_Only")
            output.seek(0)

            st.download_button(
                label="📥 Download Matched Records (Selected Columns)",
                data=output,
                file_name="matched_estimates_only.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
