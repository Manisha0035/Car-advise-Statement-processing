import streamlit as st
import pandas as pd
import io  # For handling in-memory file downloads

# Title
st.title("Statement & Estimate, Query File Matcher")

# Upload files
statement_file = st.file_uploader("Upload Statement File (Excel)", type=["xls", "xlsx"])
estimate_file = st.file_uploader("Upload Estimate File (Excel)", type=["xls", "xlsx"])
qr_file = st.file_uploader("Upload Query Results File (Excel, Optional)", type=["xls", "xlsx"])

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
        st.error("No common columns (PO or ROID) found in both files.")
        common_key = None

    if common_key:
        # Perform LEFT JOIN to keep all Statement records
        merged_df = pd.merge(statement_df, estimate_df, on=common_key, how="left", indicator=True)

        # Add new column for match status
        merged_df["Match Status"] = merged_df["_merge"].map({
            "both": "Matched with Estimates",
            "left_only": "Unmatched with Estimates (N/A)"
        })

        # Drop the merge indicator column
        merged_df.drop(columns=["_merge"], inplace=True)

        # === Optional Query Result merge ===
        if qr_file:
            qr_df = pd.read_excel(qr_file)

            if common_key in qr_df.columns:
                merged_df = merged_df.merge(qr_df, on=common_key, how="left", suffixes=("", "_QR"))

                # Dynamically detect one enriched column to determine query match
                qr_col_name = [col for col in merged_df.columns if "_QR" in col]
                if qr_col_name:
                    first_qr_col = qr_col_name[0]
                    merged_df["Match Status"] = merged_df.apply(
                        lambda row: "Matched with Query Results ✅"
                        if row["Match Status"] == "Unmatched with Estimates (N/A)" and not pd.isna(row[first_qr_col])
                        else row["Match Status"], 
                        axis=1
                    )
                else:
                    st.warning("No query result enrichment columns detected.")

            else:
                st.error(f"'{common_key}' column not found in Query Results file.")

        # Show Match Status Summary
        st.write("### Match Status Summary")
        summary_df = merged_df[[common_key, "Match Status"]]
        st.dataframe(summary_df)

        # Minimal Excel export (only CommonID and Match Status)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Match_Status")
        output.seek(0)

        st.download_button(
            label="📥 Download Match Status Only",
            data=output,
            file_name="match_status_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Unable to continue due to missing matching column.")