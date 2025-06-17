import streamlit as st
import pandas as pd
import io  # For handling in-memory file downloads

# Title
st.title("Statement & Estimate, Query File Matcher ")

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

        # Process Query Results file if uploaded
        if qr_file:
            qr_df = pd.read_excel(qr_file)

            # Debugging: Show columns in QR file
            st.write("Query Results File Columns:", qr_df.columns.tolist())

            if common_key in qr_df.columns:
                # Merge unmatched records with Query Results
                merged_df = merged_df.merge(qr_df, on=common_key, how="left", suffixes=("", "_QR"))

                # Debug: Show columns after merging
                st.write("Merged DataFrame Columns:", merged_df.columns.tolist())

                # Dynamically find the _QR column name
                qr_col_name = [col for col in merged_df.columns if "_QR" in col]
                if qr_col_name:
                    qr_col_name = qr_col_name[0]  # Pick the first matching column

                    # Update match status based on Query Results
                    merged_df["Match Status"] = merged_df.apply(
                        lambda row: "Matched with Query Results ✅"
                        if row["Match Status"] == "Unmatched with Estimates (N/A)" and not pd.isna(row[qr_col_name])
                        else row["Match Status"], 
                        axis=1
                    )
                else:
                    st.warning(f"No '_QR' column found after merging. Expected '{common_key}_QR'. Please check file structure.")

            else:
                st.error(f"'{common_key}' column not found in Query Results file.")

        # Move "Match Status" before "Case" column (if "Case" exists)
        if "Case" in merged_df.columns:
            cols = merged_df.columns.tolist()
            cols.insert(cols.index("Case"), cols.pop(cols.index("Match Status")))
            merged_df = merged_df[cols]

        # Display results
        st.write("### Matched & Unmatched Records")
        st.dataframe(merged_df)

        # Provide a download option using BytesIO
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Matched_Results")
        output.seek(0)

        st.download_button(
            label="📥 Download Matched Results",
            data=output,
            file_name="matched_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No matching records found.")