import streamlit as st
import pandas as pd
from io import BytesIO

# Set Page Title
st.title("📊 Statement & Estimate Processor with Query Matching")

# File Upload Section
statement_file = st.file_uploader("📂 Upload Statement File (Excel)", type=["xlsx"])
estimate_file = st.file_uploader("📂 Upload Estimate File (Excel)", type=["xlsx"])

if statement_file and estimate_file:
    try:
        # Load Excel files safely
        statement_xls = pd.ExcelFile(statement_file)
        estimate_xls = pd.ExcelFile(estimate_file)

        # Check if required sheets exist
        required_sheets = ["Statement", "Estimate"]
        for sheet in required_sheets:
            if sheet not in statement_xls.sheet_names or sheet not in estimate_xls.sheet_names:
                st.error(f"❌ Missing required sheet: {sheet} in uploaded files.")
                st.stop()

        # Load sheets
        statement_df = pd.read_excel(statement_xls, sheet_name="Statement")
        estimate_df = pd.read_excel(estimate_xls, sheet_name="Estimate")

        # Ensure required columns exist
        required_cols = ["PO", "ROID"]
        missing_cols = [col for col in required_cols if col not in statement_df.columns]
        if missing_cols:
            st.error(f"❌ Missing required columns in Statement file: {missing_cols}")
            st.stop()

        # Create "Common_ID" dynamically
        statement_df["Common_ID"] = statement_df["PO"].combine_first(statement_df["ROID"])

        # Define dynamic column handling for Estimate file
        all_columns = list(estimate_df.columns)
        key_columns = ["PO", "ROID"]
        estimate_cols = [col for col in all_columns if col not in key_columns]  # Exclude duplicate key columns

        # Perform LEFT JOIN on "PO" first
        merged_df = pd.merge(statement_df, estimate_df, 
                             left_on="Common_ID", 
                             right_on="PO", 
                             how="left", suffixes=("_stmt", "_est"))

        # Handle missing matches by attempting ROID match
        no_match = merged_df["Appointment Date"].isna()
        if no_match.any():
            roid_merge = pd.merge(statement_df, estimate_df, 
                                  left_on="Common_ID", 
                                  right_on="ROID", 
                                  how="left", suffixes=("_stmt", "_est"))
            
            # Fill missing values with ROID match data
            merged_df.loc[no_match, estimate_cols] = roid_merge.loc[no_match, estimate_cols]

        # Add "Match Status" Column
        merged_df["Match Status"] = merged_df["Appointment Date"].apply(
            lambda x: "Matched ✅" if pd.notna(x) else "Not Matched ❌"
        )

        # Check for Unmatched Records
        unmatched_df = merged_df[merged_df["Match Status"] == "Not Matched ❌"]

        # Upload Query File if there are unmatched records
        if not unmatched_df.empty:
            st.warning(f"⚠️ {len(unmatched_df)} records unmatched. Upload **Query Result File** for further matching.")
            query_file = st.file_uploader("📂 Upload Query Result File", type=["xlsx"])

            if query_file:
                query_xls = pd.ExcelFile(query_file)

                # Check if "QueryResult" sheet exists
                if "QueryResult" not in query_xls.sheet_names:
                    st.error("❌ Missing 'QueryResult' sheet in uploaded file.")
                    st.stop()

                query_df = pd.read_excel(query_xls, sheet_name="QueryResult")

                # Ensure "Common_ID" exists in Query File
                if "Common_ID" in query_df.columns:
                    query_merge = pd.merge(unmatched_df, query_df, 
                                           on="Common_ID", 
                                           how="left", suffixes=("", "_QR"))

                    # Update unmatched rows with Query Results
                    merged_df.loc[merged_df["Match Status"] == "Not Matched ❌", estimate_cols] = query_merge[estimate_cols]

                    # Mark any still unmatched rows as "Manual Review"
                    unmatched_final = merged_df["Appointment Date"].isna()
                    merged_df.loc[unmatched_final, "Match Status"] = "Manual Review 🛠️"

        # Display Processed Data
        st.subheader("🔍 Processed Data Overview")
        st.dataframe(merged_df)

        # Function to convert DataFrame to Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Processed_Output")
            return output.getvalue()

        # Download Button for Final Output
        excel_data = to_excel(merged_df)
        st.download_button(
            label="📥 Download Final Processed File",
            data=excel_data,
            file_name="Final_Processed_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ An error occurred: {e}")
else:
    st.info("📌 Please upload both **Statement** and **Estimate** files to proceed.")
