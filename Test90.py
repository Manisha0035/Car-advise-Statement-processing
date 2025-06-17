import streamlit as st
import pandas as pd
import io

st.title("🧾 Statement & Estimates Matcher (PO Only)")

# Upload section
statement_file = st.file_uploader("📤 Upload Statement File (.xlsx)", type=["xlsx"])
estimates_file = st.file_uploader("📤 Upload Estimates File (.xlsx)", type=["xlsx"])

# Required columns from Estimates
required_cols = [
    'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
    'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'Payable Amount',
    'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
    'Status in api', 'AP status'
]

if statement_file and estimates_file:
    # Read Excel files
    statement_df = pd.read_excel(statement_file)
    estimates_df = pd.read_excel(estimates_file)

    # Filter required columns from Estimates
    estimates_df = estimates_df[required_cols]

    # Perform Left Join using PO only
    merged_df = pd.merge(
        statement_df, estimates_df, how='left',
        on='PO', suffixes=('', '_est')
    )

    # Add Match Status column
    merged_df['Match Status'] = merged_df['Appointment date'].apply(
        lambda x: 'Matched with Estimate' if pd.notna(x) else 'Unmatched with Estimate'
    )

    # Write to Excel buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='Merged')

    # Success message
    st.success("✅ Merge completed using PO with Match Status column added!")

    # Download button
    st.download_button(
        label="📥 Download Merged File",
        data=output.getvalue(),
        file_name="Merged_Statement_Estimates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )