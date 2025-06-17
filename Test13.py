import streamlit as st
import pandas as pd
import io

st.header("Step 4 - Merge Query Result with Scraper (LEFT JOIN)")

# Upload files
scraper_file = st.file_uploader("Upload Scraper Result File", type=["xls", "xlsx"])
query_file = st.file_uploader("Upload Query Result File", type=["xls", "xlsx"])

# Input Rebate %
rebate_percent = st.number_input("Enter Rebate %", min_value=0.0, max_value=100.0, value=5.0)
rebate_decimal = rebate_percent / 100

if scraper_file and query_file:
    # Read both files
    scraper_df = pd.read_excel(scraper_file)
    query_df = pd.read_excel(query_file)

    # Convert appointment datetime
    if 'appointment_date time' in query_df.columns:
        query_df['appointment_date time'] = pd.to_datetime(query_df['appointment_date time'], errors='coerce')
        query_df['Appointment date'] = query_df['appointment_date time'].dt.date
        query_df['Appointment month'] = query_df['appointment_date time'].dt.strftime('%B')
        query_df['Appointment year'] = query_df['appointment_date time'].dt.year

    # Rename columns for uniformity
    query_df.rename(columns={
        'company(vendor name)': 'Vendor Name',
        'invoice_number': 'Invoice no',
        'vin': 'VIN',
        'transaction_fee': 'Trans fee',
        'merch_fee': 'Merch fee',
        'Status_in_api': 'Status in api',
        'ap_status': 'AP status'
    }, inplace=True)

    # Perform LEFT JOIN: query_df LEFT JOIN scraper_df ON 'ROID'
    merged_df = pd.merge(query_df, scraper_df, on='ROID', how='left')

    # Calculate rebate and amount
    merged_df['Rebate%'] = rebate_percent
    merged_df['Rebate AI'] = merged_df['Sub Total'] * (-rebate_decimal)
    merged_df['Amount'] = merged_df['Payable Amount'] + merged_df['Rebate AI']

    # Final output columns
    final_cols = [
        'Appointment date', 'Appointment month', 'Appointment year',
        'Vendor Name', 'Shop Name', 'Country', 'ROID', 'Invoice no', 'VIN',
        'Sub Total', 'Tax Total', 'AI trans Fee ', 'FMC Rebate', 'Payable Amount',
        'Rebate AI', 'Rebate%', 'Amount', 'Trans fee', 'Merch fee',
        'Status in api', 'AP status'
    ]
    final_df = merged_df[[col for col in final_cols if col in merged_df.columns]]

    # Display preview
    st.write("### Final Merged Data with LEFT JOIN")
    st.dataframe(final_df)

    # Download option
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name="Rebate_Merged")
    output.seek(0)

    st.download_button(
        label="📥 Download Final Merged File",
        data=output,
        file_name="rebate_merged_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )