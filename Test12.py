import streamlit as st
import pandas as pd
import io

st.header("Step 4: Merge Query & Scraper, Calculate Rebate")

# Rebate input
rebate_percent = st.number_input("Enter Rebate %", min_value=0.0, max_value=100.0, value=10.0, step=0.1)
rebate_decimal = rebate_percent / 100.0

# Upload files
query_file = st.file_uploader("Upload Query Results File", type=["xls", "xlsx"], key="query4")
scraper_file = st.file_uploader("Upload Scraper Results File", type=["xls", "xlsx"], key="scraper4")

if query_file and scraper_file:
    query_df = pd.read_excel(query_file)
    scraper_df = pd.read_excel(scraper_file)

    # Convert ROID to string
    query_df['ROID'] = query_df['ROID'].astype(str)
    scraper_df['ROID'] = scraper_df['ROID'].astype(str)

    # === Handle datetime split if available ===
    datetime_col = None
    for col in query_df.columns:
        if 'appointment' in col.lower() and 'datetime' in col.lower():
            datetime_col = col
            break

    if datetime_col:
        query_df[datetime_col] = pd.to_datetime(query_df[datetime_col], errors='coerce')
        query_df['Appointment date'] = query_df[datetime_col].dt.date
        query_df['Appointment month'] = query_df[datetime_col].dt.strftime('%B')
        query_df['Appointment year'] = query_df[datetime_col].dt.year
    else:
        st.warning("No 'appointment_datetime' column found to split into date/month/year.")

    # Merge Query & Scraper
    merged_df = pd.merge(query_df, scraper_df, on='ROID', how='left', suffixes=('', '_scraper'))

    # Calculate rebate and amount
    if 'Sub Total' in merged_df.columns and 'Payable Amount' in merged_df.columns:
        merged_df['Rebate%'] = rebate_percent
        merged_df['Rebate AI'] = merged_df['Sub Total'] * (-rebate_decimal)
        merged_df['Amount'] = merged_df['Payable Amount'] + merged_df['Rebate AI']
    else:
        st.warning("Missing 'Sub Total' or 'Payable Amount' in Scraper file.")

    # Desired final output columns
    final_cols = [
        'Appointment date', 'Appointment month', 'Appointment year',
        'Vendor Name', 'Shop Name', 'Country', 'ROID', 'Invoice no', 'VIN',
        'Sub Total', 'Tax Total', 'AI trans Fee ', 'FMC Rebate', 'Payable Amount',
        'Rebate AI', 'Rebate%', 'Amount', 'Trans fee', 'Merch fee',
        'Status in api', 'AP status'
    ]

    output_df = merged_df[[col for col in final_cols if col in merged_df.columns]]
    st.dataframe(output_df.head())

    # Download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Rebate_Results')
    output.seek(0)

    st.download_button(
        label="📥 Download Final Rebate Output",
        data=output,
        file_name="rebate_calculation_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )