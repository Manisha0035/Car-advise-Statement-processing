import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Tax & Rebate Calculator", layout="centered")
st.title("📟 Tax, Rebate & Payment Calculator")

# Upload section
uploaded_file = st.file_uploader("📄 Upload Excel File", type=["xlsx"])

# Rebate % input
rebate_percent = st.number_input("💰 Enter Vendor Rebate %", value=10.0, step=0.1)

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    required_cols = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']
    if not all(col in df.columns for col in required_cols):
        st.error(f"❌ File must include columns: {', '.join(required_cols)}")
    else:
        # Clean and convert columns to numeric
        for col in required_cols:
            df[col] = df[col].astype(str).str.replace(r'[\$,\u20b9,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # Drop rows with invalid values
        df.dropna(subset=required_cols, inplace=True)

        # Step 1: Calculate Tax
        df['Tax'] = df['Total (inc. Tax)'] - df['SubTotal (exc. Tax)']

        # Step 2: Calculate Rebate and Rebate %
        rebate_rate = rebate_percent / 100.0
        df['Rebate'] = df['SubTotal (exc. Tax)'] * (-rebate_rate)
        df['Rebate %'] = ((df['Rebate'] / df['SubTotal (exc. Tax)']) * 100).round(2).astype(str) + '%'

        # Step 3: Calculate Amount to Pay
        df['Amount to Pay'] = df['Payable Amount (inc. Tax)'] + df['Rebate']

        # Step 4: Process Appointment Date
        if 'appointment_datetime' in df.columns:
            df['appointment_datetime'] = pd.to_datetime(df['appointment_datetime'], errors='coerce')
            df['Appointment date'] = df['appointment_datetime'].dt.date
            df['Appointment month'] = df['appointment_datetime'].dt.strftime('%B')
            df['Appointment year'] = df['appointment_datetime'].dt.year
        else:
            st.warning("ℹ️ 'appointment_datetime' column not found, skipping date breakdown.")

        # Column renaming
        column_renames = {
            'SubTotal (exc. Tax)': 'Sub Total',
            'Tax': 'Tax Total',
            'Payable Amount (inc. Tax)': 'Payable Amount',
            'Rebate': 'Rebate AI',
            'Rebate %': 'Rebate%',
            'Amount to Pay': 'Amount to pay',
            'company': 'Vendor Name',
            'transaction_fee': 'Trans fee',
            'merch_fee': 'Merch fee',
            'Status_in_api': 'Status in api',
            'ap_status': 'AP status',
            'ai_order_id': 'ROID',
            'id': 'PO',
            'invoice_number': 'Invoice no', 
            'vin':'VIN'
      
        }

        for old_col, new_col in column_renames.items():
            if old_col in df.columns:
                df[new_col] = df[old_col]

        # Insert blank 'AI trans Fee' and 'FMC Rebate' columns between 'Tax' and 'Payable Amount'
        def insert_blank_columns(df):
            if 'Tax' in df.columns and 'Payable Amount' in df.columns:
                cols = df.columns.tolist()
                insert_index = cols.index('Tax') + 1
                for new_col in ['AI trans fee', 'FMC Rebate']:
                    if new_col not in cols:
                        cols.insert(insert_index, new_col)
                        df[new_col] = ""
                        insert_index += 1
                df = df[cols]
            return df

        df = insert_blank_columns(df)

        # Define desired column order
        desired_order = [
            'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
            'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'AI trans Fee', 'FMC Rebate', 'Payable Amount',
            'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
            'Status in api', 'AP status'
        ]

        # Filter only available columns to avoid KeyError
        final_columns = [col for col in desired_order if col in df.columns]
        df_final = df[final_columns]

        st.success("✅ Calculations complete!")
        st.dataframe(df_final)

        # Prepare file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Updated')

        st.download_button(
            label="📅 Download Result Excel",
            data=output.getvalue(),
            file_name="updated_calculations.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )