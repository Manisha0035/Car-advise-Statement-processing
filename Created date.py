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
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"🚫 Could not read the file. Error: {e}")
        st.stop()

    required_cols = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']
    missing = [col for col in required_cols if col not in df.columns]
    
    if missing:
        st.error(f"❌ Missing required columns: {', '.join(missing)}")
    else:
        # Clean and convert currency columns
        for col in required_cols:
            df[col] = df[col].astype(str).str.replace(r'[\$,\u20b9,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df.dropna(subset=required_cols, inplace=True)

        # Step 1: Calculate Tax
        df['Tax'] = df['Total (inc. Tax)'] - df['SubTotal (exc. Tax)']

        # Step 2: Calculate Rebate and Rebate %
        rebate_rate = rebate_percent / 100.0
        df['Rebate'] = df['SubTotal (exc. Tax)'] * (-rebate_rate)
        df['Rebate %'] = ((df['Rebate'] / df['SubTotal (exc. Tax)']) * 100).round(2).astype(str) + '%'

        # Step 3: Calculate Amount to Pay
        df['Amount to Pay'] = df['Payable Amount (inc. Tax)'] + df['Rebate']

        # Step 4: Process Appointment Date or fallback to created_at
        df['Used created_at'] = 'No'
        if 'appointment_datetime' in df.columns or 'created_at' in df.columns:
            if 'appointment_datetime' in df.columns:
                df['appointment_datetime'] = pd.to_datetime(df['appointment_datetime'], errors='coerce')
            else:
                df['appointment_datetime'] = pd.NaT

            if 'created_at' in df.columns:
                df['created_at'] = pd.to_datetime(df['created_at'], errors='coerce')
                used_created_mask = df['appointment_datetime'].isna() & df['created_at'].notna()
                df.loc[used_created_mask, 'appointment_datetime'] = df.loc[used_created_mask, 'created_at']
                df.loc[used_created_mask, 'Used created_at'] = 'Yes'

            # Fill Appointment date, month, and year using created_at if appointment_datetime is NaT
            df['Appointment date'] = df['appointment_datetime'].dt.date
            df['Appointment month'] = df['appointment_datetime'].dt.strftime('%B')
            df['Appointment year'] = df['appointment_datetime'].dt.year

            # Fill in the missing Appointment date, month, and year with created_at
            created_at_mask = df['appointment_datetime'].isna() & df['created_at'].notna()
            df.loc[created_at_mask, 'Appointment date'] = df.loc[created_at_mask, 'created_at'].dt.date
            df.loc[created_at_mask, 'Appointment month'] = df.loc[created_at_mask, 'created_at'].dt.strftime('%B')
            df.loc[created_at_mask, 'Appointment year'] = df.loc[created_at_mask, 'created_at'].dt.year

        else:
            st.warning("ℹ️ Neither 'appointment_datetime' nor 'created_at' found. Skipping date breakdown.")

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
            'vin': 'VIN'
        }

        for old_col, new_col in column_renames.items():
            if old_col in df.columns:
                df[new_col] = df[old_col]

        # Insert blank 'AI trans Fee' and 'FMC Rebate' columns after 'Tax Total'
        def insert_blank_columns(df):
            if 'Tax Total' in df.columns and 'Payable Amount' in df.columns:
                cols = df.columns.tolist()
                insert_index = cols.index('Tax Total') + 1
                for new_col in ['AI trans Fee', 'FMC Rebate']:
                    if new_col not in cols:
                        cols.insert(insert_index, new_col)
                        df[new_col] = ""
                        insert_index += 1
                df = df[cols]
            return df

        df = insert_blank_columns(df)

        # Ensure all expected output columns are present
        expected_columns = [
            'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
            'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'AI trans Fee',
            'FMC Rebate', 'Payable Amount', 'Rebate AI', 'Rebate%', 'Amount to pay',
            'Trans fee', 'Merch fee', 'Status in api', 'AP status', 'Used created_at'
        ]
        for col in expected_columns:
            if col not in df.columns:
                df[col] = ""

        df_final = df[expected_columns]

        st.success("✅ Calculations complete!")
        st.dataframe(df_final)

        st.markdown(f"**🔢 Rows Processed:** {len(df_final)}")
        st.markdown(f"**📌 Rows using `created_at`:** {df_final['Used created_at'].value_counts().get('Yes', 0)}")
        st.markdown(f"**💸 Total Rebate:** ₹{df_final['Rebate AI'].astype(float).sum():,.2f}")
        st.markdown(f"**✅ Total Amount to Pay:** ₹{df_final['Amount to pay'].astype(float).sum():,.2f}")

        # Excel export with row highlighting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Updated')
            workbook = writer.book
            worksheet = writer.sheets['Updated']

            # Apply light yellow fill for rows where Used created_at == "Yes"
            highlight_format = workbook.add_format({'bg_color': '#FFFACD'})  # Light yellow

            for row_num, used_created in enumerate(df_final['Used created_at'], start=1):
                if used_created == 'Yes':
                    worksheet.set_row(row_num, cell_format=highlight_format)

        st.download_button(
            label="📅 Download Result Excel",
            data=output.getvalue(),
            file_name="updated_calculations.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )