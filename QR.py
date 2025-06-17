import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Tax & Rebate Calculator", layout="centered")

st.title("🧾 Tax, Rebate & Payment Calculator")

# Upload section
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

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
            df[col] = df[col].astype(str).str.replace(r'[\$,₹,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # Drop rows with invalid values
        df.dropna(subset=required_cols, inplace=True)

        # Step 1: Calculate Tax
        df['Tax'] = df['Total (inc. Tax)'] - df['SubTotal (exc. Tax)']

        # Step 2: Calculate Rebate
        rebate_rate = rebate_percent / 100.0
        df['Rebate'] = df['SubTotal (exc. Tax)'] * (-rebate_rate)

        # Step 3: Calculate Amount to Pay
        df['Amount to Pay'] = df['Payable Amount (inc. Tax)'] + df['Rebate']

        st.success("✅ Calculations complete!")

        # Display updated DataFrame
        st.dataframe(df)

        # Prepare file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Updated')
        st.download_button(
            label="📥 Download Result Excel",
            data=output.getvalue(),
            file_name="updated_calculations.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
