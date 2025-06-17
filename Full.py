import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Statement, Estimates & Query Matcher", page_icon="🧮")
st.title("🔄 Statement, Estimates & Query Matcher with Calculations")

# Step 1: Upload Statement and Estimates Files
st.subheader("Step 1: Statement & Estimates Matcher (Auto PO or ROID)")
statement_file = st.file_uploader("📄 Upload Statement File (.xlsx)", type=["xlsx"])
estimates_file = st.file_uploader("📄 Upload Estimates File (.xlsx)", type=["xlsx"])

# Required columns from Estimates
required_cols = [
    'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
    'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'Payable Amount',
    'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
    'Status in api', 'AP status'
]

# Step 2: Upload Query Result File (for unmatched rows)
st.subheader("Step 2: Upload Query Result File (.xlsx) for Unmatched Records")
qr_file = st.file_uploader("📤 Upload Query Result File (.xlsx)", type=["xlsx"])
rebate_percent_input = st.number_input("💰 Enter Vendor Rebate %", value=10.0, step=0.1)

if statement_file and estimates_file:
    # Read the Statement and Estimates files
    statement_df = pd.read_excel(statement_file)
    estimates_df = pd.read_excel(estimates_file)

    # Filter required columns from Estimates
    estimates_df = estimates_df[required_cols]

    # Detect merge key (PO or ROID)
    if 'PO' in statement_df.columns and 'PO' in estimates_df.columns:
        merge_key = 'PO'
    elif 'ROID' in statement_df.columns and 'ROID' in estimates_df.columns:
        merge_key = 'ROID'
    else:
        st.error("❌ Neither PO nor ROID column found in both files. Please check your input files.")
        st.stop()

    st.info(f"🔗 Merging on: **{merge_key}**")

    # Perform Left Join using selected key
    merged_df = pd.merge(statement_df, estimates_df, how='left', on=merge_key, suffixes=('', '_est'))

    # Add Match Status column
    merged_df['Match Status'] = merged_df['Appointment date'].apply(
        lambda x: 'Matched with Estimate' if pd.notna(x) else 'Unmatched with Estimate'
    )

    # Add Disputed Amount if both required columns exist
    if 'Statement amount' in merged_df.columns and 'Amount to pay' in merged_df.columns:
        merged_df['Disputed amount'] = merged_df['Statement amount'] - merged_df['Amount to pay']
    else:
        st.warning("⚠️ 'Statement amount' or 'Amount to pay' column not found in the merged data. Skipping Disputed amount calculation.")

    # Display summary
    matched = (merged_df['Match Status'] == 'Matched with Estimate').sum()
    unmatched = (merged_df['Match Status'] == 'Unmatched with Estimate').sum()
    st.markdown("### 📊 Matching Summary")
    st.write(f"✅ Matched with Estimates: {matched}")
    st.write(f"❌ Unmatched with Estimates: {unmatched}")

    # If QR file is uploaded, process the unmatched rows
    if qr_file:
        # Step 3: Process Query Result and enrich unmatched records
        qr_df = pd.read_excel(qr_file)

        # Normalize column names
        qr_df.columns = [col.strip().lower() for col in qr_df.columns]

        # Drop unused columns
        qr_df = qr_df.drop(columns=['total', 'ai_tax_total'], errors='ignore')

        # Rename columns to match the required format
        rename_map = {
            'id': 'PO',
            'ai_order_id': 'ROID',
            'invoice_number': 'Invoice no',
            'vin': 'VIN',
            'company': 'Vendor Name',
            'transaction_fee': 'Trans fee',
            'merch_fee': 'Merch fee',
            'status_in_api': 'Status in api',
            'ap_status': 'AP status',
            'appointment_datetime': 'appointment_datetime',
            'subtotal (exc. tax)': 'Sub Total',
            'total (inc. tax)': 'Tax Total',
            'payable amount (inc. tax)': 'Payable Amount'
        }
        qr_df = qr_df.rename(columns=rename_map)

        # Step 1: Calculate Tax
        if 'Total (inc. Tax)' in qr_df.columns and 'SubTotal (exc. Tax)' in qr_df.columns:
            qr_df['Tax'] = (qr_df['Total (inc. Tax)'] - qr_df['SubTotal (exc. Tax)']).round(2)

        # Step 2: Calculate Rebate and Rebate %
        rebate_rate = rebate_percent_input / 100.0
        if 'SubTotal (exc. Tax)' in qr_df.columns:
            qr_df['Rebate'] = (qr_df['SubTotal (exc. Tax)'] * -rebate_rate).round(2)
        qr_df['Rebate %'] = (qr_df['Rebate'] / qr_df['SubTotal (exc. Tax)']) * 100
        qr_df['Rebate %'] = qr_df['Rebate %'].astype(str) + '%'

        # Step 3: Calculate Amount to Pay
        if 'Payable Amount (inc. Tax)' in qr_df.columns:
            qr_df['Amount to Pay'] = (qr_df['Payable Amount (inc. Tax)'] + qr_df['Rebate']).round(2)

        # Date Parsing (if the column exists)
        if 'appointment_datetime' in qr_df.columns:
            qr_df['appointment_datetime'] = pd.to_datetime(qr_df['appointment_datetime'], errors='coerce')
            qr_df['Appointment date'] = qr_df['appointment_datetime'].dt.date
            qr_df['Appointment month'] = qr_df['appointment_datetime'].dt.strftime('%B')
            qr_df['Appointment year'] = qr_df['appointment_datetime'].dt.year
        else:
            st.warning("ℹ️ 'appointment_datetime' column not found, skipping date breakdown.")

        # Final column order
        final_cols = [
            'Appointment date', 'Appointment month', 'Appointment year',
            'Vendor Name', 'ROID', 'PO', 'Invoice no', 'VIN',
            'Sub Total', 'Tax Total', 'Payable Amount', 'Rebate', 'Rebate %',
            'Amount to pay', 'Trans fee', 'Merch fee', 'Status in api', 'AP status'
        ]
        qr_df = qr_df[[col for col in final_cols if col in qr_df.columns]]

        # Fill unmatched rows with data from Query Result
        unmatched_rows = merged_df[merged_df['Match Status'] == 'Unmatched with Estimate']
        unmatched_rows = pd.merge(unmatched_rows, qr_df, how='left', on=merge_key, suffixes=('', '_qr'))

        # Replace missing values in unmatched rows with Query Result values
        for col in ['Sub Total', 'Tax Total', 'Payable Amount', 'Rebate', 'Rebate %', 'Amount to pay']:
            if f'{col}_qr' in unmatched_rows.columns:  # Ensure the column exists before using it
                unmatched_rows[col] = unmatched_rows[col].combine_first(unmatched_rows[f'{col}_qr'])

        # Update the merged data with enriched unmatched rows
        merged_df.update(unmatched_rows)

        # Output the final merged file (with enriched unmatched rows)
        output_final = BytesIO()
        merged_df.to_excel(output_final, index=False, sheet_name='Final Merged Data')
        output_final.seek(0)

        st.success("✅ Query result enriched and unmatched rows filled.")

        # Download button for the final merged data
        st.download_button(
            label="📥 Download Final Merged File",
            data=output_final.getvalue(),
            file_name="final_merged_statement_estimates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )