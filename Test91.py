import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Remap Query Result File", page_icon="🧹")
st.title("🧹 Remap Query Result File")

uploaded_file = st.file_uploader("📤 Upload Query Result File (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Standardize column names
    df.columns = [col.strip().lower() for col in df.columns]

    # Drop unwanted columns
    df = df.drop(columns=['total', 'ai_tax_total'], errors='ignore')

    # Rename columns
    rename_map = {
        'id': 'PO',
        'ai_order_id': 'ROID',
        'roid': 'ROID',
        'invoice_number': 'Invoice no',
        'vin': 'VIN',
        'company': 'Vendor Name',
        'transaction_fee': 'Trans fee',
        'merch_fee': 'Merch fee',
        'status_in_api': 'Status in api',
        'ap_status': 'AP status',
        'appointment_datetime': 'appointment_datetime'
    }
    df = df.rename(columns=rename_map)

    # Handle appointment datetime
    df['appointment_datetime'] = pd.to_datetime(df['appointment_datetime'], errors='coerce')
    df['Appointment date'] = df['appointment_datetime'].dt.date
    df['Appointment month'] = df['appointment_datetime'].dt.month
    df['Appointment year'] = df['appointment_datetime'].dt.year
    df = df.drop(columns=['appointment_datetime'], errors='ignore')

    # Reorder columns
    final_cols = [
        'Appointment date', 'Appointment month', 'Appointment year',
        'Vendor Name', 'ROID', 'PO', 'Invoice no', 'VIN',
        'Trans fee', 'Merch fee', 'Status in api', 'AP status'
    ]
    df = df[[col for col in final_cols if col in df.columns]]

    # Prepare for download
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    st.success("✅ File successfully remapped!")
    st.download_button(
        label="📥 Download Remapped Query Result",
        data=output,
        file_name="remapped_query_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )