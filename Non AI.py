import streamlit as st
import pandas as pd
import io
import datetime
import os

st.set_page_config(page_title="Statement Matcher & Tax Calculator", layout="wide")
st.title("📊 Statement Matcher & 💰 Tax Calculator")

# Tabs
main_tab, non_ai_tab = st.tabs(["Main Matcher", "🧾 Non-AI Check"])

# ---------------------- Tab 1: Main Matcher ---------------------- #
with main_tab:
    st.header("📋 Step 1: Upload Statement & Estimates Files")

    statement_file = st.file_uploader("📄 Upload Statement File (.xlsx)", type=["xlsx"], key="statement")
    estimates_file = st.file_uploader("📄 Upload Estimates File (.xlsx)", type=["xlsx"], key="estimates")

    required_cols = [
        'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
        'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'AI trans Fee', 'FMC Rebate', 'Payable Amount',
        'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
        'Status in api', 'AP status'
    ]

    rebate_enrichment_df = None
    output_final = None

    if statement_file and estimates_file:
        statement_df = pd.read_excel(statement_file)
        estimates_df = pd.read_excel(estimates_file)

        merge_key = st.selectbox("🔑 Select Merge Key", options=["PO", "ROID"], index=0)

        if merge_key not in statement_df.columns or merge_key not in estimates_df.columns:
            st.error(f"❌ Selected merge key '{merge_key}' not found in both files.")
            st.stop()

        st.info(f"🔗 Merging on: **{merge_key}**")

        estimates_df = estimates_df[required_cols]
        merged_df = pd.merge(statement_df, estimates_df, how='left', on=merge_key, indicator=True)
        merged_df['Match Status'] = merged_df['_merge'].map({
            'both': 'Matched with Estimates',
            'left_only': 'Unmatched with Estimates (N/A)'
        })
        merged_df.drop(columns=['_merge'], inplace=True)

        if 'Statement amount' in merged_df.columns and 'Amount to pay' in merged_df.columns:
            merged_df['Disputed amount'] = merged_df['Statement amount'] - merged_df['Amount to pay']

        unmatched_df = merged_df[merged_df['Match Status'] == 'Unmatched with Estimates (N/A)']

        st.subheader("🧮 Step 2: Tax & Rebate Calculator for Enrichment")
        rebate_input_file = st.file_uploader("📁 Upload file for Tax & Rebate Calculation", type=["xlsx"], key="rebate_file")

        if rebate_input_file:
            rebate_percent = st.number_input("💸 Enter Rebate %", value=10.0, step=0.1, key="rebate_pct")
            df = pd.read_excel(rebate_input_file)

            df.columns = df.columns.str.strip()
            required_cols_step2 = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']

            if not all(col in df.columns for col in required_cols_step2):
                st.error(f"❌ Required columns: {', '.join(required_cols_step2)}")
            else:
                for col in required_cols_step2:
                    df[col] = df[col].astype(str).str.replace(r'[$,₹]', '', regex=True)
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

                rebate_rate = rebate_percent / 100.0
                df['Tax'] = df['Total (inc. Tax)'] - df['SubTotal (exc. Tax)']
                df['Rebate'] = df['SubTotal (exc. Tax)'] * (-rebate_rate)
                df['Rebate %'] = df.apply(
                    lambda row: f"{((row['Rebate'] / row['SubTotal (exc. Tax)']) * 100):.2f}%" 
                    if row['SubTotal (exc. Tax)'] != 0 else '0.00%', axis=1
                )
                df['Amount to Pay'] = df['Payable Amount (inc. Tax)'] + df['Rebate']

                if 'appointment_datetime' in df.columns:
                    df['appointment_datetime'] = pd.to_datetime(df['appointment_datetime'], errors='coerce')
                    df['Appointment date'] = df['appointment_datetime'].dt.date
                    df['Appointment month'] = df['appointment_datetime'].dt.strftime('%B')
                    df['Appointment year'] = df['appointment_datetime'].dt.year

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
                    'vin': 'VIN',
                    'AI Transaction Fee': 'AI trans Fee',
                    'FMC Rebate Amount': 'FMC Rebate'
                }

                for old, new in column_renames.items():
                    if old in df.columns:
                        df[new] = df[old]

                final_cols = [col for col in required_cols if col in df.columns]
                rebate_enrichment_df = df[final_cols]

                if unmatched_df.shape[0] > 0:
                    enrich_df = pd.merge(
                        unmatched_df.drop(columns=[col for col in unmatched_df.columns if col in rebate_enrichment_df.columns and col != merge_key]),
                        rebate_enrichment_df,
                        on=merge_key,
                        how='left',
                        indicator=True
                    )
                    enrich_df['Match Status'] = enrich_df['_merge'].map({
                        'both': 'Matched with Query result',
                        'left_only': 'Still Unmatched'
                    })
                    enrich_df.drop(columns=['_merge'], inplace=True)

                    if 'Statement amount' in enrich_df.columns and 'Amount to pay' in enrich_df.columns:
                        enrich_df['Disputed amount'] = enrich_df['Statement amount'] - enrich_df['Amount to pay']

                    final_output = pd.concat([merged_df[merged_df['Match Status'] == 'Matched with Estimates'], enrich_df])

                    if {'Disputed amount', 'Rebate AI'}.issubset(final_output.columns):
                        final_output['Dispute analysis'] = final_output['Rebate AI'] + final_output['Disputed amount']

                    if 'Dispute analysis' in final_output.columns and 'Match Status' in final_output.columns:
                        cols = final_output.columns.tolist()
                        cols.remove('Match Status')
                        idx = cols.index('Dispute analysis') + 1
                        cols.insert(idx, 'Match Status')
                        final_output = final_output[cols]

                    output_final = io.BytesIO()
                    with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
                        final_output.to_excel(writer, index=False, sheet_name='Final')

                    st.success("✅ Final enriched file ready!")
                    st.download_button("📥 Download Final Enriched File", data=output_final.getvalue(), file_name="Final_Statement.xlsx", key="final_file_download")

# ---------------------- Tab 2: Non-AI Check ---------------------- #
with non_ai_tab:
    st.header("📂 Non-AI Query Result Matcher")

    non_ai_file = st.file_uploader("📄 Upload Non-AI Query Result File (.xlsx)", type=["xlsx"], key="nonai")
    statement_file_2 = st.file_uploader("📄 Upload Statement File for Non-AI Check (.xlsx)", type=["xlsx"], key="stmt_nonai")

    if non_ai_file and statement_file_2:
        non_ai_df = pd.read_excel(non_ai_file)
        statement_df_2 = pd.read_excel(statement_file_2)

        if 'id' not in non_ai_df.columns or 'PO' not in statement_df_2.columns:
            st.error("❌ Required columns missing. Ensure Non-AI file has 'id' and Statement file has 'PO'.")
        else:
            non_ai_df.rename(columns={'id': 'PO'}, inplace=True)

            non_ai_merge = pd.merge(statement_df_2, non_ai_df[['PO']], on='PO', how='left', indicator=True)
            non_ai_merge['Non AI check'] = non_ai_merge['_merge'].map({
                'both': 'Matched with Non AI query result',
                'left_only': ''
            })
            non_ai_merge.drop(columns=['_merge'], inplace=True)

            st.dataframe(non_ai_merge)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                non_ai_merge.to_excel(writer, index=False, sheet_name='Non AI Check')
            buffer.seek(0)

            st.download_button("📥 Download Non-AI Check Result", data=buffer.getvalue(), file_name="Non_AI_Check.xlsx")
