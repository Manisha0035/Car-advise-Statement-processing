import streamlit as st
import pandas as pd
import io
import datetime
import os

st.set_page_config(page_title="Statement Matcher & Tax Calculator", layout="wide")
st.title("📊 Statement Matcher & 💰 Tax Calculator")

# Define two tabs
tab1, tab2 = st.tabs(["📊 Statement Processor", "🆚 Non-AI PO Check"])

with tab1:
    st.header("📋 Step 1: Upload Statement & Estimates Files")

    statement_file = st.file_uploader("📄 Upload Statement File (.xlsx)", type=["xlsx"])
    estimates_file = st.file_uploader("📄 Upload Estimates File (.xlsx)", type=["xlsx"])

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

        merge_key = st.selectbox("🔑 Select merge key", ['PO', 'ROID'])

        if merge_key not in statement_df.columns or merge_key not in estimates_df.columns:
            st.error(f"❌ Selected key '{merge_key}' not found in both files.")
            st.stop()

        estimates_df = estimates_df[[col for col in required_cols if col in estimates_df.columns]]

        st.info(f"🔗 Merging on: **{merge_key}**")
        merged_df = pd.merge(statement_df, estimates_df, how='left', on=merge_key, indicator=True)
        merged_df['Match Status'] = merged_df['_merge'].map({
            'both': 'Matched with Estimates',
            'left_only': 'Unmatched with Estimates (N/A)'
        })
        merged_df.drop(columns=['_merge'], inplace=True)

        st.session_state["merged_df"] = merged_df.copy()

        if 'Statement amount' in merged_df.columns and 'Amount to pay' in merged_df.columns:
            merged_df['Disputed amount'] = merged_df['Statement amount'] - merged_df['Amount to pay']

        unmatched_df = merged_df[merged_df['Match Status'] == 'Unmatched with Estimates (N/A)']

        with st.expander("📄 Initial Merged File", expanded=False):
            st.dataframe(merged_df)
            output_initial = io.BytesIO()
            with pd.ExcelWriter(output_initial, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Merged')
            output_initial.seek(0)

            st.download_button(
                "📅 Download Initial Merged File",
                output_initial.getvalue(),
                "Initial_Merged_Statement_Estimates.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("---")
        st.subheader("🧲 Step 2: Tax & Rebate Calculator for Enrichment")

        rebate_input_file = st.file_uploader("📁 Upload file for Tax & Rebate Calculation", type=["xlsx"], key="rebate_file")

        if rebate_input_file:
            rebate_percent = st.number_input("💸 Enter Rebate %", value=10.0, step=0.1, key="rebate_pct")
            df = pd.read_excel(rebate_input_file)

            required_cols_step2 = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']

            if not all(col in df.columns for col in required_cols_step2):
                st.error(f"❌ Required columns: {', '.join(required_cols_step2)}")
            else:
                for col in required_cols_step2:
                    df[col] = df[col].astype(str).str.replace(r'[$,\u20b9,CA]', '', regex=True)
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

                st.success("✅ Calculations complete!")
                st.dataframe(rebate_enrichment_df)

                # Save final output and allow download
                st.session_state["final_output"] = rebate_enrichment_df

                output_final = io.BytesIO()
                with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
                    rebate_enrichment_df.to_excel(writer, index=False, sheet_name='Final Processed')
                output_final.seek(0)

                original_name = os.path.splitext(statement_file.name)[0] if statement_file else "Processed_Statement"
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                file_name = f"{original_name}_Final_Processed_{timestamp}.xlsx"

                st.download_button(
                    label="📥 Download Final Enriched File",
                    data=output_final,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

with tab2:
    st.header("🤚 Tab 2: PO Match Checker with Non-AI Reference")

    if "merged_df" not in st.session_state:
        st.warning("⚠️ Please run Tab 1 first to generate the merged output.")
    else:
        merged_df_tab2 = st.session_state["merged_df"]
        non_ai_file = st.file_uploader("📄 Upload Non-AI Reference File (with 'PO' column)", type=["xlsx"], key="non_ai_file")

        if non_ai_file:
            try:
                df_non_ai = pd.read_excel(non_ai_file)

                if 'PO' not in df_non_ai.columns or 'PO' not in merged_df_tab2.columns:
                    st.error("❌ 'PO' column not found in both files.")
                else:
                    df_non_ai['PO'] = df_non_ai['PO'].astype(str).str.strip()
                    merged_df_tab2['PO'] = merged_df_tab2['PO'].astype(str).str.strip()

                    merged_df_tab2['Non AI check'] = merged_df_tab2['PO'].apply(
                        lambda po: "Matched with Non-AI" if po in df_non_ai['PO'].values else " "
                    )

                    st.success("✅ 'Non AI check' column added.")
                    st.dataframe(merged_df_tab2)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        merged_df_tab2.to_excel(writer, index=False, sheet_name="PO_Match_Result")
                    output.seek(0)

                    st.download_button(
                        label="📅 Download PO Match Result",
                        data=output,
                        file_name="PO_Match_Result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")
