import streamlit as st
import pandas as pd
import io
import datetime
import os

st.set_page_config(page_title="Statement Processing", layout="wide")
st.title("📊 Statement Processor")

# ------------------------------------------
# 🔧 Helper Functions
# ------------------------------------------
def clean_currency(col):
    return pd.to_numeric(col.astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce').fillna(0)

def convert_dates(df, date_cols):
    for col in date_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

def download_excel(df, sheet_name, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    st.download_button(
        f"📥 Download {sheet_name}",
        output.getvalue(),
        filename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------------------------
# 📂 Tabs Setup
# ------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 Statement Processor", "Non-AI PO Check", "💳 Remittance Merge"])

# ------------------------------------------
# 📊 TAB 1 — Statement Processor
# ------------------------------------------
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

        if 'Statement amount' in merged_df.columns and 'Amount to pay' in merged_df.columns:
            merged_df['Disputed amount'] = merged_df['Statement amount'] - merged_df['Amount to pay']

        unmatched_df = merged_df[merged_df['Match Status'] == 'Unmatched with Estimates (N/A)']

        with st.expander("📄 Initial Merged File", expanded=False):
            st.dataframe(merged_df)
            merged_df = convert_dates(merged_df, ['Appointment date', 'Appointment month', 'Appointment year'])

            download_excel(merged_df, "Merged", "Initial_Merged_Statement_Estimates.xlsx")

        st.markdown("---")
        st.subheader("🧮 Step 2: Tax & Rebate Calculator for Enrichment")

        rebate_input_file = st.file_uploader("📁 Upload file for Tax & Rebate Calculation", type=["xlsx"], key="rebate_file")

        if rebate_input_file:
            rebate_percent = st.number_input("💸 Enter Rebate %", value=0.0, step=0.1, key="rebate_pct")
            df = pd.read_excel(rebate_input_file)

            required_cols_step2 = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']
            if not all(col in df.columns for col in required_cols_step2):
                st.error(f"❌ Required columns: {', '.join(required_cols_step2)}")
            else:
                for col in required_cols_step2:
                    df[col] = clean_currency(df[col])

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

                rebate_enrichment_df = df[[col for col in required_cols if col in df.columns]]
                st.success("✅ Calculations complete!")
                st.dataframe(rebate_enrichment_df)

                rebate_enrichment_df = convert_dates(rebate_enrichment_df, ['Appointment date', 'Appointment month', 'Appointment year'])

                download_excel(rebate_enrichment_df, "Updated", "updated_calculations.xlsx")

        # Step 3 - Enrich unmatched rows
        if unmatched_df.shape[0] > 0 and rebate_enrichment_df is not None:
            st.markdown("---")
            st.subheader("🔍 Step 3: Enrich Unmatched Rows Using Calculated File")

            drop_cols = [col for col in unmatched_df.columns if col in rebate_enrichment_df.columns and col != merge_key]
            enrich_df = pd.merge(
                unmatched_df.drop(columns=drop_cols),
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

            match_status_summary = final_output['Match Status'].value_counts()
            duplicate_statements = merged_df[merged_df['Match Status'] == 'Matched with Estimates'].duplicated(subset=merge_key, keep=False).sum()
            duplicate_estimates = estimates_df[estimates_df[merge_key].isin(final_output[merge_key])].duplicated(subset=merge_key, keep=False).sum()

            st.write("### Summary of Match Status")
            st.write(f"🔗 **Matched with Estimates**: {match_status_summary.get('Matched with Estimates', 0)}")
            st.write(f"🔗 **Matched with Query result**: {match_status_summary.get('Matched with Query result', 0)}")
            st.write(f"🔗 **Still Unmatched**: {match_status_summary.get('Still Unmatched', 0)}")
            st.write(f"🔁 **Duplicates in Statements**: {duplicate_statements}")
            st.write(f"🔁 **Duplicates in Estimates**: {duplicate_estimates}")

            # Final formatting
            if 'PO' in final_output.columns:
                try:
                    final_output['PO'] = pd.to_numeric(final_output['PO'], errors='raise').astype(int)
                except Exception as e:
                    st.warning(f"⚠️ Could not convert PO column to integer format: {e}")

            final_output = convert_dates(final_output, ['Appointment date', 'Appointment month', 'Appointment year'])

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            base_name = os.path.splitext(statement_file.name)[0] if statement_file else "Processed_Statement"
            file_name = f"{base_name}_Final_Processed_{timestamp}.xlsx"

            st.session_state["final_output_file_name"] = file_name
            st.session_state["final_output_df"] = final_output

            st.success("✅ Final enriched file ready!")
            download_excel(final_output, "Final Processed", file_name)

# ------------------------------------------
# 📥 TAB 2 — Non-AI PO Check
# ------------------------------------------
with tab2:
    st.header("🤚 Tab 2: PO Match Checker with Non-AI Reference")
    final_df = st.session_state.get("final_output_df")
    if final_df is None:
        st.warning("⚠️ Please run Tab 1 and complete processing to generate the final output file first.")
    else:
        non_ai_file = st.file_uploader("📄 Upload Non-AI Reference File (with 'PO' column)", type=["xlsx"], key="non_ai_file")

        if non_ai_file:
            try:
                df_non_ai = pd.read_excel(non_ai_file)

                if 'PO' not in df_non_ai.columns or 'PO' not in final_df.columns:
                    st.error("❌ 'PO' column not found in both files.")
                else:
                    df_non_ai['PO'] = df_non_ai['PO'].astype(str).str.strip()
                    final_df['PO'] = final_df['PO'].astype(str).str.strip()

                    final_df['Non AI check'] = final_df['PO'].apply(
                        lambda po: "Matched with Non-AI" if po in df_non_ai['PO'].values else " "
                    )

                    st.success("✅ 'Non AI check' column added.")
                    st.dataframe(final_df)

                    download_excel(final_df, "PO_Match_Result", st.session_state.get("final_output_file_name", "PO_Match_Result.xlsx"))
            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")

# ------------------------------------------
# 💳 TAB 3 — Remittance Merge
# ------------------------------------------
with tab3:
    st.header("💳 Tab 3: Remittance Merge Based on PO")
    final_df = st.session_state.get("final_output_df")
    if final_df is None:
        st.warning("⚠️ Please run Tab 1 and complete processing to generate the final output file first.")
    else:
        remittance_file = st.file_uploader("📤 Upload Remittance File (with shop_order_id, Paid Amounts, etc.)", type=["xlsx"], key="remittance_file")

        if remittance_file:
            try:
                remit_df = pd.read_excel(remittance_file)

                required_remit_cols = [
                    'aid_amount', 'paid_date',
                    'disputed_amount', 'disputed_date',
                    'dispute_reason', 'payment_details'
                ]

                if not all(col in remit_df.columns for col in required_remit_cols):
                    st.error(f"❌ Remittance file must include these columns:\n\n{', '.join(required_remit_cols)}")
                else:
                    remit_df['PO'] = remit_df['shop_order_id'].astype(str).str.strip()
                    final_df['PO'] = final_df['PO'].astype(str).str.strip()

                    merged_remit_df = pd.merge(
                        final_df,
                        remit_df.drop(columns=['shop_order_id']),
                        on='PO',
                        how='left'
                    )

                    for col in ['paid_date', 'disputed_date', 'paid_amount', 'disputed_amount', 'dispute_reason', 'payment_details']:
                        if col in merged_remit_df.columns:
                            merged_remit_df[col] = merged_remit_df[col].astype(str)

                    st.success("✅ Remittance columns successfully merged!")
                    st.dataframe(merged_remit_df)

                    file_name_remit = st.session_state.get("final_output_file_name", "Remittance_Merged.xlsx").replace(".xlsx", "_With_Remittance.xlsx")
                    download_excel(merged_remit_df, "Remittance_Merged", file_name_remit)
            except Exception as e:
                st.error(f"❌ Failed to process Remittance file: {e}")
