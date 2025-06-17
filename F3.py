import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Statement Matcher & Tax Calculator", layout="centered")
st.title("📊 Statement Matcher & 💰 Tax Calculator")

tab1, tab2 = st.tabs(["📋 Statement Matcher", "🧮 Tax & Rebate Calculator"])

# ============================
# 📋 TAB 1: Statement Matcher
# ============================
with tab1:
    st.header("Statement & Estimates Matcher (Auto PO or ROID)")

    # Upload section
    statement_file = st.file_uploader("📄 Upload Statement File (.xlsx)", type=["xlsx"])
    estimates_file = st.file_uploader("📄 Upload Estimates File (.xlsx)", type=["xlsx"])

    required_cols = [
        'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
        'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'AI trans Fee', 'FMC Rebate', 'Payable Amount',
        'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
        'Status in api', 'AP status'
    ]

    if statement_file and estimates_file:
        statement_df = pd.read_excel(statement_file)
        estimates_df = pd.read_excel(estimates_file)

        estimates_df = estimates_df[required_cols]

        if 'PO' in statement_df.columns and 'PO' in estimates_df.columns:
            merge_key = 'PO'
        elif 'ROID' in statement_df.columns and 'ROID' in estimates_df.columns:
            merge_key = 'ROID'
        else:
            st.error("❌ Neither PO nor ROID column found in both files.")
            st.stop()

        st.info(f"🔗 Merging on: **{merge_key}**")

        merged_df = pd.merge(statement_df, estimates_df, how='left', on=merge_key, indicator=True)
        merged_df['Match Status'] = merged_df['_merge'].map({
            'both': 'Matched with Estimates',
            'left_only': 'Unmatched with Estimates (N/A)'
        })
        merged_df.drop(columns=['_merge'], inplace=True)

        if 'Statement amount' in merged_df.columns and 'Amount to pay' in merged_df.columns:
            merged_df['Disputed amount'] = merged_df['Statement amount'] - merged_df['Amount to pay']
        else:
            st.warning("⚠️ Missing 'Statement amount' or 'Amount to pay' for dispute calc.")

        # Summary
        matched = (merged_df['Match Status'] == 'Matched with Estimates').sum()
        unmatched = (merged_df['Match Status'] == 'Unmatched with Estimates (N/A)').sum()
        dup_stmt = statement_df[merge_key].duplicated().sum()
        dup_est = estimates_df[merge_key].duplicated().sum()

        st.markdown("### 📊 Summary")
        st.write(f"✅ Matched with Estimates: {matched}")
        st.write(f"❌ Unmatched: {unmatched}")
        st.write(f"🧾 Duplicate {merge_key}s in Statement: {dup_stmt}")
        st.write(f"📄 Duplicate {merge_key}s in Estimates: {dup_est}")

        unmatched_df = merged_df[merged_df['Match Status'] == 'Unmatched with Estimates (N/A)']

        output_initial = io.BytesIO()
        with pd.ExcelWriter(output_initial, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Merged')
        output_initial.seek(0)

        st.download_button(
            "📥 Download Initial Merged File",
            output_initial.getvalue(),
            "Initial_Merged_Statement_Estimates.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if unmatched > 0:
            st.markdown("---")
            st.subheader("🔍 Upload Query Result for Enrichment")
            query_file = st.file_uploader("📄 Upload Query Result (.xlsx)", type=["xlsx"], key="query_file")

            if query_file:
                query_df = pd.read_excel(query_file)
                query_df.rename(columns={'id': 'PO', 'ai_order_id': 'ROID'}, inplace=True)

                if 'Appointment datetime' in query_df.columns:
                    query_df['Appointment date'] = pd.to_datetime(query_df['Appointment datetime']).dt.date
                    query_df['Appointment month'] = pd.to_datetime(query_df['Appointment datetime']).dt.month
                    query_df['Appointment year'] = pd.to_datetime(query_df['Appointment datetime']).dt.year
                    query_df.drop(columns=['Appointment datetime'], inplace=True)

                final_merged = pd.merge(
                    unmatched_df.drop(columns=[col for col in unmatched_df.columns if col in query_df.columns and col != merge_key]),
                    query_df,
                    on=merge_key,
                    how='left',
                    indicator=True
                )
                final_merged['Match Status'] = final_merged['_merge'].map({
                    'both': 'Matched with Query Result',
                    'left_only': 'Unmatched with Estimates and Query'
                })
                final_merged.drop(columns=['_merge'], inplace=True)

                final_output = pd.concat([
                    merged_df[merged_df['Match Status'] == 'Matched with Estimates'],
                    final_merged
                ])

                output_final = io.BytesIO()
                with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
                    final_output.to_excel(writer, index=False, sheet_name='Final Processed')

                output_final.seek(0)

                st.success("✅ Final output with query enrichment ready!")

                st.download_button(
                    "📥 Download Final Processed File",
                    output_final.getvalue(),
                    "Final_Processed_Statement.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


# ==============================
# 🧮 TAB 2: Tax & Rebate Calc
# ==============================
with tab2:
    st.header("📟 Tax, Rebate & Payment Calculator")

    uploaded_file = st.file_uploader("📄 Upload Excel File", type=["xlsx"], key="rebate_file")
    rebate_percent = st.number_input("💰 Enter Vendor Rebate %", value=10.0, step=0.1)

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        required_cols = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']

        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ Required columns: {', '.join(required_cols)}")
        else:
            for col in required_cols:
                df[col] = df[col].astype(str).str.replace(r'[\$,\u20b9,]', '', regex=True)
                df[col] = pd.to_numeric(df[col], errors='coerce')
            df.dropna(subset=required_cols, inplace=True)

            df['Tax'] = df['Total (inc. Tax)'] - df['SubTotal (exc. Tax)']
            rebate_rate = rebate_percent / 100.0
            df['Rebate'] = df['SubTotal (exc. Tax)'] * (-rebate_rate)
            df['Rebate %'] = ((df['Rebate'] / df['SubTotal (exc. Tax)']) * 100).round(2).astype(str) + '%'
            df['Amount to Pay'] = df['Payable Amount (inc. Tax)'] + df['Rebate']

            if 'appointment_datetime' in df.columns:
                df['appointment_datetime'] = pd.to_datetime(df['appointment_datetime'], errors='coerce')
                df['Appointment date'] = df['appointment_datetime'].dt.date
                df['Appointment month'] = df['appointment_datetime'].dt.strftime('%B')
                df['Appointment year'] = df['appointment_datetime'].dt.year
            else:
                st.warning("ℹ️ 'appointment_datetime' not found.")

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

            for old, new in column_renames.items():
                if old in df.columns:
                    df[new] = df[old]

            def insert_blank_columns(df):
                if 'Tax' in df.columns and 'Payable Amount' in df.columns:
                    cols = df.columns.tolist()
                    idx = cols.index('Tax') + 1
                    for new_col in ['AI trans fee', 'FMC Rebate']:
                        if new_col not in cols:
                            cols.insert(idx, new_col)
                            df[new_col] = ""
                            idx += 1
                    df = df[cols]
                return df

            df = insert_blank_columns(df)

            desired_order = [
                'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
                'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'AI trans fee', 'FMC Rebate', 'Payable Amount',
                'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
                'Status in api', 'AP status'
            ]

            final_cols = [col for col in desired_order if col in df.columns]
            df_final = df[final_cols]

            st.success("✅ Calculations complete!")
            st.dataframe(df_final)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Updated')

            output.seek(0)

            st.download_button(
                label="📅 Download Result Excel",
                data=output.getvalue(),
                file_name="updated_calculations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
