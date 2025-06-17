import streamlit as st
import pandas as pd
import io
import datetime
import os

st.set_page_config(page_title="Statement Matcher & Tax Calculator", layout="wide")
st.title("📊 Statement Matcher & 💰 Tax Calculator")

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

    unmatched_df = merged_df[merged_df['Match Status'] == 'Unmatched with Estimates (N/A)']

    with st.expander("📄 Initial Merged File", expanded=False):
        st.dataframe(merged_df)
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

    st.markdown("---")
    st.subheader("🧮 Step 2: Tax & Rebate Calculator for Enrichment")

    rebate_input_file = st.file_uploader("📁 Upload file for Tax & Rebate Calculation", type=["xlsx"], key="rebate_file")

    if rebate_input_file:
        rebate_percent = st.number_input("💸 Enter Rebate %", value=10.0, step=0.1, key="rebate_pct")
        df = pd.read_excel(rebate_input_file)
        required_cols_step2 = ['SubTotal (exc. Tax)', 'Total (inc. Tax)', 'Payable Amount (inc. Tax)']

        if not all(col in df.columns for col in required_cols_step2):
            st.error(f"❌ Required columns: {', '.join(required_cols_step2)}")
        else:
            for col in required_cols_step2:
                df[col] = df[col].astype(str).str.replace(r'[$,₹,]', '', regex=True)
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
                'FMC Rebate Amount':'FMC Rebate'
            }

            for old, new in column_renames.items():
                if old in df.columns:
                    df[new] = df[old] 

            desired_order = required_cols
            final_cols = [col for col in desired_order if col in df.columns]
            rebate_enrichment_df = df[final_cols]

            st.success("✅ Calculations complete!")
            st.dataframe(rebate_enrichment_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                rebate_enrichment_df.to_excel(writer, index=False, sheet_name='Updated')
            output.seek(0)

            st.download_button(
                label="📅 Download Tax & Rebate Result",
                data=output.getvalue(),
                file_name="updated_calculations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

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
        duplicate_statements = merged_df[merged_df['Match Status'] == 'Matched with Estimates'].duplicated(subset='PO', keep=False).sum()
        duplicate_estimates = estimates_df[estimates_df['PO'].isin(final_output['PO'])].duplicated(subset='PO', keep=False).sum()

        if not match_status_summary.empty:
            st.write("### Summary of Match Status")
            st.write(f"🔗 **Matched with Estimates**: {match_status_summary.get('Matched with Estimates', 0)}")
            st.write(f"🔗 **Matched with Query result**: {match_status_summary.get('Matched with Query result', 0)}")
            st.write(f"🔗 **Still Unmatched**: {match_status_summary.get('Still Unmatched', 0)}")
        st.write(f"🔁 **Duplicates in Statements**: {duplicate_statements}")
        st.write(f"🔁 **Duplicates in Estimates**: {duplicate_estimates}")

        output_final = io.BytesIO()
        with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
            final_output.to_excel(writer, index=False, sheet_name='Final Processed')
            workbook = writer.book
            worksheet = writer.sheets['Final Processed']

            # Highlight duplicates
            if 'PO' in final_output.columns:
                col_name = 'PO'
            elif 'ROID' in final_output.columns:
                col_name = 'ROID'
            else:
                col_name = None

            if col_name:
                col_idx = final_output.columns.get_loc(col_name)

                def get_excel_column_letter(idx):
                    letters = ''
                    while idx >= 0:
                        letters = chr(idx % 26 + 65) + letters
                        idx = idx // 26 - 1
                    return letters

                col_letter = get_excel_column_letter(col_idx)
                last_row = len(final_output) + 1
                cell_range = f'{col_letter}2:{col_letter}{last_row}'

                highlight_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                worksheet.conditional_format(cell_range, {
                    'type': 'duplicate',
                    'format': highlight_fmt
                })

        output_final.seek(0)

        original_name = os.path.splitext(statement_file.name)[0] if statement_file else "Processed_Statement"
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        file_name = f"{original_name}_Final_Processed_{timestamp}.xlsx"

        st.success("✅ Final enriched file ready!")
        st.download_button(
            "📥 Download Final Enriched Statement",
            output_final.getvalue(),
            file_name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )