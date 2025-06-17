import streamlit as st
import pandas as pd
import io

st.title("Statement & Estimates Matcher (Auto PO or ROID)")

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
    # Read Excel files
    statement_df = pd.read_excel(statement_file)
    estimates_df = pd.read_excel(estimates_file)

    # Filter required columns from Estimates
    estimates_df = estimates_df[required_cols]

    # Detect merge key
    if 'PO' in statement_df.columns and 'PO' in estimates_df.columns:
        merge_key = 'PO'
    elif 'ROID' in statement_df.columns and 'ROID' in estimates_df.columns:
        merge_key = 'ROID'
    else:
        st.error("❌ Neither PO nor ROID column found in both files. Please check your input files.")
        st.stop()

    st.info(f"🔗 Merging on: **{merge_key}**")

    # Perform LEFT JOIN to keep all Statement records
    merged_df = pd.merge(statement_df, estimates_df, on=merge_key, how="left", indicator=True)

    # Add new column for match status based on the _merge column
    merged_df["Match Status"] = merged_df["_merge"].map({
        "both": "Matched with Estimates",
        "left_only": "Unmatched with Estimates (N/A)",
    })

    # Drop the _merge column as it is no longer needed
    merged_df.drop(columns=["_merge"], inplace=True)

    # Disputed amount
    if 'Statement amount' in merged_df.columns and 'Amount to pay' in merged_df.columns:
        merged_df['Disputed amount'] = merged_df['Statement amount'] - merged_df['Amount to pay']
    else:
        st.warning("⚠️ 'Statement amount' or 'Amount to pay' column not found. Skipping Disputed amount calculation.")

    # Summary
    matched = (merged_df['Match Status'] == 'Matched with Estimates').sum()
    unmatched = (merged_df['Match Status'] == 'Unmatched with Estimates (N/A)').sum()
    dup_stmt = statement_df[merge_key].duplicated().sum()
    dup_est = estimates_df[merge_key].duplicated().sum()

    st.markdown("### 📊 Summary")
    st.write(f"✅ Matched with Estimates: {matched}")
    st.write(f"❌ Unmatched with Estimates: {unmatched}")
    st.write(f"🧾 Duplicate {merge_key}s in Statement: {dup_stmt}")
    st.write(f"📄 Duplicate {merge_key}s in Estimates: {dup_est}")

    # Save unmatched records
    unmatched_df = merged_df[merged_df['Match Status'] == 'Unmatched with Estimates (N/A)']

    # Download initial merged file
    output_initial = io.BytesIO()
    with pd.ExcelWriter(output_initial, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='Merged')

    output_initial.seek(0)

    st.success("✅ Initial merge completed!")

    st.download_button(
        label="📥 Download Initial Merged File",
        data=output_initial.getvalue(),
        file_name="Initial_Merged_Statement_Estimates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===========================
    # 🧩 Handle unmatched records
    # ===========================
    if unmatched > 0:
        st.markdown("---")
        st.markdown("### 🔍 Upload Query Result for Unmatched Records")
        query_result_file = st.file_uploader("📄 Upload Query Result File (.xlsx)", type=["xlsx"], key="query_result")

        if query_result_file:
            query_df = pd.read_excel(query_result_file)

            # Optionally standardize column names in Query Result for consistency
            if 'id' in query_df.columns:
                query_df.rename(columns={'id': 'PO'}, inplace=True)
            if 'ai_order_id' in query_df.columns:
                query_df.rename(columns={'ai_order_id': 'ROID'}, inplace=True)

            # Split 'Appointment datetime' into separate date, month, and year columns if exists
            if 'Appointment datetime' in query_df.columns:
                query_df['Appointment date'] = pd.to_datetime(query_df['Appointment datetime']).dt.date
                query_df['Appointment month'] = pd.to_datetime(query_df['Appointment datetime']).dt.month
                query_df['Appointment year'] = pd.to_datetime(query_df['Appointment datetime']).dt.year
                query_df.drop(columns=['Appointment datetime'], inplace=True)

            # Merge with unmatched records
            final_merged_unmatched = pd.merge(
                unmatched_df.drop(columns=[col for col in unmatched_df.columns if col in query_df.columns and col != merge_key]),
                query_df,
                how='left',
                on=merge_key,
            )

            # Match status for query result
            final_merged_unmatched['Match Status'] = final_merged_unmatched.apply(
                lambda row: 'Matched with Query Result' if pd.notna(row[merge_key]) and pd.notna(row.get('Appointment date'))
                else 'Unmatched with Estimates and Query Result',
                axis=1
            )

            # Final processed output
            final_output = pd.concat([
                merged_df[merged_df['Match Status'] == 'Matched with Estimates'],
                final_merged_unmatched
            ])

            # Final download
            output_final = io.BytesIO()
            with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
                final_output.to_excel(writer, index=False, sheet_name='Final Processed')

            output_final.seek(0)

            st.success("✅ Final processing completed with Query Results!")

            st.download_button(
                label="📥 Download Final Processed File",
                data=output_final.getvalue(),
                file_name="Final_Processed_Statement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )