import streamlit as st
import pandas as pd
import io

st.title("📟 Statement & Estimates Matcher (Auto PO or ROID)")

# Upload section
statement_file = st.file_uploader("📄 Upload Statement File (.xlsx)", type=["xlsx"])
estimates_file = st.file_uploader("📄 Upload Estimates File (.xlsx)", type=["xlsx"])

# Required columns from Estimates
required_cols = [
    'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
    'PO', 'ROID', 'Invoice no', 'VIN', 'Sub Total', 'Tax Total', 'Payable Amount',
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

    # Perform Left Join using selected key
    merged_df = pd.merge(
        statement_df, estimates_df, how='left',
        on=merge_key, suffixes=('', '_est')
    )

    # Add Match Status column
    merged_df['Match Status'] = merged_df['Appointment date'].apply(
        lambda x: 'Matched with Estimate' if pd.notna(x) else 'Unmatched with Estimate'
    )

    # ✅ Display summary on the Streamlit webpage
    matched = (merged_df['Match Status'] == 'Matched with Estimate').sum()
    unmatched = (merged_df['Match Status'] == 'Unmatched with Estimate').sum()
    dup_stmt = statement_df[merge_key].duplicated().sum()
    dup_est = estimates_df[merge_key].duplicated().sum()

    st.markdown("### 📊 Summary")
    st.write(f"✅ Matched with Estimates: {matched}")
    st.write(f"❌ Unmatched with Estimates: {unmatched}")
    st.write(f"🧾 Duplicate {merge_key}s in Statement: {dup_stmt}")
    st.write(f"📄 Duplicate {merge_key}s in Estimates: {dup_est}")

    # Write merged data to Excel buffer (single sheet only)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='Merged')

    output.seek(0)

    st.success("✅ Merge completed with Match Status column added!")

    # Download button
    st.download_button(
        label="📥 Download Merged File",
        data=output.getvalue(),
        file_name="Merged_Statement_Estimates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )