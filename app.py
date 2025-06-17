import streamlit as st
import pandas as pd
from io import BytesIO # type: ignore
import xlwt

st.set_page_config(page_title="Statement Processor", layout="wide")
st.title("📄 Statement Processing Tool")

# === Sidebar: Rebate Input ===
st.sidebar.header("Automation Settings")
rebate_percent = st.sidebar.number_input("Rebate % (Automation)", min_value=0.0, max_value=100.0, value=10.0)
rebate_decimal = rebate_percent / 100.0

# === Helper: CommonID Logic ===
def get_common_id(row):
    return str(row['PO']) if pd.notna(row.get('PO')) and row['PO'] != '' else str(row.get('ROID', ''))

# === Step 1: Statement Upload ===
st.header("Step 1: Upload Statement File")
statement_file = st.file_uploader("Upload Statement File (.xls/.xlsx)", type=["xls", "xlsx"])
if statement_file:
    statement_df = pd.read_excel(statement_file)
    statement_df['CommonID'] = statement_df.apply(get_common_id, axis=1).astype(str)
    st.success("✅ Statement file processed.")
    st.dataframe(statement_df.head())

# === Step 2: Estimates Upload ===
st.header("Step 2: Upload Estimates File")
estimate_file = st.file_uploader("Upload Estimates File (.xls/.xlsx)", type=["xls", "xlsx"])
if estimate_file and statement_file:
    estimate_df = pd.read_excel(estimate_file)
    estimate_df['CommonID'] = estimate_df.apply(get_common_id, axis=1).astype(str)

    merged_df = pd.merge(
        statement_df,
        estimate_df,
        on='CommonID',
        how='left',
        suffixes=('_stmt', '_est')
    )

    # Add match status
    merged_df['Matched Status'] = merged_df['Appointment date ai'].apply(
        lambda x: 'Matched with Estimates' if pd.notna(x) else 'Unmatched with Estimates'
    )

    matched = merged_df[merged_df['Matched Status'] == 'Matched with Estimates']
    unmatched = merged_df[merged_df['Matched Status'] == 'Unmatched with Estimates']

    st.subheader("🟢 Matched Estimates")
    st.dataframe(matched)

    st.subheader("🔴 Unmatched (for Query + Scraper)")
    st.dataframe(unmatched[['CommonID', 'PO_stmt', 'ROID_stmt', 'Matched Status']])

    # Download after Step 2
    def convert_step2_df_to_xls(df):
        output = BytesIO()
        wb = xlwt.Workbook()
        ws = wb.add_sheet('MatchedStatus')

        for col_num, col in enumerate(df.columns):
            ws.write(0, col_num, col)

        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            for col_num, value in enumerate(row):
                ws.write(row_num, col_num, str(value))

        wb.save(output)
        output.seek(0)
        return output

    xls_step2 = convert_step2_df_to_xls(merged_df)

    st.download_button(
        label="📥 Download Statement with Match Status (.xls)",
        data=xls_step2,
        file_name="statement_with_match_status.xls",
        mime="application/vnd.ms-excel"
    )

# === Step 3: Query + Scraper Upload ===
st.header("Step 3: Upload Query + Scraper Files")
query_file = st.file_uploader("Upload Query File (.xls/.xlsx)", type=["xls", "xlsx"])
scraper_file = st.file_uploader("Upload Scraper File (.xls/.xlsx)", type=["xls", "xlsx"])

if query_file and scraper_file and estimate_file and statement_file:
    query_df = pd.read_excel(query_file)
    scraper_df = pd.read_excel(scraper_file)

    merged_qs = pd.merge(query_df, scraper_df, on='ROID', how='left', suffixes=('_query', '_scraper'))
    merged_qs['CommonID'] = merged_qs.apply(get_common_id, axis=1).astype(str)

    unmatched['CommonID'] = unmatched['CommonID'].astype(str)
    unmatched = unmatched.drop(columns=[col for col in unmatched.columns if '_est' in col], errors='ignore')
    enriched_df = pd.merge(unmatched, merged_qs, on='CommonID', how='left')

    if 'Sub Total' in enriched_df.columns and 'Payable Amount' in enriched_df.columns:
        enriched_df['Rebate%'] = rebate_percent
        enriched_df['Rebate AI'] = enriched_df['Sub Total'] * (-rebate_decimal)
        enriched_df['Amount to pay'] = enriched_df['Payable Amount'] + enriched_df['Rebate AI']

    st.success("✅ Unmatched enriched and calculations applied.")
    st.dataframe(enriched_df.head())

# === Step 4: Final Output ===
st.header("Step 4: Final Output")
if 'enriched_df' in locals() and not enriched_df.empty:
    final_columns = [
        'Appointment date', 'Appointment month', 'Appointment year', 'Vendor Name',
        'Shop Name', 'Country', 'ROID', 'PO', 'Invoice no', 'VIN',
        'Sub Total', 'Tax Total', 'AI trans Fee ', 'FMC Rebate', 'Payable Amount',
        'Rebate AI', 'Rebate%', 'Amount to pay', 'Trans fee', 'Merch fee',
        'Status in api', 'AP status'
    ]

    final_df = pd.concat([matched, enriched_df], ignore_index=True)
    output_df = final_df[[col for col in final_columns if col in final_df.columns]]
    st.subheader("✅ Final Processed Output")
    st.dataframe(output_df.head(20))

    # Write to XLS
    def convert_df_to_xls(df):
        output = BytesIO()
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Sheet1')

        for col_num, col in enumerate(df.columns):
            ws.write(0, col_num, col)

        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            for col_num, value in enumerate(row):
                ws.write(row_num, col_num, str(value))

        wb.save(output)
        output.seek(0)
        return output

    xls_data = convert_df_to_xls(output_df)

    st.download_button(
        label="📥 Download Final Output (.xls)",
        data=xls_data,
        file_name="final_output.xls",
        mime="application/vnd.ms-excel"
    )
else:
    st.info("⚠️ Please complete Step 3 to generate final output.")