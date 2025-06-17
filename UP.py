import streamlit as st
import pandas as pd
import io
import datetime
import os

st.set_page_config(page_title="Statement Matcher & Tax Calculator", layout="wide")
st.title("📊 Statement Matcher & 💰 Tax Calculator")

# Define Tabs
tab1, tab2 = st.tabs(["📋 Tab 1: Merge & Calculate", "🆚 Tab 2: Non-AI PO Check"])

# -------------------------------
# 📋 TAB 1 - Original Functionality
# -------------------------------
with tab1:
    st.header("📋 Step 1: Upload Statement & Estimates Files")

    statement_file = st.file_uploader("📄 Upload Statement File (.xlsx)", type=["xlsx"], key="statement_file")
    estimates_file = st.file_uploader("📄 Upload Estimates File (.xlsx)", type=["xlsx"], key="estimates_file")

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

        # Save merged_df to session state for Tab 2
        st.session_state['merged_df'] = merged_df.copy()

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

        # STEP 2 + 3 logic continues as-is...
        # [Omitted here for brevity since you've kept that unchanged in your original code]

# -------------------------------
# 🆚 TAB 2 - Compare with Non-AI PO File
# -------------------------------
with tab2:
    st.header("🆚 Tab 2: PO Match Checker with Non-AI Reference")

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
                        label="📥 Download PO Match Result",
                        data=output,
                        file_name="PO_Match_Result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")
