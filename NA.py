import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="🆚 Non-AI PO Match Checker", layout="wide")

st.title("Non-AI PO Match Checker")

# File uploaders
non_ai_file = st.file_uploader("📄 Upload Non-AI Reference File (with PO column)", type=["xlsx"], key="non_ai_file")
compare_file = st.file_uploader("📄 Upload File to Compare (with PO column)", type=["xlsx"], key="compare_file")

if non_ai_file and compare_file:
    try:
        df_non_ai = pd.read_excel(non_ai_file)
        df_compare = pd.read_excel(compare_file)

        # Validate PO column
        if 'PO' not in df_non_ai.columns or 'PO' not in df_compare.columns:
            st.error("❌ 'PO' column not found in both files.")
        else:
            # Normalize PO values
            df_non_ai['PO'] = df_non_ai['PO'].astype(str).str.strip()
            df_compare['PO'] = df_compare['PO'].astype(str).str.strip()

            # Create Match Status column
            df_compare['Non AI check'] = df_compare['PO'].apply(
                lambda po: "Matched with Non-AI" if po in df_non_ai['PO'].values else " "
            )

            st.success("✅ Match Status has been added.")
            st.dataframe(df_compare)

            # Export to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_compare.to_excel(writer, index=False, sheet_name="Match_Result")
            output.seek(0)

            st.download_button(
                label="📥 Download PO Match Result",
                data=output,
                file_name="PO_Match_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
