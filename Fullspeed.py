import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF Table Extractor", layout="centered")
st.title("📄 Extract Tables from Specific PDF Pages")

# Upload PDF file
pdf_file = st.file_uploader("Upload your PDF", type=["pdf"])
start_page = st.number_input("Start Page", min_value=1, value=1)
end_page = st.number_input("End Page", min_value=1, value=1)

st.markdown("### 🛠️ Column Selection")
column_indices_str = st.text_input("Enter column indexes (comma-separated):", "0,1,2,3")
column_names_str = st.text_input("Enter column names (comma-separated):", "Invoice,Store Invoice,Service Date,Amount")

try:
    TARGET_COLUMNS = [int(i.strip()) for i in column_indices_str.split(",")]
    COLUMN_NAMES = [name.strip() for name in column_names_str.split(",")]
    
    if len(TARGET_COLUMNS) != len(COLUMN_NAMES):
        st.error("⚠️ Number of column indexes must match number of column names.")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Invalid input: {e}")
    st.stop()

if pdf_file and start_page <= end_page:
    all_tables = []
    no_table_pages = []

    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        end_page = min(end_page, total_pages)

        for i in range(start_page - 1, end_page):  # zero-indexed
            page = pdf.pages[i]
            tables = page.extract_tables()

            if not tables:
                no_table_pages.append(i + 1)
                continue

            for table in tables:
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    if df.shape[1] >= max(TARGET_COLUMNS) + 1:
                        try:
                            df = df.iloc[:, TARGET_COLUMNS]
                            df.columns = COLUMN_NAMES
                            all_tables.append(df)
                        except Exception as e:
                            st.warning(f"Error on page {i+1}: {e}")

    if all_tables:
        final_df = pd.concat(all_tables, ignore_index=True)
        st.success(f"✅ Extracted data from pages {start_page} to {end_page}")
        st.dataframe(final_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name="Extracted")
        output.seek(0)

        st.download_button(
            label="📥 Download Excel",
            data=output,
            file_name="extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No tables found in the selected pages.")

    if no_table_pages:
        st.info(f"No tables found on page(s): {', '.join(map(str, no_table_pages))}")
