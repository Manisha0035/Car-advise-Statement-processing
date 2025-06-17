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

# Target columns you want to extract (by index)
TARGET_COLUMNS = [0, 1, 2, 3]  # Adjust these if needed
COLUMN_NAMES = ["Invoice", "Store Invoice", "Service Date", "Amount"]

if pdf_file and start_page <= end_page:
    all_tables = []

    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        end_page = min(end_page, total_pages)

        for i in range(start_page - 1, end_page):  # zero-indexed
            page = pdf.pages[i]
            tables = page.extract_tables()

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
        st.success(f"✅ Extracted from pages {start_page} to {end_page}")
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