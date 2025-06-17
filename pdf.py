import streamlit as st
import pdfplumber
import pandas as pd
import io
import tempfile

# Streamlit App Title
st.set_page_config(page_title="PDF Table to Excel", layout="centered")
st.title("📄 PDF Table Extractor to Excel")
st.markdown("Upload a tabular PDF, define the page range, and download the result as Excel.")

# Upload PDF file
pdf_file = st.file_uploader("📤 Upload your PDF", type=["pdf"])

# Page range input
col1, col2 = st.columns(2)
with col1:
    start_page = st.number_input("Start Page", min_value=1, value=1)
with col2:
    end_page = st.number_input("End Page", min_value=1, value=1)

# Main extraction logic
if pdf_file and start_page <= end_page:
    if st.button("🚀 Extract and Convert"):
        all_tables = []

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_file.read())
            tmp_path = tmp.name

        with pdfplumber.open(tmp_path) as pdf:
            total_pages = len(pdf.pages)
            end_page = min(end_page, total_pages)

            for i in range(start_page - 1, end_page):
                page = pdf.pages[i]
                tables = page.extract_tables()
                st.info(f"🔍 Processing Page {i+1}...")

                if tables and any(tables):
                    for table in tables:
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0])  # First row = header
                            df["Page"] = i + 1
                            all_tables.append(df)
                else:
                    st.warning(f"⚠️ No tables found on page {i+1}.")

        # Combine and Export
        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            st.success("✅ Tables extracted! Preview below:")
            st.dataframe(final_df)

            # Save to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Extracted Tables")
            output.seek(0)

            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name="extracted_tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No tables found in the selected page range.")
