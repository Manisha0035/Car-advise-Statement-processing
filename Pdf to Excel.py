import streamlit as st
import pdfplumber
import pandas as pd
import io

st.title("📄 PDF Table to Excel Converter")
st.write("Extract tables from specific PDF pages and export them to Excel.")

# File uploader
pdf_file = st.file_uploader("Upload a PDF file", type="pdf")

# Page range input
page_range = st.text_input("Enter page range (e.g., 1-25):", value="1-5")

if pdf_file and page_range:
    try:
        # Parse start and end pages
        start_page, end_page = map(int, page_range.strip().split("-"))

        # Process PDF
        all_tables = []
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            end_page = min(end_page, total_pages)

            for i in range(start_page - 1, end_page):  # 0-based index
                page = pdf.pages[i]
                tables = page.extract_tables()

                for table in tables:
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])  # First row is header
                        df["Page"] = i + 1
                        all_tables.append(df)

        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)

            st.success(f"Extracted {len(all_tables)} tables from pages {start_page}-{end_page}.")

            st.dataframe(final_df)

            # Downloadable Excel output
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Extracted_Tables")
            output.seek(0)

            st.download_button(
                label="📥 Download Excel File",
                data=output,
                file_name="extracted_tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No tables found in the selected page range.")

    except Exception as e:
        st.error(f"Error: {e}")
