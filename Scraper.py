import pandas as pd
import streamlit as st
from data_collection_ai_amount import process_file  # Importing external function
import os

# Step 1: File upload using Streamlit
st.title("File Processor with Streamlit")
uploaded_file = st.file_uploader("Upload your Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Step 2: Read the uploaded file
    try:
        if uploaded_file.name.endswith((".xls", ".xlsx")):
            df = pd.read_excel(uploaded_file)
        else:
            st.error("Unsupported file format. Please upload an Excel file.")
            st.stop()
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # Step 3: Check for ROID column
    if 'ROID' not in df.columns:
        st.error("The uploaded file must contain a 'ROID' column.")
        st.stop()
    else:
        # Step 4: Convert ROID to text and handle missing values
        df['ROID'] = df['ROID'].fillna('Missing_ROID').astype(str)

        try:
            # Apply process_file function with proper input
            df['query_result_po'] = df['ROID'].apply(lambda x: process_file(uploaded_file, x) if isinstance(x, str) else 'Invalid_ROID')
            st.success("File processed successfully!")

            # Step 5: Display and download the results
            st.dataframe(df)
            output_file = "output_with_query_result_po.xlsx"

            # Ensure file is saved properly
            df.to_excel(output_file, index=False)

            # Safely open the file for download
            with open(output_file, 'rb') as file:
                st.download_button(
                    label="Download Processed File",
                    data=file.read(),
                    file_name=output_file,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        except Exception as e:
            st.error(f"Error processing data: {e}")