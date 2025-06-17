import streamlit as st
import pandas as pd
from io import BytesIO
import os

def clean_duplicate_columns(df):
    """Rename duplicate columns by appending a counter"""
    col_counts = {}
    new_columns = []
    
    for col in df.columns:
        if col in col_counts:
            col_counts[col] += 1
            new_columns.append(f"{col}_{col_counts[col]}")  # Rename duplicate
        else:
            col_counts[col] = 1
            new_columns.append(col)
    
    df.columns = new_columns
    return df

def format_excel(df):
    """Fix missing values, standardize text, and format dates"""
    df = clean_duplicate_columns(df)  # Handle duplicate columns

    # Convert column names to Title Case
    df.columns = [col.strip().title() for col in df.columns]

    # Fill missing 'Amount' values with 'Amount_2' if available
    if 'Amount' in df.columns and 'Amount_2' in df.columns:
        df['Amount'].fillna(df['Amount_2'], inplace=True)

    # Convert all string values to Proper Case
    df = df.applymap(lambda x: x.title() if isinstance(x, str) else x)

    # Convert date columns (if they contain 'Date' in the name)
    for col in df.columns:
        if "date" in col.lower():
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')

    # Clean 'Status' column (removing extra text)
    if 'Status' in df.columns:
        df['Status'] = df['Status'].str.extract(r'(Paid|Pending)', expand=False).fillna(df['Status'])

    # Remove duplicate rows
    df.drop_duplicates(inplace=True)

    return df

def save_to_excel(df, original_filename):
    """Save DataFrame to an Excel file in memory with a custom filename"""
    output = BytesIO()
    output_filename = f"{os.path.splitext(original_filename)[0]}_Output.xlsx"
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
    
    output.seek(0)
    return output, output_filename

# Streamlit UI
st.title("📊 Excel Formatter (Fix Duplicates & Missing Values)")
st.write("Upload an Excel file, and I'll clean duplicate columns, missing values, and format it!")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    st.write("### Original Data Preview")
    st.dataframe(df.head())

    # Process the Excel file
    formatted_df = format_excel(df)

    st.write("### Formatted Data Preview (Cleaned)")
    st.dataframe(formatted_df.head())

    # Save formatted file with the original filename + "_Output"
    formatted_excel, output_filename = save_to_excel(formatted_df, uploaded_file.name)

    st.download_button(
        label="📥 Download Formatted Excel",
        data=formatted_excel,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )