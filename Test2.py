import streamlit as st
import pandas as pd

# Streamlit App Title
st.title("📊 Statement and Estimate Matcher - Excel")

# File Upload Section
st.header("Upload Two Excel Files (.xls or .xlsx)")
statement_file = st.file_uploader("Upload Statement File (Columns: PO or ROID)", type=['xls', 'xlsx'])
estimate_file = st.file_uploader("Upload Estimate File (Columns: PO or ROID)", type=['xls', 'xlsx'])

# Function to find the correct column (PO or ROID) in each file
def get_matching_column(df):
    # Check if 'PO' exists in the column names
    if 'PO' in df.columns:
        return 'PO'
    elif 'ROID' in df.columns:
        return 'ROID'
    else:
        return None  # Return None if neither is found

# Function to Match Data Based on PO/ROID
def match_data(statement_df, estimate_df, statement_column, estimate_column):
    matched_records = []
    
    # Loop through each unique PO/ROID value in the statement file
    unique_ids = statement_df[statement_column].unique()
    for id in unique_ids:
        # Filter rows where the PO/ROID value matches in both files
        statement_row = statement_df[statement_df[statement_column] == id]
        estimate_row = estimate_df[estimate_df[estimate_column] == id]

        # If a match is found, merge relevant information
        if not estimate_row.empty:
            merged_row = {
                'PO_ROID': id,
                'Appointment_Date': statement_row['Appointment Date'].values[0] if 'Appointment Date' in statement_row.columns else 'N/A',
                'Appointment_Month': statement_row['Appointment Month'].values[0] if 'Appointment Month' in statement_row.columns else 'N/A',
                'Appointment_Year': statement_row['Appointment Year'].values[0] if 'Appointment Year' in statement_row.columns else 'N/A',
                'Vendor_Name': statement_row['Vendor Name'].values[0] if 'Vendor Name' in statement_row.columns else 'N/A',
                'Country': statement_row['Country'].values[0] if 'Country' in statement_row.columns else 'N/A',
                'ROID': statement_row['ROID'].values[0] if 'ROID' in statement_row.columns else 'N/A',
                'PO': statement_row['PO'].values[0] if 'PO' in statement_row.columns else 'N/A',
                'Invoice': statement_row['Invoice'].values[0] if 'Invoice' in statement_row.columns else 'N/A',
                'VIN': statement_row['VIN'].values[0] if 'VIN' in statement_row.columns else 'N/A',
                'Sub_Total': statement_row['Sub Total'].values[0] if 'Sub Total' in statement_row.columns else 'N/A',
                'Tax_Total': statement_row['Tax Total'].values[0] if 'Tax Total' in statement_row.columns else 'N/A',
                'Payable_Amount': statement_row['Payable Amount'].values[0] if 'Payable Amount' in statement_row.columns else 'N/A',
                'Rebate': statement_row['Rebate'].values[0] if 'Rebate' in statement_row.columns else 'N/A',
                'Amount_to_Pay': statement_row['Amount to Pay'].values[0] if 'Amount to Pay' in statement_row.columns else 'N/A',
                'Transfee': statement_row['Transfee'].values[0] if 'Transfee' in statement_row.columns else 'N/A',
                'Merchfee': statement_row['Merchfee'].values[0] if 'Merchfee' in statement_row.columns else 'N/A',
                'Status in API': estimate_row['Status in API'].values[0] if 'Status in API' in estimate_row.columns else 'N/A',
                'AP Status': estimate_row['AP Status'].values[0] if 'AP Status' in estimate_row.columns else 'N/A'
            }
            matched_records.append(merged_row)
    
    # Convert matched results to a DataFrame
    matched_df = pd.DataFrame(matched_records)
    return matched_df

# Process after both files are uploaded
if statement_file and estimate_file:
    try:
        # Read Excel files
        statement_df = pd.read_excel(statement_file)
        estimate_df = pd.read_excel(estimate_file)

        # Show input data (optional)
        st.subheader("Statement Data")
        st.dataframe(statement_df.head())

        st.subheader("Estimate Data")
        st.dataframe(estimate_df.head())

        # Identify the matching columns (PO or ROID)
        statement_column = get_matching_column(statement_df)
        estimate_column = get_matching_column(estimate_df)

        # Check if the required columns exist
        if statement_column is None or estimate_column is None:
            st.error("⚠️ Neither 'PO' nor 'ROID' found in one of the files. Please check the file contents.")
        else:
            # Show identified matching columns
            st.write(f"Matching on column: {statement_column}")

            # Perform matching
            st.subheader("🔍 Matching Results")
            matched_df = match_data(statement_df, estimate_df, statement_column, estimate_column)

            # Show matched data
            if not matched_df.empty:
                st.success("✅ Matching Completed!")
                st.dataframe(matched_df)

                # Convert DataFrame to Excel
                @st.cache_data
                def convert_df(df):
                    return df.to_excel(index=False, engine='openpyxl')

                # Create download button
                excel_data = convert_df(matched_df)
                st.download_button(
                    label="📥 Download Matched Data as Excel",
                    data=excel_data,
                    file_name='matched_records.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
            else:
                st.warning("⚠️ No matches found. Please check your files.")
            
    except Exception as e:
        st.error(f"Error reading files: {e}")

# Streamlit App Footer
st.write("Developed by [Your Name] - Data Analytics")