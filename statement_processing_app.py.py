import streamlit as st
import pandas as pd

# Title of the app
st.title("Statement Processing Application")

# Step 1: Upload Statement File
st.header("Step 1: Upload Statement File")
statement_file = st.file_uploader("Upload Statement File (CSV)", type=["csv"])
if statement_file is not None:
    statement_df = pd.read_csv(statement_file)
    st.write("Statement Data:")
    st.dataframe(statement_df)
    
    # Filtering PO or ROID
    if 'ID' in statement_df.columns:
        statement_df['Type'] = statement_df['ID'].apply(lambda x: 'PO' if 'PO' in x else ('ROID' if 'ROID' in x else 'Other'))
        statement_df = statement_df[statement_df['Type'].isin(['PO', 'ROID'])]
        statement_df = statement_df[~statement_df['Type'].duplicated(keep='first')]
        st.write("Filtered Statement Data (PO preferred):")
        st.dataframe(statement_df)

# Step 2: Upload Estimates File
st.header("Step 2: Upload Estimates File")
estimates_file = st.file_uploader("Upload Estimates File (CSV)", type=["csv"], key="estimates")
if estimates_file is not None:
    estimates_df = pd.read_csv(estimates_file)
    st.write("Estimates Data:")
    st.dataframe(estimates_df)

    # Find matches in estimates
    if 'ID' in estimates_df.columns and 'ID' in statement_df.columns:
        matches = estimates_df[estimates_df['ID'].isin(statement_df['ID'])]
        st.write(f"Number of Matches in Estimates: {len(matches)}")

# Step 3: Upload Query Results and Scraper Results
st.header("Step 3: Upload Query and Scraper Results")
query_file = st.file_uploader("Upload Query Results File (CSV)", type=["csv"], key="query")
scraper_file = st.file_uploader("Upload Scraper Results File (CSV)", type=["csv"], key="scraper")
if query_file is not None and scraper_file is not None:
    query_df = pd.read_csv(query_file)
    scraper_df = pd.read_csv(scraper_file)
    
    st.write("Query Results Data:")
    st.dataframe(query_df)
    st.write("Scraper Results Data:")
    st.dataframe(scraper_df)

    unmatched = statement_df[~statement_df['ID'].isin(estimates_df['ID'])]
    st.write("Unmatched Statements:")
    st.dataframe(unmatched)

# Step 4: Final Step with Combined Results
st.header("Step 4: Final Data Preparation")
if query_file and scraper_file:
    # Assuming the user has a common ID between the different DataFrames
    # Here, you would combine the files based on the common ID and fill out required columns.

    # This part is placeholder logic, adjust as per your actual requirements.
    final_results = pd.merge(
        pd.merge(statement_df, estimates_df, on='ID', how='left'),
        pd.concat([query_df, scraper_df], ignore_index=True), 
        on='ID', 
        how='left'
    )

    # Display final results
    st.write("Final Results:")
    st.dataframe(final_results)
    
    # Fill required columns as needed
    # Placeholder example:
    final_results['Appointment Date'] = ''  # Placeholder for actual logic
    final_results['Vendor Name'] = ''  # Placeholder for actual logic
    # Add other required columns...

    st.write("Processed Final Result:")
    st.dataframe(final_results)