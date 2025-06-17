import streamlit as st
import pandas as pd
import subprocess
import json
import io

# Helper to call external scraper script
def run_scraper_for_roid(roid):
    try:
        # Replace 'my_scraper.py' with your actual script
        result = subprocess.run(
            ['python', 'my_scraper.py', '--roid', str(roid)],
            capture_output=True, text=True
        )
        output = result.stdout.strip()
        return json.loads(output)  # Must be JSON-formatted
    except Exception as e:
        return {'ROID': roid, 'Status': 'Error', 'Error': str(e)}

st.title("Upload & Scrape by ROID")

uploaded_file = st.file_uploader("Upload Query Result Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write(df.head())

    if 'ai_order_id' not in df.columns:
        st.error("Missing 'ai_order_id' column in the file.")
    else:
        roids = df['ai_order_id'].dropna().unique()
        
        if st.button("Run Scraper for All ROIDs"):
            results = []
            progress = st.progress(0)
            status = st.empty()

            for i, roid in enumerate(roids):
                res = run_scraper_for_roid(roid)
                results.append(res)
                progress.progress((i + 1) / len(roids))
                status.text(f"Processing {i + 1} of {len(roids)}")

            results_df = pd.DataFrame(results)
            st.subheader("Scraper Results")
            st.write(results_df)

            # Download option
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                results_df.to_excel(writer, index=False)
            st.download_button("Download Excel", output.getvalue(), "scraper_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")