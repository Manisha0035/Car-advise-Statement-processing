import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import tempfile
import os
import time

st.set_page_config(page_title="CarAdvise PO Lookup", layout="wide")
st.title("🔍 CarAdvise - Shop Order Lookup")

# Login credentials
EMAIL = "appglide@caradvise.com"
PASSWORD = "CarAdvise_Admin"

# Input options
po_input = st.text_input("Enter PO Number (Shop Order ID)")
po_file = st.file_uploader("Or Upload Excel with PO Numbers", type=["xlsx", "csv"])

lookup_button = st.button("🔍 Lookup PO(s)")

if lookup_button:
    po_list = []

    # Collect POs
    if po_input:
        po_list.append(po_input.strip())
    elif po_file:
        df = pd.read_csv(po_file) if po_file.name.endswith(".csv") else pd.read_excel(po_file)
        po_list = df.iloc[:, 0].dropna().astype(str).tolist()

    if not po_list:
        st.error("Please provide at least one PO number.")
        st.stop()

    # Launch browser (not headless)
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 60)

    results = []

    try:
        # Open login page
        st.info("Opening login page...")
        driver.get("https://api.caradvise.com/admin/login")

        # Auto-fill login
        st.info("Entering email and password...")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))).send_keys(EMAIL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))).send_keys(PASSWORD + Keys.RETURN)

        st.warning("⚠️ If CAPTCHA appears, solve it manually in the browser.")
        st.info("⏳ Waiting for login to complete...")

        # Wait for Shop Orders link to confirm login
        try:
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Shop Orders")))
            st.success("✅ Login successful.")
        except:
            st.error("❌ Login failed or CAPTCHA not solved in time.")
            driver.quit()
            st.stop()

        # Go through each PO
        for po in po_list:
            st.info(f"🔍 Looking up PO: {po}")
            po_url = f"https://api.caradvise.com/admin/shop_orders/{po}"
            driver.get(po_url)
            time.sleep(5)

            soup = BeautifulSoup(driver.page_source, "html.parser")
            tables = soup.find_all("table")

            if not tables:
                st.warning(f"❌ No data found for PO: {po}")
                continue

            for table in tables:
                rows = []
                for tr in table.find_all("tr"):
                    cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                    if cols:
                        rows.append(cols)
                flat_data = {row[0]: row[1] if len(row) > 1 else "" for row in rows}
                flat_data["PO Number"] = po
                results.append(flat_data)

        # Display results
        if results:
            df = pd.DataFrame(results)
            st.success(f"✅ Found details for {len(df)} PO(s).")
            st.dataframe(df)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                df.to_excel(tmp.name, index=False)
                st.download_button("📥 Download Results", open(tmp.name, "rb"), file_name="po_lookup_results.xlsx")
                os.unlink(tmp.name)
        else:
            st.error("❌ No data found for any PO.")

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        driver.quit()
