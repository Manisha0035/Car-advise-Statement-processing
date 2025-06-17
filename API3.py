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

EMAIL = "appglide@caradvise.com"
PASSWORD = "CarAdvise_Admin"

po_input = st.text_input("Enter PO Number (Shop Order ID)")
po_file = st.file_uploader("Or Upload Excel with PO Numbers", type=["xlsx", "csv"])
lookup_button = st.button("🔍 Lookup PO(s)")

if lookup_button:
    po_list = []

    if po_input:
        po_list.append(po_input.strip())
    elif po_file:
        df_input = pd.read_csv(po_file) if po_file.name.endswith(".csv") else pd.read_excel(po_file)
        po_list = df_input.iloc[:, 0].dropna().astype(str).tolist()

    if not po_list:
        st.error("Please provide at least one PO number.")
        st.stop()

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 100)
    results = []

    try:
        st.info("Opening login page...")
        driver.get("https://api.caradvise.com/admin/login")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))).send_keys(EMAIL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))).send_keys(PASSWORD + Keys.RETURN)

        st.warning("⚠️ If CAPTCHA appears, solve it manually in the browser.")
        st.info("⏳ Waiting for login to complete...")
        try:
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Shop Orders")))
            st.success("✅ Login successful.")
        except:
            st.error("❌ Login failed or CAPTCHA not solved in time.")
            driver.quit()
            st.stop()

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

            flat_data = {}
            for table in tables:
                for tr in table.find_all("tr"):
                    cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                    if len(cols) >= 2:
                        key, value = cols[0], cols[1]
                        flat_data[key] = value

            flat_data["PO Number"] = po

            # Try to extract AP Status from dropdown
            try:
                ap_element = driver.find_element(By.XPATH, "//select[@name='status']")
                ap_status = ap_element.get_attribute("value")
                flat_data["AP Status"] = ap_status
            except:
                flat_data["AP Status"] = "Not Found"
                st.warning(f"⚠️ Could not get AP Status for PO {po}")

            results.append(flat_data)

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
