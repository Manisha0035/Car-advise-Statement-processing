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
        df = pd.read_csv(po_file) if po_file.name.endswith(".csv") else pd.read_excel(po_file)
        po_list = df.iloc[:, 0].dropna().astype(str).tolist()

    if not po_list:
        st.error("Please provide at least one PO number.")
        st.stop()

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 60)

    po_data_list = []
    ap_status_list = []

    try:
        driver.get("https://api.caradvise.com/admin/login")

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))).send_keys(EMAIL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))).send_keys(PASSWORD + Keys.RETURN)

        st.warning("⚠️ If CAPTCHA appears, solve it manually in the browser.")
        wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Shop Orders")))

        for po in po_list:
            st.info(f"🔍 Looking up PO: {po}")
            po_url = f"https://api.caradvise.com/admin/shop_orders/{po}"
            driver.get(po_url)
            time.sleep(3)

            # Extract main PO page table data
            soup = BeautifulSoup(driver.page_source, "html.parser")
            tables = soup.find_all("table")

            po_info = {
                "PO Number": po,
                "Id": "",
                "Ai Order Id": "",
                "Shop Name": "",
                "Company Name": "",
                "Vin": "",
                "Status": "",
                "Transaction Fee": "",
                "Appointment Datetime": "",
            }

            if tables:
                for table in tables:
                    for tr in table.find_all("tr"):
                        cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                        if len(cols) >= 2:
                            key, value = cols[0], cols[1]
                            if key in po_info:
                                po_info[key] = value
            else:
                st.warning(f"❌ No table data found for PO: {po}")

            po_data_list.append(po_info)

            # Extract AP Status from Edit Shop Order page
            ap_status = "Not found"
            try:
                edit_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Edit Shop Order")))
                edit_link.click()
                time.sleep(2)

                try:
                    ap_status_element = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "select[name='ap_status']"))
                    )
                    ap_status = ap_status_element.get_attribute("value")
                except:
                    ap_status_element = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='ap_status']"))
                    )
                    ap_status = ap_status_element.get_attribute("value")
            except Exception as e:
                st.warning(f"⚠️ Could not get AP Status for PO {po}: {e}")
            finally:
                driver.back()
                time.sleep(3)

            ap_status_list.append({"PO Number": po, "AP Status": ap_status})

        df_po = pd.DataFrame(po_data_list)
        df_ap = pd.DataFrame(ap_status_list)

        st.header("Results")

        st.subheader("PO Details")
        st.dataframe(df_po)

        st.subheader("AP Status")
        st.dataframe(df_ap)

        # Save both dataframes to one Excel file with two sheets
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            with pd.ExcelWriter(tmp.name, engine="xlsxwriter") as writer:
                df_po.to_excel(writer, sheet_name="PO Details", index=False)
                df_ap.to_excel(writer, sheet_name="AP Status", index=False)
            tmp.flush()
            st.download_button(
                label="📥 Download Combined Results (Excel)",
                data=open(tmp.name, "rb").read(),
                file_name="caradvise_po_lookup_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            os.unlink(tmp.name)

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        driver.quit()
