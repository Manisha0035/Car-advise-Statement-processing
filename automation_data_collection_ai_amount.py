# Import necessary libraries for data manipulation, automation, and web application interface
import pandas as pd  # For handling and processing data in a DataFrame format
import time  # For adding time delays
import random  # For generating random numbers, useful for adding random delays
import os  # For handling file operations
from selenium import webdriver  # For controlling a web browser
from selenium.webdriver.common.by import By  # For selecting elements in a web page by their attributes
from selenium.webdriver.edge.service import Service as EdgeService  # Service class for Edge browser
from selenium.webdriver.support.ui import WebDriverWait  # For adding waits to handle dynamic page loads
from selenium.webdriver.support import expected_conditions as EC  # For expected conditions in Selenium
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoSuchWindowException  # For handling exceptions in Selenium
from webdriver_manager.microsoft import EdgeChromiumDriverManager  # For managing Edge WebDriver installation
import streamlit as st  # For building web apps with Streamlit
from io import BytesIO  # For in-memory file handling
import base64  # For encoding files to base64 for downloading
from selenium.webdriver.edge.options import Options  # For setting Edge browser options
from datetime import datetime  # For handling date and time operations
from openpyxl import load_workbook  # For handling Excel files
from concurrent.futures import ThreadPoolExecutor
import concurrent.futures


# 1. Data Processing Function - Updated to Return DataFrame and ai_order_id Numbers
def process_file(file_path):
    # Specify data type for 'ai_order_id' column as string for consistency
    dtype_spec = {'ai_order_id': str}
    # Read the Excel file into a DataFrame with specified data types
    df = pd.read_excel(file_path, dtype=dtype_spec)
    # Ensure 'ai_order_id' column is treated as a string for consistent formatting
    df['ai_order_id'] = df['ai_order_id'].astype(str)
    # Use the full dataset and extract the ai_order_ids (no filtering)
    ai_order_ids = df['ai_order_id'].tolist()
    # Return the entire DataFrame and the list of ai_order_ids twice (if needed downstream)
    return df, ai_order_ids, df

# Global variable to hold the filename across the process
filename = None

# Function to generate a filename only once during each process run
def get_filename():
    global filename  # Reference the global filename variable
    if filename is None:  # Check if filename has not been generated yet
        # Create a unique filename with timestamp to avoid overwriting
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'scrape_results_{timestamp}.xlsx'  # Set the filename
    return filename  # Return the filename

# Function to ask user if they want to split the data for parallel scraping
def ask_user_to_split_data():
    # Streamlit checkbox for user to choose split option
    return st.checkbox("Split data into two parts for parallel scraping")

def scrape_authorization_numbers_multi_account(roid_numbers, df, accounts, save_interval=2, split_data=False):
    # Set up options to maximize browser window
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    drivers = []  # Initialize list to hold browser instances

    if len(accounts) == 1:
        # Single account: use one browser instance
        driver1 = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
        login(driver1, accounts[0][0], accounts[0][1])  # Login to first account
        drivers.append(driver1)  # Add driver to list
    else:
        # Multiple accounts: create browser instances for each account
        for account in accounts:
            driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
            login(driver, account[0], account[1])  # Login with each account
            drivers.append(driver)  # Add driver to list
    try:
        # Split the data into two parts if the user selected that option
        if split_data:
            split_index = len(roid_numbers) // 2
            part1 = roid_numbers[:split_index]
            part2 = roid_numbers[split_index:]
        else:
            part1 = roid_numbers
            part2 = []

        # Process first part of the data using the first two accounts
        for driver in drivers[:2]:  # Use the first two accounts
            for index, roid_number in enumerate(part1, start=1):
                # Process ai_order_id data with each account (using driver)
                driver_to_use = driver
                process_roid_data(roid_number, driver_to_use, df, index, len(part1))
        
        # Process second part of the data using the second set of accounts
        for driver in drivers[2:]:  # Use the second set of accounts
            for index, roid_number in enumerate(part2, start=1):
                # Process ai_order_id data with each account (using driver)
                driver_to_use = driver
                process_roid_data(roid_number, driver_to_use, df, index, len(part2))

    finally:
        for driver in drivers:
            driver.quit()  # Close all driver instances at the end of scraping
        save_partial_data(df)  # Final save to ensure data is not lost

    return df  # Return the final DataFrame with scraped data

# Helper function for processing individual ai_order_id data
def process_roid_data(roid_number, driver_to_use, df, index, total, save_interval=2):
    try:
        print(f"Processing {index}/{total}: ai_order_id {roid_number} using driver {index % len(driver_to_use) + 1}")
        # Construct URL for each ai_order_id
        repair_order_url = f'https://online.autointegrate.com/EditRepairOrder?jsId={roid_number}&jsro=false&roep=1'
        driver_to_use.get(repair_order_url)  # Open the page for current ai_order_id
        time.sleep(random.randint(1, 3))  # Random sleep for throttling

        # Scrape values (similar to the earlier approach)
        scrape_values(driver_to_use, roid_number, df)

        # Save data periodically based on save_interval
        if index % save_interval == 0:
            save_partial_data(df)  # Save current progress

    except (NoSuchElementException, TimeoutException, Exception) as e:
        print(f"Skipped {roid_number} due to error: {e}. Moving to next ai_order_id.")

# Function to scrape values (such as authorization number, repair order, etc.)
def scrape_values(driver_to_use, roid_number, df):
    try:
        # Scraping for authorization number, repair order, totals, etc.
        authorization_number_element = WebDriverWait(driver_to_use, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(., 'Authorization Number:')]"))
        )
        authorization_number = authorization_number_element.text.split("Authorization Number: ")[-1].strip()
        df.loc[df['ai_order_id'] == roid_number, 'Authorization Number'] = authorization_number
        print(f"Found authorization number: {authorization_number}")
        
        # Additional scraping logic for Repair Order ID, SubTotal (exc. Tax), Total (inc. Tax), etc.
        # Implement this similar to the earlier code you provided

    except Exception as e:
        print(f"Error during scraping values: {e}")
        df.loc[df['ai_order_id'] == roid_number, 'Authorization Number'] = ''


# Function to save partial data during processing
def save_partial_data(df):
    output_filename = get_filename()  # Retrieve or generate the filename

    # Check if file already exists
    if not os.path.exists(output_filename):
        # If the file doesn't exist, create a new one and save data
        df.to_excel(output_filename, index=False)
        print(f"Created new file and saved data to {output_filename}")
    else:
        # If the file exists, append data to it
        with pd.ExcelWriter(output_filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # Append data to the next available row without overwriting
            df.to_excel(writer, index=False, startrow=writer.sheets['Sheet1'].max_row, header=False)
        print(f"Appended data to existing file {output_filename}")

    return output_filename  # Return the filename

# 2. Function to automate login with provided credentials
def login(driver, username, password):
    driver.get('https://id.autointegrate.com/Account/Login')  # Open the login page

    # Wait until the page loads completely
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    try:
        # Wait and locate the username field, then enter the username
        username_field = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.ID, 'Username'))  # Updated based on HTML ID
        )
        username_field.clear()
        username_field.send_keys(username)  # Fill in the username

        # Locate and fill in the password field
        password_field = driver.find_element(By.ID, 'Password')  # Updated based on HTML ID
        password_field.clear()
        password_field.send_keys(password)

        # Locate and check the "Keep me logged in" checkbox if not selected
        remember_me_checkbox = driver.find_element(By.ID, 'RememberLogin')
        if not remember_me_checkbox.is_selected():
            remember_me_checkbox.click()  # Click to select if unchecked

        # Locate and click the login button
        login_button = driver.find_element(By.ID, 'loginButton')  # Updated based on HTML ID
        login_button.click()
        print(f"Logged in with {username}")

    except TimeoutException as e:  # Handle timeout error
        print("Error: Timeout while trying to find login fields.")
        raise e

def scrape_authorization_numbers(roid_numbers, df, accounts, group_name, save_interval=2):
    """Scrapes authorization numbers using alternating accounts within a group."""
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    # Create separate browser instances for each account in the group
    drivers = [webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options) for _ in accounts]

    # Login with each account
    for driver, account in zip(drivers, accounts):
        login(driver, account[0], account[1])

    try:
        for index, roid_number in enumerate(roid_numbers, start=1):
            driver_to_use = drivers[index % len(drivers)]  # Alternate accounts

            try:
                print(f"{group_name} Processing {index}/{len(roid_numbers)}: ai_order_id {roid_number} using Account {index % len(accounts) + 1}")

                repair_order_url = f'https://online.autointegrate.com/EditRepairOrder?jsId={roid_number}&jsro=false&roep=1'
                driver_to_use.get(repair_order_url)
                time.sleep(random.randint(1, 3))  # Random sleep to avoid detection

                # Scraping Authorization Number
                try:
                    authorization_number_element = WebDriverWait(driver_to_use, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//span[contains(., 'Authorization Number:')]"))
                    )
                    authorization_number = authorization_number_element.text.split("Authorization Number: ")[-1].strip()
                    df.loc[df['ai_order_id'] == roid_number, 'Authorization Number'] = authorization_number
                    print(f"Found authorization number: {authorization_number}")
                except Exception as e:
                    df.loc[df['ai_order_id'] == roid_number, 'Authorization Number'] = ''
                    print(f"Error extracting Authorization Number for ai_order_id {roid_number}: {e}")

                # Save data periodically
                if index % save_interval == 0:
                    save_partial_data(df)

            except (NoSuchElementException, TimeoutException, Exception) as e:
                print(f"Skipped {roid_number} due to error: {e}. Moving to next ai_order_id.")
                continue

    finally:
        for driver in drivers:
            driver.quit()  # Close all driver instances
        save_partial_data(df)  # Final save to ensure data is not lost

    return df  # Return the final DataFrame

def scrape_authorization_numbers_multi_account(roid_numbers, df, accounts):
    """Splits the scraping task into two groups and runs them in parallel."""
    # Split ai_order_ids into Group A (1-50) and Group B (51-100)
    group_A_roids = roid_numbers[:50]
    group_B_roids = roid_numbers[50:100]

    # Assign accounts to each group
    group_A_accounts = [accounts[0], accounts[1]]  # Accounts 1 & 2 for Group A
    group_B_accounts = [accounts[2], accounts[3]]  # Accounts 3 & 4 for Group B

    # Use ThreadPoolExecutor to run Group A and Group B in parallel
    with ThreadPoolExecutor(max_workers=2) as executor:
        future_A = executor.submit(scrape_authorization_numbers, group_A_roids, df, group_A_accounts, "Group A")
        future_B = executor.submit(scrape_authorization_numbers, group_B_roids, df, group_B_accounts, "Group B")

        df_A = future_A.result()
        df_B = future_B.result()

    return pd.concat([df_A, df_B])

# 3. Function to scrape authorization numbers with multiple accounts
def scrape_authorization_numbers_multi_account(roid_numbers, df, accounts, 
                                                scrape_authorization=True, 
                                                scrape_repair_order_id=True, 
                                                scrape_subtotal=True, 
                                                scrape_total=True, 
                                                scrape_payable=True, 
                                                scrape_payment_img=True, 
                                                scrape_invoice=True, 
                                                scrape_status = True,
                                                save_interval=2):
    # Set up options to maximize browser window
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    drivers = []  # Initialize list to hold browser instances

    if len(accounts) == 1:
        # Single account: use one browser instance
        driver1 = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
        login(driver1, accounts[0][0], accounts[0][1])  # Login to first account
        drivers.append(driver1)  # Add driver to list
    else:
        # Multiple accounts: create browser instances for each account
        for account in accounts:
            driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
            login(driver, account[0], account[1])  # Login with each account
            drivers.append(driver)  # Add driver to list

    try:
        for driver in drivers:
            # Wait until the search input is ready on each driver instance
            WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Repair Order ID']")))

        for index, roid_number in enumerate(roid_numbers, start=1):
            # Alternate browsers for each ai_order_id if multiple accounts
            driver_to_use = drivers[index % len(drivers)]

            try:
                print(f"Processing {index}/{len(roid_numbers)}: ai_order_id {roid_number} using driver {index % len(drivers) + 1}")
                # Construct URL for each ai_order_id
                repair_order_url = f'https://online.autointegrate.com/EditRepairOrder?jsId={roid_number}&jsro=false&roep=1'
                driver_to_use.get(repair_order_url)  # Open the page for current ai_order_id
                time.sleep(random.randint(1, 3))  # Random sleep for throttling

                # Helper function to close pop-up if it exists
                def close_popup_if_exists(driver):
                    try:
                        popup = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, "div.MuiDialog-paper"))
                        )
                        close_button = popup.find_element(By.XPATH, ".//button[@aria-label='Close']")
                        close_button.click()
                    except TimeoutException:
                        pass

                # Helper function to open a specific tab by ID
                def open_tab(driver, tab_id):
                    try:
                        tab = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, tab_id))
                        )
                        tab.click()
                    except TimeoutException:
                        pass

                close_popup_if_exists(driver_to_use)  # Close pop-up if any

                # Scraping logic for various fields based on user selection
                if scrape_authorization:
                    try:
                        authorization_number_element = WebDriverWait(driver_to_use, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(., 'Authorization Number:')]"))
                        )
                        raw = authorization_number_element.text.split("Authorization Number:")[-1]
                        # now drop anything after the first newline
                        authorization_number = raw.splitlines()[0].strip()
                        df.loc[df['ai_order_id'] == roid_number, 'Authorization Number'] = authorization_number
                        print(f"Found authorization number: {authorization_number}")
                    except Exception as e:
                        df.loc[df['ai_order_id'] == roid_number, 'Authorization Number'] = ''
                        print(f"An error occurred while extracting the authorization number: {e}")

                if scrape_repair_order_id:
                    try:
                        repair_order_id_element = WebDriverWait(driver_to_use, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//h3[contains(., 'Repair Order ID:')]"))
                        )
                        repair_order_id = repair_order_id_element.text.replace("Repair Order ID: ", "").strip()
                        df.loc[df['ai_order_id'] == roid_number, 'Repair Order ID'] = repair_order_id
                        print(f"Found Repair Order ID: {repair_order_id}")
                    except Exception as e:
                        df.loc[df['ai_order_id'] == roid_number, 'Repair Order ID'] = ''
                        print(f"An error occurred while extracting the Repair Order ID: {e}")

                if scrape_subtotal:
                    try:
                        label_element = WebDriverWait(driver_to_use, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//span[text()="SubTotal (exc. Tax)"]'))
                        )
                        parent_div = label_element.find_element(By.XPATH, './../following-sibling::div')
                        subtotal_exc_tax_element = parent_div.find_element(By.XPATH, './/span[contains(@class, "MuiTypography-totalsStaticData")]')
                        subtotal_exc_tax = subtotal_exc_tax_element.text.strip()
                        df.loc[df['ai_order_id'] == roid_number, 'SubTotal (exc. Tax)'] = subtotal_exc_tax
                        print(f"Found SubTotal (exc. Tax): {subtotal_exc_tax}")
                    except Exception as e:
                        df.loc[df['ai_order_id'] == roid_number, 'SubTotal (exc. Tax)'] = ''
                        print(f"An error occurred while extracting the SubTotal (exc. Tax): {e}")

                if scrape_status:
                    try:
                        # Wait for the status label to be present
                        label_element = WebDriverWait(driver_to_use, 10).until(
                            EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[10]/div/div[1]/div/section/div[1]/div/div/span[2]'))
                        )
                        
                        # Extract the status value
                        status_element = label_element.text.strip()
                        
                        # Assuming 'ai_order_id' is the unique identifier for the row you're updating
                        df.loc[df['ai_order_id'] == roid_number, 'Status'] = status_element
                        print(f"Found Status: {status_element}")
                        
                    except Exception as e:
                        # In case of any error, leave the field empty for that specific row
                        df.loc[df['ai_order_id'] == roid_number, 'Status'] = ''
                        print(f"An error occurred while extracting the Status: {e}")

                if scrape_total:
                    try:
                        total_inc_tax_label = WebDriverWait(driver_to_use, 15).until(
                            EC.presence_of_element_located((By.XPATH, '//span[text()="Total (inc. Tax)"]'))
                        )
                        total_inc_tax_div = total_inc_tax_label.find_element(By.XPATH, './../following-sibling::div')
                        total_inc_tax_element = total_inc_tax_div.find_element(By.XPATH, './/span[contains(@class, "MuiTypography-totalsStaticData")]')
                        total_inc_tax = total_inc_tax_element.text.strip()
                        df.loc[df['ai_order_id'] == roid_number, 'Total (inc. Tax)'] = total_inc_tax
                        print(f"Found Total (inc. Tax): {total_inc_tax}")
                    except Exception as e:
                        df.loc[df['ai_order_id'] == roid_number, 'Total (inc. Tax)'] = ''
                        print(f"An error occurred while extracting the Total (inc. Tax): {e}")

                if scrape_payable:
                    try:
                        payable_amount_inc_tax_label = WebDriverWait(driver_to_use, 15).until(
                            EC.presence_of_element_located((By.XPATH, '//span[text()="Payable Amount (inc. Tax)"]'))
                        )
                        payable_amount_inc_tax_div = payable_amount_inc_tax_label.find_element(By.XPATH, './../following-sibling::div')
                        payable_amount_inc_tax_element = payable_amount_inc_tax_div.find_element(By.XPATH, './/span[contains(@class, "MuiTypography-totalsStaticData")]')
                        payable_amount_inc_tax = payable_amount_inc_tax_element.text.strip()
                        df.loc[df['ai_order_id'] == roid_number, 'Payable Amount (inc. Tax)'] = payable_amount_inc_tax
                        print(f"Found Payable Amount (inc. Tax): {payable_amount_inc_tax}")
                    except Exception as e:
                        df.loc[df['ai_order_id'] == roid_number, 'Payable Amount (inc. Tax)'] = ''
                        print(f"An error occurred while extracting the Payable Amount (inc. Tax): {e}")

                if scrape_payment_img:
                    open_tab(driver_to_use, "simple-tab-4")  # Switch to the Payment Details tab
                    try:
                        payment_direction_img_element = WebDriverWait(driver_to_use, 30).until(
                            EC.presence_of_element_located((By.XPATH, "//img[@alt='Independent']"))
                        )
                    except Exception:
                        try:
                            payment_direction_img_element = WebDriverWait(driver_to_use, 30).until(
                                EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div[6]/div/div[2]/div[5]/div/div/div/div[5]/div/div[2]/img"))
                            )
                        except Exception as e:
                            df.loc[df['ai_order_id'] == roid_number, 'Payment Direction Img URL'] = ''
                            print(f"An error occurred while extracting the payment direction image URL: {e}")
                        else:
                            payment_direction_img_url = payment_direction_img_element.get_attribute('src')
                            df.loc[df['ai_order_id'] == roid_number, 'Payment Direction Img URL'] = payment_direction_img_url
                            print(f"Found payment direction image URL using second XPath: {payment_direction_img_url}")
                    else:
                        payment_direction_img_url = payment_direction_img_element.get_attribute('src')
                        df.loc[df['ai_order_id'] == roid_number, 'Payment Direction Img URL'] = payment_direction_img_url
                        print(f"Found payment direction image URL using first XPath: {payment_direction_img_url}")

                if scrape_invoice:
                    try:
                        invoice_number_element = WebDriverWait(driver_to_use, 20).until(
                            EC.presence_of_element_located((By.XPATH, "//button[contains(@aria-label, 'Change')]//span[contains(@class, 'MuiTypography-vehicleDetailsStaticData')]"))
                        )
                        invoice_number = invoice_number_element.text.strip()
                        df.loc[df['ai_order_id'] == roid_number, 'Invoice Number'] = invoice_number
                        print(f"Found invoice number: {invoice_number}")
                    except Exception as e:
                        df.loc[df['ai_order_id'] == roid_number, 'Invoice Number'] = ''
                        print(f"Error finding invoice number: {e}")

                # Save data periodically based on save_interval
                if index % save_interval == 0:
                    save_partial_data(df)  # Save current progress

            except (NoSuchElementException, TimeoutException, Exception) as e:
                print(f"Skipped {roid_number} due to error: {e}. Moving to next ai_order_id.")
                continue

    finally:
        for driver in drivers:
            driver.quit()  # Close all driver instances at the end of scraping
        save_partial_data(df)  # Final save to ensure data is not lost

    return df  # Return the final DataFrame with scraped data

# Main function for handling Streamlit interface and triggering the multi-account scraping


def main():
    st.title("Authorization Number Scraper")

    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])
    if uploaded_file is not None:
        with st.spinner('Processing file...'):
            df, roid_numbers, *_ = process_file(uploaded_file)

            # Sidebar Section for Field Selection
            st.sidebar.header("Select Values to Scrape")
            scrape_authorization = st.sidebar.checkbox("Scrape Authorization Number", value=True, disabled=True)
            scrape_repair_order_id = st.sidebar.checkbox("Scrape Repair Order ID", value=True, disabled=True)
            scrape_subtotal = st.sidebar.checkbox("Scrape SubTotal (exc. Tax)", value=True)
            scrape_total = st.sidebar.checkbox("Scrape Total (inc. Tax)", value=True)
            scrape_payable = st.sidebar.checkbox("Scrape Payable Amount (inc. Tax)", value=True)
            scrape_payment_img = st.sidebar.checkbox("Scrape Payment Direction Img URL", value=True)
            scrape_invoice = st.sidebar.checkbox("Scrape Invoice Number", value=True)
            scrape_status = st.sidebar.checkbox("Scrape Status", value=True)

            # Sidebar Section for Splitting Data
            st.sidebar.header("Data Split Options")
            split_data = st.sidebar.radio("How to split data?", options=["No Split", "50:50"], index=0)

            # Sidebar Section for Account Credentials
            st.sidebar.header("Account Credentials")
            accounts = []
            if split_data == "50:50":
                account_count = 4  # Require 4 accounts for 50:50 split
            else:
                account_count = st.sidebar.number_input("Number of Accounts to Use", min_value=1, max_value=10, value=1, step=1)

            

            for i in range(account_count):
                username = st.sidebar.text_input(f"Username for Account {i+1}", key=f"username_{i}")
                password = st.sidebar.text_input(f"Password for Account {i+1}", key=f"password_{i}", type="password")
                accounts.append((username, password))

            if split_data == "50:50" and len(roid_numbers) == 100:
                part1 = roid_numbers[:50]  # First half
                part2 = roid_numbers[50:]  # Second half
                first_accounts = accounts[:2]  # Accounts 1 & 2
                second_accounts = accounts[2:]  # Accounts 3 & 4

                # Debugging output
                print(f"Part 1: {len(part1)} records")
                print(f"Part 2: {len(part2)} records")
                print(f"First accounts: {first_accounts}")
                print(f"Second accounts: {second_accounts}")

                def scrape_part(part, acc_group):
                    """ Assigns alternating accounts within a group to scrape the given data. """
                    print(f"Scraping {len(part)} records with accounts: {acc_group}")
                    return scrape_authorization_numbers_multi_account(
                        part, df, acc_group, scrape_authorization, scrape_repair_order_id,
                        scrape_subtotal, scrape_total, scrape_payable, scrape_payment_img, scrape_invoice
                    )

                if st.button('Start Scraping Process'):
                    with st.spinner('Scraping authorization numbers...'):
                        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                            # Submit the scraping tasks in parallel
                            future1 = executor.submit(scrape_part, part1, first_accounts)  # Group A
                            future2 = executor.submit(scrape_part, part2, second_accounts)  # Group B

                            # Wait for results in parallel
                            df_part1 = future1.result()
                            df_part2 = future2.result()

                        # Check the results for each part
                        print("Part 1 Scraped Data:")
                        print(df_part1.head())  # Display first few rows of part 1

                        print("Part 2 Scraped Data:")
                        print(df_part2.head())  # Display first few rows of part 2

                        # Merge the results
                        df = df_part1.append(df_part2, ignore_index=True)

            else:
                if st.button('Start Scraping Process'):
                    with st.spinner('Scraping authorization numbers...'):
                        df = scrape_authorization_numbers_multi_account(
                            roid_numbers, df, accounts, scrape_authorization, scrape_repair_order_id,
                            scrape_subtotal, scrape_total, scrape_payable, scrape_payment_img, scrape_invoice
                        )

            # Display Scraped Data
            st.write(df)

            # Generate Download Link
            towrite = BytesIO()
            df.to_excel(towrite, index=False)
            towrite.seek(0)
            b64 = base64.b64encode(towrite.read()).decode()
            linko = f'<a href="data:application/octet-stream;base64,{b64}" download="updated_file.xlsx">Download updated file</a>'
            st.markdown(linko, unsafe_allow_html=True)

if __name__ == "__main__":
    main()