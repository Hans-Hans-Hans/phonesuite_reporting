from dotenv import load_dotenv
import os
from to_excel import create_workbook
from send_email import send_email
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.support.ui import Select
import time

load_dotenv()

# Configuration
url: str = os.getenv('url')
uname: str = os.getenv('usern')
pword: str = os.getenv('pword')

# Set up selenium
driver = webdriver.Chrome()
driver.get(url)

# Login
try:
     # Wait for username field
    time.sleep(2)
    driver.find_element(By.NAME, "username").send_keys(uname)
    driver.find_element(By.NAME, "password").send_keys(pword)
    
    # Find the <select> element first (replace the selector with the actual one)
    select_element = driver.find_element(By.ID, "product")  # or By.NAME, By.XPATH etc.

    # Wrap it with Select class
    select = Select(select_element)

    # Select the option by visible text
    select.select_by_visible_text("Configurator")
    
    # Press Submit
    driver.find_element(By.NAME, "Submit").click()
    time.sleep(2)
    
except Exception as e:
    print(f"An unexpected error has occured: {e}")
    
try:
    driver.find_element(By.LINK_TEXT, "Devices").click()
    
    # Explicit wait for the table rows to be loaded (up to 10 seconds)
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#device_table tbody tr"))
    )
    
    rows = driver.find_elements(By.CSS_SELECTOR, "#device_table tbody tr")
    data = []

    while True:
        # Wait for table rows to be present
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#device_table tbody tr"))
            )
        except TimeoutException:
            print("Timeout waiting for table rows.")
            break

        # Grab the rows on current page
        rows = driver.find_elements(By.CSS_SELECTOR, "#device_table tbody tr")

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            device_id = cells[0].find_element(By.CSS_SELECTOR, 'input.dev_radio').get_attribute("value")
            dev_type = cells[1].text.strip()
            name = cells[2].text.strip()
            status = cells[3].text.strip()
            ip = cells[4].text.strip()
            port = cells[5].text.strip()
            description = cells[6].text.strip()
            assignments = cells[7].text.strip()
            pickup_grp = cells[8].text.strip()
            sla = cells[9].text.strip()
            is_trunk = cells[10].text.strip()

            data.append([
                device_id, dev_type, name, status, ip, port, description,
                assignments, pickup_grp, sla, is_trunk
            ])

        # Try to find and click the "Next" button if enabled
        try:
            next_button = driver.find_element(By.ID, "device_table_next")

            # Check if the button is disabled (usually by a CSS class, e.g. 'disabled' or 'ui-state-disabled')
            classes = next_button.get_attribute("class")
            if "disabled" in classes or "ui-state-disabled" in classes:
                print("Reached last page.")
                break  # Exit loop if no more pages

            # Click the next button to go to the next page
            next_button.click()

            # Wait a bit for the page to refresh (could use explicit wait on table rows changing)
            time.sleep(2)

        except (NoSuchElementException, ElementNotInteractableException) as e:
            print(f"Next button not found or not clickable: {e}")
            break

    # After loop completes, print or process the full data list
    print(f"Total devices scraped: {len(data)}")
    for device in data:
        print(device)

except Exception as e:
    print(f"Failed to open devices tab or extract data: {e}")
    
# Sort data: unregistered entries go first
data.sort(key=lambda row: "Unregistered" not in row[3])
    
create_workbook(data)
send_email()