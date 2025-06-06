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
from selenium.webdriver.chrome.options import Options
import time

# Configure Chrome options to run headless (no GUI)
options = Options()
options.add_argument('--headless')             # Run browser in headless mode
options.add_argument('--disable-gpu')          # Disable GPU hardware acceleration (optional)
options.add_argument('--no-sandbox')           # Bypass OS security model (needed on some systems)
options.add_argument('--window-size=1920,1080')  # Set viewport size to avoid hidden elements

# Load environment variables from .env file
load_dotenv()

# Get credentials and URL from environment variables
url: str = os.getenv('url')
uname: str = os.getenv('usern')
pword: str = os.getenv('pword')

# Initialize WebDriver with headless configuration
driver = webdriver.Chrome(options=options)
driver.get(url)

def scrape():
    """
    Automates login, navigation, and scraping of Phonesuite device data using Selenium.

    The script logs into the portal, navigates to the Devices tab, paginates through the
    table, extracts relevant device data, sorts unregistered devices to the top, saves
    it as an Excel file, and emails it.
    """
    # --- LOGIN ---
    try:
        time.sleep(2)  # Wait for page to load
        driver.find_element(By.NAME, "username").send_keys(uname)
        driver.find_element(By.NAME, "password").send_keys(pword)

        # Select 'Configurator' from the dropdown
        select_element = driver.find_element(By.ID, "product")
        Select(select_element).select_by_visible_text("Configurator")

        # Submit form
        driver.find_element(By.NAME, "Submit").click()
        time.sleep(2)

    except Exception as e:
        print(f"An unexpected error occurred during login: {e}")
        return

    # --- SCRAPE DEVICE TABLE ---
    try:
        driver.find_element(By.LINK_TEXT, "Devices").click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#device_table tbody tr"))
        )

        data = []

        while True:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#device_table tbody tr"))
                )
            except TimeoutException:
                print("Timeout while waiting for table rows.")
                break

            rows = driver.find_elements(By.CSS_SELECTOR, "#device_table tbody tr")

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")

                # Extract each cell's content
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

            # --- PAGINATION ---
            try:
                next_button = driver.find_element(By.ID, "device_table_next")
                classes = next_button.get_attribute("class")

                if "disabled" in classes or "ui-state-disabled" in classes:
                    print("Reached last page.")
                    break

                next_button.click()
                time.sleep(2)

            except (NoSuchElementException, ElementNotInteractableException) as e:
                print(f"Next button not found or not clickable: {e}")
                break

        print(f"Total devices scraped: {len(data)}")
        for device in data:
            print(device)

    except Exception as e:
        print(f"Error scraping devices: {e}")
        return

    # --- SORT DATA ---
    # Unregistered entries go first in the Excel file
    data.sort(key=lambda row: "Unregistered" not in row[3])

    # --- EXPORT TO EXCEL ---
    create_workbook(data)

    # --- SEND EMAIL ---
    send_email()

# Main execution block
if __name__ == "__main__":
    scrape()
