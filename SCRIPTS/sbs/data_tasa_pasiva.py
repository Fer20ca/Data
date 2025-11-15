import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys  # for special keystrokes like TAB
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def wait_for_downloads(download_folder, timeout=60):
    """
    Polls the download folder until there are no Chrome temporary download files (.crdownload)
    or until the timeout (in seconds) is reached.
    """
    seconds = 0
    while seconds < timeout:
        if not any(fname.endswith('.crdownload') for fname in os.listdir(download_folder)):
            return True
        time.sleep(1)
        seconds += 1
    return False

# WebDriver path (adjust according to your setup)
webdriver_path = r'C:\Users\José Estrada\OneDrive - ABC Capital\Documentos\chromedriver-win64\chromedriver.exe'

# Folder where files will be saved
download_folder = r"C:\Users\José Estrada\OneDrive - ABC Capital\Tasa promedio pasiva"
if not os.path.exists(download_folder):
    os.makedirs(download_folder)

# Set up Chrome options with safe browsing disabled (to avoid download blocks)
chrome_options = Options()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
})

# Initialize the WebDriver
service = Service(webdriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Base URL for the passive deposits page
base_url = "https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIPasivaDepositoEmpresa.aspx"

# Define types for pages using a date input vs. dropdowns
types_with_date_input = ['B', 'F']
types_with_dropdowns = ['C', 'R']

# Define dates (last day of each month from March 2024 to March 2025)
dates = [
    ('27', '03', '2024'),
    ('30', '04', '2024'),
    ('31', '05', '2024'),
    ('28', '06', '2024'),
    ('31', '07', '2024'),
    ('30', '08', '2024'),
    ('30', '09', '2024'),
    ('31', '10', '2024'),
    ('29', '11', '2024'),
    ('30', '12', '2024'),
    ('31', '01', '2025'),
    ('28', '02', '2025'),
    ('31', '03', '2025')
]

# Mapping of month numbers to Spanish month names
month_mapping = {
    '01': 'Enero', '02': 'Febrero', '03': 'Marzo', '04': 'Abril',
    '05': 'Mayo', '06': 'Junio', '07': 'Julio', '08': 'Agosto',
    '09': 'Setiembre', '10': 'Octubre', '11': 'Noviembre', '12': 'Diciembre'
}

def select_currency(currency):
    try:
        if currency.upper() == "ME":
            # Click on "Moneda Extranjera"
            currency_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphContent_lbtnMex"))
            )
            currency_link.click()
            print("Selected Moneda Extranjera (ME).")
            # Wait for the page to update the currency selection
            time.sleep(5)
        elif currency.upper() == "MN":
            # MN is the default mode so no need to click any button
            print("MN is the default mode; no currency switch needed.")
        else:
            print(f"Unknown currency: {currency}")
            return False
        return True
    except Exception as e:
        print(f"Error selecting currency {currency}: {str(e)}")
        return False

# Specify the desired currency: "MN" or "ME"
desired_currency = "MN"  # Change to "ME" if you need that mode

try:
    # Process pages that use a date input for specifying the date
    for type_code in types_with_date_input:
        for day, month, year in dates:
            driver.get(f"{base_url}?tip={type_code}")

            if not select_currency(desired_currency):
                print(f"Skipping type {type_code} for {day}/{month}/{year} due to currency selection failure")
                continue

            # Wait for the date input wrapper to be present
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_cphContent_rdpDate_dateInput_wrapper"))
            )
            # Locate the actual text input element
            date_input = driver.find_element(By.ID, "ctl00_cphContent_rdpDate_dateInput")
            date_input.clear()
            date_str = f"{day}/{month}/{year}"
            date_input.send_keys(date_str)
            # Send TAB to trigger any blur/change events
            date_input.send_keys(Keys.TAB)
            print(f"Set date to {date_str} for type {type_code}.")

            # Click the "Consultar" button
            consult_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphContent_btnConsultar"))
            )
            consult_button.click()
            print(f"Clicked 'Consultar' for type {type_code} on {date_str}.")

            # Wait for the "Exportar" button to become clickable
            export_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphContent_btnExportar"))
            )
            print("Page loaded successfully after 'Consultar'.")
            # Use JavaScript click to ensure proper event firing
            driver.execute_script("arguments[0].click();", export_button)
            print(f"Clicked 'Exportar' for type {type_code} on {date_str}.")

            # Wait for the download to complete
            if wait_for_downloads(download_folder, timeout=60):
                print(f"Download completed for type {type_code} on {date_str}.")
            else:
                print(f"Timeout waiting for download for type {type_code} on {date_str}.")

    # Process pages that use dropdowns for year and month selection
    for type_code in types_with_dropdowns:
        for day, month, year in dates:
            driver.get(f"{base_url}?tip={type_code}")
            
            if not select_currency(desired_currency):
                print(f"Skipping type {type_code} for {month}/{year} due to currency selection failure")
                continue

            # Wait for and select the year from the dropdown
            year_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_cphContent_rAnio"))
            )
            year_dropdown.click()
            year_option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[text()='{year}']"))
            )
            year_option.click()
            print(f"Set year to {year} for type {type_code}.")

            # Select the month from the dropdown
            month_dropdown = driver.find_element(By.ID, "ctl00_cphContent_rMes")
            month_dropdown.click()
            month_name = month_mapping.get(month, month)
            month_option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[text()='{month_name}']"))
            )
            month_option.click()
            print(f"Set month to {month_name} for type {type_code}.")

            # Click the "Consultar" button
            consult_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphContent_btnConsultaMensual"))
            )
            consult_button.click()
            print(f"Clicked 'Consultar' for type {type_code} on {month_name}/{year}.")

            # Wait for and click the "Exportar" button
            export_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphContent_btnExportarM"))
            )
            driver.execute_script("arguments[0].click();", export_button)
            print(f"Clicked 'Exportar' for type {type_code} on {month_name}/{year}.")

            # Wait for download to complete
            if wait_for_downloads(download_folder, timeout=60):
                print(f"Download completed for type {type_code} on {month_name}/{year}.")
            else:
                print(f"Timeout waiting for download for type {type_code} on {month_name}/{year}.")

except Exception as e:
    print(f"An error occurred: {str(e)}")
finally:
    driver.quit()

print("All tasks completed.")
