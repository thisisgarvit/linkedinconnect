import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Ensure the file is in the current working directory
current_dir = os.getcwd()
excel_file_path = os.path.join(current_dir, "LinkedinAutomation.xlsx")

# Load the Excel file
try:
    df = pd.read_excel(excel_file_path, sheet_name="Sheet1", engine='openpyxl')
except Exception as e:
    print(f"Error reading the Excel file: {e}")
    exit()

# Initialize the WebDriver with existing session
options = Options()
options.add_argument("user-data-dir=/workspace/chrome-profile")  # Change this to your Chrome profile path in Gitpod
service = Service('/usr/local/bin/chromedriver')  # Path to ChromeDriver in Gitpod
driver = webdriver.Chrome(service=service, options=options)

# Function to check connection status and send requests/messages
def process_linkedin_profiles(driver, profile_urls):
    for url in profile_urls:
        driver.get(url)
        time.sleep(5)  # Wait for the profile page to load
        try:
            # Check connection status
            connect_button = driver.find_element(By.XPATH, "//button[contains(@aria-label, 'Connect')]")
            message_button = driver.find_element(By.XPATH, "//button[contains(@aria-label, 'Message')]")
            
            if connect_button:
                connect_button.click()
                time.sleep(2)
                send_button = driver.find_element(By.XPATH, "//button[@aria-label='Send now']")
                send_button.click()
                print(f"Connection request sent to {url}")
            elif message_button:
                message_button.click()
                time.sleep(2)
                message_box = driver.find_element(By.XPATH, "//textarea[@name='message']")
                message = "Hello! I hope you are doing well."
                message_box.send_keys(message)
                send_button = driver.find_element(By.XPATH, "//button[@aria-label='Send']")
                send_button.click()
                print(f"Message sent to {url}")
        except Exception as e:
            print(f"Could not process {url}: {e}")

# Process each LinkedIn profile from the Excel file
profile_urls = df['LinkedIn URL'].tolist()
process_linkedin_profiles(driver, profile_urls)

# Close the WebDriver
driver.quit()
