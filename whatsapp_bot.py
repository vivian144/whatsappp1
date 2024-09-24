import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time

# Load the Excel file
file_path = 'contacts.xlsx'  # Path to your Excel file
df = pd.read_excel(file_path)

# Set up ChromeDriver Service
service = Service(executable_path="C:/chromedriver/chromedriver-win64 (1)/chromedriver-win64/chromedriver.exe")  # Replace with your actual chromedriver path

# Initialize Selenium WebDriver with the service
driver = webdriver.Chrome(service=service)

# Open WhatsApp Web
driver.get('https://web.whatsapp.com')

# Allow time for user to scan the QR code
print("Please scan the QR code to log into WhatsApp Web.")
time.sleep(20)  # Adjust this if you need more time

# Get the message template from the user
message_template = input("Enter your message template (use {name} as placeholder): ")

# Loop through each contact and send a WhatsApp message
for index, row in df.iterrows():
    name = row['Name']  # Column A - Names
    phone = str(row['Phone Number'])  # Column B - Phone Numbers

    # Customize the message for each contact
    personalized_message = message_template.replace("{name}", name)

    print(f"Sending to {name} at {phone}: {personalized_message}")

    # Search for the phone number in WhatsApp Web
    search_box = driver.find_element(By.XPATH, '//div[@title="Search input textbox"]')
    search_box.clear()
    search_box.send_keys(phone)
    search_box.send_keys(Keys.ENTER)
    time.sleep(3)  # Wait for the contact's chat to load

    # Locate the message box and send the personalized message
    message_box = driver.find_element(By.XPATH, '//div[@title="Type a message"]')
    message_box.send_keys(personalized_message)
    message_box.send_keys(Keys.ENTER)

    # Wait a moment before sending the next message
    time.sleep(2)  # Adjust this as needed for speed

# Close the browser after sending all messages
driver.quit()

print("All messages sent successfully!")
