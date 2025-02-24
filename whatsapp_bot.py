import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service

# Path to ChromeDriver (Make sure it's updated)
chrome_driver_path = "C:\\Users\\vivia\\OneDrive\\Desktop\\project1\\chromedriver-win64\\chromedriver.exe"

# Load contacts from Excel
file_path = "contacts.xlsx"
df = pd.read_excel(file_path)

# Get message template
message_template = input("Enter your message template (use {name} as placeholder): ")

# Initialize Selenium WebDriver
service = Service(chrome_driver_path)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=service, options=options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
input("Scan the QR code and press ENTER to continue...")

# Function to validate and format phone numbers
def format_phone_number(phone):
    phone = str(phone).strip()
    phone = "".join(filter(str.isdigit, phone))  # Remove non-numeric characters
    
    if len(phone) == 10:  # Assume missing country code
        phone = "+91" + phone
    elif len(phone) == 12 and phone.startswith("91"):
        phone = "+" + phone
    elif not phone.startswith("+91") or len(phone) != 13:
        return None  # Invalid number
    
    return phone

# Loop through contacts and send messages
for index, row in df.iterrows():
    name = row["Name"]
    phone = str(row["Phone Number"])

    phone = format_phone_number(phone)
    if phone is None:
        print(f"Skipping invalid number for {name}: {row['Phone Number']}")
        continue

    personalized_message = message_template.replace("{name}", name)
    print(f"Sending to {name} at {phone}: {personalized_message}")

    # Open chat via WhatsApp URL
    whatsapp_url = f"https://web.whatsapp.com/send?phone={phone}&text={personalized_message}"
    driver.get(whatsapp_url)

    time.sleep(10)  # Wait for chat to load

    try:
        # Locate and send the message
        input_box = driver.find_element(By.XPATH, "//div[@title='Type a message']")
        input_box.send_keys(Keys.ENTER)
        print(f"Message sent to {name} ({phone})")
    except Exception as e:
        print(f"Failed to send message to {name} ({phone}): {e}")

    time.sleep(5)  # Wait before next message

# Close the browser
driver.quit()
print("All messages sent successfully!")
