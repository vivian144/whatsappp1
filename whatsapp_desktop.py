import pandas as pd
import pyautogui
import time
import webbrowser  # To open WhatsApp chat with unsaved numbers

# Load the Excel file
file_path = 'contacts.xlsx'  # Path to your Excel file
df = pd.read_excel(file_path)

# Get the message template from the user
message_template = input("Enter your message template (use {name} as placeholder): ")

# Function to open WhatsApp chat for unsaved numbers
def open_whatsapp_chat(phone):
    url = f"https://wa.me/{phone}"
    webbrowser.open(url)
    time.sleep(5)  # Allow time for WhatsApp to open the chat

# Function to send a message in WhatsApp Desktop
def send_message(name, phone, message):
    # Open chat for the unsaved number
    open_whatsapp_chat(phone)
    
    # Wait for the chat to load
    time.sleep(5)
    
    # Focus the chat box (Assuming WhatsApp Web or Desktop is fully loaded)
    # You can adjust these coordinates if necessary, depending on your screen resolution
    pyautogui.click(500, 800)  # Click in the chat input area
    
    # Wait for chat box to get focus
    time.sleep(1)

    # Write the message
    pyautogui.write(message)
    pyautogui.press('enter')  # Send the message

# Loop through each contact and send a WhatsApp message
for index, row in df.iterrows():
    name = row['Name']
    phone = str(row['Phone Number'])
    
    # Customize the message for each contact
    personalized_message = message_template.replace("{name}", name)
    
    print(f"Sending to {name} at {phone}: {personalized_message}")
    
    # Send the message via WhatsApp Desktop
    try:
        send_message(name, phone, personalized_message)
        time.sleep(2)  # Adjust the sleep time as necessary
        
    except Exception as e:
        print(f"Failed to send message to {name}. Error: {e}")

print("All messages sent successfully!")
