import win32com.client
import os

try:
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = "Test Email from Python"
    mail.Body = "This is a test email draft."
    mail.Display()  # This should open Outlook with a draft
    print("Outlook draft created successfully.")
except Exception as e:
    print(f"Error: {e}")
