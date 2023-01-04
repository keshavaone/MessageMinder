# WhatsApp Automation
A Python script that reads a list of recipients, messages, and times from an Excel file and automatically sends the messages to the specified recipients at the specified times.

## Requirements
Python 3.6 or higher
Selenium
ChromeDriver
pandas

## Setup
Install the required packages:
```
pip install selenium pandas
```
Download ChromeDriver from [here](https://chromedriver.chromium.org/) and add it to your PATH, or specify its location in the script.

Prepare an Excel file with the following columns:

Recipient: the phone number of the recipient (including the country code)
Message: the message to send
Time: the time to send the message
Add one row for each message to be sent.
Set the path to the Excel file in the script.

## Run the script:

python send_whatsapp_messages.py

## Notes
Make sure that you have the latest version of Chrome installed.
The script may not work if WhatsApp Web is open on your computer.
The script may trigger a security feature in WhatsApp that requires you to verify your phone number. In this case, you will need to manually enter the verification code on your phone.
