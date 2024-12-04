# Email-Automation
Email Automation with Selenium and python

Overview
This Python script automates the process of sending personalized emails using Gmail via Selenium WebDriver. It reads email IDs and corresponding names from an Excel file using the OpenPyXL library, and sends personalized emails to each recipient through Gmail.

Requirements
Before running the script, ensure you have the following dependencies installed:
Python 3.x libraries, you can install the required libraries using pip:
pip install selenium 
pip install openpyxl 
pip install pandas 
pip install pyautogui 
pip install keyboard

Additional Setup
ChromeDriver: You need the appropriate version of the ChromeDriver that matches your version of Chrome. Download ChromeDriver.
Excel File: Prepare an Excel file (email.xlsx) with email addresses and names. The first column should contain email IDs and the second column should contain the respective names.

Script Overview
The script performs the following actions:

Reads email IDs and names from an Excel file: It uses OpenPyXL to extract email IDs and names from an Excel file located on your local system.
Automates Gmail Login: It logs into a Gmail account using Selenium WebDriver and navigates through the login page.
Composes and sends emails: For each email ID and corresponding name, it composes a personalized email and sends it.
Handles missing names: If a name is missing (or None), it substitutes "Participant" as the default greeting.

Setup Instructions
Step 1: Prepare the Excel File
Create an Excel file named email.xlsx with two columns:

Email ID	Name
example1@example.com	John
example2@example.com	Alice
example3@example.com	None
Save the Excel file at a path you will specify in the script.

Step 2: Set Up ChromeDriver
Ensure that you have ChromeDriver downloaded and placed in a directory. Update the path variable in the script to point to your chromedriver.exe file.

Step 3: Configure the Script
Modify the following parameters in the script:

login_email: Your Gmail email address (used for login).
password: Your Gmail account password.
file_path: Path to the email.xlsx file containing email IDs and names.
path: Path to the chromedriver.exe file.
Step 4: Run the Script
Execute the script using Python:
Step 2: Set Up ChromeDriver
Ensure that you have ChromeDriver downloaded and placed in a directory. Update the path variable in the script to point to your chromedriver.exe file.

Step 3: Configure the Script
Modify the following parameters in the script:

login_email: Your Gmail email address (used for login).
password: Your Gmail account password.
file_path: Path to the email.xlsx file containing email IDs and names.
path: Path to the chromedriver.exe file.
Step 4: Run the Script
Execute the script using Python:

The script will:

Log into Gmail.
Open the Compose email window.
Populate the "To" field with the email address and personalize the greeting with the recipient's name.
Send the email and move to the next recipient.

Notes
Security Considerations: Be cautious when using your Gmail account with automation scripts. Ensure your credentials are protected, and consider using an app-specific password or OAuth 2.0 for more security.
Delays: The script includes delays (time.sleep()) to prevent issues with fast interactions, but these can be adjusted or removed as needed.
License
This project is open-source and free to use. Feel free to modify the code for your own needs.

