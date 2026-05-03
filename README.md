# IT23713680_ITPM_Assignment_1

Sinhala Transliteration Accuracy Testing 
Playwright Automation 🇱🇰

📌 Project Overview
This repository contains the automated testing suite for the IT3040 - ITPM Assignment 
The primary goal is to evaluate the transliteration accuracy of the Pixels Suite Chat Translator by identifying 50 scenarios where the system fails to correctly convert chat-style Singlish into Sinhala.
The project uses Playwright for web automation and OpenPyXL for handling Excel data.  

🎯 Key Objectives
Identify 50 negative test cases where the transliteration is inaccurate.  
Cover all 24 Singlish input types (Questions, Commands, Slang, Emojis, etc.) specified in the assignment.  
Automate the testing process and record results (Actual Output and Status) directly into an Excel file

🛠️ Technical Stack
Language: Python 3.11+
Automation Tool: Playwright (Python Sync API)
Data Handling: OpenPyXL (Excel integration)
Browser: Chromium / Google Chrome

📦 Setup & Installation
1. Prerequisites
Ensure you have Python installed on your system.  

2. Install Dependencies
Run the following commands in your terminal to set up the environment:
pip install playwright openpyxl
playwright install

🚀 Execution
To run the automated tests and generate results in the Excel sheet, execute the following command:

python test_automation.py--excel "IT23713680.xlsx" --url "https://www.pixelssuite.com/chat-translator" --output-wait-ms 8000

📂 Project Structure
test_automation.py: The core logic for web automation and data processing.  

IT23713680.xlsx: The Excel file containing the 50 test cases and recorded results.  

README.md: Project documentation and execution guide.  

Git_Link.txt: Text file containing the public repository URL

👤 Author
Name:NIRMAL BAAJ
Registration Number: IT23713680
Batch: BSc (Hons) in Information Technology - Year 3 Semester 1
