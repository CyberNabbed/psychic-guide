OUTLOOK TICKET MONITOR
======================

What is this?
-------------
This is a visual alerting tool for agents. It monitors a specific Outlook folder 
(like "Helpdesk Tickets") and displays a massive, flashing ASCII number on the 
screen representing the count of unread emails.

It is designed to be running on a secondary monitor to alert the team when 
pending work is available.

Key Features:
- Displays unread count in huge digits.
- Runs a "Matrix" style animation to grab attention.
- ONE-BUTTON CLEANUP: Press [ENTER] at any time to instantly mark all 
  unread emails in that folder as "Read".

Requirements:
- Windows OS
- Microsoft Outlook (Desktop Application installed and running)
- Python 

Setup:
1. Open a terminal/command prompt and install the requirements:
   pip install pywin32 colorama

2. **IMPORTANT: CONFIGURATION**
   Open the script file in a text editor (like Notepad).
   Look for the section at the top marked "USER CONFIGURATION".
   Change "YOUR_TARGET_FOLDER_NAME_HERE" to the exact name of the 
   Outlook folder you want to watch.

Usage:
Run the script from your terminal:
python script_name.py

- The screen will update automatically as new emails arrive.
- Press [ENTER] to clear the queue (mark all as read).
