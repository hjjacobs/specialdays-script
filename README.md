# Send a mail with a list of all special dates (birthdays and anniversaries) of your contacts using Google Apps Script

This script uses [Google Apps Script](https://script.google.com/home/) to access one's Google Contacts and calculate the age on a person's birthday & anniversaries and send a mail with for 

Using Google Contacts and the Google Apps Script environment mail me a list of all special dates of contacts (birthdays and anniversaries).
* Send an email with a list of birthdays of people in my Google contacts for the remainder of the current year. Starting in the month the script runs.
* Send an email with a list of anniversaries of people in my Google contacts for the remainder of the current year. Starting in the month the script runs.

## Steps to setup everything

1. Head to https://script.google.com/home/my and create a new project. Rename the existing file Code.gs to your liking and paste the code from sendSpecialDaysList.gs.
2. Paste your mail address in line 8
3. Customize the mail text in lines 9-17
4. Customize the mail text for each birthday in line 218
4. Customize the mail text for each anniversary in line 270
5. Enable the People API
    - Before running this script, you must enable the People API:
    - In your Google Apps Script editor, click on "Services" (the + icon)
    - Find "People API" and add it
    - Set the identifier to People and version to v1
6. Hit run and see the mails in your inbox
7. Create a trigger to run this every month for a monthly reminder

Remove line 37 and 38 when you do not want a mail for the anniversaries (and/or remover the code completely).

## Notes

This my first javascript code, it needs more refactoring. Any tips are welcome.

<hr>

Inspired by [Calculate the age of a person and write it to the event's description in your birthday calendar using Google Apps Script](https://gist.github.com/bene-we/e0a306ad6788fec5dbe45cde2de2f140).
