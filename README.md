# Google Apps Script for Sending Custom Emails

This script is designed to automate the process of sending custom emails based on the data in a Google Sheets document. It's particularly useful for situations where emails need to be sent based on specific criteria, such as application statuses.

## Features
- Fetches email templates from Google Docs and sends them as HTML emails.
- Personalizes emails with data from Google Sheets.
- Handles different email templates based on application status.

## Setup
1. **Google Sheets**: Set up a Google Sheets document with the required data.
2. **Google Docs**: Create templates for each type of email in Google Docs.
3. **Script Configuration**: 
   - Replace `your-document-id-for-accepted`, `your-document-id-for-rejected`, and `your-document-id-for-location` with the actual IDs of your Google Docs templates.
   - Replace `Your Sheet Name` with the name of your Google Sheets document.

## Usage
Run the `sendEmails` function to send emails based on the data in your Google Sheets document.

## Notes
- Ensure that the Google Apps Script has the necessary permissions to access Google Sheets and Google Docs.
- Modify the script to match the specific structure of your Google Sheets and the placeholders in your Google Docs templates.
