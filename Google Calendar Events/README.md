# Calendar Support Event Creator

This script automates the creation of Google Calendar events based on data retrieved from a Google Spreadsheet. It is designed to scan a specific sheet, identify rows based on a "support" (supp) column, and create events in a designated Google Calendar.

## Features

- Reads data from a Google Spreadsheet.
- Automatically creates calendar events in a Google Calendar.
- Prevents duplicate events by checking existing events in the time range.
- Adds event details such as event owner, instructor, and other metadata to the description.
- Includes pop-up reminders for created events.
  
## Prerequisites

1. **Google Calendar and Spreadsheet Access**:  
   Ensure you have the appropriate access to both the Google Calendar and Spreadsheet you intend to use.
   
2. **Google Apps Script Permissions**:  
   This script requires permission to access and manipulate your Google Calendar and Google Sheets.

## How it Works

1. **Calendar Setup**:  
   You must specify the Calendar ID where the events will be created by replacing `XXX` in the `calendarId` variable.

2. **Spreadsheet Setup**:  
   The script reads from a Google Spreadsheet, which you need to configure by replacing `XXX` in the following areas:
   
   - Spreadsheet URL
   - Sheet name
   - Column headers (like event name, date, time, owner, etc.)

3. **Event Creation**:  
   The script checks the `supp` column for a specific value (set by `targetName`). When a match is found, it creates a calendar event with the provided details (event name, date, time, etc.).

4. **Duplicate Prevention**:  
   Before creating a new event, the script checks the calendar for any existing events in the same time window to avoid duplicates.

## Configuration

You need to replace the following placeholders in the script before running it:

- `XXX` in `calendarId`: Replace with the Calendar ID where the events should be created.
- `XXX` in the spreadsheet URL in `SpreadsheetApp.openByUrl('XXX')`: Replace with the URL of your Google Spreadsheet.
- `XXX` in `const sheet = spreadsheet.getSheetByName('XXX')`: Replace with the name of the sheet containing the event data.
- `XXX` for target fields such as `targetName`, `eventName`, etc.: Customize these fields to match the event creation rules based on your spreadsheet structure.

## Spreadsheet Structure

The script expects a certain structure in the Google Spreadsheet, with specific headers to identify necessary data:

- **Názov**: Event Name
- **Čas**: Event Time (e.g., `09:00 - 10:00`)
- **Deň**: Event Date (in `dd.MM.yyyy` format)
- **čo treba**: Task description or action required
- **owner podujatia**: Event owner
- **lektor/ka**: Instructor
- **konto**: Email account of the person in charge
- **supp**: Support identifier to match with `targetName`

## How to Run the Script

1. Open the Google Apps Script editor by navigating to `Extensions > Apps Script` in your Google Sheets.
2. Copy and paste the script into the editor.
3. Replace the placeholder values (`XXX`) with your specific Calendar ID, Spreadsheet URL, and Sheet name.
4. Save and run the script.

## Logging

The script uses `Logger.log()` statements to provide detailed logging of the process, including which rows are processed, skipped, or result in event creation. You can view these logs in the Google Apps Script execution logs.

## Customization

- **Reminder Time**: Currently, a 10-minute pop-up reminder is set for every event. You can modify this by changing the value passed to `event.addPopupReminder(10)`.
- **Event Duration**: The event's duration is currently set to 40 minutes. You can adjust this by modifying the calculation of `eventEndTime`.

## Limitations

- The script assumes that the date format in the spreadsheet is `dd.MM.yyyy`. Ensure your date format aligns with this, or modify the date parsing logic accordingly.
- The time range for existing event detection is hardcoded as 20 minutes before and 40 minutes after the event start time.
