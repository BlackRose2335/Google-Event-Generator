# 📅 Calendar Support Event Creator

This script automates the creation of Google Calendar events based on data retrieved from a Google Spreadsheet. It is designed to scan a specific sheet, identify rows based on a "support" (`supp`) column, and create events in a designated Google Calendar.

---

## ✨ Features

- Reads data from a Google Spreadsheet.
- Automatically creates events in a Google Calendar.
- Prevents duplicate events by checking existing events in the time range.
- Adds event details (event owner, instructor, metadata) to the event description.
- Includes a pop-up reminder for the created events.

---

## 📋 Prerequisites

1. **Google Calendar and Spreadsheet Access**  
   Ensure you have the required access to both the Google Calendar and the Spreadsheet.

2. **Google Apps Script Permissions**  
   This script requires permissions to access and manage your Google Calendar and Sheets.

---

## ⚙️ How It Works

1. **Calendar Setup**  
   Replace `XXX` with your actual Calendar ID in the `calendarId` variable where events will be created.

2. **Spreadsheet Setup**  
   Configure the script by replacing `XXX` in these areas:
   - Spreadsheet URL
   - Sheet name
   - Column headers (like event name, date, time, owner, etc.)

3. **Event Creation**  
   The script scans the `supp` column, checks for a match with `targetName`, and creates an event in the calendar based on the row’s data.

4. **Duplicate Event Prevention**  
   Before creating an event, the script checks the calendar for existing events in the same time window to avoid duplication.

---

## 🛠️ Configuration

You need to replace the following placeholders in the script before running it:

- `XXX` for `calendarId`: The Calendar ID where events will be created.
- `XXX` for Spreadsheet URL in `SpreadsheetApp.openByUrl('XXX')`: Replace with the URL of your Google Spreadsheet.
- `XXX` for sheet name: Replace with the actual name of the sheet containing the event data.
- `XXX` in `targetName`, `eventName`, etc.: Customize these based on your spreadsheet structure.

---

## 📊 Spreadsheet Structure

The Google Spreadsheet must have the following columns with specific headers:

- **Názov**: Event Name
- **Čas**: Event Time (e.g., `09:00 - 10:00`)
- **Deň**: Event Date (in `dd.MM.yyyy` format)
- **čo treba**: Action or task description
- **owner podujatia**: Event owner
- **lektor/ka**: Instructor
- **konto**: Email account responsible for the event
- **supp**: Support identifier, which is compared to `targetName`

---

## 🚀 How to Run the Script

1. Open the Google Apps Script editor in your Google Sheets (`Extensions > Apps Script`).
2. Paste the provided script into the editor.
3. Replace all placeholder values (`XXX`) with actual Calendar IDs, Spreadsheet URL, and Sheet name.
4. Save and run the script.

---

## 📝 Logging

The script utilizes `Logger.log()` to track its operations. You can find information about rows processed, skipped, and events created in the Apps Script logs. These logs help with debugging and understanding the execution flow.

---

## 🔧 Customization

- **Reminder Time**  
   The script sets a 10-minute pop-up reminder for each event. You can adjust this by changing the value in `event.addPopupReminder(10)`.

- **Event Duration**  
   Events are set to last 40 minutes. You can modify this by adjusting the calculation of `eventEndTime`.

---

## ⚠️ Limitations

- **Date Format**  
   The script assumes that the date format in the spreadsheet is `dd.MM.yyyy`. Modify the date parsing logic if your spreadsheet uses a different format.

- **Event Duplication Check**  
   The script checks for existing events within a 20-minute window before the start and 40 minutes after the start of the event.

---
