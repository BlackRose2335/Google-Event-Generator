# Google Calendar Event Automation Script

This Google Apps Script automates the process of creating Google Calendar events based on a shared Google Spreadsheet. Events will be created based on specific conditions from the spreadsheet (e.g., when the name "XXX" is present in a specific column). It handles parsing event details like time, date, and event description from the spreadsheet and ensures that duplicate events are not created.

## Features

- **Automated Event Creation**: Automatically creates events on a specified Google Calendar based on data from a shared Google Spreadsheet.
- **Customizable Event Times**: Events are set to start 20 minutes before the meeting time and last for 40 minutes.
- **Prevents Duplicates**: The script checks if an event already exists before creating it, preventing duplicate entries.
- **Custom Event Details**: The script populates the event description with details like the meeting name, owner, instructor, and contact email.

## Requirements

1. **Google Calendar API**: The script uses Google Calendar API to create events. Ensure that you have proper permissions to create events on the specified calendar.
2. **Google Spreadsheet**: The source of event data must be a shared Google Spreadsheet with specific columns for event name, time, date, owner, instructor, email, and support information.

### Spreadsheet Structure

The spreadsheet should contain the following columns in its second row (the headers):
- **Deň**: The event date in the format `DD.MM.YYYY`.
- **Čas**: The time range of the event in the format `HH:mm - HH:mm`.
- **Názov**: The event name.
- **owner**: The owner of the meeting.
- **lektor/ka**: The instructor of the event.
- **konto**: The email address related to the event.
- **supp**: The support column, which should contain the value "XXX" to trigger event creation.

## How It Works

1. The script reads the data from the specified Google Spreadsheet.
2. For each row, it checks if the `supp` column contains the name "XXX"
3. If the condition is met, the script parses the `Deň` (date) and `Čas` (time) fields to create an event.
4. The event is created on the specified Google Calendar 20 minutes before the actual meeting time, and it lasts for 40 minutes.
5. If an event with the same details already exists on the calendar, the script skips creating a new event to avoid duplication.

## Usage

### 1. Setting up the Script

1. Open [Google Apps Script](https://script.google.com/).
2. Copy and paste the provided script into a new project.

### 2. Spreadsheet and Calendar Setup

1. **Spreadsheet**: 
   - Ensure that your spreadsheet has the required columns as described above.
   - Replace the `Spreadsheet URL` in the script with the URL of your Google Spreadsheet.

2. **Calendar**:
   - Replace the `calendarId` in the script with the ID of the Google Calendar where the events will be created.
   - You can find the Calendar ID in Google Calendar under Calendar Settings -> Integrate Calendar -> Calendar ID.

### 3. Running the Script

1. Open the Google Apps Script Editor.
2. Run the `createSupportEvent` function.
3. Grant the necessary permissions when prompted (e.g., access to Calendar and Spreadsheet).
4. Check your Google Calendar to see the newly created events.

### 4. Debugging and Logs

- The script logs detailed information about the process in the Google Apps Script Logger.
- Open the Apps Script Editor, then go to `View > Logs` to see real-time logging, including:
  - Rows being processed.
  - Errors or issues during execution.
  - Events that were successfully created or skipped due to duplication.

## Example

### Example Spreadsheet Row

| Deň        | Čas        | Názov         | owner     | lektor/ka | konto             | supp   |
|------------|------------|---------------|-----------|-----------|-------------------|--------|
| 21.10.2024 | 17:00 - 20:00 | Event Title | John Doe  | Jane Doe  | john@example.com  | XXX |

### Example Created Event

- **Event Name**: Support Session
- **Start Time**: 16:40 (20 minutes before the meeting starts)
- **End Time**: 17:20 (lasts 40 minutes)
- **Description**:
    - Názov: Event Title
    - owner podujatia: John Doe
    - lektor/ka: Jane Doe
    - konto: john@example.com


## Contributing

If you'd like to contribute to this project or report issues, please submit a pull request or open an issue on GitHub.

---

### Notes

1. **Calendar Permissions**: Make sure you have sufficient access to the calendar where you want to create events.
2. **Duplicate Events**: The script avoids creating duplicate events by checking for existing events within the same time frame and with the same title.
