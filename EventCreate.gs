/**
 * MIT License
 * 
 * Copyright (c) 2024 BlackRose2335 (https://github.com/BlackRose2335)
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * 1. The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * 2. THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
 * OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

function createSupportEvent() {
  const calendarId = 'XXX';
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    Logger.log(`Error: Calendar with ID ${calendarId} not found or access denied.`);
    return;
  }
  
  const spreadsheet = SpreadsheetApp.openByUrl('XXX');
  const sheet = spreadsheet.getSheetByName('XXX');
  const data = sheet.getDataRange().getValues();
  
  const headers = data[1];
  const eventName = "XXX";
  const eventNameIndex = headers.indexOf('Name');
  const timeIndex = headers.indexOf('Time');
  const dateIndex = headers.indexOf('Date');
  const actionIndex = headers.indexOf('Action');
  const ownerIndex = headers.indexOf('Owner');
  const instructorIndex = headers.indexOf('Instructor');
  const emailIndex = headers.indexOf('E-mail');
  const suppIndex = headers.indexOf('Support');

  const targetName = 'XXX';
  let eventCount = 0; 
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    Logger.log(`Processing Row ${i}: ${JSON.stringify(row)}`);

    // Ensure row is valid
    if (!row || row.length === 0) {
      Logger.log(`Row ${i} is empty or undefined. Skipping...`);
      continue;
    }

    if (suppIndex >= 0 && row[suppIndex] !== undefined) {
      Logger.log(`Row ${i}: supp = '${row[suppIndex]}'`);

      if (row[suppIndex].trim() === targetName) {
        const meetingDetails = row[timeIndex];
        const action = row[actionIndex];
        const owner = row[ownerIndex];               
        const instructor = row[instructorIndex];     
        const email = row[emailIndex];               
        
        // Log the owner for debugging
        Logger.log(`Row ${i}: owner = '${owner}'`);

        // Ensure meetingDate is handled as a string
        let meetingDate = row[dateIndex];
        Logger.log(`Row ${i}: raw meetingDate = '${meetingDate}'`);
        
        if (meetingDate instanceof Date) {
          meetingDate = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), "dd.MM.yyyy"); 
        } else if (typeof meetingDate !== "string") {
          Logger.log(`Row ${i}: meetingDate is not a string or Date object: ${meetingDate}`);
          continue; 
        }

        Logger.log(`Parsed meetingDate: ${meetingDate}`); 

        const [startTimeStr, endTimeStr] = meetingDetails.split(" - ");
        Logger.log(`Row ${i}: startTimeStr = '${startTimeStr}', endTimeStr = '${endTimeStr}'`);
        
        const [day, month, year] = meetingDate.split(".").map(Number);
        const startHour = parseInt(startTimeStr.split(":")[0], 10);
        const startMinute = parseInt(startTimeStr.split(":")[1], 10);

        const startTime = new Date(year, month - 1, day, startHour, startMinute); 
        const eventStartTime = new Date(startTime.getTime() - 20 * 60 * 1000);
        const eventEndTime = new Date(eventStartTime.getTime() + 40 * 60 * 1000);
        
        Logger.log(`Row ${i}: Checking for existing events between ${eventStartTime} and ${eventEndTime}`);
        const existingEvents = calendar.getEvents(eventStartTime, eventEndTime);
        
        const isDuplicate = existingEvents.some(event => {
          const description = event.getDescription();
          return description.includes(row[eventNameIndex]); // Check if Názov is in the event description
        });
        
        if (isDuplicate) {
          Logger.log(`Row ${i}: Event with Názov '${row[eventNameIndex]}' already exists. Skipping creation.`);
          continue; 
        }

        const description = `\
        - Name: ${row[eventNameIndex]}
        - Action: ${action}
        - Meeting owner: ${owner}
        - Instructor: ${instructor}
        - E-mail: ${email}
        `;
        
        const event = calendar.createEvent(eventName, eventStartTime, eventEndTime, {
          description: description.trim()
        });
        
        event.addPopupReminder(10);
        
        Logger.log(`Row ${i}: Created event: ${eventName} on ${eventStartTime} for ${targetName}`);
        eventCount++;
      } else {
        Logger.log(`Row ${i}: supp does not match target name. Expected '${targetName}', found '${row[suppIndex].trim()}'`);
      }
    } else {
      Logger.log(`Row ${i} does not have a valid supp index or is undefined.`);
    }
  }
  
  Logger.log(`Total events created: ${eventCount}`);
}
