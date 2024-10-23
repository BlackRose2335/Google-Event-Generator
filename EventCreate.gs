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
  const eventNameIndex = headers.indexOf('Názov');
  const timeIndex = headers.indexOf('Čas');
  const dateIndex = headers.indexOf('Deň');
  const actionIndex = headers.indexOf('čo treba');
  const ownerIndex = headers.indexOf('owner podujatia');
  const instructorIndex = headers.indexOf('lektor/ka');
  const emailIndex = headers.indexOf('konto');
  const suppIndex = headers.indexOf('supp');

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
        const eventName = "XXX";  
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
        
        const isDuplicate = existingEvents.some(event => event.getTitle() === eventName);
        
        if (isDuplicate) {
          Logger.log(`Row ${i}: Event already exists: ${eventName} on ${eventStartTime}. Skipping creation.`);
          continue; 
        }
        
        const description = `\
        - Názov: ${row[eventNameIndex]}
        - čo treba: ${action}
        - owner podujatia: ${owner}
        - lektor/ka: ${instructor}
        - konto: ${email}
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
