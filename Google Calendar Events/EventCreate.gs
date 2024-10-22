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
  const ownerIndex = headers.indexOf('owner');      
  const instructorIndex = headers.indexOf('lektor/ka'); 
  const emailIndex = headers.indexOf('konto');     
  const suppIndex = headers.indexOf('supp');        

  const targetName = 'XXX';  
  let eventCount = 0; 
  
  for (let i = 2; i < data.length; i++) { 
    const row = data[i];

    Logger.log(`Processing Row ${i}: ${JSON.stringify(row)}`);
    
    if (suppIndex >= 0 && row[suppIndex] !== undefined) {
      Logger.log(`Row ${i}: supp = ${row[suppIndex]}`);
      
      if (row[suppIndex].trim() === targetName) { 
        const eventName = "XXX";
        const meetingDetails = row[timeIndex];       
        const owner = row[ownerIndex];               
        const instructor = row[instructorIndex];    
        const email = row[emailIndex];              
        
        let meetingDate = row[dateIndex];
        if (meetingDate instanceof Date) {
          meetingDate = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), "dd.MM.yyyy");
        } else {
          Logger.log(`Row ${i}: meetingDate is not a Date object: ${meetingDate}`);
        }

        const [startTimeStr, endTimeStr] = meetingDetails.split(" - ");
        
        const [day, month, year] = meetingDate.split(".").map(Number);
        const startHour = parseInt(startTimeStr.split(":")[0], 10);
        const startMinute = parseInt(startTimeStr.split(":")[1], 10);
        
        const startTime = new Date(year, month - 1, day, startHour, startMinute); 
        
        const eventStartTime = new Date(startTime.getTime() - 20 * 60 * 1000); 
        
        const eventEndTime = new Date(eventStartTime.getTime() + 40 * 60 * 1000);
        
        const description = `
        - Názov: ${row[eventNameIndex]}
        - owner podujatia: ${owner}
        - lektor/ka: ${instructor}
        - konto: ${email}
        `;
        
        const event = calendar.createEvent(eventName, eventStartTime, eventEndTime, {
          description: description.trim()  // Remove extra spaces
        });
        
        event.addPopupReminder(10);
        event.addPopupReminder(20);
        event.addPopupReminder(30);
        
        Logger.log(`Created event: ${eventName} on ${eventStartTime} for ${targetName}`);
        eventCount++;
      }
    } else {
      Logger.log(`Row ${i} does not have a valid supp index or is undefined.`);
    }
  }
  Logger.log(`Total events created: ${eventCount}`);
}
