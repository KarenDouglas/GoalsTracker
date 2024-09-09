function createEventsFromSheet() {
  // Get the active spreadsheet and the sheet with the events data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var calendarId = 'primary'; // Use 'primary' to use your primary calendar or replace with a specific calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);

  // Get the date from cell B1
  var dateCell = sheet.getRange('B1').getValue();

  // Check if B1 has a valid date
  if (!dateCell || Object.prototype.toString.call(dateCell) !== '[object Date]') {
    Logger.log('B1 is empty or does not contain a valid date. Exiting function.');
    return; // Exit the function if B1 is empty or does not contain a valid date
  }

  // Get the data range (starting from row 5 to ignore headers in row 4)
  var dataRange = sheet.getRange(5, 1, sheet.getLastRow() - 4, 3); // Start from row 5
  var data = dataRange.getValues();

  // Loop through each row in the data
  for (var i = 0; i < data.length; i++) {
    var time = data[i][0]; // Time in column A
    var completed = data[i][1]; // Completed status in column B
    var eventTitle = data[i][2]; // Event title in column C

    // Skip rows with empty "Time" or "Event Title" cells
    if (!time || !eventTitle) {
      continue; // Skip this row
    }

    // If the event has not been created
    if (completed) {
      // Combine the date from B1 with the time from the current row
      var eventTime = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate(), time.getHours(), time.getMinutes());

      // Create the event in Google Calendar
      var event = calendar.createEvent(eventTitle, eventTime, new Date(eventTime.getTime() + 30 * 60 * 1000)); // 30 minutes duration

      // Mark the event as created by setting "Completed?" to TRUE
      sheet.getRange(i + 5, 2).setValue(true); // Row i+5 to match the spreadsheet row, Column B

      // Optionally, log the event details
      Logger.log('Created event: ' + eventTitle + ' at ' + eventTime);
    }
  }
}
