function onEdit(e) {
  const sheetName = 'Goals'; // Updated name of the sheet with your goal list
  const targetSheetName = 'Completed Goals'; // The name of the sheet for completed goals
  const statusColumn = 1; // Column number of the "Status" column (A is 1)
  const goalsDeadlineColumnNum = 5; // Column number of the "Deadline" for "Goals" sheet

  // Check if the event object (e) is provided
  if (!e) {
    Logger.log('Event object is undefined. Exiting function.');
    return; // Exit if the function is manually run or if e is undefined
  }

  const range = e.range;
  const sheet = e.source.getSheetByName(sheetName);

  // Ensure the event is triggered in the correct sheet and column
  if (range.getSheet().getName() === sheetName && range.getColumn() === statusColumn) {
    const editedValue = range.getValue().trim(); // Trim any extra whitespace

    Logger.log('Edited sheet name: ' + range.getSheet().getName());
    Logger.log('Edited column number: ' + range.getColumn());
    Logger.log('Edited value: ' + editedValue);

    // Check if the status is 'Completed'
    if (editedValue.toLowerCase() === 'completed') {
      const row = range.getRow();
      const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

      Logger.log('Moving row data: ' + rowData);

      // Move the completed goal to the "Completed Goals" sheet
      const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
      targetSheet.appendRow(rowData);

      // Delete the completed goal from the original sheet
      sheet.deleteRow(row);
    }
  }

  // Sort by Deadline column (Column C)
  const rangeToSort = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  rangeToSort.sort({ column: goalsDeadlineColumnNum, ascending: true });

  Logger.log('Sorted the goals by deadline.');
}

function moveIncompleteTasks(e) {
 const targetSheetName = 'Failure Report'; 
  // Get the source sheet (where tasks are listed)
  var sourceSheet = e.source.getSheetByName('Goals');
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName); 

  // Get the data range (assuming data starts from row 2 to ignore headers)
  var data = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  
  // Loop through the data and check for 'Incomplete' status
  for (var i = data.length - 1; i >= 0; i--) {
    var status = data[i][0]; // "Status" is in the first column
    var goalName = data[i][1]; // "Goal" is in the 2nd column
    var description = data[i][3]; // "Description is in the 4th column"
    
    if (status.toLowerCase() === 'incomplete') {
      // Add Task Name and Description to the target sheet
      targetSheet.appendRow([goalName, description]);

      // Delete the row from the source sheet
      sourceSheet.deleteRow(i + 2); // +2 to account for zero-based index and header row

    }
  }
      failureReportValidation() 
}

function failureReportValidation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Failure Report'); 

  // Get all rows of the sheet (excluding the header)
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var data = dataRange.getValues();

  // Loop through each row to apply color formatting
  for (var i = 0; i < data.length; i++) {
    var row = i + 2; // Adjust for header row offset
    var columnAValue = data[i][0]; // Value in column A
    var columnDValue = data[i][3]; // Value in column D

    if (columnAValue !== "" && columnDValue === "") {
      sheet.getRange(row, 4).setBackground('red');
    } else {
      sheet.getRange(row, 4).setBackground(null);
    }
  }
}


