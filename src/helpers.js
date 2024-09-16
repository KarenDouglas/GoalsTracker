function checkTasksAllExistInGoogleTasks() {
    // Get the active spreadsheet and the sheet by name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const taskIdIndex = headers.indexOf('TaskID');
    const taskIndex = headers.indexOf('Task');
  const taskListId = '@default'; // '@default' for default task list

  // Retrieve all tasks from Google Tasks
  const tasks = Tasks.Tasks.list(taskListId).items;
  
  if (!tasks) {
    Logger.log("No tasks found in Google Tasks.");
    return;
  }

  // Create a Set of Task IDs from Google Tasks for quick lookup
  const googleTaskIds = new Set(tasks.map(task => task.id));
  
  // Get all data in the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Start checking from the second row (skip headers)
  for (let i = 1; i < data.length; i++) {
    const sheetTaskId = data[i][taskIdIndex].trim();
    const taskTitle = data[i][taskIndex]

    if (sheetTaskId && sheetTaskId.trim() !== '') {
      if (googleTaskIds.has(sheetTaskId)) {
        Logger.log(`Task ID ${sheetTaskId} | ${taskTitle} exists in Google Tasks.`);
      } else {
        Logger.log(`Task ID ${sheetTaskId} | ${taskTitle} does NOT exist in Google Tasks.`);
      }
    }
  }
}
function isAlertConfirmed(alertTitle, alertStr) {
  // Get the UI object for the current spreadsheet
  const ui = SpreadsheetApp.getUi();

  // Display an alert with Yes and No buttons
  const response = ui.alert(
      alertTitle,
      alertStr,
    ui.ButtonSet.YES_NO
  );

  // Handle the user's response
  if (response == ui.Button.YES) {
    // User clicked "Yes"
      return true;
  } else if (response == ui.Button.NO) {
    // User clicked "No"
     return false;
  }
}
function isTaskIdInExistInGoogleTasks(id) {
    // Get the active spreadsheet and the sheet by name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const taskIdIndex = headers.indexOf('TaskID');
  const taskListId = '@default'; // '@default' for default task list

  // Retrieve all tasks from Google Tasks
  const tasks = Tasks.Tasks.list(taskListId).items;
  
  if (!tasks) {
    Logger.log("No tasks found in Google Tasks.");
    return;
  }

  // Create a Set of Task IDs from Google Tasks for quick lookup
  const googleTaskIds = new Set(tasks.map(task => task.id));

    const sheetTaskId = id

    if (sheetTaskId && sheetTaskId.trim() !== '') {
      if (googleTaskIds.has(sheetTaskId)) {
        Logger.log(`Task ID ${sheetTaskId} exists in Google Tasks.`);
        return true;
      } else {
        Logger.log(`Task ID ${sheetTaskId} does NOT exist in Google Tasks.`);
        return false
      }
    }
  
}

function getLastRow (){
    // Find the last row with data in the 'TaskID' column
  let lastDataRow = 1; // Start from the header row
  for (let i = data.length - 1; i >= 1; i--) {
    // Start from the bottom and go up
    if (data[i][taskIdIndex] && data[i][taskIdIndex].toString().trim() !== '') {
      lastDataRow = i + 1; // Rows are 1-indexed
      break;
    }
  }
  return lastDataRow;

}