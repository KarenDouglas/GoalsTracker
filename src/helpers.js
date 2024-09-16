function checkTasksAllExistInGoogleTasks() {
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
  
  // Get all data in the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Start checking from the second row (skip headers)
  for (let i = 1; i < data.length; i++) {
    const sheetTaskId = data[i][taskIdIndex];

    if (sheetTaskId && sheetTaskId.trim() !== '') {
      if (googleTaskIds.has(sheetTaskId)) {
        Logger.log(`Task ID ${sheetTaskId} exists in Google Tasks.`);
      } else {
        Logger.log(`Task ID ${sheetTaskId} does NOT exist in Google Tasks.`);
      }
    }
  }
}
const id = 'NDNHZzM5NkctcW9HQ1NBOQ'
function checkTaskExistInGoogleTasks() {
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
      } else {
        Logger.log(`Task ID ${sheetTaskId} does NOT exist in Google Tasks.`);
      }
    }
  
}
