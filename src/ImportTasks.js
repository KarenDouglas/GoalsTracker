function syncTasksFromGoogleTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const ui = SpreadsheetApp.getUi();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const taskIndex = headers.indexOf('Task');
  const notesIndex = headers.indexOf('Notes');
  const dueDateIndex = headers.indexOf('Due Date');
  const completedIndex = headers.indexOf('Completed?');
  const taskIdIndex = headers.indexOf('TaskID');
  const saveIndex = headers.indexOf('Save');

  // Check if all required columns are present
  if (
    taskIndex === -1 ||
    notesIndex === -1 ||
    dueDateIndex === -1 ||
    completedIndex === -1 ||
    taskIdIndex === -1 ||
    saveIndex === -1
  ) {
    ui.alert('One or more required columns are missing in the sheet.');
    return;
  }

  // Get existing task IDs from the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const existingTaskIds = {};

  for (let i = 1; i < data.length; i++) {
    // Start from row 2 to skip headers
    const row = data[i];
    const taskId = row[taskIdIndex];
    if (taskId && taskId.trim() !== '') {
      existingTaskIds[taskId.trim()] = true;
    }
  }

  // Fetch tasks from Google Tasks
  const taskListId = '@default';
  const tasks = Tasks.Tasks.list(taskListId).items || [];

  let newTasksCount = 0;
  const newRows = []; // Collect new rows to add

  for (const task of tasks) {
    if (!existingTaskIds[task.id]) {
      // Create a new row array matching the number of headers
      const newRow = new Array(headers.length).fill('');

      newRow[taskIndex] = task.title || '';
      newRow[notesIndex] = task.notes || '';
      newRow[dueDateIndex] = task.due ? new Date(task.due) : '';
      newRow[completedIndex] = task.status === 'completed';
      newRow[taskIdIndex] = task.id;
      newRow[saveIndex] = false; // Set 'Save' to false by default

      newRows.push(newRow);
      newTasksCount++;
    }
  }

  // Find the last row with data in the 'TaskID' column
  let lastDataRow = 1; // Start from the header row
  for (let i = data.length - 1; i >= 1; i--) {
    // Start from the bottom and go up
    if (data[i][taskIdIndex] && data[i][taskIdIndex].toString().trim() !== '') {
      lastDataRow = i + 1; // Rows are 1-indexed
      break;
    }
  }

  // Insert new rows after the last data row in 'TaskID' column
  if (newRows.length > 0) {
    sheet
      .getRange(lastDataRow + 1, 1, newRows.length, headers.length)
      .setValues(newRows);
  }

  ui.alert(`Synced ${newTasksCount} new tasks from Google Tasks to the sheet.`);
}
