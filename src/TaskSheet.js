  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const ui = SpreadsheetApp.getUi(); // Get the UI object for displaying alerts
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const taskListId = '@default'; 

  const taskIndex = headers.indexOf('Task');
  const notesIndex = headers.indexOf('Notes');
  const dueDateIndex = headers.indexOf('Due Date');
  const completedIndex = headers.indexOf('Completed?');
  const taskIdIndex = headers.indexOf('TaskID');
  const saveIndex = headers.indexOf('Save');
  const deleteIndex = headers.indexOf('Delete?');

  // Get the range of all data in the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

function addTasksToGoogleTasks() {
  // Loop through each row of data (skip the header row)
  for (let i = 1; i < data.length; i++) { // Loop only up to the last row with values
    let row = data[i]
      if (row[saveIndex] === true) { // Check if "Save" is marked as true
          // makes sure user enters a task value 
          if(row[taskIndex] === ""){
            ui.alert(`row ${i+1} must have a value entered in order to save`);
               // Reset "Save" checkbox after processing
              sheet.getRange(i + 1, saveIndex + 1).setValue(false);
             continue;
          }
        let taskId = row[taskIdIndex];
        let task = {
          title: row[taskIndex],
          notes: row[notesIndex],
          due: row[dueDateIndex] ? new Date(data[i][dueDateIndex]).toISOString() : null,
          status: row[completedIndex] === true ? 'completed' : 'needsAction'
        };

        try {
          if (taskId && taskId.trim() !== '') {
            task.id = taskId.trim(); // Assign task.id before updating
            Logger.log(`taskid before update: ${task.id}`);
            Tasks.Tasks.update(task, '@default', task.id);
            Logger.log(`taskid after update: ${task.id}`);
            Logger.log(`Updated task: ${task.title} with ID: ${task.id}`);
            ui.alert(`${task.title} has been updated: ${task.id}`);
          } else {
            const newTask = Tasks.Tasks.insert(task, '@default');
            sheet.getRange(i + 1, taskIdIndex + 1).setValue(newTask.id);
            Logger.log(`Created task: ${task.title} with new ID: ${newTask.id}`);
            ui.alert(`${task.title} has been created: ${newTask.id}`);
          }
        } catch (e) {
          Logger.log(`Error processing task ${task.title} with ID ${taskId}: ${e.message}`);
        }

        // Reset "Save" checkbox after processing
        sheet.getRange(i + 1, saveIndex + 1).setValue(false);
      }   
  }
 syncTasksFromGoogleTasks()
}

// Deletes Tasks from Google Sheets and update Google Tasks
function deleteTask(){

// loops through sheet
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    const taskId = row[taskIdIndex].trim();
    const taskTitle = row[taskIndex]

    Logger.log(`${row[deleteIndex]}`)
    // Checks if row is set to be deleted
      if (row[deleteIndex] === true) {
          // checks if sheet taskID matches a taskID in Google Tasks
          if(isTaskIdInExistInGoogleTasks(taskId)) {
              Logger.log(`${taskId} is set to be deleted`)
              // check if user confirms delete request
            if(isAlertConfirmed('Delete Task', `Are you sure you want to delete ${taskTitle}?`)){
              // deletes task from sheet and google tasks
              try {
                // Delete task from Google Tasks
                Tasks.Tasks.remove(taskListId, taskId);
                Logger.log(`Task ID ${taskId} deleted from Google Tasks.`);

                // Delete the row from the sheet after deleting from Google Tasks
                sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
                Logger.log(`Row ${i + 1} deleted from Google Sheets.`);

                // Adjust the loop index since the row count has changed after deletion
                i--; // Decrement to account for the row shift after deletion
              } catch (e) {
                Logger.log(`Failed to delete Task ID ${taskId} from Google Tasks: ${e.message}`);
                ui.alert(`Error`, `${e}`)
              }
              return;

            }else{
              // if users denies deletion  request from alert
              sheet.getRange(i + 1, deleteIndex + 1).setValue(false);
              return;
            }
          }else{
            // if taskID doesn't exist error alert sent
            ui.alert(`${taskTitle} can not be deleted`)
            sheet.getRange(i + 1, deleteIndex + 1).setValue(false);
            return
          }
      }
  }
}

//Syncs deleted data from Google Tasks and updates Google Sheets
function syncDeletedTasksFromGoogleTasks() {
  const lastRow = getLastRow();

  // loops through sheet data
  for (let i = 1; i < lastRow; i++) {
    const sheetTaskId = data[i][taskIdIndex]
    if(!isTaskIdInExistInGoogleTasks(sheetTaskId)){
       sheet.deleteRow(i+1)
       ui.alert(`${data[i][taskIndex]} is being deleted`)
      Logger.log(` ${i} DELETED task ${data[i][taskIndex]} with ID ${sheetTaskId} from Google Sheet`);
    i--
    }else{
      Logger.log(`${sheetTaskId} remains`)
    }
  }

};

function formatSheet(sheet) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const dueDateIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('Due Date');

  for (let i = 1; i < data.length; i++) {
    let dueDate = new Date(data[i][dueDateIndex]);
    let cell = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());

    if (isNaN(dueDate.getTime())) {
      cell.setBackground('white');
    } else {
      let today = new Date();
      let tomorrow = new Date(today);
      tomorrow.setDate(today.getDate() + 1);

      if (dueDate < today) {
        cell.setBackground('red'); // Past due
      } else if (dueDate.toDateString() === today.toDateString()) {
        cell.setBackground('yellow'); // Due today
      } else if (dueDate.toDateString() === tomorrow.toDateString()) {
        cell.setBackground('green'); // Due tomorrow
      } else {
        cell.setBackground('white');
      }
    }
  }
}
