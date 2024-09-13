function addTasksToGoogleTasks() {
  Logger.log('inserted from IDE')
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const ui = SpreadsheetApp.getUi(); // Get the UI object for displaying alerts
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const taskIndex = headers.indexOf('Task');
  const notesIndex = headers.indexOf('Notes');
  const dueDateIndex = headers.indexOf('Due Date');
  const completedIndex = headers.indexOf('Completed?');
  const taskIdIndex = headers.indexOf('TaskID');
  const saveIndex = headers.indexOf('Save');

  // Get the range of all data in the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();


    
  // Loop through each row of data (skip the header row)
  for (let i = 1; i < data.length; i++) { // Loop only up to the last row with values
    let row = data[i]
    const isRowEmpty = row.every(cell => cell === "");
    if(isRowEmpty){
        Logger.log('This row is empty');
      return;
    }
  
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
}





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
        cell.setBackground('pale red'); // Past due
      } else if (dueDate.toDateString() === today.toDateString()) {
        cell.setBackground('pale yellow'); // Due today
      } else if (dueDate.toDateString() === tomorrow.toDateString()) {
        cell.setBackground('pale green'); // Due tomorrow
      } else {
        cell.setBackground('white');
      }
    }
  }
}
