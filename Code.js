// Constants for sheet names
const RECURRING_SHEET_NAME = "Recurring";
const TODAYS_TASKS_SHEET_NAME = "Active";

// Constants for column names
const COL_TASK_NAME = "Task";
const COL_DAYS_UNTIL_NEXT = "Days until next schedule";
const COL_LAST_ADDED_TIME = "Last added time";
const COL_DATE_ADDED = "Date Added";
const COL_COMPLETED = "Completed";
const COL_ACTIVE_ROW = "Active row"

// Constants for additional columns in Active sheet
const ADDITIONAL_COLUMNS = [COL_DATE_ADDED, COL_COMPLETED, COL_ACTIVE_ROW];

const MAIN_DATA_OFFSET = 4;

// Function to create a two-way lookup object for headers
function createHeaderLookup(headers) {
  var lookup = {};
  for (var i = 0; i < headers.length; i++) {
    lookup[headers[i]] = i;
    lookup[i] = headers[i];
  }
  return lookup;
}

function refreshMainFilter() {
  const sheetName = "Main";
  const col = 1; // column "A"

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const filter = sheet.getFilter();
  if (!filter) {
    return;
  }
  const r = filter.getRange();
  const criteria = filter.getColumnFilterCriteria(col).copy();
  filter.remove();
  r.createFilter().setColumnFilterCriteria(col, criteria);

}

function updateTodaysTasks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recurringSheet = ss.getSheetByName(RECURRING_SHEET_NAME);
  var todaysTasksSheet = ss.getSheetByName(TODAYS_TASKS_SHEET_NAME);
  
  var recurringHeaders = recurringSheet.getRange(1, 1, 1, recurringSheet.getLastColumn()).getValues()[0];
  var rLookup = createHeaderLookup(recurringHeaders);
  
  var todaysTasksHeaders = getOrCreateTodaysTasksHeaders(todaysTasksSheet, recurringHeaders);
  var tLookup = createHeaderLookup(todaysTasksHeaders);
  
  var data = recurringSheet.getDataRange().getValues();
  data.shift(); // Remove header row
  
  var today = new Date();
  var existingTasks = [];
  if (todaysTasksSheet.getLastRow() > 1) {
    var existingTasks = todaysTasksSheet.getRange(2, 1, todaysTasksSheet.getLastRow() - 1, todaysTasksSheet.getLastColumn()).getValues();
  }
  
  var tasksForToday = [];
  var lastAddedTimeUpdates = [];
  
  data.forEach(function(row, index) {
    var newTask = processTask(row, index, rLookup, tLookup, todaysTasksHeaders, existingTasks, today);
    if (newTask.task) tasksForToday.push(newTask.task);
    if (newTask.update) lastAddedTimeUpdates.push(newTask.update);
  });
  
  if (tasksForToday.length > 0) {
    todaysTasksSheet.getRange(todaysTasksSheet.getLastRow() + 1, 1, tasksForToday.length, todaysTasksHeaders.length).setValues(tasksForToday);
  }

  refreshMainFilter();
  
  lastAddedTimeUpdates.forEach(function(update) {
    recurringSheet.getRange(update[0], update[1]).setValue(update[2]);
  });
}

function getOrCreateTodaysTasksHeaders(sheet, recurringHeaders) {
  if (sheet.getLastRow() == 0) {
    var headers = recurringHeaders.concat(ADDITIONAL_COLUMNS);
    sheet.appendRow(headers);
    return headers;
  }
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function processTask(row, index, rLookup, tLookup, todaysTasksHeaders, existingTasks, today) {
  var taskName = row[rLookup[COL_TASK_NAME]];
  var daysUntilNext = row[rLookup[COL_DAYS_UNTIL_NEXT]];
  
  if (taskName === "") return {};
  if (daysUntilNext > 0) return {};
  
  var taskExists = existingTasks.some(function(task) {
    return task[tLookup[COL_TASK_NAME]] === taskName && !task[tLookup[COL_COMPLETED]];
  });
  
  if (taskExists) return {};
  
  var newTask = createNewTask(row, rLookup, tLookup, todaysTasksHeaders, today);
  var update = [index + 2, rLookup[COL_LAST_ADDED_TIME] + 1, today];
  
  return { task: newTask, update: update };
}

function createNewTask(row, rLookup, tLookup, todaysTasksHeaders, today) {
  var newTask = new Array(todaysTasksHeaders.length);
  
  todaysTasksHeaders.forEach(function(header, index) {
    if (header in rLookup) {
      newTask[index] = row[rLookup[header]];
    }
  });
  
  newTask[tLookup[COL_DATE_ADDED]] = today;
  newTask[tLookup[COL_COMPLETED]] = false;
  newTask[tLookup[COL_ACTIVE_ROW]] = "=ROW()";
  
  return newTask;
}

/*
TODO:
  Save datetime into Active!Completed column, and Recurring!Last completed time
  Activate automatic task additions daily
  Add one-off task interface

*/


// Function to acquire the semaphore
function acquireSemaphore() {
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(10000); // Try to acquire the lock for 10 seconds
  
  if (acquired) {
    var props = PropertiesService.getScriptProperties();
    var isProcessing = props.getProperty('isProcessing');
    
    if (isProcessing === 'true') {
      lock.releaseLock();
      return false;
    } else {
      props.setProperty('isProcessing', 'true');
      lock.releaseLock();
      return true;
    }
  }
  
  return false;
}

// Function to release the semaphore
function releaseSemaphore() {
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(10000);
  
  if (acquired) {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('isProcessing', 'false');
    lock.releaseLock();
  }
}

function onEdit(e) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var debugSheet = getOrCreateDebugSheet(ss);
  var sheet = e.source.getActiveSheet();
  
  logDebug(debugSheet, "onEdit triggered. Sheet: " + sheet.getName());
  
  // Check if the edit was made in the "Main" sheet
  if (sheet.getName() !== "Main") {
    logDebug(debugSheet, "Edit not in Main sheet. Refreshing Main view then exiting.");
    refreshMainFilter();
    return;
  }

    
  // Get the column numbers for "Complete", "Row", and "Reassign" columns
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var completeColNum = headers.indexOf("Complete") + 1;

  var reassignColNum = headers.indexOf("Reassign") + 1;


  if (completeColNum < e.range.getColumn() || completeColNum > e.range.getLastColumn()) {
    if (reassignColNum < e.range.getColumn() || reassignColNum > e.range.getLastColumn()) {
      logDebug(debugSheet, "Edit column in 'Main' isn't 'Complete' or 'Reassign'. Exit to avoid annoying refreshing.");
      return;
    } else {
      logDebug(debugSheet, "Edited column 'Reassign'");
    }
  } else {
    logDebug(debugSheet, "Edited column 'Complete'");
  }

  // If we're already processing, undo the edit and inform the user
  if (scriptProperties.getProperty('isScriptEditing') === 'true') {
    logDebug(debugSheet, "Edit attempted while processing. Undoing edit.");
    e.range.setValue(e.oldValue);
    //SpreadsheetApp.getUi().alert("Please wait. The sheet is currently being processed.");
    return;
  }

  // Increment the edit instance counter
  var editInstance = Number(scriptProperties.getProperty('editInstance') || 0) + 1;
  scriptProperties.setProperty('editInstance', editInstance.toString());
  logDebug(debugSheet, "Edit instance: " + editInstance);

  // Wait
  Utilities.sleep(3000);

  // Check if another edit has occurred
  var currentInstance = Number(scriptProperties.getProperty('editInstance') || 0);
  if (currentInstance !== editInstance) {
    logDebug(debugSheet, "Another edit occurred. Deferring to newer instance.");
    return;
  }

  // Set processing flag
  scriptProperties.setProperty('isScriptEditing', 'true');

  // Use a try-finally block to ensure the processing flag is always reset
  try {
    logDebug(debugSheet, "Starting processing.");    
    var rowColNum = headers.indexOf(COL_ACTIVE_ROW) + 1;
    var reprocessingColNum = headers.indexOf("Reprocessing") + 1;
    var recurringForeignKeyCol = headers.indexOf("Recurring key") + 1;

       
    // Hide "Complete" and "Reassign" columns
    sheet.hideColumns(completeColNum);
    sheet.hideColumns(reassignColNum);
    sheet.showColumns(reprocessingColNum);
    logDebug(debugSheet, "Hidden 'Complete' and 'Reassign' columns");
    
    logDebug(debugSheet, "Headers: Complete column: " + completeColNum + ", Row column: " + rowColNum + ", Reassign column: " + reassignColNum);
    
    // Function to get the current state of the "Complete" and "Reassign" columns
    function getColumnState() {
      var data = sheet.getRange(MAIN_DATA_OFFSET, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      return data.map(row => ({
        complete: row[completeColNum - 1],
        reassign: row[reassignColNum - 1]
      }));
    }
    
    // Get the state after the edit
    var currentState = getColumnState();
    
    // Process the state immediately
    var activeSheet = ss.getSheetByName("Active");
    var activeData = activeSheet.getDataRange().getValues();
    var activeHeaders = activeData[0];
    var completedColIndex = activeHeaders.indexOf("Completed");
    var rowColIndex = activeHeaders.indexOf(COL_ACTIVE_ROW);
    var ownerColIndex = activeHeaders.indexOf("Owner");

    var recurringSheet = ss.getSheetByName("Recurring");
    var recurringData = recurringSheet.getDataRange().getValues();
    var recurringHeaders = recurringData[0];
    var recurringKeyColIndex = recurringHeaders.indexOf("Recurring key");
    var recurringLastCompletedIndex = recurringHeaders.indexOf("Last completed time");

    
    logDebug(debugSheet, "Processing state. Headers: Completed column in Active: " + (completedColIndex + 1) + 
             ", Row column in Active: " + (rowColIndex + 1) + ", Owner column in Active: " + (ownerColIndex + 1));
    
    var changesCount = 0;

    currentState.reverse();

    currentState.forEach((state, rev_index) => {
      var index = currentState.length - rev_index - 1;
      if (state.complete || state.reassign) {
        changesCount++;
        var sourceRow = sheet.getRange(index + MAIN_DATA_OFFSET, rowColNum).getValue();
        logDebug(debugSheet, "Processing row " + (index + MAIN_DATA_OFFSET) + " in Main, sourceRow: " + sourceRow);
        
        if (state.complete) {
          // Clear the checkbox
          sheet.getRange(index + MAIN_DATA_OFFSET, completeColNum).setValue(false);
          logDebug(debugSheet, "Cleared checkbox in Main sheet row " + (index + MAIN_DATA_OFFSET));

          var recurringKey = sheet.getRange(index + MAIN_DATA_OFFSET, recurringForeignKeyCol).getValue();
          if (recurringKey) {
            logDebug(debugSheet, "Looking for a recurring task with key " + recurringKey + " in " + recurringData.length + " tasks");
            logDebug(debugSheet, JSON.stringify(recurringData));

            if(!recurringData.some(function(row, rindex) {
              if (row[recurringKeyColIndex] == recurringKey) {
                logDebug(debugSheet, "Found at row " + (rindex+1));
                logDebug(debugSheet, JSON.stringify(row));

                recurringSheet.getRange(rindex + 1, recurringLastCompletedIndex + 1).setValue(new Date());
                return true;
              }
              return false;
            })) {
              logDebug(debugSheet, "ERROR: didn't find a recurring task with key " + recurringKey);
            }
          }

          activeSheet.getRange(sourceRow, completedColIndex + 1).setValue(new Date());
          logDebug(debugSheet, "Updated Active sheet row " + (sourceRow) + " to completed");
        }
        
        if (state.reassign) {
          if (state.reassign.toLowerCase() === "unassign") {
            activeSheet.getRange(sourceRow, ownerColIndex + 1).setValue("");
            logDebug(debugSheet, "Unassigned owner in Active sheet row " + (sourceRow));
          } else {
            activeSheet.getRange(sourceRow, ownerColIndex + 1).setValue(state.reassign);
            logDebug(debugSheet, "Reassigned owner to '" + state.reassign + "' in Active sheet row " + (sourceRow));
          }
          
          // Clear the Reassign cell in the Main sheet
          sheet.getRange(index + MAIN_DATA_OFFSET, reassignColNum).setValue("");
          logDebug(debugSheet, "Cleared Reassign cell in Main sheet row " + (index + MAIN_DATA_OFFSET));
        }
      }
    });
    
    logDebug(debugSheet, "Processed " + changesCount + " changes");

    refreshMainFilter();
    
  } catch (error) {
    logDebug(debugSheet, "Error occurred: " + error.toString());
  } finally {    
    // Unhide "Complete" and "Reassign" columns
    sheet.showColumns(completeColNum);
    sheet.showColumns(reassignColNum);
    sheet.hideColumns(reprocessingColNum);

    logDebug(debugSheet, "Unhidden 'Complete' and 'Reassign' columns");

    // Reset processing flag
    scriptProperties.deleteProperty('isScriptEditing');
    logDebug(debugSheet, "Processing complete. Reset processing flag.");
  }
}

// Add this function to your script
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Reset Processing State', 'resetProcessingState')
    .addToUi();
}

function resetProcessingState() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('isScriptEditing');
  scriptProperties.deleteProperty('editInstance');
  SpreadsheetApp.getUi().alert('Processing state has been reset.');
}

function getOrCreateDebugSheet(ss) {
  var debugSheet = ss.getSheetByName("Debug Log");
  if (!debugSheet) {
    debugSheet = ss.insertSheet("Debug Log");
    debugSheet.appendRow(["Timestamp", "Message"]);
  }
  return debugSheet;
}

function logDebug(sheet, message) {
  sheet.appendRow([new Date(), message]);
}