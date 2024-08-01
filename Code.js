// Constants for sheet names
const RECURRING_SHEET_NAME = "Recurring";
const TODAYS_TASKS_SHEET_NAME = "Active";
const DEBUG_SHEET_NAME = "Debug Log";

// Constants for column names in Recurring sheet
const COL_TASK_NAME = "Task";
const COL_SCHEDULE_FROM_COMPLETION = "Schedule from completion";
const COL_NEXT_SCHEDULED_DATE = "Next scheduled date";
const COL_RECURRING_KEY = "Recurring key";
const COL_DAYS = "Days";

// Constants for column names in Active sheet
const COL_DATE_ADDED = "Date Added";
const COL_COMPLETED = "Completed";
const COL_COMPLETED_DATE = "Completed Date";
const COL_ACTIVE_ROW = "Active row";
const COL_OWNER = "Owner";
const COL_URGENCY = "Urgency";
const ACT_RECURRING_KEY = COL_RECURRING_KEY;
const RESORT_CHECKBOX_CELL = "K2";

// Constants for additional columns in Active sheet
const ADDITIONAL_COLUMNS = [COL_DATE_ADDED, COL_COMPLETED, COL_URGENCY, COL_COMPLETED_DATE, COL_ACTIVE_ROW];

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

function findBadKey(data, rLookup) {
    var recurringKeys = [];
    var badKey;
    if (data.every(function (row) {
        var key = row[rLookup[COL_RECURRING_KEY]];
        if (key === "") {
            badKey = key;
            return false;
        }
        if (recurringKeys.includes(key)) {
            badKey = key;
            return false;
        } else {
            recurringKeys.push(key);
            return true;
        }
    })) {
        return false;
    } else {
        return badKey;
    }
}

function updateTodaysTasks() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var recurringSheet = ss.getSheetByName(RECURRING_SHEET_NAME);
    var todaysTasksSheet = ss.getSheetByName(TODAYS_TASKS_SHEET_NAME);

    updateActiveSheet();

    var recurringHeaders = recurringSheet.getRange(1, 1, 1, recurringSheet.getLastColumn()).getValues()[0];
    // Cut out any headers after an empty cell ("")
    recurringHeaders = recurringHeaders.slice(0, recurringHeaders.indexOf(""));
    var rLookup = createHeaderLookup(recurringHeaders);

    var todaysTasksHeaders = getOrCreateTodaysTasksHeaders(todaysTasksSheet, recurringHeaders);
    var tLookup = createHeaderLookup(todaysTasksHeaders);

    var data = recurringSheet.getDataRange().getValues();
    data.shift(); // Remove header row

    var badKey = findBadKey(data, rLookup);

    if (badKey) {
        throw new Error("Duplicate or missing recurring keys found: " + badKey);
    }

    var today = new Date();
    var existingTasks = [];
    if (todaysTasksSheet.getLastRow() > 1) {
        var existingTasks = todaysTasksSheet.getRange(2, 1, todaysTasksSheet.getLastRow() - 1, todaysTasksSheet.getLastColumn()).getValues();
    }

    var tasksForToday = [];
    var lastAddedTimeUpdates = [];

    data.forEach(function (row, index) {
        var newTask = processTask(row, index, rLookup, tLookup, todaysTasksHeaders, existingTasks, today);
        if (newTask.task) tasksForToday.push(newTask.task);
        if (newTask.update) lastAddedTimeUpdates.push(newTask.update);
    });

    if (tasksForToday.length > 0) {
        todaysTasksSheet.getRange(todaysTasksSheet.getLastRow() + 1, 1, tasksForToday.length, todaysTasksHeaders.length).setValues(tasksForToday);
    }

    updateActiveSheet();

    lastAddedTimeUpdates.forEach(function (update) {
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
    var nextScheduledDate = row[rLookup[COL_NEXT_SCHEDULED_DATE]] || today;
    var days = row[rLookup[COL_DAYS]];
    var recurringKey = row[rLookup[COL_RECURRING_KEY]];

    if (taskName === "") return {};
    var addAfterTime = new Date(new Date(nextScheduledDate).getTime() - .2 * 24 * 60 * 60 * 1000);
    if (addAfterTime > today) return {}; // If the next scheduled date is more than .2 days in the future, don't add it

    var taskExists = existingTasks.some(function (task) {
        return task[tLookup[COL_RECURRING_KEY]] == recurringKey && (!task[tLookup[COL_COMPLETED]] || task[tLookup[COL_COMPLETED]] === "false");
    });

    if (taskExists) return {};

    var newTask = createNewTask(row, rLookup, tLookup, todaysTasksHeaders, today);
    var newNextScheduledDate = row[rLookup[COL_SCHEDULE_FROM_COMPLETION]] ? "" : new Date(nextScheduledDate.getTime() + days * 24 * 60 * 60 * 1000);

    // For daily tasks, if the next scheduled date is still in the past, skip forward day by day until it's in the future.
    while (days < 2 && newNextScheduledDate && newNextScheduledDate < today) {
      newNextScheduledDate = new Date(newNextScheduledDate.getTime() + days * 24 * 60 * 60 * 1000);
    }
    var update = [index + 2, rLookup[COL_NEXT_SCHEDULED_DATE] + 1, newNextScheduledDate];

    return { task: newTask, update: update };
}

function createNewTask(row, rLookup, tLookup, todaysTasksHeaders, today) {
    var newTask = new Array(todaysTasksHeaders.length);

    todaysTasksHeaders.forEach(function (header, index) {
        if (header in rLookup) {
            newTask[index] = row[rLookup[header]];
        }
    });

    newTask[tLookup[COL_DATE_ADDED]] = today;
    newTask[tLookup[COL_COMPLETED]] = false;
    newTask[tLookup[COL_ACTIVE_ROW]] = "=ROW()";

    return newTask;
}


// Function that creates a one-off task with name taken from a cell and adds it to the "Active" sheet
function addOneOffTask(taskName, owner) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (taskName === "") {
        return false;
    }

    var activeSheet = ss.getSheetByName(TODAYS_TASKS_SHEET_NAME);
    var activeData = activeSheet.getDataRange().getValues();
    var activeHeaders = activeData[0];
    var tLookup = createHeaderLookup(activeHeaders);

    var newTask = new Array(activeHeaders.length);
    newTask[tLookup[COL_TASK_NAME]] = taskName;
    newTask[tLookup[COL_OWNER]] = owner;
    newTask[tLookup[COL_DATE_ADDED]] = new Date();
    newTask[tLookup[COL_COMPLETED]] = false;
    newTask[tLookup[COL_ACTIVE_ROW]] = "=ROW()";

    logDebug(getOrCreateDebugSheet(ss), "Adding one-off task: " + newTask[tLookup[COL_TASK_NAME]]);

    // Append the new task to the "Active" sheet
    activeSheet.appendRow(newTask);

    return true;
}

function onEdit(e) {
    var sheet = e.source.getActiveSheet();

    if (sheet.getName() != TODAYS_TASKS_SHEET_NAME) return;
    if (sheet.getSelection().getCurrentCell().getA1Notation() != RESORT_CHECKBOX_CELL) return;

    var cell = sheet.getRange(RESORT_CHECKBOX_CELL);

    if (!cell.getValue()) return;

    cell.setValue(false);

    updateActiveSheet();

    var debugSheet = getOrCreateDebugSheet(SpreadsheetApp.getActiveSpreadsheet());
    logDebug(debugSheet, "onEdit resort checkbox clicked");

}

function updateActiveSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = getOrCreateDebugSheet(ss);
    var activeSheet = ss.getSheetByName(TODAYS_TASKS_SHEET_NAME);
    var activeData = activeSheet.getDataRange().getValues();
    var activeHeaders = activeData[0];
    var tLookup = createHeaderLookup(activeHeaders);
    var rLookup;
    var recurringSheet;
    var recurringData;

    var today = new Date();

    // If the "Task" column has text in row 2, add it as a one-off task
    var newTask = activeSheet.getRange("A2").getValue();
    if (newTask !== "") {
        var owner = activeSheet.getRange("B2").getValue();
        if (owner === "") {
            owner = "Shared";
        }
        addOneOffTask(newTask, owner);

        // Clear the cells
        activeSheet.getRange("A2").setValue("");
        activeSheet.getRange("B2").setValue("");
    }

    // Find "Completed" TRUE columns that don't have a "Completed Date" and add one
    activeData.forEach(function (row, index) {
        if (index < 3) return; // Skip header, row 2, and row 3. Data starts at row 4
        if (row[tLookup[COL_COMPLETED]] === true && row[tLookup[COL_COMPLETED_DATE]] === "") {
            activeSheet.getRange(index + 1, tLookup[COL_COMPLETED_DATE] + 1).setValue(today);

            // Update the "Next scheduled date" in "Recurring" if "Schedule from completion" is TRUE
            if (row[tLookup[COL_SCHEDULE_FROM_COMPLETION]] === true) {
                var recurringKey = row[tLookup[ACT_RECURRING_KEY]];
                if (!rLookup) {
                    recurringSheet = ss.getSheetByName(RECURRING_SHEET_NAME);
                    recurringData = recurringSheet.getDataRange().getValues();
                    var rHeaders = recurringData[0];
                    rLookup = createHeaderLookup(rHeaders);
                }

                // Find the index of the matching recurring row
                var recurringRowIndex = 0;
                var recurringRow;
                for (var i = 1; i < recurringData.length; i++) {
                    if (recurringData[i][rLookup[COL_RECURRING_KEY]] == recurringKey) {
                        recurringRowIndex = i;
                        recurringRow = recurringData[i];
                        break;
                    }
                }
                if (recurringRowIndex > 0) {
                    var days = recurringRow[rLookup[COL_DAYS]];
                    var newNextScheduledDate = new Date(today.getTime() + days * 24 * 60 * 60 * 1000);
                    recurringSheet.getRange(recurringRowIndex + 1, rLookup[COL_NEXT_SCHEDULED_DATE] + 1).setValue(newNextScheduledDate);
                } else {
                    logDebug(debugSheet, "Recurring key not found: " + recurringKey);
                }
            }


        }

        // If owner is blank, set it to "Shared"
        if (row[tLookup[COL_OWNER]] === "") {
            activeSheet.getRange(index + 1, tLookup[COL_OWNER] + 1).setValue("Shared");
        }
    });

    addActiveFilter();
}


// Function to add the filter to show only FALSE values in column "Complete" in the "Task" sheet, then sort by the "Owner" column
function addActiveFilter() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TODAYS_TASKS_SHEET_NAME);
    var filter = sheet.getFilter();
    if (filter) {
        filter.remove();
    }
    // Get headers
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rLookup = createHeaderLookup(headers);


    // Filter header is in row 3, data starts in row 4
    var range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
    sheet.getRange(3, rLookup[COL_COMPLETED] + 1, sheet.getLastRow() - 2, 1).insertCheckboxes();
    filter = range.createFilter();
    filter.setColumnFilterCriteria(rLookup[COL_COMPLETED] + 1, SpreadsheetApp.newFilterCriteria().whenTextEqualTo("FALSE").build());
    filter.sort(rLookup[COL_OWNER] + 1, true);

    // Hide all columns except "Task", "Owner", "Completed", and "Urgency"
    headers.forEach(function (header, index) {
        if (header !== COL_TASK_NAME && header !== COL_COMPLETED && header !== COL_OWNER && header !== COL_URGENCY) {
            sheet.hideColumns(index + 1);
        }
    });
    // Hide row 1
    sheet.hideRows(1);
    sheet.hideRows(3);
    // Freeze up to row 2
    sheet.setFrozenRows(2);
}

// Function to remove the filter from the "Task" sheet
function removeActiveFilter() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TODAYS_TASKS_SHEET_NAME);
    var filter = sheet.getFilter();
    if (filter) {
        filter.remove();
    }

    // Unhide all columns
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(function (header, index) {
        sheet.showColumns(index + 1);
    });
    sheet.showRows(1);
    sheet.showRows(3);
}

// Add this function to your script
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Task Actions')
        .addItem('Add Scheduled Tasks', 'updateTodaysTasks')
        .addItem('Resort Active Sheet', 'updateActiveSheet')
        .addItem('Remove Active Filter', 'removeActiveFilter')
        .addItem('Reset Processing State', 'resetProcessingState')
        .addToUi();
}

function resetProcessingState() {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('isScriptEditing');
    scriptProperties.deleteProperty('editInstance');
    // SpreadsheetApp.getUi().alert('Processing state has been reset.');
    updateActiveSheet();
}

function getOrCreateDebugSheet(ss) {
    var debugSheet = ss.getSheetByName(DEBUG_SHEET_NAME);
    if (!debugSheet) {
        debugSheet = ss.insertSheet(DEBUG_SHEET_NAME);
        debugSheet.appendRow(["Timestamp", "Message"]);
    }
    return debugSheet;
}

function logDebug(sheet, message) {
    sheet.appendRow([new Date(), message]);
}