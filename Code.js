// Constants for sheet names
const RECURRING_SHEET_NAME = "Recurring";
const TODAYS_TASKS_SHEET_NAME = "Active";
const MAIN_SHEET_NAME = "Main";
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
const COL_ACTIVE_ROW = "Active row";
const COL_OWNER = "Owner";

// Constants for column names in Main sheet
const COL_COMPLETE = "Complete";
const COL_REASSIGN = "Reassign";
const COL_REPROCESSING = "Reprocessing";
const COL_URGENCY = "Urgency";

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
    const col = 2; // column "B"

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
    const filter = sheet.getFilter();
    if (!filter) {
        return;
    }
    const r = filter.getRange();
    const criteria = filter.getColumnFilterCriteria(col).copy();
    filter.remove();
    r.createFilter().setColumnFilterCriteria(col, criteria);
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

    refreshMainFilter();

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
    var nextScheduledDate = row[rLookup[COL_NEXT_SCHEDULED_DATE]];
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
function addOneOffTaskFromCell(cellLocation, owner) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var taskName = ss.getRange(cellLocation).getValue();

    if (taskName === "") {
        return false;
    }

    var activeSheet = ss.getSheetByName(TODAYS_TASKS_SHEET_NAME);
    var activeData = activeSheet.getDataRange().getValues();
    var activeHeaders = activeData[0];
    var tLookup = createHeaderLookup(activeHeaders);

    var newTask = new Array(activeHeaders.length);
    newTask[tLookup[COL_TASK_NAME]] = ss.getRange(cellLocation).getValue();
    newTask[tLookup[COL_OWNER]] = owner;
    newTask[tLookup[COL_DATE_ADDED]] = new Date();
    newTask[tLookup[COL_COMPLETED]] = false;
    newTask[tLookup[COL_ACTIVE_ROW]] = "=ROW()";

    logDebug(getOrCreateDebugSheet(ss), "Adding one-off task: " + newTask[tLookup[COL_TASK_NAME]]);

    // Append the new task to the "Active" sheet
    activeSheet.appendRow(newTask);

    // Clear the cell
    ss.getRange(cellLocation).setValue("");

    return true;
}

const oneOffTaskRows = [2, 51, 100];

// Function to create one off task using the text in Main sheet J2
function addOneOffTaskA() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ss.getRange("B" + oneOffTaskRows[0]).getValue();
    addOneOffTaskFromCell("J" + oneOffTaskRows[0], owner) && refreshMainFilter();
}
function addOneOffTaskB() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ss.getRange("B" + oneOffTaskRows[1]).getValue();
    addOneOffTaskFromCell("J" + oneOffTaskRows[1], owner) && refreshMainFilter();
}
function addOneOffTaskC() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ""
    addOneOffTaskFromCell("J" + oneOffTaskRows[2], owner) && refreshMainFilter();
}


function onEdit(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = getOrCreateDebugSheet(ss);
    var sheet = e.source.getActiveSheet();

    logDebug(debugSheet, "onEdit triggered. Sheet: " + sheet.getName());

    // Check if the edit was made in the "Main" sheet
    if (sheet.getName() !== MAIN_SHEET_NAME) {
        logDebug(debugSheet, "Edit not in Main sheet. Refreshing Main view then exiting.");
        refreshMainFilter();
        return;
    }

    // Get the column numbers for "Complete", "Row", and "Reassign" columns
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var completeColNum = headers.indexOf(COL_COMPLETE) + 1;
    var reassignColNum = headers.indexOf(COL_REASSIGN) + 1;

    // Log the cell that was edited
    logDebug(debugSheet, "Edit cell: " + e.range.getA1Notation());

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

    // Check if it's a one-off task row, and ignore
    if (oneOffTaskRows.includes(e.range.getRow())) {
        logDebug(debugSheet, "Edit in one-off task row. Ignoring.");
        return;
    }

    var scriptProperties = PropertiesService.getScriptProperties();
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

    runMainUpdate();
}

function runMainUpdate() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = ss.getSheetByName(TODAYS_TASKS_SHEET_NAME);
    var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    var debugSheet = getOrCreateDebugSheet(ss);
    var headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];

    // Set processing flag
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('isScriptEditing', 'true');

    // Use a try-finally block to ensure the processing flag is always reset
    try {
        logDebug(debugSheet, "Starting processing.");
        var rowColNum = headers.indexOf(COL_ACTIVE_ROW) + 1;
        var reprocessingColNum = headers.indexOf(COL_REPROCESSING) + 1;
        var recurringForeignKeyCol = headers.indexOf(COL_RECURRING_KEY) + 1;
        var completeColNum = headers.indexOf(COL_COMPLETE) + 1;
        var reassignColNum = headers.indexOf(COL_REASSIGN) + 1;

        // Hide "Complete" and "Reassign" columns
        mainSheet.hideColumns(completeColNum);
        mainSheet.hideColumns(reassignColNum);
        mainSheet.showColumns(reprocessingColNum);
        logDebug(debugSheet, "Hidden 'Complete' and 'Reassign' columns");

        logDebug(debugSheet, "Headers: Complete column: " + completeColNum + ", Row column: " + rowColNum + ", Reassign column: " + reassignColNum);

        // Process the state immediately
        var activeData = activeSheet.getDataRange().getValues();
        var activeHeaders = activeData[0];
        var completedColIndex = activeHeaders.indexOf(COL_COMPLETED);
        var rowColIndex = activeHeaders.indexOf(COL_ACTIVE_ROW);
        var ownerColIndex = activeHeaders.indexOf(COL_OWNER);

        var recurringSheet = ss.getSheetByName(RECURRING_SHEET_NAME);
        var recurringData = recurringSheet.getDataRange().getValues();
        var recurringHeaders = recurringData[0];
        var recurringKeyColIndex = recurringHeaders.indexOf(COL_RECURRING_KEY);
        var recurringNextScheduledIndex = recurringHeaders.indexOf(COL_NEXT_SCHEDULED_DATE);
        var recurringScheduleFromCompletionIndex = recurringHeaders.indexOf(COL_SCHEDULE_FROM_COMPLETION);
        var recurringDaysIndex = recurringHeaders.indexOf(COL_DAYS);

        logDebug(debugSheet, "Processing state. Headers: Completed column in Active: " + (completedColIndex + 1) +
            ", Row column in Active: " + (rowColIndex + 1) + ", Owner column in Active: " + (ownerColIndex + 1));

        var changesCount = 0;

        // Function to get the current state of the "Complete" and "Reassign" columns
        function getColumnState() {
            var data = mainSheet.getRange(MAIN_DATA_OFFSET, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
            return data.map(row => ({
                complete: row[completeColNum - 1],
                reassign: row[reassignColNum - 1]
            }));
        }

        // Get the state after the edit
        var currentState = getColumnState();

        currentState.reverse();

        currentState.forEach((state, rev_index) => {
            var index = currentState.length - rev_index - 1;

            // Skip one-off task rows
            if (oneOffTaskRows.includes(index + MAIN_DATA_OFFSET)) {
                return;
            }

            if (state.complete || state.reassign) {
                var sourceRow = mainSheet.getRange(index + MAIN_DATA_OFFSET, rowColNum).getValue();
                logDebug(debugSheet, "Processing row " + (index + MAIN_DATA_OFFSET) + " in Main, sourceRow: " + sourceRow);

                if (state.complete) {
                    changesCount++;
                    // Clear the checkbox
                    // mainSheet.getRange(index + MAIN_DATA_OFFSET, completeColNum).setValue(false);
                    // logDebug(debugSheet, "Cleared checkbox in Main sheet row " + (index + MAIN_DATA_OFFSET));

                    var recurringKey = mainSheet.getRange(index + MAIN_DATA_OFFSET, recurringForeignKeyCol).getValue();
                    if (recurringKey) {
                        logDebug(debugSheet, "Looking for a recurring task with key " + recurringKey + " in " + recurringData.length + " tasks");

                        if (!recurringData.some(function (row, rindex) {
                            if (row[recurringKeyColIndex] == recurringKey) {
                                logDebug(debugSheet, "Found at row " + (rindex + 1));
                                logDebug(debugSheet, JSON.stringify(row));

                                if (row[recurringScheduleFromCompletionIndex]) {
                                    recurringSheet.getRange(rindex + 1, recurringNextScheduledIndex + 1).setValue(new Date(new Date().getTime() + row[recurringDaysIndex] * 24 * 60 * 60 * 1000));
                                }
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
                    changesCount++;
                    if (state.reassign.toLowerCase() === "unassign") {
                        activeSheet.getRange(sourceRow, ownerColIndex + 1).setValue("");
                        logDebug(debugSheet, "Unassigned owner in Active sheet row " + (sourceRow));
                    } else {
                        activeSheet.getRange(sourceRow, ownerColIndex + 1).setValue(state.reassign);
                        logDebug(debugSheet, "Reassigned owner to '" + state.reassign + "' in Active sheet row " + (sourceRow));
                    }

                    // Clear the Reassign cell in the Main sheet
                    // mainSheet.getRange(index + MAIN_DATA_OFFSET, reassignColNum).setValue("");
                    // logDebug(debugSheet, "Cleared Reassign cell in Main sheet row " + (index + MAIN_DATA_OFFSET));
                }
            }
        });

        // Clear all the checkboxes
        mainSheet.getRange(MAIN_DATA_OFFSET, completeColNum, mainSheet.getLastRow() - MAIN_DATA_OFFSET, 1).setValue("");
        logDebug(debugSheet, "Cleared all checkboxes in Main sheet");
        // Clear all the reassign cells
        mainSheet.getRange(MAIN_DATA_OFFSET, reassignColNum, mainSheet.getLastRow() - MAIN_DATA_OFFSET, 1).setValue("");
        logDebug(debugSheet, "Cleared all Reassign cells in Main sheet");

        logDebug(debugSheet, "Processed " + changesCount + " changes");

        refreshMainFilter();

    } catch (error) {
        logDebug(debugSheet, "Error occurred: " + error.toString());
    } finally {
        // Reset processing flag
        scriptProperties.deleteProperty('isScriptEditing');
        logDebug(debugSheet, "Processing complete. Reset processing flag.");

        // Unhide "Complete" and "Reassign" columns
        mainSheet.showColumns(completeColNum);
        mainSheet.showColumns(reassignColNum);
        mainSheet.hideColumns(reprocessingColNum);

        logDebug(debugSheet, "Unhidden 'Complete' and 'Reassign' columns");

    }
}

// Add this function to your script
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Task Actions')
        .addItem('Add Scheduled Tasks', 'updateTodaysTasks')
        .addItem('Add Main Filter', 'addMainFilter')
        .addItem('Remove Main Filter', 'removeMainFilter')
        .addItem('Reset Processing State', 'resetProcessingState')
        .addToUi();
}

function resetProcessingState() {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('isScriptEditing');
    scriptProperties.deleteProperty('editInstance');
    // SpreadsheetApp.getUi().alert('Processing state has been reset.');
    addMainFilter();
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


// Function to add the filter:
//    =AND(NOT(ISBLANK(B:B)), NOT(REGEXMATCH(B:B, "Task")))
// to the "Main" sheet, column "B"

function addMainFilter() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
    var filter = sheet.getFilter();
    if (filter) {
        filter.remove();
    }

    var range = sheet.getRange("B:B");
    filter = range.createFilter();
    filter.setColumnFilterCriteria(2, SpreadsheetApp.newFilterCriteria().whenFormulaSatisfied("=AND(NOT(ISBLANK(B:B)), NOT(REGEXMATCH(B:B, \"Task\")))").build());

    // Hide all columns except a few
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(function (header, index) {
        if (header !== COL_TASK_NAME && header !== COL_COMPLETE && header !== COL_REASSIGN && header !== COL_URGENCY) {
            sheet.hideColumns(index + 1);
        }
    });

    // Hide row 1
    sheet.hideRows(1);
}


// Reversal of addMainFilter

function removeMainFilter() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
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
}