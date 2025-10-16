function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Event Menu')
      .addItem('Send Selected Row(s) to calendar', 'sendToA11yCalendar') // Changed menu item text for clarity
      .addToUi();
}

/**
 * Creates all-day calendar events for each selected row in the "[SheetName]" sheet.
 * It extracts the DFR SME, Start Date, and End Date from columns 3, 1, and 2 respectively.
 */
function sendToA11yCalendar() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("[SheetName]");

  // Get the active range, which can be a single cell, a row, or multiple rows/columns.
  var activeRange = SpreadsheetApp.getActiveRange();

  // Get all values from the active range. This returns a 2D array.
  // Each inner array represents a row, and elements within it are column values.
  var values = activeRange.getValues();

  // Get the starting row index of the active range. This is important if the selection
  // doesn't start from row 1, as sheet.getRange().getValues() returns values relative
  // to the top-left of the selected range, not the sheet.
  var startRow = activeRange.getRow();

  // Define the column indices (0-based) for the data we need.
  // Assuming:
  // Column 1 (A) = Start Date -> index 0
  // Column 2 (B) = End Date   -> index 1
  // Column 3 (C) = Assigned SME -> index 2
  const START_DATE_COL_INDEX = 0;
  const END_DATE_COL_INDEX = 1;
  const ASSIGNED_SME_COL_INDEX = 2; // Column C

  // Get the calendar by its ID.
  var calendar = CalendarApp.getCalendarById('[CalendarID]');

  // Check if the calendar was found.
  if (!calendar) {
    SpreadsheetApp.getUi().alert('Error: Calendar not found. Please check the Calendar ID.');
    return;
  }

  // Iterate over each row in the selected range.
  for (var i = 0; i < values.length; i++) {
    var rowData = values[i]; // This is an array representing the current row's data

    // Extract values using the defined column indices.
    var startDate = rowData[START_DATE_COL_INDEX];
    var endDate = rowData[END_DATE_COL_INDEX];
    var assignedSme = rowData[ASSIGNED_SME_COL_INDEX];

    // Basic validation: Check if essential data is present.
    // Assuming startDate and assignedSme are mandatory.
    if (!startDate || !assignedSme) {
      // Log an error or skip the row if data is incomplete.
      console.warn('Skipping row ' + (startRow + i) + ' due to missing Start Date or Assigned SME.');
      continue; // Move to the next row
    }

    // Construct the event title.
    var title = "DFR " + assignedSme;

    try {
      // Create the all-day event.
      // If endDate is empty, create a single-day event.
      if (endDate) {
        calendar.createAllDayEvent(title, startDate, endDate);
      } else {
        calendar.createAllDayEvent(title, startDate); // Single day event if no end date
      }
      console.log('Event created: ' + title + ' from ' + startDate + ' to ' + (endDate || startDate));
    } catch (e) {
      console.error('Error creating event for row ' + (startRow + i) + ': ' + e.message);
      SpreadsheetApp.getUi().alert('Error creating event for row ' + (startRow + i) + ': ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert('Events created successfully for selected rows!');
}


/**
 * Triggered automatically on spreadsheet edits.
 * Checks if the edit is in the "Assigned SME" column and updates the corresponding calendar event.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object containing information about the edit.
 */
function updateCalendarEventOnEdit(e) {
  const sheetName = "[SheetName]";
  const SME_COLUMN = 3; // Column C (1-based index)
  const START_DATE_COLUMN = 1; // Column A (1-based index)
  const END_DATE_COLUMN = 2; // Column B (1-based index)

  var range = e.range;
  var sheet = range.getSheet();

  // Ensure the edit is on the correct sheet and in the SME column.
  if (sheet.getName() !== sheetName || range.getColumn() !== SME_COLUMN) {
    return; // Not the sheet or column we're interested in.
  }

  var editedRow = range.getRow();
  var newSme = e.value; // The new value entered in the cell
  var oldSme = e.oldValue; // The value before the edit

  // If there's no old value, it means a new value was added to an empty cell,
  // or the cell was cleared. We only want to update if an existing SME changed.
  if (!oldSme || oldSme === newSme) {
    return;
  }

  // Get the start and end dates from the same row.
  var startDate = sheet.getRange(editedRow, START_DATE_COLUMN).getValue();
  var endDate = sheet.getRange(editedRow, END_DATE_COLUMN).getValue();

  // Basic validation for dates and new SME
  if (!startDate || !newSme) {
    console.warn('Cannot update event: Missing Start Date or New SME in row ' + editedRow);
    return;
  }

  var calendar = CalendarApp.getCalendarById('[CalendarID]');

  if (!calendar) {
    console.error('Error: Calendar not found for update. Please check the Calendar ID.');
    return;
  }

  // Construct the old and new event titles for searching and updating.
  var oldTitle = "DFR " + oldSme;
  var newTitle = "DFR " + newSme;

  try {
    // Search for events that match the old title and date range.
    // getEvents(startTime, endTime) requires Date objects.
    // For all-day events, the start date is the beginning of the day, and end date is the end of the day.
    // If endDate is not provided, assume it's a single-day event.
    var searchEndDate = endDate ? new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate() + 1) : new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 1);

    var events = calendar.getEvents(startDate, searchEndDate, { search: oldTitle });

    if (events.length > 0) {
      // Assuming there's only one relevant event per row/SME.
      // If multiple events match, you might need more specific criteria.
      for (var i = 0; i < events.length; i++) {
        var event = events[i];
        // Double-check if the event title truly matches the old title to avoid false positives
        if (event.getTitle() === oldTitle) {
          event.setTitle(newTitle);
          console.log('Updated event: From "' + oldTitle + '" to "' + newTitle + '" for row ' + editedRow);
          SpreadsheetApp.getUi().alert('Calendar event updated successfully for row ' + editedRow + '!');
          return; // Assuming one event per row, exit after updating.
        }
      }
    } else {
      console.warn('No existing calendar event found for "' + oldTitle + '" on ' + startDate.toDateString() + ' (row ' + editedRow + ').');
      SpreadsheetApp.getUi().alert('No existing calendar event found to update for row ' + editedRow + '.');
    }
  } catch (e) {
    console.error('Error updating calendar event for row ' + editedRow + ': ' + e.message);
    SpreadsheetApp.getUi().alert('Error updating calendar event for row ' + editedRow + ': ' + e.message);
  }
}
