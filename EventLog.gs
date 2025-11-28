/**
 * Event Logging
 * Persistent event log that survives campaign resets.
 */

/**
 * Logs an event to the Event Log sheet. This log persists across campaign resets.
 * @param {string} eventType The type of event (e.g., 'Day Processed', 'Location Found', 'Combat', etc.)
 * @param {string} description A brief description of the event
 * @param {string} details Optional additional details about the event
 */
function logEvent(eventType, description, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let eventLogSheet = ss.getSheetByName(SHEET_NAMES.EVENT_LOG);

    if (!eventLogSheet) {
      eventLogSheet = createOrGetSheet(ss, 'Event Log', '#9c27b0');
      setupEventLog(eventLogSheet);
    }

    const currentDay = ss.getRangeByName('CurrentDay') ? ss.getRangeByName('CurrentDay').getValue() : 0;
    const timestamp = new Date();

    eventLogSheet.appendRow([
      timestamp,
      currentDay,
      eventType,
      description,
      details || ''
    ]);

    const lastRow = eventLogSheet.getLastRow();
    eventLogSheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    eventLogSheet.getRange(lastRow, 2).setHorizontalAlignment('center');

  } catch (e) {
    console.error('Error logging event:', e);
  }
}
