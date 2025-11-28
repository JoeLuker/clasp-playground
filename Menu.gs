/**
 * Menu and Sidebar Initialization
 * Handles the Campaign menu and sidebar display.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Campaign')
    .addItem('Show Campaign Manager', 'showSidebar')
    .addSeparator()
    .addItem('Initialize/Reset Campaign', 'initializeCompleteCampaign')
    .addSeparator()
    .addItem('Log Custom Event', 'logCustomEvent')
    .addSeparator()
    .addItem('Import Hex Map', 'importHexMapPrompt')
    .addItem('Set Current Hex', 'setCurrentHexPrompt')
    .addToUi();
}

/**
 * Creates and displays the HTML sidebar.
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle('Campaign Manager')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Helper function to include other HTML files (CSS, JS) into the main Sidebar HTML.
 * @param {string} filename The name of the HTML file to include.
 * @returns {string} The content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Prompts the user to log a custom event manually.
 */
function logCustomEvent() {
  const ui = SpreadsheetApp.getUi();

  const eventTypeResponse = ui.prompt(
    'Log Custom Event',
    'Enter the event type (e.g., "Combat", "Location Found", "NPC Encounter"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (eventTypeResponse.getSelectedButton() !== ui.Button.OK) return;
  const eventType = eventTypeResponse.getResponseText().trim();
  if (!eventType) return;

  const descriptionResponse = ui.prompt(
    'Event Description',
    'Enter a brief description of the event:',
    ui.ButtonSet.OK_CANCEL
  );

  if (descriptionResponse.getSelectedButton() !== ui.Button.OK) return;
  const description = descriptionResponse.getResponseText().trim();
  if (!description) return;

  const detailsResponse = ui.prompt(
    'Event Details',
    'Enter additional details (optional, leave blank to skip):',
    ui.ButtonSet.OK_CANCEL
  );

  const details = detailsResponse.getSelectedButton() === ui.Button.OK
    ? detailsResponse.getResponseText().trim()
    : '';

  logEvent(eventType, description, details);
  ui.alert('Event Logged', 'The event has been added to the Event Log.', ui.ButtonSet.OK);
}
