/**
 * Pathfinder Campaign Manager - Complete System
 * Backend script for Google Sheets.
 * Manages all sheet interactions and provides data to the HTML sidebar.
 */

// ============= MENU AND SIDEBAR INITIALIZATION =============

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
 * Prompts user to import hex map JSON data.
 */
function importHexMapPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Import Hex Map',
    'Paste the JSON data from your hex map editor:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const jsonData = response.getResponseText();
    const result = importHexMap(jsonData);
    ui.alert('Import Result', result, ui.ButtonSet.OK);
  }
}

/**
 * Prompts user to set current hex location.
 */
function setCurrentHexPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Current Hex',
    'Enter hex coordinates (format: q:r:s, e.g., 0:0:0):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const hexCoords = response.getResponseText().trim();
    const result = setCurrentHex(hexCoords);
    ui.alert('Location Updated', result, ui.ButtonSet.OK);
  }
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

// ============= DATA GETTER FUNCTIONS FOR SIDEBAR =============

/**
 * Gets all the key data points from the spreadsheet to display on the sidebar.
 * @returns {object} An object containing all the necessary dashboard and status data.
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { error: "Cannot access spreadsheet. Please ensure the script is bound to a Google Sheet and you have permission to access it." };
    }
    return {
      currentDay: ss.getRangeByName('CurrentDay').getValue(),
      totalMiles: ss.getRangeByName('TotalMiles').getValue(),
      currentHex: ss.getRangeByName('CurrentHex') ? ss.getRangeByName('CurrentHex').getValue() : '',
      hexTerrain: ss.getRangeByName('HexTerrain') ? ss.getRangeByName('HexTerrain').getValue() : '',
      temperature: ss.getRangeByName('Temperature').getValue(),
      terrain: ss.getRangeByName('Terrain').getValue(),
      pathType: ss.getRangeByName('PathType').getValue(),
      weather: ss.getRangeByName('Weather').getValue(),
      travelPace: ss.getRangeByName('TravelPace').getValue(),
      animalsGrazing: ss.getRangeByName('AnimalsGrazing').getValue(),
      foodDays: ss.getRangeByName('FoodDaysLeft').getValue(),
      foodStatus: ss.getRangeByName('FoodStatus').getValue(),
      waterDays: ss.getRangeByName('WaterDaysLeft').getValue(),
      waterStatus: ss.getRangeByName('WaterStatus').getValue(),
      avgHp: ss.getRangeByName('AvgHPPercent').getDisplayValue(),
      checksNeeded: ss.getRangeByName('ChecksNeeded').getValue(),
      criticalCount: ss.getRangeByName('CriticalCount').getValue(),
      alertText: ss.getRangeByName('AlertText').getValue()
    };
  } catch (e) {
    Logger.log('Error in getDashboardData: ' + e.toString());
    if (e.toString().includes('failed while accessing document')) {
      return { error: "Permission error: The script cannot access the spreadsheet. Please ensure:\n1. The script is bound to your Google Sheet\n2. You have authorized the script\n3. You have permission to edit the spreadsheet\n\nTry running 'Initialize/Reset Campaign' from the menu to re-authorize." };
    }
    return { error: `Error: ${e.message}. Please run 'Initialize/Reset Campaign' from the menu if the sheet is not initialized.` };
  }
}

/**
 * Updates the resources found values.
 * @param {number} foodFound Amount of food found (in days)
 * @param {number} waterFound Amount of water found (in gallons)
 * @returns {string} Confirmation message
 */
function updateResourcesFound(foodFound, waterFound) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (foodFound > 0) {
    ss.getRangeByName('FoodFound').setValue(foodFound);
    logEvent('Resource Found', `Found ${foodFound} days of food`, '');
  }
  if (waterFound > 0) {
    ss.getRangeByName('WaterFound').setValue(waterFound);
    logEvent('Resource Found', `Found ${waterFound} gallons of water`, '');
  }
  return 'Resources updated successfully';
}

/**
 * Updates environment settings from the sidebar.
 * @param {object} environment Object containing temperature, terrain, pathType, weather, travelPace, animalsGrazing
 * @returns {string} Confirmation message
 */
function updateEnvironment(environment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRangeByName('Temperature').setValue(environment.temperature);
  ss.getRangeByName('Terrain').setValue(environment.terrain);
  ss.getRangeByName('PathType').setValue(environment.pathType);
  ss.getRangeByName('Weather').setValue(environment.weather);
  ss.getRangeByName('TravelPace').setValue(environment.travelPace);
  ss.getRangeByName('AnimalsGrazing').setValue(environment.animalsGrazing);
  
  logEvent('Environment Changed', 
    `${environment.terrain} terrain, ${environment.weather} weather, ${environment.travelPace} pace`,
    `Temperature: ${environment.temperature}, Path: ${environment.pathType}, Grazing: ${environment.animalsGrazing ? 'Yes' : 'No'}`);
  
  return 'Environment updated successfully';
}

// ============= ACTION FUNCTIONS CALLED BY SIDEBAR =============

/**
 * Processes the day's events, consumes resources, and returns a status message.
 * @returns {string} A confirmation message for the user.
 */
function processDay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const currentDay = ss.getRangeByName('CurrentDay').getValue();
  const miles = ss.getRangeByName('CalcMilesToday').getValue();
  const newDay = currentDay + 1;
  ss.getRangeByName('CurrentDay').setValue(newDay);
  ss.getRangeByName('TotalMiles').setValue(ss.getRangeByName('TotalMiles').getValue() + miles);
  
  const foodUsed = ss.getRangeByName('FoodDaily').getValue();
  const waterUsed = ss.getRangeByName('WaterDaily').getValue();
  const fodderUsed = ss.getRangeByName('FodderDaily').getValue();
  const provisionUsed = ss.getRangeByName('ProvisionDaily').getValue();
  const foodFound = ss.getRangeByName('FoodFound').getValue();
  const waterFound = ss.getRangeByName('WaterFound').getValue();
  
  ss.getRangeByName('FoodStock').setValue(
    ss.getRangeByName('FoodStock').getValue() - foodUsed + foodFound
  );
  ss.getRangeByName('WaterStock').setValue(
    ss.getRangeByName('WaterStock').getValue() - waterUsed + waterFound
  );
  ss.getRangeByName('FodderStock').setValue(
    ss.getRangeByName('FodderStock').getValue() - fodderUsed
  );
  ss.getRangeByName('ProvisionStock').setValue(
    ss.getRangeByName('ProvisionStock').getValue() - provisionUsed
  );
  updateDeprivation(ss);
  const log = ss.getSheetByName('Log');
  log.appendRow([
    newDay, new Date(), miles,
    foodUsed,
    waterUsed,
    fodderUsed,
    provisionUsed,
    ''
  ]);
  ss.getRangeByName('FoodFound').setValue(0);
  ss.getRangeByName('WaterFound').setValue(0);
  
  // Log the day processing event
  const details = `Traveled ${Math.round(miles)} miles. Resources: Food -${foodUsed}${foodFound > 0 ? ' +' + foodFound : ''}, Water -${waterUsed}${waterFound > 0 ? ' +' + waterFound : ''}, Fodder -${fodderUsed}, Provisions -${provisionUsed}`;
  logEvent('Day Processed', `Advanced to Day ${newDay}`, details);
  
  return `Advanced to Day ${newDay}. Traveled ${Math.round(miles)} miles.`;
}

/**
 * Gets the daily preview data for display.
 * @returns {object} An object with the calculated preview data.
 */
function previewDay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    miles: Math.round(ss.getRangeByName('CalcMilesToday').getValue()),
    food: ss.getRangeByName('FoodDaily').getValue(),
    water: ss.getRangeByName('WaterDaily').getValue(),
    fodder: ss.getRangeByName('FodderDaily').getValue(),
    provisions: ss.getRangeByName('ProvisionDaily').getValue(),
  };
}

/**
 * Updates deprivation stats for all characters in the roster.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 */
function updateDeprivation(ss) {
  const hasFood = ss.getRangeByName('FoodFound').getValue() > 0 || ss.getRangeByName('FoodDaily').getValue() < ss.getRangeByName('FoodStock').getValue();
  const hasWater = ss.getRangeByName('WaterFound').getValue() > 0 || ss.getRangeByName('WaterDaily').getValue() < ss.getRangeByName('WaterStock').getValue();
  const hoursPerDay = ss.getRangeByName('HOURS_PER_DAY').getValue();
  const characters = ['Char1', 'Char2', 'Char3', 'Char4'];
  characters.forEach(charPrefix => {
    const charName = ss.getRangeByName(`${charPrefix}_Name`).getValue();
    if (charName === '') return;
    if (hasFood) {
      ss.getRangeByName(`${charPrefix}_DaysNoFood`).setValue(0);
    } else {
      ss.getRangeByName(`${charPrefix}_DaysNoFood`).setValue(ss.getRangeByName(`${charPrefix}_DaysNoFood`).getValue() + 1);
    }
    if (hasWater) {
      ss.getRangeByName(`${charPrefix}_HoursNoWater`).setValue(0);
    } else {
      ss.getRangeByName(`${charPrefix}_HoursNoWater`).setValue(ss.getRangeByName(`${charPrefix}_HoursNoWater`).getValue() + hoursPerDay);
    }
    ss.getRangeByName(`${charPrefix}_HoursNoSleep`).setValue(ss.getRangeByName(`${charPrefix}_HoursNoSleep`).getValue() + hoursPerDay);
  });
}

// ============= FULL SYSTEM INITIALIZATION =============

function initializeCompleteCampaign() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm Full Reset',
    'This will DELETE ALL EXISTING SHEETS in this spreadsheet and build a completely new campaign manager. The Event Log will be preserved. Are you sure you want to continue?',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  
  // Preserve Event Log if it exists
  let eventLogSheet = ss.getSheetByName('Event Log');
  let eventLogData = null;
  if (eventLogSheet) {
    // Save the data
    const lastRow = eventLogSheet.getLastRow();
    if (lastRow > 1) {
      eventLogData = eventLogSheet.getRange(1, 1, lastRow, eventLogSheet.getLastColumn()).getValues();
    }
  }
  
  const tempSheet = ss.insertSheet('temp');
  const allSheets = ss.getSheets();
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName !== 'temp' && sheetName !== 'Event Log') {
      ss.deleteSheet(sheet);
    }
  });
  createAllSheets(ss);
  createAllNamedRanges(ss);
  addAllFormulas(ss);
  
  // Restore or create Event Log
  if (eventLogSheet) {
    // Sheet was preserved, just make sure it's set up correctly
    setupEventLog(eventLogSheet);
    if (eventLogData && eventLogData.length > 1) {
      // Restore the data (skip header row)
      eventLogSheet.getRange(2, 1, eventLogData.length - 1, eventLogData[0].length).setValues(eventLogData.slice(1));
    }
  } else {
    // Create new Event Log
    eventLogSheet = createOrGetSheet(ss, 'Event Log', '#9c27b0');
    setupEventLog(eventLogSheet);
  }
  
  // Log the initialization
  logEvent('Campaign Initialized', 'System reset and all sheets recreated');
  
  ss.deleteSheet(tempSheet);
  
  onOpen(); 
  ui.alert('Campaign Initialized', 'System ready. Open the sidebar from the Campaign menu.', ui.ButtonSet.OK);
}

function createAllSheets(ss) {
  const dashboard = createOrGetSheet(ss, 'Dashboard', '#4285f4');
  const calculations = createOrGetSheet(ss, 'Calculations', '#0f9d58');
  const roster = createOrGetSheet(ss, 'Roster', '#ea4335');
  const status = createOrGetSheet(ss, 'Status Monitor', '#fbbc04');
  const constants = createOrGetSheet(ss, 'Constants', '#9c27b0');
  const references = createOrGetSheet(ss, 'Reference Tables', '#9e9e9e');
  const caravan = createOrGetSheet(ss, 'Caravan', '#ff6f00');
  const exploration = createOrGetSheet(ss, 'Exploration', '#00897b');
  const log = createOrGetSheet(ss, 'Log', '#673ab7');
  const hexMap = createOrGetSheet(ss, 'Hex Map', '#00bcd4');
  setupDashboard(dashboard);
  setupCalculations(calculations);
  setupRoster(roster);
  setupStatus(status);
  setupConstants(constants);
  setupReferences(references);
  setupCaravan(caravan);
  setupExploration(exploration);
  setupLog(log);
  setupHexMap(hexMap);
}

function createAllNamedRanges(ss) {
    const ranges = {
    'STARVATION_DAYS': 'Constants!B2', 'DEHYDRATION_HOURS': 'Constants!B3', 'SURVIVAL_DC_BASE': 'Constants!B4', 'EXHAUSTION_HOURS': 'Constants!B5', 'HP_CRITICAL': 'Constants!B7', 'HP_BLOODIED': 'Constants!B8', 'HP_WOUNDED': 'Constants!B9', 'DAYS_CRITICAL': 'Constants!B11', 'DAYS_LOW': 'Constants!B12', 'DAYS_GOOD': 'Constants!B13', 'BASE_SPEED': 'Constants!B15', 'TRAVEL_HOURS_PER_DAY': 'Constants!B16', 'FODDER_PER_MOUNT': 'Constants!B17', 'FAST_MULTIPLIER': 'Constants!B18', 'SLOW_MULTIPLIER': 'Constants!B19', 'NORMAL_MULTIPLIER': 'Constants!B20', 'HOT_WATER_MULT': 'Constants!B22', 'SEVERE_HEAT_MULT': 'Constants!B23', 'EXTREME_HEAT_MULT': 'Constants!B24', 'COLD_WATER_MULT': 'Constants!B25', 'MAX_DAYS': 'Constants!B27', 'HOURS_PER_DAY': 'Constants!B28', 'UNREST_PENALTY_MULT': 'Constants!B29',
    'CurrentDay': 'Dashboard!B4', 'TotalMiles': 'Dashboard!B5', 'CurrentHex': 'Hex Map!B4', 'HexTerrain': 'Hex Map!B5', 'Temperature': 'Dashboard!B8', 'Terrain': 'Dashboard!B9', 'PathType': 'Dashboard!B10', 'Weather': 'Dashboard!B11', 'Altitude': 'Dashboard!B12', 'TravelPace': 'Dashboard!F4', 'AnimalsGrazing': 'Dashboard!F5', 'FoodFound': 'Dashboard!F6', 'WaterFound': 'Dashboard!F7', 'FoodStock': 'Dashboard!B14', 'WaterStock': 'Dashboard!B15', 'FodderStock': 'Dashboard!B16', 'ProvisionStock': 'Dashboard!B17',
    'CalcBaseSpeed': 'Calculations!B5', 'CalcTerrainMod': 'Calculations!B6', 'CalcPaceMod': 'Calculations!B7', 'CalcWeatherMod': 'Calculations!B8', 'CalcMilesToday': 'Calculations!B9', 'FoodBase': 'Calculations!B13', 'FoodMod': 'Calculations!C13', 'FoodDaily': 'Calculations!D13', 'FoodDaysLeft': 'Calculations!E13', 'FoodStatus': 'Calculations!F13', 'WaterBase': 'Calculations!B14', 'WaterMod': 'Calculations!C14', 'WaterDaily': 'Calculations!D14', 'WaterDaysLeft': 'Calculations!E14', 'WaterStatus': 'Calculations!F14', 'FodderBase': 'Calculations!B15', 'FodderMod': 'Calculations!C15', 'FodderDaily': 'Calculations!D15', 'FodderDaysLeft': 'Calculations!E15', 'FodderStatus': 'Calculations!F15', 'ProvisionBase': 'Calculations!B16', 'ProvisionMod': 'Calculations!C16', 'ProvisionDaily': 'Calculations!D16', 'ProvisionDaysLeft': 'Calculations!E16', 'ProvisionStatus': 'Calculations!F16',
    'Char1_Name': 'Roster!A4', 'Char1_Class': 'Roster!B4', 'Char1_CON': 'Roster!C4', 'Char1_Fort': 'Roster!D4', 'Char1_HP': 'Roster!E4', 'Char1_MaxHP': 'Roster!F4', 'Char1_DaysNoFood': 'Roster!G4', 'Char1_HoursNoWater': 'Roster!H4', 'Char1_HoursNoSleep': 'Roster!I4', 'Char1_CheckNeeded': 'Roster!J4', 'Char1_CheckDC': 'Roster!K4', 'Char1_CheckType': 'Roster!L4', 'Char1_Status': 'Roster!M4',
    'Char2_Name': 'Roster!A5', 'Char2_Class': 'Roster!B5', 'Char2_CON': 'Roster!C5', 'Char2_Fort': 'Roster!D5', 'Char2_HP': 'Roster!E5', 'Char2_MaxHP': 'Roster!F5', 'Char2_DaysNoFood': 'Roster!G5', 'Char2_HoursNoWater': 'Roster!H5', 'Char2_HoursNoSleep': 'Roster!I5', 'Char2_CheckNeeded': 'Roster!J5', 'Char2_CheckDC': 'Roster!K5', 'Char2_CheckType': 'Roster!L5', 'Char2_Status': 'Roster!M5',
    'Char3_Name': 'Roster!A6', 'Char3_Class': 'Roster!B6', 'Char3_CON': 'Roster!C6', 'Char3_Fort': 'Roster!D6', 'Char3_HP': 'Roster!E6', 'Char3_MaxHP': 'Roster!F6', 'Char3_DaysNoFood': 'Roster!G6', 'Char3_HoursNoWater': 'Roster!H6', 'Char3_HoursNoSleep': 'Roster!I6', 'Char3_CheckNeeded': 'Roster!J6', 'Char3_CheckDC': 'Roster!K6', 'Char3_CheckType': 'Roster!L6', 'Char3_Status': 'Roster!M6',
    'Char4_Name': 'Roster!A7', 'Char4_Class': 'Roster!B7', 'Char4_CON': 'Roster!C7', 'Char4_Fort': 'Roster!D7', 'Char4_HP': 'Roster!E7', 'Char4_MaxHP': 'Roster!F7', 'Char4_DaysNoFood': 'Roster!G7', 'Char4_HoursNoWater': 'Roster!H7', 'Char4_HoursNoSleep': 'Roster!I7', 'Char4_CheckNeeded': 'Roster!J7', 'Char4_CheckDC': 'Roster!K7', 'Char4_CheckType': 'Roster!L7', 'Char4_Status': 'Roster!M7',
    'PartySize': 'Roster!P4', 'MountCount': 'Roster!P5', 'ChecksNeeded': 'Roster!P6', 'CriticalCount': 'Roster!P7', 'AvgHPPercent': 'Roster!P8', 'TotalSize': 'Roster!P10',
    'Mount1_Name': 'Roster!A19', 'Mount1_Type': 'Roster!B19', 'Mount1_HD': 'Roster!C19', 'Mount1_HP': 'Roster!D19', 'Mount1_MaxHP': 'Roster!E19', 'Mount1_Speed': 'Roster!F19', 'Mount1_Carrying': 'Roster!G19', 'Mount1_MaxLoad': 'Roster!H19', 'Mount1_Grazing': 'Roster!I19', 'Mount1_Status': 'Roster!J19',
    'CaravanName': 'Caravan!B4', 'CaravanLevel': 'Caravan!D4', 'CaravanOffense': 'Caravan!B5', 'CaravanDefense': 'Caravan!D5', 'CaravanMobility': 'Caravan!B6', 'CaravanMorale': 'Caravan!D6', 'CaravanUnrest': 'Caravan!B7', 'CaravanMaxUnrest': 'Caravan!D7', 'CaravanFortune': 'Caravan!B8', 'CaravanPrestige': 'Caravan!D8', 'CaravanAttack': 'Caravan!G4', 'CaravanAC': 'Caravan!I4', 'CaravanSecurity': 'Caravan!G5', 'CaravanResolve': 'Caravan!I5', 'CaravanSpeed': 'Caravan!G6', 'CaravanHP': 'Caravan!I6', 'CaravanCargo': 'Caravan!G7', 'CaravanTravelers': 'Caravan!I7', 'CaravanConsumption': 'Caravan!G8',
    'TerritoryName': 'Exploration!B4', 'TerritoryCR': 'Exploration!B5', 'ExplorationSkill': 'Exploration!B6', 'ExplorationDC': 'Exploration!B7', 'CurrentDP': 'Exploration!B8', 'DaysExplored': 'Exploration!B9',
    'StatusFood': 'Status Monitor!B5', 'StatusWater': 'Status Monitor!B6', 'StatusFodder': 'Status Monitor!B7', 'StatusProvision': 'Status Monitor!B8', 'AlertFood': 'Status Monitor!C5', 'AlertWater': 'Status Monitor!C6', 'AlertFodder': 'Status Monitor!C7', 'AlertProvision': 'Status Monitor!C8', 'PartyAvgHP': 'Status Monitor!G5', 'PartyChecks': 'Status Monitor!G6', 'PartyCritical': 'Status Monitor!G7', 'AlertText': 'Status Monitor!A11',
    'TerrainTable': 'Reference Tables!A3:D12', 'PaceTable': 'Reference Tables!F3:G5', 'TempTable': 'Reference Tables!I3:J9', 'WeatherTable': 'Reference Tables!A15:B20', 'ForageTable': 'Reference Tables!D15:E19', 'AltitudeTable': 'Reference Tables!G15:H18',
    'AllCharNames': 'Roster!A4:A15', 'AllCharHP': 'Roster!E4:E15', 'AllCharMaxHP': 'Roster!F4:F15', 'AllCharChecks': 'Roster!J4:J15', 'AllCharStatus': 'Roster!M4:M15', 'AllMountNames': 'Roster!A19:A30', 'AllWagonHP': 'Caravan!C12:C25', 'AllWagonCargo': 'Caravan!E12:E25', 'AllWagonTravelers': 'Caravan!D12:D25', 'AllWagonConsumption': 'Caravan!F12:F25', 'AllTravelerJobs': 'Caravan!I12:I25',
    'PartyHeaders': 'Roster!A3:M3', 'PartyTable': 'Roster!A4:M15', 'MountHeaders': 'Roster!A18:K18', 'MountTable': 'Roster!A19:K30'
  };
  Object.entries(ranges).forEach(([name, rangeA1]) => {
    try {
      ss.setNamedRange(name, ss.getRange(rangeA1));
    } catch (e) {
      console.error(`Failed to create named range "${name}" for range "${rangeA1}". Error: ${e.message}`);
    }
  });
}

function addAllFormulas(ss) {
  const calc = ss.getSheetByName('Calculations');
  calc.getRange('B5').setFormula('=BASE_SPEED');
  calc.getRange('B6').setFormula('=IFNA(IF(PathType="Highway", VLOOKUP(Terrain, TerrainTable, 2, FALSE), IF(PathType="Road or Trail", VLOOKUP(Terrain, TerrainTable, 3, FALSE), VLOOKUP(Terrain, TerrainTable, 4, FALSE))), 1)');
  calc.getRange('B7').setFormula('=IFNA(VLOOKUP(TravelPace, PaceTable, 2, FALSE), 1)');
  calc.getRange('B8').setFormula('=IFNA(VLOOKUP(Weather, WeatherTable, 2, FALSE), 1)');
  calc.getRange('B9').setFormula('=CalcBaseSpeed * CalcTerrainMod * CalcPaceMod * CalcWeatherMod');
  calc.getRange('B13').setFormula('=COUNTIF(AllCharNames, "<>")');
  calc.getRange('C13').setFormula('=IF(TravelPace = "Fast", FAST_MULTIPLIER, IF(TravelPace = "Slow", SLOW_MULTIPLIER, NORMAL_MULTIPLIER))');
  calc.getRange('D13').setFormula('=FoodBase * FoodMod');
  calc.getRange('E13').setFormula('=IFERROR(IF(FoodDaily > 0, MIN(INT(FoodStock / FoodDaily), MAX_DAYS), MAX_DAYS), MAX_DAYS)');
  calc.getRange('F13').setFormula('=IF(FoodDaysLeft < DAYS_CRITICAL, "CRITICAL", IF(FoodDaysLeft < DAYS_LOW, "LOW", IF(FoodDaysLeft < DAYS_GOOD, "OK", "GOOD")))');
  calc.getRange('B14').setFormula('=COUNTIF(AllCharNames, "<>")');
  calc.getRange('C14').setFormula('=IFNA(VLOOKUP(Temperature, TempTable, 2, FALSE), 1)');
  calc.getRange('D14').setFormula('=WaterBase * WaterMod');
  calc.getRange('E14').setFormula('=IFERROR(IF(WaterDaily > 0, MIN(INT(WaterStock / WaterDaily), MAX_DAYS), MAX_DAYS), MAX_DAYS)');
  calc.getRange('F14').setFormula('=IF(WaterDaysLeft < DAYS_CRITICAL, "CRITICAL", IF(WaterDaysLeft < DAYS_LOW, "LOW", IF(WaterDaysLeft < DAYS_GOOD, "OK", "GOOD")))');
  calc.getRange('B15').setFormula('=COUNTIF(AllMountNames, "<>") * FODDER_PER_MOUNT');
  calc.getRange('C15').setFormula('=IF(AnimalsGrazing, 0, 1)');
  calc.getRange('D15').setFormula('=FodderBase * FodderMod');
  calc.getRange('E15').setFormula('=IFERROR(IF(FodderDaily > 0, MIN(INT(FodderStock / FodderDaily), MAX_DAYS), MAX_DAYS), MAX_DAYS)');
  calc.getRange('F15').setFormula('=IF(FodderDaysLeft < DAYS_CRITICAL, "CRITICAL", IF(FodderDaysLeft < DAYS_LOW, "LOW", IF(FodderDaysLeft < DAYS_GOOD, "OK", "GOOD")))');
  calc.getRange('B16').setFormula('=IFERROR(CaravanConsumption, 0)');
  calc.getRange('C16').setFormula('=1');
  calc.getRange('D16').setFormula('=ProvisionBase * ProvisionMod');
  calc.getRange('E16').setFormula('=IFERROR(IF(ProvisionDaily > 0, MIN(INT(ProvisionStock / ProvisionDaily), MAX_DAYS), MAX_DAYS), MAX_DAYS)');
  calc.getRange('F16').setFormula('=IF(ProvisionDaysLeft < DAYS_CRITICAL, "CRITICAL", IF(ProvisionDaysLeft < DAYS_LOW, "LOW", IF(ProvisionDaysLeft < DAYS_GOOD, "OK", "GOOD")))');
  const roster = ss.getSheetByName('Roster');
  for (let i = 1; i <= 4; i++) {
    const row = 3 + i;
    roster.getRange(`J${row}`).setFormula(`=IF(LEN(A${row}), IF(OR(G${row} > STARVATION_DAYS, H${row} > (DEHYDRATION_HOURS + C${row})), "YES", "NO"), "")`);
    roster.getRange(`K${row}`).setFormula(`=IF(J${row}="YES", IF(G${row} > STARVATION_DAYS, SURVIVAL_DC_BASE + G${row} - STARVATION_DAYS, SURVIVAL_DC_BASE + INT((H${row} - DEHYDRATION_HOURS - C${row}) / HOURS_PER_DAY)), 0)`);
    roster.getRange(`L${row}`).setFormula(`=IF(J${row}="YES", IF(G${row} > STARVATION_DAYS, "Starvation", "Dehydration"), "OK")`);
    roster.getRange(`M${row}`).setFormula(`=IF(LEN(A${row}), IF(F${row}>0, IF(E${row}/F${row} < HP_CRITICAL, "Critical", IF(E${row}/F${row} < HP_BLOODIED, "Bloodied", IF(E${row}/F${row} < HP_WOUNDED, "Wounded", "Healthy"))), "Healthy"), "")`);
  }
  roster.getRange('P4').setFormula('=COUNTIF(AllCharNames, "<>")');
  roster.getRange('P5').setFormula('=COUNTIF(AllMountNames, "<>")');
  roster.getRange('P6').setFormula('=COUNTIF(AllCharChecks, "YES")');
  roster.getRange('P7').setFormula('=COUNTIF(AllCharStatus, "Critical")');
  roster.getRange('P8').setFormula('=IFERROR(AVERAGE(FILTER(AllCharHP, NOT(ISBLANK(AllCharHP))) / FILTER(AllCharMaxHP, AllCharMaxHP>0)), 1)');
  roster.getRange('P10').setFormula('=PartySize + MountCount');
  roster.getRange('I19').setFormula('=AnimalsGrazing');
  roster.getRange('J19').setFormula('=IF(LEN(A19), IF(E19>0, IF(D19/E19 < HP_BLOODIED, "Injured", "Healthy"), "Healthy"), "")');
  const caravan = ss.getSheetByName('Caravan');
  caravan.getRange('G4').setFormula('=CaravanOffense + MIN(5, COUNTIF(AllTravelerJobs, "Guard"))');
  caravan.getRange('I4').setFormula('=10 + CaravanDefense');
  caravan.getRange('G5').setFormula('=CaravanOffense + MIN(5, COUNTIF(AllTravelerJobs, "Guide"))');
  caravan.getRange('I5').setFormula('=CaravanMorale + MIN(5, COUNTIF(AllTravelerJobs, "Entertainer"))');
  caravan.getRange('G6').setFormula('=BASE_SPEED + CaravanMobility * 4');
  caravan.getRange('I6').setFormula('=SUM(AllWagonHP)');
  caravan.getRange('G7').setFormula('=SUM(AllWagonCargo)');
  caravan.getRange('I7').setFormula('=SUM(AllWagonTravelers)');
  caravan.getRange('G8').setFormula('=IFERROR(SUM(AllWagonConsumption) + COUNTIF(AllTravelerJobs, "<>") / 2 - MIN(5, COUNTIF(AllTravelerJobs, "Cook")) * 2, 0)');
  const status = ss.getSheetByName('Status Monitor');
  status.getRange('B5').setFormula('=FoodDaysLeft');
  status.getRange('B6').setFormula('=WaterDaysLeft');
  status.getRange('B7').setFormula('=FodderDaysLeft');
  status.getRange('B8').setFormula('=ProvisionDaysLeft');
  status.getRange('C5').setFormula('=FoodStatus');
  status.getRange('C6').setFormula('=WaterStatus');
  status.getRange('C7').setFormula('=FodderStatus');
  status.getRange('C8').setFormula('=ProvisionStatus');
  status.getRange('G5').setFormula('=AvgHPPercent');
  status.getRange('G6').setFormula('=ChecksNeeded');
  status.getRange('G7').setFormula('=CriticalCount');
  status.getRange('A11').setFormula('=TRIM(CONCATENATE(IF(FoodStatus = "CRITICAL", "Food CRITICAL! ", ""), IF(WaterStatus = "CRITICAL", "Water CRITICAL! ", ""), IF(FodderStatus = "CRITICAL", "Fodder CRITICAL! ", ""), IF(ChecksNeeded > 0, ChecksNeeded & " checks needed! ", ""), IF(CriticalCount > 0, CriticalCount & " members critical! ", ""), IF(CaravanUnrest > CaravanMorale, "MUTINY RISK!", "")))');
  const exploration = ss.getSheetByName('Exploration');
  // Exploration DC lookup based on Territory CR (Table 4-1: Exploration DCs)
  exploration.getRange('B7').setFormula('=IFNA(VLOOKUP(TerritoryCR, D5:E24, 2, FALSE), 16 + TerritoryCR)');
  // Location discovery scores (Base + Terrain Mod + Hidden Mod)
  exploration.getRange('E13').setFormula('=B13+C13+D13');
  exploration.getRange('E14').setFormula('=B14+C14+D14');
  exploration.getRange('E15').setFormula('=IF(B15<>"", B15+C15+D15, "")');
  // Way Sign DCs based on complexity
  exploration.getRange('C24').setFormula('=IF(B24="Simple", TerritoryCR+10, IF(B24="Moderate", TerritoryCR+15, IF(B24="Complex", TerritoryCR+20, "")))');
  exploration.getRange('C25').setFormula('=IF(B25="Simple", TerritoryCR+10, IF(B25="Moderate", TerritoryCR+15, IF(B25="Complex", TerritoryCR+20, "")))');
  exploration.getRange('C26').setFormula('=IF(B26="Simple", TerritoryCR+10, IF(B26="Moderate", TerritoryCR+15, IF(B26="Complex", TerritoryCR+20, "")))');
}

// ============= HELPER AND SHEET SETUP FUNCTIONS =============

function createOrGetSheet(ss, name, color) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }
  sheet.setTabColor(color);
  return sheet;
}

function setupDashboard(sheet) {
  sheet.clear();
  sheet.getRange('A1:G1').merge().setValue('CAMPAIGN DASHBOARD').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  
  sheet.getRange('A3:C3').merge().setValue('CURRENT STATE').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A4:C5').setValues([['Current Day:', 1, ''], ['Total Miles:', 0, '']]);
  
  sheet.getRange('A7:C7').merge().setValue('ENVIRONMENT').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A8:C12').setValues([['Temperature:', 'Normal', ''], ['Terrain:', 'Plains', ''], ['Path Type:', 'Road or Trail', ''], ['Weather:', 'Clear', ''], ['Altitude:', 'Normal', '']]);
  
  sheet.getRange('E3:G3').merge().setValue('DAILY ACTIONS').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('E4:G7').setValues([['Travel Pace:', 'Normal', ''], ['Animals Grazing:', false, ''], ['Food Found:', 0, 'days'], ['Water Found:', 0, 'gallons']]);
  
  sheet.getRange('A13:C13').merge().setValue('RESOURCE STOCKS').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A14:C17').setValues([['Food:', 100, 'days'], ['Water:', 200, 'gallons'], ['Fodder:', 500, 'lbs'], ['Provisions:', 50, 'units']]);
  sheet.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Extreme Cold', 'Severe Cold', 'Cold', 'Normal', 'Hot', 'Severe Heat', 'Extreme Heat']).build());
  sheet.getRange('B9').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Desert, sandy', 'Forest', 'Hills', 'Jungle', 'Moor', 'Mountains', 'Plains', 'Swamp', 'Tundra, frozen']).build());
  sheet.getRange('B10').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Highway', 'Road or Trail', 'Trackless']).build());
  sheet.getRange('B11').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Clear', 'Light Rain', 'Heavy Rain', 'Storm', 'Snow', 'Blizzard']).build());
  sheet.getRange('F4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Slow', 'Normal', 'Fast']).build());
  sheet.getRange('F5').insertCheckboxes();
}

function setupCalculations(sheet) {
  sheet.clear();
  sheet.getRange('A1:F1').merge().setValue('LIVE CALCULATIONS').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange('A3:F3').merge().setValue('MOVEMENT CALCULATION').setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A4:C4').setValues([['Component', 'Value', 'Notes']]);
  sheet.getRange('A5:A9').setValues([['Base Speed'], ['Terrain Modifier'], ['Pace Modifier'], ['Weather Modifier'], ['MILES TODAY']]);
  sheet.getRange('A11:F11').merge().setValue('RESOURCE CONSUMPTION').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('A12:F12').setValues([['Resource', 'Base', 'Modifier', 'Daily Use', 'Days Left', 'Alert']]);
  sheet.getRange('A13:A16').setValues([['Food'], ['Water'], ['Fodder'], ['Provisions']]);
}

function setupRoster(sheet) {
  sheet.clear();
  sheet.getRange('A1:M1').merge().setValue('PARTY & MOUNT ROSTER').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange('A3:M3').setValues([['Name', 'Class', 'CON', 'Fort', 'HP', 'Max HP', 'Days No Food', 'Hours No Water', 'Hours No Sleep', 'Check?', 'DC', 'Type', 'Status']]).setFontWeight('bold').setBackground('#ffebee');
  sheet.getRange('A4:I7').setValues([
    ['Feldspar', 'Fighter', 14, 5, 52, 52, 0, 0, 0],
    ['Raven', 'Cleric', 12, 2, 38, 38, 0, 0, 0],
    ['Tharn', 'Rogue', 13, 3, 40, 40, 0, 0, 0],
    ['Tabah', 'Wizard', 10, 1, 30, 30, 0, 0, 0]
  ]);
  sheet.getRange('A18:K18').setValues([['Mount', 'Type', 'HD', 'HP', 'Max HP', 'Speed', 'Carrying', 'Max Load', 'Grazing?', 'Status', 'Notes']]).setFontWeight('bold').setBackground('#ffebee');
  sheet.getRange('A19:H19').setValues([['Thunder', 'Horse', 2, 15, 15, 50, 120, 300]]);
  sheet.getRange('O3:P9').setValues([
    ['Party Size:', ''], 
    ['Mount Count:', ''], 
    ['Checks Needed:', ''], 
    ['Critical Members:', ''], 
    ['Average HP%:', ''], 
    ['', ''], // Corrected Spacer Row
    ['Total Size:', '']
  ]).setFontWeight('bold');
}

function setupStatus(sheet) {
  sheet.clear();
  sheet.getRange('A1:H1').merge().setValue('STATUS MONITOR').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange('A3:D3').merge().setValue('RESOURCE STATUS').setFontWeight('bold').setBackground('#fff9c4');
  sheet.getRange('A4:D4').setValues([['Resource', 'Days Left', 'Status', '']]);
  sheet.getRange('A5:A8').setValues([['Food'], ['Water'], ['Fodder'], ['Provisions']]);
  sheet.getRange('F3:H3').merge().setValue('PARTY HEALTH').setFontWeight('bold').setBackground('#fff9c4');
  sheet.getRange('F4:G7').setValues([['Metric', 'Value'], ['Average HP', ''], ['Checks Needed', ''], ['Critical Members', '']]);
}

function setupConstants(sheet) {
  sheet.clear();
  sheet.getRange('A1:B1').merge().setValue('GAME CONSTANTS').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2:B28').setValues([
    ['Starvation Days:', 3], ['Dehydration Hours:', 24], ['Survival DC Base:', 10], ['Exhaustion Hours:', 48], ['', ''],
    ['HP Critical %:', 0.25], ['HP Bloodied %:', 0.50], ['HP Wounded %:', 0.75], ['', ''],
    ['Days Critical:', 3], ['Days Low:', 7], ['Days Good:', 14], ['', ''],
    ['Base Speed (miles/day):', 24], ['Travel Hours/Day:', 8], ['Fodder per Mount:', 20], ['Fast Multiplier:', 1.5], ['Slow Multiplier:', 0.75], ['Normal Multiplier:', 1.0], ['', ''],
    ['Hot Water Mult:', 2.0], ['Severe Heat Mult:', 4.0], ['Extreme Heat Mult:', 6.0], ['Cold Water Mult:', 1.0], ['', ''],
    ['Max Days Display:', 999], ['Hours per Day:', 24], ['Unrest Penalty Mult:', 1], ['', '']
  ]);
}

function setupReferences(sheet) {
  sheet.clear();
  sheet.getRange('A1:J1').merge().setValue('REFERENCE TABLES - Pathfinder 1E Rules').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  
  // Terrain Table (Table 7-8: Terrain and Overland Movement)
  sheet.getRange('A2:D2').setValues([['TERRAIN', 'Highway', 'Road or Trail', 'Trackless']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A3:D12').setValues([
    ['Desert, sandy', 1.0, 0.5, 0.5],
    ['Forest', 1.0, 1.0, 0.5],
    ['Hills', 1.0, 0.75, 0.5],
    ['Jungle', 1.0, 0.75, 0.25],
    ['Moor', 1.0, 1.0, 0.75],
    ['Mountains', 0.75, 0.75, 0.5],
    ['Plains', 1.0, 1.0, 0.75],
    ['Swamp', 1.0, 0.75, 0.5],
    ['Tundra, frozen', 1.0, 0.75, 0.75]
  ]);
  
  // Pace Table
  sheet.getRange('F2:G2').setValues([['PACE', 'MODIFIER']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('F3:G5').setValues([['Slow', 0.75], ['Normal', 1.0], ['Fast', 1.25]]);
  
  // Temperature Table
  sheet.getRange('I2:J2').setValues([['TEMPERATURE', 'WATER MULT']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('I3:J9').setValues([
    ['Extreme Cold', 1.0],
    ['Severe Cold', 1.0],
    ['Cold', 1.0],
    ['Normal', 1.0],
    ['Hot', 2.0],
    ['Severe Heat', 4.0],
    ['Extreme Heat', 6.0]
  ]);
  
  // Weather Table
  sheet.getRange('A14:B14').setValues([['WEATHER', 'MODIFIER']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A15:B20').setValues([
    ['Clear', 1.0],
    ['Light Rain', 0.9],
    ['Heavy Rain', 0.75],
    ['Storm', 0.5],
    ['Snow', 0.5],
    ['Blizzard', 0.25]
  ]);
  
  // Foraging Table
  sheet.getRange('D14:E14').setValues([['FORAGING', 'DC']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('D15:E19').setValues([
    ['Abundant', 10],
    ['Average', 15],
    ['Sparse', 20],
    ['Barren', 25],
    ['Desolate', 30]
  ]);
  
  // Altitude Table
  sheet.getRange('G14:H14').setValues([['ALTITUDE', 'MODIFIER']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('G15:H18').setValues([
    ['Normal', 1.0],
    ['High', 0.9],
    ['Very High', 0.75],
    ['Extreme', 0.5]
  ]);
}

function setupCaravan(sheet) {
  sheet.clear();
  sheet.getRange('A1:J1').merge().setValue('CARAVAN MANAGEMENT').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange('A3:D3').merge().setValue('PRIMARY STATS').setFontWeight('bold').setBackground('#ffe0b2');
  sheet.getRange('A4:D8').setValues([['Name:', 'Desert Wind', 'Level:', 5], ['Offense:', 3, 'Defense:', 2], ['Mobility:', 2, 'Morale:', 4], ['Unrest:', 0, 'Max Unrest:', 10], ['Fortune:', 2, 'Prestige:', 0]]);
  sheet.getRange('F3:I3').merge().setValue('DERIVED STATS').setFontWeight('bold').setBackground('#ffe0b2');
  sheet.getRange('F4:I8').setValues([['Attack:', '', 'AC:', ''], ['Security:', '', 'Resolve:', ''], ['Speed:', '', 'HP:', ''], ['Cargo:', '', 'Travelers:', ''], ['Consumption:', '', '', '']]);
  sheet.getRange('A10:F10').merge().setValue('WAGONS').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('A11:F11').setValues([['Name', 'Type', 'HP', 'Travelers', 'Cargo', 'Consumption']]).setFontWeight('bold');
  sheet.getRange('A12:F14').setValues([['Supply', 'Covered', 30, 6, 4000, 2], ['Fortune', 'Armored', 40, 4, 2000, 3], ['Royal', 'Luxury', 50, 8, 1000, 4]]);
  sheet.getRange('H10:K10').merge().setValue('TRAVELERS').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('H11:K11').setValues([['Name', 'Job', 'PC?', 'Notes']]).setFontWeight('bold');
  sheet.getRange('H12:K14').setValues([['John', 'Driver', false, ''], ['Sarah', 'Cook', false, ''], ['Marcus', 'Guard', false, '']]);
}

function setupExploration(sheet) {
  sheet.clear();
  sheet.getRange('A1:J1').merge().setValue('EXPLORATION TRACKER - Pathfinder 1E Rules').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  
  // Territory Info Section
  sheet.getRange('A3:B3').merge().setValue('TERRITORY INFORMATION').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('A4:B9').setValues([
    ['Territory Name:', 'Darkwood'],
    ['Territory CR:', 5],
    ['Exploration Skill:', 'Survival'],
    ['Exploration DC:', ''],
    ['Current Discovery Points:', 0],
    ['Days Explored:', 0]
  ]);
  
  // Exploration DC Reference Table (Table 4-1)
  sheet.getRange('D3:F3').merge().setValue('EXPLORATION DC REFERENCE').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('D4:F4').setValues([['CR', 'DC', '']]).setFontWeight('bold');
  sheet.getRange('D5:F24').setValues([
    [1, 17, ''], [2, 19, ''], [3, 21, ''], [4, 22, ''], [5, 23, ''],
    [6, 24, ''], [7, 25, ''], [8, 26, ''], [9, 27, ''], [10, 28, ''],
    [11, 29, ''], [12, 30, ''], [13, 31, ''], [14, 32, ''], [15, 33, ''],
    [16, 34, ''], [17, 35, ''], [18, 36, ''], [19, 37, ''], [20, 38, '']
  ]);
  
  // Locations Section
  sheet.getRange('A11').setValue('LOCATIONS').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('A12:F12').setValues([['Name', 'Base Score', 'Terrain Mod', 'Hidden Mod', 'Final Score', 'Status']]).setFontWeight('bold');
  sheet.getRange('A13:F15').setValues([
    ['Hidden Temple', 6, 2, 4, '', 'Undiscovered'],
    ['Ancient Ruins', 3, 2, 0, '', 'Undiscovered'],
    ['', '', '', '', '', '']
  ]);
  
  // Discovery Score Modifiers Reference (Table 4-2)
  sheet.getRange('H11').setValue('DISCOVERY SCORE MODIFIERS').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('H12:I12').setValues([['Condition', 'Modifier']]).setFontWeight('bold');
  sheet.getRange('H13:I20').setValues([
    ['Desert or plains terrain', '+1'],
    ['Forest, hills, or marsh terrain', '+2'],
    ['Mountain terrain', '+3'],
    ['Location is traveled to/from often', '-4'],
    ['Location is mobile', '+4'],
    ['Location is unusually large', '-2'],
    ['Location is unusually small', '+2'],
    ['Location is deliberately hidden', '+2 to +6']
  ]);
  
  // Way Signs Section
  sheet.getRange('A22').setValue('WAY SIGNS').setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A23:F23').setValues([['Description', 'Complexity', 'DC', 'Discovery Points', 'Status', 'Notes']]).setFontWeight('bold');
  sheet.getRange('A24:F26').setValues([
    ['Old Map Found', 'Simple', '', 1, 'Undiscovered', 'CR + 10'],
    ['Traveler\'s Journal', 'Moderate', '', 3, 'Undiscovered', 'CR + 15'],
    ['Aerial Reconnaissance', 'Complex', '', 5, 'Undiscovered', 'CR + 20']
  ]);
  
  // Exploration Actions Log
  sheet.getRange('A28').setValue('EXPLORATION LOG').setFontWeight('bold').setBackground('#f3e5f5');
  sheet.getRange('A29:F29').setValues([['Day', 'Action', 'Skill Check', 'Result', 'DP Gained', 'Notes']]).setFontWeight('bold');
}

function setupLog(sheet) {
  sheet.clear();
  sheet.getRange('A1:H1').setValues([['Day', 'Date', 'Miles', 'Food', 'Water', 'Fodder', 'Provisions', 'Notes']]).setFontWeight('bold').setBackground('#e1bee7');
  sheet.setFrozenRows(1);
}

function setupHexMap(sheet) {
  sheet.clear();
  sheet.getRange('A1:F1').merge().setValue('HEX MAP TRACKER').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  
  // Current Location
  sheet.getRange('A3:B3').merge().setValue('CURRENT LOCATION').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('A4:B6').setValues([
    ['Current Hex (q:r:s):', '0:0:0'],
    ['Hex Terrain:', ''],
    ['Pathfinder Terrain:', '']
  ]);
  
  // Hex Map Data
  sheet.getRange('D3:F3').merge().setValue('HEX MAP DATA').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('D4:F4').setValues([['Hex (q:r:s)', 'Map Terrain', 'Pathfinder Terrain']]).setFontWeight('bold');
  
  // Terrain Mapping Reference
  sheet.getRange('A8:B8').merge().setValue('TERRAIN MAPPING').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('A9:B9').setValues([['Map Terrain', 'Pathfinder Terrain']]).setFontWeight('bold');
  sheet.getRange('A10:B30').setValues([
    ['plains', 'Plains'],
    ['farmland', 'Plains'],
    ['grassland', 'Plains'],
    ['scrubland', 'Plains'],
    ['forest', 'Forest'],
    ['pine-forest', 'Forest'],
    ['dense-jungle', 'Jungle'],
    ['hills', 'Hills'],
    ['grassy-hills', 'Hills'],
    ['forest-hills', 'Hills'],
    ['pine-hills', 'Hills'],
    ['jungle-hills', 'Hills'],
    ['mountains', 'Mountains'],
    ['pine-mountains', 'Mountains'],
    ['ice-mountains', 'Mountains'],
    ['desert', 'Desert, sandy'],
    ['badlands', 'Desert, sandy'],
    ['swamp', 'Swamp'],
    ['marsh', 'Swamp'],
    ['water', 'Plains'], // Water hexes - special handling needed
    ['deep-water', 'Plains']
  ]);
}

function setupEventLog(sheet) {
  // Only set up if the sheet is empty or doesn't have headers
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1:E1').setValues([['Timestamp', 'Day', 'Event Type', 'Description', 'Details']]).setFontWeight('bold').setBackground('#9c27b0');
    sheet.setFrozenRows(1);
    // Format header row
    sheet.getRange('A1:E1').setFontColor('#ffffff');
    // Set column widths
    sheet.setColumnWidth(1, 150); // Timestamp
    sheet.setColumnWidth(2, 60);  // Day
    sheet.setColumnWidth(3, 120); // Event Type
    sheet.setColumnWidth(4, 250); // Description
    sheet.setColumnWidth(5, 300); // Details
  }
}

/**
 * Logs an event to the Event Log sheet. This log persists across campaign resets.
 * @param {string} eventType The type of event (e.g., 'Day Processed', 'Location Found', 'Combat', etc.)
 * @param {string} description A brief description of the event
 * @param {string} details Optional additional details about the event
 */
function logEvent(eventType, description, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let eventLogSheet = ss.getSheetByName('Event Log');
    
    // Create Event Log if it doesn't exist
    if (!eventLogSheet) {
      eventLogSheet = createOrGetSheet(ss, 'Event Log', '#9c27b0');
      setupEventLog(eventLogSheet);
    }
    
    const currentDay = ss.getRangeByName('CurrentDay') ? ss.getRangeByName('CurrentDay').getValue() : 0;
    const timestamp = new Date();
    
    // Append the event
    eventLogSheet.appendRow([
      timestamp,
      currentDay,
      eventType,
      description,
      details || ''
    ]);
    
    // Format the new row
    const lastRow = eventLogSheet.getLastRow();
    eventLogSheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    eventLogSheet.getRange(lastRow, 2).setHorizontalAlignment('center');
    
  } catch (e) {
    console.error('Error logging event:', e);
  }
}

/**
 * Imports hex map data from JSON and populates the Hex Map sheet.
 * @param {string} jsonData The JSON string containing hex map data
 * @returns {string} Confirmation message
 */
function importHexMap(jsonData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hexMapSheet = ss.getSheetByName('Hex Map');
    if (!hexMapSheet) {
      return 'Error: Hex Map sheet not found. Please initialize the campaign first.';
    }
    
    const mapData = JSON.parse(jsonData);
    const hexes = mapData.terrain?.hexes || {};
    
    // Clear existing hex data (keep headers and current location)
    const lastRow = hexMapSheet.getLastRow();
    if (lastRow > 10) {
      hexMapSheet.deleteRows(11, lastRow - 10);
    }
    
    // Prepare hex data for import
    const hexRows = [];
    Object.keys(hexes).forEach(hexKey => {
      const hex = hexes[hexKey];
      const mapTerrain = hex.tile?.tile_name || '';
      const pfTerrain = mapTerrainToPathfinder(mapTerrain);
      hexRows.push([
        `${hex.q}:${hex.r}:${hex.s}`,
        mapTerrain,
        pfTerrain
      ]);
    });
    
    // Sort by q, then r, then s
    hexRows.sort((a, b) => {
      const [q1, r1, s1] = a[0].split(':').map(Number);
      const [q2, r2, s2] = b[0].split(':').map(Number);
      if (q1 !== q2) return q1 - q2;
      if (r1 !== r2) return r1 - r2;
      return s1 - s2;
    });
    
    // Write hex data
    if (hexRows.length > 0) {
      hexMapSheet.getRange(5, 4, hexRows.length, 3).setValues(hexRows);
    }
    
    logEvent('Hex Map Imported', `Imported ${hexRows.length} hexes from ${mapData.metadata?.title || 'map'}`, '');
    
    return `Successfully imported ${hexRows.length} hexes from the map.`;
  } catch (e) {
    return `Error importing hex map: ${e.message}`;
  }
}

/**
 * Maps hex map terrain types to Pathfinder terrain types.
 * @param {string} mapTerrain The terrain type from the hex map
 * @returns {string} The corresponding Pathfinder terrain type
 */
function mapTerrainToPathfinder(mapTerrain) {
  const terrainMap = {
    'plains': 'Plains',
    'farmland': 'Plains',
    'grassland': 'Plains',
    'scrubland': 'Plains',
    'forest': 'Forest',
    'pine-forest': 'Forest',
    'tree': 'Forest',
    'dense-jungle': 'Jungle',
    'jungle-tree': 'Jungle',
    'hills': 'Hills',
    'grassy-hills': 'Hills',
    'forest-hills': 'Hills',
    'pine-hills': 'Hills',
    'jungle-hills': 'Hills',
    'ice-hills': 'Hills',
    'mountains': 'Mountains',
    'pine-mountains': 'Mountains',
    'ice-mountains': 'Mountains',
    'ice-mountain': 'Mountains',
    'ice-pine-mountains': 'Mountains',
    'forest-mountain': 'Mountains',
    'dead-tree-mountains': 'Mountains',
    'jungle-mountains': 'Mountains',
    'desert': 'Desert, sandy',
    'badlands': 'Desert, sandy',
    'swamp': 'Swamp',
    'marsh': 'Swamp',
    'water': 'Plains', // Special handling - water hexes
    'deep-water': 'Plains',
    'beach': 'Plains',
    'icy': 'Tundra, frozen',
    'ice-tree': 'Tundra, frozen',
    'volcano': 'Mountains',
    'dead-tree': 'Forest',
    'dead-tree-hills': 'Hills'
  };
  
  return terrainMap[mapTerrain] || 'Plains';
}

/**
 * Updates the current hex location and automatically updates terrain.
 * @param {string} hexCoords Hex coordinates in format "q:r:s"
 * @returns {string} Confirmation message
 */
function setCurrentHex(hexCoords) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hexMapSheet = ss.getSheetByName('Hex Map');
    if (!hexMapSheet) {
      return 'Error: Hex Map sheet not found.';
    }
    
    // Update current hex
    ss.getRangeByName('CurrentHex').setValue(hexCoords);
    
    // Find hex data
    const hexDataRange = hexMapSheet.getRange('D5:F' + hexMapSheet.getLastRow());
    const hexData = hexDataRange.getValues();
    const hexRow = hexData.find(row => row[0] === hexCoords);
    
    if (hexRow) {
      const mapTerrain = hexRow[1];
      const pfTerrain = hexRow[2] || mapTerrainToPathfinder(mapTerrain);
      
      // Update hex terrain display
      ss.getRangeByName('HexTerrain').setValue(mapTerrain);
      
      // Auto-update Pathfinder terrain if it's different
      const currentTerrain = ss.getRangeByName('Terrain').getValue();
      if (pfTerrain && pfTerrain !== currentTerrain) {
        ss.getRangeByName('Terrain').setValue(pfTerrain);
        logEvent('Location Changed', `Moved to hex ${hexCoords}`, `Terrain auto-updated to ${pfTerrain}`);
      } else {
        logEvent('Location Changed', `Moved to hex ${hexCoords}`, `Terrain: ${mapTerrain}`);
      }
      
      return `Location updated to hex ${hexCoords}. Terrain: ${pfTerrain}`;
    } else {
      ss.getRangeByName('HexTerrain').setValue('Unknown');
      logEvent('Location Changed', `Moved to hex ${hexCoords}`, 'Hex not found in map data');
      return `Location updated to hex ${hexCoords}, but hex not found in map data.`;
    }
  } catch (e) {
    return `Error setting hex location: ${e.message}`;
  }
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

