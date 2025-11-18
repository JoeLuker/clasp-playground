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

// ============= DATA GETTER FUNCTIONS FOR SIDEBAR =============

/**
 * Gets all the key data points from the spreadsheet to display on the sidebar.
 * @returns {object} An object containing all the necessary dashboard and status data.
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      currentDay: ss.getRangeByName('CurrentDay').getValue(),
      totalMiles: ss.getRangeByName('TotalMiles').getValue(),
      temperature: ss.getRangeByName('Temperature').getValue(),
      terrain: ss.getRangeByName('Terrain').getValue(),
      weather: ss.getRangeByName('Weather').getValue(),
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
    return { error: "Sheet not initialized. Please run 'Initialize/Reset Campaign' from the menu." };
  }
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
  ss.getRangeByName('CurrentDay').setValue(currentDay + 1);
  ss.getRangeByName('TotalMiles').setValue(ss.getRangeByName('TotalMiles').getValue() + miles);
  ss.getRangeByName('FoodStock').setValue(
    ss.getRangeByName('FoodStock').getValue() - ss.getRangeByName('FoodDaily').getValue() + ss.getRangeByName('FoodFound').getValue()
  );
  ss.getRangeByName('WaterStock').setValue(
    ss.getRangeByName('WaterStock').getValue() - ss.getRangeByName('WaterDaily').getValue() + ss.getRangeByName('WaterFound').getValue()
  );
  ss.getRangeByName('FodderStock').setValue(
    ss.getRangeByName('FodderStock').getValue() - ss.getRangeByName('FodderDaily').getValue()
  );
  ss.getRangeByName('ProvisionStock').setValue(
    ss.getRangeByName('ProvisionStock').getValue() - ss.getRangeByName('ProvisionDaily').getValue()
  );
  updateDeprivation(ss);
  const log = ss.getSheetByName('Log');
  log.appendRow([
    currentDay + 1, new Date(), miles,
    ss.getRangeByName('FoodDaily').getValue(),
    ss.getRangeByName('WaterDaily').getValue(),
    ss.getRangeByName('FodderDaily').getValue(),
    ss.getRangeByName('ProvisionDaily').getValue(),
    ''
  ]);
  ss.getRangeByName('FoodFound').setValue(0);
  ss.getRangeByName('WaterFound').setValue(0);
  return `Advanced to Day ${currentDay + 1}. Traveled ${Math.round(miles)} miles.`;
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
    'This will DELETE ALL EXISTING SHEETS in this spreadsheet and build a completely new campaign manager. Are you sure you want to continue?',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const tempSheet = ss.insertSheet('temp');
  const allSheets = ss.getSheets();
  allSheets.forEach(sheet => {
    if (sheet.getName() !== 'temp') {
      ss.deleteSheet(sheet);
    }
  });
  createAllSheets(ss);
  createAllNamedRanges(ss);
  addAllFormulas(ss);
  
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
  setupDashboard(dashboard);
  setupCalculations(calculations);
  setupRoster(roster);
  setupStatus(status);
  setupConstants(constants);
  setupReferences(references);
  setupCaravan(caravan);
  setupExploration(exploration);
  setupLog(log);
}

function createAllNamedRanges(ss) {
    const ranges = {
    'STARVATION_DAYS': 'Constants!B2', 'DEHYDRATION_HOURS': 'Constants!B3', 'SURVIVAL_DC_BASE': 'Constants!B4', 'EXHAUSTION_HOURS': 'Constants!B5', 'HP_CRITICAL': 'Constants!B7', 'HP_BLOODIED': 'Constants!B8', 'HP_WOUNDED': 'Constants!B9', 'DAYS_CRITICAL': 'Constants!B11', 'DAYS_LOW': 'Constants!B12', 'DAYS_GOOD': 'Constants!B13', 'BASE_SPEED': 'Constants!B15', 'FODDER_PER_MOUNT': 'Constants!B16', 'FAST_MULTIPLIER': 'Constants!B17', 'SLOW_MULTIPLIER': 'Constants!B18', 'NORMAL_MULTIPLIER': 'Constants!B19', 'HOT_WATER_MULT': 'Constants!B21', 'SEVERE_HEAT_MULT': 'Constants!B22', 'EXTREME_HEAT_MULT': 'Constants!B23', 'COLD_WATER_MULT': 'Constants!B24', 'MAX_DAYS': 'Constants!B26', 'HOURS_PER_DAY': 'Constants!B27', 'UNREST_PENALTY_MULT': 'Constants!B28',
    'CurrentDay': 'Dashboard!B4', 'TotalMiles': 'Dashboard!B5', 'Temperature': 'Dashboard!B8', 'Terrain': 'Dashboard!B9', 'Weather': 'Dashboard!B10', 'Altitude': 'Dashboard!B11', 'TravelPace': 'Dashboard!F4', 'AnimalsGrazing': 'Dashboard!F5', 'FoodFound': 'Dashboard!F6', 'WaterFound': 'Dashboard!F7', 'FoodStock': 'Dashboard!B14', 'WaterStock': 'Dashboard!B15', 'FodderStock': 'Dashboard!B16', 'ProvisionStock': 'Dashboard!B17',
    'CalcBaseSpeed': 'Calculations!B5', 'CalcTerrainMod': 'Calculations!B6', 'CalcPaceMod': 'Calculations!B7', 'CalcWeatherMod': 'Calculations!B8', 'CalcMilesToday': 'Calculations!B9', 'FoodBase': 'Calculations!B13', 'FoodMod': 'Calculations!C13', 'FoodDaily': 'Calculations!D13', 'FoodDaysLeft': 'Calculations!E13', 'FoodStatus': 'Calculations!F13', 'WaterBase': 'Calculations!B14', 'WaterMod': 'Calculations!C14', 'WaterDaily': 'Calculations!D14', 'WaterDaysLeft': 'Calculations!E14', 'WaterStatus': 'Calculations!F14', 'FodderBase': 'Calculations!B15', 'FodderMod': 'Calculations!C15', 'FodderDaily': 'Calculations!D15', 'FodderDaysLeft': 'Calculations!E15', 'FodderStatus': 'Calculations!F15', 'ProvisionBase': 'Calculations!B16', 'ProvisionMod': 'Calculations!C16', 'ProvisionDaily': 'Calculations!D16', 'ProvisionDaysLeft': 'Calculations!E16', 'ProvisionStatus': 'Calculations!F16',
    'Char1_Name': 'Roster!A4', 'Char1_Class': 'Roster!B4', 'Char1_CON': 'Roster!C4', 'Char1_Fort': 'Roster!D4', 'Char1_HP': 'Roster!E4', 'Char1_MaxHP': 'Roster!F4', 'Char1_DaysNoFood': 'Roster!G4', 'Char1_HoursNoWater': 'Roster!H4', 'Char1_HoursNoSleep': 'Roster!I4', 'Char1_CheckNeeded': 'Roster!J4', 'Char1_CheckDC': 'Roster!K4', 'Char1_CheckType': 'Roster!L4', 'Char1_Status': 'Roster!M4',
    'Char2_Name': 'Roster!A5', 'Char2_Class': 'Roster!B5', 'Char2_CON': 'Roster!C5', 'Char2_Fort': 'Roster!D5', 'Char2_HP': 'Roster!E5', 'Char2_MaxHP': 'Roster!F5', 'Char2_DaysNoFood': 'Roster!G5', 'Char2_HoursNoWater': 'Roster!H5', 'Char2_HoursNoSleep': 'Roster!I5', 'Char2_CheckNeeded': 'Roster!J5', 'Char2_CheckDC': 'Roster!K5', 'Char2_CheckType': 'Roster!L5', 'Char2_Status': 'Roster!M5',
    'Char3_Name': 'Roster!A6', 'Char3_Class': 'Roster!B6', 'Char3_CON': 'Roster!C6', 'Char3_Fort': 'Roster!D6', 'Char3_HP': 'Roster!E6', 'Char3_MaxHP': 'Roster!F6', 'Char3_DaysNoFood': 'Roster!G6', 'Char3_HoursNoWater': 'Roster!H6', 'Char3_HoursNoSleep': 'Roster!I6', 'Char3_CheckNeeded': 'Roster!J6', 'Char3_CheckDC': 'Roster!K6', 'Char3_CheckType': 'Roster!L6', 'Char3_Status': 'Roster!M6',
    'Char4_Name': 'Roster!A7', 'Char4_Class': 'Roster!B7', 'Char4_CON': 'Roster!C7', 'Char4_Fort': 'Roster!D7', 'Char4_HP': 'Roster!E7', 'Char4_MaxHP': 'Roster!F7', 'Char4_DaysNoFood': 'Roster!G7', 'Char4_HoursNoWater': 'Roster!H7', 'Char4_HoursNoSleep': 'Roster!I7', 'Char4_CheckNeeded': 'Roster!J7', 'Char4_CheckDC': 'Roster!K7', 'Char4_CheckType': 'Roster!L7', 'Char4_Status': 'Roster!M7',
    'PartySize': 'Roster!P4', 'MountCount': 'Roster!P5', 'ChecksNeeded': 'Roster!P6', 'CriticalCount': 'Roster!P7', 'AvgHPPercent': 'Roster!P8', 'TotalSize': 'Roster!P10',
    'Mount1_Name': 'Roster!A19', 'Mount1_Type': 'Roster!B19', 'Mount1_HD': 'Roster!C19', 'Mount1_HP': 'Roster!D19', 'Mount1_MaxHP': 'Roster!E19', 'Mount1_Speed': 'Roster!F19', 'Mount1_Carrying': 'Roster!G19', 'Mount1_MaxLoad': 'Roster!H19', 'Mount1_Grazing': 'Roster!I19', 'Mount1_Status': 'Roster!J19',
    'CaravanName': 'Caravan!B4', 'CaravanLevel': 'Caravan!D4', 'CaravanOffense': 'Caravan!B5', 'CaravanDefense': 'Caravan!D5', 'CaravanMobility': 'Caravan!B6', 'CaravanMorale': 'Caravan!D6', 'CaravanUnrest': 'Caravan!B7', 'CaravanMaxUnrest': 'Caravan!D7', 'CaravanFortune': 'Caravan!B8', 'CaravanPrestige': 'Caravan!D8', 'CaravanAttack': 'Caravan!G4', 'CaravanAC': 'Caravan!I4', 'CaravanSecurity': 'Caravan!G5', 'CaravanResolve': 'Caravan!I5', 'CaravanSpeed': 'Caravan!G6', 'CaravanHP': 'Caravan!I6', 'CaravanCargo': 'Caravan!G7', 'CaravanTravelers': 'Caravan!I7', 'CaravanConsumption': 'Caravan!G8',
    'TerritoryName': 'Exploration!B3', 'TerritoryCR': 'Exploration!B4', 'ExplorationDC': 'Exploration!B5', 'ExplorationSkill': 'Exploration!B6', 'CurrentDP': 'Exploration!B7', 'DaysExplored': 'Exploration!B8',
    'StatusFood': 'Status Monitor!B5', 'StatusWater': 'Status Monitor!B6', 'StatusFodder': 'Status Monitor!B7', 'StatusProvision': 'Status Monitor!B8', 'AlertFood': 'Status Monitor!C5', 'AlertWater': 'Status Monitor!C6', 'AlertFodder': 'Status Monitor!C7', 'AlertProvision': 'Status Monitor!C8', 'PartyAvgHP': 'Status Monitor!G5', 'PartyChecks': 'Status Monitor!G6', 'PartyCritical': 'Status Monitor!G7', 'AlertText': 'Status Monitor!A11',
    'TerrainTable': 'Reference Tables!A3:B10', 'PaceTable': 'Reference Tables!D3:E5', 'TempTable': 'Reference Tables!G3:H9', 'WeatherTable': 'Reference Tables!A13:B18', 'ForageTable': 'Reference Tables!D13:E17', 'AltitudeTable': 'Reference Tables!G13:H16',
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
  calc.getRange('B6').setFormula('=IFNA(VLOOKUP(Terrain, TerrainTable, 2, FALSE), 1)');
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
  exploration.getRange('B5').setFormula('=16 + TerritoryCR');
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
  sheet.getRange('A8:C11').setValues([['Temperature:', 'Normal', ''], ['Terrain:', 'Plains', ''], ['Weather:', 'Clear', ''], ['Altitude:', 'Normal', '']]);
  
  sheet.getRange('E3:G3').merge().setValue('DAILY ACTIONS').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('E4:G7').setValues([['Travel Pace:', 'Normal', ''], ['Animals Grazing:', false, ''], ['Food Found:', 0, 'days'], ['Water Found:', 0, 'gallons']]);
  
  sheet.getRange('A13:C13').merge().setValue('RESOURCE STOCKS').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A14:C17').setValues([['Food:', 100, 'days'], ['Water:', 200, 'gallons'], ['Fodder:', 500, 'lbs'], ['Provisions:', 50, 'units']]);
  sheet.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Extreme Cold', 'Severe Cold', 'Cold', 'Normal', 'Hot', 'Severe Heat', 'Extreme Heat']).build());
  sheet.getRange('B9').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Plains', 'Forest', 'Hills', 'Mountains', 'Desert', 'Swamp', 'Road', 'Jungle']).build());
  sheet.getRange('B10').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Clear', 'Light Rain', 'Heavy Rain', 'Storm', 'Snow', 'Blizzard']).build());
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
    ['Base Speed:', 24], ['Fodder per Mount:', 20], ['Fast Multiplier:', 1.5], ['Slow Multiplier:', 0.75], ['Normal Multiplier:', 1.0], ['', ''],
    ['Hot Water Mult:', 2.0], ['Severe Heat Mult:', 4.0], ['Extreme Heat Mult:', 6.0], ['Cold Water Mult:', 1.0], ['', ''],
    ['Max Days Display:', 999], ['Hours per Day:', 24], ['Unrest Penalty Mult:', 1]
  ]);
}

function setupReferences(sheet) {
  sheet.clear();
  sheet.getRange('A1:H1').merge().setValue('REFERENCE TABLES').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange('A2:B10').setValues([['TERRAIN', 'MODIFIER'], ['Plains', 1.0], ['Forest', 0.5], ['Hills', 0.75], ['Mountains', 0.5], ['Desert', 0.75], ['Swamp', 0.5], ['Road', 1.25], ['Jungle', 0.25]]);
  sheet.getRange('D2:E5').setValues([['PACE', 'MODIFIER'], ['Slow', 0.75], ['Normal', 1.0], ['Fast', 1.25]]);
  sheet.getRange('G2:H9').setValues([['TEMPERATURE', 'WATER MULT'], ['Extreme Cold', 1.0], ['Severe Cold', 1.0], ['Cold', 1.0], ['Normal', 1.0], ['Hot', 2.0], ['Severe Heat', 4.0], ['Extreme Heat', 6.0]]);
  sheet.getRange('A12:B18').setValues([['WEATHER', 'MODIFIER'], ['Clear', 1.0], ['Light Rain', 0.9], ['Heavy Rain', 0.75], ['Storm', 0.5], ['Snow', 0.5], ['Blizzard', 0.25]]);
  sheet.getRange('D12:E17').setValues([['FORAGING', 'DC'], ['Abundant', 10], ['Average', 15], ['Sparse', 20], ['Barren', 25], ['Desolate', 30]]);
  sheet.getRange('G12:H16').setValues([['ALTITUDE', 'MODIFIER'], ['Normal', 1.0], ['High', 0.9], ['Very High', 0.75], ['Extreme', 0.5]]);
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
  sheet.getRange('A1:F1').merge().setValue('EXPLORATION TRACKER').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange('A3:B8').setValues([['Territory:', 'Darkwood'], ['CR:', 5], ['DC:', ''], ['Skill:', 'Survival'], ['Current DP:', 0], ['Days Explored:', 0]]);
  sheet.getRange('A10').setValue('LOCATIONS').setFontWeight('bold');
  sheet.getRange('A11:E11').setValues([['Name', 'Base', 'Modifiers', 'Final', 'Status']]);
  sheet.getRange('A20').setValue('WAY SIGNS').setFontWeight('bold');
  sheet.getRange('A21:D21').setValues([['Description', 'Complexity', 'DP', 'Status']]);
}

function setupLog(sheet) {
  sheet.clear();
  sheet.getRange('A1:H1').setValues([['Day', 'Date', 'Miles', 'Food', 'Water', 'Fodder', 'Provisions', 'Notes']]).setFontWeight('bold').setBackground('#e1bee7');
  sheet.setFrozenRows(1);
}

