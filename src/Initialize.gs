/**
 * Campaign Initialization
 * Full system initialization, sheet creation, named ranges, and formulas.
 */

function initializeCompleteCampaign() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm Full Reset',
    'This will DELETE ALL EXISTING SHEETS in this spreadsheet and build a completely new campaign manager. The Event Log will be preserved. Are you sure you want to continue?',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  let eventLogSheet = ss.getSheetByName(SHEET_NAMES.EVENT_LOG);
  let eventLogData = null;
  if (eventLogSheet) {
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

  if (eventLogSheet) {
    setupEventLog(eventLogSheet);
    if (eventLogData && eventLogData.length > 1) {
      eventLogSheet.getRange(2, 1, eventLogData.length - 1, eventLogData[0].length).setValues(eventLogData.slice(1));
    }
  } else {
    eventLogSheet = createOrGetSheet(ss, 'Event Log', '#9c27b0');
    setupEventLog(eventLogSheet);
  }

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

function createAllNamedRanges(ss) {
  const ranges = {
    // Constants
    'STARVATION_DAYS': 'Constants!B2', 'DEHYDRATION_HOURS': 'Constants!B3', 'SURVIVAL_DC_BASE': 'Constants!B4',
    'EXHAUSTION_HOURS': 'Constants!B5', 'HP_CRITICAL': 'Constants!B7', 'HP_BLOODIED': 'Constants!B8',
    'HP_WOUNDED': 'Constants!B9', 'DAYS_CRITICAL': 'Constants!B11', 'DAYS_LOW': 'Constants!B12',
    'DAYS_GOOD': 'Constants!B13', 'BASE_SPEED': 'Constants!B15', 'TRAVEL_HOURS_PER_DAY': 'Constants!B16',
    'FODDER_PER_MOUNT': 'Constants!B17', 'FAST_MULTIPLIER': 'Constants!B18', 'SLOW_MULTIPLIER': 'Constants!B19',
    'NORMAL_MULTIPLIER': 'Constants!B20', 'HOT_WATER_MULT': 'Constants!B22', 'SEVERE_HEAT_MULT': 'Constants!B23',
    'EXTREME_HEAT_MULT': 'Constants!B24', 'COLD_WATER_MULT': 'Constants!B25', 'MAX_DAYS': 'Constants!B27',
    'HOURS_PER_DAY': 'Constants!B28', 'UNREST_PENALTY_MULT': 'Constants!B29',

    // Dashboard
    'CurrentDay': 'Dashboard!B4', 'TotalMiles': 'Dashboard!B5', 'CurrentHex': 'Hex Map!B4',
    'HexTerrain': 'Hex Map!B5', 'Temperature': 'Dashboard!B8', 'Terrain': 'Dashboard!B9',
    'PathType': 'Dashboard!B10', 'Weather': 'Dashboard!B11', 'Altitude': 'Dashboard!B12',
    'TravelPace': 'Dashboard!F4', 'AnimalsGrazing': 'Dashboard!F5', 'FoodFound': 'Dashboard!F6',
    'WaterFound': 'Dashboard!F7', 'FoodStock': 'Dashboard!B14', 'WaterStock': 'Dashboard!B15',
    'FodderStock': 'Dashboard!B16', 'ProvisionStock': 'Dashboard!B17',

    // Calculations
    'CalcBaseSpeed': 'Calculations!B5', 'CalcTerrainMod': 'Calculations!B6', 'CalcPaceMod': 'Calculations!B7',
    'CalcWeatherMod': 'Calculations!B8', 'CalcMilesToday': 'Calculations!B9',
    'FoodBase': 'Calculations!B13', 'FoodMod': 'Calculations!C13', 'FoodDaily': 'Calculations!D13',
    'FoodDaysLeft': 'Calculations!E13', 'FoodStatus': 'Calculations!F13',
    'WaterBase': 'Calculations!B14', 'WaterMod': 'Calculations!C14', 'WaterDaily': 'Calculations!D14',
    'WaterDaysLeft': 'Calculations!E14', 'WaterStatus': 'Calculations!F14',
    'FodderBase': 'Calculations!B15', 'FodderMod': 'Calculations!C15', 'FodderDaily': 'Calculations!D15',
    'FodderDaysLeft': 'Calculations!E15', 'FodderStatus': 'Calculations!F15',
    'ProvisionBase': 'Calculations!B16', 'ProvisionMod': 'Calculations!C16', 'ProvisionDaily': 'Calculations!D16',
    'ProvisionDaysLeft': 'Calculations!E16', 'ProvisionStatus': 'Calculations!F16',

    // Roster - Characters
    'Char1_Name': 'Roster!A4', 'Char1_Class': 'Roster!B4', 'Char1_CON': 'Roster!C4', 'Char1_Fort': 'Roster!D4',
    'Char1_HP': 'Roster!E4', 'Char1_MaxHP': 'Roster!F4', 'Char1_DaysNoFood': 'Roster!G4',
    'Char1_HoursNoWater': 'Roster!H4', 'Char1_HoursNoSleep': 'Roster!I4', 'Char1_CheckNeeded': 'Roster!J4',
    'Char1_CheckDC': 'Roster!K4', 'Char1_CheckType': 'Roster!L4', 'Char1_Status': 'Roster!M4',

    'Char2_Name': 'Roster!A5', 'Char2_Class': 'Roster!B5', 'Char2_CON': 'Roster!C5', 'Char2_Fort': 'Roster!D5',
    'Char2_HP': 'Roster!E5', 'Char2_MaxHP': 'Roster!F5', 'Char2_DaysNoFood': 'Roster!G5',
    'Char2_HoursNoWater': 'Roster!H5', 'Char2_HoursNoSleep': 'Roster!I5', 'Char2_CheckNeeded': 'Roster!J5',
    'Char2_CheckDC': 'Roster!K5', 'Char2_CheckType': 'Roster!L5', 'Char2_Status': 'Roster!M5',

    'Char3_Name': 'Roster!A6', 'Char3_Class': 'Roster!B6', 'Char3_CON': 'Roster!C6', 'Char3_Fort': 'Roster!D6',
    'Char3_HP': 'Roster!E6', 'Char3_MaxHP': 'Roster!F6', 'Char3_DaysNoFood': 'Roster!G6',
    'Char3_HoursNoWater': 'Roster!H6', 'Char3_HoursNoSleep': 'Roster!I6', 'Char3_CheckNeeded': 'Roster!J6',
    'Char3_CheckDC': 'Roster!K6', 'Char3_CheckType': 'Roster!L6', 'Char3_Status': 'Roster!M6',

    'Char4_Name': 'Roster!A7', 'Char4_Class': 'Roster!B7', 'Char4_CON': 'Roster!C7', 'Char4_Fort': 'Roster!D7',
    'Char4_HP': 'Roster!E7', 'Char4_MaxHP': 'Roster!F7', 'Char4_DaysNoFood': 'Roster!G7',
    'Char4_HoursNoWater': 'Roster!H7', 'Char4_HoursNoSleep': 'Roster!I7', 'Char4_CheckNeeded': 'Roster!J7',
    'Char4_CheckDC': 'Roster!K7', 'Char4_CheckType': 'Roster!L7', 'Char4_Status': 'Roster!M7',

    // Roster - Summary
    'PartySize': 'Roster!P4', 'MountCount': 'Roster!P5', 'ChecksNeeded': 'Roster!P6',
    'CriticalCount': 'Roster!P7', 'AvgHPPercent': 'Roster!P8', 'TotalSize': 'Roster!P10',

    // Roster - Mounts
    'Mount1_Name': 'Roster!A19', 'Mount1_Type': 'Roster!B19', 'Mount1_HD': 'Roster!C19',
    'Mount1_HP': 'Roster!D19', 'Mount1_MaxHP': 'Roster!E19', 'Mount1_Speed': 'Roster!F19',
    'Mount1_Carrying': 'Roster!G19', 'Mount1_MaxLoad': 'Roster!H19', 'Mount1_Grazing': 'Roster!I19',
    'Mount1_Status': 'Roster!J19',

    // Caravan
    'CaravanName': 'Caravan!B4', 'CaravanLevel': 'Caravan!D4', 'CaravanOffense': 'Caravan!B5',
    'CaravanDefense': 'Caravan!D5', 'CaravanMobility': 'Caravan!B6', 'CaravanMorale': 'Caravan!D6',
    'CaravanUnrest': 'Caravan!B7', 'CaravanMaxUnrest': 'Caravan!D7', 'CaravanFortune': 'Caravan!B8',
    'CaravanPrestige': 'Caravan!D8', 'CaravanAttack': 'Caravan!G4', 'CaravanAC': 'Caravan!I4',
    'CaravanSecurity': 'Caravan!G5', 'CaravanResolve': 'Caravan!I5', 'CaravanSpeed': 'Caravan!G6',
    'CaravanHP': 'Caravan!I6', 'CaravanCargo': 'Caravan!G7', 'CaravanTravelers': 'Caravan!I7',
    'CaravanConsumption': 'Caravan!G8',

    // Exploration
    'TerritoryName': 'Exploration!B4', 'TerritoryCR': 'Exploration!B5', 'ExplorationSkill': 'Exploration!B6',
    'ExplorationDC': 'Exploration!B7', 'CurrentDP': 'Exploration!B8', 'DaysExplored': 'Exploration!B9',

    // Status Monitor
    'StatusFood': 'Status Monitor!B5', 'StatusWater': 'Status Monitor!B6', 'StatusFodder': 'Status Monitor!B7',
    'StatusProvision': 'Status Monitor!B8', 'AlertFood': 'Status Monitor!C5', 'AlertWater': 'Status Monitor!C6',
    'AlertFodder': 'Status Monitor!C7', 'AlertProvision': 'Status Monitor!C8', 'PartyAvgHP': 'Status Monitor!G5',
    'PartyChecks': 'Status Monitor!G6', 'PartyCritical': 'Status Monitor!G7', 'AlertText': 'Status Monitor!A11',

    // Reference Tables
    'TerrainTable': 'Reference Tables!A3:D11', 'PaceTable': 'Reference Tables!F3:G5',
    'TempTable': 'Reference Tables!I3:J9', 'WeatherTable': 'Reference Tables!A15:B20',
    'ForageTable': 'Reference Tables!D15:E19', 'AltitudeTable': 'Reference Tables!G15:H18',

    // Ranges
    'AllCharNames': 'Roster!A4:A15', 'AllCharHP': 'Roster!E4:E15', 'AllCharMaxHP': 'Roster!F4:F15',
    'AllCharChecks': 'Roster!J4:J15', 'AllCharStatus': 'Roster!M4:M15', 'AllMountNames': 'Roster!A19:A30',
    'AllWagonHP': 'Caravan!C12:C25', 'AllWagonCargo': 'Caravan!E12:E25', 'AllWagonTravelers': 'Caravan!D12:D25',
    'AllWagonConsumption': 'Caravan!F12:F25', 'AllTravelerJobs': 'Caravan!I12:I25',
    'PartyHeaders': 'Roster!A3:M3', 'PartyTable': 'Roster!A4:M15',
    'MountHeaders': 'Roster!A18:K18', 'MountTable': 'Roster!A19:K30'
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
  const calc = ss.getSheetByName(SHEET_NAMES.CALCULATIONS);
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

  const roster = ss.getSheetByName(SHEET_NAMES.ROSTER);
  for (let i = 1; i <= ROW_OFFSETS.PARTY_COUNT; i++) {
    const row = ROW_OFFSETS.PARTY_START + i - 1;
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

  const caravan = ss.getSheetByName(SHEET_NAMES.CARAVAN);
  caravan.getRange('G4').setFormula('=CaravanOffense + MIN(5, COUNTIF(AllTravelerJobs, "Guard"))');
  caravan.getRange('I4').setFormula('=10 + CaravanDefense');
  caravan.getRange('G5').setFormula('=CaravanOffense + MIN(5, COUNTIF(AllTravelerJobs, "Guide"))');
  caravan.getRange('I5').setFormula('=CaravanMorale + MIN(5, COUNTIF(AllTravelerJobs, "Entertainer"))');
  caravan.getRange('G6').setFormula('=BASE_SPEED + CaravanMobility * 4');
  caravan.getRange('I6').setFormula('=SUM(AllWagonHP)');
  caravan.getRange('G7').setFormula('=SUM(AllWagonCargo)');
  caravan.getRange('I7').setFormula('=SUM(AllWagonTravelers)');
  caravan.getRange('G8').setFormula('=IFERROR(SUM(AllWagonConsumption) + COUNTIF(AllTravelerJobs, "<>") / 2 - MIN(5, COUNTIF(AllTravelerJobs, "Cook")) * 2, 0)');

  const status = ss.getSheetByName(SHEET_NAMES.STATUS);
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

  const exploration = ss.getSheetByName(SHEET_NAMES.EXPLORATION);
  exploration.getRange('B7').setFormula('=IFNA(VLOOKUP(TerritoryCR, D5:E24, 2, FALSE), 16 + TerritoryCR)');
  exploration.getRange('E13').setFormula('=B13+C13+D13');
  exploration.getRange('E14').setFormula('=B14+C14+D14');
  exploration.getRange('E15').setFormula('=IF(B15<>"", B15+C15+D15, "")');
  exploration.getRange('C24').setFormula('=IF(B24="Simple", TerritoryCR+10, IF(B24="Moderate", TerritoryCR+15, IF(B24="Complex", TerritoryCR+20, "")))');
  exploration.getRange('C25').setFormula('=IF(B25="Simple", TerritoryCR+10, IF(B25="Moderate", TerritoryCR+15, IF(B25="Complex", TerritoryCR+20, "")))');
  exploration.getRange('C26').setFormula('=IF(B26="Simple", TerritoryCR+10, IF(B26="Moderate", TerritoryCR+15, IF(B26="Complex", TerritoryCR+20, "")))');
}
