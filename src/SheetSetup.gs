/**
 * Sheet Setup Functions
 * Individual sheet configuration and initial data population.
 */

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

  sheet.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(VALID_TEMPERATURES).build());
  sheet.getRange('B9').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(VALID_TERRAINS).build());
  sheet.getRange('B10').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(VALID_PATH_TYPES).build());
  sheet.getRange('B11').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(VALID_WEATHER).build());
  sheet.getRange('F4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(VALID_PACES).build());
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
    ['', ''],
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
  sheet.getRange('A2:B30').setValues([
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

  sheet.getRange('A2:D2').setValues([['TERRAIN', 'Highway', 'Road or Trail', 'Trackless']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A3:D11').setValues([
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

  sheet.getRange('F2:G2').setValues([['PACE', 'MODIFIER']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('F3:G5').setValues([['Slow', 0.75], ['Normal', 1.0], ['Fast', 1.25]]);

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

  sheet.getRange('A14:B14').setValues([['WEATHER', 'MODIFIER']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A15:B20').setValues([
    ['Clear', 1.0],
    ['Light Rain', 0.9],
    ['Heavy Rain', 0.75],
    ['Storm', 0.5],
    ['Snow', 0.5],
    ['Blizzard', 0.25]
  ]);

  sheet.getRange('D14:E14').setValues([['FORAGING', 'DC']]).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('D15:E19').setValues([
    ['Abundant', 10],
    ['Average', 15],
    ['Sparse', 20],
    ['Barren', 25],
    ['Desolate', 30]
  ]);

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

  sheet.getRange('A3:B3').merge().setValue('TERRITORY INFORMATION').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('A4:B9').setValues([
    ['Territory Name:', 'Darkwood'],
    ['Territory CR:', 5],
    ['Exploration Skill:', 'Survival'],
    ['Exploration DC:', ''],
    ['Current Discovery Points:', 0],
    ['Days Explored:', 0]
  ]);

  sheet.getRange('D3:F3').merge().setValue('EXPLORATION DC REFERENCE').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('D4:F4').setValues([['CR', 'DC', '']]).setFontWeight('bold');
  sheet.getRange('D5:F24').setValues([
    [1, 17, ''], [2, 19, ''], [3, 21, ''], [4, 22, ''], [5, 23, ''],
    [6, 24, ''], [7, 25, ''], [8, 26, ''], [9, 27, ''], [10, 28, ''],
    [11, 29, ''], [12, 30, ''], [13, 31, ''], [14, 32, ''], [15, 33, ''],
    [16, 34, ''], [17, 35, ''], [18, 36, ''], [19, 37, ''], [20, 38, '']
  ]);

  sheet.getRange('A11').setValue('LOCATIONS').setFontWeight('bold').setBackground('#fff3e0');
  sheet.getRange('A12:F12').setValues([['Name', 'Base Score', 'Terrain Mod', 'Hidden Mod', 'Final Score', 'Status']]).setFontWeight('bold');
  sheet.getRange('A13:F15').setValues([
    ['Hidden Temple', 6, 2, 4, '', 'Undiscovered'],
    ['Ancient Ruins', 3, 2, 0, '', 'Undiscovered'],
    ['', '', '', '', '', '']
  ]);

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

  sheet.getRange('A22').setValue('WAY SIGNS').setFontWeight('bold').setBackground('#e8f5e9');
  sheet.getRange('A23:F23').setValues([['Description', 'Complexity', 'DC', 'Discovery Points', 'Status', 'Notes']]).setFontWeight('bold');
  sheet.getRange('A24:F26').setValues([
    ['Old Map Found', 'Simple', '', 1, 'Undiscovered', 'CR + 10'],
    ['Traveler\'s Journal', 'Moderate', '', 3, 'Undiscovered', 'CR + 15'],
    ['Aerial Reconnaissance', 'Complex', '', 5, 'Undiscovered', 'CR + 20']
  ]);

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

  sheet.getRange('A3:B3').merge().setValue('CURRENT LOCATION').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('A4:B6').setValues([
    ['Current Hex (q:r:s):', '0:0:0'],
    ['Hex Terrain:', ''],
    ['Pathfinder Terrain:', '']
  ]);

  sheet.getRange('D3:F3').merge().setValue('HEX MAP DATA').setFontWeight('bold').setBackground('#e1f5fe');
  sheet.getRange('D4:F4').setValues([['Hex (q:r:s)', 'Map Terrain', 'Pathfinder Terrain']]).setFontWeight('bold');

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
    ['water', 'Plains'],
    ['deep-water', 'Plains']
  ]);
}

function setupEventLog(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1:E1').setValues([['Timestamp', 'Day', 'Event Type', 'Description', 'Details']]).setFontWeight('bold').setBackground('#9c27b0');
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 60);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 250);
    sheet.setColumnWidth(5, 300);
  }
}
