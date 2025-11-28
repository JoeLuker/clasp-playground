/**
 * Sidebar API
 * Functions called by the HTML sidebar for data retrieval and actions.
 */

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

    const rangeNames = [
      'CurrentDay', 'TotalMiles', 'CurrentHex', 'HexTerrain',
      'Temperature', 'Terrain', 'PathType', 'Weather', 'TravelPace',
      'AnimalsGrazing', 'FoodDaysLeft', 'FoodStatus', 'WaterDaysLeft',
      'WaterStatus', 'AvgHPPercent', 'ChecksNeeded', 'CriticalCount', 'AlertText'
    ];

    const values = {};
    rangeNames.forEach(name => {
      try {
        const range = ss.getRangeByName(name);
        if (range) {
          values[name] = name === 'AvgHPPercent' ? range.getDisplayValue() : range.getValue();
        } else {
          values[name] = name === 'CurrentHex' || name === 'HexTerrain' ? '' : null;
        }
      } catch (e) {
        Logger.log(`Warning: Could not read range ${name}: ${e.message}`);
        values[name] = name === 'CurrentHex' || name === 'HexTerrain' ? '' : null;
      }
    });

    return {
      currentDay: values.CurrentDay || 0,
      totalMiles: values.TotalMiles || 0,
      currentHex: values.CurrentHex || '',
      hexTerrain: values.HexTerrain || '',
      temperature: values.Temperature || 'Normal',
      terrain: values.Terrain || 'Plains',
      pathType: values.PathType || 'Road or Trail',
      weather: values.Weather || 'Clear',
      travelPace: values.TravelPace || 'Normal',
      animalsGrazing: values.AnimalsGrazing || false,
      foodDays: values.FoodDaysLeft || 0,
      foodStatus: values.FoodStatus || 'GOOD',
      waterDays: values.WaterDaysLeft || 0,
      waterStatus: values.WaterStatus || 'GOOD',
      avgHp: values.AvgHPPercent || '100%',
      checksNeeded: values.ChecksNeeded || 0,
      criticalCount: values.CriticalCount || 0,
      alertText: values.AlertText || ''
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
  try {
    if (typeof foodFound !== 'number' || foodFound < 0 || isNaN(foodFound)) {
      throw new Error('Food found must be a non-negative number');
    }
    if (typeof waterFound !== 'number' || waterFound < 0 || isNaN(waterFound)) {
      throw new Error('Water found must be a non-negative number');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Cannot access spreadsheet');
    }

    if (foodFound > 0) {
      ss.getRangeByName('FoodFound').setValue(foodFound);
      logEvent('Resource Found', `Found ${foodFound} days of food`, '');
    }
    if (waterFound > 0) {
      ss.getRangeByName('WaterFound').setValue(waterFound);
      logEvent('Resource Found', `Found ${waterFound} gallons of water`, '');
    }
    return 'Resources updated successfully';
  } catch (e) {
    Logger.log('Error in updateResourcesFound: ' + e.toString());
    throw e;
  }
}

/**
 * Updates environment settings from the sidebar.
 * @param {object} environment Object containing temperature, terrain, pathType, weather, travelPace, animalsGrazing
 * @returns {string} Confirmation message
 */
function updateEnvironment(environment) {
  try {
    if (!environment || typeof environment !== 'object') {
      throw new Error('Environment object is required');
    }

    if (environment.temperature && !VALID_TEMPERATURES.includes(environment.temperature)) {
      throw new Error(`Invalid temperature: ${environment.temperature}`);
    }
    if (environment.terrain && !VALID_TERRAINS.includes(environment.terrain)) {
      throw new Error(`Invalid terrain: ${environment.terrain}`);
    }
    if (environment.pathType && !VALID_PATH_TYPES.includes(environment.pathType)) {
      throw new Error(`Invalid path type: ${environment.pathType}`);
    }
    if (environment.weather && !VALID_WEATHER.includes(environment.weather)) {
      throw new Error(`Invalid weather: ${environment.weather}`);
    }
    if (environment.travelPace && !VALID_PACES.includes(environment.travelPace)) {
      throw new Error(`Invalid travel pace: ${environment.travelPace}`);
    }
    if (environment.animalsGrazing !== undefined && typeof environment.animalsGrazing !== 'boolean') {
      throw new Error('AnimalsGrazing must be a boolean');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Cannot access spreadsheet');
    }

    const updates = [];
    if (environment.temperature) updates.push({ name: 'Temperature', value: environment.temperature });
    if (environment.terrain) updates.push({ name: 'Terrain', value: environment.terrain });
    if (environment.pathType) updates.push({ name: 'PathType', value: environment.pathType });
    if (environment.weather) updates.push({ name: 'Weather', value: environment.weather });
    if (environment.travelPace) updates.push({ name: 'TravelPace', value: environment.travelPace });
    if (environment.animalsGrazing !== undefined) updates.push({ name: 'AnimalsGrazing', value: environment.animalsGrazing });

    updates.forEach(update => {
      const range = ss.getRangeByName(update.name);
      if (range) {
        range.setValue(update.value);
      }
    });

    logEvent('Environment Changed',
      `${environment.terrain || 'Unknown'} terrain, ${environment.weather || 'Unknown'} weather, ${environment.travelPace || 'Unknown'} pace`,
      `Temperature: ${environment.temperature || 'Unknown'}, Path: ${environment.pathType || 'Unknown'}, Grazing: ${environment.animalsGrazing ? 'Yes' : 'No'}`);

    return 'Environment updated successfully';
  } catch (e) {
    Logger.log('Error in updateEnvironment: ' + e.toString());
    throw e;
  }
}

/**
 * Processes the day's events, consumes resources, and returns a status message.
 * @returns {string} A confirmation message for the user.
 */
function processDay() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Cannot access spreadsheet');
    }

    const currentDay = ss.getRangeByName('CurrentDay').getValue() || 0;
    const miles = ss.getRangeByName('CalcMilesToday').getValue() || 0;
    const totalMiles = ss.getRangeByName('TotalMiles').getValue() || 0;
    const foodUsed = ss.getRangeByName('FoodDaily').getValue() || 0;
    const waterUsed = ss.getRangeByName('WaterDaily').getValue() || 0;
    const fodderUsed = ss.getRangeByName('FodderDaily').getValue() || 0;
    const provisionUsed = ss.getRangeByName('ProvisionDaily').getValue() || 0;
    const foodFound = ss.getRangeByName('FoodFound').getValue() || 0;
    const waterFound = ss.getRangeByName('WaterFound').getValue() || 0;
    const foodStock = ss.getRangeByName('FoodStock').getValue() || 0;
    const waterStock = ss.getRangeByName('WaterStock').getValue() || 0;
    const fodderStock = ss.getRangeByName('FodderStock').getValue() || 0;
    const provisionStock = ss.getRangeByName('ProvisionStock').getValue() || 0;

    const newDay = currentDay + 1;

    ss.getRangeByName('CurrentDay').setValue(newDay);
    ss.getRangeByName('TotalMiles').setValue(totalMiles + miles);
    ss.getRangeByName('FoodStock').setValue(foodStock - foodUsed + foodFound);
    ss.getRangeByName('WaterStock').setValue(waterStock - waterUsed + waterFound);
    ss.getRangeByName('FodderStock').setValue(fodderStock - fodderUsed);
    ss.getRangeByName('ProvisionStock').setValue(provisionStock - provisionUsed);
    ss.getRangeByName('FoodFound').setValue(0);
    ss.getRangeByName('WaterFound').setValue(0);

    updateDeprivation(ss);

    const log = ss.getSheetByName(SHEET_NAMES.LOG);
    log.appendRow([
      newDay, new Date(), miles,
      foodUsed,
      waterUsed,
      fodderUsed,
      provisionUsed,
      ''
    ]);

    const details = `Traveled ${Math.round(miles)} miles. Resources: Food -${foodUsed}${foodFound > 0 ? ' +' + foodFound : ''}, Water -${waterUsed}${waterFound > 0 ? ' +' + waterFound : ''}, Fodder -${fodderUsed}, Provisions -${provisionUsed}`;
    logEvent('Day Processed', `Advanced to Day ${newDay}`, details);

    return `Advanced to Day ${newDay}. Traveled ${Math.round(miles)} miles.`;
  } catch (e) {
    Logger.log('Error in processDay: ' + e.toString());
    throw e;
  }
}

/**
 * Gets the daily preview data for display.
 * @returns {object} An object with the calculated preview data.
 */
function previewDay() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Cannot access spreadsheet');
    }

    return {
      miles: Math.round(ss.getRangeByName('CalcMilesToday').getValue() || 0),
      food: ss.getRangeByName('FoodDaily').getValue() || 0,
      water: ss.getRangeByName('WaterDaily').getValue() || 0,
      fodder: ss.getRangeByName('FodderDaily').getValue() || 0,
      provisions: ss.getRangeByName('ProvisionDaily').getValue() || 0,
    };
  } catch (e) {
    Logger.log('Error in previewDay: ' + e.toString());
    throw e;
  }
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
