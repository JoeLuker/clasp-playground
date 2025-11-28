/**
 * Hex Map Management
 * Import, navigation, and terrain mapping for hex-based exploration.
 */

/**
 * Alternative import method for large JSON files.
 * Call this directly from the Apps Script editor with your JSON as a string.
 * @param {string} jsonData The JSON string (can be very large)
 * @returns {string} Result message
 */
function importHexMapLarge(jsonData) {
  try {
    return importHexMap(jsonData);
  } catch (e) {
    Logger.log('Import error: ' + e.toString());
    return `Error: ${e.message}`;
  }
}

/**
 * Prompts user to import hex map JSON data.
 */
function importHexMapPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Import Hex Map',
    'Paste the JSON data from your hex map editor.\n\nNote: If your JSON is very large, the prompt may truncate it. In that case, you can:\n1. Save the JSON to a text file\n2. Use the Apps Script editor to call importHexMap() directly with the JSON string\n\nPaste JSON here:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const jsonData = response.getResponseText().trim();
    if (!jsonData) {
      ui.alert('Error', 'No JSON data provided.', ui.ButtonSet.OK);
      return;
    }

    try {
      const result = importHexMap(jsonData);
      ui.alert('Import Result', result, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Import Error', `Failed to import hex map: ${e.message}\n\nMake sure the JSON is valid and complete.`, ui.ButtonSet.OK);
      Logger.log('Hex map import error: ' + e.toString());
    }
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
 * Imports hex map data from JSON and populates the Hex Map sheet.
 * @param {string} jsonData The JSON string containing hex map data
 * @returns {string} Confirmation message
 */
function importHexMap(jsonData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Cannot access spreadsheet. Please ensure the script is bound to a Google Sheet.');
    }

    const hexMapSheet = ss.getSheetByName(SHEET_NAMES.HEX_MAP);
    if (!hexMapSheet) {
      throw new Error('Hex Map sheet not found. Please run "Initialize/Reset Campaign" first.');
    }

    if (!jsonData || jsonData.trim() === '') {
      throw new Error('Empty JSON data provided.');
    }

    let mapData;
    try {
      mapData = JSON.parse(jsonData);
    } catch (parseError) {
      throw new Error(`Invalid JSON format: ${parseError.message}`);
    }

    if (!mapData.terrain || !mapData.terrain.hexes) {
      throw new Error('JSON does not contain terrain.hexes data. Make sure you copied the complete map JSON.');
    }

    const hexes = mapData.terrain.hexes;
    const hexCount = Object.keys(hexes).length;

    if (hexCount === 0) {
      throw new Error('No hex data found in the JSON.');
    }

    const lastRow = hexMapSheet.getLastRow();
    if (lastRow > 10) {
      hexMapSheet.deleteRows(11, lastRow - 10);
    }

    const hexRows = [];
    Object.keys(hexes).forEach(hexKey => {
      const hex = hexes[hexKey];
      if (!hex || typeof hex.q === 'undefined' || typeof hex.r === 'undefined' || typeof hex.s === 'undefined') {
        Logger.log(`Skipping invalid hex: ${hexKey}`);
        return;
      }
      const mapTerrain = hex.tile?.tile_name || '';
      const pfTerrain = mapTerrainToPathfinder(mapTerrain);
      hexRows.push([
        `${hex.q}:${hex.r}:${hex.s}`,
        mapTerrain,
        pfTerrain
      ]);
    });

    if (hexRows.length === 0) {
      throw new Error('No valid hex data could be extracted from the JSON.');
    }

    hexRows.sort((a, b) => {
      try {
        const [q1, r1, s1] = a[0].split(':').map(Number);
        const [q2, r2, s2] = b[0].split(':').map(Number);
        if (q1 !== q2) return q1 - q2;
        if (r1 !== r2) return r1 - r2;
        return s1 - s2;
      } catch (e) {
        return 0;
      }
    });

    if (hexRows.length > 0) {
      for (let i = 0; i < hexRows.length; i += BATCH_SIZE) {
        const batch = hexRows.slice(i, i + BATCH_SIZE);
        const startRow = 5 + i;
        hexMapSheet.getRange(startRow, 4, batch.length, 3).setValues(batch);
      }
    }

    logEvent('Hex Map Imported', `Imported ${hexRows.length} hexes from ${mapData.metadata?.title || 'map'}`, `Total hexes in JSON: ${hexCount}`);

    return `Successfully imported ${hexRows.length} hexes from "${mapData.metadata?.title || 'the map'}".`;
  } catch (e) {
    Logger.log('Hex map import error: ' + e.toString());
    throw e;
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
    'water': 'Plains',
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
    const hexMapSheet = ss.getSheetByName(SHEET_NAMES.HEX_MAP);
    if (!hexMapSheet) {
      return 'Error: Hex Map sheet not found.';
    }

    ss.getRangeByName('CurrentHex').setValue(hexCoords);

    const hexDataRange = hexMapSheet.getRange('D5:F' + hexMapSheet.getLastRow());
    const hexData = hexDataRange.getValues();
    const hexRow = hexData.find(row => row[0] === hexCoords);

    if (hexRow) {
      const mapTerrain = hexRow[1];
      const pfTerrain = hexRow[2] || mapTerrainToPathfinder(mapTerrain);

      ss.getRangeByName('HexTerrain').setValue(mapTerrain);

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
