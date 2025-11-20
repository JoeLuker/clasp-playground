# Hex Map Import Guide

## Method 1: Using the Menu (Small Maps)

1. Go to **Campaign > Import Hex Map**
2. Paste your JSON data
3. Click OK

**Note:** This method has a character limit (~50,000 characters). If your map is large, use Method 2.

## Method 2: Direct Function Call (Large Maps)

For large hex maps that exceed the prompt dialog limit:

1. Open **Extensions > Apps Script**
2. In the function dropdown, select `importHexMapLarge`
3. In the code editor, modify the function call to include your JSON:

```javascript
function testImport() {
  const jsonData = `{
    "metadata": { ... },
    "terrain": { ... }
  }`;
  const result = importHexMapLarge(jsonData);
  Logger.log(result);
}
```

4. Run the `testImport` function
5. Check the execution log for the result

## Method 3: Using a Named Range (Recommended for Very Large Maps)

1. Create a new sheet called "Hex Map JSON"
2. Paste your JSON into cell A1
3. Run this in the Apps Script editor:

```javascript
function importFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jsonSheet = ss.getSheetByName('Hex Map JSON');
  if (!jsonSheet) {
    Logger.log('Error: Create a sheet named "Hex Map JSON" and paste your JSON in A1');
    return;
  }
  const jsonData = jsonSheet.getRange('A1').getValue();
  const result = importHexMapLarge(jsonData);
  Logger.log(result);
}
```

## Troubleshooting

**"Invalid JSON format" error:**
- Make sure you copied the complete JSON
- Check for any truncation in the prompt dialog
- Use Method 2 or 3 for large maps

**"No hex data found" error:**
- Verify your JSON has a `terrain.hexes` structure
- Make sure you're using the correct JSON format from your map editor

**Import seems to hang:**
- Large maps (1000+ hexes) may take a moment
- Check the execution log for progress
- The function processes hexes in batches of 1000

