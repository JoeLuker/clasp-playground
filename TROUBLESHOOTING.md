# Troubleshooting Guide

## "Service Spreadsheets failed while accessing document" Error

This error typically means the script cannot access your Google Sheet. Here's how to fix it:

### Solution 1: Re-authorize the Script

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. In the Apps Script editor, click **Run** on any function (like `onOpen`)
4. When prompted, click **Review Permissions**
5. Select your Google account
6. Click **Advanced** → **Go to [Project Name] (unsafe)**
7. Click **Allow** to grant permissions

### Solution 2: Check Script Binding

Make sure the script is properly bound to your sheet:

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Check that you see your `Code.gs` file
4. If not, the script might be standalone - you need to:
   - Copy the Script ID from the Apps Script editor
   - Make sure `.clasp.json` has the correct Script ID
   - Run `clasp push` again

### Solution 3: Check Spreadsheet Permissions

1. Make sure you have **Edit** access to the spreadsheet
2. If the spreadsheet is shared, ensure you're logged into the correct Google account
3. Try opening the spreadsheet directly in your browser

### Solution 4: Re-initialize the Campaign

1. In your Google Sheet, go to **Campaign > Initialize/Reset Campaign**
2. This will re-authorize and set up all sheets
3. If this fails, the script may need to be re-bound

### Solution 5: Check Script ID

1. Verify your `.clasp.json` has the correct Script ID:
   ```bash
   cat .clasp.json
   ```
2. The Script ID should match the one in your Apps Script editor
3. If it's wrong, update it and run `clasp push`

### Solution 6: Re-push the Code

If all else fails:

```bash
clasp push --force
```

Then refresh your Google Sheet and try again.

## Common Issues

### "Sheet not initialized" Error
- Run **Campaign > Initialize/Reset Campaign** from the menu

### Menu Not Appearing
- Refresh the Google Sheet
- Make sure you ran `clasp push` successfully
- Check that `onOpen()` function exists in Code.gs

### Sidebar Not Loading
- Check the browser console for errors (F12)
- Make sure all HTML files were pushed (`Sidebar.html`, `ClientScript.html`, `Stylesheet.html`)
- Try closing and reopening the sidebar

### Named Ranges Not Found
- Run **Campaign > Initialize/Reset Campaign** to recreate all named ranges
- Check that all sheets exist (Dashboard, Calculations, Roster, etc.)

## Getting Help

If you continue to have issues:

1. Check the Apps Script execution log:
   - **Extensions > Apps Script > View > Execution log**
2. Check for error messages in the sidebar
3. Verify your `.clasp.json` has the correct Script ID
4. Make sure you're logged into the correct Google account

