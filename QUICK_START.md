# Quick Start Guide: Connect to Google Sheet

## Step-by-Step Connection Process

### 1. Install and Login
```bash
npm install -g @google/clasp
clasp login
```

### 2. Create Your Google Sheet
- Go to [Google Sheets](https://sheets.google.com)
- Create a new spreadsheet
- Name it "Pathfinder Campaign Manager" (or whatever you like)

### 3. Get the Script ID
- In your Google Sheet, click **Extensions > Apps Script**
- Click the **⚙️ Project Settings** icon (left sidebar)
- Find **Script ID** and copy it

### 4. Link Your Local Project
- Open `.clasp.json` in this project
- Paste your Script ID:
```json
{
  "scriptId": "paste-your-script-id-here",
  "rootDir": "."
}
```

### 5. Push Your Code
```bash
clasp push
```

### 6. Initialize the Campaign
- Go back to your Google Sheet
- You should see a **Campaign** menu in the toolbar
- Click **Campaign > Initialize/Reset Campaign**
- Confirm to create all sheets and setup

### 7. Open the Sidebar
- Click **Campaign > Show Campaign Manager**
- The sidebar will open with your campaign dashboard!

## Troubleshooting

**No "Campaign" menu?**
- Make sure you ran `clasp push` successfully
- Refresh your Google Sheet
- Check that `onOpen()` function exists in Code.gs

**Script not working?**
- Run `clasp open` to check the Apps Script editor
- Look for any error messages
- Make sure you authorized the script when prompted

**Want to make changes?**
- Edit files locally
- Run `clasp push` to update
- Refresh your Google Sheet

