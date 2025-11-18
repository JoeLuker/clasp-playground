# Pathfinder Campaign Manager - Google Apps Script Project

A complete Google Apps Script project for managing Pathfinder campaigns in Google Sheets, developed locally using clasp.

## Setup Instructions

### 1. Install clasp

```bash
npm install -g @google/clasp
```

### 2. Authenticate with Google

```bash
clasp login
```

This will open a browser window for you to authorize clasp to access your Google account.

### 3. Create or Link a Project

#### Option A: Create a new Apps Script project

```bash
clasp create --type sheets --title "Pathfinder Campaign Manager"
```

This will create a new Google Apps Script project and update `.clasp.json` with the script ID.

#### Option B: Link to an existing project

If you already have a Google Apps Script project:

1. Open the project in the [Apps Script editor](https://script.google.com)
2. Copy the Script ID from the URL or Project Settings
3. Update `.clasp.json` with your script ID:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "."
}
```

### 4. Push Code to Google

```bash
clasp push
```

This uploads all your local files to the Apps Script project.

### 5. Set Up the Spreadsheet

1. Create a new Google Sheet or use an existing one
2. In the Apps Script editor (after pushing), go to **Project Settings**
3. Under **Google Cloud Platform (GCP) Project**, note your project
4. In your Google Sheet, go to **Extensions > Apps Script**
5. Make sure the script is attached to your sheet
6. Run `initializeCompleteCampaign()` from the editor to set up all sheets and named ranges

## Development Workflow

### Making Changes

1. Edit files locally in VS Code (or your preferred editor)
2. Push changes to Google:
   ```bash
   clasp push
   ```
3. Test in the Google Sheet

### Pulling Changes

If you make changes in the web editor and want to sync them locally:

```bash
clasp pull
```

### Opening the Web Editor

```bash
clasp open
```

This opens the Apps Script editor in your browser.

## Project Structure

```
.
├── Code.gs              # Main Apps Script backend code
├── Sidebar.html         # Main HTML sidebar template
├── Stylesheet.html      # CSS styles
├── ClientScript.html    # Client-side JavaScript
├── .clasp.json          # Clasp configuration
├── .claspignore         # Files to ignore when pushing
└── README.md            # This file
```

## Features

- **Campaign Dashboard**: Track day, miles, environment, and resources
- **Party Management**: Manage characters, mounts, and their status
- **Resource Tracking**: Food, water, fodder, and provisions
- **Caravan System**: Manage wagons, travelers, and caravan stats
- **Exploration Tracker**: Track territory exploration
- **Status Monitoring**: Real-time alerts for critical conditions
- **Daily Processing**: Automatically advance days and consume resources

## Usage

1. Open your Google Sheet
2. Go to **Campaign > Initialize/Reset Campaign** to set up the system
3. Use **Campaign > Show Campaign Manager** to open the sidebar
4. Configure your party, environment, and resources in the sheets
5. Use the sidebar to process days and monitor status

## Notes

- The script creates multiple sheets with color-coded tabs
- Named ranges are used extensively for formula references
- The sidebar provides a clean interface for daily operations
- All calculations are performed in the spreadsheet using formulas

