# Pathfinder Campaign Manager - Google Apps Script Project

A complete Google Apps Script project for managing Pathfinder campaigns in Google Sheets, developed locally using clasp.

## Getting Started (First Time Setup)

### Clone or Download

If you're cloning this repository:

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd clasp-playground
   ```

2. **Set up your clasp configuration:**
   - Copy the example file: `cp .clasp.json.example .clasp.json`
   - Edit `.clasp.json` and add your Script ID (see instructions below)

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

### 5. Connect to a Google Sheet

**Recommended Method: Create a Sheet-Bound Script**

This is the easiest way - the script will be directly attached to your Google Sheet:

1. **Create a new Google Sheet** (or use an existing one)
   - Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
   - Give it a name like "Pathfinder Campaign Manager"

2. **Open Apps Script from the Sheet**
   - In your Google Sheet, click **Extensions > Apps Script**
   - This opens the Apps Script editor with a script bound to your sheet
   - You'll see a default `Code.gs` file with a simple function

3. **Get the Script ID**
   - In the Apps Script editor, click the **Project Settings** icon (⚙️ gear) in the left sidebar
   - Scroll down to find **Script ID**
   - Copy the Script ID (it looks like: `1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7q8r9s0t`)

4. **Set up your local `.clasp.json`**
   - If you don't have `.clasp.json`, copy the example: `cp .clasp.json.example .clasp.json`
   - Open `.clasp.json` in your project
   - Paste the Script ID:
   ```json
   {
     "scriptId": "YOUR_SCRIPT_ID_HERE",
     "rootDir": "."
   }
   ```
   - **Note:** `.clasp.json` is gitignored to keep your Script ID private

5. **Push your code to the sheet**
   ```bash
   clasp push
   ```
   - This will upload all your files (Code.gs, Sidebar.html, etc.) to the Apps Script project
   - The script is now connected to your Google Sheet!

### 6. Initialize the Campaign System

After connecting to your sheet:

1. Open your Google Sheet
2. Go to **Extensions > Apps Script** (or run `clasp open`)
3. In the Apps Script editor, select the function `initializeCompleteCampaign` from the dropdown
4. Click the **Run** button (▶️)
5. Authorize the script when prompted (first time only)
6. Go back to your Google Sheet - you should see a **Campaign** menu in the toolbar
7. Click **Campaign > Initialize/Reset Campaign**
8. Confirm the reset - this will create all the necessary sheets and named ranges
9. Once initialized, use **Campaign > Show Campaign Manager** to open the sidebar

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

## GitHub Setup

This project is configured for GitHub with sensitive files excluded:

### Protected Files (in .gitignore)
- `.clasp.json` - Contains your Script ID (keep private)
- `.clasprc.json` - Contains authentication tokens (never commit)
- `node_modules/` - Dependencies
- `.env` files - Environment variables

### Setting Up GitHub

1. **Create a new repository on GitHub:**
   - Go to [GitHub](https://github.com/new)
   - Create a new repository (don't initialize with README if you already have one)

2. **Connect your local repository:**
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git branch -M main
   git push -u origin main
   ```

3. **For collaborators:**
   - They should clone the repo
   - Copy `.clasp.json.example` to `.clasp.json`
   - Add their own Script ID
   - Run `clasp login` to authenticate

## Notes

- The script creates multiple sheets with color-coded tabs
- Named ranges are used extensively for formula references
- The sidebar provides a clean interface for daily operations
- All calculations are performed in the spreadsheet using formulas
- Sensitive configuration files are gitignored for security

