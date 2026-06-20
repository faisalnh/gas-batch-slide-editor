# Batch Slides Editor

A Google Apps Script web app for batch-editing Google Slides presentations stored in Google Drive folders. It provides a simple web interface to delete, insert, replace, rename, and update text across multiple slide decks at once.

This project is designed for teams that manage many related Google Slides files and need to apply the same update repeatedly without opening each presentation manually.

## Features

- **Delete slides in bulk**
  - Remove a slide or slide range from every Google Slides file in a target Drive folder.

- **Insert slides in bulk**
  - Copy one or more slides from a source presentation into every presentation in a destination folder.
  - Insert at a specific position or append to the end.

- **Replace slides by matching file name**
  - Replace a slide in each target presentation using a slide from a matching source presentation.
  - Source and target files are matched by file name.

- **Replace text across slide decks**
  - Find and replace text in all slides or in a specific slide number.
  - Optional case-sensitive matching.

- **Rename Google Slides files in bulk**
  - Find and replace text in presentation file names inside a Drive folder.
  - Optional case-sensitive matching.

- **Progress summary UI**
  - Shows processed files, edited files, skipped files, errors, current file, and total changes.

## Project Structure

```text
.
├── appsscript.json   # Apps Script manifest
├── code.js           # Server-side Apps Script logic
├── index.html        # Web app UI, styles, and client-side script
├── .clasp.json       # Clasp project configuration
├── .claspignore      # Files ignored by clasp push
└── .gitignore
```

## Requirements

- Google account with access to Google Drive and Google Slides.
- Permission to edit the target presentations and folders.
- Optional for local development: [Clasp](https://github.com/google/clasp), the Google Apps Script CLI.

## Apps Script Services Used

This app uses built-in Google Apps Script services:

- `HtmlService` for serving the web interface.
- `DriveApp` for reading folders, listing Google Slides files, and renaming files.
- `SlidesApp` for editing presentations and slides.
- `CacheService` for temporary job progress/status tracking.
- `Utilities` for short throttling delays between files.

## Deployment

### Option 1: Deploy from Apps Script Editor

1. Open the project in the Google Apps Script editor.
2. Make sure the files `code.js`, `index.html`, and `appsscript.json` are present.
3. Click **Deploy** → **New deployment**.
4. Select **Web app**.
5. Configure deployment settings:
   - **Execute as:** User accessing the web app
   - **Who has access:** Domain users, or adjust based on your needs
6. Click **Deploy**.
7. Authorize the requested Google Drive and Slides permissions.
8. Open the generated web app URL.

### Option 2: Deploy with Clasp

Install and authenticate Clasp:

```bash
npm install -g @google/clasp
clasp login
```

Push local files to Apps Script:

```bash
clasp push
```

Open the Apps Script project:

```bash
clasp open
```

Then deploy the project as a web app from the Apps Script editor.

## Usage

1. Open the deployed web app URL.
2. Choose one of the available task tabs:
   - **Delete**
   - **Insert**
   - **Replace**
   - **Text Replace**
   - **File Name Replace**
3. Paste the required Google Drive folder or Google Slides presentation URLs.
4. Enter slide numbers, text values, or replacement settings.
5. Click the run button for the selected task.
6. Wait for the progress summary to complete.

## Important Notes

- This app directly modifies Google Slides files in the selected folders.
- Make a backup before running large batch operations.
- Slide numbers are 1-based, matching the numbering users see in Google Slides.
- For replacement by folder, source and target presentations must have matching file names.
- The web app is configured in `appsscript.json` to run as the accessing user and to be available to the domain.
- Job progress is stored temporarily in Apps Script cache and may expire after several minutes.

## Current Manifest Configuration

```json
{
  "timeZone": "Asia/Jakarta",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_ACCESSING",
    "access": "DOMAIN"
  }
}
```

## License

No license has been specified yet. Add a license file if this repository will be shared or reused publicly.
