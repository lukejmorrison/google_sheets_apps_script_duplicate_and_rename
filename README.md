# Google Sheets Archiving Script

A Google Apps Script solution for creating protected, dated archives of Google Sheets data with a single click.

## Overview

This script provides an easy way to create date-stamped snapshots of your Google Sheets. It duplicates the active sheet, names it with the current date, adds archival information, and protects it from future edits. This is ideal for preserving historical data, creating audit trails, or maintaining point-in-time records.

## Features

- **One-Click Archiving**: Instantly create archived copies of your current sheet
- **Date-Based Naming**: Automatically names archives using YYYYMMDD format
- **Conflict Handling**: Adds numerical suffixes when same-day archives exist
- **Archival Headers**: Adds timestamp and protection notice to the archived sheet
- **Total Protection**: Archives are protected from edits by all users (including the owner)
- **Simple Integration**: Can be added as a button directly in your spreadsheet

## How It Works

When executed, the script:

1. Duplicates the currently active sheet
2. Names the copy using today's date (format: YYYYMMDD)
3. Adds "Archived on [Month Day]" and "Document has been protected" headers
4. Applies protection settings that prevent all users from editing
5. If a sheet with today's date already exists, appends a counter (e.g., "20231015_2")

## Installation

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Copy the contents of `duplicateAndRename.gs` into the script editor
4. Save the project
5. Authorize the script when prompted

## Important Security Note

The script includes the `@OnlyCurrentDoc` annotation which is **required** for the script to function properly. This annotation:
- Restricts the script's access to only the current spreadsheet
- Reduces the permissions requested during authorization
- Prevents Google from blocking the script due to excessive permission requests

Without this annotation, Google's security measures may block the script or present users with alarming permission requests that could discourage usage.

## Usage

### Method 1: Run from Script Editor
1. Open the Apps Script editor
2. Select the `duplicateAndRenameTab` function
3. Click the Run button (▶)

### Method 2: Add a Button to Your Sheet
1. Go to **Insert > Drawing** in your Google Sheet
2. Create a button shape with appropriate text (e.g., "Archive Sheet")
3. Click **Save and Close**
4. Select the inserted drawing
5. Click the three dots (⋮) in the corner and select **Assign script**
6. Enter `duplicateAndRenameTab` and click OK
7. Click the button whenever you want to create an archive

## Accessing Protected Archives

The archived sheets are protected from all users by default. To make changes:
1. Go to **Data > Protected sheets and ranges**
2. Select the protected sheet
3. Click the three dots (⋮) and select **Remove protection**, or add yourself as an editor

## Notes

- Make sure you're on the correct sheet before running the archive function
- The script includes detailed logging to help with troubleshooting
- The archiving process includes formatting of the header text (bold, size 12)

## License

Feel free to modify and use this script according to your needs.

## Author

Created by Luke Morrison
