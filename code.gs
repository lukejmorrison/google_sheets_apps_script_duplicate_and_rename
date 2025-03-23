/**
 * Created by Luke.
 *
 * This script contains functions to manage sheets in a Google Spreadsheet.
 *
 * **Main Function:**
 * - `duplicateAndRenameTab()`: Duplicates the active sheet, renames the copy with today's date in "yyyyMMdd" format
 *   (e.g., "20231015" for October 15, 2023), and adds "Archived on [Month Day]," and "Document has been protected"
 *   in the first two rows. If a sheet with the same date exists, it appends a counter (e.g., "_2") to ensure a unique name.
 *   The new sheet is then protected, with all editors (including the owner) removed, making it uneditable without manually
 *   adjusting the protection settings. This is useful for creating dated archives or snapshots of the sheet's data, ensuring
 *   historical versions remain unchanged for record-keeping or reference.
 *
 * **Additional Function:**
 * - `myFunction()`: Logs the values from cells A42 to A52 of the active sheet. This may be used for debugging or other purposes
 *   unrelated to the main archiving functionality.
 *
 * **Requirements:**
 * - The script includes the "@OnlyCurrentDoc" annotation to restrict access to the current spreadsheet only, enhancing security.
 * - Ensure the script has permission to run by authorizing it when prompted in the Google Apps Script editor or when first using it.
 *
 * **Usage:**
 * - Run the `duplicateAndRenameTab` function manually from the script editor, set up a trigger (e.g., time-based) to automate it,
 *   or link it to a button in the sheet (see instructions below).
 *
 * **How to Link to a Button in the Sheet:**
 * 1. **Insert a Button**: Go to **Insert** > **Drawing** in the Google Sheet menu. Create a shape (e.g., a rectangle), add text like
 *    "Archive Sheet," then click **Save and Close** to place it in the sheet.
 * 2. **Assign the Script**: Click the drawing, then click the three dots (...) in the top-right corner. Select **Assign script**,
 *    enter "duplicateAndRenameTab" (without quotes), and click **OK**.
 * 3. **Run It**: Click the button to execute the script. Authorize it if prompted (first time only). The active sheet will be duplicated,
 *    renamed, and protected. Ensure you’re on the desired sheet before clicking.
 *
 * **Protection Details:**
 * - The archived sheet is locked for all users, including the owner. To edit it later, go to **Data > Protected sheets and ranges**,
 *   find the sheet (e.g., "20231015"), and either add yourself as an editor or remove the protection manually. This ensures the archive
 *   remains secure until intentionally unlocked.
 *
 * **Note:**
 * - The script file can be named anything e.g., "code.gs".
 * - Check the execution log (**View** > **Logs** in the script editor) if the script doesn’t work as expected.
 */

/**
 * @OnlyCurrentDoc
 */

/**
 * @OnlyCurrentDoc
 */
function duplicateAndRenameTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();

  Logger.log("Starting duplicateAndRenameTab function.");

  // Get today's date and format it
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");

  // Generate a unique sheet name
  var baseName = formattedDate;
  var newSheetName = baseName;
  var counter = 2;
  while (ss.getSheetByName(newSheetName) !== null) {
    newSheetName = baseName + "_" + counter;
    counter++;
    Logger.log("Sheet name exists. Trying: '" + newSheetName + "'");
  }

  // Copy and rename the sheet
  var newSheet = currentSheet.copyTo(ss);
  newSheet.setName(newSheetName);
  Logger.log("Renamed new sheet to: '" + newSheetName + "'");

  // Insert rows and set values
  newSheet.insertRowsBefore(1, 2);
  var archiveDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM dd,");
  newSheet.getRange("A1").setValue("Archived on " + archiveDate);
  newSheet.getRange("A2").setValue("Document has been protected");
  Logger.log("Set text in A1 and A2.");

  // Format cells
  newSheet.getRange("A1:A2").setFontWeight("bold").setFontSize(12);
  Logger.log("Formatted text in A1:A2 (bold, size 12).");

  // Apply protection and remove all editors (including the owner)
  var protection = newSheet.protect();
  protection.setDescription("Archived sheet - locked for all, including owner");
  protection.removeEditors(protection.getEditors()); // Remove all current editors
  Logger.log("Applied protection to the sheet and removed all editors, including owner.");

  Logger.log("Finished duplicateAndRenameTab function.");
}