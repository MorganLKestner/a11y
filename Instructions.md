# **Google Sheets to Google Calendar Event Management Script**

This document provides instructions for setting up and using a Google Apps Script that automates the creation and updating of calendar events directly from a Google Sheet. 

## **1\. Purpose of the Script**

This script is designed to streamline the process of managing calendar events related to your team's schedule. It allows users to:

* **Create multiple all-day calendar events** based on selected rows in a Google Sheet, pulling information such as Start Date, End Date, and Assigned Responsibility.  
* **Automatically update existing calendar events** when the "Assigned Responsibility" changes in the Google Sheet, ensuring the calendar reflects the most current assignments.

## **2\. Prerequisites**

* A Google Sheet with your scheduling data. You will need to specify its name in the script.  
* A Google Calendar where events will be created and updated. You will need its Calendar ID.  
  * **Important:** You will need to update the Calendar ID in the script with your team's specific calendar ID.  
* The Google Sheet should have the following data in these columns:  
  * **Column A (1):** Start Date  
  * **Column B (2):** End Date (Optional \- if left blank, a single-day event will be created)  
  * **Column C (3):** Assigned Responsibility (e.g., SME, Team Member, Project Lead)

## **3\. Script Code**

Add the script code file from the resource folder `app-script-google.md`

## **4\. Setting Up the Script**

Follow these steps to add the script to your Google Sheet:

1. **Open your Google Sheet:** Go to your relevant Google Sheet.  
2. **Open Apps Script:** From the top menu, click Extensions \> Apps Script. This will open a new tab with the Apps Script editor.  
3. **Paste the Code:** Delete any existing code in the Code.gs file and paste the entire script provided in Section 3 into the editor.  
4. **Configure the Script:**  
   * In both the sendToCalendar and updateCalendarEventOnEdit functions, locate the \--- CONFIGURATION START \--- and \--- CONFIGURATION END \--- sections.  
   * **Replace "\[YOUR\_SHEET\_NAME\]"** with the exact name of your Google Sheet (e.g., "My Project Schedule").  
   * **Replace "\[YOUR\_CALENDAR\_ID\]"** with the ID of the Google Calendar you want to use. You can find your calendar ID in Google Calendar settings.  
   * Optionally, adjust EVENT\_TITLE\_PREFIX if you want a different prefix for your event titles.  
5. **Save the Script:** Click the floppy disk icon (Save project) or press Ctrl \+ S (Windows/Linux) / Cmd \+ S (Mac).  
6. **Authorize the Script (First Time Only):**  
   * When you save or try to run the script for the first time, you might be prompted to review permissions.  
   * Click Review permissions.  
   * Select your Google account.  
   * Click Allow to grant the script access to your Google Sheet and Google Calendar.

## **5\. Setting Up the onEdit Trigger for Automatic Updates**

To enable automatic event updates when the Assigned Responsibility changes, you need to set up an installable onEdit trigger:

1. **Go to the Apps Script editor** (if not already open).  
2. On the left sidebar, click the **Triggers** icon (looks like an alarm clock).  
3. Click the **\+ Add Trigger** button in the bottom right corner.  
4. Configure the trigger as follows:  
   * **Choose which function to run:** Select updateCalendarEventOnEdit.  
   * **Choose deployment to run:** Head (this is usually the default).  
   * **Select event source:** From spreadsheet.  
   * **Select event type:** On edit.  
   * **Failure notification settings:** Choose how often you want to be notified of failures (e.g., Notify me immediately or Notify me daily).  
5. Click **Save**.  
6. If prompted again, authorize the script.

## **6\. How to Use the Script**

### **6.1. Manually Creating Events for Selected Rows**

1. **Open your Google Sheet.**  
2. **Select the rows** for which you want to create calendar events. You can select a single row, multiple contiguous rows, or even a range of cells that spans multiple rows.  
3. From the Google Sheet menu bar, click on the Event **Menu \> Send Selected Row(s) to the calendar**.  
4. A confirmation alert will appear once the events have been created.

### **6.2. Automatically Updating Events on Assigned Responsibility Change**

1. **Ensure the onEdit trigger is set up** as described in Section 5\.  
2. In your Google Sheet, navigate to the sheet you configured.  
3. **Edit a cell in Column C (Assigned Responsibility)** for an existing entry that already has a corresponding calendar event.  
4. After you finish editing the cell (e.g., by pressing Enter or clicking out of the cell), the script will automatically run in the background.  
5. If an event matching the old responsibility and date range is found, its title will be updated in the calendar. A confirmation alert will appear in the spreadsheet.

## **7\. Important Notes and Considerations**

* **Configuration:** Remember to update the SHEET\_NAME, CALENDAR\_ID, and EVENT\_TITLE\_PREFIX variables in the script's \--- CONFIGURATION \--- sections.  
* **Column Order:** The script relies on the specific column order (Start Date in A, End Date in B, Assigned Responsibility in C). If your sheet's layout changes, you will need to adjust the START\_DATE\_COL\_INDEX, END\_DATE\_COL\_INDEX, ASSIGNED\_RESPONSIBILITY\_COL\_INDEX constants in sendToCalendar and START\_DATE\_COLUMN, END\_DATE\_COLUMN, ASSIGNED\_RESPONSIBILITY\_COLUMN in updateCalendarEventOnEdit accordingly.  
* **Event Titles:** The script uses the configured EVENT\_TITLE\_PREFIX \+ Assigned Responsibility for event titles. This consistent naming convention is crucial for the onEdit trigger to find and update existing events.  
* **Duplicate Events:** If you run the manual Send Selected Row(s) to calendar option multiple times on the same rows, it will create duplicate events in your calendar. The onEdit trigger only updates existing events; it does not prevent new duplicates from being created via the manual menu option.  
* **Error Messages:** The script includes SpreadsheetApp.getUi().alert() messages for success and common errors. More detailed error logs can be found in the Apps Script editor under Executions (the clock icon on the left sidebar).  
* **Permissions:** Google Apps Script requires permissions to access your Google Sheet and Google Calendar. Ensure these permissions are granted during the initial setup.