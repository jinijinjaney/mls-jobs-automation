**MLS Jobs Google Apps Script**

This Google Apps Script automates the creation and management of job folders and notes in Google Drive based on rows in a Google Sheet. Each row generates a dedicated folder and a single Job Notes.txt file containing the row’s details. The folder’s URL is automatically written back to column G of the same row.

**Features:**

- **Automatic folder creation:** Creates a folder under MLS Jobs/2025/<Bid#> <Client> when a new row is added.
- **Job Notes file generation:** Generates a Job Notes.txt file containing all row details, including Bid#, Client, TMK, Address, Status, and Last Updated timestamp.
- **URL mapping:** Writes the folder’s URL back to column G of the sheet.
- **Automatic updates:** If a row is updated, the script refreshes the Job Notes.txt file with the new details automatically.
- **Real-time spreadsheet edits:** Manual edits in the spreadsheet trigger immediate updates to the corresponding Job Notes.txt.
- **Custom menu:** Provides an MLS Jobs menu in the sheet to manually process the current row or all rows.

**Setup Instructions:**

1. **Open the Google Sheet**
   - Ensure your sheet has the following headers in row 1:
     Bid# | Client | TMK | Address | Status

2. **Open the Script Editor**
   - In the Google Sheet, go to:
     Extensions → Apps Script

3. **Paste the Script**
   - Copy the provided .gs script and paste it in the editor.

4. **Update Configuration (if needed)**
   Inside the script, locate the configuration section and update the constants if your setup differs:
   - ROOT_FOLDER_PATH = 'MLS Jobs/2025' // Parent folder path in Drive
   - TARGET_SHEET_NAME = 'Sheet1'       // Sheet name
   - FOLDER_URL_COLUMN = 7              // Column to store folder URL (default: G)

5. **Save the Script**
   - Click the Save icon in Apps Script.

6. **Install the onEdit Trigger**
   1. In the Apps Script editor, click the Triggers icon (clock icon on the left side).
   2. Click + Add Trigger.
   3. Configure the trigger:
      - Choose which function to run: onEdit
      - Deployment: Head
      - Event source: From spreadsheet
      - Event type: On edit
   4. Click Save.

7. **Test the Script**
   - Return to your Google Sheet.
   - Add a new row and watch the script create a folder, add a Job Notes.txt file, and write the folder URL in column G.
   - Update an existing row and confirm the Job Notes.txt file updates automatically.

8. **Use the MLS Jobs Menu**
   - A custom MLS Jobs menu will appear in your sheet:
     - Process all rows – updates folders and notes for all rows.
     - Process current row – updates folder and notes for the selected row.

**Important Usage Note:**

- When manually entering information directly into the spreadsheet, **allow the sheet to save first** before moving to the next column or row. This ensures the script correctly captures the data, updates the Job Notes.txt file properly, and avoids duplications or overwritten information.

**Folder Structure Example:**

MLS Jobs/2025/
├── 1001 Alice Johnson/
│ └── Job Notes.txt
├── 1002 Bright Co./
│ └── Job Notes.txt

**Author:**
Liezyl Jane Claros  
GitHub Profile: https://github.com/jinijinjaney  
Email: liezyljaneclaros10@gmail.com
