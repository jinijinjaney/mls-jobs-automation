/**
 * MLS Jobs - Google Apps Script
 *
 * Creates/updates a folder for each row, adds a Job Notes.txt file,
 * and writes the folder URL back to column G.
 */

/* CONFIG */
const ROOT_FOLDER_PATH = 'MLS Jobs/2025'; 
const TARGET_SHEET_NAME = 'Sheet1';  
const FOLDER_URL_COLUMN = 7;   // G 
const HEADER_ROW = 1;
const COLUMNS = {
  BID: 1,      // A
  CLIENT: 2,   // B
  TMK: 3,      // C
  ADDRESS: 4,  // D
  STATUS: 5    // E
};


/* Helpers */
function cleanName(name) {
  if (!name) return 'unnamed';
  return name.toString().replace(/[\/\\\?\%\*\:\|\"<>\.]/g, '-').trim();
}

/* Menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MLS Jobs')
    .addItem('Process all rows', 'processAllRows')
    .addItem('Process current row', 'processCurrentRowMenu')
    .addToUi();
}

function processCurrentRowMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const row = sheet.getActiveRange().getRow();
  if (row <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert('Select a data row.');
    return;
  }
  createOrUpdateFolderForRow(sheet, row);
  SpreadsheetApp.getUi().alert('Processed row ' + row);
}

/* onEdit trigger */
function onEdit(e) {
  try {
    if (!e) return;
    const sheet = e.source.getActiveSheet();
    if (!sheet || sheet.getName() !== TARGET_SHEET_NAME) return;

    const row = e.range.getRow();
    if (row <= HEADER_ROW) return;

    const editedStartCol = e.range.getColumn();
    const editedNumCols = e.range.getNumColumns ? e.range.getNumColumns() : 1;
    const editedCols = [];
    for (let c = editedStartCol; c < editedStartCol + editedNumCols; c++) editedCols.push(c);

    const allowedCols = [COLUMNS.BID, COLUMNS.CLIENT, COLUMNS.TMK, COLUMNS.ADDRESS, COLUMNS.STATUS];
    const touched = editedCols.some(c => allowedCols.indexOf(c) !== -1);
    if (!touched) return; 

  const bid = sheet.getRange(row, COLUMNS.BID).getValue().toString().trim();
  const client = sheet.getRange(row, COLUMNS.CLIENT).getValue().toString().trim();

  if (!bid || !client) {
    return;
  }

    createOrUpdateFolderForRow(sheet, row);
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/* Main logic */
function createOrUpdateFolderForRow(sheet, row) {
  try {
    function safeGet(r, c) {
      const cell = sheet.getRange(r, c);
      const d = cell.getDisplayValue();
      if (d !== undefined && d !== null && d !== '') return d;
      return cell.getValue();
    }

    const bid = safeGet(row, COLUMNS.BID);
    const client = safeGet(row, COLUMNS.CLIENT);
    const newFolderName = cleanName(bid + ' ' + client);

    const parent = getOrCreateFolderByPath(ROOT_FOLDER_PATH);

    const urlCell = sheet.getRange(row, FOLDER_URL_COLUMN);
    const existingUrl = urlCell.getValue();
    let folder;

    if (existingUrl) {
      try {
        const folderId = existingUrl.match(/[-\w]{25,}/);
        if (folderId) {
          folder = DriveApp.getFolderById(folderId[0]);

          if (folder.getName() !== newFolderName) {
            folder.setName(newFolderName);
          }
        }
      } catch (err) {
        Logger.log("Couldn't open existing folder, will fallback: " + err);
      }
    }

    if (!folder) {
      folder = findOrCreateFolderInParent(parent, newFolderName);
    }

    const notes = buildNotes(sheet, row);

    const notesFileBase = 'Job Notes';
    const notesFileName = 'Job Notes.txt';

    const filesIter = folder.getFiles();
    while (filesIter.hasNext()) {
      const f = filesIter.next();
      const fname = f.getName();
      if (fname && fname.toLowerCase().startsWith(notesFileBase.toLowerCase())) {
        f.setTrashed(true);
      }
    }

    folder.createFile(notesFileName, notes);
    urlCell.setValue(folder.getUrl());

  } catch (err) {
    Logger.log('createOrUpdateFolderForRow error: ' + err);
  }
}

/* Drive helpers */
function getOrCreateFolderByPath(path) {
  if (!path) return DriveApp.getRootFolder();
  const parts = path.split('/').filter(p => p && p.trim() !== '');
  let parent = DriveApp.getRootFolder();
  for (let i = 0; i < parts.length; i++) {
    const name = parts[i].trim();
    const folders = parent.getFoldersByName(name);
    if (folders.hasNext()) {
      parent = folders.next();
    } else {
      parent = parent.createFolder(name);
    }
  }
  return parent;
}

function findOrCreateFolderInParent(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

/* Build notes content */
function buildNotes(sheet, row) {
  function safeGet(r, c) {
    const cell = sheet.getRange(r, c);
    const d = cell.getDisplayValue();
    if (d !== undefined && d !== null && d !== '') return d;
    return cell.getValue();
  }

  const bid = safeGet(row, COLUMNS.BID);
  const client = safeGet(row, COLUMNS.CLIENT);
  const tmk = safeGet(row, COLUMNS.TMK);
  const address = safeGet(row, COLUMNS.ADDRESS);
  const status = safeGet(row, COLUMNS.STATUS);

  const updated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  return [
    'Job Notes',
    '---------',
    'Bid#: ' + (bid || ''),
    'Client: ' + (client || ''),
    'TMK: ' + (tmk || ''),
    'Address: ' + (address || ''),
    'Status: ' + (status || ''),
    'Row: ' + row,
    'Last Updated: ' + updated
  ].join('\n');
}

/* Bulk process */
function processAllRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    const bid = sheet.getRange(r, COLUMNS.BID).getValue();
    const client = sheet.getRange(r, COLUMNS.CLIENT).getValue();
    if (bid && client) {
      createOrUpdateFolderForRow(sheet, r);
    }
  }
}
