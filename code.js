const CONFIG = {
  STATE_SHEET: "USER_PROGRESS_DATA",
  RESULT_SHEET: "RESULTS"
};

/**
 * Text Merge Sorter logic.
 */
function showSorter() {
  const html = HtmlService.createTemplateFromFile('UI')
    .evaluate()
    .setTitle('Text Merge Sorter')
    .setWidth(800)
    .setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Text Merge Sorter');
}

function getUserId() {
  const email = Session.getActiveUser().getEmail();
  return email && email !== "" ? email : "User_" + SpreadsheetApp.getActiveSpreadsheet().getId().substring(0, 5);
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userId = getUserId();

  const stateSheet = ss.getSheetByName(CONFIG.STATE_SHEET);
  if (stateSheet) {
    const data = stateSheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userId && data[i][1]) return JSON.parse(data[i][1]);
    }
  }

  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No data found in rows 2+');

  const rawData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const songList = rawData.map(row => ({
    name: row[0] ? row[0].toString().trim() : "Unknown"
  })).filter(s => s.name !== "");

  return {
    userId: userId,
    songList: songList,
    queue: songList.map((_, i) => [i]),
    leftArr: [],
    rightArr: [],
    merged: [],
    history: [],
    status: 'STARTING'
  };
}

function backgroundSave(stateJson) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.STATE_SHEET) || ss.insertSheet(CONFIG.STATE_SHEET);
  if (!sheet.isSheetHidden()) sheet.hideSheet();

  const userId = getUserId();
  const data = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === userId) { row = i + 1; break; }
  }
  if (row === -1) {
    sheet.appendRow([userId, ""]);
    row = sheet.getLastRow();
  }
  sheet.getRange(row, 2).setValue(stateJson);
}

/**
 * Saves and integrates directly with the "Paste Rankings Here" source of truth.
 */
function saveFinalResult(userId, songList, resultIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TAB_NAME = 'Paste Rankings Here';
  const sheet = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);

  const headers = sheet.getRange(1, 1, 1, Math.max(2, sheet.getLastColumn())).getValues()[0];
  let colIndex = -1;
  const userName = userId.split('@')[0];

  for (let i = 1; i < headers.length; i++) {
    if (headers[i] === userName) { colIndex = i + 1; break; }
  }

  if (colIndex === -1) {
    colIndex = 2;
    for (let i = 1; i < headers.length; i++) {
      if (!headers[i]) { colIndex = i + 1; break; }
      colIndex = i + 2;
    }
    sheet.getRange(1, colIndex).setValue(userName).setFontWeight('bold');
    sheet.getRange(1, 1).setValue("User Name").setFontWeight('bold');
    sheet.getRange(2, 1).setValue("Ranked List").setFontWeight('bold');
  }

  const finalRows = resultIndices.map((idx, rank) => [`${rank + 1}. ${songList[idx].name}`]);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, colIndex, lastRow - 1, 1).clearContent();
  sheet.getRange(2, colIndex, finalRows.length, 1).setValues(finalRows);

  SpreadsheetApp.getUi().alert("Results saved for " + userName + "! You can now run the 'Sync ALL Sheets' command to distribute your rankings.");
}