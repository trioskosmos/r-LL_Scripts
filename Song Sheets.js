/**
 * --- CONFIGURATION ---
 * All group tabs are defined dynamically in the 'Sheet Manager' tab.
 * Add a "Global" or "All Songs" tab there if you want a master list!
 */
const STATIC_TARGETS = [];

const CONFIG_TAB_NAME = "Sheet Manager";

/**
 * Returns combined list of static and user-defined tabs from Sheet Manager.
 */
function getTargetSheetConfigs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configs = [...STATIC_TARGETS];

  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
  if (configSheet) {
    const lastCol = configSheet.getLastColumn();
    const lastRow = configSheet.getLastRow();
    if (lastCol >= 1 && lastRow >= 1) {
      const data = configSheet.getRange(1, 1, lastRow, lastCol).getValues();

      // Each column is one custom tab
      for (let col = 0; col < lastCol; col++) {
        const tabName = String(data[0][col]).trim();
        // Skip labels or empty headers
        if (!tabName || tabName === "Custom Tab Name" || tabName === "Artist Reference") continue;

        const ids = [];
        for (let row = 1; row < lastRow; row++) {
          const val = String(data[row][col]).trim();
          // Skip labels like "ID:" or "Song:" and empty values
          if (val && !/^id:$/i.test(val) && !/^song:$/i.test(val)) {
            ids.push(val);
          }
        }

        if (ids.length > 0) {
          configs.push({
            tabName: tabName,
            condition: (baseRow) => {
              return ids.some(id => {
                // Split ONLY by newlines to support pasted lists, but keep semicolons 
                // intact as they are often part of the Artist Info/ID strings.
                const individualItems = id.split('\n').map(s => s.trim()).filter(s => s !== "");

                return individualItems.some(item => {
                  // Normalize: lowercase and strip rank prefix (e.g., "1. ")
                  let search = item.toLowerCase().replace(/^\d+\.\s+/, "").trim();
                  if (!search) return false;

                  // Extract just the song name if format is "Song - Artist"
                  const searchBase = search.split(/\s+-\s+/)[0].trim();

                  // Check common columns in Base sheet: A(0), B(1), H(7)
                  const checkCols = [0, 1, 7];
                  return checkCols.some(idx => {
                    const rowVal = baseRow[idx];
                    if (rowVal === undefined || rowVal === null) return false;

                    const cellVal = String(rowVal).trim().toLowerCase();
                    if (!cellVal) return false;

                    // 1. Exact match (best for specific IDs or clean song names)
                    if (cellVal === search || cellVal === searchBase) return true;
                    // 2. Contains (handles ID:96 in ID:96;ID:97 or "Song" in "Song (TV Size)")
                    if (cellVal.includes(search) || cellVal.includes(searchBase)) return true;
                    // 3. Reverse contains (if Base has "Song Name" and search is just "Song")
                    if (search.length > 3 && cellVal.includes(search)) return true;
                    if (searchBase.length > 3 && cellVal.includes(searchBase)) return true;

                    return false;
                  });
                });
              });
            }
          });
        }
      }
    }
  }

  return configs;
}

/**
 * Creates a custom menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Ensure UI elements are ready
  setupPasteSheetUI();

  // 1. UPDATE RANKINGS MENU
  ui.createMenu('📊 Update System (v3)')
    .addItem('🚀 1. Sync & Update ALL Sheets', 'syncFromInbox')
    .addSeparator()
    .addItem('🔄 2. Update Membership Only', 'updateFilteredTabs')
    .addSeparator()
    .addItem('🔍 Generate Artist Reference', 'generateArtistReference')
    .addItem('🛠️ Setup Sheet Manager', 'setupSheetManager')
    .addItem('❓ Update FAQ Tab', 'setupFAQ')
    .addItem('🐞 Run Diagnostics', 'runDiagnostics')
    .addToUi();

  ui.createMenu('🛠️ Analysis')
    .addItem('🚀 Run Opps Analysis (Friends/Rivals)', 'runFullAnalysis')
    .addItem('🌶️ Run Hot Takes & Glazes', 'runHotTakesAnalysis')
    .addItem('📈 Run More Analysis', 'runMoreAnalysis')
    .addItem('🧪 Run Spice Index', 'runSpiceAnalysis')
    .addToUi();
}

/**
 * Opens a dialog box for pasting the text list and entering a name.
 */
function showTextImportDialog() {
  const html = HtmlService.createHtmlOutput(
    '<style>body{font-family:sans-serif; padding:10px;} textarea{width:100%; height:200px; margin-top:5px;} input{width:100%; padding:8px; margin-bottom:10px; box-sizing: border-box;}</style>' +
    '<b>1. Enter Your Name:</b><br>' +
    '<input type="text" id="userName" placeholder="Header for this column..."><br>' +
    '<b>2. Paste List (Format: Rank. Song - Artist):</b><br>' +
    '<textarea id="textData" placeholder="1. Song Name - Artist"></textarea><br><br>' +
    '<button id="btn" style="width:100%; padding:10px; background:#4285f4; color:white; border:none; cursor:pointer;" onclick="submit()">Process & Close</button>' +
    '<script>' +
    '  function submit() {' +
    '    var name = document.getElementById("userName").value;' +
    '    var text = document.getElementById("textData").value;' +
    '    if(!name || !text) { alert("Please fill in both fields"); return; }' +
    '    document.getElementById("btn").disabled = true;' +
    '    document.getElementById("btn").innerText = "Processing...";' +
    '    google.script.run' +
    '      .withSuccessHandler(function() { google.script.host.close(); })' +
    '      .processPastedText(text, name);' +
    '  }' +
    '</script>'
  ).setWidth(500).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, 'Import Ranked List');
}

/**
 * Parses text, finds next column, uses user name as header.
 */
function processPastedText(text, userName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const configs = getTargetSheetConfigs();
  const config = configs.find(c => c.tabName === sheet.getName());
  if (!config) {
    ui.alert("This sheet is not a configured group tab.");
    return;
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
  let targetColIndex = 5;
  for (let i = 4; i < headerRow.length; i++) {
    if (headerRow[i] === "" || headerRow[i] === null) {
      targetColIndex = i + 1;
      break;
    }
    if (i === headerRow.length - 1) targetColIndex = i + 2;
  }

  sheet.getRange(1, targetColIndex).setValue(userName);

  const lines = text.split('\n');
  const externalMap = {};
  const regex = /^(\d+)\.\s+(.+?)\s+-\s+(.+)$/;

  lines.forEach(line => {
    let match = line.trim().match(regex);
    if (match) {
      externalMap[match[2].trim()] = parseInt(match[1]);
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const songNamesInSheet = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const outputColumn = songNamesInSheet.map(row => [externalMap[row[0].toString().trim()] || ""]);

  sheet.getRange(2, targetColIndex, outputColumn.length, 1).setValues(outputColumn);
  sortByPoints();
  ss.toast("Scores updated for " + userName);
}

/**
 * Sorts by Points (Column C) - Lowest to Highest.
 */
function sortByPoints() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!isTargetSheet(sheet.getName())) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const sortRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  sortRange.sort({ column: 3, ascending: true });

  const rankValues = [];
  for (let i = 1; i <= lastRow - 1; i++) { rankValues.push([i]); }
  sheet.getRange(2, 1, rankValues.length, 1).setValues(rankValues);
}

/**
 * Sorts Alphabetically by Song Name (Column B).
 */
function sortAlphabetically() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!isTargetSheet(sheet.getName())) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const sortRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  sortRange.sort({ column: 2, ascending: true });
  sheet.getRange(2, 1, lastRow - 1, 1).setValue("-");
}

function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function isTargetSheet(name) {
  return getTargetSheetConfigs().some(config => config.tabName === name);
}

/**
 * FEATURE: Setup the Sheet Manager tab for user-defined song lists.
 */
function setupSheetManager() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME) || ss.insertSheet(CONFIG_TAB_NAME);
  if (sheet.getLastRow() === 0) {
    // Set up example structure
    const examples = [
      ['Global / All Songs', 'Specific Group 1'],
      ['ID:', 'ID:001'],
      ['', 'ID:002']
    ];
    sheet.getRange(1, 1, 3, 3).setValues(examples);
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontWeight('bold').setBackground('#efefef');

    SpreadsheetApp.getUi().alert("Sheet Manager set up in 'Vertical Column' mode!\n\n1. Put your Tab Name in Row 1.\n2. Paste your list of Artist IDs (from Base Col H) directly below it.\n3. Run 'Sync Membership from Base' to create/update these tabs.");
  } else {
    ss.toast("Sheet Manager already exists.");
  }
}

/**
 * FEATURE: Creates a single tab listing all unique Artist Strings and their songs.
 */
function generateArtistReference() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base');
  if (!baseSheet) { SpreadsheetApp.getUi().alert("Base sheet not found!"); return; }

  const data = baseSheet.getDataRange().getValues().slice(1);
  const refSheet = ss.getSheetByName('Artist Reference') || ss.insertSheet('Artist Reference');
  refSheet.clear();

  // Group songs by Artist Info string
  const artistMap = {};
  data.forEach(row => {
    const artist = String(row[7]).trim() || "Unknown";
    const song = String(row[1]).trim();
    if (!artistMap[artist]) artistMap[artist] = [];
    artistMap[artist].push(song);
  });

  // Prepare output rows
  const sortedArtists = Object.keys(artistMap).sort();
  const output = [['Artist Info (Col H)', 'Matching Songs']];

  sortedArtists.forEach(artist => {
    output.push([artist, artistMap[artist].join(', ')]);
  });

  refSheet.getRange(1, 1, output.length, 2).setValues(output);
  refSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#efefef');
  refSheet.setColumnWidth(1, 400);
  refSheet.setColumnWidth(2, 600);
  refSheet.getRange("A:B").setWrap(true).setVerticalAlignment("top");

  ss.toast("Artist Reference tab generated!");
}

/**
 * Syncs membership from 'Base' and creates dynamic formulas for Points and Average.
 */
function updateFilteredTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base');
  if (!baseSheet) { SpreadsheetApp.getUi().alert("Base sheet not found!"); return; }

  const baseRows = baseSheet.getDataRange().getValues().slice(1);

  const configs = getTargetSheetConfigs();
  configs.forEach(config => {
    let targetSheet = ss.getSheetByName(config.tabName) || ss.insertSheet(config.tabName);

    let existingData = targetSheet.getDataRange().getValues();
    let scoreMap = {};
    let headers = ["Rank", "Song", "Points", "Average"];

    if (existingData.length > 0) {
      let currentHeaders = existingData[0];
      if (currentHeaders.length > 4) { headers = currentHeaders; }
      if (existingData.length > 1) {
        existingData.slice(1).forEach(row => {
          if (row[1]) scoreMap[row[1]] = row.slice(4);
        });
      }
    }

    let currentSongs = [];
    baseRows.forEach(row => {
      if (config.condition(row)) {
        let name = row[1];
        let scores = scoreMap[name] || [];
        let userColCount = Math.max(0, headers.length - 4);
        while (scores.length < userColCount) scores.push("");
        currentSongs.push({ name: name, scores: scores });
      }
    });

    let finalRows = currentSongs.map((item, index) => {
      let rIdx = index + 2;
      // Use "infinite row" range (e.g., E2:2) so formulas work even as columns are added later
      let sumFormula = `=SUM(E${rIdx}:${rIdx})`;
      let avgFormula = `=IFERROR(AVERAGE(E${rIdx}:${rIdx}), 0)`;

      return [index + 1, item.name, sumFormula, avgFormula, ...item.scores];
    });

    targetSheet.clearContents();
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    if (finalRows.length > 0) {
      targetSheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
      targetSheet.getRange(2, 4, finalRows.length, 1).setNumberFormat("0.00");
    }
  });
  ss.toast("Membership synced.");
}


/**
 * CORE FEATURE: Process the "Paste Rankings Here" sheet and distribute relative ranks.
 */
function syncFromInbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 0. Ensure Membership is synced first (catch new songs/tabs)
  updateFilteredTabs();

  const TAB_NAME = 'Paste Rankings Here';
  const LOG_TAB = 'Sync Log';
  const inboxSheet = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);

  // 1. Prepare/Verify Inbox Structure
  if (inboxSheet.getLastRow() < 1 || inboxSheet.getRange(1, 1).getValue() !== "User Name") {
    inboxSheet.clear();
    inboxSheet.getRange(1, 1).setValue("User Name").setFontWeight('bold');
    inboxSheet.getRange(2, 1).setValue("Ranked List").setFontWeight('bold');
    inboxSheet.setColumnWidth(1, 120);
    SpreadsheetApp.getUi().alert("I've set up the '" + TAB_NAME + "' sheet. \n\nColumn A contains labels. \nStart from Column B: Put the User Name in Row 1 and paste their list in Row 2 downwards!");
    return;
  }

  // Diagnostics Tab
  const logSheet = ss.getSheetByName(LOG_TAB) || ss.insertSheet(LOG_TAB);
  logSheet.clear().appendRow(['Time', 'User', 'Status', 'Details']).getRange(1, 1, 1, 4).setFontWeight('bold');
  const addLog = (user, status, details) => logSheet.appendRow([new Date().toLocaleTimeString(), user, status, details]);

  const lastCol = inboxSheet.getLastColumn();
  if (lastCol < 2) {
    addLog('System', 'Error', 'No user columns found starting from Column B');
    SpreadsheetApp.getUi().alert("Please add at least one user's data starting from Column B.");
    return;
  }

  const fullData = inboxSheet.getDataRange().getValues();
  const userMap = {};
  const regex = /^(\d+)\.\s+(.+?)(?:\s+-\s+.+)?$/;

  // 2. PARSE COLUMNS
  for (let col = 1; col < lastCol; col++) {
    const userName = String(fullData[0][col]).trim();
    if (!userName) continue;

    let parsedRows = [];
    for (let row = 1; row < fullData.length; row++) {
      let cellValue = String(fullData[row][col]).trim();
      if (!cellValue) continue;

      let lines = cellValue.split('\n');
      lines.forEach(line => {
        let match = line.trim().match(regex);
        if (match) {
          parsedRows.push({
            originalRank: parseInt(match[1]),
            songName: match[2].trim()
          });
        }
      });
    }

    if (parsedRows.length > 0) {
      userMap[userName] = { list: parsedRows, count: parsedRows.length };
    }
  }

  if (Object.keys(userMap).length === 0) {
    addLog('System', 'Error', 'No valid rankings found in columns');
    SpreadsheetApp.getUi().alert("No valid rankings found. Ensure format is '1. Song Name'");
    return;
  }

  // 3. DISTRIBUTE TO SUB-SHEETS
  let totalUpdates = 0;
  let totalCleared = 0;
  const validUserNames = Object.keys(userMap).map(n => n.toLowerCase());

  const configsSync = getTargetSheetConfigs();
  configsSync.forEach(config => {
    const targetSheet = ss.getSheetByName(config.tabName);
    if (!targetSheet) return;

    const sheetData = targetSheet.getDataRange().getValues();
    if (sheetData.length < 2) return;

    const headers = sheetData[0];
    const normalizedHeaders = headers.map(h => String(h).trim().toLowerCase());

    // --- CLEANUP STEP: Clear users NOT in the userMap ---
    for (let i = 4; i < headers.length; i++) {
      const headerName = String(headers[i]).trim();
      if (!headerName) continue;

      if (!validUserNames.includes(headerName.toLowerCase())) {
        // This user is no longer in the master list - Clear their column
        const lastRowInTarget = targetSheet.getLastRow();
        if (lastRowInTarget > 1) {
          targetSheet.getRange(2, i + 1, lastRowInTarget - 1, 1).clearContent();
          addLog('System', 'Cleanup', `Cleared user "${headerName}" from tab "${config.tabName}" (not found in master paste list).`);
          totalCleared++;
        }
      }
    }

    // Normalizing song names in the sheet for robustness
    const songMapInSheet = {};
    sheetData.slice(1).forEach((r, idx) => {
      const name = String(r[1]).trim();
      songMapInSheet[name.toLowerCase()] = { originalName: name, rowIndex: idx };
    });

    // Headers are already defined above for the cleanup step

    Object.keys(userMap).forEach(userName => {
      const userRanking = userMap[userName];
      let matchedCount = 0;
      let failedMatches = [];
      const matchedSongs = [];

      userRanking.list.forEach(item => {
        const normalizedInput = item.songName.toLowerCase();
        if (songMapInSheet[normalizedInput]) {
          matchedSongs.push({
            ...item,
            sheetName: songMapInSheet[normalizedInput].originalName,
            rowIndex: songMapInSheet[normalizedInput].rowIndex
          });
          matchedCount++;
        } else {
          failedMatches.push(item.songName);
        }
      });

      if (matchedSongs.length === 0) {
        addLog(userName, 'Skip', `Tab "${config.tabName}": 0 songs matched.`);
        return;
      }

      matchedSongs.sort((a, b) => a.originalRank - b.originalRank);

      const relativeMap = {};
      let currentRelRank = 0;
      let lastOrigRank = -1;

      matchedSongs.forEach(item => {
        if (item.originalRank > lastOrigRank) currentRelRank++;
        relativeMap[item.sheetName] = currentRelRank;
        lastOrigRank = item.originalRank;
      });

      // Find or create User Column - scan header row for existing name or first empty slot after D
      let colIndex = normalizedHeaders.indexOf(userName.toLowerCase()) + 1;
      if (colIndex === 0) {
        // User doesn't exist yet - find the first empty column starting from E (index 5)
        colIndex = 5; // Start at Column E
        for (let i = 4; i < headers.length; i++) {
          if (headers[i] === "" || headers[i] === null || headers[i] === undefined) {
            colIndex = i + 1;
            break;
          }
          colIndex = i + 2; // If all are full, use next after last
        }
        targetSheet.getRange(1, colIndex).setValue(userName).setFontWeight('bold');
      }

      // Write values efficiently
      const songNamesList = sheetData.slice(1).map(r => String(r[1]).trim());
      const output = songNamesList.map(name => [relativeMap[name] || ""]);
      targetSheet.getRange(2, colIndex, output.length, 1).setValues(output);

      totalUpdates++;
      const colLetter = columnToLetter(colIndex);
      addLog(userName, 'Success', `Tab "${config.tabName}": matched ${matchedCount}/${userRanking.count} songs. Written to Col ${colLetter}. ${failedMatches.length > 0 ? "Missed: " + failedMatches.join(", ") : ""}`);
    });
  });

  // 4. RE-RANK all target sheets by points
  const configs = getTargetSheetConfigs();
  configs.forEach(config => {
    const sheet = ss.getSheetByName(config.tabName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const lastCol = sheet.getLastColumn();
    const sortRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    sortRange.sort({ column: 3, ascending: true }); // Sort by Points (Column C)

    // Re-sequence Rank (Column A)
    const rankValues = [];
    for (let i = 1; i <= lastRow - 1; i++) { rankValues.push([i]); }
    sheet.getRange(2, 1, rankValues.length, 1).setValues(rankValues);
  });

  ss.toast(`Sync Complete! Updates: ${totalUpdates}, Cleared: ${totalCleared}. All custom tabs synced.`);

  // 4. Run Analysis automatically
  try {
    runFullAnalysis();
    runHotTakesAnalysis();
    runMoreAnalysis();
    runSpiceAnalysis();
    ss.toast("Sync & All Analyses Complete!");
  } catch (e) {
    ss.toast("Sync finished, but analysis encountered an error.");
  }
}

/**
 * FEATURE: Setup or Update the FAQ tab.
 */
function setupFAQ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const faqSheet = ss.getSheetByName('FAQ') || ss.insertSheet('FAQ');
  faqSheet.clear();
  faqSheet.setColumnWidth(1, 40);
  faqSheet.setColumnWidth(2, 300);
  faqSheet.setColumnWidth(3, 600);

  const content = [
    ['#', 'TOPIC', 'INSTRUCTIONS / DETAILS'],
    ['1', '🚀 HOW DO I UPDATE?', 'GO TO: "📊 Update System" > "🚀 1. Sync & Update ALL Sheets".\nThis is the "Golden Button" that does EVERYTHING: Matches songs, updates every group tab, and refreshes all analysis.'],
    ['', '', ''],
    ['2', '📝 RANKING SONGS', '1. Go to "Paste Rankings Here" tab.\n2. Put your name in Row 1 (Starting Column B).\n3. Paste your list ("1. Song Name") below it.'],
    ['', '', ''],
    ['3', '📂 CUSTOM TABS', 'Use the "Sheet Manager" tab. Define a Tab Name and the Artist IDs for that group. The system will handle the song filtering automatically.'],
    ['', '', ''],
    ['4', '🔄 MEMBERSHIP ONLY', 'If you just changed a group list in Sheet Manager but haven\'t changed any user rankings, run "🔄 2. Update Membership Only" to quickly refresh the song list in those tabs.'],
    ['', '', ''],
    ['5', '🌶️ ANALYSES', 'Analysis runs automatically during a full sync. You can also run "Opps" or "Hot Takes" manually from the "🛠️ Analysis" menu.'],
    ['', '', ''],
    ['6', '⚡ SYNC LOG', 'Check the "Sync Log" tab after any update. It will tell you if any songs failed to match (typos, etc).'],
    ['', '', ''],
    ['---', 'NEED HELP?', 'If the menus at the top are missing, just REFRESH the browser page!']
  ];

  const range = faqSheet.getRange(1, 1, content.length, 3);
  range.setValues(content).setVerticalAlignment("top").setWrap(true);

  // Styling
  faqSheet.getRange("1:1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold");
  faqSheet.getRange("A:B").setFontWeight("bold");
  faqSheet.getRange("C1").setHorizontalAlignment("left");

  ss.toast("FAQ Refreshed for clarity!");
}

/**
 * Ensures the Paste sheet has a visible prompt/button area.
 */
function setupPasteSheetUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Paste Rankings Here') || ss.insertSheet('Paste Rankings Here');

  const cell = sheet.getRange("A1");
  if (cell.getValue() !== "User Name") {
    sheet.getRange("A1:A2").setValues([["User Name"], ["(Put lists below)"]]).setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setColumnWidth(1, 150);
  }

  // Adding a high-visibility instruction cell
  const instructionCell = sheet.getRange("A4");
  instructionCell.setValue("👉 CLICK SYNC IN THE MENU ABOVE TO UPDATE ALL SHEETS 🚀")
    .setFontWeight("bold")
    .setFontColor("red")
    .setBackground("#fff2cc");

  sheet.getRange("A4:A5").merge();
}

/**
 * DEBUG TOOL: Diagnoses why songs might not be matching.
 */
function runDiagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debugSheet = ss.getSheetByName('Debug Log') || ss.insertSheet('Debug Log');
  debugSheet.clear();

  const log = [['Type', 'Category', 'Details']];

  // 1. Map Base Sheet
  const baseSheet = ss.getSheetByName('Base');
  if (!baseSheet) {
    log.push(['Error', 'Infrastructure', 'Base Sheet not found!']);
  } else {
    const baseHeaders = baseSheet.getRange(1, 1, 1, baseSheet.getLastColumn()).getValues()[0];
    log.push(['Info', 'Base Header Map', baseHeaders.map((h, i) => `[Col ${i + 1}: ${h}]`).join(' | ')]);

    // Find song column
    const songIndex = baseHeaders.findIndex(h => /song/i.test(String(h)));
    log.push(['Info', 'Base Sheet', `Found "Song" at index ${songIndex} (Col ${songIndex + 1})`]);
  }

  // 2. Check Sheet Manager Configs
  const configs = getTargetSheetConfigs();
  log.push(['Info', 'Configuration', `${configs.length} Tabs detected in Sheet Manager`]);

  configs.forEach(conf => {
    log.push(['Section', 'Processing Tab', `>>> ${conf.tabName} <<<`]);

    const targetSheet = ss.getSheetByName(conf.tabName);
    if (!targetSheet) {
      log.push(['Warning', conf.tabName, 'Sheet does not exist yet. Run Membership Sync.']);
    } else {
      const headers = targetSheet.getRange(1, 1, 1, Math.max(1, targetSheet.getLastColumn())).getValues()[0];
      const dataRows = targetSheet.getLastRow() - 1;
      log.push(['Info', conf.tabName, `Sheet found. Rows: ${dataRows}, Raw Headers: ${JSON.stringify(headers)}`]);

      // Analyze user columns for Matrix requirements
      const userCols = headers.slice(4).filter(h => h !== "" && h !== null);
      log.push(['Analysis', conf.tabName, `User Columns detected for Matrix: ${JSON.stringify(userCols)}`]);
      if (userCols.length < 2) {
        log.push(['Critical', conf.tabName, 'DIVERGENCE MATRIX WILL FAIL: Need at least 2 users with ranking columns.']);
      }
    }

    // 3. Test Matching Logic & Opps Eligibility for this config
    if (baseSheet) {
      const baseData = baseSheet.getDataRange().getValues();
      const baseRows = baseData.slice(1);

      const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
      let searchIds = [];
      if (configSheet) {
        const fullConfig = configSheet.getDataRange().getValues();
        const headerRow = fullConfig[0];
        const colIdx = headerRow.findIndex(h => String(h).trim().toLowerCase() === conf.tabName.toLowerCase());
        if (colIdx !== -1) {
          searchIds = fullConfig.slice(1).map(r => String(r[colIdx]).trim()).filter(v => v && !/^id:$/i.test(v) && !/^song:$/i.test(v));
        }
      }
      log.push(['Info', conf.tabName, `Looking for these items: ${JSON.stringify(searchIds)}`]);

      let matchCount = 0;
      let missedIds = [...searchIds];

      baseRows.forEach((row, idx) => {
        if (conf.condition(row)) {
          matchCount++;
          const songNameLower = String(row[1] || row[0]).toLowerCase();
          const artistInfoLower = String(row[7] || "").toLowerCase();
          const col0Lower = String(row[0] || "").toLowerCase();

          missedIds = missedIds.filter(id => {
            const cleanId = id.toLowerCase();
            return !songNameLower.includes(cleanId) &&
              !artistInfoLower.includes(cleanId) &&
              !col0Lower.includes(cleanId);
          });
        }
      });

      log.push(['Result', conf.tabName, `Total matches in Base: ${matchCount}`]);
      if (missedIds.length > 0) {
        log.push(['FAIL', conf.tabName, `FAILED TO FIND: ${JSON.stringify(missedIds)}`]);
      }

      // --- NEW: OPPS ELIGIBILITY CHECK ---
      if (targetSheet && targetSheet.getLastRow() > 1) {
        const tData = targetSheet.getDataRange().getValues();
        const tHeaders = tData[0];
        const tRows = tData.slice(1);
        const sysCols = ['Rank', 'Song', 'Points', 'Average'];

        const uCols = [];
        tHeaders.forEach((h, idx) => {
          const cH = String(h || "").trim();
          if (cH && !sysCols.some(s => s.toLowerCase() === cH.toLowerCase())) {
            uCols.push({ name: cH, idx: idx });
          }
        });

        if (uCols.length >= 2) {
          const vRows = tRows.filter(r => uCols.every(u => r[u.idx] !== "" && !isNaN(parseFloat(r[u.idx]))));
          log.push(['Opps Check', conf.tabName, `Matrix Status: ${vRows.length > 0 ? "READY" : "FAILED"}. Details: ${uCols.length} users found, and found ${vRows.length} songs have scores from EVERY user.`]);

          if (vRows.length === 0) {
            uCols.forEach(u => {
              const hasDataCount = tRows.filter(r => r[u.idx] !== "" && !isNaN(parseFloat(r[u.idx]))).length;
              log.push(['Opps Hint', conf.tabName, `User "${u.name}" has scores for ${hasDataCount}/${tRows.length} songs.`]);
            });
          }
        }
      }
    }
  });

  debugSheet.clear();
  debugSheet.getRange(1, 1, log.length, 3).setValues(log);
  debugSheet.setColumnWidth(1, 120);
  debugSheet.setColumnWidth(2, 200);
  debugSheet.setColumnWidth(3, 800);
  debugSheet.getRange(1, 1, log.length, 3).setWrap(true);

  SpreadsheetApp.getUi().alert("EXHAUSTIVE Diagnostics run! Check the 'Debug Log' tab.");
}