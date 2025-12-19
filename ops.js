/**
 * Configuration: This matches your sync script structure.
 */
const NON_USER_COLUMNS = ['Rank', 'Song', 'Points', 'Average'];

/**
 * Perform divergence analysis across all group sheets.
 */
function runFullAnalysis() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const LOG_CATEGORY = 'Analysis Debug';
    const log = (msg, details) => {
        const logSheet = ss.getSheetByName('Debug Log') || ss.insertSheet('Debug Log');
        logSheet.appendRow([LOG_CATEGORY, msg, details || '']);
    };

    log('Start', 'Initiating full analysis sequence');

    try {
        let oppsSheet = ss.getSheetByName('Opps') || ss.insertSheet('Opps');
        log('Sheet Init', 'Opps sheet identified and cleared');
        oppsSheet.clear();

        let allSheetsData = [];
        let globalPartnerCounts = {};
        let globalDivergence = {};
        let allUniqueUsers = new Set();
        let groupConsensus = [];

        const configs = getTargetSheetConfigs();
        log('Config Check', `Targeting ${configs.length} potential group tabs`);

        configs.forEach(config => {
            let sheet = ss.getSheetByName(config.tabName);
            if (!sheet) {
                log('Skip Tab', `${config.tabName} sheet not found in spreadsheet`);
                return;
            }
            log('Tab Processing', `Processing tab: ${config.tabName}`);

            let data = sheet.getDataRange().getValues();
            if (data.length < 2) {
                log('Skip Tab', `${config.tabName} has no data rows`);
                return;
            }
            log('Tab Processing', `${config.tabName}: Found ${data.length - 1} data rows`);

            let headers = data[0];
            let rows = data.slice(1);

            let userCols = [];
            headers.forEach((h, idx) => {
                const cleanH = String(h || "").trim();
                const sysCols = ['Rank', 'Song', 'Points', 'Average'];
                if (cleanH !== "" && !sysCols.some(sysCol => sysCol.toLowerCase() === cleanH.toLowerCase())) {
                    let hasData = rows.some(r => r[idx] !== "" && r[idx] !== null && r[idx] !== undefined);
                    if (hasData) {
                        userCols.push({ name: cleanH, index: idx });
                        allUniqueUsers.add(cleanH);
                    }
                }
            });

            log('Tab Processing', `${config.tabName}: Detected ${userCols.length} users: ${userCols.map(u => u.name).join(', ')}`);

            if (userCols.length < 2) {
                log('Skip Tab', `${config.tabName}: Not enough users (${userCols.length})`);
                return;
            }

            let validRows = rows.filter(row => userCols.every(u => row[u.index] !== "" && !isNaN(parseFloat(row[u.index]))));

            log('Tab Processing', `${config.tabName}: Found ${validRows.length} valid rows for analysis`);

            if (validRows.length === 0) {
                log('Skip Tab', `${config.tabName}: 0 songs shared by all users`);
                return;
            }

            let matrix = {};
            let totalGroupDiv = 0;
            let pairCount = 0;

            userCols.forEach(u1 => {
                matrix[u1.name] = {};
                if (!globalDivergence[u1.name]) globalDivergence[u1.name] = {};

                userCols.forEach(u2 => {
                    let rawDiff = validRows.reduce((sum, row) => sum + Math.abs(row[u1.index] - row[u2.index]), 0);
                    let normalizedDiff = (rawDiff / Math.pow(validRows.length, 2)) * 100;
                    matrix[u1.name][u2.name] = normalizedDiff;

                    if (!globalDivergence[u1.name][u2.name]) globalDivergence[u1.name][u2.name] = { total: 0, count: 0 };
                    globalDivergence[u1.name][u2.name].total += normalizedDiff;
                    globalDivergence[u1.name][u2.name].count++;

                    if (u1.name !== u2.name) {
                        totalGroupDiv += normalizedDiff;
                        pairCount++;
                    }
                });
            });
            log('Tab Processing', `${config.tabName}: Divergence matrix calculated`);

            let userPartners = [];
            userCols.forEach(u1 => {
                let scores = Object.keys(matrix[u1.name])
                    .filter(name => name !== u1.name)
                    .map(name => ({ name: name, score: matrix[u1.name][name] }))
                    .sort((a, b) => a.score - b.score);

                if (scores.length > 0) {
                    let least = scores[0];
                    let most = scores[scores.length - 1];
                    userPartners.push([u1.name, most.name, most.score.toFixed(2), least.name, least.score.toFixed(2)]);
                    updateCount(globalPartnerCounts, most.name, 'most');
                    updateCount(globalPartnerCounts, least.name, 'least');
                }
            });
            log('Tab Processing', `${config.tabName}: User partners identified`);

            groupConsensus.push([
                config.tabName,
                (totalGroupDiv / pairCount).toFixed(2),
                validRows.length + " songs"
            ]);

            allSheetsData.push({
                name: config.tabName,
                matrix: matrix,
                userNames: userCols.map(u => u.name).sort(),
                partners: userPartners
            });
            log('Tab Processing', `${config.tabName}: Data aggregated`);
        });

        if (allSheetsData.length === 0) {
            oppsSheet.getRange(1, 1).setValue("No divergence data found across any tabs. Check shared song scores.");
            log('End', 'Finished with 0 results');
            return;
        }

        // --- PRINTING ---
        log('Printing', `Writing results to Opps tab for ${allSheetsData.length} groups`);
        let cursor = 1;
        let sortedUniqueUsers = Array.from(allUniqueUsers).sort();

        oppsSheet.getRange(cursor, 1).setValue('--- GROUP CONSENSUS SUMMARY ---').setFontWeight('bold').setFontSize(12);
        cursor += 2;
        let gcHeader = [['Group Name', 'Avg. Internal Divergence (Lower = More Unified)', 'Songs Analyzed']];
        oppsSheet.getRange(cursor, 1, 1, 3).setValues(gcHeader).setBackground('#fff2cc').setFontWeight('bold');
        if (groupConsensus.length > 0) {
            oppsSheet.getRange(cursor + 1, 1, groupConsensus.length, 3).setValues(groupConsensus.sort((a, b) => a[1] - b[1]));
            cursor += groupConsensus.length + 3;
        }
        log('Printing', 'Group Consensus Summary written');

        oppsSheet.getRange(cursor, 1).setValue('--- GLOBAL: User-to-User Divergence Matrix ---').setFontWeight('bold').setFontSize(12);
        cursor += 2;
        let gHeader = [['User Name', ...sortedUniqueUsers]];
        let gRows = sortedUniqueUsers.map(u1 => [u1, ...sortedUniqueUsers.map(u2 => {
            let div = globalDivergence[u1] ? globalDivergence[u1][u2] : null;
            return (div && div.count > 0) ? (div.total / div.count).toFixed(2) : "0.00";
        })]);
        oppsSheet.getRange(cursor, 1, 1, gHeader[0].length).setValues(gHeader).setBackground('#cfe2f3').setFontWeight('bold');
        oppsSheet.getRange(cursor + 1, 1, gRows.length, gRows[0].length).setValues(gRows);
        cursor += gRows.length + 3;
        log('Printing', 'Global Divergence Matrix written');

        allSheetsData.forEach(res => {
            oppsSheet.getRange(cursor, 1).setValue(`--- TAB: ${res.name} (Per-Group Matrix) ---`).setFontWeight('bold').setFontSize(12).setFontColor('red');
            cursor += 2;

            let lmHeader = [['User Name', ...res.userNames]];
            let lmRows = res.userNames.map(u1 => [u1, ...res.userNames.map(u2 => res.matrix[u1][u2].toFixed(2))]);
            oppsSheet.getRange(cursor, 1, 1, lmHeader[0].length).setValues(lmHeader).setBackground('#eeeeee').setFontWeight('bold');
            oppsSheet.getRange(cursor + 1, 1, lmRows.length, lmRows[0].length).setValues(lmRows);
            cursor += lmRows.length + 2;

            oppsSheet.getRange(cursor, 1, 1, 5).setValues([['User', 'Rival (Most Diff)', 'Score', 'Friend (Least Diff)', 'Score']]).setBackground('#f3f3f3').setFontStyle('italic');
            oppsSheet.getRange(cursor + 1, 1, res.partners.length, 5).setValues(res.partners);
            cursor += res.partners.length + 4;
            log('Printing', `Per-Group Matrix and Partners for ${res.name} written`);
        });

        log('End', 'Analysis successfully printed to Opps tab');
        ss.toast("Opps analysis complete!");
    } catch (e) {
        log('CRITICAL ERROR', e.message + " | " + e.stack);
        ss.toast("Analysis encountered an error. Check Debug Log.");
    }
}

/**
 * FEATURE: Hot Takes & Glazes Analysis
 * Identifies songs where users differ most from group consensus.
 */
function runHotTakesAnalysis() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const takesSheet = ss.getSheetByName('Takes') || ss.insertSheet('Takes');
    takesSheet.clear();

    const configs = getTargetSheetConfigs();
    let startCol = 1;

    configs.forEach(config => {
        const sheet = ss.getSheetByName(config.tabName);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;

        const headers = data[0];
        const rows = data.slice(1);
        const nonUserCols = ['Rank', 'Song', 'Points', 'Average'];

        const userCols = [];
        headers.forEach((h, idx) => {
            if (!nonUserCols.includes(h) && h !== "") {
                if (rows.some(r => r[idx] !== "" && !isNaN(parseFloat(r[idx])))) {
                    userCols.push({ name: h, index: idx });
                }
            }
        });

        if (userCols.length === 0) return;

        const takesInGroup = [];
        rows.forEach(row => {
            const songName = String(row[1]);
            const ranks = userCols.map(u => parseFloat(row[u.index])).filter(r => !isNaN(r));
            if (ranks.length < 2) return;

            const averageRank = ranks.reduce((a, b) => a + b, 0) / ranks.length;

            userCols.forEach(u => {
                const userRank = parseFloat(row[u.index]);
                if (isNaN(userRank)) return;

                const delta = userRank - averageRank;
                const normalized = (delta / rows.length) * 100;

                takesInGroup.push({
                    user: u.name,
                    song: songName,
                    score: normalized,
                    userRank: userRank.toFixed(0),
                    avgRank: averageRank.toFixed(1)
                });
            });
        });

        if (takesInGroup.length === 0) return;

        // --- RENDER GROUP BLOCK ---
        let cursor = 1;
        takesSheet.getRange(cursor, startCol).setValue(`--- GROUP: ${config.tabName} ---`).setFontWeight('bold').setFontSize(14).setFontColor('#444444');
        cursor += 2;

        // A. GLAZES
        takesSheet.getRange(cursor, startCol).setValue('BIGGEST GLAZES').setFontWeight('bold').setFontColor('#4a86e8');
        cursor += 1;
        const glazes = takesInGroup.filter(t => t.score < 0).sort((a, b) => a.score - b.score);
        writeTakesHorizontalBatch(takesSheet, cursor, startCol, glazes);
        cursor += Math.max(2, glazes.length) + 2;

        // B. HOT TAKES
        takesSheet.getRange(cursor, startCol).setValue('HOTTEST TAKES').setFontWeight('bold').setFontColor('#cc0000');
        cursor += 1;
        const hotTakes = takesInGroup.filter(t => t.score > 0).sort((a, b) => b.score - a.score);
        writeTakesHorizontalBatch(takesSheet, cursor, startCol, hotTakes);

        // Format Columns
        takesSheet.setColumnWidth(startCol, 80); // Score
        takesSheet.setColumnWidth(startCol + 1, 100); // User
        takesSheet.setColumnWidth(startCol + 2, 200); // Song
        takesSheet.setColumnWidth(startCol + 3, 50); // Rank
        takesSheet.setColumnWidth(startCol + 4, 50); // Avg

        // Gap Column
        takesSheet.setColumnWidth(startCol + 5, 40);

        // Increment for next group
        startCol += 6;
    });

    ss.toast("Song Dispute and Hot Takes analysis (Side-by-Side) complete!");
}

/**
 * rumi analysis additions below
 */

/**
 * MAIN: Runs all "More Analysis" features - displays 2 per row
 */
function runMoreAnalysis() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const SHEET_NAME = 'More Analysis';
    const configs = getTargetSheetConfigs();

    const allData = collectAllSongData(configs, ss);

    if (Object.keys(allData.songs).length === 0) return;

    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    sheet.clear();

    // Collect all features as data objects
    const features = [
        getMostControversialSongs(allData),
        getMostConsistentlyRanked(allData),
        getHotTakesPerUser(allData),
        getSongDisputeAnalysis(configs, ss),
        getUniversallyTopBottom(allData),
        getComebackSongs(allData),
        getSubunitPopularity(allData),
        getOutlierUsers(allData, configs, ss),
    ].filter(f => f !== null);

    // Display features 2 per row
    displayFeaturesPairwise(sheet, features);
}

/**
 * Display features in pairs (2 side by side)
 */
function displayFeaturesPairwise(sheet, features) {
    const GAP_COLUMNS = 2;

    // Calculate max width of all left-column features
    let maxLeftWidth = 0;
    for (let i = 0; i < features.length; i += 2) {
        const feature1 = features[i];
        const feature1Width = feature1.headers[0].length;
        maxLeftWidth = Math.max(maxLeftWidth, feature1Width);
    }

    // Position of right column (consistent for all pairs)
    const col1 = 1;
    const col2 = col1 + maxLeftWidth + GAP_COLUMNS;

    let currentRow = 1;

    // Display pairs with fixed right column position
    for (let i = 0; i < features.length; i += 2) {
        const feature1 = features[i];
        const feature2 = features[i + 1] || null;

        currentRow = displayFeaturePair(sheet, currentRow, feature1, feature2, col1, col2);
    }
}

/**
 * Display a pair of features side by side at fixed column positions
 */
function displayFeaturePair(sheet, startRow, feature1, feature2, col1, col2) {
    let currentRow = startRow;

    // Write titles
    sheet.getRange(currentRow, col1).setValue(feature1.title)
        .setFontWeight('bold')
        .setFontSize(12)
        .setFontColor(feature1.titleColor);

    if (feature2) {
        sheet.getRange(currentRow, col2).setValue(feature2.title)
            .setFontWeight('bold')
            .setFontSize(12)
            .setFontColor(feature2.titleColor);
    }

    currentRow += 1;

    // Write headers
    const feature1Width = feature1.headers[0].length;
    sheet.getRange(currentRow, col1, 1, feature1Width)
        .setValues([feature1.headers[0]])
        .setBackground(feature1.headerBgColor)
        .setFontWeight('bold');

    if (feature2) {
        const feature2Width = feature2.headers[0].length;
        sheet.getRange(currentRow, col2, 1, feature2Width)
            .setValues([feature2.headers[0]])
            .setBackground(feature2.headerBgColor)
            .setFontWeight('bold');
    }

    currentRow += 1;

    // Write data rows (both features at same vertical position)
    const maxRows = Math.max(feature1.rows.length, feature2 ? feature2.rows.length : 0);

    for (let i = 0; i < maxRows; i++) {
        if (i < feature1.rows.length) {
            const row = feature1.rows[i];
            sheet.getRange(currentRow, col1, 1, row.length)
                .setValues([row])
                .setBackground(feature1.rowBgColors[i] || '#ffffff');
        }

        if (feature2 && i < feature2.rows.length) {
            const row = feature2.rows[i];
            sheet.getRange(currentRow, col2, 1, row.length)
                .setValues([row])
                .setBackground(feature2.rowBgColors[i] || '#ffffff');
        }

        currentRow += 1;
    }

    return currentRow + 2; // Spacing between pairs
}

/**
 * Helper: Collect all song data across all group tabs
 */
function collectAllSongData(configs, ss) {
    const NON_USER_COLUMNS = ['Rank', 'Song', 'Points', 'Average'];
    const songs = {};
    const userRanks = {};

    // Build a map of song names to artists from Base sheet
    const baseSheet = ss.getSheetByName('Base');
    const artistMap = {};
    if (baseSheet) {
        const baseData = baseSheet.getDataRange().getValues();
        for (let i = 1; i < baseData.length; i++) {
            const songName = String(baseData[i][1]).trim();
            const artist = String(baseData[i][7]).trim();
            artistMap[songName] = artist;
        }
    }

    configs.forEach(config => {
        const sheet = ss.getSheetByName(config.tabName);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;

        const headers = data[0];
        const rows = data.slice(1);

        const userCols = [];
        headers.forEach((h, idx) => {
            if (!NON_USER_COLUMNS.includes(h) && h !== "") {
                const hasData = rows.some(r => r[idx] !== "" && !isNaN(parseFloat(r[idx])));
                if (hasData) userCols.push({ name: h, index: idx });
            }
        });

        rows.forEach(row => {
            const songName = String(row[1]);
            const artist = artistMap[songName] || "";

            if (!songs[songName]) {
                songs[songName] = { ranks: [], artist: artist };
            }

            userCols.forEach(u => {
                const rank = parseFloat(row[u.index]);
                if (!isNaN(rank)) {
                    songs[songName].ranks.push({ user: u.name, rank: rank });

                    if (!userRanks[u.name]) userRanks[u.name] = [];
                    userRanks[u.name].push(rank);
                }
            });
        });
    });

    // Calculate stats for songs
    Object.keys(songs).forEach(songName => {
        const ranks = songs[songName].ranks.map(r => r.rank);
        const avg = ranks.reduce((a, b) => a + b, 0) / ranks.length;
        const variance = ranks.reduce((sum, r) => sum + Math.pow(r - avg, 2), 0) / ranks.length;
        const stdDev = Math.sqrt(variance);

        songs[songName].avgRank = avg;
        songs[songName].stdDev = stdDev;
        songs[songName].min = Math.min(...ranks);
        songs[songName].max = Math.max(...ranks);
        songs[songName].count = ranks.length;
    });

    return { songs, userRanks };
}

// ============================================
// FEATURE 1: MOST CONTROVERSIAL SONGS
// ============================================
function getMostControversialSongs(allData) {
    const topN = 20;

    const controversial = Object.keys(allData.songs)
        .map(songName => ({
            song: songName,
            stdDev: allData.songs[songName].stdDev,
            avgRank: allData.songs[songName].avgRank,
            min: allData.songs[songName].min,
            max: allData.songs[songName].max
        }))
        .sort((a, b) => b.stdDev - a.stdDev)
        .slice(0, topN);

    const rows = controversial.map((item, idx) => [
        idx + 1,
        item.song,
        item.stdDev.toFixed(2),
        item.avgRank.toFixed(1),
        `${item.min}-${item.max}`
    ]);

    const rowBgColors = rows.map(row => {
        const disagreement = parseFloat(row[2]);
        if (disagreement > 30) return '#ffcccc';
        else if (disagreement > 20) return '#ffdddd';
        return '#ffffff';
    });

    return {
        title: 'MOST CONTROVERSIAL SONGS (highest std deviation songs (disagreement))',
        titleColor: '#c00000',
        headers: [['Rank', 'Song', 'Disagreement', 'Avg Rank', 'Range (Min-Max)']],
        headerBgColor: '#ffe6e6',
        rows: rows,
        rowBgColors: rowBgColors
    };
}

// ============================================
// FEATURE 2: MOST CONSISTENTLY RANKED
// ============================================
function getMostConsistentlyRanked(allData) {
    const consistent = Object.keys(allData.songs)
        .map(songName => ({
            song: songName,
            stdDev: allData.songs[songName].stdDev,
            avgRank: allData.songs[songName].avgRank,
            count: allData.songs[songName].count
        }))
        .sort((a, b) => a.stdDev - b.stdDev)
        .slice(0, 20);

    const rows = consistent.map((item, idx) => [
        idx + 1,
        item.song,
        (100 - Math.min(item.stdDev * 5, 100)).toFixed(0) + '%',
        item.avgRank.toFixed(1)
    ]);

    return {
        title: 'MOST CONSISTENTLY RANKED (lowest std deviation songs)',
        titleColor: '#0066cc',
        headers: [['Rank', 'Song', 'Consistency', 'Avg Rank']],
        headerBgColor: '#cce5ff',
        rows: rows,
        rowBgColors: rows.map(() => '#e6f2ff')
    };
}

// ============================================
// FEATURE 3: HOT TAKES PER USER
// ============================================
function getHotTakesPerUser(allData) {
    const hotTakes = [];

    Object.keys(allData.songs).forEach(songName => {
        const song = allData.songs[songName];
        const avgRank = song.avgRank;

        song.ranks.forEach(rankObj => {
            const deviation = rankObj.rank - avgRank;
            if (Math.abs(deviation) > 25) {
                hotTakes.push({
                    song: songName,
                    user: rankObj.user,
                    theirRank: rankObj.rank,
                    groupAvg: avgRank,
                    deviation: deviation,
                    type: deviation > 0 ? 'Overrated!' : 'Underrated!'
                });
            }
        });
    });

    hotTakes.sort((a, b) => Math.abs(b.deviation) - Math.abs(a.deviation));
    const topTakes = hotTakes.slice(0, 20);

    const rows = topTakes.map((t, idx) => [
        idx + 1,
        t.song,
        t.user + ' (' + t.type + ')',
        t.theirRank.toFixed(0),
        t.groupAvg.toFixed(1),
        t.deviation.toFixed(1)
    ]);

    const rowBgColors = rows.map(row => {
        const dev = Math.abs(parseFloat(row[5]));
        if (dev > 50) return '#ff9999';
        else if (dev > 30) return '#ffcccc';
        return '#ffffff';
    });

    return {
        title: 'HOTTEST TAKES (regurgitation of takes tab)',
        titleColor: '#ff0000',
        headers: [['Rank', 'Song', 'Hot Take Artist', 'Their Rank', 'Group Avg', 'Deviation']],
        headerBgColor: '#ffcccc',
        rows: rows,
        rowBgColors: rowBgColors
    };
}

// ============================================
// FEATURE 4: MOST DISPUTED SONGS
// ============================================
function getSongDisputeAnalysis(configs, ss) {
    const SHOW_TOP_SONGS = 20;
    let songDisputes = {};

    configs.forEach(config => {
        const configSheet = ss.getSheetByName(config.tabName);
        if (!configSheet) return;

        const data = configSheet.getDataRange().getValues();
        if (data.length < 2) return;

        const headers = data[0];
        const rows = data.slice(1);
        const NON_USER_COLUMNS = ['Rank', 'Song', 'Points', 'Average'];

        const userCols = [];
        headers.forEach((h, idx) => {
            if (!NON_USER_COLUMNS.includes(h) && h !== "") {
                const hasData = rows.some(r => r[idx] !== "" && !isNaN(parseFloat(r[idx])));
                if (hasData) userCols.push({ name: h, index: idx });
            }
        });

        if (userCols.length < 2) return;

        rows.forEach(row => {
            const songName = String(row[1]);
            let maxDisagreement = 0;
            let maxPair = "";
            let avgDisagreement = 0;
            let pairCount = 0;

            for (let i = 0; i < userCols.length; i++) {
                for (let j = i + 1; j < userCols.length; j++) {
                    const rank1 = parseFloat(row[userCols[i].index]);
                    const rank2 = parseFloat(row[userCols[j].index]);

                    if (isNaN(rank1) || isNaN(rank2)) continue;

                    const disagreement = Math.abs(rank1 - rank2);
                    avgDisagreement += disagreement;
                    pairCount++;

                    if (disagreement > maxDisagreement) {
                        maxDisagreement = disagreement;
                        maxPair = `${userCols[i].name} vs ${userCols[j].name}`;
                    }
                }
            }

            if (pairCount > 0) {
                avgDisagreement /= pairCount;

                if (!songDisputes[songName]) {
                    songDisputes[songName] = {
                        maxPair: maxPair,
                        maxDisagreement: maxDisagreement,
                        avgDisagreement: avgDisagreement
                    };
                }
            }
        });
    });

    const topDisputes = Object.keys(songDisputes)
        .map(songName => ({
            song: songName,
            ...songDisputes[songName]
        }))
        .sort((a, b) => b.maxDisagreement - a.maxDisagreement)
        .slice(0, SHOW_TOP_SONGS);

    if (topDisputes.length === 0) return null;

    const rows = topDisputes.map((item, idx) => [
        idx + 1,
        item.song,
        item.maxPair,
        item.maxDisagreement,
        item.avgDisagreement.toFixed(1)
    ]);

    const rowBgColors = rows.map(row => {
        const maxDiff = row[3];
        if (maxDiff > 100) return '#ff9999';
        else if (maxDiff > 50) return '#ffcc99';
        else if (maxDiff > 20) return '#ffff99';
        return '#ffffff';
    });

    return {
        title: 'MOST DISPUTED SONGS (1v1 fight to the death)',
        titleColor: '#cc6600',
        headers: [['Rank', 'Song', 'Biggest Fight', 'Max Diff', 'Avg Diff']],
        headerBgColor: '#fff2cc',
        rows: rows,
        rowBgColors: rowBgColors
    };
}

// ============================================
// FEATURE 5: UNIVERSALLY TOP/BOTTOM 10
// ============================================
function getUniversallyTopBottom(allData) {
    const songStats = Object.keys(allData.songs).map(songName => {
        const song = allData.songs[songName];

        // Calculate Mean Reciprocal Rank with discount function
        // 1/(1+rank) gives smoother weighting - rank 1 = 0.5, rank 2 = 0.33, etc
        const mrr = song.ranks.reduce((sum, r) => sum + (1 / (1 + r.rank)), 0) / song.ranks.length;

        return {
            song: songName,
            avgRank: song.avgRank,
            stdDev: song.stdDev,
            mrr: mrr
        };
    });

    const topSongs = songStats.sort((a, b) => a.avgRank - b.avgRank).slice(0, 10);
    const bottomSongs = songStats.sort((a, b) => b.avgRank - a.avgRank).slice(0, 10);

    const topRows = topSongs.map((s, idx) => [
        idx + 1,
        s.song,
        s.avgRank.toFixed(1),
        s.mrr.toFixed(2),
        (100 - Math.min(s.stdDev * 2, 100)).toFixed(0) + '%'
    ]);

    const bottomRows = bottomSongs.map((s, idx) => [
        idx + 1,
        s.song,
        s.avgRank.toFixed(1),
        s.mrr.toFixed(2),
        (100 - Math.min(s.stdDev * 2, 100)).toFixed(0) + '%'
    ]);

    const allRows = topRows.concat(bottomRows);
    const bgColors = topRows.map(() => '#ccffcc').concat(bottomRows.map(() => '#ffcccc'));

    return {
        title: 'UNIVERSALLY TOP/BOTTOM 10',
        description: 'Best and worst songs with strong consensus',
        titleColor: '#1a73e8',
        headers: [['Rank', 'Song', 'Avg Rank', 'MRR', 'Agreement']],
        headerBgColor: '#d9e8f5',
        rows: allRows,
        rowBgColors: bgColors
    };
}

// ============================================
// FEATURE 6: OUTLIER USERS
// ============================================
function getOutlierUsers(allData, configs, ss) {
    const NON_USER_COLUMNS = ['Rank', 'Song', 'Points', 'Average'];
    const userDistances = {};

    configs.forEach(config => {
        const sheet = ss.getSheetByName(config.tabName);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;

        const headers = data[0];
        const rows = data.slice(1);

        const userCols = [];
        headers.forEach((h, idx) => {
            if (!NON_USER_COLUMNS.includes(h) && h !== "") {
                const hasData = rows.some(r => r[idx] !== "" && !isNaN(parseFloat(r[idx])));
                if (hasData) userCols.push({ name: h, index: idx });
            }
        });

        userCols.forEach(u1 => {
            if (!userDistances[u1.name]) userDistances[u1.name] = { total: 0, count: 0 };

            userCols.forEach(u2 => {
                if (u1.name === u2.name) return;

                let distance = 0;
                let comparisons = 0;

                rows.forEach(row => {
                    const rank1 = parseFloat(row[u1.index]);
                    const rank2 = parseFloat(row[u2.index]);

                    if (!isNaN(rank1) && !isNaN(rank2)) {
                        distance += Math.abs(rank1 - rank2);
                        comparisons++;
                    }
                });

                if (comparisons > 0) {
                    userDistances[u1.name].total += distance / comparisons;
                    userDistances[u1.name].count++;
                }
            });
        });
    });

    const outliers = Object.keys(userDistances)
        .map(userName => ({
            user: userName,
            avgDistance: userDistances[userName].count > 0 ? userDistances[userName].total / userDistances[userName].count : 0
        }))
        .sort((a, b) => b.avgDistance - a.avgDistance);

    const rows = outliers.map((u, idx) => [
        idx + 1,
        u.user,
        u.avgDistance.toFixed(1)
    ]);

    return {
        title: 'OUTLIER RANKING (whos the spiciest)',
        description: 'Users with most unique/different taste from others',
        titleColor: '#d33527',
        headers: [['Rank', 'User', 'Avg Distance']],
        headerBgColor: '#f4cccc',
        rows: rows,
        rowBgColors: rows.map(row => {
            const dist = parseFloat(row[2]);
            if (dist > 25) return '#ff9999';
            else if (dist > 15) return '#ffcc99';
            return '#ccffcc';
        })
    };
}

// ============================================
// FEATURE 7: COMEBACK/SLEEPER SONGS
// ============================================
function getComebackSongs(allData) {
    const sleepers = [];

    Object.keys(allData.songs).forEach(songName => {
        const song = allData.songs[songName];
        const avgRank = song.avgRank;

        let maxLover = null;
        let minRank = Infinity;

        song.ranks.forEach(rankObj => {
            if (rankObj.rank < minRank) {
                minRank = rankObj.rank;
                maxLover = rankObj.user;
            }
        });

        if (maxLover && minRank < 30 && avgRank > 60) {
            sleepers.push({
                song: songName,
                avgRank: avgRank,
                maxLover: maxLover,
                theirRank: minRank,
                gap: (avgRank - minRank).toFixed(0)
            });
        }
    });

    sleepers.sort((a, b) => b.gap - a.gap);
    const topSleepers = sleepers.slice(0, 20);

    const rows = topSleepers.map((s, idx) => [
        idx + 1,
        s.song,
        s.avgRank.toFixed(1),
        s.maxLover,
        s.theirRank,
        s.gap
    ]);

    return {
        title: 'SLEEPER SONGS (songs loved by at least one user but not by many)',
        description: 'Underrated gems loved by at least one user',
        titleColor: '#27ae60',
        headers: [['Rank', 'Song', 'Avg Rank', 'Lover', 'Their Rank', 'Gap']],
        headerBgColor: '#d5f4e6',
        rows: rows,
        rowBgColors: rows.map(row => {
            const gap = parseInt(row[5]);
            if (gap > 80) return '#66ff66';
            else if (gap > 50) return '#99ff99';
            return '#ccffcc';
        })
    };
}

// ============================================
// FEATURE 8: SUBUNIT POPULARITY
// ============================================
function getSubunitPopularity(allData) {
    const subunitMap = {
        'ID:160': 'CatChu!',
        'ID:161': 'KALEIDOSCORE',
        'ID:162': '5yncri5e!'
    };

    const subunits = {
        'CatChu!': { ranks: [], count: 0 },
        'KALEIDOSCORE': { ranks: [], count: 0 },
        '5yncri5e!': { ranks: [], count: 0 }
    };

    Object.keys(allData.songs).forEach(songName => {
        const song = allData.songs[songName];
        const artist = song.artist || "";

        // Check which subunit this song belongs to
        Object.keys(subunitMap).forEach(idKey => {
            if (artist.includes(idKey)) {
                const subunitName = subunitMap[idKey];
                subunits[subunitName].ranks = subunits[subunitName].ranks.concat(
                    song.ranks.map(r => r.rank)
                );
                subunits[subunitName].count++;
            }
        });
    });

    const results = Object.keys(subunits)
        .map(subunit => {
            const data = subunits[subunit];
            const ranks = data.ranks;

            if (ranks.length === 0) return null;

            const avgRank = ranks.reduce((a, b) => a + b, 0) / ranks.length;
            const variance = ranks.reduce((sum, r) => sum + Math.pow(r - avgRank, 2), 0) / ranks.length;
            const stdDev = Math.sqrt(variance);

            return {
                subunit: subunit,
                avgRank: avgRank,
                stdDev: stdDev
            };
        })
        .filter(r => r !== null)
        .sort((a, b) => a.avgRank - b.avgRank);

    const rows = results.map((r, idx) => [
        idx + 1,
        r.subunit,
        r.avgRank.toFixed(1),
        r.stdDev.toFixed(2)
    ]);

    return {
        title: 'SUBUNIT POPULARITY (uh oh)',
        description: 'Ranking performance of KALEIDOSCORE, 5yncri5e!, and CatChu!',
        titleColor: '#e67e22',
        headers: [['Rank', 'Subunit', 'Avg Rank', 'Deviation']],
        headerBgColor: '#fdebd0',
        rows: rows,
        rowBgColors: rows.map((row, idx) => {
            if (idx === 0) return '#66ff66';
            else if (idx === rows.length - 1) return '#ff9999';
            return '#ffff99';
        })
    };
}



// ============================================
// MAIN: SPICE ANALYSIS (DEDICATED TAB)
// ============================================
function runSpiceAnalysis() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const SHEET_NAME = 'Spice Index';
    const configs = getTargetSheetConfigs();
    const allData = collectAllSongData(configs, ss);

    if (Object.keys(allData.songs).length === 0) return;

    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    sheet.clear();

    const spiceData = getSpiceMeterAnalysis(allData, configs, ss);

    // Header styling
    sheet.getRange(1, 1).setValue(spiceData.title).setFontWeight('bold').setFontSize(16).setFontColor(spiceData.titleColor);
    sheet.getRange(2, 1).setValue(spiceData.description).setFontStyle('italic');

    const headers = spiceData.headers[0];
    sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setBackground(spiceData.headerBgColor).setFontWeight('bold');

    if (spiceData.rows.length > 0) {
        sheet.getRange(5, 1, spiceData.rows.length, spiceData.rows[0].length).setValues(spiceData.rows);

        // Apply row background colors
        spiceData.rowBgColors.forEach((color, idx) => {
            sheet.getRange(5 + idx, 1, 1, headers.length).setBackground(color);
        });

        // Formatting
        sheet.setColumnWidth(1, 150); // User
        sheet.setColumnWidth(2, 120); // Global
        for (let i = 2; i < headers.length; i++) {
            sheet.setColumnWidth(i + 1, 120);
        }

        // Bold the Global column
        sheet.getRange(4, 2, spiceData.rows.length + 1, 1).setFontWeight('bold');

        // Final borders
        sheet.getRange(4, 1, spiceData.rows.length + 1, headers.length).setBorder(true, true, true, true, true, true);
    }

    ss.toast("Spice Index updated!");
}

// ============================================
// FEATURE 9: THE SPICE METER
// ============================================
function getSpiceMeterAnalysis(allData, configs, ss) {
    const userRMSPerGroup = {}; // { user: { groupName: sqDiffs[], overallSqDiffs: [] } }
    const groupNames = configs.map(c => c.tabName);
    const allUniqueUsers = Object.keys(allData.userRanks);

    allUniqueUsers.forEach(user => {
        userRMSPerGroup[user] = { overallSqDiffs: [] };
        groupNames.forEach(gn => userRMSPerGroup[user][gn] = []);
    });

    const SYS_COLS = ['Rank', 'Song', 'Points', 'Average'];

    configs.forEach(config => {
        const groupName = config.tabName;
        const sheet = ss.getSheetByName(groupName);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;

        const headers = data[0];
        const rows = data.slice(1);

        const userColsInGroup = [];
        headers.forEach((h, idx) => {
            const cleanH = String(h || "").trim();
            if (cleanH !== "" && !SYS_COLS.some(s => s.toLowerCase() === cleanH.toLowerCase())) {
                const hasData = rows.some(r => r[idx] !== "" && !isNaN(parseFloat(r[idx])));
                if (hasData) userColsInGroup.push({ name: cleanH, index: idx });
            }
        });

        rows.forEach(row => {
            const userRanksForSong = userColsInGroup.map(u => ({
                user: u.name,
                rank: parseFloat(row[u.index])
            })).filter(r => !isNaN(r.rank));

            if (userRanksForSong.length > 1) {
                userRanksForSong.forEach(userRankObj => {
                    const otherRanks = userRanksForSong.filter(r => r.user !== userRankObj.user).map(r => r.rank);
                    const avgOther = otherRanks.reduce((a, b) => a + b, 0) / otherRanks.length;
                    const sqDiff = Math.pow(userRankObj.rank - avgOther, 2);

                    if (userRMSPerGroup[userRankObj.user]) {
                        userRMSPerGroup[userRankObj.user][groupName].push(sqDiff);
                        userRMSPerGroup[userRankObj.user].overallSqDiffs.push(sqDiff);
                    }
                });
            }
        });
    });

    const rows = allUniqueUsers.map(user => {
        const rowData = [user];
        let overallRMS = 0;

        if (userRMSPerGroup[user].overallSqDiffs.length > 0) {
            overallRMS = Math.sqrt(userRMSPerGroup[user].overallSqDiffs.reduce((a, b) => a + b, 0) / userRMSPerGroup[user].overallSqDiffs.length);
        }
        rowData.push(overallRMS.toFixed(1));

        groupNames.forEach(gn => {
            const diffs = userRMSPerGroup[user][gn];
            if (diffs && diffs.length > 0) {
                const rms = Math.sqrt(diffs.reduce((a, b) => a + b, 0) / diffs.length);
                rowData.push(rms.toFixed(1));
            } else {
                rowData.push("-");
            }
        });

        return rowData;
    });

    // Sort by overall RMS (Spiciest at top)
    rows.sort((a, b) => parseFloat(b[1]) - parseFloat(a[1]));

    const resultHeaders = [['User', 'Global Spice', ...groupNames]];

    return {
        title: 'THE SPICE METER (Group Breakdown)',
        description: 'Root Mean Squared deviation from others. Higher = More Unique.',
        titleColor: '#e67e22',
        headers: resultHeaders,
        headerBgColor: '#fdebd0',
        rows: rows,
        rowBgColors: rows.map(row => {
            const overall = parseFloat(row[1]);
            if (overall > 35) return '#ffcccc'; // Spicy
            if (overall < 15) return '#cceeff'; // Basic
            return '#ffffff';
        })
    };
}


/**
 * Helper to write a small batch of takes horizontally.
 */
function writeTakesHorizontalBatch(sheet, startRow, startCol, takes) {
    const header = [['%', 'User', 'Song', 'Rank', 'Avg']];
    sheet.getRange(startRow, startCol, 1, 5).setValues(header).setBackground('#efefef').setFontWeight('bold');

    if (takes.length === 0) {
        sheet.getRange(startRow + 1, startCol).setValue("(No takes)");
        return;
    }

    const rows = takes.map(t => [
        t.score.toFixed(1),
        t.user,
        t.song,
        t.userRank,
        t.avgRank
    ]);

    sheet.getRange(startRow + 1, startCol, rows.length, 5).setValues(rows);
    sheet.getRange(startRow + 1, startCol, rows.length, 1).setHorizontalAlignment("center");
}

function updateCount(obj, user, type) {
    if (!obj[user]) obj[user] = { most: 0, least: 0 };
    obj[user][type]++;
}