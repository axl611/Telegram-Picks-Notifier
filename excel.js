// excel.js - Excel manager for picks tracking
const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE = path.join(process.cwd(), 'picks.xlsx');
const BET_AMOUNT = 2000;

const TIPSTERS = ['Abuelo', 'Cristian Rey', 'Roberto Rey'];

// Current schema version - increment when structure changes
const SCHEMA_VERSION = 3;

// Expected columns for tipster sheets (v3 - no Time, no Type, combined Pick)
const TIPSTER_COLUMNS = ['Date', 'Sport', 'Match', 'Pick', 'Odds', 'Bet', 'Result', 'Profit/Loss', 'Balance'];

// Expected columns for General sheet
const GENERAL_COLUMNS = ['Date', 'Tipster', 'Sport', 'Match', 'Pick', 'Odds', 'Bet', 'Result', 'Profit/Loss'];

async function initExcel() {
    const workbook = new ExcelJS.Workbook();

    // Create tipster sheets
    for (const tipster of TIPSTERS) {
        const sheet = workbook.addWorksheet(tipster);
        sheet.addRow(TIPSTER_COLUMNS);
        styleHeader(sheet.getRow(1));
        setColumnWidths(sheet, [12, 10, 35, 30, 8, 10, 10, 14, 14]);
    }

    // Create General sheet
    const general = workbook.addWorksheet('General');
    general.addRow(GENERAL_COLUMNS);
    styleHeader(general.getRow(1));
    setColumnWidths(general, [12, 15, 10, 35, 30, 8, 10, 10, 14]);

    // Create Summary sheet
    const summary = workbook.addWorksheet('Summary');
    buildSummarySheet(summary, {});

    // Add schema version marker (hidden)
    const infoSheet = workbook.addWorksheet('_Info');
    infoSheet.getCell('A1').value = 'SchemaVersion';
    infoSheet.getCell('B1').value = SCHEMA_VERSION;
    infoSheet.state = 'hidden';

    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log('Excel file initialized: ' + EXCEL_FILE);
}

function styleHeader(row) {
    row.eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
}

function setColumnWidths(sheet, widths) {
    sheet.columns = widths.map((w, i) => ({ width: w }));
}

function buildSummarySheet(sheet, stats) {
    sheet._merges = {};
    
    sheet.getCell('A1').value = 'PICKS TRACKER - SUMMARY';
    sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F4E79' } };

    const headers = ['', 'Abuelo', 'Cristian Rey', 'Roberto Rey', 'TOTAL'];
    for (let i = 0; i < headers.length; i++) {
        const cell = sheet.getCell(2, i + 1);
        cell.value = headers[i];
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        cell.alignment = { horizontal: 'center' };
    }

    const dataRows = [
        ['Total Picks', 0, 0, 0, 0],
        ['Wins', 0, 0, 0, 0],
        ['Losses', 0, 0, 0, 0],
        ['Pushes', 0, 0, 0, 0],
        ['Win Rate %', '0%', '0%', '0%', '0%'],
        ['Total Profit/Loss', '$0', '$0', '$0', '$0'],
        ['Yield %', '0%', '0%', '0%', '0%'],
    ];

    if (stats && Object.keys(stats).length > 0) {
        const tipsterList = TIPSTERS;
        const totalWins = tipsterList.reduce((s, t) => s + (stats[t] ? stats[t].wins : 0), 0);
        const totalLosses = tipsterList.reduce((s, t) => s + (stats[t] ? stats[t].losses : 0), 0);
        const totalPushes = tipsterList.reduce((s, t) => s + (stats[t] ? stats[t].pushes : 0), 0);
        const totalPicks = tipsterList.reduce((s, t) => s + (stats[t] ? stats[t].totalPicks : 0), 0);
        const totalProfit = tipsterList.reduce((s, t) => s + (stats[t] ? stats[t].totalProfit : 0), 0);
        const totalWinRate = totalPicks > 0 ? ((totalWins / totalPicks) * 100).toFixed(1) + '%' : '0%';
        const totalYield = totalPicks > 0 ? ((totalProfit / (totalPicks * BET_AMOUNT)) * 100).toFixed(1) + '%' : '0%';

        dataRows[0] = ['Total Picks', ...tipsterList.map(t => stats[t] ? stats[t].totalPicks : 0), totalPicks];
        dataRows[1] = ['Wins', ...tipsterList.map(t => stats[t] ? stats[t].wins : 0), totalWins];
        dataRows[2] = ['Losses', ...tipsterList.map(t => stats[t] ? stats[t].losses : 0), totalLosses];
        dataRows[3] = ['Pushes', ...tipsterList.map(t => stats[t] ? stats[t].pushes : 0), totalPushes];
        dataRows[4] = ['Win Rate %', ...tipsterList.map(t => stats[t] ? stats[t].winRate : '0%'), totalWinRate];
        dataRows[5] = ['Total Profit/Loss', ...tipsterList.map(t => stats[t] ? '$' + stats[t].totalProfit.toFixed(2) : '$0'), '$' + totalProfit.toFixed(2)];
        dataRows[6] = ['Yield %', ...tipsterList.map(t => stats[t] ? stats[t].yield_ : '0%'), totalYield];
    }

    for (let r = 0; r < dataRows.length; r++) {
        for (let c = 0; c < dataRows[r].length; c++) {
            const cell = sheet.getCell(r + 3, c + 1);
            cell.value = dataRows[r][c];
            if (c === 0) cell.font = { bold: true };
        }
    }

    sheet.getColumn(1).width = 20;
    sheet.getColumn(2).width = 15;
    sheet.getColumn(3).width = 15;
    sheet.getColumn(4).width = 15;
    sheet.getColumn(5).width = 15;
}

// ── Check and migrate schema if needed ──
async function ensureSchema() {
    let workbook;
    
    try {
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);
    } catch (err) {
        // File doesn't exist or is corrupted - create new
        console.log('   Creating new Excel file...');
        await initExcel();
        return null;
    }

    // Check schema version
    let infoSheet = workbook.getWorksheet('_Info');
    let currentVersion = 0;
    
    if (infoSheet) {
        currentVersion = infoSheet.getCell('B1').value || 0;
    }

    if (currentVersion >= SCHEMA_VERSION) {
        // Schema is up to date
        return workbook;
    }

    console.log('   Migrating schema from v' + currentVersion + ' to v' + SCHEMA_VERSION + '...');
    
    // Migrate the workbook
    const migratedWorkbook = await migrateWorkbook(workbook, currentVersion);
    
    // Update schema version
    if (!migratedWorkbook.getWorksheet('_Info')) {
        infoSheet = migratedWorkbook.addWorksheet('_Info');
        infoSheet.state = 'hidden';
    }
    infoSheet.getCell('A1').value = 'SchemaVersion';
    infoSheet.getCell('B1').value = SCHEMA_VERSION;
    
    await migratedWorkbook.xlsx.writeFile(EXCEL_FILE);
    console.log('   Migration complete.');
    
    return migratedWorkbook;
}

// ── Migrate workbook to new schema ──
async function migrateWorkbook(workbook, fromVersion) {
    // Migrate each tipster sheet
    for (const tipster of TIPSTERS) {
        let sheet = workbook.getWorksheet(tipster);
        
        if (!sheet) {
            // Create missing sheet
            sheet = workbook.addWorksheet(tipster);
            sheet.addRow(TIPSTER_COLUMNS);
            styleHeader(sheet.getRow(1));
            setColumnWidths(sheet, [12, 10, 35, 30, 8, 10, 10, 14, 14]);
            continue;
        }

        // Get existing data
        const oldData = [];
        sheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return; // skip header
            oldData.push(row.values.slice(1));
        });

        // Get current headers
        const currentHeaders = sheet.getRow(1).values.slice(1);
        
        // Check if migration needed
        const needsMigration = currentHeaders.length !== TIPSTER_COLUMNS.length ||
            !TIPSTER_COLUMNS.every((h, i) => currentHeaders[i] === h);

        if (needsMigration) {
            console.log('   Migrating sheet: ' + tipster);
            
            // Map old columns to new
            const migratedData = oldData.map(row => {
                return migrateRow(row, currentHeaders, TIPSTER_COLUMNS);
            });

            // Clear and rebuild
            const rowCount = sheet.rowCount;
            for (let i = rowCount; i >= 1; i--) {
                sheet.spliceRows(i, 1);
            }

            sheet.addRow(TIPSTER_COLUMNS);
            styleHeader(sheet.getRow(1));
            
            migratedData.forEach(row => sheet.addRow(row));
            setColumnWidths(sheet, [12, 10, 35, 30, 8, 10, 10, 14, 14]);
        }
    }

    // Migrate General sheet
    let general = workbook.getWorksheet('General');
    if (!general) {
        general = workbook.addWorksheet('General');
        general.addRow(GENERAL_COLUMNS);
        styleHeader(general.getRow(1));
    } else {
        const currentHeaders = general.getRow(1).values.slice(1);
        const needsMigration = currentHeaders.length !== GENERAL_COLUMNS.length ||
            !GENERAL_COLUMNS.every((h, i) => currentHeaders[i] === h);

        if (needsMigration) {
            const oldData = [];
            general.eachRow((row, rowNum) => {
                if (rowNum === 1) return;
                oldData.push(row.values.slice(1));
            });

            const migratedData = oldData.map(row => {
                return migrateRow(row, currentHeaders, GENERAL_COLUMNS);
            });

            const rowCount = general.rowCount;
            for (let i = rowCount; i >= 1; i--) {
                general.spliceRows(i, 1);
            }

            general.addRow(GENERAL_COLUMNS);
            styleHeader(general.getRow(1));
            migratedData.forEach(row => general.addRow(row));
        }
    }
    setColumnWidths(general, [12, 15, 10, 35, 30, 8, 10, 10, 14]);

    // Ensure Summary sheet exists
    if (!workbook.getWorksheet('Summary')) {
        const summary = workbook.addWorksheet('Summary');
        buildSummarySheet(summary, {});
    }

    return workbook;
}

// ── Migrate a single row from old to new columns ──
function migrateRow(oldRow, oldHeaders, newHeaders) {
    const newRow = newHeaders.map(h => '');
    
    oldHeaders.forEach((oldHeader, i) => {
        if (oldHeader && oldRow[i] !== undefined) {
            const newIndex = newHeaders.indexOf(oldHeader);
            if (newIndex !== -1) {
                newRow[newIndex] = oldRow[i];
            }
        }
    });

    // Handle specific migrations
    // v1->v2: If we had 'Pick' and 'Type' separate, combine them
    const pickIndex = oldHeaders.indexOf('Pick');
    const typeIndex = oldHeaders.indexOf('Type');
    const newPickIndex = newHeaders.indexOf('Pick');
    
    if (pickIndex !== -1 && typeIndex !== -1 && newPickIndex !== -1) {
        const pickVal = oldRow[pickIndex] || '';
        const typeVal = oldRow[typeIndex] || '';
        
        if (pickVal && typeVal && typeVal !== 'ML') {
            // If they had separate values, try to combine intelligently
            if (pickVal.toLowerCase().includes(typeVal.toLowerCase())) {
                newRow[newPickIndex] = pickVal;
            } else {
                newRow[newPickIndex] = pickVal + ' (' + typeVal + ')';
            }
        } else {
            newRow[newPickIndex] = pickVal || typeVal || '';
        }
    }

    return newRow;
}

async function addPick(pickData) {
    if (pickData.match) {
        pickData.match = pickData.match.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
    }
    
    if (!pickData.match || pickData.match.length < 3) {
        console.log('   Skipping invalid pick - match too short');
        return;
    }
    
    if (pickData.match.length > 60) {
        pickData.match = pickData.match.substring(0, 55) + '...';
    }

    // Ensure schema is up to date
    let workbook = await ensureSchema();
    
    if (!workbook) {
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);
    }

    // Normalize tipster name
    let tipsterName = pickData.tipster || 'Unknown';
    const tipsterLower = tipsterName.toLowerCase();
    if (tipsterLower.includes('abuelo')) {
        tipsterName = 'Abuelo';
    } else if (tipsterLower.includes('cristian')) {
        tipsterName = 'Cristian Rey';
    } else if (tipsterLower.includes('roberto') || tipsterLower.includes('beto')) {
        tipsterName = 'Roberto Rey';
    }

    // Row data matching TIPSTER_COLUMNS
    const rowData = [
        pickData.date || '',
        pickData.sport || 'SOCCER',
        pickData.match || '',
        pickData.pick || '',
        pickData.odds || '',
        BET_AMOUNT,
        '',  // result
        '',  // profitloss
        ''   // balance
    ];

    // Add to tipster sheet
    const tipsterSheet = workbook.getWorksheet(tipsterName);
    if (tipsterSheet) {
        tipsterSheet.addRow(rowData);
    }

    // Add to General sheet
    const generalSheet = workbook.getWorksheet('General');
    if (generalSheet) {
        const generalRowData = [
            pickData.date || '',
            tipsterName,
            pickData.sport || 'SOCCER',
            pickData.match || '',
            pickData.pick || '',
            pickData.odds || '',
            BET_AMOUNT,
            '',  // result
            ''   // profitloss
        ];
        generalSheet.addRow(generalRowData);
    }

    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log('Pick saved: ' + tipsterName + ' | ' + pickData.match + ' | ' + pickData.pick + ' | Odds: ' + pickData.odds);
}

async function updateSummary() {
    // Ensure schema is up to date
    let workbook = await ensureSchema();
    
    if (!workbook) {
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);
    }

    const stats = {};

    for (const tipster of TIPSTERS) {
        const sheet = workbook.getWorksheet(tipster);
        if (!sheet) continue;

        let wins = 0, losses = 0, pushes = 0, totalProfit = 0, totalPicks = 0;

        sheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            // Columns: Date, Sport, Match, Pick, Odds, Bet, Result, Profit/Loss, Balance
            const result = (row.getCell(7).value || '').toString().toUpperCase();
            const odds = parseFloat(row.getCell(5).value) || 0;
            if (result === 'W') { wins++; totalProfit += (odds - 1) * BET_AMOUNT; totalPicks++; }
            else if (result === 'L') { losses++; totalProfit -= BET_AMOUNT; totalPicks++; }
            else if (result === 'P') { pushes++; }
        });

        const winRate = totalPicks > 0 ? ((wins / totalPicks) * 100).toFixed(1) + '%' : '0%';
        const totalBet = totalPicks * BET_AMOUNT;
        const yield_ = totalBet > 0 ? ((totalProfit / totalBet) * 100).toFixed(1) + '%' : '0%';

        stats[tipster] = { wins, losses, pushes, totalProfit, totalPicks, winRate, yield_ };
        console.log('   [Summary] ' + tipster + ': ' + totalPicks + ' picks, ' + wins + 'W-' + losses + 'L-' + pushes + 'P');
    }

    // Rebuild summary
    const summarySheet = workbook.getWorksheet('Summary');
    
    const rowCount = summarySheet.rowCount;
    for (let i = rowCount; i >= 1; i--) {
        summarySheet.spliceRows(i, 1);
    }
    
    buildSummarySheet(summarySheet, stats);

    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log('Summary updated.');
}

module.exports = { initExcel, addPick, updateSummary };