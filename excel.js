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
    buildSummarySheet(summary);

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

function buildSummarySheet(sheet) {
    // Clear any existing content
    const rowCount = sheet.rowCount;
    for (let i = rowCount; i >= 1; i--) {
        sheet.spliceRows(i, 1);
    }

    sheet.getCell('A1').value = 'PICKS TRACKER - SUMMARY';
    sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F4E79' } };

    // Headers: Row 2
    const headers = ['', 'Abuelo', 'Cristian Rey', 'Roberto Rey', 'TOTAL'];
    for (let i = 0; i < headers.length; i++) {
        const cell = sheet.getCell(2, i + 1);
        cell.value = headers[i];
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        cell.alignment = { horizontal: 'center' };
    }

    // Row 3: Total Picks - Count rows with results (any value in Result column)
    sheet.getCell(3, 1).value = 'Total Picks';
    sheet.getCell(3, 1).font = { bold: true };
    
    sheet.getCell(3, 2).value = { formula: `COUNTA(Abuelo!A2:A1000)-COUNTA(Abuelo!G2:G1000)` };  // Count rows minus blanks in Result col
    sheet.getCell(3, 3).value = { formula: `COUNTA('Cristian Rey'!A2:A1000)-COUNTA('Cristian Rey'!G2:G1000)` };
    sheet.getCell(3, 4).value = { formula: `COUNTA('Roberto Rey'!A2:A1000)-COUNTA('Roberto Rey'!G2:G1000)` };
    sheet.getCell(3, 5).value = { formula: `B3+C3+D3` };

    // Row 4: Wins - Count cells in Result column that equal "W"
    sheet.getCell(4, 1).value = 'Wins';
    sheet.getCell(4, 1).font = { bold: true };
    
    sheet.getCell(4, 2).value = { formula: `COUNTIF(Abuelo!G:G,"W")` };
    sheet.getCell(4, 3).value = { formula: `COUNTIF('Cristian Rey'!G:G,"W")` };
    sheet.getCell(4, 4).value = { formula: `COUNTIF('Roberto Rey'!G:G,"W")` };
    sheet.getCell(4, 5).value = { formula: `B4+C4+D4` };

    // Row 5: Losses - Count cells in Result column that equal "L"
    sheet.getCell(5, 1).value = 'Losses';
    sheet.getCell(5, 1).font = { bold: true };
    
    sheet.getCell(5, 2).value = { formula: `COUNTIF(Abuelo!G:G,"L")` };
    sheet.getCell(5, 3).value = { formula: `COUNTIF('Cristian Rey'!G:G,"L")` };
    sheet.getCell(5, 4).value = { formula: `COUNTIF('Roberto Rey'!G:G,"L")` };
    sheet.getCell(5, 5).value = { formula: `B5+C5+D5` };

    // Row 6: Pushes - Count cells in Result column that equal "P"
    sheet.getCell(6, 1).value = 'Pushes';
    sheet.getCell(6, 1).font = { bold: true };
    
    sheet.getCell(6, 2).value = { formula: `COUNTIF(Abuelo!G:G,"P")` };
    sheet.getCell(6, 3).value = { formula: `COUNTIF('Cristian Rey'!G:G,"P")` };
    sheet.getCell(6, 4).value = { formula: `COUNTIF('Roberto Rey'!G:G,"P")` };
    sheet.getCell(6, 5).value = { formula: `B6+C6+D6` };

    // Row 7: Win Rate % = Wins / Total Picks (excluding pushes)
    sheet.getCell(7, 1).value = 'Win Rate %';
    sheet.getCell(7, 1).font = { bold: true };
    
    sheet.getCell(7, 2).value = { formula: `IF((B4+B5)=0,0,B4/(B4+B5))` };
    sheet.getCell(7, 3).value = { formula: `IF((C4+C5)=0,0,C4/(C4+C5))` };
    sheet.getCell(7, 4).value = { formula: `IF((D4+D5)=0,0,D4/(D4+D5))` };
    sheet.getCell(7, 5).value = { formula: `IF((B4+B5+C4+C5+D4+D5)=0,0,(B4+C4+D4)/(B4+B5+C4+C5+D4+D5))` };

    // Format Win Rate as percentage
    for (let c = 2; c <= 5; c++) {
        sheet.getCell(7, c).numFmt = '0.0%';
    }

    // Row 8: Total Profit/Loss - Sum of Profit/Loss column (H column)
    sheet.getCell(8, 1).value = 'Total Profit/Loss';
    sheet.getCell(8, 1).font = { bold: true };
    
    sheet.getCell(8, 2).value = { formula: `SUMIF(Abuelo!G:G,"W",Abuelo!H:H)+SUMIF(Abuelo!G:G,"L",Abuelo!H:H)` };
    sheet.getCell(8, 3).value = { formula: `SUMIF('Cristian Rey'!G:G,"W",'Cristian Rey'!H:H)+SUMIF('Cristian Rey'!G:G,"L",'Cristian Rey'!H:H)` };
    sheet.getCell(8, 4).value = { formula: `SUMIF('Roberto Rey'!G:G,"W",'Roberto Rey'!H:H)+SUMIF('Roberto Rey'!G:G,"L",'Roberto Rey'!H:H)` };
    sheet.getCell(8, 5).value = { formula: `B8+C8+D8` };

    // Format as currency
    for (let c = 2; c <= 5; c++) {
        sheet.getCell(8, c).numFmt = '$#,##0.00';
    }

    // Row 9: Yield % = Total Profit / (Total Picks * Bet Amount) = Total Profit / Total Staked
    sheet.getCell(9, 1).value = 'Yield %';
    sheet.getCell(9, 1).font = { bold: true };
    
    // For each tipster: Profit / (Total Picks * 2000)
    sheet.getCell(9, 2).value = { formula: `IF(B3=0,0,B8/(B3*2000))` };
    sheet.getCell(9, 3).value = { formula: `IF(C3=0,0,C8/(C3*2000))` };
    sheet.getCell(9, 4).value = { formula: `IF(D3=0,0,D8/(D3*2000))` };
    sheet.getCell(9, 5).value = { formula: `IF(E3=0,0,E8/(E3*2000))` };

    // Format Yield as percentage
    for (let c = 2; c <= 5; c++) {
        sheet.getCell(9, c).numFmt = '0.0%';
    }

    // Set column widths
    sheet.getColumn(1).width = 20;
    sheet.getColumn(2).width = 15;
    sheet.getColumn(3).width = 15;
    sheet.getColumn(4).width = 15;
    sheet.getColumn(5).width = 15;

    // Center align all numeric cells
    for (let r = 3; r <= 9; r++) {
        for (let c = 2; c <= 5; c++) {
            sheet.getCell(r, c).alignment = { horizontal: 'right' };
        }
    }
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

    // Rebuild Summary sheet with live formulas
    const summarySheet = workbook.getWorksheet('Summary');
    buildSummarySheet(summarySheet);

    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log('   [Summary] Updated with live formulas');
}

// Update a pick's result (Win/Loss/Push) after game finishes
async function updatePickResult(pick, result) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const tipsterSheet = workbook.getWorksheet(pick.tipster);
    const generalSheet = workbook.getWorksheet('General');

    if (!tipsterSheet) {
        console.error(`[Excel] Sheet not found for tipster: ${pick.tipster}`);
        return;
    }

    // Map result to single letter
    const resultLetter = result === 'Win' ? 'W' : result === 'Loss' ? 'L' : 'P';
    const odds = parseFloat(pick.odds) || 0;
    let profitLoss = 0;

    if (resultLetter === 'W') {
        profitLoss = (odds - 1) * BET_AMOUNT;
    } else if (resultLetter === 'L') {
        profitLoss = -BET_AMOUNT;
    }
    // Push = 0 profit/loss

    // Find and update in tipster sheet
    let found = false;
    tipsterSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;

        const rowDate = row.getCell(1).value;
        const rowMatch = row.getCell(3).value;
        
        // Match by date and match name
        if (rowDate === pick.date && rowMatch === pick.match) {
            row.getCell(7).value = resultLetter;      // Result column
            row.getCell(8).value = profitLoss;        // Profit/Loss column
            found = true;
        }
    });

    if (found) {
        console.log(`[Excel] Updated result for ${pick.tipster}: ${pick.match} = ${resultLetter}`);
    } else {
        console.log(`[Excel] Warning: Could not find pick to update: ${pick.date} ${pick.match}`);
    }

    // Also update General sheet
    generalSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;

        const rowDate = row.getCell(1).value;
        const rowTipster = row.getCell(2).value;
        const rowMatch = row.getCell(4).value;

        if (rowDate === pick.date && rowTipster === pick.tipster && rowMatch === pick.match) {
            row.getCell(8).value = resultLetter;      // Result column
            row.getCell(9).value = profitLoss;        // Profit/Loss column
        }
    });

    await workbook.xlsx.writeFile(EXCEL_FILE);
}

module.exports = { initExcel, addPick, updateSummary, updatePickResult };