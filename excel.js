// excel.js - Excel manager for picks tracking
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const EXCEL_FILE = path.join(process.cwd(), 'picks.xlsx');
const CSV_FILE = path.join(process.cwd(), 'picks_backup.csv');
const BET_AMOUNT = 2000;

const TIPSTERS = ['Abuelo', 'Cristian Rey', 'Roberto Rey'];

// Current schema version - increment when structure changes
const SCHEMA_VERSION = 3;

// Expected columns for tipster sheets (v3 - no Time, no Type, combined Pick)
const TIPSTER_COLUMNS = ['Date', 'Sport', 'Match', 'Pick', 'Odds', 'Bet', 'Result', 'Profit/Loss', 'Balance'];

// Expected columns for General sheet
const GENERAL_COLUMNS = ['Date', 'Tipster', 'Sport', 'Match', 'Pick', 'Odds', 'Bet', 'Result', 'Profit/Loss'];

async function initExcel() {
    // Check if file already exists
    if (fs.existsSync(EXCEL_FILE)) {
        console.log('Excel file already exists: ' + EXCEL_FILE);
        // Use ensureSchema to load and verify the existing file
        await ensureSchema();
        return;
    }
    
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
        if (fs.existsSync(EXCEL_FILE)) {
            // File exists but couldn't be read (locked or temporarily unavailable) — never delete it
            console.error('   Could not read Excel file (may be locked or in use): ' + err.message);
            return null;
        }
        // File genuinely missing — create a fresh one
        console.log('   Excel file not found, creating new one...');
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
        const newRow = tipsterSheet.addRow(rowData);
        const rowNum = newRow.number;
        
        // Add formula to Profit/Loss column (H): IF(G="W",(F-1)*2000,IF(G="L",-2000,0))
        newRow.getCell(8).value = { formula: `IF(G${rowNum}="W",(F${rowNum}-1)*2000,IF(G${rowNum}="L",-2000,0))` };
        
        // Add formula to Balance column (I): running total
        if (rowNum === 2) {
            // First data row: Balance = Profit/Loss
            newRow.getCell(9).value = { formula: `H${rowNum}` };
        } else {
            // Subsequent rows: Balance = Previous Balance + This Profit/Loss
            newRow.getCell(9).value = { formula: `I${rowNum - 1}+H${rowNum}` };
        }
        
        // Format Profit/Loss column as currency
        newRow.getCell(8).numFmt = '$#,##0.00';
        newRow.getCell(9).numFmt = '$#,##0.00';
        
        // Add data validation to Result column (G) - dropdown with W, L, P
        tipsterSheet.dataValidations.add({
            type: 'list',
            formula1: '"W,L,P"',
            showDropDown: true,
            sqref: `G${rowNum}`
        });
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
        const newGeneralRow = generalSheet.addRow(generalRowData);
        const genRowNum = newGeneralRow.number;
        
        // Add formula to Profit/Loss column (I): IF(H="W",(G-1)*2000,IF(H="L",-2000,0))
        newGeneralRow.getCell(9).value = { formula: `IF(H${genRowNum}="W",(G${genRowNum}-1)*2000,IF(H${genRowNum}="L",-2000,0))` };
        newGeneralRow.getCell(9).numFmt = '$#,##0.00';
        
        // Add data validation to Result column (H) - dropdown with W, L, P
        generalSheet.dataValidations.add({
            type: 'list',
            formula1: '"W,L,P"',
            showDropDown: true,
            sqref: `H${genRowNum}`
        });
    }

    await workbook.xlsx.writeFile(EXCEL_FILE);

    // Append to CSV backup
    appendPickToCsv(tipsterName, pickData);

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

// Normalize date for comparison (Excel may return Date object or string)
function normalizeDateForCompare(val) {
    if (val == null || val === '') return '';
    if (val instanceof Date) {
        return val.getDate() + ' ' + val.toLocaleString('en', { month: 'short' });
    }
    return String(val).trim();
}

// Normalize match for comparison (Excel may truncate with "...")
function normalizeMatchForCompare(val) {
    if (val == null) return '';
    return String(val).replace(/\s+/g, ' ').trim();
}

function datesMatchExcel(a, b) {
    return normalizeDateForCompare(a) === normalizeDateForCompare(b);
}

function matchesMatchExcel(rowMatch, pickMatch) {
    const r = normalizeMatchForCompare(rowMatch);
    const p = normalizeMatchForCompare(pickMatch);
    if (r === p) return true;
    if (r.startsWith(p) || p.startsWith(r)) return true;
    return false;
}

// Update a pick's result (Win/Loss/Push) after game finishes
async function updatePickResult(pick, result) {
    console.log(`[Excel] updatePickResult called: tipster=${pick.tipster}, date=${pick.date}, match=${pick.match}, result=${result}`);
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

    // Find and update in tipster sheet (use normalized comparison)
    let found = false;
    tipsterSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;

        const rowDate = row.getCell(1).value;
        const rowMatch = row.getCell(3).value;

        if (datesMatchExcel(rowDate, pick.date) && matchesMatchExcel(rowMatch, pick.match)) {
            row.getCell(7).value = resultLetter;
            found = true;
            console.log(`[Excel] Updated tipster row ${rowNum}: ${resultLetter}`);
        }
    });

    if (found) {
        console.log(`[Excel] Updated result for ${pick.tipster}: ${pick.match} = ${resultLetter}`);
    } else {
        console.log(`[Excel] Warning: Could not find pick to update: date="${pick.date}" match="${pick.match}"`);
    }

    // Also update General sheet
    if (generalSheet) {
        generalSheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;

            const rowDate = row.getCell(1).value;
            const rowTipster = (row.getCell(2).value || '').toString().trim();
            const rowMatch = row.getCell(4).value;

            if (datesMatchExcel(rowDate, pick.date) && rowTipster === pick.tipster && matchesMatchExcel(rowMatch, pick.match)) {
                row.getCell(8).value = resultLetter;
            }
        });
    }

    await workbook.xlsx.writeFile(EXCEL_FILE);

    // Update CSV backup with result
    updateResultInCsv(pick.tipster, pick.date, pick.match, resultLetter);

    console.log(`[Excel] Saved picks.xlsx`);
}

// ── CSV Backup (append-only fallback so data is never lost) ──

const CSV_HEADERS = 'Timestamp,Tipster,Date,Sport,Match,Pick,Odds,Bet,Result';

function escapeCsv(val) {
    const s = String(val ?? '');
    if (s.includes(',') || s.includes('"') || s.includes('\n')) {
        return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
}

function ensureCsvFile() {
    if (!fs.existsSync(CSV_FILE)) {
        fs.writeFileSync(CSV_FILE, CSV_HEADERS + '\n', 'utf8');
        console.log('[CSV] Created backup file: ' + CSV_FILE);
    }
}

function appendPickToCsv(tipster, pickData) {
    try {
        ensureCsvFile();
        const row = [
            new Date().toISOString(),
            escapeCsv(tipster),
            escapeCsv(pickData.date),
            escapeCsv(pickData.sport),
            escapeCsv(pickData.match),
            escapeCsv(pickData.pick),
            escapeCsv(pickData.odds),
            BET_AMOUNT,
            ''
        ].join(',');
        fs.appendFileSync(CSV_FILE, row + '\n', 'utf8');
    } catch (err) {
        console.error('[CSV] Backup write error: ' + err.message);
    }
}

function updateResultInCsv(tipster, date, match, result) {
    try {
        if (!fs.existsSync(CSV_FILE)) return;
        const lines = fs.readFileSync(CSV_FILE, 'utf8').split('\n');
        const normMatch = (val) => String(val ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
        const normDate = (val) => String(val ?? '').trim().toLowerCase();
        let updated = false;
        for (let i = 1; i < lines.length; i++) {
            if (!lines[i].trim()) continue;
            // Parse CSV line (handles quoted fields)
            const fields = parseCsvLine(lines[i]);
            if (!fields || fields.length < 9) continue;
            const csvTipster = fields[1].trim();
            const csvDate = fields[2].trim();
            const csvMatch = fields[4].trim();
            const csvResult = fields[8].trim();
            if (csvResult) continue; // already has a result
            if (csvTipster !== tipster) continue;
            if (normDate(csvDate) !== normDate(date)) continue;
            if (normMatch(csvMatch) !== normMatch(match) &&
                !normMatch(csvMatch).startsWith(normMatch(match)) &&
                !normMatch(match).startsWith(normMatch(csvMatch))) continue;
            fields[8] = result;
            lines[i] = fields.map(f => escapeCsv(f)).join(',');
            updated = true;
        }
        if (updated) {
            fs.writeFileSync(CSV_FILE, lines.join('\n'), 'utf8');
        }
    } catch (err) {
        console.error('[CSV] Backup result update error: ' + err.message);
    }
}

function parseCsvLine(line) {
    const fields = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (inQuotes) {
            if (ch === '"' && line[i + 1] === '"') {
                current += '"';
                i++;
            } else if (ch === '"') {
                inQuotes = false;
            } else {
                current += ch;
            }
        } else {
            if (ch === '"') {
                inQuotes = true;
            } else if (ch === ',') {
                fields.push(current);
                current = '';
            } else {
                current += ch;
            }
        }
    }
    fields.push(current);
    return fields;
}

module.exports = { initExcel, addPick, updateSummary, updatePickResult };