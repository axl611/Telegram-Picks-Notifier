// sheets.js - Google Sheets manager for picks tracking
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const BET_AMOUNT = 2000;
const STARTING_BALANCE = 0;
const TIPSTERS = ['Abuelo', 'Cristian Rey', 'Roberto Rey'];
const TIPSTER_COLUMNS = ['Date', 'Sport', 'Match', 'Pick', 'Odds', 'Bet', 'Result', 'Profit/Loss', 'Balance'];
const GENERAL_COLUMNS = ['Date', 'Tipster', 'Sport', 'Match', 'Pick', 'Odds', 'Bet', 'Result', 'Profit/Loss'];

let api;
let spreadsheetId;

// ── Auth ──────────────────────────────────────────────────────────────────────

async function ensureApi() {
    if (api) return;

    const credFile = process.env.GOOGLE_CREDENTIALS_FILE || path.join(process.cwd(), 'service-account.json');

    let auth;
    if (fs.existsSync(credFile)) {
        // Service account JSON key file
        auth = new google.auth.GoogleAuth({
            keyFile: credFile,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
    } else if (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
        // Service account via env vars
        auth = new google.auth.GoogleAuth({
            credentials: {
                client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
            },
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
    } else {
        // Application Default Credentials (gcloud auth application-default login)
        auth = new google.auth.GoogleAuth({
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
    }

    api = google.sheets({ version: 'v4', auth });
    spreadsheetId = process.env.GOOGLE_SHEET_ID;
    if (!spreadsheetId) throw new Error('GOOGLE_SHEET_ID missing in .env');
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function colLetter(n) {
    let s = '';
    while (n > 0) {
        n--;
        s = String.fromCharCode(65 + (n % 26)) + s;
        n = Math.floor(n / 26);
    }
    return s;
}

function q(title) {
    return `'${title}'`;
}

async function getSheetMeta() {
    const resp = await api.spreadsheets.get({ spreadsheetId });
    return resp.data.sheets.map(s => ({
        title: s.properties.title,
        sheetId: s.properties.sheetId,
    }));
}

// ── Sheet Setup ───────────────────────────────────────────────────────────────

async function ensureHeaders(title, columns) {
    const range = `${q(title)}!A1:${colLetter(columns.length)}1`;
    const resp = await api.spreadsheets.values.get({ spreadsheetId, range });
    const existing = (resp.data.values || [])[0] || [];

    if (columns.length === existing.length && columns.every((c, i) => existing[i] === c)) return;

    await api.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'RAW',
        requestBody: { values: [columns] },
    });
}

function conditionalColorRule(sheetId, colIndex, value, bgColor) {
    return {
        addConditionalFormatRule: {
            rule: {
                ranges: [{ sheetId, startRowIndex: 1, startColumnIndex: colIndex, endColumnIndex: colIndex + 1, endRowIndex: 1000 }],
                booleanRule: {
                    condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: value }] },
                    format: { backgroundColor: bgColor },
                },
            },
            index: 0,
        },
    };
}

async function applyFormatting() {
    const meta = await getSheetMeta();
    const requests = [];

    const headerBg = { red: 0.12, green: 0.31, blue: 0.47 };
    const headerFg = { red: 1, green: 1, blue: 1 };

    // ── Tipster sheets ──
    for (const tipster of TIPSTERS) {
        const sheet = meta.find(s => s.title === tipster);
        if (!sheet) continue;
        const sid = sheet.sheetId;

        // Header row: bold white text on dark blue
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1 },
                cell: {
                    userEnteredFormat: {
                        textFormat: { bold: true, foregroundColorStyle: { rgbColor: headerFg } },
                        backgroundColor: headerBg,
                        horizontalAlignment: 'CENTER',
                    },
                },
                fields: 'userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)',
            },
        });

        // Data rows: reset to default (no background, black text, not bold)
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 1, endRowIndex: 1000 },
                cell: {
                    userEnteredFormat: {
                        textFormat: { bold: false, foregroundColorStyle: { rgbColor: { red: 0, green: 0, blue: 0 } } },
                        backgroundColor: { red: 1, green: 1, blue: 1 },
                    },
                },
                fields: 'userEnteredFormat(textFormat,backgroundColor)',
            },
        });

        // Currency: Bet (F=5), Profit/Loss (H=7), Balance (I=8)
        for (const col of [5, 7, 8]) {
            requests.push({
                repeatCell: {
                    range: { sheetId: sid, startRowIndex: 1, startColumnIndex: col, endColumnIndex: col + 1, endRowIndex: 1000 },
                    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
                    fields: 'userEnteredFormat.numberFormat',
                },
            });
        }

        // W/L/P dropdown on Result (G=6)
        requests.push({
            setDataValidation: {
                range: { sheetId: sid, startRowIndex: 1, startColumnIndex: 6, endColumnIndex: 7, endRowIndex: 1000 },
                rule: {
                    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'W' }, { userEnteredValue: 'L' }, { userEnteredValue: 'P' }] },
                    showCustomUi: true,
                    strict: false,
                },
            },
        });

        // Conditional colors for Result column
        requests.push(
            conditionalColorRule(sid, 6, 'W', { red: 0.57, green: 0.82, blue: 0.31 }),
            conditionalColorRule(sid, 6, 'L', { red: 1, green: 0.4, blue: 0.4 }),
            conditionalColorRule(sid, 6, 'P', { red: 1, green: 1, blue: 0 }),
        );

        // Column widths
        const widths = [100, 80, 280, 240, 60, 80, 60, 110, 110];
        for (let i = 0; i < widths.length; i++) {
            requests.push({
                updateDimensionProperties: {
                    range: { sheetId: sid, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
                    properties: { pixelSize: widths[i] },
                    fields: 'pixelSize',
                },
            });
        }
    }

    // ── General sheet ──
    const genSheet = meta.find(s => s.title === 'General');
    if (genSheet) {
        const sid = genSheet.sheetId;

        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1 },
                cell: {
                    userEnteredFormat: {
                        textFormat: { bold: true, foregroundColorStyle: { rgbColor: headerFg } },
                        backgroundColor: headerBg,
                        horizontalAlignment: 'CENTER',
                    },
                },
                fields: 'userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)',
            },
        });

        // Data rows: reset to default
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 1, endRowIndex: 1000 },
                cell: {
                    userEnteredFormat: {
                        textFormat: { bold: false, foregroundColorStyle: { rgbColor: { red: 0, green: 0, blue: 0 } } },
                        backgroundColor: { red: 1, green: 1, blue: 1 },
                    },
                },
                fields: 'userEnteredFormat(textFormat,backgroundColor)',
            },
        });

        // Currency: Bet (G=6), Profit/Loss (I=8)
        for (const col of [6, 8]) {
            requests.push({
                repeatCell: {
                    range: { sheetId: sid, startRowIndex: 1, startColumnIndex: col, endColumnIndex: col + 1, endRowIndex: 1000 },
                    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
                    fields: 'userEnteredFormat.numberFormat',
                },
            });
        }

        requests.push({
            setDataValidation: {
                range: { sheetId: sid, startRowIndex: 1, startColumnIndex: 7, endColumnIndex: 8, endRowIndex: 1000 },
                rule: {
                    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'W' }, { userEnteredValue: 'L' }, { userEnteredValue: 'P' }] },
                    showCustomUi: true,
                    strict: false,
                },
            },
        });

        requests.push(
            conditionalColorRule(sid, 7, 'W', { red: 0.57, green: 0.82, blue: 0.31 }),
            conditionalColorRule(sid, 7, 'L', { red: 1, green: 0.4, blue: 0.4 }),
            conditionalColorRule(sid, 7, 'P', { red: 1, green: 1, blue: 0 }),
        );

        const widths = [100, 120, 80, 280, 240, 60, 80, 60, 110];
        for (let i = 0; i < widths.length; i++) {
            requests.push({
                updateDimensionProperties: {
                    range: { sheetId: sid, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
                    properties: { pixelSize: widths[i] },
                    fields: 'pixelSize',
                },
            });
        }
    }

    // ── Summary sheet ──
    const sumSheet = meta.find(s => s.title === 'Summary');
    if (sumSheet) {
        const sid = sumSheet.sheetId;

        // Title style
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 },
                cell: {
                    userEnteredFormat: {
                        textFormat: { bold: true, fontSize: 14, foregroundColorStyle: { rgbColor: { red: 0.12, green: 0.31, blue: 0.47 } } },
                    },
                },
                fields: 'userEnteredFormat.textFormat',
            },
        });

        // Header row (row 2)
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 1, endRowIndex: 2 },
                cell: {
                    userEnteredFormat: {
                        textFormat: { bold: true, foregroundColorStyle: { rgbColor: headerFg } },
                        backgroundColor: headerBg,
                        horizontalAlignment: 'CENTER',
                    },
                },
                fields: 'userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)',
            },
        });

        // Label column bold
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 2, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1 },
                cell: { userEnteredFormat: { textFormat: { bold: true } } },
                fields: 'userEnteredFormat.textFormat',
            },
        });

        // Currency for Total P/L (row 8, index 7)
        requests.push({
            repeatCell: {
                range: { sheetId: sid, startRowIndex: 7, endRowIndex: 8, startColumnIndex: 1, endColumnIndex: 5 },
                cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
                fields: 'userEnteredFormat.numberFormat',
            },
        });

        // Percentage for Win Rate (row 7, index 6) and Yield (row 9, index 8)
        for (const rowIdx of [6, 8]) {
            requests.push({
                repeatCell: {
                    range: { sheetId: sid, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 1, endColumnIndex: 5 },
                    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
                    fields: 'userEnteredFormat.numberFormat',
                },
            });
        }

        // Column widths
        const widths = [160, 120, 120, 120, 120];
        for (let i = 0; i < widths.length; i++) {
            requests.push({
                updateDimensionProperties: {
                    range: { sheetId: sid, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
                    properties: { pixelSize: widths[i] },
                    fields: 'pixelSize',
                },
            });
        }
    }

    if (requests.length > 0) {
        await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
    }
}

// ── Summary formulas ──────────────────────────────────────────────────────────

async function buildSummarySheet() {
    const values = [
        ['PICKS TRACKER - SUMMARY', '', '', '', ''],
        ['', 'Abuelo', 'Cristian Rey', 'Roberto Rey', 'TOTAL'],
        ['Total Picks',
            '=COUNTA(Abuelo!A2:A1000)',
            "=COUNTA('Cristian Rey'!A2:A1000)",
            "=COUNTA('Roberto Rey'!A2:A1000)",
            '=B3+C3+D3'],
        ['Wins',
            '=COUNTIF(Abuelo!G:G,"W")',
            "=COUNTIF('Cristian Rey'!G:G,\"W\")",
            "=COUNTIF('Roberto Rey'!G:G,\"W\")",
            '=B4+C4+D4'],
        ['Losses',
            '=COUNTIF(Abuelo!G:G,"L")',
            "=COUNTIF('Cristian Rey'!G:G,\"L\")",
            "=COUNTIF('Roberto Rey'!G:G,\"L\")",
            '=B5+C5+D5'],
        ['Pushes',
            '=COUNTIF(Abuelo!G:G,"P")',
            "=COUNTIF('Cristian Rey'!G:G,\"P\")",
            "=COUNTIF('Roberto Rey'!G:G,\"P\")",
            '=B6+C6+D6'],
        ['Win Rate %',
            '=IF((B4+B5)=0,0,B4/(B4+B5))',
            '=IF((C4+C5)=0,0,C4/(C4+C5))',
            '=IF((D4+D5)=0,0,D4/(D4+D5))',
            '=IF((B4+B5+C4+C5+D4+D5)=0,0,(B4+C4+D4)/(B4+B5+C4+C5+D4+D5))'],
        ['Total Profit/Loss',
            '=SUMIF(Abuelo!G:G,"W",Abuelo!H:H)+SUMIF(Abuelo!G:G,"L",Abuelo!H:H)',
            "=SUMIF('Cristian Rey'!G:G,\"W\",'Cristian Rey'!H:H)+SUMIF('Cristian Rey'!G:G,\"L\",'Cristian Rey'!H:H)",
            "=SUMIF('Roberto Rey'!G:G,\"W\",'Roberto Rey'!H:H)+SUMIF('Roberto Rey'!G:G,\"L\",'Roberto Rey'!H:H)",
            '=B8+C8+D8'],
        ['Yield %',
            `=IF(B3=0,0,B8/(B3*${BET_AMOUNT}))`,
            `=IF(C3=0,0,C8/(C3*${BET_AMOUNT}))`,
            `=IF(D3=0,0,D8/(D3*${BET_AMOUNT}))`,
            `=IF(E3=0,0,E8/(E3*${BET_AMOUNT}))`],
    ];

    await api.spreadsheets.values.update({
        spreadsheetId,
        range: 'Summary!A1:E9',
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
    });
}

// ── Date / Match comparison ───────────────────────────────────────────────────

function normalizeDateForCompare(val) {
    if (val == null || val === '') return '';
    return String(val).trim();
}

function normalizeMatchForCompare(val) {
    if (val == null) return '';
    return String(val).replace(/\s+/g, ' ').trim();
}

function datesMatch(a, b) {
    return normalizeDateForCompare(a) === normalizeDateForCompare(b);
}

function matchesMatch(rowMatch, pickMatch) {
    const r = normalizeMatchForCompare(rowMatch);
    const p = normalizeMatchForCompare(pickMatch);
    if (r === p) return true;
    if (r.startsWith(p) || p.startsWith(r)) return true;
    return false;
}

// ── Public API ────────────────────────────────────────────────────────────────

async function initSheets() {
    await ensureApi();

    const meta = await getSheetMeta();
    const existingTitles = meta.map(s => s.title);

    // Create missing sheets in a single batch request
    const toCreate = [...TIPSTERS, 'General', 'Summary'].filter(t => !existingTitles.includes(t));
    if (toCreate.length > 0) {
        await api.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: toCreate.map(title => ({ addSheet: { properties: { title } } })),
            },
        });
    }

    // Set headers
    for (const tipster of TIPSTERS) {
        await ensureHeaders(tipster, TIPSTER_COLUMNS);
    }
    await ensureHeaders('General', GENERAL_COLUMNS);

    // Build summary with live formulas
    await buildSummarySheet();

    // Apply all formatting, conditional colors, data validation, column widths
    await applyFormatting();

    console.log('[Sheets] Initialized: https://docs.google.com/spreadsheets/d/' + spreadsheetId);
}

async function addPick(pickData) {
    await ensureApi();

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

    let tipsterName = pickData.tipster || 'Unknown';
    const tl = tipsterName.toLowerCase();
    if (tl.includes('abuelo')) tipsterName = 'Abuelo';
    else if (tl.includes('cristian')) tipsterName = 'Cristian Rey';
    else if (tl.includes('roberto') || tl.includes('beto')) tipsterName = 'Roberto Rey';

    // ── Tipster sheet ──
    const tipResp = await api.spreadsheets.values.get({
        spreadsheetId,
        range: `${q(tipsterName)}!A:A`,
    });
    const tipRowCount = (tipResp.data.values || []).length;
    const newRow = tipRowCount + 1;

    const plFormula = `=IF(G${newRow}="W",E${newRow}*${BET_AMOUNT},IF(G${newRow}="L",-${BET_AMOUNT},0))`;
    const balFormula = newRow === 2
        ? `=${STARTING_BALANCE}+H${newRow}`
        : `=I${newRow - 1}+H${newRow}`;

    await api.spreadsheets.values.append({
        spreadsheetId,
        range: `${q(tipsterName)}!A:I`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
            values: [[
                pickData.date || '',
                pickData.sport || 'SOCCER',
                pickData.match || '',
                pickData.pick || '',
                pickData.odds || '',
                BET_AMOUNT,
                '',
                plFormula,
                balFormula,
            ]],
        },
    });

    // ── General sheet ──
    const genResp = await api.spreadsheets.values.get({
        spreadsheetId,
        range: 'General!A:A',
    });
    const genRowCount = (genResp.data.values || []).length;
    const genRow = genRowCount + 1;

    const genPlFormula = `=IF(H${genRow}="W",F${genRow}*${BET_AMOUNT},IF(H${genRow}="L",-${BET_AMOUNT},0))`;

    await api.spreadsheets.values.append({
        spreadsheetId,
        range: 'General!A:I',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
            values: [[
                pickData.date || '',
                tipsterName,
                pickData.sport || 'SOCCER',
                pickData.match || '',
                pickData.pick || '',
                pickData.odds || '',
                BET_AMOUNT,
                '',
                genPlFormula,
            ]],
        },
    });

    console.log('Pick saved: ' + tipsterName + ' | ' + pickData.match + ' | ' + pickData.pick + ' | Odds: ' + pickData.odds);
}

async function updateSummary() {
    await ensureApi();
    await buildSummarySheet();
    console.log('   [Summary] Updated');
}

async function updatePickResult(pick, result) {
    await ensureApi();

    const resultLetter = result === 'Win' ? 'W' : result === 'Loss' ? 'L' : result === 'Push' ? 'P' : result;
    console.log(`[Sheets] updatePickResult: tipster=${pick.tipster}, date=${pick.date}, match=${pick.match}, result=${resultLetter}`);

    // Read tipster sheet rows
    const tipResp = await api.spreadsheets.values.get({
        spreadsheetId,
        range: `${q(pick.tipster)}!A2:G`,
    });
    const tipData = tipResp.data.values || [];

    const batchData = [];
    let found = false;

    for (let i = 0; i < tipData.length; i++) {
        const [date, , match] = tipData[i];
        if (datesMatch(date, pick.date) && matchesMatch(match, pick.match)) {
            const rowNum = i + 2;
            batchData.push({ range: `${q(pick.tipster)}!G${rowNum}`, values: [[resultLetter]] });
            found = true;
            console.log(`[Sheets] Updated tipster row ${rowNum}: ${resultLetter}`);
        }
    }

    // Read General sheet
    const genResp = await api.spreadsheets.values.get({
        spreadsheetId,
        range: 'General!A2:H',
    });
    const genData = genResp.data.values || [];

    for (let i = 0; i < genData.length; i++) {
        const [date, tipster, , match] = genData[i];
        if (datesMatch(date, pick.date) && (tipster || '').trim() === pick.tipster && matchesMatch(match, pick.match)) {
            const rowNum = i + 2;
            batchData.push({ range: `General!H${rowNum}`, values: [[resultLetter]] });
            console.log(`[Sheets] Updated General row ${rowNum}: ${resultLetter}`);
        }
    }

    if (batchData.length > 0) {
        await api.spreadsheets.values.batchUpdate({
            spreadsheetId,
            requestBody: { valueInputOption: 'RAW', data: batchData },
        });
    }

    if (!found) {
        console.log(`[Sheets] Warning: Could not find pick: date="${pick.date}" match="${pick.match}"`);
    } else {
        console.log('[Sheets] Result saved');
    }
}

async function getPicksWithoutResults() {
    await ensureApi();

    const picksToCheck = [];

    for (const tipster of TIPSTERS) {
        const resp = await api.spreadsheets.values.get({
            spreadsheetId,
            range: `${q(tipster)}!A2:G`,
        });
        const rows = resp.data.values || [];

        for (const row of rows) {
            const [date, sport, match, pick, odds, bet, result] = row;
            if (!date || !match) continue;
            if (result && result.trim()) continue;

            picksToCheck.push({
                tipster,
                date: (date || '').trim(),
                match: (match || '').trim(),
                pick: (pick || '').trim(),
                sport: ((sport || 'soccer')).toLowerCase(),
                odds: odds || '',
            });
        }
    }

    return picksToCheck;
}

module.exports = { initSheets, addPick, updateSummary, updatePickResult, getPicksWithoutResults };
