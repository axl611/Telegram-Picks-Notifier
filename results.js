// results.js - Event-based result checker with sport-specific delays
// Checks soccer/NBA after 1h 20min, MLB/NFL/NHL after 3 hours
const https = require('https');
const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE = path.join(process.cwd(), 'picks.xlsx');
const BET_AMOUNT = 2000;
const TIPSTERS = ['Abuelo', 'Cristian Rey', 'Roberto Rey'];

// Sport-specific delays (in milliseconds)
const RESULT_CHECK_DELAYS = {
    'soccer': 80 * 60 * 1000,      // 1h 20min
    'nba': 80 * 60 * 1000,         // 1h 20min
    'mlb': 3 * 60 * 60 * 1000,     // 3 hours
    'nfl': 3 * 60 * 60 * 1000,     // 3 hours
    'nhl': 3 * 60 * 60 * 1000      // 3 hours
};

// Direct sport key for non-soccer sports
const SPORT_API_KEYS = {
    'nba': 'basketball_nba',
    'mlb': 'baseball_mlb',
    'nfl': 'americanfootball_nfl',
    'nhl': 'icehockey_nhl'
};

// Common soccer leagues to try (auto-discovery not possible in timer context)
const COMMON_SOCCER_LEAGUES = [
    'soccer_epl', 'soccer_efl_champ', 'soccer_england_league1',
    'soccer_spain_la_liga', 'soccer_germany_bundesliga',
    'soccer_italy_serie_a', 'soccer_france_ligue_one',
    'soccer_usa_mls', 'soccer_mexico_ligamx',
    'soccer_uefa_champs_league', 'soccer_uefa_europa_league',
    'soccer_brazil_campeonato', 'soccer_argentina_primera_division'
];

// Store active timers by pick ID so we can cancel if needed
const activeTimers = new Map();

let _telegramToken = null;
let _telegramChatId = null;
let _oddsApiKey = null;

function init(telegramToken, telegramChatId, oddsApiKey) {
    _telegramToken = telegramToken;
    _telegramChatId = telegramChatId;
    _oddsApiKey = oddsApiKey;
}

function sendTelegram(text) {
    return new Promise(resolve => {
        if (!_telegramToken || !_telegramChatId) return resolve();
        const body = JSON.stringify({ chat_id: _telegramChatId, text, parse_mode: 'Markdown' });
        const req = https.request(
            {
                hostname: 'api.telegram.org',
                path: '/bot' + _telegramToken + '/sendMessage',
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
            },
            res => { res.on('data', () => {}); res.on('end', resolve); }
        );
        req.on('error', () => resolve());
        req.write(body);
        req.end();
    });
}

// Generic API GET
function apiGet(urlPath) {
    return new Promise((resolve, reject) => {
        const req = https.request(
            { hostname: 'api.the-odds-api.com', path: urlPath, method: 'GET', headers: { Accept: 'application/json' } },
            res => {
                const remaining = res.headers['x-requests-remaining'];
                if (remaining !== undefined) {
                    console.log(`   [Results] API credits remaining: ${remaining}`);
                }
                let data = '';
                res.on('data', chunk => { data += chunk; });
                res.on('end', () => {
                    try {
                        const json = JSON.parse(data);
                        resolve(Array.isArray(json) ? json : (json.data || []));
                    } catch (err) {
                        reject(err);
                    }
                });
            }
        );
        req.on('error', reject);
        req.end();
    });
}

// Fetch scores for a single sport key (2 credits with daysFrom)
function fetchScoresByKey(sportKey) {
    const url = `/v4/sports/${sportKey}/scores?daysFrom=1&apiKey=${_oddsApiKey}`;
    return apiGet(url);
}

// Fetch scores — auto-resolves soccer to multiple league keys
async function fetchScoresFromAPI(sport) {
    const sportLower = sport.toLowerCase();
    const directKey = SPORT_API_KEYS[sportLower];
    if (directKey) {
        const scores = await fetchScoresByKey(directKey);
        console.log(`   [Results] Fetched scores for ${directKey}: ${scores.length} games`);
        return scores;
    }

    // Soccer: try common leagues, collect all scores
    console.log(`   [Results] Soccer — querying ${COMMON_SOCCER_LEAGUES.length} league(s)...`);
    let allScores = [];
    for (const key of COMMON_SOCCER_LEAGUES) {
        try {
            const scores = await fetchScoresByKey(key);
            if (scores.length > 0) {
                console.log(`   [Results] ${key}: ${scores.length} game(s)`);
                allScores = allScores.concat(scores);
            }
        } catch { /* league may not be active */ }
    }
    console.log(`   [Results] Total soccer scores: ${allScores.length}`);
    return allScores;
}

// Normalize strings for matching — strips OCR noise like "Sí", "FC", etc.
const TEAM_NOISE_WORDS = new Set(['sí', 'si', 'no', 'yes', 'fc', 'cf', 'sc', 'ac', 'afc', 'bc']);
function norm(str) {
    return (str || '')
        .split(/\s+/)
        .map(w => w.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, ''))
        .filter(w => w && !TEAM_NOISE_WORDS.has(w))
        .join(' ')
        .trim();
}

// Find matching score for a pick
function findScoreForPick(pick, scores) {
    const pickTeams = pick.match.split(/\s+vs\.?\s+/i).map(t => t.trim());
    
    if (pickTeams.length < 2) {
        console.log(`   [Results] Could not parse teams from match: ${pick.match}`);
        return null;
    }
    
    const [homeTeam, awayTeam] = pickTeams;
    const homeNorm = norm(homeTeam);
    const awayNorm = norm(awayTeam);
    
    for (const score of scores) {
        const homeScore = norm(score.home_team);
        const awayScore = norm(score.away_team);
        
        const homeMatch = homeNorm.includes(homeScore) || homeScore.includes(homeNorm) || awayNorm.includes(homeScore) || homeScore.includes(awayNorm);
        const awayMatch = homeNorm.includes(awayScore) || awayScore.includes(homeNorm) || awayNorm.includes(awayScore) || awayScore.includes(awayNorm);
        
        if (homeMatch && awayMatch) {
            return score;
        }
    }
    
    return null;
}

// Get team score from game
function teamScore(game, teamName) {
    if (!game.scores) return null;
    const n = norm(teamName);
    const entry = game.scores.find(s => norm(s.name).includes(n) || n.includes(norm(s.name)));
    return entry ? parseInt(entry.score) : null;
}

// Evaluate pick result based on game scores
function evaluatePick(pickText, game) {
    if (!game || !game.completed || !game.scores || game.scores.length < 2) return null;

    const homeScore = teamScore(game, game.home_team);
    const awayScore = teamScore(game, game.away_team);
    if (homeScore === null || awayScore === null) return null;

    const total = homeScore + awayScore;
    const lower = (pickText || '').toLowerCase();

    console.log(`   [Results] Score: ${game.home_team} ${homeScore} - ${awayScore} ${game.away_team} (total ${total})`);

    // BTTS (Both Teams to Score)
    if (/ambos equipos|btts/i.test(lower)) {
        const both = homeScore > 0 && awayScore > 0;
        if (/\bno\b/i.test(lower)) return both ? 'L' : 'W';
        return both ? 'W' : 'L';
    }

    // Over
    const overM = lower.match(/over\s+(\d+\.?\d*)/);
    if (overM) {
        const line = parseFloat(overM[1]);
        if (total === line) return 'P';
        return total > line ? 'W' : 'L';
    }

    // Under
    const underM = lower.match(/under\s+(\d+\.?\d*)/);
    if (underM) {
        const line = parseFloat(underM[1]);
        if (total === line) return 'P';
        return total < line ? 'W' : 'L';
    }

    // Asian Handicap
    const ahM = lower.match(/ah\s+([+-]?\d+\.?\d*)/);
    if (ahM) {
        const handicap = parseFloat(ahM[1]);
        const pickedTeam = norm(pickText.replace(/ah.*/i, '').split(/\s+vs\.?\s+/i)[0].trim());
        const homeN = norm(game.home_team);
        const awayN = norm(game.away_team);
        let diff;
        if (homeN.includes(pickedTeam) || pickedTeam.includes(homeN)) {
            diff = (homeScore - awayScore) + handicap;
        } else if (awayN.includes(pickedTeam) || pickedTeam.includes(awayN)) {
            diff = (awayScore - homeScore) + handicap;
        } else {
            return null;
        }
        if (diff === 0) return 'P';
        return diff > 0 ? 'W' : 'L';
    }

    // ML (default) - pick first team in match string
    const mlTeam = norm(pickText.replace(/\s*ml\s*$/i, '').trim());
    const homeN  = norm(game.home_team);
    const awayN  = norm(game.away_team);
    const pickedHome = homeN.includes(mlTeam) || mlTeam.includes(homeN);
    const pickedAway = awayN.includes(mlTeam) || mlTeam.includes(awayN);
    
    if (!pickedHome && !pickedAway) return null;
    if (homeScore === awayScore) return 'L'; // draw = loss for soccer ML
    const homeWon = homeScore > awayScore;
    if (pickedHome) return homeWon ? 'W' : 'L';
    if (pickedAway) return homeWon ? 'L' : 'W';

    return null;
}

function calcPL(result, odds) {
    const o = parseFloat(odds) || 0;
    if (result === 'W') return parseFloat(((o - 1) * BET_AMOUNT).toFixed(2));
    if (result === 'L') return -BET_AMOUNT;
    if (result === 'P') return 0;
    return null;
}

// Generate unique ID for a pick
function generatePickId(tipster, date, match) {
    return `${tipster}|${date}|${match}`.replace(/\s+/g, '_');
}

// Normalize date for comparison (Excel may return Date object or string like "2 Mar")
function normalizeDateForCompare(val) {
    if (val == null || val === '') return '';
    if (val instanceof Date) {
        return val.getDate() + ' ' + val.toLocaleString('en', { month: 'short' });
    }
    const s = String(val).trim();
    return s;
}

// Normalize match for comparison (trim, collapse spaces; Excel may truncate with "...")
function normalizeMatchForCompare(val) {
    if (val == null) return '';
    return String(val).replace(/\s+/g, ' ').trim();
}

function datesMatch(a, b) {
    const na = normalizeDateForCompare(a);
    const nb = normalizeDateForCompare(b);
    return na === nb;
}

function matchesMatch(rowMatch, pickMatch) {
    const r = normalizeMatchForCompare(rowMatch);
    const p = normalizeMatchForCompare(pickMatch);
    if (r === p) return true;
    // Excel truncates match at 60 chars (55 + '...'), so check if one starts the other
    if (r.startsWith(p) || p.startsWith(r)) return true;
    return false;
}

// Schedule result check for a pick
async function scheduleResultCheck(tipster, date, match, pick, odds, sport) {
    if (!sport) {
        console.log(`   [Results] Missing sport for pick, skipping: ${match}`);
        return;
    }

    const pickId = generatePickId(tipster, date, match);
    const sportLower = sport.toLowerCase();
    const delay = RESULT_CHECK_DELAYS[sportLower];

    if (!delay) {
        console.log(`   [Results] Unknown sport for pick, skipping: ${sportLower}`);
        return;
    }

    const delayMinutes = Math.round(delay / 60 / 1000);
    console.log(`   [Results] ⏰ Scheduled check for ${sportLower} (${delayMinutes}min): ${match}`);

    // Cancel any existing timer for this pick
    if (activeTimers.has(pickId)) {
        clearTimeout(activeTimers.get(pickId));
    }

    // Schedule the check
    const timer = setTimeout(async () => {
        try {
            console.log(`   [Results] ── Checking result for: ${tipster} | ${date} | ${match} ──`);

            const scores = await fetchScoresFromAPI(sportLower);
            console.log(`   [Results] Got ${scores.length} score(s) for ${sportLower}`);

            const matchingScore = findScoreForPick({ match, sport }, scores);

            if (!matchingScore) {
                console.log(`   [Results] ⚠️  No matching score found for: ${match}`);
                activeTimers.delete(pickId);
                return;
            }

            console.log(`   [Results] Found game: ${matchingScore.home_team} vs ${matchingScore.away_team}, completed=${matchingScore.completed}`);

            if (!matchingScore.completed) {
                console.log(`   [Results] Game not completed yet, rescheduling in 5min: ${match}`);
                // Reschedule for 5 minutes later
                const retryTimer = setTimeout(() => {
                    scheduleResultCheck(tipster, date, match, pick, odds, sport);
                }, 5 * 60 * 1000);
                activeTimers.set(pickId, retryTimer);
                return;
            }

            const result = evaluatePick(pick, matchingScore);

            if (result) {
                const pl = calcPL(result, odds);

                console.log(`   [Results] Updating Excel: tipster=${tipster}, date=${date}, match=${match}, result=${result}`);

                // Update Excel
                let workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(EXCEL_FILE);
                console.log(`   [Results] Loaded workbook, looking for row...`);

                const tipsterSheet = workbook.getWorksheet(tipster);
                const generalSheet = workbook.getWorksheet('General');

                if (!tipsterSheet) {
                    console.error(`   [Results] Sheet not found for tipster: ${tipster}`);
                    activeTimers.delete(pickId);
                    return;
                }

                let found = false;

                // Find and update in tipster sheet (use normalized date/match comparison)
                tipsterSheet.eachRow((row, rowNum) => {
                    if (rowNum === 1) return;

                    const rowDate = row.getCell(1).value;
                    const rowMatch = row.getCell(3).value;

                    const dateOk = datesMatch(rowDate, date);
                    const matchOk = matchesMatch(rowMatch, match);

                    if (dateOk && matchOk) {
                        row.getCell(7).value = result;      // Result column ONLY
                        // Profit/Loss (column 8) and Balance (column 9) are formula-based - they auto-calculate

                        // Color the result cell
                        const colors = { 'W': 'FF92D050', 'L': 'FFFF6666', 'P': 'FFFFFF00' };
                        row.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[result] } };

                        found = true;
                        console.log(`   [Results] Matched tipster row ${rowNum}: date=${normalizeDateForCompare(rowDate)}, match=${normalizeMatchForCompare(rowMatch)}`);
                    }
                });

                if (found) {
                    // Also update General sheet
                    let generalUpdated = 0;
                    if (generalSheet) generalSheet.eachRow((row, rowNum) => {
                        if (rowNum === 1) return;

                        const rowDate = row.getCell(1).value;
                        const rowTipster = (row.getCell(2).value || '').toString().trim();
                        const rowMatch = row.getCell(4).value;

                        if (datesMatch(rowDate, date) && rowTipster === tipster && matchesMatch(rowMatch, match)) {
                            row.getCell(8).value = result;   // Result column ONLY
                            generalUpdated++;
                        }
                    });
                    if (generalSheet) console.log(`   [Results] Updated General sheet: ${generalUpdated} row(s)`);

                    await workbook.xlsx.writeFile(EXCEL_FILE);
                    console.log(`   [Results] Saved picks.xlsx`);

                    // Update summary (rebuilds formulas)
                    const { updateSummary } = require('./excel');
                    await updateSummary();

                    const emoji = result === 'W' ? '✅' : result === 'L' ? '❌' : '➡️';
                    console.log(`   [Results] ${emoji} ${tipster} | ${match} | ${pick} → ${result} ($${pl})`);

                    await sendTelegram(
                        `${emoji} *Result: ${result}*\n` +
                        `*Tipster:* ${tipster}\n` +
                        `*Match:* ${match}\n` +
                        `*Pick:* ${pick}\n` +
                        `*Odds:* ${odds}\n` +
                        `*P&L:* $${pl}`
                    );
                } else {
                    console.log(`   [Results] No matching row in Excel for date="${date}" match="${match}" (tipster=${tipster}). Check date format in sheet (e.g. "2 Mar").`);
                }
            } else {
                console.log(`   [Results] Could not evaluate pick: ${match} | ${pick}`);
            }

            activeTimers.delete(pickId);
        } catch (err) {
            console.error(`   [Results] Error checking result:`, err.message);
            activeTimers.delete(pickId);
        }
    }, delay);

    activeTimers.set(pickId, timer);
}

// Cancel all pending checks (on shutdown)
function cancelAllChecks() {
    console.log(`[Results] Cancelling ${activeTimers.size} pending result checks`);
    for (const timer of activeTimers.values()) {
        clearTimeout(timer);
    }
    activeTimers.clear();
}

function startResultsScheduler() {
    console.log('📊 Event-based result checker started (sport-specific delays)');
}

module.exports = { init, startResultsScheduler, cancelAllChecks, scheduleResultCheck };