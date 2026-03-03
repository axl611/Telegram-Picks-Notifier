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

// Map from display sport names to API keys
const SPORT_API_KEYS = {
    'soccer': 'soccer',
    'nba': 'basketball_nba',
    'mlb': 'baseball_mlb',
    'nfl': 'americanfootball_nfl',
    'nhl': 'icehockey_nhl'
};

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

// Fetch scores from Odds API
function fetchScoresFromAPI(sport) {
    return new Promise((resolve, reject) => {
        const apiSport = SPORT_API_KEYS[sport.toLowerCase()] || sport;
        const url = `/v4/sports/${apiSport}/scores?daysFrom=1&apiKey=${_oddsApiKey}`;
        
        const req = https.request(
            { hostname: 'api.the-odds-api.com', path: url, method: 'GET', headers: { Accept: 'application/json' } },
            res => {
                let data = '';
                res.on('data', chunk => { data += chunk; });
                res.on('end', () => {
                    try {
                        const json = JSON.parse(data);
                        console.log(`   [Results] Fetched scores for ${sport}: ${json.data?.length || 0} games`);
                        resolve(json.data || []);
                    } catch (err) {
                        console.error(`   [Results] Error parsing scores for ${sport}:`, err.message);
                        reject(err);
                    }
                });
            }
        );
        req.on('error', reject);
        req.end();
    });
}

// Normalize strings for matching
function norm(str) {
    return (str || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9 ]/g, '')
        .replace(/\s+/g, ' ')
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
            console.log(`   [Results] Checking: ${match}`);

            const scores = await fetchScoresFromAPI(sportLower);
            const matchingScore = findScoreForPick({ match, sport }, scores);

            if (!matchingScore) {
                console.log(`   [Results] ⚠️  No matching score found: ${match}`);
                activeTimers.delete(pickId);
                return;
            }

            if (!matchingScore.completed) {
                console.log(`   [Results] Game not completed yet: ${match}`);
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

                // Update Excel
                let workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(EXCEL_FILE);

                const tipsterSheet = workbook.getWorksheet(tipster);
                const generalSheet = workbook.getWorksheet('General');

                if (!tipsterSheet) {
                    console.error(`   [Results] Sheet not found for tipster: ${tipster}`);
                    activeTimers.delete(pickId);
                    return;
                }

                let found = false;
                let balance = 0;

                // Find and update in tipster sheet
                tipsterSheet.eachRow((row, rowNum) => {
                    if (rowNum === 1) return;

                    const rowDate = (row.getCell(1).value || '').toString();
                    const rowMatch = (row.getCell(3).value || '').toString();

                    if (rowDate === date && rowMatch === match) {
                        row.getCell(7).value = result;      // Result column
                        row.getCell(8).value = pl;          // Profit/Loss column

                        // Calculate running balance
                        balance = 0;
                        tipsterSheet.eachRow((r, rn) => {
                            if (rn === 1 || rn > rowNum) return;
                            const v = parseFloat(r.getCell(8).value);
                            if (!isNaN(v)) balance += v;
                        });
                        balance += pl;

                        row.getCell(9).value = balance;      // Balance column

                        // Color the result cell
                        const colors = { 'W': 'FF92D050', 'L': 'FFFF6666', 'P': 'FFFFFF00' };
                        row.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[result] } };

                        found = true;
                    }
                });

                if (found) {
                    // Also update General sheet
                    generalSheet.eachRow((row, rowNum) => {
                        if (rowNum === 1) return;

                        const rowDate = (row.getCell(1).value || '').toString();
                        const rowTipster = (row.getCell(2).value || '').toString();
                        const rowMatch = (row.getCell(4).value || '').toString();

                        if (rowDate === date && rowTipster === tipster && rowMatch === match) {
                            row.getCell(8).value = result;   // Result column
                            row.getCell(9).value = pl;       // Profit/Loss column
                        }
                    });

                    await workbook.xlsx.writeFile(EXCEL_FILE);

                    // Update summary
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
                        `*P&L:* $${pl}\n` +
                        `*Balance:* $${balance.toFixed(2)}`
                    );
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