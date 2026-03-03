// results-checker.js
// Test script to force W/L/P result checks for the 3 tipster tabs (bypasses the API schedule).
// The main app uses results.js with sport-specific delays; this script runs on demand and
// fills the Result column in picks.xlsx for any pick that doesn't have one yet.
// Run: node results-checker.js           (update Excel)
// Run: node results-checker.js --dry-run (show results, don't write)

const https = require('https');
const config = require('./config');
const excel = require('./excel');

// Sport-specific delays (in milliseconds)
const RESULT_CHECK_DELAYS = {
    'soccer': 80 * 60 * 1000,    // 1h 20min
    'nba': 80 * 60 * 1000,       // 1h 20min
    'mlb': 3 * 60 * 60 * 1000,   // 3 hours
    'nfl': 3 * 60 * 60 * 1000,   // 3 hours
    'nhl': 3 * 60 * 60 * 1000    // 3 hours
};

// Direct sport key for non-soccer sports
const SPORT_API_KEYS = {
    'nba': 'basketball_nba',
    'mlb': 'baseball_mlb',
    'nfl': 'americanfootball_nfl',
    'nhl': 'icehockey_nhl'
};

// Store active timers by pick ID so we can cancel if needed
const activeTimers = new Map();

// ── Generic JSON GET (used by all API calls) ──
function apiGet(url, logCredits = false) {
    return new Promise((resolve, reject) => {
        https.get(url, (res) => {
            if (logCredits) {
                const remaining = res.headers['x-requests-remaining'];
                const used = res.headers['x-requests-used'];
                if (remaining !== undefined) {
                    console.log(`[ResultsChecker] API credits: ${used} used, ${remaining} remaining`);
                }
            }
            let data = '';
            res.on('data', chunk => { data += chunk; });
            res.on('end', () => {
                try {
                    const json = JSON.parse(data);
                    resolve(Array.isArray(json) ? json : (json.data || []));
                } catch (err) {
                    console.error(`[ResultsChecker] JSON parse error: ${data.substring(0, 300)}`);
                    reject(err);
                }
            });
        }).on('error', reject);
    });
}

// ── Get scores for a sport key ──
// Without daysFrom: 1 credit (live + upcoming only)
// With daysFrom=1: 2 credits (adds today's completed games)
async function fetchScores(sportKey, daysFrom) {
    const daysParam = daysFrom ? `&daysFrom=${daysFrom}` : '';
    const url = `https://api.the-odds-api.com/v4/sports/${sportKey}/scores?apiKey=${config.ODDS_API_KEY}${daysParam}`;
    const scores = await apiGet(url, true);
    const cost = daysFrom ? 2 : 1;
    console.log(`[ResultsChecker]   ${sportKey}: ${scores.length} game(s) (${cost} credit)`);
    return scores;
}

// Common soccer leagues — used as fallback when events lookup fails (completed games)
const COMMON_SOCCER_LEAGUES = [
    'soccer_epl', 'soccer_efl_champ', 'soccer_england_league1',
    'soccer_spain_la_liga', 'soccer_germany_bundesliga',
    'soccer_italy_serie_a', 'soccer_france_ligue_one',
    'soccer_usa_mls', 'soccer_mexico_ligamx',
    'soccer_uefa_champs_league', 'soccer_uefa_europa_league',
    'soccer_brazil_campeonato', 'soccer_argentina_primera_division'
];

// ── Fetch scores for a sport, stop early when match is found ──
// daysFrom=1 (today's completed): 2 credits/league — used by production
// daysFrom=3 (last 3 days): 2 credits/league — used by test script manually
async function fetchScoresForSport(sport, teamNames, daysFrom = 1) {
    const sportLower = sport.toLowerCase();

    if (SPORT_API_KEYS[sportLower]) {
        return fetchScores(SPORT_API_KEYS[sportLower], daysFrom);
    }

    // Soccer: check common leagues, stop as soon as we find our teams
    console.log(`[ResultsChecker] Soccer — checking common leagues (daysFrom=${daysFrom})...`);
    let allScores = [];
    for (const key of COMMON_SOCCER_LEAGUES) {
        const scores = await fetchScores(key, daysFrom);
        if (scores.length === 0) continue;
        allScores = allScores.concat(scores);
        for (const tn of teamNames) {
            for (const sc of scores) {
                if (fuzzyMatch(tn, sc.home_team) || fuzzyMatch(tn, sc.away_team)) {
                    console.log(`[ResultsChecker] Found match in ${key}!`);
                    return allScores;
                }
            }
        }
    }
    return allScores;
}

// Kept for the scheduled timer path (non-standalone)
function fetchScoresFromAPI(sport, daysFrom = 1) {
    const apiSport = SPORT_API_KEYS[sport.toLowerCase()];
    if (apiSport) {
        return fetchScores(apiSport, daysFrom);
    }
    // Soccer fallback: can't auto-discover in timer context, try common leagues
    const commonSoccer = [
        'soccer_epl', 'soccer_spain_la_liga', 'soccer_germany_bundesliga',
        'soccer_italy_serie_a', 'soccer_france_ligue_one',
        'soccer_england_championship', 'soccer_usa_mls',
        'soccer_mexico_ligamx', 'soccer_uefa_champs_league'
    ];
    return (async () => {
        let all = [];
        for (const key of commonSoccer) {
            try {
                const scores = await fetchScores(key, daysFrom);
                all = all.concat(scores);
            } catch { /* skip unavailable leagues */ }
        }
        return all;
    })();
}

// ── Clean team name: strip OCR/parser noise ──
// \b doesn't work with accented chars (í, é, etc.), so split by spaces and filter
const TEAM_NOISE_WORDS = new Set(['sí', 'si', 'no', 'yes', 'fc', 'cf', 'sc', 'ac', 'afc', 'bc']);
function cleanTeamName(name) {
    return (name || '')
        .split(/\s+/)
        .filter(w => !TEAM_NOISE_WORDS.has(w.toLowerCase()))
        .join(' ')
        .trim();
}

// ── Fuzzy match team names ──
function fuzzyMatch(str1, str2) {
    const normalize = s => cleanTeamName(s).toLowerCase().trim();
    const s1 = normalize(str1);
    const s2 = normalize(str2);

    if (!s1 || !s2) return false;
    if (s1 === s2) return true;
    if (s1.includes(s2) || s2.includes(s1)) return true;

    // Word overlap (at least 1 meaningful word in common, ignoring short words)
    const words1 = s1.split(/\s+/).filter(w => w.length > 2);
    const words2 = new Set(s2.split(/\s+/).filter(w => w.length > 2));
    const overlap = words1.filter(w => words2.has(w)).length;
    return overlap > 0;
}

// ── Find matching score for a pick ──
function findScoreForPick(pick, scores) {
    const pickTeams = pick.match.split(/\s+vs\.?\s+/i).map(t => cleanTeamName(t));

    if (pickTeams.length < 2) {
        console.log(`[ResultsChecker] Could not parse teams from match: ${pick.match}`);
        return null;
    }

    const [homeTeam, awayTeam] = pickTeams;
    console.log(`[ResultsChecker]   Looking for: "${homeTeam}" vs "${awayTeam}"`);

    for (const score of scores) {
        const homeOk = fuzzyMatch(homeTeam, score.home_team) || fuzzyMatch(awayTeam, score.home_team);
        const awayOk = fuzzyMatch(awayTeam, score.away_team) || fuzzyMatch(homeTeam, score.away_team);

        if (homeOk && awayOk) {
            return score;
        }
    }

    return null;
}

// ── Determine pick result ──
function determineResult(pick, score) {
    if (!score.completed || !score.scores || score.scores.length < 2) {
        return null;  // Game not completed
    }
    
    const homeScore = parseInt(score.scores[0].score);
    const awayScore = parseInt(score.scores[1].score);
    
    const pickLower = pick.pick.toLowerCase();
    
    // Moneyline (ML)
    if (pickLower.includes('ml')) {
        const teamInPick = pick.match.split(/\s+vs\.?\s+/i)[0].trim().toLowerCase();
        const homeTeamLower = score.home_team.toLowerCase();
        const isPickingHome = fuzzyMatch(teamInPick, homeTeamLower);
        
        const pickWins = isPickingHome ? (homeScore > awayScore) : (awayScore > homeScore);
        
        if (homeScore === awayScore) return 'Push';
        return pickWins ? 'Win' : 'Loss';
    }
    
    // Over/Under
    if (pickLower.includes('over') || pickLower.includes('under')) {
        const totalScore = homeScore + awayScore;
        const overUnderMatch = pick.pick.match(/over\s+([\d.]+)|under\s+([\d.]+)/i);
        
        if (!overUnderMatch) return null;
        
        const threshold = parseFloat(overUnderMatch[1] || overUnderMatch[2]);
        const isOver = pickLower.includes('over');
        
        const pickWins = isOver ? (totalScore > threshold) : (totalScore < threshold);
        
        if (totalScore === threshold) return 'Push';
        return pickWins ? 'Win' : 'Loss';
    }
    
    // Both Teams to Score (BTTS)
    if (pickLower.includes('btts')) {
        const bothScored = homeScore > 0 && awayScore > 0;
        return bothScored ? 'Win' : 'Loss';
    }
    
    // Default: treat as ML on first team
    const homeTeamLower = score.home_team.toLowerCase();
    const pickTeamLower = pick.match.split(/\s+vs\.?\s+/i)[0].trim().toLowerCase();
    const isPickingHome = fuzzyMatch(pickTeamLower, homeTeamLower);
    
    const pickWins = isPickingHome ? (homeScore > awayScore) : (awayScore > homeScore);
    
    if (homeScore === awayScore) return 'Push';
    return pickWins ? 'Win' : 'Loss';
}

// ── Schedule result check for a single pick ──
async function scheduleResultCheck(pick, pickId) {
    // Only schedule if we have a sport and haven't already checked
    if (!pick.sport || pick.resultChecked) {
        return;
    }
    
    const sport = pick.sport.toLowerCase();
    const delay = RESULT_CHECK_DELAYS[sport];
    
    if (!delay) {
        console.log(`[ResultsChecker] Unknown sport for pick: ${sport}`);
        return;
    }
    
    const delayMinutes = Math.round(delay / 60 / 1000);
    console.log(`[ResultsChecker] Scheduled result check for ${sport} match (${delayMinutes}min): ${pickId}`);
    
    // Cancel any existing timer for this pick
    if (activeTimers.has(pickId)) {
        clearTimeout(activeTimers.get(pickId));
    }
    
    // Schedule the check
    const timer = setTimeout(async () => {
        try {
            console.log(`[ResultsChecker] Checking result for pick: ${pickId}`);
            
            const scores = await fetchScoresFromAPI(sport);
            const matchingScore = findScoreForPick(pick, scores);
            
            if (!matchingScore) {
                console.log(`[ResultsChecker] No matching score found for: ${pick.match}`);
                activeTimers.delete(pickId);
                return;
            }
            
            if (!matchingScore.completed) {
                console.log(`[ResultsChecker] Game not completed yet: ${pick.match}`);
                // Reschedule for 5 minutes later
                scheduleResultCheck(pick, pickId);
                activeTimers.delete(pickId);
                return;
            }
            
            const result = determineResult(pick, matchingScore);
            
            if (result) {
                pick.result = result;
                pick.resultChecked = true;
                
                console.log(`[ResultsChecker] Result: ${pick.match} = ${result}`);
                
                // Update Excel with result
                try {
                    await excel.updatePickResult(pick, result);
                } catch (err) {
                    console.error(`[ResultsChecker] Error updating Excel:`, err.message);
                }
            }
            
            activeTimers.delete(pickId);
        } catch (err) {
            console.error(`[ResultsChecker] Error checking result:`, err.message);
            activeTimers.delete(pickId);
        }
    }, delay);
    
    activeTimers.set(pickId, timer);
}

// ── Cancel pending check ──
function cancelResultCheck(pickId) {
    if (activeTimers.has(pickId)) {
        clearTimeout(activeTimers.get(pickId));
        activeTimers.delete(pickId);
        console.log(`[ResultsChecker] Cancelled pending check for: ${pickId}`);
    }
}

// ── Cancel all pending checks (on shutdown) ──
function cancelAllChecks() {
    console.log(`[ResultsChecker] Cancelling ${activeTimers.size} pending result checks`);
    for (const timer of activeTimers.values()) {
        clearTimeout(timer);
    }
    activeTimers.clear();
}

// ── Standalone test runner ──
// Usage:
//   node results-checker.js              → today's completed games, update Excel
//   node results-checker.js --dry-run    → today's completed, show results, DON'T write
//   node results-checker.js --days3      → last 3 days completed, update Excel
async function runStandalone() {
    const args = process.argv.slice(2);
    const dryRun = args.includes('--dry-run');
    const daysFrom = args.includes('--days3') ? 3 : args.includes('--days2') ? 2 : 1;

    if (!config.ODDS_API_KEY) {
        console.error('[ResultsChecker] ODDS_API_KEY missing in .env');
        process.exit(1);
    }
    const ExcelJS = require('exceljs');
    const path = require('path');
    const EXCEL_FILE = path.join(process.cwd(), 'picks.xlsx');
    const fs = require('fs');
    if (!fs.existsSync(EXCEL_FILE)) {
        console.error('[ResultsChecker] picks.xlsx not found');
        process.exit(1);
    }

    if (dryRun) console.log('[ResultsChecker] *** DRY RUN – will NOT write to Excel ***');

    console.log('[ResultsChecker] Loading picks.xlsx (3 tipster tabs + General + Summary)...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const TIPSTERS = ['Abuelo', 'Cristian Rey', 'Roberto Rey'];
    console.log('[ResultsChecker] Scanning for picks with empty Result column...');

    const picksToCheck = [];
    for (const tipster of TIPSTERS) {
        const sheet = workbook.getWorksheet(tipster);
        if (!sheet) continue;
        sheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rawDate = row.getCell(1).value;
            const date = rawDate instanceof Date
                ? rawDate.getDate() + ' ' + rawDate.toLocaleString('en', { month: 'short' })
                : (rawDate || '').toString().trim();
            const sport = (row.getCell(2).value || '').toString().trim();
            const match = (row.getCell(3).value || '').toString().trim();
            const resultCell = row.getCell(7).value;
            if (!date || !match) return;
            const hasResult = resultCell !== undefined && resultCell !== null && String(resultCell).trim() !== '';
            if (hasResult) return;
            const pickText = (row.getCell(4).value || '').toString().trim();
            picksToCheck.push({ tipster, date, match, pick: pickText, sport: (sport || 'soccer').toLowerCase() });
            console.log(`   → ${tipster} | ${date} | ${sport || 'soccer'} | ${match} | ${pickText}`);
        });
    }

    if (picksToCheck.length === 0) {
        console.log('[ResultsChecker] No picks without results. Nothing to do.');
        return;
    }
    console.log(`[ResultsChecker] Found ${picksToCheck.length} pick(s) without result.\n`);

    // Group by sport
    const bySport = {};
    for (const pick of picksToCheck) {
        const s = pick.sport || 'soccer';
        if (!bySport[s]) bySport[s] = [];
        bySport[s].push(pick);
    }
    const sportList = Object.keys(bySport);

    let updated = 0;
    let paidCalls = 0;

    for (const sport of sportList) {
        const picks = bySport[sport];
        console.log(`\n[ResultsChecker] ── ${sport.toUpperCase()} (${picks.length} pick(s)) ──`);

        // Extract all team names from picks (for soccer league auto-discovery)
        const allTeamNames = [];
        for (const p of picks) {
            const teams = p.match.split(/\s+vs\.?\s+/i).map(t => cleanTeamName(t));
            allTeamNames.push(...teams);
        }

        let scores;
        try {
            scores = await fetchScoresForSport(sport, allTeamNames, daysFrom);
        } catch (err) {
            console.error(`[ResultsChecker] API error for ${sport}:`, err.message);
            continue;
        }

        // Show completed games so you can verify API data
        const completed = scores.filter(g => g.completed);
        console.log(`[ResultsChecker] ${scores.length} game(s) returned, ${completed.length} completed.`);
        for (const g of completed) {
            const s0 = g.scores?.[0]; const s1 = g.scores?.[1];
            console.log(`   ${g.home_team} ${s0?.score ?? '?'} - ${s1?.score ?? '?'} ${g.away_team}`);
        }

        for (const pick of picks) {
            console.log(`\n[ResultsChecker] Checking: ${pick.tipster} | ${pick.match} | Pick: ${pick.pick}`);
            const matchingScore = findScoreForPick(pick, scores);
            if (!matchingScore) {
                console.log(`[ResultsChecker]   No matching game in API for: ${pick.match}`);
                continue;
            }
            console.log(`[ResultsChecker]   Matched: ${matchingScore.home_team} vs ${matchingScore.away_team} (completed=${matchingScore.completed})`);
            if (!matchingScore.completed) {
                console.log(`[ResultsChecker]   Game not completed yet`);
                continue;
            }
            const result = determineResult(pick, matchingScore);
            if (result) {
                console.log(`[ResultsChecker]   Result: ${result}`);
                if (!dryRun) {
                    await excel.updatePickResult(pick, result);
                    console.log(`[ResultsChecker]   Written to Excel`);
                } else {
                    console.log(`[ResultsChecker]   (dry-run, skipped Excel write)`);
                }
                updated++;
            } else {
                console.log(`[ResultsChecker]   Could not evaluate: ${pick.pick}`);
            }
        }
    }
    console.log(`\n[ResultsChecker] Done. Results found: ${updated} of ${picksToCheck.length}.`);
    if (dryRun) console.log('[ResultsChecker] (dry-run – no changes were saved)');
}

if (require.main === module) {
    runStandalone().catch(err => {
        console.error('[ResultsChecker]', err);
        process.exit(1);
    });
}

module.exports = {
    scheduleResultCheck,
    cancelResultCheck,
    cancelAllChecks,
    determineResult,
    findScoreForPick,
    RESULT_CHECK_DELAYS
};
