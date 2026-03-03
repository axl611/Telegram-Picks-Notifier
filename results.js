// results.js - Polls for results every 30 minutes
// Uses 365Scores (free) first, falls back to Odds API if needed
const https = require('https');
const { findMatchFallback } = require('./fallback-scores');
const { updatePickResult, updateSummary, getPicksWithoutResults } = require('./sheets');

const BET_AMOUNT = 2000;
const POLL_INTERVAL_MS = 30 * 60 * 1000; // 30 minutes

let _telegramToken = null;
let _telegramChatId = null;
let _oddsApiKey = null;
let _pollTimer = null;

function init(telegramToken, telegramChatId, oddsApiKey) {
    _telegramToken = telegramToken;
    _telegramChatId = telegramChatId;
    _oddsApiKey = oddsApiKey;
}

// ── Telegram ──────────────────────────────────────────────────────────────────

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

// ── Odds API (fallback, costs credits) ────────────────────────────────────────

const SPORT_API_KEYS = {
    'nba': 'basketball_nba',
    'mlb': 'baseball_mlb',
    'nfl': 'americanfootball_nfl',
    'nhl': 'icehockey_nhl'
};

const COMMON_SOCCER_LEAGUES = [
    'soccer_epl', 'soccer_efl_champ', 'soccer_england_league1',
    'soccer_spain_la_liga', 'soccer_germany_bundesliga',
    'soccer_italy_serie_a', 'soccer_france_ligue_one',
    'soccer_usa_mls', 'soccer_mexico_ligamx',
    'soccer_uefa_champs_league', 'soccer_uefa_europa_league',
    'soccer_brazil_campeonato', 'soccer_argentina_primera_division'
];

function apiGet(urlPath) {
    return new Promise((resolve, reject) => {
        const req = https.request(
            { hostname: 'api.the-odds-api.com', path: urlPath, method: 'GET', headers: { Accept: 'application/json' } },
            res => {
                const remaining = res.headers['x-requests-remaining'];
                if (remaining !== undefined) {
                    console.log(`   [Results] Odds API credits remaining: ${remaining}`);
                }
                let data = '';
                res.on('data', chunk => { data += chunk; });
                res.on('end', () => {
                    try {
                        const json = JSON.parse(data);
                        resolve(Array.isArray(json) ? json : (json.data || []));
                    } catch (err) { reject(err); }
                });
            }
        );
        req.on('error', reject);
        req.end();
    });
}

function fetchScoresByKey(sportKey) {
    return apiGet(`/v4/sports/${sportKey}/scores?daysFrom=1&apiKey=${_oddsApiKey}`);
}

async function fetchScoresFromAPI(sport) {
    const directKey = SPORT_API_KEYS[sport.toLowerCase()];
    if (directKey) {
        const scores = await fetchScoresByKey(directKey);
        console.log(`   [Results] Odds API ${directKey}: ${scores.length} game(s)`);
        return scores;
    }

    console.log(`   [Results] Odds API — querying ${COMMON_SOCCER_LEAGUES.length} soccer league(s)...`);
    let allScores = [];
    for (const key of COMMON_SOCCER_LEAGUES) {
        try {
            const scores = await fetchScoresByKey(key);
            if (scores.length > 0) allScores = allScores.concat(scores);
        } catch { /* skip */ }
    }
    return allScores;
}

// ── Team name matching ────────────────────────────────────────────────────────

const TEAM_NOISE_WORDS = new Set(['sí', 'si', 'no', 'yes', 'fc', 'cf', 'sc', 'ac', 'afc', 'bc',
    'total', 'resultado', 'goles', 'ganador', 'marcador']);

function norm(str) {
    return (str || '')
        .replace(/\bel\s*abuelo\b/gi, '')
        .replace(/\babuel(o|ito)\b/gi, '')
        .replace(/\bcristian\s*rey\b/gi, '')
        .replace(/\broberto\s*rey\b/gi, '')
        .split(/\s+/)
        .map(w => w.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, ''))
        .filter(w => w && !TEAM_NOISE_WORDS.has(w))
        .join(' ')
        .trim();
}

function findScoreForPick(pick, scores) {
    const pickTeams = pick.match.split(/\s+vs\.?\s+/i).map(t => t.trim());
    if (pickTeams.length < 2) return null;

    const [homeTeam, awayTeam] = pickTeams;
    const homeNorm = norm(homeTeam);
    const awayNorm = norm(awayTeam);

    for (const score of scores) {
        const hn = norm(score.home_team);
        const an = norm(score.away_team);
        const homeOk = homeNorm.includes(hn) || hn.includes(homeNorm) || awayNorm.includes(hn) || hn.includes(awayNorm);
        const awayOk = homeNorm.includes(an) || an.includes(homeNorm) || awayNorm.includes(an) || an.includes(awayNorm);
        if (homeOk && awayOk) return score;
    }
    return null;
}

// ── Pick evaluation ───────────────────────────────────────────────────────────

function teamScore(game, teamName) {
    if (!game.scores) return null;
    const n = norm(teamName);
    const entry = game.scores.find(s => norm(s.name).includes(n) || n.includes(norm(s.name)));
    return entry ? parseInt(entry.score) : null;
}

function evaluatePick(pickText, game) {
    if (!game || !game.completed || !game.scores || game.scores.length < 2) return null;

    const homeScore = teamScore(game, game.home_team);
    const awayScore = teamScore(game, game.away_team);
    if (homeScore === null || awayScore === null) return null;

    const total = homeScore + awayScore;
    const lower = (pickText || '').toLowerCase();

    console.log(`   [Results] Score: ${game.home_team} ${homeScore} - ${awayScore} ${game.away_team} (total ${total})`);

    // BTTS
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

    // ML (default)
    const mlTeam = norm(pickText.replace(/\s*ml\s*$/i, '').trim());
    const homeN = norm(game.home_team);
    const awayN = norm(game.away_team);
    const pickedHome = homeN.includes(mlTeam) || mlTeam.includes(homeN);
    const pickedAway = awayN.includes(mlTeam) || mlTeam.includes(awayN);

    if (!pickedHome && !pickedAway) return null;
    if (homeScore === awayScore) return 'L';
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

// ── Core: check a single pick ─────────────────────────────────────────────────

async function checkSinglePick(pick) {
    const teams = pick.match.split(/\s+vs\.?\s+/i).map(t => norm(t));
    let matchingScore = null;

    // 1) Try 365Scores (free)
    if (teams.length >= 2) {
        try {
            matchingScore = await findMatchFallback(pick.date, teams[0], teams[1]);
            if (matchingScore) console.log(`   [Results] Found via 365Scores`);
        } catch {}
    }

    // 2) Fall back to Odds API if needed
    if (!matchingScore && _oddsApiKey) {
        console.log(`   [Results] Not in 365Scores — trying Odds API...`);
        try {
            const scores = await fetchScoresFromAPI(pick.sport || 'soccer');
            matchingScore = findScoreForPick(pick, scores);
        } catch (err) {
            console.log(`   [Results] Odds API error: ${err.message}`);
        }
    }

    if (!matchingScore) return null;
    if (!matchingScore.completed) {
        console.log(`   [Results] Game not completed yet: ${matchingScore.home_team} vs ${matchingScore.away_team}`);
        return null;
    }

    return { score: matchingScore, result: evaluatePick(pick.pick, matchingScore) };
}

// ── Poll cycle: check all pending picks ───────────────────────────────────────

async function pollForResults() {
    try {
        const pending = await getPicksWithoutResults();

        if (pending.length === 0) {
            console.log(`[Results] No pending picks — skipping`);
            return;
        }

        console.log(`[Results] ── Checking ${pending.length} pending pick(s) ──`);

        let updated = 0;
        for (const pick of pending) {
            console.log(`[Results] ${pick.tipster} | ${pick.match} | ${pick.pick}`);

            const outcome = await checkSinglePick(pick);
            if (!outcome || !outcome.result) {
                console.log(`   [Results] No result yet`);
                continue;
            }

            const pl = calcPL(outcome.result, pick.odds);
            console.log(`   [Results] → ${outcome.result} ($${pl})`);

            await updatePickResult(pick, outcome.result);
            updated++;

            const emoji = outcome.result === 'W' ? '✅' : outcome.result === 'L' ? '❌' : '➡️';
            await sendTelegram(
                `${emoji} *Result: ${outcome.result}*\n` +
                `*Tipster:* ${pick.tipster}\n` +
                `*Match:* ${pick.match}\n` +
                `*Pick:* ${pick.pick}\n` +
                `*Odds:* ${pick.odds}\n` +
                `*P&L:* $${pl}`
            );
        }

        if (updated > 0) {
            await updateSummary();
            console.log(`[Results] Updated ${updated} result(s)`);
        }
    } catch (err) {
        console.error(`[Results] Poll error: ${err.message}`);
    }
}

// ── Scheduler ─────────────────────────────────────────────────────────────────

function startResultsScheduler() {
    console.log(`📊 Results checker started (polls every 30 min when picks are pending)`);

    // Run first check shortly after startup (10 seconds)
    setTimeout(() => pollForResults(), 10 * 1000);

    // Then every 30 minutes
    _pollTimer = setInterval(() => pollForResults(), POLL_INTERVAL_MS);
}

function cancelAllChecks() {
    if (_pollTimer) {
        clearInterval(_pollTimer);
        _pollTimer = null;
    }
    console.log('[Results] Scheduler stopped');
}

// scheduleResultCheck is kept for compatibility with index.js — new picks will
// be picked up by the next 30-min poll cycle automatically
function scheduleResultCheck() {}

module.exports = { init, startResultsScheduler, cancelAllChecks, scheduleResultCheck };
