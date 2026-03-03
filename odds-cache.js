// odds-cache.js
// Fetches and caches upcoming matches from Odds API to validate team names.
// Minimizes API calls by caching data locally.

const https = require('https');
const config = require('./config');

let cachedMatches = null;
let cacheTimestamp = null;
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

// ── Fetch matches from Odds API ──
function fetchFromOddsAPI(sport) {
    return new Promise((resolve, reject) => {
        const url = `https://api.the-odds-api.com/v4/sports/${sport}/events?apiKey=${config.ODDS_API_KEY}`;
        https.get(url, (res) => {
            let data = '';
            res.on('data', chunk => { data += chunk; });
            res.on('end', () => {
                try {
                    const json = JSON.parse(data);
                    resolve(json.data || []);
                } catch (err) {
                    reject(err);
                }
            });
        }).on('error', reject);
    });
}

// ── Initialize cache with all sports ──
async function initCache() {
    if (cachedMatches && cacheTimestamp && (Date.now() - cacheTimestamp < CACHE_TTL_MS)) {
        console.log('[OddsCache] Cache is fresh, skipping refresh.');
        return cachedMatches;
    }

    console.log('[OddsCache] Fetching matches from Odds API...');
    const sports = ['nba', 'nfl', 'mlb', 'nhl', 'soccer'];
    const allMatches = {};

    try {
        for (const sport of sports) {
            try {
                const matches = await fetchFromOddsAPI(sport);
                allMatches[sport] = matches;
                console.log(`[OddsCache] Fetched ${matches.length} matches for ${sport.toUpperCase()}`);
            } catch (err) {
                console.error(`[OddsCache] Error fetching ${sport}:`, err.message);
                allMatches[sport] = [];
            }
        }
        cachedMatches = allMatches;
        cacheTimestamp = Date.now();
        console.log('[OddsCache] Cache initialized successfully.');
    } catch (err) {
        console.error('[OddsCache] Fatal error initializing cache:', err.message);
        cachedMatches = {};
        cacheTimestamp = Date.now();
    }

    return cachedMatches;
}

// ── Find a match by team names (fuzzy match) ──
function findMatch(teamA, teamB) {
    if (!cachedMatches) return null;

    const normTeamA = teamA.toLowerCase();
    const normTeamB = teamB.toLowerCase();

    for (const [sport, matches] of Object.entries(cachedMatches)) {
        for (const match of matches) {
            const home = (match.home_team || '').toLowerCase();
            const away = (match.away_team || '').toLowerCase();

            // Check if both teams match (in either order)
            if ((home.includes(normTeamA) || normTeamA.includes(home)) &&
                (away.includes(normTeamB) || normTeamB.includes(away))) {
                return {
                    sport,
                    homeTeam: match.home_team,
                    awayTeam: match.away_team,
                    commenceTime: match.commence_time,
                    rawMatch: match
                };
            }

            if ((away.includes(normTeamA) || normTeamA.includes(away)) &&
                (home.includes(normTeamB) || normTeamB.includes(home))) {
                return {
                    sport,
                    homeTeam: match.away_team,
                    awayTeam: match.home_team,
                    commenceTime: match.commence_time,
                    rawMatch: match
                };
            }
        }
    }

    return null;
}

// ── Validate and normalize team names ──
async function validateAndNormalizeTeams(teamA, teamB) {
    // Initialize cache if needed
    if (!cachedMatches) {
        await initCache();
    }

    const match = findMatch(teamA, teamB);
    if (match) {
        return {
            normalized: true,
            homeTeam: match.homeTeam,
            awayTeam: match.awayTeam,
            sport: match.sport,
            commenceTime: match.commenceTime
        };
    }

    // Fallback: return original names if no match found
    return {
        normalized: false,
        homeTeam: teamA,
        awayTeam: teamB,
        sport: null,
        commenceTime: null
    };
}

// ── Force cache refresh ──
async function refreshCache() {
    console.log('[OddsCache] Forcing cache refresh...');
    cacheTimestamp = 0;
    return initCache();
}

module.exports = {
    initCache,
    validateAndNormalizeTeams,
    refreshCache,
    getCachedMatches: () => cachedMatches
};
