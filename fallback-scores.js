// fallback-scores.js - 365Scores fallback for leagues not covered by the Odds API
// Free, no API key required. Searches by team name → competitor ID → recent games.
const https = require('https');

const BASE = 'https://webws.365scores.com';
const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

// In-memory caches (reset each run)
const competitorCache = new Map();   // teamName → { id, name }
const gameCache = new Map();         // competitorId:dateKey → games[]

function httpGet(url) {
    return new Promise((resolve, reject) => {
        const opts = { headers: { 'User-Agent': UA, Accept: 'application/json' } };
        https.get(url, opts, res => {
            let data = '';
            res.on('data', chunk => { data += chunk; });
            res.on('end', () => {
                try {
                    const json = JSON.parse(data);
                    resolve(json);
                } catch {
                    reject(new Error('365Scores JSON parse error'));
                }
            });
        }).on('error', reject);
    });
}

async function searchCompetitor(teamName) {
    const key = teamName.toLowerCase().trim();
    if (competitorCache.has(key)) return competitorCache.get(key);

    const url = `${BASE}/web/search/?query=${encodeURIComponent(teamName)}&langId=1`;
    const data = await httpGet(url);
    const comps = (data.competitors || []).filter(c => c.sportId === 1);

    if (comps.length === 0) {
        competitorCache.set(key, null);
        return null;
    }

    // Prefer exact-ish match (name contains the search term)
    const nameL = key;
    const best = comps.find(c => c.name.toLowerCase().includes(nameL))
              || comps.find(c => nameL.includes(c.name.toLowerCase()))
              || comps[0];

    const result = { id: best.id, name: best.name };
    competitorCache.set(key, result);
    return result;
}

// dateStr format: "2 Mar", "02 Mar", "15 Feb", etc.
function parseDateRange(dateStr) {
    const months = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
    const m = String(dateStr).match(/(\d{1,2})\s+(\w{3})/i);
    if (!m) return null;
    const day = parseInt(m[1]);
    const mon = months[m[2].toLowerCase()];
    if (mon === undefined) return null;

    const now = new Date();
    const year = now.getFullYear();
    const target = new Date(year, mon, day);

    // Search window: target day -1 to +1 (handles timezone offsets)
    const from = new Date(target); from.setDate(from.getDate() - 1);
    const to = new Date(target); to.setDate(to.getDate() + 1);

    const fmt = d => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    return { fromStr: fmt(from), toStr: fmt(to), target };
}

async function fetchGamesForCompetitor(competitorId, dateStr) {
    const cacheKey = `${competitorId}:${dateStr}`;
    if (gameCache.has(cacheKey)) return gameCache.get(cacheKey);

    const range = parseDateRange(dateStr);
    if (!range) return [];

    const url = `${BASE}/web/games/?langId=1&competitors=${competitorId}&startDate=${range.fromStr}&endDate=${range.toStr}`;
    const data = await httpGet(url);
    const games = data.games || [];
    gameCache.set(cacheKey, games);
    return games;
}

// Returns an object shaped like the Odds API score so determineResult() works unchanged:
// { home_team, away_team, completed, scores: [{name, score}, {name, score}] }
async function findMatchFallback(dateStr, homeTeamRaw, awayTeamRaw) {
    console.log(`[365Scores] Fallback lookup: "${homeTeamRaw}" vs "${awayTeamRaw}" on ${dateStr}`);

    // Search both teams
    const [homeComp, awayComp] = await Promise.all([
        searchCompetitor(homeTeamRaw),
        searchCompetitor(awayTeamRaw)
    ]);

    if (!homeComp && !awayComp) {
        console.log(`[365Scores] Neither team found in search`);
        return null;
    }

    // Try fetching games using whichever competitor we found
    const competitorId = homeComp ? homeComp.id : awayComp.id;
    const games = await fetchGamesForCompetitor(competitorId, dateStr);

    if (games.length === 0) {
        console.log(`[365Scores] No games found for competitor ${competitorId} around ${dateStr}`);
        return null;
    }

    // Find the game that involves both teams (fuzzy match)
    for (const g of games) {
        const hName = (g.homeCompetitor || {}).name || '';
        const aName = (g.awayCompetitor || {}).name || '';
        const hScore = g.homeCompetitor?.score;
        const aScore = g.awayCompetitor?.score;

        const homeMatch = fuzzy(homeTeamRaw, hName) || fuzzy(homeTeamRaw, aName);
        const awayMatch = fuzzy(awayTeamRaw, aName) || fuzzy(awayTeamRaw, hName);

        if (homeMatch && awayMatch) {
            const ended = g.statusGroup === 4 || (g.statusText || '').toLowerCase() === 'ended';
            console.log(`[365Scores] Found: ${hName} ${hScore} - ${aScore} ${aName} (${g.statusText})`);

            return {
                home_team: hName,
                away_team: aName,
                completed: ended,
                scores: [
                    { name: hName, score: String(Math.floor(hScore ?? 0)) },
                    { name: aName, score: String(Math.floor(aScore ?? 0)) }
                ]
            };
        }
    }

    console.log(`[365Scores] Games found but no match for "${homeTeamRaw}" vs "${awayTeamRaw}"`);
    return null;
}

function fuzzy(search, candidate) {
    const norm = s => (s || '').toLowerCase().replace(/\bfc\b|\bcf\b|\bsc\b/gi, '').replace(/\s+/g, ' ').trim();
    const a = norm(search);
    const b = norm(candidate);
    if (!a || !b) return false;
    if (a === b) return true;
    if (b.includes(a) || a.includes(b)) return true;
    const words = a.split(' ').filter(w => w.length > 2);
    return words.some(w => b.includes(w));
}

module.exports = { findMatchFallback };
