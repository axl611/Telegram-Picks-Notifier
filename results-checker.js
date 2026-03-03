// results-checker.js
// Monitors picks and checks their results at the right time based on sport
// Sport-specific delays:
// - Soccer, NBA: 1h 20min (80 minutes)
// - MLB, NFL, NHL: 3 hours

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

// ── Fetch scores from Odds API ──
function fetchScoresFromAPI(sport) {
    return new Promise((resolve, reject) => {
        const apiSport = SPORT_API_KEYS[sport.toLowerCase()] || sport;
        const url = `https://api.the-odds-api.com/v4/sports/${apiSport}/scores?daysFrom=1&apiKey=${config.ODDS_API_KEY}`;
        
        https.get(url, (res) => {
            let data = '';
            res.on('data', chunk => { data += chunk; });
            res.on('end', () => {
                try {
                    const json = JSON.parse(data);
                    console.log(`[ResultsChecker] Fetched scores for ${sport}: ${json.data?.length || 0} games`);
                    resolve(json.data || []);
                } catch (err) {
                    console.error(`[ResultsChecker] Error parsing scores for ${sport}:`, err.message);
                    reject(err);
                }
            });
        }).on('error', reject);
    });
}

// ── Fuzzy match team names ──
function fuzzyMatch(str1, str2) {
    const normalize = s => s.toLowerCase().trim();
    const s1 = normalize(str1);
    const s2 = normalize(str2);
    
    if (s1 === s2) return true;
    
    // Check if one contains the other (for partial matches like "Boston" vs "Boston Celtics")
    if (s1.includes(s2) || s2.includes(s1)) return true;
    
    // Check for word overlap (e.g., "Celtics" in both)
    const words1 = new Set(s1.split(/\s+/));
    const words2 = new Set(s2.split(/\s+/));
    const overlap = [...words1].filter(w => words2.has(w)).length;
    
    return overlap > 0;
}

// ── Find matching score for a pick ──
function findScoreForPick(pick, scores) {
    const pickTeams = pick.match.split(/\s+vs\.?\s+/i).map(t => t.trim());
    
    if (pickTeams.length < 2) {
        console.log(`[ResultsChecker] Could not parse teams from match: ${pick.match}`);
        return null;
    }
    
    const [homeTeam, awayTeam] = pickTeams;
    
    for (const score of scores) {
        const homeMatch = fuzzyMatch(homeTeam, score.home_team);
        const awayMatch = fuzzyMatch(awayTeam, score.away_team);
        
        if (homeMatch && awayMatch) {
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

module.exports = {
    scheduleResultCheck,
    cancelResultCheck,
    cancelAllChecks,
    determineResult,
    findScoreForPick,
    RESULT_CHECK_DELAYS
};
