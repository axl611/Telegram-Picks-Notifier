#!/usr/bin/env node
// test-results-today.js - Force API calls to verify results for TODAY's matches only
// Usage: node test-results-today.js

const ExcelJS = require('exceljs');
const https = require('https');
const path = require('path');
const config = require('./config');

const EXCEL_FILE = path.join(process.cwd(), 'picks.xlsx');
const BET_AMOUNT = 2000;
const TIPSTERS = ['Abuelo', 'Cristian Rey', 'Roberto Rey'];

// Sport-specific delays (not used in this test - we check immediately)
const SPORT_API_KEYS = {
    'soccer': 'soccer',
    'nba': 'basketball_nba',
    'mlb': 'baseball_mlb',
    'nfl': 'americanfootball_nfl',
    'nhl': 'icehockey_nhl'
};

// Get today's date in "Day Month" format (e.g., "2 Mar")
function getTodayFormatted() {
    const now = new Date();
    return now.getDate() + ' ' + now.toLocaleString('en', { month: 'short' });
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

// Fetch scores from Odds API
function fetchScoresFromAPI(sport) {
    return new Promise((resolve, reject) => {
        const apiSport = SPORT_API_KEYS[sport.toLowerCase()] || sport;
        const url = `/v4/sports/${apiSport}/scores?daysFrom=1&apiKey=${config.ODDS_API_KEY}`;
        
        const req = https.request(
            { hostname: 'api.the-odds-api.com', path: url, method: 'GET', headers: { Accept: 'application/json' } },
            res => {
                let data = '';
                res.on('data', chunk => { data += chunk; });
                res.on('end', () => {
                    try {
                        const json = JSON.parse(data);
                        resolve(json.data || []);
                    } catch (err) {
                        console.error(`❌ Error parsing scores for ${sport}:`, err.message);
                        reject(err);
                    }
                });
            }
        );
        req.on('error', reject);
        req.end();
    });
}

// Find matching score for a pick
function findScoreForPick(pick, scores) {
    const pickTeams = pick.match.split(/\s+vs\.?\s+/i).map(t => t.trim());
    
    if (pickTeams.length < 2) {
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

// Main test function
async function testTodaysResults() {
    const todayFormatted = getTodayFormatted();
    console.log(`\n🧪 TEST: Checking results for TODAY (${todayFormatted})`);
    console.log('================================================\n');

    try {
        // Load workbook
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);

        let totalPicks = 0;
        let resultsFound = 0;
        let resultsFailed = 0;

        // Check each tipster sheet
        for (const tipster of TIPSTERS) {
            const sheet = workbook.getWorksheet(tipster);
            if (!sheet) continue;

            const tipstersToday = [];

            // Collect all picks for today
            sheet.eachRow((row, rowNum) => {
                if (rowNum === 1) return;
                
                const date = (row.getCell(1).value || '').toString().trim();
                const sport = (row.getCell(2).value || '').toString().trim();
                const match = (row.getCell(3).value || '').toString().trim();
                const pick = (row.getCell(4).value || '').toString().trim();
                const odds = (row.getCell(5).value || '').toString().trim();
                const result = (row.getCell(7).value || '').toString().trim();

                // Only process today's picks that don't have a result yet
                if (date === todayFormatted && !result && match && sport) {
                    tipstersToday.push({
                        rowNum,
                        tipster,
                        date,
                        sport,
                        match,
                        pick,
                        odds,
                        result
                    });
                    totalPicks++;
                }
            });

            if (tipstersToday.length > 0) {
                console.log(`📋 ${tipster}: ${tipstersToday.length} pick(s) today`);

                // Group picks by sport and fetch scores
                const sportMap = {};
                for (const p of tipstersToday) {
                    if (!sportMap[p.sport]) sportMap[p.sport] = [];
                    sportMap[p.sport].push(p);
                }

                // Check each sport
                for (const [sport, picks] of Object.entries(sportMap)) {
                    console.log(`   Fetching ${sport.toUpperCase()} scores...`);
                    
                    try {
                        const scores = await fetchScoresFromAPI(sport);
                        
                        if (scores.length === 0) {
                            console.log(`   ⚠️  No scores found for ${sport}`);
                            continue;
                        }

                        console.log(`   ✅ Fetched ${scores.length} games from API`);

                        // Check each pick
                        for (const p of picks) {
                            const matchingScore = findScoreForPick(p, scores);

                            if (!matchingScore) {
                                console.log(`      ❌ No matching score: ${p.match}`);
                                resultsFailed++;
                                continue;
                            }

                            if (!matchingScore.completed) {
                                console.log(`      ⏳ Game not completed: ${p.match}`);
                                resultsFailed++;
                                continue;
                            }

                            const resultLetter = evaluatePick(p.pick, matchingScore);

                            if (!resultLetter) {
                                console.log(`      ❓ Could not evaluate: ${p.match} | ${p.pick}`);
                                resultsFailed++;
                                continue;
                            }

                            // Calculate profit/loss
                            const o = parseFloat(p.odds) || 0;
                            let pl = 0;
                            if (resultLetter === 'W') {
                                pl = (o - 1) * BET_AMOUNT;
                            } else if (resultLetter === 'L') {
                                pl = -BET_AMOUNT;
                            }

                            // Update the cell with result
                            const row = sheet.getRow(p.rowNum);
                            row.getCell(7).value = resultLetter;
                            
                            // Color the result cell
                            const colors = { 'W': 'FF92D050', 'L': 'FFFF6666', 'P': 'FFFFFF00' };
                            row.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[resultLetter] } };

                            const emoji = resultLetter === 'W' ? '✅' : resultLetter === 'L' ? '❌' : '➡️';
                            console.log(`      ${emoji} ${p.match} | ${p.pick} → ${resultLetter} ($${pl})`);
                            resultsFound++;
                        }
                    } catch (err) {
                        console.log(`   ❌ Error fetching ${sport} scores:`, err.message);
                    }
                }
            }
        }

        // Save the workbook
        if (resultsFound > 0) {
            await workbook.xlsx.writeFile(EXCEL_FILE);
            console.log(`\n✅ Saved ${resultsFound} result(s) to Excel`);
        }

        console.log(`\n📊 Summary:`);
        console.log(`   Total picks today: ${totalPicks}`);
        console.log(`   Results found & updated: ${resultsFound}`);
        console.log(`   Could not process: ${resultsFailed}`);
        console.log(`   Still pending: ${totalPicks - resultsFound - resultsFailed}`);
        console.log('');

    } catch (err) {
        console.error('❌ Error:', err.message);
        process.exit(1);
    }
}

// Run the test
testTodaysResults().catch(err => {
    console.error('❌ Fatal error:', err);
    process.exit(1);
});
