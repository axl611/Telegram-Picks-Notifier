// parser.js - Pick parser with Odds API team validation
const { validateAndNormalizeTeams } = require('./odds-cache');

const SPORT_TEAMS = {
    NHL: ['flames', 'ducks', 'bruins', 'sabres', 'canadiens', 'senators', 'leafs', 'hurricanes', 'panthers', 'lightning', 'capitals', 'rangers', 'islanders', 'flyers', 'penguins', 'blues', 'blackhawks', 'avalanche', 'stars', 'wild', 'predators', 'jets', 'oilers', 'canucks', 'sharks', 'kings', 'coyotes', 'golden knights', 'kraken'],
    NBA: ['lakers', 'celtics', 'warriors', 'bulls', 'heat', 'bucks', 'nets', 'knicks', 'sixers', 'raptors', 'hawks', 'hornets', 'pistons', 'pacers', 'cavaliers', 'magic', 'wizards', 'nuggets', 'timberwolves', 'thunder', 'blazers', 'jazz', 'suns', 'kings', 'clippers', 'spurs', 'rockets', 'grizzlies', 'pelicans', 'mavericks'],
    MLB: ['yankees', 'red sox', 'dodgers', 'giants', 'cubs', 'cardinals', 'braves', 'mets', 'phillies', 'nationals', 'marlins', 'brewers', 'reds', 'pirates', 'padres', 'rockies', 'diamondbacks', 'astros', 'rangers', 'angels', 'athletics', 'mariners', 'white sox', 'tigers', 'guardians', 'twins', 'royals', 'rays', 'blue jays', 'orioles'],
    NFL: ['patriots', 'bills', 'dolphins', 'jets', 'ravens', 'bengals', 'browns', 'steelers', 'titans', 'colts', 'jaguars', 'texans', 'chiefs', 'raiders', 'chargers', 'broncos', 'cowboys', 'giants', 'eagles', 'commanders', 'bears', 'lions', 'packers', 'vikings', 'falcons', 'panthers', 'saints', 'buccaneers', 'rams', 'seahawks', '49ers', 'cardinals'],
    UFC: ['ufc', 'mma', 'fight', 'bout', 'pelea'],
    SOCCER: ['fc', 'united', 'city', 'athletic', 'atletico', 'real', 'barcelona', 'leon', 'necaxa', 'monterrey', 'cruz azul', 'america', 'chivas', 'tigres', 'pumas', 'toluca', 'santos', 'pachuca', 'queretaro', 'inter miami', 'la galaxy', 'seattle', 'portland', 'atlanta', 'orlando', 'nashville', 'austin', 'charlotte', 'liverpool', 'arsenal', 'chelsea', 'tottenham', 'manchester', 'juventus', 'milan', 'roma', 'napoli', 'bayern', 'dortmund', 'ajax', 'psv', 'porto', 'benfica', 'sevilla', 'valencia', 'getafe', 'club']
};

function detectSport(matchText) {
    if (!matchText) return 'SOCCER';
    const lower = matchText.toLowerCase();
    for (const [sport, teams] of Object.entries(SPORT_TEAMS)) {
        if (sport === 'SOCCER') continue;
        if (teams.some(t => lower.includes(t))) return sport;
    }
    if (SPORT_TEAMS.UFC.some(t => lower.includes(t))) return 'UFC';
    return 'SOCCER';
}

// ── Find two teams by matching known team names ──
function findTeamsInText(text) {
    const lower = text.toLowerCase();
    const foundTeams = [];

    // Build a list of all known teams with their display names
    const allTeams = [];
    for (const [sport, teams] of Object.entries(SPORT_TEAMS)) {
        for (const team of teams) {
            allTeams.push({ name: team, sport });
        }
    }

    // Sort by length descending so longer team names match first (e.g., 'golden knights' before 'knights')
    allTeams.sort((a, b) => b.name.length - a.name.length);

    // Find all matches in the text with their positions
    const matches = [];
    for (const team of allTeams) {
        const regex = new RegExp(`\\b${team.name}\\b`, 'gi');
        let m;
        while ((m = regex.exec(lower)) !== null) {
            matches.push({ name: team.name, pos: m.index, sport: team.sport });
        }
    }

    // Sort by position and pick the first two distinct teams
    matches.sort((a, b) => a.pos - b.pos);
    const seen = new Set();
    for (const match of matches) {
        if (!seen.has(match.name)) {
            foundTeams.push(match.name);
            seen.add(match.name);
            if (foundTeams.length === 2) break;
        }
    }

    return foundTeams.length === 2 ? foundTeams : null;
}

function extractOdds(text) {
    if (!text) return null;

    // Decimal odds (e.g. 1.55, 2.10)
    const decMatches = text.match(/\b(\d{1,2}\.\d{2,3})\b/g);
    if (decMatches) {
        for (const m of decMatches) {
            const num = parseFloat(m);
            if (num >= 1.10 && num <= 15.00) return num;
        }
    }

    // American odds (e.g. -125, +150, -110)
    const amMatch = text.match(/([+-]\d{3,4})\b/);
    if (amMatch) {
        const american = parseInt(amMatch[1]);
        if (american > 0) return parseFloat(((american / 100) + 1).toFixed(3));
        if (american < 0) return parseFloat(((100 / Math.abs(american)) + 1).toFixed(3));
    }

    return null;
}

function cleanOCR(text) {
    return text
        // Remove Twitter watermark (but be careful not to remove @ between team names)
        // Target: "Twitter @...", "Twi @...", "@PicksPra", "@Picks Pra", etc. at start/end or with clear boundaries
        .replace(/Twitter\s+@\w+/gi, '')
        .replace(/^@\w+/gm, '')  // @ at line start
        .replace(/@\w+\s*$/gm, '') // @ at line end
        .replace(/\bTwi\s*@\w+/gi, '')  // Twi @...
        .replace(/\bTwitte\b/gi, '')  // Twitte/Twitter fragments
        // Remove other noise
        .replace(/Tipster\s*Costo/gi, '')
        .replace(/Costo/gi, '')
        .replace(/\$\d+\.?\d*\s*MXN/gi, '')
        .replace(/ABIERTO/gi, '')
        .replace(/Tipster/gi, '')
        .replace(/\s{2,}/g, ' ')
        .trim();
    // Note: do NOT remove SGP — Google Vision reads it correctly
}

function extractDate(text) {
    const match = text.match(/(\d{1,2})\/(\d{2})/);
    if (match) {
        const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return match[1] + ' ' + (months[parseInt(match[2]) - 1] || match[2]);
    }
    return null;
}

function extractTipsterFromOCR(text) {
    const tipsters = [
        { pattern: /roberto\s*rey|roberto/i,       name: 'Roberto Rey',  key: 'roberto'  },
        { pattern: /cristian\s*rey|cristian/i,      name: 'Cristian Rey', key: 'cristian' },
        { pattern: /el\s*abuelo|abuelo|abuelito/i,  name: 'El Abuelo',    key: 'abuelo'   }
    ];
    for (const t of tipsters) {
        if (t.pattern.test(text)) return t;
    }
    return null;
}

// ── Build SGP pick string — collects all legs, deduplicates ──
function buildSGPPick(text, match) {
    const legs = [];
    const seen = new Set();

    const addLeg = (leg) => {
        const key = leg.toLowerCase().replace(/\s+/g, '');
        if (!seen.has(key)) { seen.add(key); legs.push(leg); }
    };

    // ML leg — "1x2 PA" present means pick the first team to win
    if (/1x2|PA/i.test(text)) {
        const parts = (match || '').split(/\s+vs\.?\s+/i);
        if (parts.length > 0) addLeg(parts[0].trim() + ' ML');
    }

    // Over legs — skip if line is prefixed by "total de goles" context (player total)
    // Deduplicated by line value
    const overLines = new Set();
    for (const m of text.matchAll(/m[aá]s de (\d+\.?\d*)/gi)) {
        if (!overLines.has(m[1])) { overLines.add(m[1]); addLeg('Over ' + m[1]); }
    }

    // Under legs — deduplicated by line value
    const underLines = new Set();
    for (const m of text.matchAll(/menos de (\d+\.?\d*)/gi)) {
        if (!underLines.has(m[1])) { underLines.add(m[1]); addLeg('Under ' + m[1]); }
    }

    // BTTS leg
    if (/ambos equipos|ambos anotan|btts/i.test(text)) addLeg('BTTS Yes');

    return legs.length > 0 ? 'SGP: ' + legs.join(' + ') : 'SGP';
}

// ── Build pick text — SGP checked FIRST ──
function buildPickText(text, match) {
    if (!text) return match ? match.split(/\s+vs\.?\s+/i)[0].trim() + ' ML' : 'ML';

    const lower = text.toLowerCase();

    // SGP — Google Vision reads the red badge correctly as "SGP"
    // Also handle "crear apuesta" as fallback
    if (/\bSGP\b/.test(text) || /crear apuesta|misma apuesta/i.test(lower)) {
        return buildSGPPick(text, match);
    }

    // BTTS
    if (/ambos equipos|mbos equipos|ambos anotan|btts/i.test(lower)) {
        return 'Ambos equipos marcan';
    }

    // Over
    const overMatch = lower.match(/m[aá]s de (\d+\.?\d*)|over (\d+\.?\d*)/i);
    if (overMatch) return 'Over ' + (overMatch[1] || overMatch[2]);

    // Under
    const underMatch = lower.match(/menos de (\d+\.?\d*)|under (\d+\.?\d*)/i);
    if (underMatch) return 'Under ' + (underMatch[1] || underMatch[2]);

    // Asian Handicap
    if (/handicap|hándicap|\bah\b|asian/i.test(lower)) {
        const ahMatch = lower.match(/[+-]?\d+\.?\d*/);
        return ahMatch ? 'AH ' + ahMatch[0] : 'AH';
    }

    // ML via "1x2 PA" indicator
    if (/1x2|PA/i.test(text)) {
        const parts = (match || '').split(/\s+vs\.?\s+/i);
        return parts.length > 0 ? parts[0].trim() + ' ML' : 'ML';
    }

    // Default: ML on first team
    if (match) {
        const parts = match.split(/\s+vs\.?\s+/i);
        return parts[0].trim() + ' ML';
    }

    return 'ML';
}

// ── Main parser ──
async function parsePicks(ocrText, tipsterFromCaption) {
    const cleaned = cleanOCR(ocrText);
    const fullText = cleaned.replace(/\n/g, ' ');
    const picks = [];

    const tipsterFromOCR = extractTipsterFromOCR(fullText);
    const tipster = tipsterFromCaption || tipsterFromOCR;

    const date = extractDate(fullText);
    const odds = extractOdds(fullText);
    console.log('   [Parser] Odds:', odds);

    // Strip tipster names, pick phrases, and betting terms BEFORE extracting teams
    const matchText = fullText
        // Tipster names
        .replace(/\bel\s*abuelo\b/gi, '')
        .replace(/\babuel(o|ito)\b/gi, '')
        .replace(/\bcristian\s*rey\b/gi, '')
        .replace(/\broberto\s*rey\b/gi, '')
        // Pick phrases
        .replace(/ambos equipos marcan[:\s]*(sí|si|no|yes)?/gi, '')
        .replace(/ambos anotan[:\s]*(sí|si|no|yes)?/gi, '')
        .replace(/mbos equipos[:\s]*(sí|si|no|yes)?/gi, '')
        .replace(/btts[:\s]*(sí|si|no|yes)?/gi, '')
        .replace(/m[aá]s de \d+\.?\d*/gi, '')
        .replace(/menos de \d+\.?\d*/gi, '')
        .replace(/over \d+\.?\d*/gi, '')
        .replace(/under \d+\.?\d*/gi, '')
        .replace(/handicap[:\s]*[+-]?\d+\.?\d*/gi, '')
        .replace(/hándicap[:\s]*[+-]?\d+\.?\d*/gi, '')
        .replace(/\b(1x2|PA|ML|OU|AH|BTTS|SGP)\b/gi, '')
        .replace(/crear apuesta|misma apuesta/gi, '')
        // Betting terms that leak into team names
        .replace(/\btotal\s*(de\s*(goles|esquinas|tarjetas|puntos|sets))?\b/gi, '')
        .replace(/\bresultado\b/gi, '')
        .replace(/\bgoles\b/gi, '')
        .replace(/\bganador\b/gi, '')
        .replace(/\bmarcador\b/gi, '')
        .replace(/\s{2,}/g, ' ')
        .trim();

    // Extract match
    let match = null;
    const garbage = /\b(Twit|Twitter|pe|pra|ar|CES|FE|LW|ET|ca)\b/gi;

    // Try "Team A vs. Team B" first (also accept @ or other OCR artifacts)
    const vsMatch = matchText.match(/([A-Za-záéíóúÁÉÍÓÚñÑ]+(?:\s+[A-Za-záéíóúÁÉÍÓÚñÑ]+){0,3})\s*(?:vs\.?|@|,)\s*([A-Za-záéíóúÁÉÍÓÚñÑ]+(?:\s+[A-Za-záéíóúÁÉÍÓÚñÑ]+){0,3})/i);
    if (vsMatch) {
        const teamA = vsMatch[1].replace(garbage, '').trim();
        const teamB = vsMatch[2].replace(garbage, '').trim();
        if (teamA.length > 1 && teamB.length > 1) match = teamA + ' vs. ' + teamB;
    }

    // Fallback: "ABC Team1 XYZ Team2" — two team-like tokens separated by known sport keywords
    // e.g. "BOS Celtics MIL Bucks" -> "BOS Celtics vs. MIL Bucks"
    if (!match) {
        const slipMatch = matchText.match(/\b([A-Z]{2,3}\s+[A-Za-z]+)\s+([A-Z]{2,3}\s+[A-Za-z]+)\b/);
        if (slipMatch) {
            const teamA = slipMatch[1].trim();
            const teamB = slipMatch[2].trim();
            if (teamA.length > 3 && teamB.length > 3) match = teamA + ' vs. ' + teamB;
        }
    }

    // Final fallback: Find any two known teams in the text, assume they're playing each other
    if (!match) {
        const foundTeams = findTeamsInText(matchText);
        if (foundTeams && foundTeams.length === 2) {
            // Validate and normalize team names via Odds API cache
            const validation = await validateAndNormalizeTeams(foundTeams[0], foundTeams[1]);
            if (validation.normalized) {
                match = validation.homeTeam + ' vs. ' + validation.awayTeam;
                console.log('   [Parser] Normalized via Odds API:', match);
            } else {
                match = foundTeams[0] + ' vs. ' + foundTeams[1];
                console.log('   [Parser] Found teams by keyword match (not in cache):', match);
            }
        }
    }

    const pickText = buildPickText(fullText, match);
    console.log('   [Parser] Pick:', pickText);
    console.log('   [Parser] Match:', match);
    console.log('   [Parser] Tipster:', tipster ? (typeof tipster === 'string' ? tipster : tipster.name) : tipsterFromCaption);

    if (!match) {
        console.log('   [Parser] No valid match found, skipping.');
        return picks;
    }

    // Resolve tipster name — image OCR takes priority over caption
    let tipsterName = 'Unknown';
    if (tipster) {
        // tipster can be a string (from caption) or an object (from OCR)
        if (typeof tipster === 'string') {
            tipsterName = tipster;
        } else if (tipster.name) {
            tipsterName = tipster.name;
        }
    } else if (tipsterFromCaption) {
        if (typeof tipsterFromCaption === 'string') {
            tipsterName = tipsterFromCaption;
        } else if (tipsterFromCaption.name) {
            tipsterName = tipsterFromCaption.name;
        }
    }

    // Normalize tipster name to match Excel sheet names
    const tLower = tipsterName.toLowerCase();
    if (tLower.includes('abuelo'))        tipsterName = 'Abuelo';
    else if (tLower.includes('cristian')) tipsterName = 'Cristian Rey';
    else if (tLower.includes('roberto') || tLower.includes('beto')) tipsterName = 'Roberto Rey';

    picks.push({
        tipster: tipsterName,
        tipsterKey: tipster ? tipster.key : null,
        sport: detectSport(match),
        match,
        pick: pickText,
        odds,
        date: date || getTodayFormatted(),
        raw: fullText
    });

    return picks;
}

function getTodayFormatted() {
    const now = new Date();
    return now.getDate() + ' ' + now.toLocaleString('en', { month: 'short' });
}

module.exports = { parsePicks, extractTipsterFromOCR, detectSport };