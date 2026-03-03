// parser.js - Pick parser for Phase 2
const ExcelJS = require('exceljs');
const path = require('path');

// ── Sport detection from team names ──
const SPORT_TEAMS = {
    NHL: ['flames', 'ducks', 'bruins', 'sabres', 'canadiens', 'senators', 'leafs', 'hurricanes', 'panthers', 'lightning', 'capitals', 'rangers', 'islanders', 'flyers', 'penguins', 'blues', 'blackhawks', 'avalanche', 'stars', 'wild', 'predators', 'jets', 'oilers', 'canucks', 'sharks', 'kings', 'coyotes', 'golden knights', 'kraken'],
    NBA: ['lakers', 'celtics', 'warriors', 'bulls', 'heat', 'bucks', 'nets', 'knicks', 'sixers', 'raptors', 'hawks', 'hornets', 'pistons', 'pacers', 'cavaliers', 'magic', 'wizards', 'nuggets', 'timberwolves', 'thunder', 'blazers', 'jazz', 'suns', 'kings', 'clippers', 'spurs', 'rockets', 'grizzlies', 'pelicans', 'mavericks'],
    MLB: ['yankees', 'red sox', 'dodgers', 'giants', 'cubs', 'cardinals', 'braves', 'mets', 'phillies', 'nationals', 'marlins', 'brewers', 'reds', 'pirates', 'padres', 'rockies', 'diamondbacks', 'astros', 'rangers', 'angels', 'athletics', 'mariners', 'white sox', 'tigers', 'guardians', 'twins', 'royals', 'rays', 'blue jays', 'orioles'],
    NFL: ['patriots', 'bills', 'dolphins', 'jets', 'ravens', 'bengals', 'browns', 'steelers', 'titans', 'colts', 'jaguars', 'texans', 'chiefs', 'raiders', 'chargers', 'broncos', 'cowboys', 'giants', 'eagles', 'commanders', 'bears', 'lions', 'packers', 'vikings', 'falcons', 'panthers', 'saints', 'buccaneers', 'rams', 'seahawks', '49ers', 'cardinals'],
    UFC: ['ufc', 'mma', 'fight', 'bout', 'pelea'],
    SOCCER: ['fc', 'united', 'city', 'athletic', 'atletico', 'real', 'barcelona', 'leon', 'necaxa', 'monterrey', 'cruz azul', 'america', 'chivas', 'tigres', 'pumas', 'toluca', 'santos', 'pachuca', 'queretaro', 'inter miami', 'la galaxy', 'seattle', 'portland', 'atlanta', 'orlando', 'nashville', 'austin', 'charlotte', 'liverpool', 'arsenal', 'chelsea', 'tottenham', 'manchester', 'juventus', 'milan', 'roma', 'napoli', 'bayern', 'dortmund', 'ajax', 'psv', 'porto', 'benfica', 'sevilla', 'valencia', 'club']
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

// ── Odds extraction - find decimal number like 1.55, 2.10 ──
function extractOdds(text) {
    if (!text) return null;
    
    const decimalPattern = /\b(\d{1,2}\.\d{2,3})\b/g;
    const matches = text.match(decimalPattern);
    
    if (matches) {
        for (const match of matches) {
            const num = parseFloat(match);
            if (num >= 1.10 && num <= 15.00) {
                return num;
            }
        }
    }
    
    return null;
}

// ── Build combined Pick text ──
function buildPickText(text, match) {
    if (!text) return match || 'ML';
    
    const lower = text.toLowerCase();
    
    // BTTS - fix partial OCR "mbos equipos" -> "Ambos equipos marcan"
    if (/ambos equipos|mbos equipos|ambos anotan|btts/i.test(lower)) {
        return 'Ambos equipos marcan';
    }
    
    // Totals
    const overMatch = lower.match(/mas de (\d+\.?\d*)|más de (\d+\.?\d*)|over (\d+\.?\d*)/i);
    if (overMatch) {
        const line = overMatch[1] || overMatch[2] || overMatch[3];
        return 'Over ' + line;
    }
    
    const underMatch = lower.match(/menos de (\d+\.?\d*)|under (\d+\.?\d*)/i);
    if (underMatch) {
        const line = underMatch[1] || underMatch[2];
        return 'Under ' + line;
    }
    
    // Handicap
    if (/handicap|hándicap|ah|asian/i.test(lower)) {
        const ahMatch = lower.match(/[+-]?\d+\.?\d*/);
        if (ahMatch) return 'AH ' + ahMatch[0];
        return 'AH';
    }
    
    // Check for "crear apuesta" or combo picks
    if (/crear apuesta|sgm|misma apuesta/i.test(lower)) {
        // Extract the combo info
        return text.trim();
    }
    
    // Default to ML with team name if available
    if (match) {
        // Extract first team from match
        const teams = match.split(/\s+vs\.?\s+/i);
        if (teams.length > 0) {
            return teams[0].trim() + ' ML';
        }
    }
    
    return 'ML';
}

// ── Clean OCR noise ──
function cleanOCR(text) {
    return text
        .replace(/Twitter\s*@PicksPra/gi, '')
        .replace(/@PicksPra/gi, '')
        .replace(/Tipster\s*Costo/gi, '')
        .replace(/Costo/gi, '')
        .replace(/\$\d+\.?\d*\s*MXN/gi, '')
        .replace(/ABIERTO/gi, '')
        .replace(/Tipster/gi, '')
        .replace(/\s{2,}/g, ' ')
        .trim();
}

// ── Extract date ──
function extractDate(text) {
    const match = text.match(/(\d{1,2})\/(\d{2})/);
    if (match) {
        const day = match[1];
        const month = match[2];
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const monthName = months[parseInt(month) - 1];
        return day + ' ' + monthName;
    }
    return null;
}

// ── Extract tipster from OCR ──
function extractTipsterFromOCR(text) {
    const tipsters = [
        { pattern: /roberto\s*rey|roberto/i, name: 'Roberto Rey', key: 'roberto' },
        { pattern: /cristian\s*rey|cristian/i, name: 'Cristian Rey', key: 'cristian' },
        { pattern: /el\s*abuelo|abuelo|abuelito/i, name: 'El Abuelo', key: 'abuelo' }
    ];
    for (const tipster of tipsters) {
        if (tipster.pattern.test(text)) return tipster;
    }
    return null;
}

// ── Main parser ──
function parsePicks(ocrText, tipsterFromCaption) {
    const cleaned = cleanOCR(ocrText);
    const fullText = cleaned.replace(/\n/g, ' ');

    const picks = [];

    // Detect tipster
    const tipsterFromOCR = extractTipsterFromOCR(fullText);
    const tipster = tipsterFromOCR || tipsterFromCaption;

    // Extract date
    const date = extractDate(fullText);

    // Extract odds
    const odds = extractOdds(fullText);
    console.log('   [Parser] Extracted odds:', odds);

    // Extract match
    let match = null;
    const vsPattern = /([A-Za-záéíóúÁÉÍÓÚñÑ]+(?:\s+[A-Za-záéíóúÁÉÍÓÚñÑ]+)?)\s+vs\.?\s+([A-Za-záéíóúÁÉÍÓÚñÑ]+(?:\s+[A-Za-záéíóúÁÉÍÓÚñÑ]+)?)/i;
    const vsMatch = fullText.match(vsPattern);
    
    if (vsMatch) {
        let teamA = vsMatch[1].trim();
        let teamB = vsMatch[2].trim();
        
        const garbage = ['Twit', 'Twitter', 'ML', 'OU', 'AH', 'BTTS', 'pe', 'pra', 'ar', 'CES', 'FE', 'LW', 'ET', 'o', 'ca'];
        for (const g of garbage) {
            teamA = teamA.replace(new RegExp('\\b' + g + '\\b', 'gi'), '').trim();
            teamB = teamB.replace(new RegExp('\\b' + g + '\\b', 'gi'), '').trim();
        }
        
        match = teamA + ' vs. ' + teamB;
    }

    // Build combined pick text
    const pickText = buildPickText(fullText, match);
    console.log('   [Parser] Pick text:', pickText);

    picks.push({
        tipster: tipster ? tipster.name : 'Unknown',
        tipsterKey: tipster ? tipster.key : null,
        sport: detectSport(match),
        match: match,
        pick: pickText,
        odds: odds,
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