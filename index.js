const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const fs = require('fs');
const https = require('https');
// Tesseract replaced by Google Vision
const config = require('./config');
const { parsePicks } = require('./parser');
const { initExcel, addPick, updateSummary } = require('./excel');
const { init: initResults, startResultsScheduler } = require('./results');
const { initCache } = require('./odds-cache');


// Initialize Excel - will create new file if missing, or load existing one
initExcel().catch(err => console.error('Excel init error: ' + err.message));

// Initialize Odds API cache (fetches upcoming matches once on startup)
initCache().catch(err => console.error('OddsCache init error: ' + err.message));

// Initialize results checker with credentials from .env
initResults(config.TELEGRAM_TOKEN, config.TELEGRAM_CHAT_ID, config.ODDS_API_KEY);

const TARGET_GROUPS = ['Axl', 'PRAIZA MAR26 👽💰 PICKS'];
const MY_NAME = 'Axl';
const ALERT_PREFIX = '🔔ALERT:';
const HEARTBEAT_INTERVAL_MINUTES = 5;
const COUNTERS_FILE = './counters.json';
const CONTEXT_WINDOW_MS = 30 * 1000;
const IMAGE_WAIT_MS = 5 * 1000;

const TELEGRAM_TOKEN = config.TELEGRAM_TOKEN;
const TELEGRAM_CHAT_ID = config.TELEGRAM_CHAT_ID;

const KEYWORD_GROUPS = {
    abuelo: { keywords: ['abuelo', 'abuelito'], label: 'del Abuelo', name: 'Abuelo' },
    cristian: { keywords: ['cristian'], label: 'de Cristian', name: 'Cristian Rey' },
    roberto: { keywords: ['roberto', 'beto'], label: 'de Roberto', name: 'Roberto Rey' }
};

const contextWindows = {};
const pendingAlerts = {};

function loadCounters() {
    try {
        if (fs.existsSync(COUNTERS_FILE)) return JSON.parse(fs.readFileSync(COUNTERS_FILE, 'utf8'));
    } catch (_) {}
    return {};
}

function saveCounters() {
    try { fs.writeFileSync(COUNTERS_FILE, JSON.stringify(counters, null, 2)); }
    catch (err) { console.error('Could not save counters: ' + err.message); }
}

const counters = loadCounters();
let isReady = false;
let heartbeatTimer = null;
let reconnectAttempts = 0;
const processedIds = new Set();

function normalize(text) {
    return text.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function matchedKeywordGroup(text) {
    const norm = normalize(text);
    for (const [groupKey, group] of Object.entries(KEYWORD_GROUPS)) {
        if (group.keywords.some(kw => norm.includes(normalize(kw)))) {
            return { groupKey, label: group.label, name: group.name };
        }
    }
    return null;
}

function matchedGroup(chatName) {
    return TARGET_GROUPS.find(g => normalize(chatName).includes(normalize(g))) || null;
}

function getTodayFormatted() {
    const now = new Date();
    return now.getDate() + ' ' + now.toLocaleString('en', { month: 'short' });
}

function getTodayKey() {
    return new Date().toISOString().split('T')[0];
}

function getNextCount(senderName, groupKey) {
    const key = getTodayKey() + ':' + senderName + ':' + groupKey;
    if (!counters[key]) counters[key] = 0;
    counters[key]++;
    saveCounters();
    return counters[key];
}

function timestamp() {
    return '[' + new Date().toLocaleTimeString() + ']';
}

async function extractTextFromImage(base64Data) {
    try {
        const apiKey = config.GOOGLE_VISION_KEY;
        if (!apiKey) { console.error('No GOOGLE_VISION_KEY in .env'); return ''; }

        const body = JSON.stringify({
            requests: [{
                image: { content: base64Data },
                features: [{ type: 'TEXT_DETECTION', maxResults: 1 }]
            }]
        });

        const result = await new Promise((resolve, reject) => {
            const options = {
                hostname: 'vision.googleapis.com',
                path: '/v1/images:annotate?key=' + apiKey,
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
            };
            const req = https.request(options, res => {
                let raw = '';
                res.on('data', c => raw += c);
                res.on('end', () => {
                    try { resolve(JSON.parse(raw)); }
                    catch (e) { reject(new Error('Vision parse error: ' + raw.slice(0, 100))); }
                });
            });
            req.on('error', reject);
            req.write(body);
            req.end();
        });

        const annotation = result.responses &&
                           result.responses[0] &&
                           result.responses[0].fullTextAnnotation;
        if (!annotation) {
            console.log('   Google Vision: no text found in image.');
            return '';
        }
        return annotation.text.trim();
    } catch (err) {
        console.error('Google Vision error: ' + err.message);
        return '';
    }
}

function sendTelegramText(text) {
    return new Promise((resolve) => {
        const body = JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: text });
        const options = {
            hostname: 'api.telegram.org',
            path: '/bot' + TELEGRAM_TOKEN + '/sendMessage',
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
        };
        const req = https.request(options, res => { res.on('data', () => {}); res.on('end', resolve); });
        req.on('error', err => { console.error('Telegram text error: ' + err.message); resolve(); });
        req.write(body);
        req.end();
    });
}

function sendTelegramImage(base64Data, mimeType, caption) {
    return new Promise((resolve) => {
        try {
            const imageBuffer = Buffer.from(base64Data, 'base64');
            const boundary = '----FormBoundary' + Math.random().toString(36);
            const ext = mimeType.includes('png') ? 'png' : 'jpg';
            const parts = '--' + boundary + '\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n' + TELEGRAM_CHAT_ID + '\r\n' +
                '--' + boundary + '\r\nContent-Disposition: form-data; name="caption"\r\n\r\n' + caption + '\r\n' +
                '--' + boundary + '\r\nContent-Disposition: form-data; name="photo"; filename="pick.' + ext + '"\r\nContent-Type: ' + mimeType + '\r\n\r\n';
            const body = Buffer.concat([Buffer.from(parts), imageBuffer, Buffer.from('\r\n--' + boundary + '--\r\n')]);
            const options = {
                hostname: 'api.telegram.org',
                path: '/bot' + TELEGRAM_TOKEN + '/sendPhoto',
                method: 'POST',
                headers: { 'Content-Type': 'multipart/form-data; boundary=' + boundary, 'Content-Length': body.length }
            };
            const req = https.request(options, res => { res.on('data', () => {}); res.on('end', resolve); });
            req.on('error', err => { console.error('Telegram image error: ' + err.message); resolve(); });
            req.write(body);
            req.end();
        } catch (err) { console.error('Telegram image build error: ' + err.message); resolve(); }
    });
}

function openContextWindow(chatName, groupKey, label, name, sender, count) {
    if (contextWindows[chatName] && contextWindows[chatName].timer) clearTimeout(contextWindows[chatName].timer);
    const timer = setTimeout(() => {
        console.log(timestamp() + ' Context window closed for "' + chatName + '"');
        delete contextWindows[chatName];
    }, CONTEXT_WINDOW_MS);
    contextWindows[chatName] = { groupKey, label, name, sender, count, timer };
    console.log(timestamp() + ' Context window opened for "' + chatName + '" (' + label + ') - 30s');
}

function closeContextWindow(chatName) {
    if (contextWindows[chatName]) {
        if (contextWindows[chatName].timer) clearTimeout(contextWindows[chatName].timer);
        delete contextWindows[chatName];
    }
}

function queueTextAlert(chatName, alertMessage, tipsterName) {
    if (pendingAlerts[chatName] && pendingAlerts[chatName].timer) clearTimeout(pendingAlerts[chatName].timer);
    const timer = setTimeout(async () => {
        delete pendingAlerts[chatName];
        await sendTelegramText(alertMessage);
    }, IMAGE_WAIT_MS);
    pendingAlerts[chatName] = { alertMessage, tipsterName, timer };
}

function startHeartbeat() {
    if (heartbeatTimer) clearInterval(heartbeatTimer);
    heartbeatTimer = setInterval(() => {
        console.log(timestamp() + (isReady ? ' Heartbeat - listening...' : ' Heartbeat - NOT connected'));
    }, HEARTBEAT_INTERVAL_MINUTES * 60 * 1000);
}

async function processImage(msg, alertMessage, media, tipsterFromCaption) {
    await sendTelegramImage(media.data, media.mimetype, alertMessage);
    console.log('   Image sent to Telegram.');

    if (!media.mimetype.startsWith('image/')) return;

    setImmediate(async () => {
        try {
            console.log('   Running OCR...');
            const ocrText = await extractTextFromImage(media.data);

            if (!ocrText) {
                console.log('   OCR returned no text.');
                return;
            }

            console.log('   OCR: ' + ocrText.replace(/\n/g, ' | '));

            const picks = await parsePicks(ocrText, tipsterFromCaption);
            console.log('   Parsed ' + picks.length + ' pick(s)');

            const { scheduleResultCheck } = require('./results');

            for (const pick of picks) {
                await addPick(pick);
                
                // Schedule result check with sport-specific delays
                if (pick.sport) {
                    scheduleResultCheck(
                        pick.tipster,
                        pick.date,
                        pick.match,
                        pick.pick,
                        pick.odds,
                        pick.sport
                    );
                }
            }

            await updateSummary();
            console.log('   Excel updated.');

        } catch (err) {
            console.error('   Error: ' + err.message);
        }
    });
}

function createClient() {
    const client = new Client({
        authStrategy: new LocalAuth(),
        puppeteer: { headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] }
    });

    client.on('qr', qr => {
        console.log('\nScan this QR code:\n');
        qrcode.generate(qr, { small: true });
    });

    client.on('ready', () => {
        isReady = true;
        reconnectAttempts = 0;
        console.log('\n✅ WhatsApp connected!');
        console.log('👀 Watching: ' + TARGET_GROUPS.join(', '));
        console.log('🔑 Tracking: Abuelo, Cristian, Roberto\n');

        // Start results scheduler once WhatsApp is connected
        startResultsScheduler();
    });

    client.on('auth_failure', () => {
        isReady = false;
        console.error('Authentication failed. Delete .wwebjs_auth folder and restart.');
    });

    client.on('disconnected', async reason => {
        isReady = false;
        console.log('Disconnected: ' + reason);
        reconnectAttempts++;
        const waitSeconds = Math.min(reconnectAttempts * 10, 60);
        setTimeout(async () => {
            try { await client.initialize(); }
            catch (err) { console.error('Reconnect failed: ' + err.message); }
        }, waitSeconds * 1000);
    });

    async function handleMessage(msg) {
        try {
            const body = msg.body || '';
            if (body.startsWith(ALERT_PREFIX)) return;

            const msgId = msg.id && msg.id.id ? msg.id.id : null;
            if (msgId) {
                if (processedIds.has(msgId)) return;
                processedIds.add(msgId);
            }

            const chat = await msg.getChat();
            const chatName = chat.name || '';
            if (!matchedGroup(chatName)) return;

            const isImage = msg.hasMedia && ['image', 'video', 'document'].includes(msg.type);

            let sender = msg.fromMe ? MY_NAME : 'Unknown';
            if (!msg.fromMe) {
                const contact = await msg.getContact();
                sender = contact.pushname || contact.name || contact.number || 'Unknown';
            }

            const match = body ? matchedKeywordGroup(body) : null;

            if (match && !isImage) {
                const count = getNextCount(sender, match.groupKey);
                const date = getTodayFormatted();
                const alertMessage = '(' + sender + ') Pick ' + match.label + ' - ' + date + ' - #' + count;
                console.log('\n' + timestamp() + ' ' + alertMessage);
                openContextWindow(chatName, match.groupKey, match.label, match.name, sender, count);
                queueTextAlert(chatName, alertMessage, match.name);
                return;
            }

            if (match && isImage) {
                const count = getNextCount(sender, match.groupKey);
                const date = getTodayFormatted();
                const alertMessage = '(' + sender + ') Pick ' + match.label + ' - ' + date + ' - #' + count;
                console.log('\n' + timestamp() + ' ' + alertMessage + ' (image)');

                if (pendingAlerts[chatName]) {
                    clearTimeout(pendingAlerts[chatName].timer);
                    delete pendingAlerts[chatName];
                }

                const media = await msg.downloadMedia();
                if (media) await processImage(msg, alertMessage, media, match.name);
                openContextWindow(chatName, match.groupKey, match.label, match.name, sender, count);
                return;
            }

            if (isImage) {
                const media = await msg.downloadMedia();
                if (!media) return;

                if (pendingAlerts[chatName]) {
                    const pending = pendingAlerts[chatName];
                    clearTimeout(pending.timer);
                    delete pendingAlerts[chatName];
                    await processImage(msg, pending.alertMessage, media, pending.tipsterName);
                    return;
                }

                if (contextWindows[chatName]) {
                    const ctx = contextWindows[chatName];
                    const date = getTodayFormatted();
                    const alertMessage = '(' + sender + ') Pick ' + ctx.label + ' - ' + date;
                    await processImage(msg, alertMessage, media, ctx.name);
                    return;
                }
                return;
            }

            if (!isImage && !match && contextWindows[chatName]) {
                closeContextWindow(chatName);
            }

        } catch (err) {
            console.error('Error: ' + err.message);
        }
    }

    client.on('message_create', handleMessage);
    return client;
}

console.log('\n========================================');
console.log(' WhatsApp Picks Notifier');
console.log('========================================\n');

startHeartbeat();
const client = createClient();
client.initialize();