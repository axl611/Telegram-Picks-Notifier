// config.js
// Loads environment variables (from .env when present) and exposes a small
// config object used by the application. Do not store secrets in Git.
require('dotenv').config();

const config = {
  TELEGRAM_TOKEN:  process.env.TELEGRAM_TOKEN  || '',
  TELEGRAM_CHAT_ID: process.env.TELEGRAM_CHAT_ID || '',
  ODDS_API_KEY:        process.env.ODDS_API_KEY        || '',
  GOOGLE_VISION_KEY:   process.env.GOOGLE_VISION_KEY   || '',
};

module.exports = config;