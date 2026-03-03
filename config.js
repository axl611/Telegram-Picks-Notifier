// config.js
// Loads environment variables (from .env when present) and exposes a small
// config object used by the application. Do not store secrets in Git.

require('dotenv').config();

const config = {
  TELEGRAM_TOKEN: process.env.TELEGRAM_TOKEN || '',
  TELEGRAM_CHAT_ID: process.env.TELEGRAM_CHAT_ID || '',
  // Add other env-backed values here. Example:
  // JWT_SECRET: process.env.JWT_SECRET || 'dev-placeholder'
};

module.exports = config;
