# Telegram-Picks-Notifier

## Installation & Setup

### Prerequisites
- Node.js (v14 or higher)
- npm (Node Package Manager)

### Step 1: Clone the Repository
```bash
git clone <repository-url>
cd Telegram-Picks-Notifier
```

### Step 2: Install Dependencies
```bash
npm install
```

## Dependencies

This project requires the following npm packages:

### 1. **dotenv** (^17.3.1)
- **Purpose**: Load environment variables from a `.env` file
- **Usage**: Manages sensitive configuration like API keys and credentials
- **Installation**: `npm install dotenv`

### 2. **googleapis** (^171.4.0)
- **Purpose**: Google API client library
- **Usage**: Integrate with Google Sheets for data storage and management
- **Installation**: `npm install googleapis`

### 3. **node-notifier** (^10.0.1)
- **Purpose**: Send desktop notifications
- **Usage**: Display system notifications for alerts and updates
- **Installation**: `npm install node-notifier`

### 4. **qrcode-terminal** (^0.12.0)
- **Purpose**: Generate QR codes in the terminal
- **Usage**: Display QR codes for WhatsApp Web authentication
- **Installation**: `npm install qrcode-terminal`

### 5. **whatsapp-web.js** (^1.34.6)
- **Purpose**: WhatsApp Web client using Puppeteer
- **Usage**: Automate WhatsApp messaging and interaction
- **Installation**: `npm install whatsapp-web.js`

### 6. **puppeteer** (latest)
- **Purpose**: Headless Chrome/Chromium automation library
- **Usage**: Required by whatsapp-web.js for browser automation
- **Installation**: `npm install puppeteer`

## Built-in Node.js Modules (No Installation Required)
- **https**: For making HTTP requests
- **fs**: File system operations
- **path**: Path utilities and file path handling

## System Dependencies

For Puppeteer/Chromium to work properly on Linux, install the following system libraries:

```bash
sudo apt-get update && sudo apt-get install -y libnspr4 libnss3 libatk1.0-0t64 libatk-bridge2.0-0t64 libgdk-pixbuf2.0-0 libgtk-3-0t64 libgbm-dev libasound2t64
```

**Libraries installed:**
- `libnspr4` - Netscape Portable Runtime
- `libnss3` - Network Security Service
- `libatk1.0-0t64` - Accessibility Toolkit
- `libatk-bridge2.0-0t64` - Accessibility Bridge
- `libgdk-pixbuf2.0-0` - Image library
- `libgtk-3-0t64` - GTK+ 3 library
- `libgbm-dev` - Generic Buffer Management
- `libasound2t64` - ALSA sound library

## Environment Variables Configuration

Create a `.env` file in the root directory with the following variables:

```env
TELEGRAM_TOKEN=<your_telegram_bot_token>
TELEGRAM_CHAT_ID=<your_telegram_chat_id>
ODDS_API_KEY=<your_odds_api_key>
GOOGLE_VISION_KEY=<your_google_vision_api_key>
GOOGLE_SHEET_ID=<your_google_sheet_id>
```

**Configuration Details:**
- **TELEGRAM_TOKEN**: Your Telegram Bot API token from BotFather
- **TELEGRAM_CHAT_ID**: The chat ID where the bot will send messages
- **ODDS_API_KEY**: API key for sports odds data
- **GOOGLE_VISION_KEY**: Google Cloud Vision API key for OCR functionality
- **GOOGLE_SHEET_ID**: The ID of your Google Sheet for storing picks and results

⚠️ **Important**: Never commit the `.env` file to Git. It contains sensitive credentials.

## Complete Installation Command
```bash
npm install dotenv googleapis node-notifier qrcode-terminal whatsapp-web.js puppeteer
```