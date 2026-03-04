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

Create a `.env` file in the root directory with the following variables (use `=` between key and value, not `:`):

```env
TELEGRAM_TOKEN=<your_telegram_bot_token>
TELEGRAM_CHAT_ID=<your_telegram_chat_id>
ODDS_API_KEY=<your_odds_api_key>
GOOGLE_VISION_KEY=<your_google_vision_api_key>
GOOGLE_SHEET_ID=<your_google_sheet_id>
```

> ⚠️ Make sure there are **no spaces** around the `=` sign and that you restart the application after editing `.env`.

If dotenv can't parse the file (e.g. you used colons), environment variables will appear empty and the app will report missing keys.

Errors like `Sheets init error: Could not load the default credentials` mean the Google API client didn't find any credentials. See the **Google Sheets Authentication** section below.

**Configuration Details:**
- **TELEGRAM_TOKEN**: Your Telegram Bot API token from BotFather
- **TELEGRAM_CHAT_ID**: The chat ID where the bot will send messages
- **ODDS_API_KEY**: API key for sports odds data
- **GOOGLE_VISION_KEY**: Google Cloud Vision API key for OCR functionality
- **GOOGLE_SHEET_ID**: The ID of your Google Sheet for storing picks and results

⚠️ **Important**: Never commit the `.env` file to Git. It contains sensitive credentials.



## Google Sheets Authentication

The application needs authorized credentials in order to access your spreadsheet. There are three supported methods, choose one:

1. **Service account key file (recommended)**
   * Create a service account in the Google Cloud Console with the `Sheets API` enabled.
   * Download the JSON key file and place it in the project root as `service-account.json`.
   * Alternatively you can store it elsewhere and set `GOOGLE_CREDENTIALS_FILE` to its path.

2. **Environment variables** (useful for CI or containers)
   - `GOOGLE_SERVICE_ACCOUNT_EMAIL` – the service account's email address
   - `GOOGLE_PRIVATE_KEY` – the private key contents (newline characters must be escaped as `\n`)

   Example:
   ```bash
   export GOOGLE_SERVICE_ACCOUNT_EMAIL="your@service-account.iam.gserviceaccount.com"
   export GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   ```

3. **Application Default Credentials**
   Run `gcloud auth application-default login` on the machine where the bot runs. This stores credentials that the client library will automatically pick up.
   
   > ⚠️ If `gcloud` isn’t installed (e.g. you see “Command 'gcloud' not found”), you can skip this method entirely and instead use one of the other two options above. Installing the Google Cloud SDK is optional and only needed for this credential flow.
4. **OAuth 2.0 (Using oauth-client.json)**
   If you have an `oauth-client.json` file from Google Cloud Console (Desktop/Installed app credentials):
   
   **First Run: Authorize the App**
   1. Run the application:
      ```bash
      node index.js
      ```
   
   2. The app will prompt you to visit:
      ```
      📱 Visit this link in your browser:
      http://localhost:3000
      ```
   
   3. Open that URL in your browser and you'll be redirected to Google's authorization screen
   
   4. Click "Allow" to grant permissions
   
   5. The app automatically captures the authorization code and saves the token
   
   6. You'll see a success message and the app continues
   
   **Subsequent Runs: No Authorization Needed**
   - The auth token is cached in `oauth-token.json` (in `.gitignore`)
   - Simply run `node index.js` and it will use the cached token
   - Tokens auto-refresh as needed
   
   ⚠️ **Never commit `oauth-token.json` to Git**—it contains sensitive auth data
Once credentials are available, make sure the spreadsheet ID is set (`GOOGLE_SHEET_ID`) and restart the app; the earlier error will disappear.

## Complete Installation Command
```bash
npm install dotenv googleapis node-notifier qrcode-terminal whatsapp-web.js puppeteer
```