#!/bin/bash

# Installation script for Telegram-Picks-Notifier dependencies
# This script installs all required npm packages

echo "=========================================="
echo "Installing Telegram-Picks-Notifier"
echo "Dependencies"
echo "=========================================="
echo ""

# Check if npm is installed
if ! command -v npm &> /dev/null; then
    echo "❌ npm is not installed. Please install Node.js and npm first."
    exit 1
fi

echo "✓ npm found"
echo ""

# Install dependencies
echo "Installing dependencies..."
echo ""

npm install dotenv googleapis node-notifier qrcode-terminal whatsapp-web.js puppeteer

# Check if installation was successful
if [ $? -eq 0 ]; then
    echo ""
    echo "=========================================="
    echo "✓ All dependencies installed successfully!"
    echo "=========================================="
    echo ""
    echo "Installed packages:"
    echo "  • dotenv (^17.3.1)"
    echo "  • googleapis (^171.4.0)"
    echo "  • node-notifier (^10.0.1)"
    echo "  • qrcode-terminal (^0.12.0)"
    echo "  • whatsapp-web.js (^1.34.6)"
    echo "  • puppeteer (latest)"
    echo ""
else
    echo ""
    echo "❌ Installation failed. Please check the error messages above."
    exit 1
fi
