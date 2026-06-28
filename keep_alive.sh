#!/bin/bash
cd ~/tg-bot

echo "🤖 Bot শুরু হচ্ছে..."

while true; do
    python bot.py
    echo "⚠️ Bot বন্ধ হয়ে গেছে! ৫ সেকেন্ড পর আবার চালু হবে..."
    sleep 5
done
