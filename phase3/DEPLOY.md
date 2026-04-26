Phase 3 Deployment - Telegram Bot

GET BOT TOKEN
1. Open Telegram -> search for @BotFather
2. Send: /newbot
3. Choose a name: ElderCare Bot
4. Choose a username: eldercare_yourname_bot
5. Copy the token BotFather gives you (looks like 1234567890:ABCdef...)

GET YOUR CHAT ID
1. Search for your new bot by its username and tap Start
2. Send it any message (e.g. "hello")
3. Open this URL in a browser (replace TOKEN with your actual token):
   https://api.telegram.org/botTOKEN/getUpdates
4. Find "chat":{"id": 123456789} in the response - that number is your CHAT_ID

ADD TO APPS SCRIPT (same project as Phase 2)
1. Open the ElderCare Apps Script project from Phase 2
2. Click + to create a new file -> name it Telegram
3. Paste the content of phase3/Telegram.gs
4. At the top of the file, replace:
   - PASTE_TELEGRAM_BOT_TOKEN_HERE with your token
   - PASTE_TELEGRAM_CHAT_ID_HERE with your chat ID (as a string, e.g. '123456789')

SET DAILY ALERT TRIGGER (8am every day)
1. In Apps Script, click the clock icon (Triggers) on the left sidebar
2. Click + Add Trigger (bottom right)
3. Function to run: dailyMorningAlert
4. Event source: Time-driven
5. Time-based trigger type: Day timer
6. Time of day: 8am to 9am
7. Click Save

SET WEEKLY SUMMARY TRIGGER (Sunday morning)
1. Click + Add Trigger again
2. Function to run: weeklySummaryAlert
3. Event source: Time-driven
4. Time-based trigger type: Week timer
5. Day of week: Every Sunday
6. Time of day: 9am to 10am
7. Click Save

TEST
1. In the Apps Script editor, select dailyMorningAlert from the function dropdown
2. Click Run
3. Check your Telegram - you should receive the morning alert message
