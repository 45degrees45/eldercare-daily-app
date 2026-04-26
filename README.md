# ElderCare

Mobile-first PWA and Telegram bot for tracking elderly parent care tasks, supplement reordering, and health metrics. Built on Google Sheets and Google Apps Script.

## Architecture

```
PWA (hosted on Apps Script Web App)
  ^^ google.script.run
Apps Script Backend (Code.gs)
  ^^ SpreadsheetApp
Google Sheet (TASKS, SUPPLIES, HEALTH_LOG, LOG)

Telegram.gs (scheduled triggers) -> Telegram Bot API -> Your phone
```

## Phases

- Phase 1 - Google Sheet setup (phase1/)
- Phase 2 - Web App PWA + Backend (phase2/)
- Phase 3 - Telegram alerts (phase3/)
- Phase 4 - Google Calendar integration (future)

## Deploy Order

1. Phase 1 first (creates the Sheet)
2. Phase 2 second (creates the web app - needs Sheet ID from Phase 1)
3. Phase 3 last (add to same Apps Script project as Phase 2)

See each phase*/DEPLOY.md for step-by-step instructions.

## Key Features

- Swipe-to-complete today's tasks
- Smart reorder alerts (calculates order date from qty + daily usage + delivery time)
- Daily health log (energy, sleep, focus, mood, confidence)
- Trend charts (last 14 days)
- Telegram morning briefing and weekly summary
