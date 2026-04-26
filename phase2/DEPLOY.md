Phase 2 Deployment - Web App

SETUP
1. Open the ElderCare Google Sheet from Phase 1
2. Extensions -> Apps Script
3. Create 4 files in Apps Script editor:
   - Code.gs    (paste content of phase2/Code.gs)
   - Index.html (paste content of phase2/Index.html)
   - Styles.html (paste content of phase2/Styles.html)
   - App.html   (paste content of phase2/App.html)
4. In Code.gs line 1: replace PASTE_YOUR_SHEET_ID_HERE with your Sheet ID from Phase 1

DEPLOY AS WEB APP
1. Click Deploy -> New deployment
2. Type: Web App
3. Execute as: Me
4. Who has access: Anyone (or "Anyone with Google account" for privacy)
5. Click Deploy and copy the Web App URL

ADD TO PHONE HOME SCREEN
iPhone: Open Web App URL in Safari -> tap Share -> Add to Home Screen -> name it ElderCare
Android: Open URL in Chrome -> three dots menu -> Add to Home Screen

TEST CHECKLIST
- Today screen loads and shows seeded tasks (vitamins, magnesium, etc.)
- Tap a task card - it slides right and appears in Done section
- Supplies screen loads showing Magnesium + Multivitamin with "Setup needed" badge
- Tap Ordered on a supply - badge changes to Stock OK
- Health sliders move and values update live
- Tap Save Health Log - toast shows overall score
- Trends screen shows chart after at least 1 health entry
- Telegram alerts (Phase 3) fire at 8am daily
