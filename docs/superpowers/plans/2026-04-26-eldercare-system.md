# ElderCare System Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a mobile-first PWA + Telegram bot system for tracking an elderly parent's care tasks, supplement reordering, and health metrics — with Google Sheets as the intelligence layer.

**Architecture:** Google Apps Script serves the PWA (HTML/CSS/JS) and acts as the backend API, reading/writing a Google Sheet. A Telegram bot sends push alerts for reorder reminders and daily task nudges. Google Calendar integration is deferred as a future phase.

**Tech Stack:** Google Apps Script (backend + host), Vanilla JS PWA (frontend), Google Sheets (data), Telegram Bot API (alerts)

---

## File Structure

```
059_2026APR_ElderCare/
├── phase1/
│   ├── setup.gs              # Creates all Sheet tabs + columns + formulas
│   └── DEPLOY.md
├── phase2/
│   ├── Code.gs               # Apps Script backend — doGet, API handlers
│   ├── Index.html            # PWA shell — nav, screen containers
│   ├── Styles.html           # All CSS — mobile-first, swipe UX
│   ├── App.html              # All JS — screen logic, fetch, animations
│   └── DEPLOY.md
├── phase3/
│   ├── Telegram.gs           # Bot setup, daily trigger, alert logic
│   └── DEPLOY.md
└── README.md
```

### Sheet Structure

| Sheet | Purpose |
|---|---|
| TASKS | Recurring care tasks (bills, vitamins, plants) |
| SUPPLIES | Supplement inventory + auto reorder dates |
| HEALTH_LOG | Daily health metrics (energy, sleep, mood, etc.) |
| LOG | Hidden — Apps Script error log |

---

## Task 1: Google Sheet Setup (Phase 1)

**Files:**
- Create: `phase1/setup.gs`
- Create: `phase1/DEPLOY.md`

- [ ] **Step 1: Create `phase1/setup.gs`**

```javascript
function setupElderCare() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- TASKS sheet ---
  let tasks = ss.getSheetByName('TASKS') || ss.insertSheet('TASKS');
  tasks.clearContents();
  tasks.getRange(1, 1, 1, 9).setValues([[
    'ID', 'Title', 'Category', 'Frequency',
    'Last Done', 'Next Due', 'Status', 'Notes', 'Active'
  ]]);
  tasks.getRange('A1:I1').setFontWeight('bold').setBackground('#34A853').setFontColor('white');
  const taskCatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Health', 'Finance', 'Home', 'Garden', 'Service'], true).build();
  tasks.getRange('C2:C200').setDataValidation(taskCatRule);
  const freqRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['daily', 'weekly', 'monthly', 'once'], true).build();
  tasks.getRange('D2:D200').setDataValidation(freqRule);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Paused', 'Done'], true).build();
  tasks.getRange('G2:G200').setDataValidation(statusRule);

  // --- SUPPLIES sheet ---
  let supplies = ss.getSheetByName('SUPPLIES') || ss.insertSheet('SUPPLIES');
  supplies.clearContents();
  supplies.getRange(1, 1, 1, 9).setValues([[
    'ID', 'Item', 'Category', 'Qty On Hand',
    'Daily Usage', 'Delivery Days', 'Last Ordered', 'Reorder Date', 'Notes'
  ]]);
  supplies.getRange('A1:I1').setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  const supCatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Vitamin', 'Mineral', 'Medicine', 'Grocery', 'Household'], true).build();
  supplies.getRange('C2:C200').setDataValidation(supCatRule);

  // --- HEALTH_LOG sheet ---
  let health = ss.getSheetByName('HEALTH_LOG') || ss.insertSheet('HEALTH_LOG');
  health.clearContents();
  health.getRange(1, 1, 1, 8).setValues([[
    'Date', 'Energy (1-10)', 'Sleep (hrs)', 'Focus (1-10)',
    'Mood (1-10)', 'Confidence (1-10)', 'Overall', 'Notes'
  ]]);
  health.getRange('A1:H1').setFontWeight('bold').setBackground('#FF6D00').setFontColor('white');

  // --- LOG sheet (hidden) ---
  let log = ss.getSheetByName('LOG') || ss.insertSheet('LOG');
  log.clearContents();
  log.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Action', 'Detail']]);
  ss.setActiveSheet(tasks);
  try { ss.getSheetByName('LOG').hideSheet(); } catch(e) {}

  // --- Seed starter tasks ---
  tasks.getRange('A2:I6').setValues([
    [1, 'Take morning vitamins', 'Health', 'daily', '', '', 'Active', '', true],
    [2, 'Take magnesium', 'Health', 'daily', '', '', 'Active', '', true],
    [3, 'Pay electricity bill', 'Finance', 'monthly', '', '', 'Active', '', true],
    [4, 'Water the plants', 'Garden', 'weekly', '', '', 'Active', '', true],
    [5, 'Clean the house', 'Home', 'weekly', '', '', 'Active', '', true],
  ]);

  // --- Seed starter supplies ---
  supplies.getRange('A2:I3').setValues([
    [1, 'Magnesium tablets', 'Mineral', 50, 1, 7, '', '', ''],
    [2, 'Multivitamin', 'Vitamin', 60, 1, 7, '', '', ''],
  ]);

  SpreadsheetApp.getUi().alert('ElderCare sheet setup complete!');
}
```

- [ ] **Step 2: Create `phase1/DEPLOY.md`**

```
Phase 1 Deployment — Sheet Setup

1. Create a new Google Sheet, name it: ElderCare
2. Extensions -> Apps Script -> delete default content
3. Paste setup.gs into Code.gs
4. Click Save -> Run setupElderCare
5. Approve permissions
6. Verify 4 sheets: TASKS, SUPPLIES, HEALTH_LOG, LOG (hidden)
7. Copy the Sheet ID from the URL — needed for Phase 2

Sheet ID is in: https://docs.google.com/spreadsheets/d/SHEET_ID/edit
```

- [ ] **Step 3: Manual verify**
  - Open a new Google Sheet named `ElderCare`
  - Run `setupElderCare`
  - Confirm 4 sheets created with correct headers and seed data

---

## Task 2: Apps Script Backend (Phase 2 — Code.gs)

**Files:**
- Create: `phase2/Code.gs`

- [ ] **Step 1: Create `phase2/Code.gs` with doGet + helpers**

```javascript
const SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ElderCare')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _log(action, detail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const log = ss.getSheetByName('LOG');
  log.appendRow([new Date(), action, detail]);
}
```

- [ ] **Step 2: Add `getTodayTasks()` to Code.gs**

```javascript
function getTodayTasks() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('TASKS');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const [id, title, category, frequency, lastDone, nextDue, status, notes, active] = data[i];
    if (!active || status !== 'Active' || !title) continue;

    let isDueToday = false;
    if (!lastDone || lastDone === '') {
      isDueToday = true;
    } else {
      const last = new Date(lastDone);
      last.setHours(0, 0, 0, 0);
      const daysSince = Math.floor((today - last) / 86400000);
      if (frequency === 'daily' && daysSince >= 1) isDueToday = true;
      if (frequency === 'weekly' && daysSince >= 7) isDueToday = true;
      if (frequency === 'monthly' && daysSince >= 28) isDueToday = true;
    }

    if (isDueToday) {
      tasks.push({ id, title, category, frequency, lastDone: lastDone ? lastDone.toISOString() : '', row: i + 1 });
    }
  }
  return tasks;
}
```

- [ ] **Step 3: Add `markTaskDone(id)` to Code.gs**

```javascript
function markTaskDone(taskId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('TASKS');
  const data = sheet.getDataRange().getValues();
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == taskId) {
      sheet.getRange(i + 1, 5).setValue(today);
      _log('markTaskDone', 'ID=' + taskId + ' title=' + data[i][1]);
      return { success: true, title: data[i][1] };
    }
  }
  return { success: false, error: 'Task not found' };
}
```

- [ ] **Step 4: Add `getSupplies()` to Code.gs**

```javascript
function getSupplies() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SUPPLIES');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const supplies = [];
  for (let i = 1; i < data.length; i++) {
    const [id, item, category, qty, dailyUsage, deliveryDays, lastOrdered, reorderDate, notes] = data[i];
    if (!item) continue;

    let daysUntilReorder = null;
    let status = 'ok';
    if (lastOrdered && dailyUsage > 0) {
      const reorder = new Date(reorderDate);
      reorder.setHours(0, 0, 0, 0);
      daysUntilReorder = Math.floor((reorder - today) / 86400000);
      if (daysUntilReorder <= 0) status = 'order_now';
      else if (daysUntilReorder <= 3) status = 'order_soon';
    } else {
      status = 'needs_setup';
    }

    supplies.push({
      id, item, category, qty, dailyUsage, deliveryDays,
      lastOrdered: lastOrdered ? new Date(lastOrdered).toISOString() : '',
      daysUntilReorder, status, row: i + 1
    });
  }
  return supplies;
}
```

- [ ] **Step 5: Add `markOrdered(id)` to Code.gs**

```javascript
function markOrdered(supplyId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SUPPLIES');
  const data = sheet.getDataRange().getValues();
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == supplyId) {
      const qty = data[i][3];
      const dailyUsage = data[i][4];
      const deliveryDays = data[i][5];
      sheet.getRange(i + 1, 7).setValue(today);
      const reorderDate = new Date(today.getTime() + ((qty / dailyUsage) - deliveryDays) * 86400000);
      sheet.getRange(i + 1, 8).setValue(reorderDate);
      _log('markOrdered', 'ID=' + supplyId + ' item=' + data[i][1]);
      return { success: true, item: data[i][1], reorderDate: reorderDate.toISOString() };
    }
  }
  return { success: false, error: 'Supply not found' };
}
```

- [ ] **Step 6: Add `logHealth()` and `getHealthTrends()` to Code.gs**

```javascript
function logHealth(entry) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('HEALTH_LOG');
  const today = new Date();
  const overall = ((entry.energy + entry.focus + entry.mood + entry.confidence) / 4).toFixed(1);
  sheet.appendRow([today, entry.energy, entry.sleep, entry.focus, entry.mood, entry.confidence, overall, entry.notes || '']);
  _log('logHealth', 'overall=' + overall);
  return { success: true, overall };
}

function getHealthTrends() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('HEALTH_LOG');
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0]).slice(-30);
  return rows.map(r => ({
    date: new Date(r[0]).toISOString(),
    energy: r[1], sleep: r[2], focus: r[3],
    mood: r[4], confidence: r[5], overall: r[6]
  }));
}
```

---

## Task 3: PWA Shell — Index.html + Styles.html

**Files:**
- Create: `phase2/Index.html`
- Create: `phase2/Styles.html`

- [ ] **Step 1: Create `phase2/Index.html`**

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
  <meta name="mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <title>ElderCare</title>
  <style><?!= include('Styles'); ?></style>
</head>
<body>
  <div id="app">
    <nav class="bottom-nav">
      <button class="nav-btn active" onclick="showScreen('today')" id="nav-today">
        <span class="nav-icon">📋</span><span class="nav-label">Today</span>
      </button>
      <button class="nav-btn" onclick="showScreen('supplies')" id="nav-supplies">
        <span class="nav-icon">📦</span><span class="nav-label">Supplies</span>
      </button>
      <button class="nav-btn" onclick="showScreen('health')" id="nav-health">
        <span class="nav-icon">💚</span><span class="nav-label">Health</span>
      </button>
      <button class="nav-btn" onclick="showScreen('charts')" id="nav-charts">
        <span class="nav-icon">📈</span><span class="nav-label">Trends</span>
      </button>
    </nav>

    <div id="screen-today" class="screen active">
      <div class="screen-header">
        <h1>Today</h1>
        <span id="today-date" class="header-date"></span>
      </div>
      <div id="tasks-loading" class="loading-spinner">Loading...</div>
      <div id="tasks-list" class="card-list"></div>
      <div id="tasks-done" class="done-section" style="display:none">
        <h3 class="done-label">Done today</h3>
        <div id="tasks-done-list" class="card-list muted"></div>
      </div>
    </div>

    <div id="screen-supplies" class="screen">
      <div class="screen-header"><h1>Supplies</h1></div>
      <div id="supplies-loading" class="loading-spinner">Loading...</div>
      <div id="supplies-list" class="card-list"></div>
    </div>

    <div id="screen-health" class="screen">
      <div class="screen-header">
        <h1>Health Log</h1>
        <span class="header-date">How is mum today?</span>
      </div>
      <form id="health-form" class="health-form">
        <div class="metric-row">
          <label>Energy <span class="metric-val" id="val-energy">5</span></label>
          <input type="range" min="1" max="10" value="5" id="slider-energy" oninput="updateVal('energy',this.value)">
        </div>
        <div class="metric-row">
          <label>Sleep (hrs) <span class="metric-val" id="val-sleep">7</span></label>
          <input type="range" min="1" max="12" value="7" step="0.5" id="slider-sleep" oninput="updateVal('sleep',this.value)">
        </div>
        <div class="metric-row">
          <label>Focus <span class="metric-val" id="val-focus">5</span></label>
          <input type="range" min="1" max="10" value="5" id="slider-focus" oninput="updateVal('focus',this.value)">
        </div>
        <div class="metric-row">
          <label>Mood <span class="metric-val" id="val-mood">5</span></label>
          <input type="range" min="1" max="10" value="5" id="slider-mood" oninput="updateVal('mood',this.value)">
        </div>
        <div class="metric-row">
          <label>Confidence <span class="metric-val" id="val-confidence">5</span></label>
          <input type="range" min="1" max="10" value="5" id="slider-confidence" oninput="updateVal('confidence',this.value)">
        </div>
        <div class="metric-row">
          <label>Notes (optional)</label>
          <input type="text" id="health-notes" placeholder="Anything unusual?" class="text-input">
        </div>
        <button type="button" onclick="submitHealth()" class="btn-primary">Save Health Log</button>
      </form>
    </div>

    <div id="screen-charts" class="screen">
      <div class="screen-header"><h1>Trends</h1></div>
      <div id="charts-loading" class="loading-spinner">Loading...</div>
      <div id="charts-container"></div>
    </div>
  </div>
  <script><?!= include('App'); ?></script>
</body>
</html>
```

- [ ] **Step 2: Create `phase2/Styles.html`**

```css
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  background: #f5f5f7; color: #1c1c1e;
  -webkit-tap-highlight-color: transparent;
}
#app { max-width: 480px; margin: 0 auto; min-height: 100vh; position: relative; }

.bottom-nav {
  position: fixed; bottom: 0; left: 50%; transform: translateX(-50%);
  width: 100%; max-width: 480px; display: flex; background: white;
  border-top: 1px solid #e5e5ea; z-index: 100;
  padding-bottom: env(safe-area-inset-bottom);
}
.nav-btn {
  flex: 1; padding: 10px 0; background: none; border: none;
  display: flex; flex-direction: column; align-items: center; gap: 2px;
  cursor: pointer; color: #8e8e93; transition: color .2s;
}
.nav-btn.active { color: #34A853; }
.nav-icon { font-size: 22px; }
.nav-label { font-size: 10px; font-weight: 500; }

.screen { display: none; padding: 16px 16px 90px; min-height: 100vh; }
.screen.active { display: block; }
.screen-header {
  display: flex; justify-content: space-between; align-items: center;
  margin-bottom: 20px;
}
.screen-header h1 { font-size: 28px; font-weight: 700; }
.header-date { font-size: 13px; color: #8e8e93; }

.card-list { display: flex; flex-direction: column; gap: 10px; }
.task-card {
  background: white; border-radius: 14px; padding: 16px 18px;
  display: flex; align-items: center; gap: 14px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08); cursor: pointer;
  transition: transform .15s, opacity .3s;
}
.task-card:active { transform: scale(0.97); }
.task-icon {
  width: 48px; height: 48px; border-radius: 12px;
  display: flex; align-items: center; justify-content: center; font-size: 22px;
}
.cat-health { background: #e8f5e9; }
.cat-finance { background: #e3f2fd; }
.cat-home { background: #fce4ec; }
.cat-garden { background: #f3e5f5; }
.cat-service { background: #fff8e1; }
.task-info { flex: 1; }
.task-title { font-size: 15px; font-weight: 600; }
.task-sub { font-size: 12px; color: #8e8e93; margin-top: 2px; }
.task-check { font-size: 22px; }
.task-card.muted { opacity: 0.45; }
.done-section { margin-top: 24px; }
.done-label {
  font-size: 13px; color: #8e8e93; margin-bottom: 10px;
  font-weight: 600; text-transform: uppercase; letter-spacing: .5px;
}

.supply-card {
  background: white; border-radius: 14px; padding: 16px 18px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
.supply-header { display: flex; justify-content: space-between; align-items: center; }
.supply-name { font-size: 15px; font-weight: 600; }
.supply-badge { font-size: 11px; padding: 3px 10px; border-radius: 20px; font-weight: 600; }
.badge-ok { background: #e8f5e9; color: #2e7d32; }
.badge-soon { background: #fff8e1; color: #f57f17; }
.badge-now { background: #fce4ec; color: #c62828; }
.badge-setup { background: #e3f2fd; color: #1565c0; }
.supply-detail { font-size: 12px; color: #8e8e93; margin-top: 6px; }
.btn-order {
  margin-top: 10px; width: 100%; padding: 10px; background: #4285F4;
  color: white; border: none; border-radius: 10px;
  font-size: 14px; font-weight: 600; cursor: pointer;
}
.btn-order:active { opacity: .8; }

.health-form { display: flex; flex-direction: column; gap: 20px; }
.metric-row {
  background: white; border-radius: 14px; padding: 16px 18px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
.metric-row label {
  display: flex; justify-content: space-between;
  font-size: 14px; font-weight: 600; margin-bottom: 10px;
}
.metric-val { font-size: 18px; color: #34A853; font-weight: 700; }
input[type=range] { width: 100%; accent-color: #34A853; height: 6px; }
.text-input {
  width: 100%; border: 1px solid #e5e5ea; border-radius: 10px;
  padding: 10px 12px; font-size: 14px; outline: none;
}
.btn-primary {
  width: 100%; padding: 16px; background: #34A853; color: white;
  border: none; border-radius: 14px; font-size: 16px; font-weight: 700; cursor: pointer;
}
.btn-primary:active { opacity: .85; }

.chart-section {
  background: white; border-radius: 14px; padding: 16px;
  margin-bottom: 14px; box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
.chart-title { font-size: 14px; font-weight: 600; margin-bottom: 12px; }
.sparkline { display: flex; align-items: flex-end; gap: 4px; height: 60px; }
.spark-bar {
  flex: 1; background: #34A853; border-radius: 3px 3px 0 0;
  min-width: 6px; transition: height .3s;
}
.spark-bar.sleep { background: #4285F4; }
.spark-bar.mood { background: #FF6D00; }
.chart-labels { display: flex; justify-content: space-between; margin-top: 4px; }
.chart-label { font-size: 9px; color: #8e8e93; }

.toast {
  position: fixed; bottom: 80px; left: 50%; transform: translateX(-50%);
  background: #1c1c1e; color: white; padding: 10px 20px; border-radius: 20px;
  font-size: 13px; font-weight: 500; z-index: 200;
  opacity: 0; transition: opacity .3s; pointer-events: none;
}
.toast.show { opacity: 1; }
.loading-spinner { text-align: center; color: #8e8e93; padding: 40px; font-size: 14px; }
```

---

## Task 4: PWA App Logic — App.html

**Files:**
- Create: `phase2/App.html`

- [ ] **Step 1: Create `phase2/App.html`**

```javascript
const state = { tasks: [], doneTasks: [], supplies: [], healthLogged: false };

const CAT_ICONS = { Health: '💊', Finance: '💳', Home: '🏠', Garden: '🌱', Service: '🔧' };
const STATUS_LABELS = { ok: 'Stock OK', order_soon: 'Order Soon', order_now: 'Order NOW', needs_setup: 'Setup needed' };
const STATUS_BADGES = { ok: 'badge-ok', order_soon: 'badge-soon', order_now: 'badge-now', needs_setup: 'badge-setup' };

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('today-date').textContent =
    new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'short' });
  loadTodayTasks();
});

function showScreen(name) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('screen-' + name).classList.add('active');
  document.getElementById('nav-' + name).classList.add('active');
  if (name === 'supplies') loadSupplies();
  if (name === 'charts') loadCharts();
}

function loadTodayTasks() {
  document.getElementById('tasks-loading').style.display = 'block';
  document.getElementById('tasks-list').innerHTML = '';
  google.script.run
    .withSuccessHandler(renderTasks)
    .withFailureHandler(() => showToast('Error loading tasks'))
    .getTodayTasks();
}

function renderTasks(tasks) {
  document.getElementById('tasks-loading').style.display = 'none';
  const list = document.getElementById('tasks-list');
  if (!tasks.length) {
    list.innerHTML = '<p style="text-align:center;color:#8e8e93;padding:40px;font-size:14px;">All done for today!</p>';
    return;
  }
  list.innerHTML = tasks.map(t => `
    <div class="task-card" id="task-${t.id}" onclick="doneTask(${t.id}, '${t.title.replace(/'/g,"\\'")}', this)">
      <div class="task-icon cat-${t.category.toLowerCase()}">${CAT_ICONS[t.category] || '📌'}</div>
      <div class="task-info">
        <div class="task-title">${t.title}</div>
        <div class="task-sub">${t.frequency} · ${t.category}</div>
      </div>
      <div class="task-check">○</div>
    </div>
  `).join('');
}

function doneTask(id, title, el) {
  el.style.transition = 'transform .3s, opacity .3s';
  el.style.transform = 'translateX(80px)';
  el.style.opacity = '0';
  setTimeout(() => {
    el.remove();
    showToast('Done: ' + title);
    const doneList = document.getElementById('tasks-done-list');
    document.getElementById('tasks-done').style.display = 'block';
    doneList.innerHTML += `<div class="task-card muted">
      <div class="task-info"><div class="task-title">${title}</div></div>
      <div class="task-check">✓</div>
    </div>`;
  }, 280);
  google.script.run
    .withFailureHandler(() => showToast('Save failed'))
    .markTaskDone(id);
}

function loadSupplies() {
  document.getElementById('supplies-loading').style.display = 'block';
  document.getElementById('supplies-list').innerHTML = '';
  google.script.run
    .withSuccessHandler(renderSupplies)
    .withFailureHandler(() => showToast('Error loading supplies'))
    .getSupplies();
}

function renderSupplies(supplies) {
  document.getElementById('supplies-loading').style.display = 'none';
  const list = document.getElementById('supplies-list');
  if (!supplies.length) {
    list.innerHTML = '<p style="text-align:center;color:#8e8e93;padding:40px">No supplies set up yet</p>';
    return;
  }
  list.innerHTML = supplies.map(s => {
    const daysText = s.daysUntilReorder !== null
      ? (s.daysUntilReorder <= 0 ? 'Overdue to order!' : 'Order in ' + s.daysUntilReorder + ' days')
      : 'Set last order date to activate';
    const showBtn = s.status === 'order_now' || s.status === 'order_soon' || s.status === 'needs_setup';
    return `<div class="supply-card">
      <div class="supply-header">
        <div class="supply-name">${s.item}</div>
        <span class="supply-badge ${STATUS_BADGES[s.status]}">${STATUS_LABELS[s.status]}</span>
      </div>
      <div class="supply-detail">${daysText} · Qty: ${s.qty} · Usage: ${s.dailyUsage}/day</div>
      ${showBtn ? `<button class="btn-order" onclick="orderDone(${s.id}, '${s.item.replace(/'/g,"\\'")}', this)">Ordered</button>` : ''}
    </div>`;
  }).join('');
}

function orderDone(id, item, btn) {
  btn.disabled = true;
  btn.textContent = 'Saving...';
  google.script.run
    .withSuccessHandler(() => { showToast('Ordered ' + item + '! Reorder date set.'); loadSupplies(); })
    .withFailureHandler(() => { btn.disabled = false; btn.textContent = 'Ordered'; showToast('Save failed'); })
    .markOrdered(id);
}

function updateVal(metric, val) {
  document.getElementById('val-' + metric).textContent = val;
}

function submitHealth() {
  const entry = {
    energy: +document.getElementById('slider-energy').value,
    sleep: +document.getElementById('slider-sleep').value,
    focus: +document.getElementById('slider-focus').value,
    mood: +document.getElementById('slider-mood').value,
    confidence: +document.getElementById('slider-confidence').value,
    notes: document.getElementById('health-notes').value
  };
  const btn = document.querySelector('.btn-primary');
  btn.disabled = true;
  btn.textContent = 'Saving...';
  google.script.run
    .withSuccessHandler(r => { showToast('Health logged! Overall: ' + r.overall + '/10'); btn.textContent = 'Saved'; btn.style.background = '#2e7d32'; })
    .withFailureHandler(() => { btn.disabled = false; btn.textContent = 'Save Health Log'; showToast('Save failed'); })
    .logHealth(entry);
}

function loadCharts() {
  document.getElementById('charts-loading').style.display = 'block';
  google.script.run
    .withSuccessHandler(renderCharts)
    .withFailureHandler(() => showToast('Error loading trends'))
    .getHealthTrends();
}

function renderCharts(rows) {
  document.getElementById('charts-loading').style.display = 'none';
  if (!rows.length) {
    document.getElementById('charts-container').innerHTML =
      '<p style="text-align:center;color:#8e8e93;padding:40px">No health data yet. Log some entries first!</p>';
    return;
  }
  const last14 = rows.slice(-14);
  const metrics = [
    { key: 'energy', label: 'Energy (1-10)', color: '' },
    { key: 'sleep', label: 'Sleep (hrs)', color: 'sleep' },
    { key: 'mood', label: 'Mood (1-10)', color: 'mood' },
    { key: 'overall', label: 'Overall Score', color: '' }
  ];
  document.getElementById('charts-container').innerHTML = metrics.map(m => {
    const vals = last14.map(r => r[m.key]);
    const max = Math.max(...vals, 1);
    const avg = (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1);
    const bars = vals.map(v => {
      const pct = Math.round((v / max) * 100);
      return '<div class="spark-bar ' + m.color + '" style="height:' + pct + '%"></div>';
    }).join('');
    return '<div class="chart-section"><div class="chart-title">' + m.label + '</div>' +
      '<div class="sparkline">' + bars + '</div>' +
      '<div class="chart-labels">' +
        '<span class="chart-label">' + new Date(last14[0].date).toLocaleDateString('en-GB', {day:'numeric',month:'short'}) + '</span>' +
        '<span class="chart-label">avg ' + avg + '</span>' +
        '<span class="chart-label">' + new Date(last14[last14.length-1].date).toLocaleDateString('en-GB', {day:'numeric',month:'short'}) + '</span>' +
      '</div></div>';
  }).join('');
}

function showToast(msg) {
  let t = document.querySelector('.toast');
  if (!t) { t = document.createElement('div'); t.className = 'toast'; document.body.appendChild(t); }
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2500);
}
```

---

## Task 5: Phase 2 Deploy Instructions

**Files:**
- Create: `phase2/DEPLOY.md`

- [ ] **Step 1: Create `phase2/DEPLOY.md`**

```
Phase 2 Deployment — Web App

SETUP
1. Open the ElderCare Google Sheet from Phase 1
2. Extensions -> Apps Script
3. Create 4 files in Apps Script:
   - Code.gs    -> paste phase2/Code.gs
   - Index.html -> paste phase2/Index.html
   - Styles.html -> paste phase2/Styles.html
   - App.html   -> paste phase2/App.html
4. In Code.gs line 1: replace PASTE_YOUR_SHEET_ID_HERE with your Sheet ID

DEPLOY AS WEB APP
1. Click Deploy -> New deployment
2. Type: Web App
3. Execute as: Me
4. Who has access: Anyone (or "Anyone with Google account")
5. Click Deploy -> copy the Web App URL

ADD TO PHONE HOME SCREEN
iPhone: Open URL in Safari -> Share -> Add to Home Screen
Android: Open URL in Chrome -> three dots -> Add to Home Screen

TEST CHECKLIST
- Today screen loads and shows seeded tasks
- Tap a task card — it slides right and appears in Done section
- Supplies screen loads showing Magnesium + Multivitamin with "needs_setup" badge
- Tap "Ordered" on a supply -> badge changes to "Stock OK"
- Health sliders move and values update live
- Save Health Log -> toast shows overall score
- Trends screen shows "No health data yet" until 1+ entries exist
```

---

## Task 6: Telegram Bot (Phase 3)

**Files:**
- Create: `phase3/Telegram.gs`
- Create: `phase3/DEPLOY.md`

- [ ] **Step 1: Create `phase3/Telegram.gs`**

```javascript
const BOT_TOKEN = 'PASTE_TELEGRAM_BOT_TOKEN_HERE';
const CHAT_ID   = 'PASTE_TELEGRAM_CHAT_ID_HERE';

function sendTelegram(message) {
  const url = 'https://api.telegram.org/bot' + BOT_TOKEN + '/sendMessage';
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: CHAT_ID, text: message, parse_mode: 'HTML' })
  });
}

function dailyMorningAlert() {
  const tasks = getTodayTasks();
  const supplies = getSupplies();

  let msg = '<b>ElderCare — Good Morning</b>\n\n';

  if (tasks.length) {
    msg += '<b>Today\'s Tasks:</b>\n';
    tasks.forEach(t => { msg += '- ' + t.title + ' (' + t.category + ')\n'; });
  } else {
    msg += 'No tasks due today!\n';
  }

  const urgent = supplies.filter(s => s.status === 'order_now' || s.status === 'order_soon');
  if (urgent.length) {
    msg += '\n<b>Supplies to Order:</b>\n';
    urgent.forEach(s => {
      const days = s.daysUntilReorder !== null
        ? ' (' + Math.abs(s.daysUntilReorder) + ' days ' + (s.daysUntilReorder <= 0 ? 'overdue' : 'left') + ')'
        : '';
      msg += '- ' + s.item + days + ' -> ' + (s.status === 'order_now' ? 'ORDER NOW' : 'Order soon') + '\n';
    });
  }

  msg += '\nOpen ElderCare app to mark tasks done.';
  sendTelegram(msg);
}

function weeklySummaryAlert() {
  const rows = getHealthTrends();
  const last7 = rows.slice(-7);
  if (!last7.length) return;

  const avg = (key) => (last7.reduce((s, r) => s + (r[key] || 0), 0) / last7.length).toFixed(1);
  let msg = '<b>ElderCare — Weekly Summary</b>\n\n';
  msg += 'Energy avg: ' + avg('energy') + '/10\n';
  msg += 'Sleep avg: ' + avg('sleep') + ' hrs\n';
  msg += 'Mood avg: ' + avg('mood') + '/10\n';
  msg += 'Overall avg: ' + avg('overall') + '/10\n';

  const prevWeek = rows.slice(-14, -7);
  if (prevWeek.length) {
    const prevAvg = (prevWeek.reduce((s, r) => s + (r.overall || 0), 0) / prevWeek.length).toFixed(1);
    const diff = (avg('overall') - prevAvg).toFixed(1);
    msg += '\nVs last week: ' + (diff > 0 ? '+' : '') + diff;
  }
  sendTelegram(msg);
}
```

- [ ] **Step 2: Create `phase3/DEPLOY.md`**

```
Phase 3 Deployment — Telegram Bot

GET BOT TOKEN
1. Open Telegram -> search @BotFather
2. Send: /newbot
3. Name it: ElderCare Bot
4. Username: eldercare_yourname_bot
5. Copy the token BotFather gives you

GET YOUR CHAT ID
1. Start a chat with your new bot (search by username, tap Start)
2. Send it any message (e.g. "hello")
3. Open this URL in browser (replace TOKEN with your actual token):
   https://api.telegram.org/botTOKEN/getUpdates
4. Find "chat":{"id": 123456789} — that number is your CHAT_ID

ADD TO APPS SCRIPT
1. Open ElderCare Apps Script (same project as Phase 2)
2. Create new file: Telegram.gs
3. Paste phase3/Telegram.gs content
4. Replace BOT_TOKEN and CHAT_ID values at top of file

SET DAILY TRIGGER (8am alert)
1. Apps Script -> Triggers (clock icon) -> Add Trigger
2. Function: dailyMorningAlert
3. Event source: Time-driven
4. Type: Day timer -> 8am to 9am
5. Save

SET WEEKLY TRIGGER (Sunday summary)
1. Add another Trigger
2. Function: weeklySummaryAlert
3. Type: Week timer -> Every Sunday -> 9am to 10am
4. Save

TEST
Run dailyMorningAlert manually once -> check Telegram for the message.
```

---

## Self-Review

**Spec coverage:**
- Swipe/tap task cards -> Task 3+4 (slide animation + doneTask) ✓
- Smart reorder logic (qty / daily_usage - delivery_days) -> Task 2 Step 5 markOrdered ✓
- Health log with sliders (energy, sleep, focus, mood, confidence) -> Task 3 Index.html + Task 4 submitHealth ✓
- Trend charts (last 14 days sparklines) -> Task 4 renderCharts ✓
- Telegram morning alert + weekly summary -> Task 6 ✓
- Google Calendar deferred -> noted in README ✓
- Amazon monthly ordering mapped to SUPPLIES tab + markOrdered reorder date calc ✓
- Seed data includes magnesium + multivitamin ✓

**Type consistency:**
- getTodayTasks / markTaskDone / getSupplies / markOrdered / logHealth / getHealthTrends defined in Task 2, called identically in Task 4 (App.html) and Task 6 (Telegram.gs) ✓
