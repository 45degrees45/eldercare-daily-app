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
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const log = ss.getSheetByName('LOG');
    if (!log) return; // LOG sheet missing — degrade gracefully
    log.appendRow([new Date(), action, detail]);
  } catch(e) {
    console.error('_log failed:', e);
  }
}

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
      // 'once' frequency: shown when lastDone is empty (handled above), never reshown after completion
    }

    if (isDueToday) {
      tasks.push({ id, title, category, frequency, lastDone: lastDone ? new Date(lastDone).toISOString() : '', row: i + 1 });
    }
  }
  return tasks;
}

function markTaskDone(taskId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('TASKS');
  const data = sheet.getDataRange().getValues();
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(taskId)) {
      sheet.getRange(i + 1, 5).setValue(today);
      _log('markTaskDone', 'ID=' + taskId + ' title=' + data[i][1]);
      return { success: true, title: data[i][1] };
    }
  }
  return { success: false, error: 'Task not found' };
}

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

function markOrdered(supplyId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SUPPLIES');
  const data = sheet.getDataRange().getValues();
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(supplyId)) {
      const qty = data[i][3];
      const dailyUsage = data[i][4];
      const deliveryDays = data[i][5];
      if (!dailyUsage || dailyUsage <= 0) {
        return { success: false, error: 'Daily usage must be greater than 0' };
      }
      sheet.getRange(i + 1, 7).setValue(today);
      const reorderDate = new Date(today.getTime() + ((qty / dailyUsage) - deliveryDays) * 86400000);
      sheet.getRange(i + 1, 8).setValue(reorderDate);
      _log('markOrdered', 'ID=' + supplyId + ' item=' + data[i][1]);
      return { success: true, item: data[i][1], reorderDate: reorderDate.toISOString() };
    }
  }
  return { success: false, error: 'Supply not found' };
}

function logHealth(entry) {
  const required = ['energy', 'sleep', 'focus', 'mood', 'confidence'];
  for (const key of required) {
    if (!Number.isFinite(entry[key])) {
      return { success: false, error: 'Invalid value for ' + key };
    }
  }
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
