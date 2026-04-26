// Fill in these two values during deployment (see phase3/DEPLOY.md)
const BOT_TOKEN = 'PASTE_TELEGRAM_BOT_TOKEN_HERE';
const CHAT_ID   = 'PASTE_TELEGRAM_CHAT_ID_HERE';

function sendTelegram(message) {
  const url = 'https://api.telegram.org/bot' + BOT_TOKEN + '/sendMessage';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: CHAT_ID, text: message, parse_mode: 'HTML' })
    });
  } catch(e) {
    console.error('Telegram send failed:', e);
  }
}

function dailyMorningAlert() {
  const tasks = getTodayTasks();
  const supplies = getSupplies();

  let msg = '<b>ElderCare - Good Morning</b>\n\n';

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

  msg += '\nOpen the ElderCare app to mark tasks done.';
  sendTelegram(msg);
}

function weeklySummaryAlert() {
  const rows = getHealthTrends();
  const last7 = rows.slice(-7);
  if (!last7.length) return;

  const avg = (key) => (last7.reduce((s, r) => s + (Number(r[key]) || 0), 0) / last7.length).toFixed(1);
  let msg = '<b>ElderCare - Weekly Summary</b>\n\n';
  msg += 'Energy avg: ' + avg('energy') + '/10\n';
  msg += 'Sleep avg: ' + avg('sleep') + ' hrs\n';
  msg += 'Mood avg: ' + avg('mood') + '/10\n';
  msg += 'Overall avg: ' + avg('overall') + '/10\n';

  const prevWeek = rows.slice(-14, -7);
  if (prevWeek.length) {
    const prevAvg = (prevWeek.reduce((s, r) => s + (Number(r.overall) || 0), 0) / prevWeek.length).toFixed(1);
    const diff = (parseFloat(avg('overall')) - parseFloat(prevAvg)).toFixed(1);
    msg += '\nVs last week: ' + (diff > 0 ? '+' : '') + diff;
  }
  sendTelegram(msg);
}
