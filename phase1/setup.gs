function setupElderCare() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- TASKS sheet ---
  let tasks = ss.getSheetByName('TASKS') || ss.insertSheet('TASKS');
  tasks.clear();
  tasks.getRange(1, 1, 1, 9).setValues([[
    'ID', 'Title', 'Category', 'Frequency',
    'Last Done', 'Next Due', 'Status', 'Notes', 'Active'
  ]]);
  tasks.getRange('A1:I1').setFontWeight('bold').setBackground('#34A853').setFontColor('white');
  tasks.setFrozenRows(1);
  tasks.setColumnWidth(2, 200); // Title column wider
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
  supplies.clear();
  supplies.getRange(1, 1, 1, 9).setValues([[
    'ID', 'Item', 'Category', 'Qty On Hand',
    'Daily Usage', 'Delivery Days', 'Last Ordered', 'Reorder Date', 'Notes'
  ]]);
  supplies.getRange('A1:I1').setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  supplies.setFrozenRows(1);
  supplies.setColumnWidth(2, 180); // Item column wider
  const supCatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Vitamin', 'Mineral', 'Medicine', 'Grocery', 'Household'], true).build();
  supplies.getRange('C2:C200').setDataValidation(supCatRule);

  // --- HEALTH_LOG sheet ---
  let health = ss.getSheetByName('HEALTH_LOG') || ss.insertSheet('HEALTH_LOG');
  health.clear();
  health.getRange(1, 1, 1, 8).setValues([[
    'Date', 'Energy (1-10)', 'Sleep (hrs)', 'Focus (1-10)',
    'Mood (1-10)', 'Confidence (1-10)', 'Overall', 'Notes'
  ]]);
  health.getRange('A1:H1').setFontWeight('bold').setBackground('#FF6D00').setFontColor('white');
  health.setFrozenRows(1);

  // --- LOG sheet (hidden) ---
  let log = ss.getSheetByName('LOG') || ss.insertSheet('LOG');
  log.clear();
  log.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Action', 'Detail']]);
  ss.setActiveSheet(tasks);
  try { ss.getSheetByName('LOG').hideSheet(); } catch(e) { console.error('Could not hide LOG sheet:', e); }

  // --- Seed starter tasks ---
  // Only seed if sheet is empty (first run)
  if (tasks.getLastRow() < 2) {
    tasks.getRange('A2:I6').setValues([
      [1, 'Take morning vitamins', 'Health', 'daily', '', '', 'Active', '', true],
      [2, 'Take magnesium', 'Health', 'daily', '', '', 'Active', '', true],
      [3, 'Pay electricity bill', 'Finance', 'monthly', '', '', 'Active', '', true],
      [4, 'Water the plants', 'Garden', 'weekly', '', '', 'Active', '', true],
      [5, 'Clean the house', 'Home', 'weekly', '', '', 'Active', '', true],
    ]);
  }

  // --- Seed starter supplies ---
  // Only seed if sheet is empty (first run)
  if (supplies.getLastRow() < 2) {
    supplies.getRange('A2:I3').setValues([
      [1, 'Magnesium tablets', 'Mineral', 50, 1, 7, '', '', ''],
      [2, 'Multivitamin', 'Vitamin', 60, 1, 7, '', '', ''],
    ]);
  }

  SpreadsheetApp.getUi().alert('ElderCare sheet setup complete!');
}
