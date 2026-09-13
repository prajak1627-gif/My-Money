/**
 * ==============================================
 *  บันทึกรายรับ-รายจ่าย — Backend (Google Apps Script + Google Sheets)
 * ==============================================
 * วิธีติดตั้ง: ดูไฟล์ README.md
 */

const SHEET_NAMES = {
  WALLETS: 'Wallets',
  CATEGORIES: 'Categories',
  TRANSACTIONS: 'Transactions',
  PERIODS: 'Periods'
};

const SHEET_HEADERS = {
  Wallets: ['ID', 'Name', 'Icon', 'Color', 'InitialBalance', 'CreatedAt'],
  Categories: ['ID', 'Name', 'Type', 'Icon', 'Color', 'CreatedAt'],
  Transactions: ['ID', 'Date', 'Type', 'WalletID', 'CategoryID', 'Amount', 'Note', 'PeriodID', 'CreatedAt'],
  Periods: ['ID', 'Name', 'StartDate', 'EndDate', 'Status', 'CreatedAt']
};

/* ---------------------------------------------
 *  Web App entry point
 * -------------------------------------------- */
function doGet(e) {
  setupSheets(); // ensure sheets exist on first load
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('บันทึกรายรับ-รายจ่าย')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------------------------------------
 *  JSON API (สำหรับหน้าเว็บที่ฝากไว้ที่อื่น เช่น GitHub Pages
 *  เรียกผ่าน fetch() ด้วย POST, Content-Type: text/plain
 *  body: { "action": "ชื่อฟังก์ชัน", "params": [ ...อาร์กิวเมนต์ ] }
 * -------------------------------------------- */
const API_ALLOWED_FUNCTIONS_ = {
  getInitialData, getDashboardData, getWalletCategoryBreakdown,
  getWallets, addWallet, updateWallet, deleteWallet,
  getCategories, addCategory, updateCategory, deleteCategory,
  getPeriods, addPeriod, updatePeriod, deletePeriod, closePeriodAndStartNew,
  getTransactions, addTransaction, updateTransaction, deleteTransaction
};

function doPost(e) {
  setupSheets();

  let action, params;
  try {
    const body = JSON.parse(e.postData.contents);
    action = body.action;
    params = body.params || [];
  } catch (err) {
    return jsonResponse_({ success: false, error: 'รูปแบบคำขอไม่ถูกต้อง (invalid request body)' });
  }

  const fn = API_ALLOWED_FUNCTIONS_[action];
  if (!fn) {
    return jsonResponse_({ success: false, error: 'ไม่รู้จักคำสั่ง: ' + action });
  }

  try {
    const result = fn.apply(null, params);
    return jsonResponse_({ success: true, data: result });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message || String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ---------------------------------------------
 *  Sheet setup / seeding
 * -------------------------------------------- */
function setupSheets() {
  const ss = getSS();
  Object.keys(SHEET_HEADERS).forEach(name => ensureSheet_(ss, name, SHEET_HEADERS[name]));
  seedDefaultsIfEmpty_();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function seedDefaultsIfEmpty_() {
  const ss = getSS();

  const walletSheet = ss.getSheetByName(SHEET_NAMES.WALLETS);
  if (walletSheet.getLastRow() < 2) {
    addWallet({ name: 'เงินสด', icon: '💵', color: '#2F7A56', initialBalance: 0 });
    addWallet({ name: 'บัญชีธนาคาร', icon: '🏦', color: '#16302A', initialBalance: 0 });
  }

  const catSheet = ss.getSheetByName(SHEET_NAMES.CATEGORIES);
  if (catSheet.getLastRow() < 2) {
    const defaultExpense = [
      ['อาหาร', '🍜', '#B44B36'],
      ['เดินทาง', '🚗', '#B8912B'],
      ['ช้อปปิ้ง', '🛍️', '#7A5AB8'],
      ['บิล/ค่าน้ำค่าไฟ', '💡', '#3E7A9A'],
      ['บันเทิง', '🎬', '#B8506F'],
      ['สุขภาพ', '💊', '#4E9A6E'],
      ['อื่นๆ', '📦', '#8A8A82']
    ];
    defaultExpense.forEach(c => addCategory({ name: c[0], type: 'expense', icon: c[1], color: c[2] }));

    const defaultIncome = [
      ['เงินเดือน', '💼', '#2F7A56'],
      ['โบนัส', '🎁', '#2F9A6E'],
      ['รายได้เสริม', '➕', '#3E7A56']
    ];
    defaultIncome.forEach(c => addCategory({ name: c[0], type: 'income', icon: c[1], color: c[2] }));
  }

  const periodSheet = ss.getSheetByName(SHEET_NAMES.PERIODS);
  if (periodSheet.getLastRow() < 2) {
    const today = new Date();
    const start = new Date(today.getFullYear(), today.getMonth(), 1);
    addPeriod({ name: 'รอบเดือน ' + Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM yyyy'), startDate: start, endDate: '', status: 'active' });
  }
}

/* ---------------------------------------------
 *  Generic sheet helpers
 * -------------------------------------------- */
function readSheet_(name) {
  const sheet = getSS().getSheetByName(name);
  const range = sheet.getDataRange().getValues();
  const headers = range.shift();
  return range
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

function findRowById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based row index
  }
  return -1;
}

function toIsoDate_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

/* ---------------------------------------------
 *  WALLETS  (กระเป๋า)
 * -------------------------------------------- */
function getWallets() {
  return readSheet_(SHEET_NAMES.WALLETS).map(w => ({
    id: w.ID, name: w.Name, icon: w.Icon, color: w.Color,
    initialBalance: Number(w.InitialBalance) || 0
  }));
}

function addWallet(data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.WALLETS);
  const id = Utilities.getUuid();
  sheet.appendRow([id, data.name, data.icon || '💰', data.color || '#16302A', Number(data.initialBalance) || 0, new Date()]);
  return id;
}

function updateWallet(id, data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.WALLETS);
  const row = findRowById_(sheet, id);
  if (row < 0) throw new Error('ไม่พบกระเป๋านี้');
  sheet.getRange(row, 2, 1, 4).setValues([[data.name, data.icon, data.color, Number(data.initialBalance) || 0]]);
  return true;
}

function deleteWallet(id) {
  const txs = readSheet_(SHEET_NAMES.TRANSACTIONS);
  if (txs.some(t => String(t.WalletID) === String(id))) {
    throw new Error('ลบไม่ได้ เนื่องจากมีรายการที่ผูกกับกระเป๋านี้อยู่');
  }
  const sheet = getSS().getSheetByName(SHEET_NAMES.WALLETS);
  const row = findRowById_(sheet, id);
  if (row > 0) sheet.deleteRow(row);
  return true;
}

/* ---------------------------------------------
 *  CATEGORIES  (หมวดหมู่)
 * -------------------------------------------- */
function getCategories() {
  return readSheet_(SHEET_NAMES.CATEGORIES).map(c => ({
    id: c.ID, name: c.Name, type: c.Type, icon: c.Icon, color: c.Color
  }));
}

function addCategory(data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.CATEGORIES);
  const id = Utilities.getUuid();
  sheet.appendRow([id, data.name, data.type, data.icon || '📁', data.color || '#8A8A82', new Date()]);
  return id;
}

function updateCategory(id, data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.CATEGORIES);
  const row = findRowById_(sheet, id);
  if (row < 0) throw new Error('ไม่พบหมวดหมู่นี้');
  sheet.getRange(row, 2, 1, 4).setValues([[data.name, data.type, data.icon, data.color]]);
  return true;
}

function deleteCategory(id) {
  const txs = readSheet_(SHEET_NAMES.TRANSACTIONS);
  if (txs.some(t => String(t.CategoryID) === String(id))) {
    throw new Error('ลบไม่ได้ เนื่องจากมีรายการที่ผูกกับหมวดหมู่นี้อยู่');
  }
  const sheet = getSS().getSheetByName(SHEET_NAMES.CATEGORIES);
  const row = findRowById_(sheet, id);
  if (row > 0) sheet.deleteRow(row);
  return true;
}

/* ---------------------------------------------
 *  PERIODS  (รอบบัญชี / ตัดรอบ)
 * -------------------------------------------- */
function getPeriods() {
  return readSheet_(SHEET_NAMES.PERIODS)
    .map(p => ({
      id: p.ID, name: p.Name,
      startDate: toIsoDate_(p.StartDate),
      endDate: p.EndDate ? toIsoDate_(p.EndDate) : '',
      status: p.Status
    }))
    .sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
}

function addPeriod(data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.PERIODS);
  const id = Utilities.getUuid();
  sheet.appendRow([
    id, data.name,
    data.startDate ? new Date(data.startDate) : new Date(),
    data.endDate ? new Date(data.endDate) : '',
    data.status || 'active',
    new Date()
  ]);
  return id;
}

function updatePeriod(id, data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.PERIODS);
  const row = findRowById_(sheet, id);
  if (row < 0) throw new Error('ไม่พบรอบบัญชีนี้');
  sheet.getRange(row, 2, 1, 4).setValues([[
    data.name,
    data.startDate ? new Date(data.startDate) : '',
    data.endDate ? new Date(data.endDate) : '',
    data.status || 'active'
  ]]);
  return true;
}

// ตัดรอบบัญชี: ปิดรอบปัจจุบัน (ใส่วันที่สิ้นสุด) แล้วเปิดรอบใหม่ต่อจากวันถัดไปโดยอัตโนมัติ
function closePeriodAndStartNew(id, endDate, newPeriodName) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.PERIODS);
  const row = findRowById_(sheet, id);
  if (row < 0) throw new Error('ไม่พบรอบบัญชีนี้');

  const end = endDate ? new Date(endDate) : new Date();
  sheet.getRange(row, 4).setValue(end);       // EndDate
  sheet.getRange(row, 5).setValue('closed');  // Status

  const nextStart = new Date(end);
  nextStart.setDate(nextStart.getDate() + 1);
  const name = newPeriodName || ('รอบเดือน ' + Utilities.formatDate(nextStart, Session.getScriptTimeZone(), 'MMMM yyyy'));
  return addPeriod({ name: name, startDate: nextStart, endDate: '', status: 'active' });
}

function deletePeriod(id) {
  const txs = readSheet_(SHEET_NAMES.TRANSACTIONS);
  if (txs.some(t => String(t.PeriodID) === String(id))) {
    throw new Error('ลบไม่ได้ เนื่องจากมีรายการที่ผูกกับรอบบัญชีนี้อยู่');
  }
  const sheet = getSS().getSheetByName(SHEET_NAMES.PERIODS);
  const row = findRowById_(sheet, id);
  if (row > 0) sheet.deleteRow(row);
  return true;
}

/* ---------------------------------------------
 *  TRANSACTIONS  (รายรับ-รายจ่าย)
 * -------------------------------------------- */
function findPeriodIdForDate_(dateObj) {
  const periods = readSheet_(SHEET_NAMES.PERIODS);
  const d = new Date(dateObj);
  for (const p of periods) {
    const start = new Date(p.StartDate);
    const end = p.EndDate ? new Date(p.EndDate) : null;
    if (d >= start && (!end || d <= end)) return p.ID;
  }
  return '';
}

function addTransaction(data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const id = Utilities.getUuid();
  const date = data.date ? new Date(data.date) : new Date();
  const periodId = data.periodId || findPeriodIdForDate_(date);
  sheet.appendRow([
    id, date, data.type, data.walletId, data.categoryId,
    Number(data.amount) || 0, data.note || '', periodId, new Date()
  ]);
  return id;
}

function updateTransaction(id, data) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const row = findRowById_(sheet, id);
  if (row < 0) throw new Error('ไม่พบรายการนี้');
  const date = data.date ? new Date(data.date) : new Date();
  const periodId = data.periodId || findPeriodIdForDate_(date);
  sheet.getRange(row, 2, 1, 7).setValues([[
    date, data.type, data.walletId, data.categoryId,
    Number(data.amount) || 0, data.note || '', periodId
  ]]);
  return true;
}

function deleteTransaction(id) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const row = findRowById_(sheet, id);
  if (row > 0) sheet.deleteRow(row);
  return true;
}

function getTransactions(filters) {
  filters = filters || {};
  const wallets = getWallets();
  const categories = getCategories();
  const allTxRaw = readSheet_(SHEET_NAMES.TRANSACTIONS);

  // คำนวณยอดคงเหลือสะสม (คงเหลือ) ของแต่ละกระเป๋า ต้องไล่ตามลำดับเวลาจริง (เก่า -> ใหม่)
  // ก่อนตัดกรองใดๆ เพื่อให้ยอดคงเหลือถูกต้องเสมอ ไม่ว่าจะกรองแบบไหนภายหลัง
  const chronological = allTxRaw.slice().sort((a, b) => new Date(a.Date) - new Date(b.Date));
  const runningByWallet = {};
  wallets.forEach(w => (runningByWallet[w.id] = w.initialBalance));
  const balanceById = {};
  chronological.forEach(t => {
    const wId = t.WalletID;
    if (!(wId in runningByWallet)) runningByWallet[wId] = 0;
    if (t.Type === 'income') runningByWallet[wId] += Number(t.Amount) || 0;
    else runningByWallet[wId] -= Number(t.Amount) || 0;
    balanceById[t.ID] = runningByWallet[wId];
  });

  let txs = allTxRaw;
  if (filters.periodId) txs = txs.filter(t => String(t.PeriodID) === String(filters.periodId));
  if (filters.walletId) txs = txs.filter(t => String(t.WalletID) === String(filters.walletId));
  if (filters.categoryId) txs = txs.filter(t => String(t.CategoryID) === String(filters.categoryId));
  if (filters.type) txs = txs.filter(t => t.Type === filters.type);
  if (filters.dateStart) {
    const start = new Date(filters.dateStart);
    txs = txs.filter(t => new Date(t.Date) >= start);
  }
  if (filters.dateEnd) {
    const end = new Date(filters.dateEnd);
    end.setHours(23, 59, 59, 999);
    txs = txs.filter(t => new Date(t.Date) <= end);
  }

  const mapped = txs.map(t => {
    const wallet = wallets.find(w => w.id === t.WalletID);
    const cat = categories.find(c => c.id === t.CategoryID);
    return {
      id: t.ID,
      date: toIsoDate_(t.Date),
      type: t.Type,
      walletId: t.WalletID,
      walletName: wallet ? wallet.name : '-',
      walletIcon: wallet ? wallet.icon : '❓',
      categoryId: t.CategoryID,
      categoryName: cat ? cat.name : 'ไม่ระบุ',
      categoryIcon: cat ? cat.icon : '📁',
      categoryColor: cat ? cat.color : '#8A8A82',
      amount: Number(t.Amount) || 0,
      note: t.Note,
      periodId: t.PeriodID,
      balance: balanceById[t.ID] !== undefined ? balanceById[t.ID] : (wallet ? wallet.initialBalance : 0)
    };
  });

  const sortAsc = filters.sort === 'asc';
  mapped.sort((a, b) => sortAsc ? new Date(a.date) - new Date(b.date) : new Date(b.date) - new Date(a.date));

  // เลขลำดับ (ID แถวที่แสดงในตาราง) นับจากรายการเก่าสุด = 1 เสมอ ไม่ว่าจะเรียงลำดับแบบไหน
  const chronoIndex = {};
  chronological.forEach((t, i) => (chronoIndex[t.ID] = i + 1));
  mapped.forEach(t => (t.rowNo = chronoIndex[t.id] || 0));

  return mapped;
}

/* ---------------------------------------------
 *  DASHBOARD
 * -------------------------------------------- */
function getInitialData() {
  return {
    wallets: getWallets(),
    categories: getCategories(),
    periods: getPeriods()
  };
}

function getDashboardData(periodId) {
  const wallets = getWallets();
  const categories = getCategories();
  const periods = getPeriods();
  const allTx = readSheet_(SHEET_NAMES.TRANSACTIONS);

  // ยอดคงเหลือของแต่ละกระเป๋า = ยอดตั้งต้น + รายรับทั้งหมด - รายจ่ายทั้งหมด (ตลอดกาล ไม่ขึ้นกับรอบบัญชี)
  const walletBalances = wallets.map(w => {
    let income = 0, expense = 0;
    allTx.forEach(t => {
      if (String(t.WalletID) === String(w.id)) {
        if (t.Type === 'income') income += Number(t.Amount);
        else expense += Number(t.Amount);
      }
    });
    return {
      id: w.id, name: w.name, icon: w.icon, color: w.color,
      balance: w.initialBalance + income - expense
    };
  });
  const totalBalance = walletBalances.reduce((s, w) => s + w.balance, 0);

  // เลือกรอบบัญชีที่จะสรุป
  let period = null;
  if (periodId) period = periods.find(p => p.id === periodId);
  if (!period) period = periods.find(p => p.status === 'active') || periods[0];

  let periodTx = allTx;
  if (period) {
    const start = new Date(period.startDate);
    const end = period.endDate ? new Date(period.endDate) : null;
    periodTx = allTx.filter(t => {
      const d = new Date(t.Date);
      return d >= start && (!end || d <= end);
    });
  }

  let totalIncome = 0, totalExpense = 0;
  const expenseByCategory = {};
  periodTx.forEach(t => {
    const amt = Number(t.Amount) || 0;
    if (t.Type === 'income') {
      totalIncome += amt;
    } else {
      totalExpense += amt;
      expenseByCategory[t.CategoryID] = (expenseByCategory[t.CategoryID] || 0) + amt;
    }
  });

  const donutData = Object.keys(expenseByCategory)
    .map(catId => {
      const cat = categories.find(c => c.id === catId);
      return {
        name: cat ? cat.name : 'ไม่ระบุ',
        icon: cat ? cat.icon : '📁',
        color: cat ? cat.color : '#8A8A82',
        value: expenseByCategory[catId]
      };
    })
    .sort((a, b) => b.value - a.value);

  // รายการล่าสุด — ใช้ allTx ที่โหลดไว้แล้ว ไม่อ่านชีตซ้ำ (เร็วขึ้นมาก)
  const chronological = allTx.slice().sort((a, b) => new Date(a.Date) - new Date(b.Date));
  const runningByWallet = {};
  wallets.forEach(w => (runningByWallet[w.id] = w.initialBalance));
  const balanceById = {};
  const rowNoById = {};
  chronological.forEach((t, i) => {
    if (!(t.WalletID in runningByWallet)) runningByWallet[t.WalletID] = 0;
    if (t.Type === 'income') runningByWallet[t.WalletID] += Number(t.Amount) || 0;
    else runningByWallet[t.WalletID] -= Number(t.Amount) || 0;
    balanceById[t.ID] = runningByWallet[t.WalletID];
    rowNoById[t.ID] = i + 1;
  });

  const recentTx = allTx
    .slice()
    .sort((a, b) => new Date(b.Date) - new Date(a.Date))
    .slice(0, 8)
    .map(t => {
      const wallet = wallets.find(w => w.id === t.WalletID);
      const cat = categories.find(c => c.id === t.CategoryID);
      return {
        id: t.ID,
        rowNo: rowNoById[t.ID] || 0,
        date: toIsoDate_(t.Date),
        type: t.Type,
        walletId: t.WalletID,
        walletName: wallet ? wallet.name : '-',
        walletIcon: wallet ? wallet.icon : '❓',
        categoryId: t.CategoryID,
        categoryName: cat ? cat.name : 'ไม่ระบุ',
        categoryIcon: cat ? cat.icon : '📁',
        categoryColor: cat ? cat.color : '#8A8A82',
        amount: Number(t.Amount) || 0,
        note: t.Note,
        periodId: t.PeriodID,
        balance: balanceById[t.ID] !== undefined ? balanceById[t.ID] : (wallet ? wallet.initialBalance : 0)
      };
    });

  // สรุปแต่ละกระเป๋าตามหมวดหมู่ — คำนวณจาก periodTx ที่มีอยู่แล้ว ไม่เรียกฟังก์ชันแยก (ลดรอบเรียก API ฝั่ง client)
  const walletBreakdown = wallets.map(w => {
    const walletTx = periodTx.filter(t => String(t.WalletID) === String(w.id));
    let wIncome = 0, wExpense = 0;
    const catMap = {};
    walletTx.forEach(t => {
      const amt = Number(t.Amount) || 0;
      if (t.Type === 'income') wIncome += amt; else wExpense += amt;
      if (!catMap[t.CategoryID]) catMap[t.CategoryID] = { income: 0, expense: 0 };
      if (t.Type === 'income') catMap[t.CategoryID].income += amt;
      else catMap[t.CategoryID].expense += amt;
    });
    const categoryBreakdown = Object.keys(catMap).map(catId => {
      const cat = categories.find(c => c.id === catId);
      return {
        id: catId,
        name: cat ? cat.name : 'ไม่ระบุ',
        icon: cat ? cat.icon : '📁',
        color: cat ? cat.color : '#8A8A82',
        type: cat ? cat.type : 'expense',
        income: catMap[catId].income,
        expense: catMap[catId].expense
      };
    }).sort((a, b) => (b.income + b.expense) - (a.income + a.expense));

    return { id: w.id, name: w.name, icon: w.icon, color: w.color, income: wIncome, expense: wExpense, categoryBreakdown };
  });

  return {
    walletBalances,
    totalBalance,
    totalIncome,
    totalExpense,
    net: totalIncome - totalExpense,
    donutData,
    recentTx,
    walletBreakdown,
    currentPeriod: period,
    periods
  };
}

/* ---------------------------------------------
 *  สรุปแต่ละกระเป๋าตามหมวดหมู่ (สำหรับ Dashboard)
 * -------------------------------------------- */
function getWalletCategoryBreakdown(periodId) {
  const wallets = getWallets();
  const categories = getCategories();
  const periods = getPeriods();
  const allTx = readSheet_(SHEET_NAMES.TRANSACTIONS);

  let period = periodId ? periods.find(p => p.id === periodId) : (periods.find(p => p.status === 'active') || periods[0]);
  let periodTx = allTx;
  if (period) {
    const start = new Date(period.startDate);
    const end = period.endDate ? new Date(period.endDate) : null;
    periodTx = allTx.filter(t => {
      const d = new Date(t.Date);
      return d >= start && (!end || d <= end);
    });
  }

  return wallets.map(w => {
    const walletTx = periodTx.filter(t => String(t.WalletID) === String(w.id));
    let income = 0, expense = 0;
    const catMap = {};

    walletTx.forEach(t => {
      const amt = Number(t.Amount) || 0;
      if (t.Type === 'income') income += amt;
      else expense += amt;

      if (!catMap[t.CategoryID]) catMap[t.CategoryID] = { income: 0, expense: 0 };
      if (t.Type === 'income') catMap[t.CategoryID].income += amt;
      else catMap[t.CategoryID].expense += amt;
    });

    const categoryBreakdown = Object.keys(catMap).map(catId => {
      const cat = categories.find(c => c.id === catId);
      return {
        id: catId,
        name: cat ? cat.name : 'ไม่ระบุ',
        icon: cat ? cat.icon : '📁',
        color: cat ? cat.color : '#8A8A82',
        type: cat ? cat.type : 'expense',
        income: catMap[catId].income,
        expense: catMap[catId].expense
      };
    }).sort((a, b) => (b.income + b.expense) - (a.income + a.expense));

    return {
      id: w.id,
      name: w.name,
      icon: w.icon,
      color: w.color,
      income,
      expense,
      categoryBreakdown
    };
  });
}
