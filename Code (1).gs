// ═══════════════════════════════════════════════════════════════
//  FINTRA BUDGET — Apps Script (Code.gs)
//  Versi dengan CORS headers untuk integrasi website budget.fintra.id
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// Sheet name constants — jangan diubah, dipakai oleh website
const SHEET_TRANSAKSI  = 'Transaksi';
const SHEET_ANGGARAN   = 'Input Anggaran';
const SHEET_AKUN       = 'Akun';
const SHEET_KAT_INC    = 'Kategori Pendapatan';
const SHEET_KAT_EXP    = 'Kategori Pengeluaran & Anggaran';
const SHEET_SETTINGS   = 'Settings';
const SHEET_PENGATURAN = 'Pengaturan';

// ── CORS HELPER ─────────────────────────────────────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  // Note: Google Apps Script deployed as "Anyone" already handles OPTIONS
  // For full browser CORS, add the header below if needed:
  // .setHeaders({ 'Access-Control-Allow-Origin': '*' });
  // However GAS Web Apps set this automatically when deployed as "Anyone"
}

function ok(action, data, method) {
  return corsResponse({
    success:   true,
    method:    method || 'GET',
    action:    action,
    data:      data,
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  });
}

function err(msg) {
  return corsResponse({
    success:   false,
    error:     msg,
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  });
}

// ── doGet ────────────────────────────────────────────────────────
function doGet(e) {
  const params = e.parameter || {};

  // If api=1, handle as API request
  if (params.api === '1') {
    try {
      return handleGet(params);
    } catch(ex) {
      return err(ex.message);
    }
  }

  // Otherwise: serve the built-in Apps Script dashboard (legacy)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Fintra Budget Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleGet(params) {
  const action = params.action || '';
  const year   = params.year  ? parseInt(params.year)  : null;
  const month  = params.month ? parseInt(params.month) : null;

  switch(action) {
    case 'health':
      return ok('health', { status: 'ok', spreadsheet: SS.getName() });

    case 'getInitialData':
      return ok('getInitialData', getInitialData());

    case 'getDashboardData':
      return ok('getDashboardData', getDashboardData({ year, month }));

    case 'getTransactions':
      return ok('getTransactions', getTransactions(params));

    case 'getBudgets':
      return ok('getBudgets', getBudgets(params));

    case 'getAccounts':
      return ok('getAccounts', getAccounts());

    case 'getCategories':
      return ok('getCategories', getCategories());

    default:
      return err('Action tidak dikenali: ' + action);
  }
}

// ── doPost ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    let payload;
    try {
      // Website sends form-urlencoded with payload as JSON string in "data" field
      // Postman / server can also send raw JSON body — handle both
      if (e.parameter && e.parameter.data) {
        payload = JSON.parse(e.parameter.data);
      } else if (e.postData && e.postData.contents) {
        payload = JSON.parse(e.postData.contents);
      } else {
        return err('Tidak ada data yang dikirim.');
      }
    } catch(ex) {
      return err('Body tidak valid: ' + ex.message);
    }

    const action = payload.action || '';

    switch(action) {
      case 'addTransaction':
        return ok('addTransaction', addTransaction(payload), 'POST');
      case 'addBudget':
        return ok('addBudget', addBudget(payload), 'POST');
      default:
        return err('Action tidak dikenali: ' + action);
    }
  } catch(ex) {
    return err(ex.message);
  }
}

// ═══════════════════════════════════════════════════════════════
//  READ FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function getInitialData() {
  const cats = getCategories();
  const accounts = getAccounts();
  const optSheet = SS.getSheetByName(SHEET_PENGATURAN);
  let options = { years: [], months: [] };
  if (optSheet) {
    // Years: current year ± 1
    const now = new Date();
    options.years  = [now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1];
    options.months = Array.from({length:12}, (_,i) => i+1);
  }
  return { options, accounts, categories: cats };
}

function getAccounts() {
  const sheet = SS.getSheetByName(SHEET_AKUN);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const accounts = [];
  // Header row is row 1 (index 0), data starts row 2
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1]; // Column B = Akun name
    if (name) accounts.push(String(name).trim());
  }
  return accounts;
}

function getCategories() {
  const income  = [];
  const expense = [];

  const incSheet = SS.getSheetByName(SHEET_KAT_INC);
  if (incSheet) {
    const data = incSheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) { // skip header + total
      const k = data[i][1];
      if (k && String(k).trim() && String(k).trim() !== 'Total') income.push(String(k).trim());
    }
  }

  const expSheet = SS.getSheetByName(SHEET_KAT_EXP);
  if (expSheet) {
    const data = expSheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) { // skip header + total
      const k = data[i][1];
      if (k && String(k).trim() && String(k).trim() !== 'Total') expense.push(String(k).trim());
    }
  }

  return { income, expense };
}

function getTransactions(params) {
  const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0]; // But we know the structure
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Columns: [0]=blank, [1]=blank, [2]=Tanggal, [3]=Keterangan, [4]=Jumlah, [5]=blank, [6]=Tipe, [7]=Kategori, [8]=Akun
    const dateRaw   = row[2];
    const ket       = row[3];
    const jumlah    = row[4];
    const tipe      = row[6];
    const kategori  = row[7];
    const akun      = row[8];

    if (!dateRaw) continue;

    const d = dateRaw instanceof Date ? dateRaw : new Date(dateRaw);
    if (isNaN(d)) continue;

    const rowYear  = d.getFullYear();
    const rowMonth = d.getMonth() + 1;

    // Filters
    if (params.year  && parseInt(params.year)  !== rowYear)  continue;
    if (params.month && parseInt(params.month) !== rowMonth) continue;
    if (params.tipe  && params.tipe  !== String(tipe).trim())     continue;
    if (params.kategori && params.kategori !== String(kategori).trim()) continue;
    if (params.akun  && params.akun  !== String(akun).trim())     continue;

    rows.push({
      tanggal:   Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      keterangan: String(ket||'').trim(),
      jumlah:    Number(jumlah||0),
      tipe:      String(tipe||'').trim(),
      kategori:  String(kategori||'').trim(),
      akun:      String(akun||'').trim(),
    });
  }
  return rows;
}

function getBudgets(params) {
  const sheet = SS.getSheetByName(SHEET_ANGGARAN);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Columns: [0]=blank, [1]=blank, [2]=TanggalInput, [3]=TahunAnggaran, [4]=BulanAnggaran, [5]=Jumlah, [6]=Kategori
    const tglRaw  = row[2];
    const tahun   = row[3];
    const bulan   = row[4];
    const jumlah  = row[5];
    const kat     = row[6];

    if (!tahun || !bulan) continue;
    if (params.year  && parseInt(params.year)  !== parseInt(tahun))  continue;
    if (params.month && parseInt(params.month) !== parseInt(bulan))  continue;
    if (params.kategori && params.kategori !== String(kat).trim())   continue;

    const dateStr = tglRaw instanceof Date
      ? Utilities.formatDate(tglRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(tglRaw||'');

    rows.push({
      tanggalInput:  dateStr,
      tahunAnggaran: parseInt(tahun),
      bulanAnggaran: parseInt(bulan),
      jumlah:        Number(jumlah||0),
      kategori:      String(kat||'').trim(),
    });
  }
  return rows;
}

function getDashboardData({ year, month }) {
  const now = new Date();
  const y = year  || now.getFullYear();
  const m = month || (now.getMonth() + 1);

  const txns    = getTransactions({ year: y, month: m });
  const budgets = getBudgets({ year: y, month: m });

  // Metrics
  const income  = txns.filter(t => t.tipe === 'Pendapatan').reduce((s,t) => s + Math.abs(t.jumlah), 0);
  const spent   = txns.filter(t => t.tipe === 'Pengeluaran').reduce((s,t) => s + Math.abs(t.jumlah), 0);
  const net     = income - spent;
  const savePct = income > 0 ? net / income : 0;

  // Total budget for the month (sum of all budget entries for this month)
  const totalBudget = budgets.reduce((s,b) => s + Math.abs(b.jumlah), 0);

  // Projected spend
  const daysInMonth = new Date(y, m, 0).getDate();
  const today       = now.getFullYear() === y && (now.getMonth()+1) === m ? now.getDate() : daysInMonth;
  const avgDaily    = today > 0 ? spent / today : 0;
  const projected   = avgDaily * daysInMonth;
  const projUtil    = totalBudget > 0 ? projected / totalBudget : 0;

  // Expense breakdown by category
  const breakdown = {};
  txns.filter(t => t.tipe === 'Pengeluaran').forEach(t => {
    const k = t.kategori || 'Lain-Lain';
    breakdown[k] = (breakdown[k]||0) + Math.abs(t.jumlah);
  });

  // Budget per category
  const budgetByKat = {};
  budgets.forEach(b => { budgetByKat[b.kategori] = (budgetByKat[b.kategori]||0) + b.jumlah; });

  const breakdownArr = Object.entries(breakdown)
    .map(([kategori, spentAmt]) => ({
      kategori,
      spent:    spentAmt,
      budget:   budgetByKat[kategori] || 0,
      pct:      budgetByKat[kategori] > 0 ? spentAmt / budgetByKat[kategori] : null
    }))
    .sort((a,b) => b.spent - a.spent);

  return {
    period: { year: y, month: m },
    metrics: {
      totalPendapatan:      income,
      totalPengeluaran:     spent,
      saldoBersih:          net,
      persentaseDitabung:   savePct,
      totalBudget:          totalBudget,
      spent:                spent,
      income:               income,
      averageDailySpend:    Math.round(avgDaily),
      projectedMonthEndSpend: Math.round(projected),
      projectedUtilization: projUtil,
      daysInMonth:          daysInMonth,
      daysElapsed:          today,
    },
    transactions: txns,
    budgets:      budgets,
    breakdown:    breakdownArr,
  };
}

// ═══════════════════════════════════════════════════════════════
//  WRITE FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function addTransaction(payload) {
  // Validate
  if (!payload.tanggal)    throw new Error('Field "tanggal" wajib diisi.');
  if (!payload.keterangan) throw new Error('Field "keterangan" wajib diisi.');
  if (!payload.jumlah || Number(payload.jumlah) <= 0) throw new Error('Field "jumlah" harus > 0.');
  if (!payload.tipe)       throw new Error('Field "tipe" wajib diisi.');
  if (!payload.kategori)   throw new Error('Field "kategori" wajib diisi.');
  if (!payload.akun)       throw new Error('Field "akun" wajib diisi.');

  const sheet = SS.getSheetByName(SHEET_TRANSAKSI);
  if (!sheet) throw new Error('Sheet Transaksi tidak ditemukan.');

  const tgl    = new Date(payload.tanggal);
  const jumlah = payload.tipe === 'Pengeluaran'
    ? -Math.abs(Number(payload.jumlah))
    :  Math.abs(Number(payload.jumlah));

  // Append: blank, blank, date, keterangan, jumlah, blank, tipe, kategori, akun
  sheet.appendRow([
    '', '',
    tgl,
    payload.keterangan.trim(),
    jumlah,
    '',
    payload.tipe,
    payload.kategori,
    payload.akun,
  ]);

  return { message: 'Transaksi berhasil ditambahkan.', rows: sheet.getLastRow() - 1 };
}

function addBudget(payload) {
  if (!payload.tahunAnggaran) throw new Error('Field "tahunAnggaran" wajib diisi.');
  if (!payload.bulanAnggaran) throw new Error('Field "bulanAnggaran" wajib diisi.');
  if (!payload.jumlah || Number(payload.jumlah) <= 0) throw new Error('Field "jumlah" harus > 0.');
  if (!payload.kategori)      throw new Error('Field "kategori" wajib diisi.');

  const sheet = SS.getSheetByName(SHEET_ANGGARAN);
  if (!sheet) throw new Error('Sheet "Input Anggaran" tidak ditemukan.');

  const tglInput = payload.tanggalInput ? new Date(payload.tanggalInput) : new Date();
  const timestamp = new Date();

  // Append: blank, blank, tanggalInput, tahun, bulan, jumlah, kategori, blank, timestamp
  sheet.appendRow([
    '', '',
    tglInput,
    parseInt(payload.tahunAnggaran),
    parseInt(payload.bulanAnggaran),
    Math.abs(Number(payload.jumlah)),
    payload.kategori,
    '',
    timestamp,
  ]);

  return { message: 'Anggaran berhasil ditambahkan.', rows: sheet.getLastRow() - 1 };
}
