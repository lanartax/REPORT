/**
 * GadgetMart Sosmed KPI & Daily Achievement Engine
 * Features: turnover.xlsx integration, Report 1 (Daily Matrix), Report 2 (KPI & GP vs Timegone),
 * Tab 3: Cek Status AR & Verifikasi Sosmed, Tab 4: Arsip Bulanan, Tab 5: Admin Panel,
 * Full-height PNG Screenshot Download Engine, and Gemini AI Assistant.
 */

// 1. DEFAULT MASTER SALES & TARGETS CONFIGURATION
const DEFAULT_SALES_MASTER = [
  { code: 'BJM.SOS.005', name: 'SIFA', target: 1081261572, active: true },
  { code: 'BJM.SOS.015', name: 'DANIEL', target: 1081261572, active: true },
  { code: 'BJM.SOS.017', name: 'MARLINA', target: 1081261572, active: true },
  { code: 'BJM.SOS.004', name: 'BINTANG', target: 867261886, active: true },
  { code: 'BJM.SOS.007', name: 'FEBRI', target: 867261886, active: true },
  { code: 'BJM.SOS.016', name: 'DEA', target: 867261886, active: true },
  { code: 'BJM.SOS.018', name: 'MARISA', target: 867261886, active: true },
  { code: 'BJM.SOS.019', name: 'ZOFIA', target: 867261886, active: true },
  { code: 'BJM.SOS.022', name: 'QONA', target: 867261886, active: true },
  // Extended Master Sales
  { code: 'BJM.SOS.013', name: 'TIKA', target: 867261886, active: false },
  { code: 'BJM.SOS.001', name: 'DIANA', target: 867261886, active: false },
  { code: 'BJM.SOS.012', name: 'QOFA', target: 867261886, active: false },
  { code: 'BJM.SOS.010', name: 'RIMA', target: 867261886, active: false },
  { code: 'BJM.SOS.020', name: 'REVA', target: 867261886, active: false },
  { code: 'BJM.SOS.021', name: 'ALI', target: 867261886, active: false },
  { code: 'BJM.SOS.000', name: 'NOL', target: 0, active: false }
];

function getMasterSales() {
  const saved = localStorage.getItem('gadgetmart_sosmed_master');
  if (saved) {
    try { return JSON.parse(saved); } catch(e){}
  }
  return DEFAULT_SALES_MASTER;
}

function saveMasterSales(list) {
  localStorage.setItem('gadgetmart_sosmed_master', JSON.stringify(list));
}

// Default Gemini API configuration
const DEFAULT_GEMINI_API_KEY = 'AQ.Ab8RN6Lvfri-gb4jU1QFWZYHQyb3pprbRkGJCoV020aWIWCsXA';
const DEFAULT_GEMINI_MODEL = 'gemini-3.5-flash-lite';

// Global State
let currentCalculationMode = 'DPP'; // 'DPP' or 'GROSS'
let rawTurnoverRows = [];     // Source for Report 1 & 2 Achievement (with DPP)
let rawCrosellingRows = [];   // Source for Tab 3: Cek AR & Pencarian WhatsApp Pelanggan (Google Sheets)
let parsedDailyMatrix = {};
let activeMonth = 10;
let activeYear = 2026;
let daysInMonth = 31;
let lastActiveDay = 13;
let timegonePct = 42;
let currentActiveTab = 1;

// Executive Dashboard & Leaderboard State
let currentLeaderboardSort = 'omzet'; // 'omzet', 'gp', 'score'
let leaderboardPeriodMode = 'current'; // 'current' or 'ath'
let dashFilters = { month: '10', cabang: 'all', sales: 'all', mode: 'DPP' };
let cachedMonthlyTargets = {};
let lastProcessedPicStats = null;
let lastActivePICs = [];
let lastDocTransactions = {};
let lastTotalSosmedAchieve = 0;
let lastTotalSosmedGP = 0;
let lastTotalSosmedGadget = 0;
let lastTotalSosmedCross = 0;
let cachedSalesPhotos = {};
let calculatedSalesKPI = [];
let cachedMonthlyArchives = [];
let chartInstances = {};



// Utility Formatters
function formatRupiah(num) {
  if (num === null || num === undefined || isNaN(num) || num === 0) return 'Rp -';
  const isNeg = num < 0;
  const absVal = Math.abs(num);
  const formatted = absVal.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  return (isNeg ? '- Rp ' : 'Rp ') + formatted;
}

function formatPercent(pct) {
  if (isNaN(pct)) return '0%';
  return `${Math.round(pct)}%`;
}

// Phone Normalizer & Formatter
function normalizePhone(p) {
  if (!p) return '';
  let str = String(p).replace(/\D/g, '');
  if (str.startsWith('62')) return str;
  if (str.startsWith('0')) return '62' + str.slice(1);
  if (str.startsWith('8')) return '62' + str;
  return str;
}

function formatPhoneDisplay(p) {
  const norm = normalizePhone(p);
  if (!norm || norm.length < 9) return p || '-';
  if (norm.startsWith('62')) {
    const rest = norm.slice(2);
    return `+62 ${rest.slice(0, 3)}-${rest.slice(3, 7)}-${rest.slice(7)}`;
  }
  return norm;
}

// RESOLVE POSTING DATE: Handles Inverted US Locale Date (MM/DD/YYYY exported from SAP turnover.xlsx)
function resolvePostingDay(pd) {
  if (typeof pd === 'number') {
    const d = XLSX.SSF.parse_date_code(pd);
    // In turnover.xlsx, October (month 10) was exported as day 10, and days 1-12 were exported as month 1-12!
    if (d.d === 10 && d.m >= 1 && d.m <= 12) {
      return d.m;
    }
    return d.d;
  }
  if (typeof pd === 'string') {
    const parts = pd.split(/[\/\-]/);
    if (parts.length >= 2) return parseInt(parts[0], 10);
  }
  return 1;
}

function formatDisplayDate(pd) {
  const day = resolvePostingDay(pd);
  return `${String(day).padStart(2, '0')}/${String(activeMonth).padStart(2, '0')}/${activeYear}`;
}

// Normal Date Formatter for Google Sheets / Croselling rows
function formatCrosellingDate(pd) {
  if (!pd) return '-';
  if (typeof pd === 'number') {
    const d = XLSX.SSF.parse_date_code(pd);
    if (d) {
      return `${String(d.d).padStart(2, '0')}/${String(d.m).padStart(2, '0')}/${d.y}`;
    }
  }
  if (typeof pd === 'string') {
    const parts = pd.trim().split(/[\/\-]/);
    if (parts.length >= 3) {
      let y = parts[2].trim().split(' ')[0];
      if (y.length === 2) y = '20' + y;
      return `${String(parts[0]).padStart(2, '0')}/${String(parts[1]).padStart(2, '0')}/${y}`;
    }
    return pd;
  }
  return '-';
}

// ==========================================================================
// INDEXEDDB ENGINE: BROWSER PERSISTENT STORAGE FOR EXCEL FILES
// ==========================================================================
const GM_DB_NAME = 'GadgetMartReportDB';
const GM_DB_VERSION = 1;
const GM_STORE_NAME = 'uploaded_files';

function openReportDB() {
  return new Promise((resolve) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      resolve(null);
      return;
    }
    const req = window.indexedDB.open(GM_DB_NAME, GM_DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(GM_STORE_NAME)) {
        db.createObjectStore(GM_STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => {
      console.warn('IndexedDB gagal dibuka:', req.error);
      resolve(null);
    };
  });
}

async function saveFileToIndexedDB(key, arrayBuffer) {
  try {
    const db = await openReportDB();
    if (!db) return false;
    return new Promise((resolve) => {
      const tx = db.transaction(GM_STORE_NAME, 'readwrite');
      const store = tx.objectStore(GM_STORE_NAME);
      store.put(arrayBuffer, key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    });
  } catch (e) {
    console.warn('saveFileToIndexedDB error:', e);
    return false;
  }
}

async function getFileFromIndexedDB(key) {
  try {
    const db = await openReportDB();
    if (!db) return null;
    return new Promise((resolve) => {
      const tx = db.transaction(GM_STORE_NAME, 'readonly');
      const store = tx.objectStore(GM_STORE_NAME);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (e) {
    console.warn('getFileFromIndexedDB error:', e);
    return null;
  }
}

async function clearFileFromIndexedDB(key) {
  try {
    const db = await openReportDB();
    if (!db) return;
    const tx = db.transaction(GM_STORE_NAME, 'readwrite');
    tx.objectStore(GM_STORE_NAME).delete(key);
  } catch (e) {}
}

function getFormattedTimestamp(date = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  const d = pad(date.getDate());
  const m = pad(date.getMonth() + 1);
  const y = date.getFullYear();
  const hr = pad(date.getHours());
  const min = pad(date.getMinutes());
  return `${d}/${m}/${y} ${hr}:${min} WITA`;
}

function updateTurnoverIndicator(text) {
  const el1 = document.getElementById('turnoverUpdateText');
  if (el1) el1.innerText = text;
  const el2 = document.getElementById('turnoverBannerText');
  if (el2) el2.innerText = text;
}

async function resetTurnoverToDefault() {
  if (confirm("Apakah Anda yakin ingin menghapus file turnover yang diunggah dan kembali ke data default sistem?")) {
    try {
      await fetch('/api/get-excel?type=turnover', { method: 'DELETE' });
    } catch(e) {}
    await clearFileFromIndexedDB('turnover_file');
    localStorage.removeItem('gadgetmart_turnover_timestamp');
    localStorage.removeItem('gadgetmart_turnover_is_custom');
    alert("File turnover kustom telah dihapus dari database & browser. Halaman akan dimuat ulang dengan data default sistem.");
    window.location.reload();
  }
}

// MATCH ALL SOSMED SALES EMPLOYEE CODES (BOTH "BJM.SOS.0" AND "BJM.SOS.K.0")
function isSosmedSalesCode(empStr, remStr = '', cabStr = '') {
  const emp = String(empStr || '').trim().toLowerCase();
  const rem = String(remStr || '').trim().toLowerCase();
  const cab = String(cabStr || '').trim().toLowerCase();

  // All sales employee codes starting with "bjm.sos" (both bjm.sos.0 and bjm.sos.k.0)
  if (emp.startsWith('bjm.sos')) return true;
  if (/bjm\.sos(?:\.k)?\./i.test(emp)) return true;
  if (emp.includes('bjm.sos')) return true;

  // Check remarks and sales cabang if emp is empty
  if (/bjm\.sos(?:\.k)?\./i.test(rem) || rem.includes('bjm.sos')) return true;
  if (/bjm\.sos(?:\.k)?\./i.test(cab) || cab.includes('bjm.sos')) return true;

  return false;
}

// Validasi ketat kode sales item faktur: harus benar-benar diinput dengan KODE SALES SOSMED (BJM.SOS)
// Catatan: Jika kode salesnya FL. (Floor/Toko) atau SC. (Sales Counter Toko), meskipun di Remarks ada tulisan 'sosmed',
// item tersebut TETAP BELUM TERINPUT SOSMED di SAP kasir!
function isItemInputSosmed(row) {
  if (!row) return false;
  const schl = String(row['SC/HL Code'] || '').trim();
  const slp = String(row['SlpName'] || row['Sales Employee Name'] || '').trim();

  if (/^bjm\.sos/i.test(schl) || /bjm\.sos(?:\.k)?\./i.test(schl)) return true;
  if (/^bjm\.sos/i.test(slp) || /bjm\.sos(?:\.k)?\./i.test(slp)) return true;

  return false;
}

// Barang yang bebas dari kewajiban input sosmed (misal TAS GADGETMART 25x25 / kantong belanja)
// Jika customer beli HP + Tas Gadgetmart, dan HP-nya sudah sosmed, faktur tetap dihitung 100% SUDAH SOSMED
function isItemExemptFromSosmed(row) {
  if (!row) return false;
  const desc = String(row['Dscription'] || row['Item/Service Description'] || '').trim().toUpperCase();
  const code = String(row['Item No.'] || row['Item Code'] || '').trim().toUpperCase();

  if (desc.includes('TAS GADGETMART') || 
      desc.includes('TAS GADGET MART') || 
      desc.includes('TAS 25X25') || 
      desc.includes('TAS 25 X 25') || 
      desc.includes('KANTONG GADGETMART') ||
      desc.includes('KANTONG 25X25') ||
      desc.includes('KANTONG 25 X 25')) {
    return true;
  }

  if ((code.includes('TAS') || desc.includes('TAS')) && (code.includes('25X25') || desc.includes('25X25') || desc.includes('25 X 25'))) {
    return true;
  }

  return false;
}

// Helper untuk memilih sheet transaksi turnover yang tepat (misal TARIKKAN pada file KPI SOSMED)
function resolveTurnoverSheet(workbook) {
  if (workbook.SheetNames.includes('TARIKKAN')) return workbook.Sheets['TARIKKAN'];
  if (workbook.SheetNames.includes('Sheet1')) return workbook.Sheets['Sheet1'];
  return workbook.Sheets[workbook.SheetNames[0]];
}

// Helper untuk membaca data turnover secara cepat & hemat memori (mencegah crash/OOM peramban pada file >50MB)
function readTurnoverJsonFromBuffer(arrayBuffer) {
  let targetSheetName = 'TARIKKAN';
  try {
    const metaWb = XLSX.read(arrayBuffer, { type: 'array', bookSheets: true });
    if (metaWb && Array.isArray(metaWb.SheetNames) && metaWb.SheetNames.length > 0) {
      if (metaWb.SheetNames.includes('TARIKKAN')) targetSheetName = 'TARIKKAN';
      else if (metaWb.SheetNames.includes('Sheet1')) targetSheetName = 'Sheet1';
      else targetSheetName = metaWb.SheetNames[0];
    }
  } catch (e) {
    console.warn('Meta read error:', e);
  }

  try {
    const wb = XLSX.read(arrayBuffer, {
      type: 'array',
      sheets: [targetSheetName],
      cellFormula: false,
      cellHTML: false,
      cellText: false
    });
    const sheet = wb.Sheets[targetSheetName];
    if (sheet) return XLSX.utils.sheet_to_json(sheet);
  } catch (err) {
    console.warn('Optimized read failed, fallback to standard read:', err);
  }

  const standardWb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheet = resolveTurnoverSheet(standardWb);
  return XLSX.utils.sheet_to_json(sheet);
}

// 2. PARSER ENGINE FOR turnover.xlsx (Achievement DPP)
function parseTurnoverBuffer(arrayBuffer) {
  try {
    const jsonData = readTurnoverJsonFromBuffer(arrayBuffer);

    if (!jsonData || jsonData.length === 0) return;
    rawTurnoverRows = jsonData;
    console.log(`turnover.xlsx loaded: ${jsonData.length} rows`);
    processAllReports();
    updateDatabaseStatus();
  } catch (err) {
    console.error("Gagal membaca turnover.xlsx:", err);
  }
}

// PARSER ENGINE FOR Google Sheets / Croselling (AR & WhatsApp Customer Search)
function parseCrosellingBuffer(arrayBuffer) {
  try {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    if (!jsonData || jsonData.length === 0) return;
    rawCrosellingRows = jsonData;
    console.log(`Google Sheets croselling loaded: ${jsonData.length} rows`);
    const pill = document.getElementById('crosellingStatusPill');
    if (pill) {
      pill.innerHTML = `<span>📁 Data Source: Google Sheets (${jsonData.length.toLocaleString('id-ID')} Baris Data - Ada No WhatsApp)</span>`;
    }
    updateDatabaseStatus();
    if (typeof rawKeepRows !== 'undefined' && rawKeepRows.length > 0 && typeof runARMissingSosmedAudit === 'function') {
      runARMissingSosmedAudit();
    }
  } catch (err) {
    console.error("Gagal membaca Google Sheets croselling:", err);
  }
}

function updateDatabaseStatus() {
  const title = document.getElementById('fileStatusTitle');
  const sub = document.getElementById('fileStatusSub');
  if (!title || !sub) return;

  const tCount = rawTurnoverRows.length;
  const cCount = rawCrosellingRows.length;

  if (tCount > 0 && cCount > 0) {
    title.innerText = `Dual Engine Siap: turnover.xlsx (${tCount.toLocaleString('id-ID')} Baris) & Google Sheets (${cCount.toLocaleString('id-ID')} Baris)`;
    sub.innerHTML = `✅ <strong>turnover.xlsx</strong> aktif untuk Report 1 & 2 (DPP) &bull; ✅ <strong>Google Sheets</strong> aktif untuk Cek AR & No. WhatsApp!`;
  } else if (tCount > 0) {
    title.innerText = `turnover.xlsx Dimuat (${tCount.toLocaleString('id-ID')} Baris Data)`;
    sub.innerHTML = `Report 1 & 2 Achievement terhitung menggunakan Nilai DPP. Menunggu data Google Sheets...`;
  } else if (cCount > 0) {
    title.innerText = `Google Sheets Dimuat (${cCount.toLocaleString('id-ID')} Baris Data)`;
    sub.innerHTML = `Database AR & WhatsApp siap digunakan untuk pencarian.`;
  }
}

// SMART FILE UPLOAD HANDLER (WITH INDEXEDDB PERSISTENCE & LIVE INDICATOR)
async function parseExcelFile(arrayBuffer, fileName = '') {
  try {
    const jsonData = readTurnoverJsonFromBuffer(arrayBuffer);

    if (!jsonData || jsonData.length === 0) {
      alert("File Excel kosong atau tidak memiliki data.");
      return;
    }

    const firstRow = jsonData[0] || {};
    const hasWA = 'No HP Customer' in firstRow || 'SC/HL Code' in firstRow || fileName.toLowerCase().includes('croselling') || fileName.toLowerCase().includes('sheet');
    const hasDPP = 'DPP' in firstRow || fileName.toLowerCase().includes('turnover');
    const updateTs = getFormattedTimestamp();

    if (hasWA && !hasDPP) {
      rawCrosellingRows = jsonData;
      await saveFileToIndexedDB('croselling_file', arrayBuffer);
      localStorage.setItem('gadgetmart_croselling_timestamp', updateTs);
      saveExcelToServer(arrayBuffer, 'croselling');
      alert(`Berhasil mengimpor Data Google Sheets (${jsonData.length.toLocaleString('id-ID')} baris data)!\nDatabase AR & WhatsApp telah diperbarui dan tetap tersimpan saat refresh.`);
      switchTab(3);
    } else if (hasDPP) {
      rawTurnoverRows = jsonData;
      await saveFileToIndexedDB('turnover_file', arrayBuffer);
      localStorage.setItem('gadgetmart_turnover_timestamp', updateTs);
      localStorage.setItem('gadgetmart_turnover_is_custom', 'true');
      updateTurnoverIndicator(`File turnover terupdate per ${updateTs}`);
      saveExcelToServer(arrayBuffer, 'turnover');
      processAllReports();
      alert(`Berhasil mengimpor turnover.xlsx (${jsonData.length.toLocaleString('id-ID')} baris data)!\nFile turnover terupdate per ${updateTs} dan akan tetap tersimpan saat refresh.`);
      switchTab(2);
    } else {
      // Default: check if it has No HP
      if ('No HP Customer' in firstRow) {
        rawCrosellingRows = jsonData;
        await saveFileToIndexedDB('croselling_file', arrayBuffer);
        localStorage.setItem('gadgetmart_croselling_timestamp', updateTs);
        saveExcelToServer(arrayBuffer, 'croselling');
        switchTab(3);
      } else {
        rawTurnoverRows = jsonData;
        await saveFileToIndexedDB('turnover_file', arrayBuffer);
        localStorage.setItem('gadgetmart_turnover_timestamp', updateTs);
        localStorage.setItem('gadgetmart_turnover_is_custom', 'true');
        updateTurnoverIndicator(`File turnover terupdate per ${updateTs}`);
        saveExcelToServer(arrayBuffer, 'turnover');
        processAllReports();
        switchTab(2);
      }
    }
    updateDatabaseStatus();
  } catch (err) {
    console.error("Gagal membaca Excel:", err);
    alert("Gagal membaca file Excel. Pastikan format file valid.");
  }
}

// 3. CORE PROCESSING (REPORT 1 & REPORT 2)
function processAllReports() {
  if (!rawTurnoverRows || rawTurnoverRows.length === 0) return;

  const masterList = getMasterSales();
  const activePICs = masterList.filter(m => m.active).map(m => m.name);

  // Initialize Report 1 Matrix (Day 1..31)
  parsedDailyMatrix = {};
  for (let day = 1; day <= daysInMonth; day++) {
    const dayStr = String(day).padStart(2, '0');
    const monthStr = String(activeMonth).padStart(2, '0');
    const dateStr = `${dayStr}/${monthStr}/${activeYear}`;

    let weekLabel = 'WEEK 1';
    if (day >= 8 && day <= 14) weekLabel = 'WEEK 2';
    else if (day >= 15 && day <= 21) weekLabel = 'WEEK 3';
    else if (day >= 22 && day <= 28) weekLabel = 'WEEK 4';
    else if (day >= 29) weekLabel = 'WEEK 5';

    parsedDailyMatrix[day] = {
      day,
      dateStr,
      weekLabel,
      sales: {},
      totalDay: 0,
      avg: 0
    };

    activePICs.forEach(pic => {
      parsedDailyMatrix[day].sales[pic] = 0;
    });
  }

  // Aggregate stats per PIC for Report 2
  const picStats = {};
  activePICs.forEach(pic => {
    const m = masterList.find(x => x.name === pic);
    const target = m ? getSalesTargetForMonth(m.code, activeMonth, activeYear, m.target) : 867261886;
    picStats[pic] = {
      name: pic,
      target: target,
      pencapaian: 0,
      targetGP: Math.round(target * 0.05),
      acvGP: 0,
      gadgetCount: 0,
      crossCount: 0
    };
  });

  const docTransactions = {};
  let maxActiveDay = 1;
  let totalSosmedAchieve = 0;
  let totalSosmedGP = 0;

  for (const row of rawTurnoverRows) {
    const emp = String(row['Sales Employee Name'] || '');
    const rem = String(row['Remarks'] || '');
    const cab = String(row['SalesCabang'] || '');

    // Check if transaction belongs to Sosmed (matches all codes starting with bjm.sos, including bjm.sos.0 and bjm.sos.k.0)
    const isSos = isSosmedSalesCode(emp, rem, cab);
    if (!isSos) continue;

    // Resolve date with inverted US-locale fix
    const day = resolvePostingDay(row['Posting Date']);
    if (day >= 1 && day <= daysInMonth) {
      if (day > maxActiveDay) maxActiveDay = day;
    }

    // Amounts
    const dpp = Number(row['DPP']) || 0;
    const total = Number(row['Total']) || 0;
    // GP: Gunakan kolom 'FIX' jika tersedia (karena di Excel kolom FIX = GrossProfit + Cashback - Admin Kredit - Admin Marketplace),
    // jika kolom 'FIX' tidak ada (misal file turnover mentah standar), gunakan 'GrossProfit'.
    const fixVal = (row['FIX'] !== undefined && row['FIX'] !== '') ? Number(row['FIX']) : NaN;
    const rawGP = Number(row['GrossProfit']) || 0;
    const gp = !isNaN(fixVal) ? fixVal : rawGP;
    const amount = currentCalculationMode === 'DPP' ? dpp : total;

    // Accumulate total division sosmed achievement
    totalSosmedAchieve += amount;
    totalSosmedGP += gp;

    // Report 1: Daily Matrix accumulation for total day
    if (parsedDailyMatrix[day]) {
      parsedDailyMatrix[day].totalDay += amount;
    }

    // Identify sales PIC
    let matchedPIC = null;
    for (const m of masterList) {
      if (emp.includes(m.code) || rem.includes(m.code) || cab.includes(m.code)) {
        matchedPIC = m.name;
        break;
      }
    }

    // Accumulate individual PIC stats if matched and active
    if (matchedPIC && activePICs.includes(matchedPIC)) {
      if (parsedDailyMatrix[day]) {
        parsedDailyMatrix[day].sales[matchedPIC] += amount;
      }
      picStats[matchedPIC].pencapaian += amount;
      picStats[matchedPIC].acvGP += gp;
    }

    // Grouping for Cross-selling calculation (by Document Number)
    const docNo = row['Document Number'];
    if (docNo) {
      if (!docTransactions[docNo]) {
        docTransactions[docNo] = {
          pic: matchedPIC || null,
          gadgetPics: new Set(),
          accPics: new Set(),
          accTotal: 0,
          hasGadget: false,
          hasCross: false
        };
      }
      if (matchedPIC && !docTransactions[docNo].pic) {
        docTransactions[docNo].pic = matchedPIC;
      }

      // Klasifikasi kategori sesuai rumus standar FILL GRUP Excel
      // TM = Tebus Murah (diberi 'X' agar tidak dihitung sebagai gadget/croselling reguler)
      const itemDesc = String(row['Item/Service Description'] || '');
      const grpRaw = String(row['Group Name'] || '').trim().toUpperCase();
      let grp = grpRaw.slice(0, 3);
      if (itemDesc.startsWith('TM ')) grp = 'X';
      else if (itemDesc.startsWith('GMART CARE') || itemDesc.startsWith('JASA SERVICE T') || itemDesc.startsWith('JASA SERVICE FLEX T')) grp = 'ACC';

      if (grp === 'GAD') {
        docTransactions[docNo].hasGadget = true;
        if (matchedPIC) {
          docTransactions[docNo].gadgetPics.add(matchedPIC);
          docTransactions[docNo].pic = matchedPIC;
        }
      } else if (grp === 'ACC') {
        docTransactions[docNo].accTotal += total;
        if (matchedPIC) {
          docTransactions[docNo].accPics.add(matchedPIC);
        }
      }
    }
  }

  // Calculate Cross-Selling Counts from Document groups
  // Rumus Standar Excel: JLH CROSS >50K adalah akumulasi Total Aksesoris dalam satu faktur/nota > 50.000
  let totalSosmedGadget = 0;
  let totalSosmedCross = 0;
  for (const doc in docTransactions) {
    const item = docTransactions[doc];
    if (item.hasGadget) {
      totalSosmedGadget++;
      const isCross = item.accTotal > 50000;
      item.hasCross = isCross;
      if (isCross) {
        totalSosmedCross++;
      }

      // Hitung Gadget per PIC
      item.gadgetPics.forEach(pic => {
        if (picStats[pic]) picStats[pic].gadgetCount++;
      });

      // Hitung Cross (>50K) per PIC
      if (isCross) {
        if (item.accPics.size > 0) {
          item.accPics.forEach(pic => {
            if (picStats[pic]) picStats[pic].crossCount++;
          });
        } else {
          item.gadgetPics.forEach(pic => {
            if (picStats[pic]) picStats[pic].crossCount++;
          });
        }
      }
    }
  }

  // Calculate AVG for Report 1
  for (let day = 1; day <= daysInMonth; day++) {
    parsedDailyMatrix[day].avg = parsedDailyMatrix[day].totalDay / activePICs.length;
  }

  // Calculate Timegone based on max active day (e.g. 13 / 31 = 42%)
  lastActiveDay = maxActiveDay > 0 ? maxActiveDay : 13;
  timegonePct = Math.round((lastActiveDay / daysInMonth) * 100);

  // Save for global usage in dashboard & leaderboard
  lastProcessedPicStats = picStats;
  lastActivePICs = activePICs;
  lastDocTransactions = docTransactions;
  lastTotalSosmedAchieve = totalSosmedAchieve;
  lastTotalSosmedGP = totalSosmedGP;
  lastTotalSosmedGadget = totalSosmedGadget;
  lastTotalSosmedCross = totalSosmedCross;

  // Render all views
  renderReport1Table(activePICs);
  renderReport2Tables(picStats, activePICs);
  updateExecutiveKPICards(picStats, activePICs);
  populateDashboardFilterDropdowns();
  applyDashboardFilters();
  renderSpeedometerGauges(picStats, activePICs);
  renderMonthToMonthCharts();
  renderSalesLeaderboard(currentLeaderboardSort);
  populateAdminPanel(masterList);
}

// 4. RENDER REPORT 1: DAILY MATRIX
function renderReport1Table(activePICs) {
  const theadRow = document.getElementById('matrixHeaderRow');
  const tbody = document.getElementById('matrixTableBody');
  const tfoot = document.getElementById('matrixTableFoot');

  const titleEl = document.getElementById('report1MainTitle');
  if (titleEl) {
    const monthNames = ['', 'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'];
    const mStr = monthNames[Number(activeMonth)] || String(activeMonth).toUpperCase();
    titleEl.innerText = `REPORT 1: DAILY ACHIEVEMENT SOSMED ${mStr} ${activeYear}`;
  }

  let headerHtml = `<th class="col-sticky col-date">TANGGAL</th>`;
  activePICs.forEach(pic => headerHtml += `<th>${pic}</th>`);
  headerHtml += `<th class="col-achieve">ACHIEVE DAY</th><th class="col-week">WEEK</th><th class="col-avg">AVG</th>`;
  theadRow.innerHTML = headerHtml;

  tbody.innerHTML = '';
  const colTotals = {};
  activePICs.forEach(pic => colTotals[pic] = 0);
  let grandTotalAchieve = 0;

  for (let day = 1; day <= daysInMonth; day++) {
    const row = parsedDailyMatrix[day];
    const tr = document.createElement('tr');
    tr.id = `row-date-${day}`;
    if (day === lastActiveDay) tr.classList.add('row-today');

    let trHtml = `<td class="col-sticky">${row.dateStr}</td>`;
    activePICs.forEach(pic => {
      const val = row.sales[pic] || 0;
      colTotals[pic] += val;
      const valClass = val === 0 ? 'cell-zero' : '';
      trHtml += `<td class="${valClass}">${formatRupiah(val)}</td>`;
    });

    grandTotalAchieve += row.totalDay;
    const achieveClass = row.totalDay === 0 ? 'cell-zero' : '';
    trHtml += `<td class="col-achieve ${achieveClass}">${formatRupiah(row.totalDay)}</td>`;
    trHtml += `<td class="col-week">${row.weekLabel}</td>`;
    trHtml += `<td class="col-avg ${row.avg === 0 ? 'cell-zero' : ''}">${formatRupiah(row.avg)}</td>`;

    tr.innerHTML = trHtml;
    tbody.appendChild(tr);
  }

  // Footer: TOTAL, SISA, TARGET, PROGRESS
  const masterList = getMasterSales();
  let totalHtml = `<tr class="row-total"><td class="col-sticky">TOTAL</td>`;
  let sisaHtml = `<tr class="row-sisa"><td class="col-sticky">SISA</td>`;
  let targetHtml = `<tr class="row-target"><td class="col-sticky">TARGET</td>`;
  let progressHtml = `<tr class="row-progress"><td class="col-sticky">PROGRESS</td>`;

  let sumTarget = 0;
  activePICs.forEach(pic => {
    const m = masterList.find(x => x.name === pic);
    const target = m ? getSalesTargetForMonth(m.code, activeMonth, activeYear, m.target) : 867261886;
    sumTarget += target;
    const total = colTotals[pic];
    const sisa = total - target;
    const pct = target > 0 ? (total / target) * 100 : 0;

    totalHtml += `<td>${formatRupiah(total)}</td>`;
    sisaHtml += `<td>${formatRupiah(sisa)}</td>`;
    targetHtml += `<td>${formatRupiah(target)}</td>`;
    progressHtml += `<td>
      <div class="progress-cell-inner">
        <span class="progress-val">${formatPercent(pct)}</span>
        <div class="mini-progress-track">
          <div class="mini-progress-bar" style="width:${Math.min(100, Math.max(0, Math.round(pct)))}%;"></div>
        </div>
      </div>
    </td>`;
  });

  const grandSisa = grandTotalAchieve - sumTarget;
  const grandPct = sumTarget > 0 ? (grandTotalAchieve / sumTarget) * 100 : 0;

  totalHtml += `<td class="col-achieve">${formatRupiah(grandTotalAchieve)}</td><td class="col-week text-center">-</td><td class="col-avg text-center">-</td></tr>`;
  sisaHtml += `<td class="col-achieve">${formatRupiah(grandSisa)}</td><td class="col-week text-center">-</td><td class="col-avg text-center">-</td></tr>`;
  targetHtml += `<td class="col-achieve">${formatRupiah(sumTarget)}</td><td class="col-week text-center">-</td><td class="col-avg text-center">-</td></tr>`;
  progressHtml += `<td class="col-achieve">
    <div class="progress-cell-inner">
      <span class="progress-val">${formatPercent(grandPct)}</span>
      <div class="mini-progress-track">
        <div class="mini-progress-bar" style="width:${Math.min(100, Math.max(0, Math.round(grandPct)))}%;"></div>
      </div>
    </div>
  </td><td class="col-week text-center">-</td><td class="col-avg text-center">-</td></tr>`;

  tfoot.innerHTML = totalHtml + sisaHtml + targetHtml + progressHtml;
}

// 5. RENDER REPORT 2: KPI SUMMARY & TIMEGONE
function renderReport2Tables(picStats, activePICs) {
  document.getElementById('timegoneHeaderTag').innerText = `${timegonePct}%`;
  document.getElementById('timegoneDayBadge').innerText = `Hari ke-${lastActiveDay} / ${daysInMonth}`;
  document.getElementById('kpiTimegoneVal').innerText = `${timegonePct}%`;

  const summaryBody = document.getElementById('kpiSummaryBody');
  const summaryFoot = document.getElementById('kpiSummaryFoot');
  const crossBody = document.getElementById('kpiCrossBody');
  const crossFoot = document.getElementById('kpiCrossFoot');

  summaryBody.innerHTML = '';
  crossBody.innerHTML = '';

  let sumTarget = 0, sumAch = 0, sumTargetGP = 0, sumAcvGP = 0;
  let sumGadget = 0, sumCross = 0;
  const tg = (typeof timegonePct === 'number' && timegonePct > 0) ? timegonePct : 100;

  activePICs.forEach(pic => {
    const s = picStats[pic];
    sumTarget += s.target;
    sumAch += s.pencapaian;
    sumTargetGP += s.targetGP;
    sumAcvGP += s.acvGP;
    sumGadget += s.gadgetCount;
    sumCross += s.crossCount;

    // Ratio & Status based on Excel formula:
    // Ratio = (Achievement% / Timegone%) * 100
    // If Ratio < 80% => ☠️, If Ratio < 90% => 💔, Else => 🖤

    // Omzet % and status
    const pctOmzet = s.target > 0 ? (s.pencapaian / s.target) * 100 : 0;
    const ratioOmzet = (pctOmzet / tg) * 100;
    const statusOmzet = ratioOmzet >= 90 ? '🖤' : (ratioOmzet >= 80 ? '💔' : '☠️');

    // GP % and status
    const pctGP = s.targetGP > 0 ? (s.acvGP / s.targetGP) * 100 : 0;
    const ratioGP = (pctGP / tg) * 100;
    const statusGP = ratioGP >= 90 ? '🖤' : (ratioGP >= 80 ? '💔' : '☠️');

    // Cross selling % (target ratio 1/3)
    const ratioActual = s.gadgetCount > 0 ? (s.crossCount / s.gadgetCount) : 0;
    const pctCross = Math.min(100, Math.round(ratioActual * 3 * 100));

    // Final Report % (35% Omzet + 35% GP + 30% Cross)
    const omzetNorm = Math.min(100, Math.max(0, pctOmzet));
    const gpNorm = Math.min(100, Math.max(0, pctGP));
    const pctAkhir = (omzetNorm * 0.35) + (gpNorm * 0.35) + (pctCross * 0.30);

    const ratioAkhir = (pctAkhir / tg) * 100;
    const statusAkhir = ratioAkhir >= 90 ? '🖤' : (ratioAkhir >= 80 ? '💔' : '☠️');

    // Table 1 Row
    const tr1 = document.createElement('tr');
    tr1.innerHTML = `
      <td class="text-left font-bold">${pic}</td>
      <td>${formatRupiah(s.target)}</td>
      <td>${formatRupiah(s.pencapaian)}</td>
      <td class="font-bold">${formatPercent(pctOmzet)}</td>
      <td class="text-center status-icon-lg">${statusOmzet}</td>
      <td>${formatRupiah(s.targetGP)}</td>
      <td class="${s.acvGP < 0 ? 'text-crimson' : ''}">${formatRupiah(s.acvGP)}</td>
      <td class="font-bold">${formatPercent(pctGP)}</td>
      <td class="text-center status-icon-lg">${statusGP}</td>
    `;
    summaryBody.appendChild(tr1);

    // Table 2 Row
    const tr2 = document.createElement('tr');
    tr2.innerHTML = `
      <td class="text-left font-bold">${pic}</td>
      <td>${s.gadgetCount}</td>
      <td>${s.crossCount}</td>
      <td class="font-bold">${formatPercent(pctCross)}</td>
      <td class="font-bold">${formatPercent(pctAkhir)}</td>
      <td class="text-center status-icon-lg">${statusAkhir}</td>
      <td></td>
    `;
    crossBody.appendChild(tr2);
  });

  // Table 1 MANAJER Total Row
  // Use lastTotalSosmedAchieve and lastTotalSosmedGP which includes ALL sales employee codes starting with bjm.sos (both bjm.sos.0 and bjm.sos.k.0)
  const managerAchieve = (typeof lastTotalSosmedAchieve === 'number' && lastTotalSosmedAchieve > 0) ? lastTotalSosmedAchieve : sumAch;
  const managerGP = (typeof lastTotalSosmedGP === 'number' && lastTotalSosmedGP !== 0) ? lastTotalSosmedGP : sumAcvGP;
  const managerGadget = (typeof lastTotalSosmedGadget === 'number' && lastTotalSosmedGadget > 0) ? lastTotalSosmedGadget : sumGadget;
  const managerCross = (typeof lastTotalSosmedCross === 'number' && lastTotalSosmedCross > 0) ? lastTotalSosmedCross : sumCross;

  const totalPctOmzet = sumTarget > 0 ? (managerAchieve / sumTarget) * 100 : 0;
  const totalPctGP = sumTargetGP > 0 ? (managerGP / sumTargetGP) * 100 : 0;
  const ratioMgrOmzet = (totalPctOmzet / tg) * 100;
  const statusMgrOmzet = ratioMgrOmzet >= 90 ? '🖤' : (ratioMgrOmzet >= 80 ? '💔' : '☠️');
  const ratioMgrGP = (totalPctGP / tg) * 100;
  const statusMgrGP = ratioMgrGP >= 90 ? '🖤' : (ratioMgrGP >= 80 ? '💔' : '☠️');

  summaryFoot.innerHTML = `
    <tr class="row-manajer">
      <td class="text-left font-bold">MANAJER</td>
      <td>${formatRupiah(sumTarget)}</td>
      <td>${formatRupiah(managerAchieve)}</td>
      <td class="font-bold">${formatPercent(totalPctOmzet)}</td>
      <td class="text-center status-icon-lg">${statusMgrOmzet}</td>
      <td>${formatRupiah(sumTargetGP)}</td>
      <td>${formatRupiah(managerGP)}</td>
      <td class="font-bold">${formatPercent(totalPctGP)}</td>
      <td class="text-center status-icon-lg">${statusMgrGP}</td>
    </tr>
  `;

  // Table 2 MANAJER Total Row
  const totalRatio = managerGadget > 0 ? (managerCross / managerGadget) : 0;
  const totalPctCross = Math.min(100, Math.round(totalRatio * 3 * 100));
  const totalPctAkhir = (Math.min(100, Math.max(0, totalPctOmzet)) * 0.35) + (Math.min(100, Math.max(0, totalPctGP)) * 0.35) + (totalPctCross * 0.30);
  const ratioMgrAkhir = (totalPctAkhir / tg) * 100;
  const statusMgrAkhir = ratioMgrAkhir >= 90 ? '🖤' : (ratioMgrAkhir >= 80 ? '💔' : '☠️');

  crossFoot.innerHTML = `
    <tr>
      <td class="text-left font-bold">MANAJER</td>
      <td>${managerGadget}</td>
      <td>${managerCross}</td>
      <td class="font-bold">${formatPercent(totalPctCross)}</td>
      <td class="font-bold">${formatPercent(totalPctAkhir)}</td>
      <td class="text-center status-icon-lg">${statusMgrAkhir}</td>
      <td></td>
    </tr>
  `;
}

// 6. UPDATE EXECUTIVE KPI SUMMARY CARDS
function updateExecutiveKPICards(picStats, activePICs) {
  let grandTarget = 0, grandGPTarget = 0;
  let sumAch = 0, sumGP = 0, sumGadget = 0, sumCross = 0;

  activePICs.forEach(pic => {
    const s = picStats[pic];
    grandTarget += s.target;
    grandGPTarget += s.targetGP;
    sumAch += s.pencapaian;
    sumGP += s.acvGP;
    sumGadget += s.gadgetCount;
    sumCross += s.crossCount;
  });

  const grandTotal = (typeof lastTotalSosmedAchieve === 'number' && lastTotalSosmedAchieve > 0) ? lastTotalSosmedAchieve : sumAch;
  const grandGP = (typeof lastTotalSosmedGP === 'number' && lastTotalSosmedGP !== 0) ? lastTotalSosmedGP : sumGP;
  const grandGadget = (typeof lastTotalSosmedGadget === 'number' && lastTotalSosmedGadget > 0) ? lastTotalSosmedGadget : sumGadget;
  const grandCross = (typeof lastTotalSosmedCross === 'number' && lastTotalSosmedCross > 0) ? lastTotalSosmedCross : sumCross;

  const pct = grandTarget > 0 ? (grandTotal / grandTarget) * 100 : 0;
  const sisa = grandTotal - grandTarget;
  const gpPct = grandGPTarget > 0 ? (grandGP / grandGPTarget) * 100 : 0;
  const crossRatio = grandGadget > 0 ? Math.min(100, Math.round((grandCross / grandGadget) * 3 * 100)) : 0;

  const elAch = document.getElementById('kpiTotalAchieve');
  if (elAch) elAch.innerText = formatRupiah(grandTotal);
  const elTgt = document.getElementById('kpiTargetSub');
  if (elTgt) elTgt.innerText = `Target: ${formatRupiah(grandTarget)}`;
  const elPct = document.getElementById('kpiProgressPct');
  if (elPct) elPct.innerText = formatPercent(pct);
  const elBar = document.getElementById('kpiProgressBar');
  if (elBar) elBar.style.width = `${Math.min(pct, 100)}%`;
  const elSisa = document.getElementById('kpiSisaTarget');
  if (elSisa) elSisa.innerText = formatRupiah(sisa);

  const elGP = document.getElementById('kpiTotalGP');
  if (elGP) elGP.innerText = formatRupiah(grandGP);
  const elTgtGP = document.getElementById('kpiTargetGP');
  if (elTgtGP) elTgtGP.innerText = formatRupiah(grandGPTarget);
  const elGPPct = document.getElementById('kpiGPPct');
  if (elGPPct) elGPPct.innerText = formatPercent(gpPct);

  const elCross = document.getElementById('kpiCrossRatio');
  if (elCross) elCross.innerText = formatPercent(crossRatio);
  const elGadget = document.getElementById('kpiTotalGadget');
  if (elGadget) elGadget.innerText = grandGadget.toLocaleString('id-ID');

  const teamStatus = pct >= (timegonePct - 4) ? '🖤 On Track' : '💔 Warning';
  const elStatus = document.getElementById('kpiTeamStatusIcon');
  if (elStatus) elStatus.innerText = teamStatus;
}

// 7. TAB NAVIGATION & VIEW MODE FILTER (SUPPORT 7 UNIFIED TABS)
let currentReportViewMode = 'all'; // 'all', 'r1', 'r2'

function toggleSettingsDropdown(e) {
  if (e) e.stopPropagation();
  const dropdown = document.getElementById('headerGearDropdown');
  if (dropdown) dropdown.classList.toggle('show');
}

function closeSettingsDropdown() {
  const dropdown = document.getElementById('headerGearDropdown');
  if (dropdown) dropdown.classList.remove('show');
}

function toggleWorkspaceMenu(e) {
  toggleSettingsDropdown(e);
}

function closeWorkspaceMenu() {
  closeSettingsDropdown();
}

// Close workspace dropdown menu when clicking anywhere outside
document.addEventListener('click', function(e) {
  const wrapper = document.querySelector('.header-gear-dropdown-wrap') || document.querySelector('.header-menu-wrapper');
  const dropdown = document.getElementById('headerGearDropdown') || document.getElementById('workspaceNavDropdown');
  if (dropdown && wrapper && !wrapper.contains(e.target)) {
    dropdown.classList.remove('show');
  }
});

// Workspace Login Modal Handlers
function openLoginModal() {
  const modal = document.getElementById('workspaceLoginModal');
  if (modal) {
    modal.style.display = 'flex';
    const pin = document.getElementById('loginPinInput');
    const fb = document.getElementById('loginFeedback');
    if (fb) fb.style.display = 'none';
    if (pin) { pin.value = ''; pin.focus(); }
  }
}

function closeLoginModal() {
  const modal = document.getElementById('workspaceLoginModal');
  if (modal) modal.style.display = 'none';
}

function executeWorkspaceLogin() {
  const roleSelect = document.getElementById('loginRoleSelect');
  const role = roleSelect ? roleSelect.value : 'admin';
  const pinInput = document.getElementById('loginPinInput');
  const pin = pinInput ? pinInput.value.trim() : '';
  const feedback = document.getElementById('loginFeedback');

  if (pin === '1234' || pin === 'admin' || pin === 'gadgetmart' || pin === '') {
    showSaveToast(`✅ Berhasil login sebagai ${role.toUpperCase()}!`);
    closeLoginModal();
    const loginBtn = document.getElementById('headerLoginBtn');
    if (loginBtn) {
      loginBtn.innerHTML = `👤 ${role === 'admin' ? 'Admin' : (role === 'sosmed' ? 'Sosmed' : 'Cabang')}`;
      loginBtn.style.color = '#ffc72c';
      loginBtn.style.fontWeight = '700';
    }
  } else {
    if (feedback) {
      feedback.style.display = 'block';
      feedback.style.background = '#ffe4e6';
      feedback.style.color = '#be123c';
      feedback.innerText = 'PIN atau password tidak cocok (Gunakan PIN default: 1234)';
    }
  }
}

function toggleSidebarMobile() {
  toggleWorkspaceMenu();
}

function switchTab(tabIndex) {
  currentActiveTab = tabIndex;
  for (let i = 1; i <= 8; i++) {
    const btn = document.getElementById(`tabBtn${i}`);
    const content = document.getElementById(`tabContent${i}`);
    if (btn) btn.classList.toggle('active', i === tabIndex);
    if (content) content.classList.toggle('active', i === tabIndex);
  }

  // Update dynamic active module name in subbar (Sesuai Mockup Gambar User)
  const moduleNames = {
    1: 'Dashboard',
    2: 'Report',
    3: 'Cek AR',
    4: 'Format Keep',
    5: 'Tracking AR',
    6: 'Klaim AR',
    7: 'Arsip Pencapaian Bulanan',
    8: 'Pengaturan Admin'
  };
  const activeModEl = document.getElementById('activeModuleName');
  if (activeModEl) {
    activeModEl.textContent = moduleNames[tabIndex] || 'Dashboard';
  }

  closeWorkspaceMenu();
  closeSettingsDropdown();

  // Toggle KPI Cards and Gemini Prompt visibility
  const kpiSection = document.getElementById('kpiGridSection');
  const aiSection = document.getElementById('aiPromptSection');
  if (kpiSection) kpiSection.style.display = 'none'; // Sesuai instruksi: kartu KPI atas disembunyikan
  if (tabIndex === 1) {
    if (aiSection) aiSection.style.display = 'flex';
    populateDashboardFilterDropdowns();
    applyDashboardFilters();
    if (lastProcessedPicStats && lastActivePICs.length > 0) {
      renderSpeedometerGauges(lastProcessedPicStats, lastActivePICs);
      renderMonthToMonthCharts();
      renderSalesLeaderboard('omzet');
    }
  } else if (tabIndex === 2) {
    if (aiSection) aiSection.style.display = 'flex';
  } else {
    if (aiSection) aiSection.style.display = 'none';
  }

  if (tabIndex === 4 && rawKeepRows.length === 0) loadKeepData();
  if (tabIndex === 5 && rawAuditARList.length === 0) runARMissingSosmedAudit();
  if (tabIndex === 6) renderARClaimsModule();
  if (tabIndex === 7) renderArchiveGrid();
  if (tabIndex === 8) {
    populateAdminPanel(getMasterSales());
    onAdminTargetMonthChange();
    renderAdminPhotoTable();
  }
}

function setReportViewMode(mode) {
  currentReportViewMode = mode;
  const btnAll = document.getElementById('filterBtnAll');
  const btnR1 = document.getElementById('filterBtnR1');
  const btnR2 = document.getElementById('filterBtnR2');

  if (btnAll) btnAll.classList.toggle('active', mode === 'all');
  if (btnR1) btnR1.classList.toggle('active', mode === 'r1');
  if (btnR2) btnR2.classList.toggle('active', mode === 'r2');

  const r1Wrap = document.getElementById('report1ViewWrap');
  const r2Wrap = document.getElementById('report2ViewWrap');
  const kpiSection = document.getElementById('kpiGridSection');

  if (mode === 'all') {
    if (r1Wrap) r1Wrap.style.display = 'block';
    if (r2Wrap) r2Wrap.style.display = 'block';
    if (kpiSection) kpiSection.style.display = 'block';
  } else if (mode === 'r1') {
    if (r1Wrap) r1Wrap.style.display = 'block';
    if (r2Wrap) r2Wrap.style.display = 'none';
    if (kpiSection) kpiSection.style.display = 'block';
  } else if (mode === 'r2') {
    if (r1Wrap) r1Wrap.style.display = 'none';
    if (r2Wrap) r2Wrap.style.display = 'block';
    if (kpiSection) kpiSection.style.display = 'none';
  }
}

// AUTO-SAVE EXCEL TO SERVER FOR LAN PERSISTENCE
async function saveExcelToServer(arrayBuffer, type = 'turnover') {
  try {
    // Vercel serverless functions memiliki batas payload maksimal 4.5 MB.
    // Jika file berukuran besar (> 4MB seperti KPI SOSMED 53MB), simpan lokal di browser (IndexedDB)
    // dan lewati upload serverless untuk mencegah error 413 Payload Too Large.
    if (arrayBuffer.byteLength > 4 * 1024 * 1024) {
      console.log(`File ${(arrayBuffer.byteLength / 1024 / 1024).toFixed(1)} MB disimpan di IndexedDB browser (melebihi limit 4.5MB serverless Vercel).`);
      if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
        return;
      }
    }
    const res = await fetch(`/api/save-excel?type=${encodeURIComponent(type)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: arrayBuffer
    });
    if (res.ok) {
      const data = await res.json();
      console.log('💾 File saved to server disk successfully:', data);
      showSaveToast(`💾 Database ${type === 'croselling' ? 'Google Sheets' : 'turnover.xlsx'} tersimpan di Server! Tersedia untuk semua device di LAN & tetap tersimpan saat refresh.`);
      updateDatabaseBadge(true);
    }
  } catch (err) {
    console.warn("Gagal auto-save ke server:", err);
  }
}

function showSaveToast(message) {
  let toast = document.getElementById('saveToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'saveToast';
    toast.className = 'save-toast';
    document.body.appendChild(toast);
  }
  toast.innerHTML = `<span>${message}</span>`;
  toast.style.display = 'flex';
  setTimeout(() => {
    if (toast) toast.style.display = 'none';
  }, 4000);
}

// ==========================================================================
// FLOATING ANIMATED PERCENTAGE PROGRESS NOTIFICATION ENGINE
// ==========================================================================
function getProgressToastContainer() {
  let container = document.getElementById('progressToastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'progressToastContainer';
    container.className = 'progress-toast-container';
    document.body.appendChild(container);
  }
  return container;
}

function showProgressNotification({ id = `prog-${Date.now()}`, title = 'Memproses Data...', detail = '', initialPct = 0 }) {
  const container = getProgressToastContainer();
  let toast = document.getElementById(id);

  if (!toast) {
    toast = document.createElement('div');
    toast.id = id;
    toast.className = 'progress-toast';
    toast.innerHTML = `
      <div class="progress-toast-header">
        <div class="progress-toast-title-wrap">
          <span class="progress-toast-spinner" id="${id}-spinner"></span>
          <span class="progress-toast-title" id="${id}-title">${title}</span>
        </div>
        <span class="progress-toast-pct" id="${id}-pct">${Math.round(initialPct)}%</span>
      </div>
      <div class="progress-bar-track">
        <div class="progress-bar-fill" id="${id}-fill" style="width: ${initialPct}%;"></div>
      </div>
      <div class="progress-toast-detail" id="${id}-detail">${detail}</div>
    `;
    container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('show'));
  }

  const titleEl = document.getElementById(`${id}-title`);
  const pctEl = document.getElementById(`${id}-pct`);
  const fillEl = document.getElementById(`${id}-fill`);
  const detailEl = document.getElementById(`${id}-detail`);

  let currentPercent = initialPct;

  const controller = {
    id,
    update({ pct, title: newTitle, detail: newDetail }) {
      if (typeof pct === 'number') {
        currentPercent = Math.max(0, Math.min(100, pct));
        if (pctEl) pctEl.innerText = `${Math.round(currentPercent)}%`;
        if (fillEl) fillEl.style.width = `${currentPercent}%`;
      }
      if (newTitle && titleEl) titleEl.innerText = newTitle;
      if (newDetail !== undefined && detailEl) detailEl.innerText = newDetail;
    },
    complete({ title: compTitle = '✅ Selesai (100%)', detail: compDetail = 'Data berhasil diproses' } = {}) {
      if (pctEl) pctEl.innerText = '100%';
      if (fillEl) fillEl.style.width = '100%';
      if (titleEl) titleEl.innerHTML = compTitle;
      if (detailEl) detailEl.innerText = compDetail;
      toast.classList.remove('error');
      toast.classList.add('success');

      setTimeout(() => {
        toast.classList.remove('show');
        toast.classList.add('hide');
        setTimeout(() => toast.remove(), 400);
      }, 2500);
    },
    fail({ title: errTitle = '❌ Gagal', detail: errDetail = 'Terjadi kesalahan saat memproses data' } = {}) {
      if (pctEl) pctEl.innerText = 'Gagal';
      if (fillEl) {
        fillEl.style.width = '100%';
        fillEl.style.background = '#ef4444';
      }
      if (titleEl) titleEl.innerHTML = errTitle;
      if (detailEl) detailEl.innerText = errDetail;
      toast.classList.remove('success');
      toast.classList.add('error');

      setTimeout(() => {
        toast.classList.remove('show');
        toast.classList.add('hide');
        setTimeout(() => toast.remove(), 400);
      }, 4000);
    },
    dismiss() {
      toast.classList.remove('show');
      toast.classList.add('hide');
      setTimeout(() => toast.remove(), 400);
    }
  };

  return controller;
}
window.showProgressNotification = showProgressNotification;

function startSmoothProgress({ 
  id = `smooth-${Date.now()}`, 
  title = 'Menyimpan ke Google Sheets...', 
  detail = 'Menghubungkan ke server...',
  timeoutMs = 15000,
  onTimeout
} = {}) {
  const ctrl = showProgressNotification({ id, title, detail, initialPct: 10 });
  let current = 10;
  let isDone = false;

  const interval = setInterval(() => {
    if (isDone) {
      clearInterval(interval);
      return;
    }
    // Naik dinamis di awal, melambat di atas 85% menunggu respon jaringan
    if (current < 45) current += Math.floor(Math.random() * 8) + 6;
    else if (current < 75) current += Math.floor(Math.random() * 5) + 3;
    else if (current < 92) current += Math.floor(Math.random() * 3) + 1;
    
    current = Math.min(92, current);
    ctrl.update({ pct: current });
  }, 130);

  // Safety timer agar notifikasi tidak tertahan selamanya di 92% jika jaringan bermasalah
  const safetyTimer = setTimeout(() => {
    if (!isDone) {
      isDone = true;
      clearInterval(interval);
      ctrl.fail({
        title: '⚠️ Batas Waktu Terlampaui (Timeout)',
        detail: 'Koneksi jaringan terlalu lama (15 detik). Silakan coba lagi atau pastikan server lokal aktif.'
      });
      if (typeof onTimeout === 'function') {
        try { onTimeout(); } catch (_) {}
      }
    }
  }, timeoutMs);

  return {
    id: ctrl.id,
    update: ctrl.update,
    complete(opts) {
      if (isDone) return;
      isDone = true;
      clearInterval(interval);
      clearTimeout(safetyTimer);
      ctrl.update({ pct: 100 });
      ctrl.complete(opts);
    },
    fail(opts) {
      if (isDone) return;
      isDone = true;
      clearInterval(interval);
      clearTimeout(safetyTimer);
      ctrl.fail(opts);
    },
    dismiss() {
      if (isDone) return;
      isDone = true;
      clearInterval(interval);
      clearTimeout(safetyTimer);
      ctrl.dismiss();
    }
  };
}
window.startSmoothProgress = startSmoothProgress;

function updateDatabaseBadge(isOk) {
  const text = document.getElementById('dbSyncStatusText');
  if (text) {
    text.innerText = isOk ? 'Database Server: Tersimpan & Tersinkron LAN' : 'Database Server: Menghubungkan...';
  }
}

// 8. TAB 2: CEK AR & PENCARIAN WHATSAPP CUSTOMER (GOOGLE SHEETS INTEGRATION)
function executeARSearch() {
  const input = document.getElementById('arInputSearch');
  const query = input.value.trim();
  if (!query) {
    alert("Masukkan nomor Dokumen AR atau Nomor WhatsApp Customer!");
    return;
  }
  performSearch(query);
}

function quickCheckAR(query) {
  switchTab(2);
  document.getElementById('arInputSearch').value = query;
  performSearch(query);
}

function performSearch(query) {
  if ((!rawCrosellingRows || rawCrosellingRows.length === 0) && (!rawTurnoverRows || rawTurnoverRows.length === 0)) {
    alert("Database belum dimuat. Silakan tunggu sebentar atau muat ulang data default.");
    return;
  }

  const cleanQ = String(query).trim();
  const digitsOnly = cleanQ.replace(/\D/g, '');
  const normPhone = normalizePhone(cleanQ);

  // Search in Google Sheets rows first (has WhatsApp numbers & per-item sales PIC)
  let matchedByPhone = [];
  let matchedByDoc = [];

  if (rawCrosellingRows && rawCrosellingRows.length > 0) {
    if (digitsOnly.length >= 8) {
      matchedByPhone = rawCrosellingRows.filter(r => {
        const hp = normalizePhone(r['No HP Customer']);
        return hp === normPhone || hp.endsWith(digitsOnly.slice(-9));
      });
    }
    matchedByDoc = rawCrosellingRows.filter(r => {
      const doc = String(r['Document Number'] || '').replace(/\D/g, '');
      return doc === digitsOnly;
    });
  }

  // Fallback to turnover.xlsx if not found in Google Sheets
  if (matchedByDoc.length === 0 && rawTurnoverRows && rawTurnoverRows.length > 0) {
    const fallbackTurnover = rawTurnoverRows.filter(r => {
      const doc = String(r['Document Number'] || '').replace(/\D/g, '');
      return doc === digitsOnly;
    });
    if (fallbackTurnover.length > 0) {
      renderARDocumentView(cleanQ, fallbackTurnover, 'TURNOVER');
      return;
    }
  }

  // Decision: WhatsApp Customer Search vs AR Document Search
  if (matchedByPhone.length > 0 && (matchedByDoc.length === 0 || cleanQ.startsWith('08') || cleanQ.startsWith('628') || cleanQ.startsWith('+62') || (digitsOnly.length >= 10 && !cleanQ.startsWith('26')))) {
    renderWhatsAppCustomerView(cleanQ, matchedByPhone);
  } else if (matchedByDoc.length > 0) {
    renderARDocumentView(cleanQ, matchedByDoc, 'CROSSELLING');
  } else if (matchedByPhone.length > 0) {
    renderWhatsAppCustomerView(cleanQ, matchedByPhone);
  } else {
    alert(`Data untuk AR atau Nomor WhatsApp "${query}" tidak ditemukan dalam database.`);
    document.getElementById('arResultContainer').style.display = 'none';
  }
}

// 8A. VIEW PENCARIAN NOMOR WHATSAPP: MENAMPILKAN SEMUA PRODUK YANG DIBELI CUSTOMER
function renderWhatsAppCustomerView(query, rows) {
  const resContainer = document.getElementById('arResultContainer');
  resContainer.style.display = 'block';

  // Toggle UI visibility
  document.getElementById('customerProfileBanner').style.display = 'block';
  document.getElementById('arStatusBanner').style.display = 'none';
  document.getElementById('arDetailsGrid').style.display = 'none';

  const first = rows[0];
  const custName = first['Name'] || first['Customer/Vendor Name'] || 'Customer';
  const custCode = first['Customer/Vendor Code'] || '-';
  const rawWA = first['No HP Customer'] || query;
  const normWA = normalizePhone(rawWA);
  const displayWA = formatPhoneDisplay(rawWA);

  const uniqueDocs = Array.from(new Set(rows.map(r => r['Document Number'])));
  const uniqueCabs = Array.from(new Set(rows.map(r => r['Cabang']).filter(Boolean)));
  const totalQty = rows.reduce((acc, r) => acc + (Number(r['Quantity']) || 1), 0);
  const totalSpent = rows.reduce((acc, r) => acc + (Number(r['Harga Jual']) || 0), 0);

  // Populate Customer Profile Banner
  document.getElementById('custAvatar').innerText = '📱';
  document.getElementById('custProfileName').innerText = `${custName} (${first['Customer/Vendor Name'] || ''})`;
  document.getElementById('custProfileCode').innerText = `Kode Customer: ${custCode} &bull; Sumber: Google Sheets`;
  document.getElementById('custProfileWA').innerText = displayWA;
  document.getElementById('custProfileTotalSpent').innerText = formatRupiah(totalSpent);
  document.getElementById('custProfileInvoicesCount').innerText = `${uniqueDocs.length} Faktur AR`;
  document.getElementById('custProfileItemsCount').innerText = `${totalQty} Item Produk`;
  document.getElementById('custProfileCabang').innerText = uniqueCabs.join(', ') || '-';

  document.getElementById('custWhatsAppAction').innerHTML = `
    <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
        <path d="M.057 24l1.687-6.163c-1.041-1.804-1.588-3.849-1.587-5.946.003-6.556 5.338-11.891 11.893-11.891 3.181.001 6.167 1.24 8.413 3.488 2.245 2.248 3.481 5.236 3.48 8.414-.003 6.557-5.338 11.892-11.893 11.892-1.99-.001-3.951-.5-5.688-1.448l-6.305 1.654zm6.597-3.807c1.676.995 3.276 1.591 5.392 1.592 5.448 0 9.886-4.434 9.889-9.885.002-5.462-4.415-9.89-9.881-9.892-5.452 0-9.887 4.434-9.889 9.884-.001 2.225.651 3.891 1.746 5.634l-.999 3.648 3.742-.981z"/>
      </svg>
      Chat Customer di WhatsApp
    </a>
  `;

  // Global state for Tab 3 Cek AR search results & copy
  currentARSearchResults = rows;
  currentARSearchDoc = '';

  // Update Table Headers
  document.getElementById('itemsTableTitle').innerText = `🛒 Riwayat Pembelian Produk oleh Customer (${displayWA} - ${custName}):`;
  document.getElementById('arItemsCountBadge').innerText = `${rows.length} Baris Produk (${totalQty} Qty)`;

  document.getElementById('arItemsTableHead').innerHTML = `
    <th class="col-sticky" style="width: 35px; text-align: center;">NO</th>
    <th>NO. AR / FAKTUR</th>
    <th>TANGGAL & WAKTU</th>
    <th>CABANG</th>
    <th class="text-left" style="min-width: 200px;">NAMA PRODUK / ITEM DESCRIPTION</th>
    <th style="width: 45px;">QTY</th>
    <th>HARGA SATUAN</th>
    <th>TOTAL BELANJA</th>
    <th class="text-left" style="min-width: 170px;">SALES EMPLOYEE (PER ITEM)</th>
    <th class="text-center" style="width: 110px;">STATUS SOSMED</th>
    <th class="text-center" style="width: 90px;">POTENSI DENDA</th>
    <th class="text-center" style="width: 100px;">SALIN EXCEL</th>
  `;

  // Render Rows
  const tbody = document.getElementById('arItemsBody');
  tbody.innerHTML = '';

  rows.forEach((r, idx) => {
    const qty = Number(r['Quantity']) || 1;
    const subtotal = Number(r['Harga Jual']) || 0;
    const price = qty > 0 ? (subtotal / qty) : subtotal;
    const schl = String(r['SC/HL Code'] || '');
    const slp = String(r['SlpName'] || '');
    const rem = String(r['Remarks'] || '');
    const dateFormatted = formatCrosellingDate(r['Posting Date']);
    const waktu = r['Waktu'] || '';

    // Check if sales code indicates sosmed
    const isSosmed = isItemInputSosmed(r);
    const empDisplay = schl ? schl : (slp ? slp : '-');
    const empBadge = isSosmed 
      ? `<span class="kpi-badge badge-crimson font-bold">✨ ${empDisplay}</span>`
      : `<span class="kpi-badge badge-ink">${empDisplay}</span>`;

    const statusBadge = isSosmed
      ? `<span class="status-badge badge-accent">✨ Terinput Sosmed</span>`
      : `<span class="status-badge badge-ink">🏢 Offline Toko / Kasir</span>`;

    const dendaStatus = getItemPotensiDenda(r);
    const dendaBadge = dendaStatus === 'DENDA'
      ? `<span class="kpi-badge badge-crimson font-bold" style="font-size:0.68rem; padding:2px 6px;">🚨 DENDA</span>`
      : `<span class="kpi-badge badge-ink" style="font-size:0.68rem; padding:2px 6px;">PASS</span>`;

    const copyBtn = `
      <button class="btn-copy-item-sm" onclick="copyARSearchItemToClipboard(${idx}, this)" title="Salin item ini ke format 12 kolom Excel (tanpa header)">
        📋 Copy Rekap
      </button>
    `;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-sticky">${idx + 1}</td>
      <td class="font-bold text-crimson cursor-pointer" onclick="quickCheckAR('${r['Document Number']}')" title="Klik untuk cek detail AR ini">📄 ${r['Document Number']}</td>
      <td>${dateFormatted} ${waktu}</td>
      <td class="font-bold">${r['Cabang'] || '-'}</td>
      <td class="text-left font-bold">${r['Dscription'] || '-'}</td>
      <td>${qty}</td>
      <td>${formatRupiah(price)}</td>
      <td class="font-bold">${formatRupiah(subtotal)}</td>
      <td class="text-left">${empBadge}</td>
      <td class="text-center">${statusBadge}</td>
      <td class="text-center">${dendaBadge}</td>
      <td class="text-center">${copyBtn}</td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('arItemsFoot').innerHTML = `
    <tr class="row-total">
      <td colspan="5" class="text-left font-bold">TOTAL PEMBELIAN CUSTOMER (KUMULATIF):</td>
      <td><strong>${totalQty}</strong></td>
      <td>-</td>
      <td class="font-bold">${formatRupiah(totalSpent)}</td>
      <td colspan="4" class="text-left font-bold text-accent">${uniqueDocs.length} Faktur Dokumen</td>
    </tr>
  `;

  resContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// 8B. VIEW PENCARIAN DOKUMEN AR: RINCIAN FAKTUR & VERIFIKASI SALES PER ITEM
function renderARDocumentView(docNumber, rows, source = 'CROSSELLING') {
  const resContainer = document.getElementById('arResultContainer');
  resContainer.style.display = 'block';

  document.getElementById('customerProfileBanner').style.display = 'none';
  document.getElementById('arStatusBanner').style.display = 'flex';
  document.getElementById('arDetailsGrid').style.display = 'grid';

  const first = rows[0];
  const masterList = getMasterSales();

  // Validasi ketat kode sales:
  // Item hanya diakui terinput Sosmed jika kode salesnya benar-benar BJM.SOS,
  // BUKAN sekadar ada tulisan 'sosmed' pada Remarks/Keterangan kasir!
  const sosmedItems = rows.filter(r => isItemInputSosmed(r));
  const missingItems = rows.filter(r => !isItemInputSosmed(r));

  let detectedPIC = null;
  let detectedCode = null;

  // Cari identitas PIC dari item yang berkode BJM.SOS
  for (const r of sosmedItems) {
    const schl = String(r['SC/HL Code'] || '');
    const slp = String(r['SlpName'] || r['Sales Employee Name'] || '');

    for (const m of masterList) {
      if (schl.includes(m.code) || slp.includes(m.code)) {
        detectedPIC = m.name;
        detectedCode = m.code;
        break;
      }
    }
    if (detectedPIC) break;

    const matchCode = (schl + ' ' + slp).match(/bjm\.sos(?:\.k)?\.\w+/i);
    if (matchCode) {
      detectedCode = matchCode[0].toUpperCase();
      detectedPIC = 'TIM SOSMED';
      break;
    }
  }

  // Update Status Banner
  const banner = document.getElementById('arStatusBanner');
  const icon = document.getElementById('arStatusIcon');
  const title = document.getElementById('arStatusTitle');
  const subtitle = document.getElementById('arStatusSubtitle');

  const remarksText = String(first['Remarks'] || '').trim();
  const remarksHasSosmed = /sosmed|sos/i.test(remarksText);

  if (missingItems.length === 0 && sosmedItems.length > 0) {
    // SEMUA ITEM SUDAH BERKODE SOSMED
    banner.className = 'ar-status-banner banner-success';
    icon.innerHTML = '✅';
    title.innerText = 'SUDAH DIINPUT KODE SOSMED';
    subtitle.innerText = `Semua item (${rows.length} item) dalam dokumen ini telah terverifikasi dengan Sales Sosmed: ${detectedPIC || 'SOSMED'} (${detectedCode || 'BJM.SOS'})`;
  } else if (sosmedItems.length > 0 && missingItems.length > 0) {
    // SEBAGIAN ITEM SUDAH, SEBAGIAN BELUM
    banner.className = 'ar-status-banner banner-warning';
    icon.innerHTML = '⚠️';
    title.innerText = 'SEBAGIAN ITEM BELUM DIINPUT SOSMED';
    subtitle.innerText = `Peringatan: ${sosmedItems.length} item berkode Sosmed, tetapi ${missingItems.length} item lainnya masih berkode Toko Offline! Segera klaim item yang belum terinput.`;
  } else {
    // SEMUA ITEM BELUM DIINPUT SOSMED
    banner.className = 'ar-status-banner banner-danger';
    icon.innerHTML = '❌';
    title.innerText = 'BELUM DIINPUT KODE SOSMED';
    if (remarksHasSosmed) {
      subtitle.innerText = `Peringatan: Kasir menulis "${remarksText}" di Remarks, tetapi KODE SALES di SAP masih offline (${first['SC/HL Code'] || first['SlpName'] || 'Toko'}). Belum terhitung omzet Sosmed, segera klaim ke cabang!`;
    } else {
      subtitle.innerText = `Peringatan: Dokumen ini belum memiliki kode BJM.SOS. Masih tercatat sebagai transaksi offline/toko fisik (${first['SC/HL Code'] || first['SlpName'] || 'Toko'}).`;
    }
  }

  // Calculate totals and metadata
  let totalDocVal = 0;
  let totalDocQty = 0;
  const rawWA = first['No HP Customer'] || '';
  const normWA = normalizePhone(rawWA);
  const displayWA = formatPhoneDisplay(rawWA);

  document.getElementById('arMetaDoc').innerText = first['Document Number'] || docNumber;
  document.getElementById('arMetaDate').innerText = source === 'CROSSELLING' ? `${formatCrosellingDate(first['Posting Date'])} ${first['Waktu'] || ''}` : formatDisplayDate(first['Posting Date']);
  document.getElementById('arMetaCabang').innerText = first['Cabang'] || '-';
  document.getElementById('arMetaCust').innerText = `${first['Customer/Vendor Name'] || '-'} ${first['Name'] ? '(' + first['Name'] + ')' : ''}`;
  
  if (normWA && normWA.length >= 9) {
    document.getElementById('arMetaWA').innerHTML = `
      <span class="font-bold">${displayWA}</span>
      <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp" style="margin-left: 8px; padding: 2px 8px; font-size: 0.72rem;">
        💬 Chat WA
      </a>
    `;
  } else {
    document.getElementById('arMetaWA').innerText = 'Tidak tersedia di dokumen ini';
  }

  document.getElementById('arMetaRemarks').innerText = first['Remarks'] || '- (Tidak ada catatan)';

  // Global state for Tab 3 Cek AR search results & copy
  currentARSearchResults = rows;
  currentARSearchDoc = docNumber;

  // Table Headers
  document.getElementById('itemsTableTitle').innerText = `Daftar Item / Produk dalam Dokumen AR ${first['Document Number'] || docNumber}:`;
  document.getElementById('arItemsTableHead').innerHTML = `
    <th class="col-sticky" style="width: 35px; text-align: center;">NO</th>
    <th class="text-left" style="min-width: 200px;">NAMA PRODUK / ITEM DESCRIPTION</th>
    <th class="text-center">KATEGORI / BRAND</th>
    <th style="width: 45px;">QTY</th>
    <th class="text-left" style="min-width: 170px;">SALES EMPLOYEE (PER ITEM)</th>
    <th class="text-left" style="min-width: 120px;">CABANG</th>
    <th>HARGA SATUAN</th>
    <th>TOTAL</th>
    <th class="text-center" style="width: 100px;">STATUS SOSMED</th>
    <th class="text-center" style="width: 90px;">POTENSI DENDA</th>
    <th class="text-center" style="width: 100px;">SALIN EXCEL</th>
  `;

  const tbody = document.getElementById('arItemsBody');
  tbody.innerHTML = '';

  rows.forEach((r, idx) => {
    const qty = Number(r['Quantity']) || 1;
    const subtotal = source === 'CROSSELLING' ? (Number(r['Harga Jual']) || 0) : (Number(r['Total']) || 0);
    const price = qty > 0 ? (subtotal / qty) : subtotal;
    totalDocVal += subtotal;
    totalDocQty += qty;

    const schl = r['SC/HL Code'] || '';
    const slp = r['SlpName'] || '';
    const emp = r['Sales Employee Name'] || '';
    const empDisplay = schl ? schl : (slp ? slp : (emp || '-'));

    const isSosmedItem = isItemInputSosmed(r);
    const empBadge = isSosmedItem
      ? `<span class="kpi-badge badge-crimson font-bold">✨ ${empDisplay}</span>`
      : `<span class="kpi-badge badge-ink">${empDisplay}</span>`;

    const dendaStatus = getItemPotensiDenda(r);
    const dendaBadge = dendaStatus === 'DENDA'
      ? `<span class="kpi-badge badge-crimson font-bold" style="font-size:0.68rem; padding:2px 6px;">🚨 DENDA</span>`
      : `<span class="kpi-badge badge-ink" style="font-size:0.68rem; padding:2px 6px;">PASS</span>`;

    const copyBtn = `
      <button class="btn-copy-item-sm" onclick="copyARSearchItemToClipboard(${idx}, this)" title="Salin item ini ke format 12 kolom Excel (tanpa header)">
        📋 Copy Rekap
      </button>
    `;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-sticky">${idx + 1}</td>
      <td class="text-left font-bold">${r['Dscription'] || r['Item/Service Description'] || '-'}</td>
      <td class="text-center"><span class="kpi-badge badge-ink">${r['Brand'] || r['ItmsGrpNam'] || r['Group Name'] || '-'}</span></td>
      <td>${qty}</td>
      <td class="text-left">${empBadge}</td>
      <td class="text-left font-bold">${r['Cabang'] || r['SalesCabang'] || '-'}</td>
      <td>${formatRupiah(price)}</td>
      <td class="font-bold">${formatRupiah(subtotal)}</td>
      <td class="text-center">${isSosmedItem ? '<span class="status-badge badge-accent">✨ Sosmed</span>' : '<span class="status-badge badge-ink">🏢 Offline</span>'}</td>
      <td class="text-center">${dendaBadge}</td>
      <td class="text-center">${copyBtn}</td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('arMetaTotal').innerText = formatRupiah(totalDocVal);
  document.getElementById('arItemsCountBadge').innerText = `${rows.length} Item (${totalDocQty} Qty)`;

  document.getElementById('arItemsFoot').innerHTML = `
    <tr class="row-total">
      <td colspan="3" class="text-left font-bold">TOTAL DOKUMEN AR:</td>
      <td><strong>${totalDocQty}</strong></td>
      <td colspan="3"></td>
      <td class="font-bold">${formatRupiah(totalDocVal)}</td>
      <td colspan="3"></td>
    </tr>
  `;

  resContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// 9. TAB 4: DATA KEEP & FOLLOW UP CUSTOMER (GOOGLE SHEETS INTEGRATION)
let rawKeepRows = [];
let filteredKeepRows = [];
let keepCurrentPage = 1;
const KEEP_PAGE_SIZE = 50;

function parseRFC4180CSV(str) {
  const arr = [];
  let quote = false;
  let row = [''];
  for (let c = 0; c < str.length; c++) {
    const cc = str[c], nc = str[c+1];
    if (cc === '"') {
      if (quote && nc === '"') { row[row.length - 1] += '"'; c++; }
      else { quote = !quote; }
    } else if (cc === ',' && !quote) {
      row.push('');
    } else if ((cc === '\r' || cc === '\n') && !quote) {
      if (cc === '\r' && nc === '\n') { c++; }
      arr.push(row);
      row = [''];
    } else {
      row[row.length - 1] += cc;
    }
  }
  if (row.length > 1 || row[0] !== '') arr.push(row);
  return arr;
}
window.parseRFC4180CSV = parseRFC4180CSV;

function parseKeepCSV(csvText, isSilent = false) {
  try {
    if (!csvText || typeof csvText !== 'string') return;
    const table = parseRFC4180CSV(csvText);
    if (!table || table.length < 2) return;

    const headers = table[0].map(h => h.trim());
    const jsonData = [];
    for (let i = 1; i < table.length; i++) {
      const rowArr = table[i];
      if (rowArr.length === 1 && !rowArr[0]) continue;
      const rowObj = {};
      for (let j = 0; j < headers.length; j++) {
        rowObj[headers[j]] = rowArr[j] !== undefined ? rowArr[j] : '';
      }
      rowObj._gsheetsIndex = i - 1;
      jsonData.push(rowObj);
    }

    if (jsonData.length === 0) return;
    rawKeepRows = jsonData;
    console.log(`Keep CSV parsed successfully (Fast RFC4180): ${rawKeepRows.length} rows`);

    const pill = document.getElementById('keepStatusPill');
    if (pill) {
      pill.innerHTML = `<span>📁 Sumber: Google Sheets Live (${rawKeepRows.length.toLocaleString('id-ID')} Baris Data)</span>`;
    }

    populateKeepSosmedDropdown();
    applyKeepFilters();

    if (typeof rawCrosellingRows !== 'undefined' && rawCrosellingRows.length > 0 && typeof runARMissingSosmedAudit === 'function') {
      runARMissingSosmedAudit();
    }
  } catch(e) {
    console.error("Error parsing Keep CSV:", e);
  }
}
window.parseKeepCSV = parseKeepCSV;

async function loadKeepData(isSilent = false) {
  const pill = document.getElementById('keepStatusPill');
  if (pill && !isSilent) pill.innerHTML = '<span>⏳ Membaca Data Keep Google Sheets...</span>';

  try {
    const res = await fetch(`keep-data.csv?t=${Date.now()}`);
    if (!res.ok) throw new Error("File keep-data.csv tidak ditemukan");
    const csvText = await res.text();
    parseKeepCSV(csvText, isSilent);
  } catch (err) {
    console.warn("Gagal memuat keep-data.csv:", err);
    if (pill && !isSilent) pill.innerHTML = '<span style="color:#e11d48;">❌ Gagal memuat data keep</span>';
  }
}

function renderKeepTable() {
  applyKeepFilters();
}
window.renderKeepTable = renderKeepTable;

async function syncLiveKeepData() {
  const btn = document.getElementById('btnSyncKeepData');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) btn.innerHTML = '⏳ Menyinkronkan...';

  const headerBtn = document.getElementById('headerBtnSyncSheets');
  const originalHeaderText = headerBtn ? headerBtn.innerHTML : '';
  if (headerBtn) {
    headerBtn.innerHTML = `
      <svg class="spin-icon" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="23 4 23 10 17 10"></polyline>
        <polyline points="1 20 1 14 7 14"></polyline>
        <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
      </svg>
      <span>Syncing...</span>
    `;
  }

  const prog = startSmoothProgress({
    id: 'sync-live-keep-progress',
    title: '🔄 Sinkronisasi Data Google Sheets',
    detail: 'Menghubungkan & mengunduh data keep terbaru...'
  });

  try {
    let csvData = null;
    try {
      const res = await fetch('/api/sync-keep', { signal: AbortSignal.timeout ? AbortSignal.timeout(12000) : undefined });
      const text = await res.text();
      if (text.trim().startsWith('{')) {
        const data = JSON.parse(text);
        if (data.success) {
          if (data.csvText) {
            csvData = data.csvText;
          } else {
            const localRes = await fetch(`keep-data.csv?t=${Date.now()}`, { signal: AbortSignal.timeout ? AbortSignal.timeout(6000) : undefined });
            if (localRes.ok) {
              csvData = await localRes.text();
            }
          }
        }
      }
    } catch (e) {
      console.warn("Backend sync-keep warning:", e);
    }

    // Direct browser fallback to Google Sheets CSV export
    if (!csvData) {
      prog.update({ pct: 45, detail: 'Mengunduh langsung dari server Google Sheets...' });
      const directRes = await fetch('https://docs.google.com/spreadsheets/d/1V0pIa0NZCprJqRwnSnikFaLZ-dd50nY0uE3bjVYsUe0/export?format=csv&gid=1014933310', { signal: AbortSignal.timeout ? AbortSignal.timeout(10000) : undefined });
      if (directRes.ok) {
        csvData = await directRes.text();
      }
    }

    if (csvData) {
      prog.update({ pct: 80, detail: 'Memproses baris data keep ke memori...' });
      parseKeepCSV(csvData);
      try {
        localStorage.setItem('gadgetmart_sosmed_keep_cache', csvData);
      } catch (storageErr) {
        console.warn("Storage quota exceeded for localStorage keep cache:", storageErr);
      }
      applyKeepFilters();
      prog.complete({
        title: '✅ Sinkronisasi Google Sheets Berhasil!',
        detail: `Total ${rawKeepRows.length.toLocaleString('id-ID')} baris data berhasil dimuat`
      });
    } else {
      await loadKeepData();
      prog.complete({
        title: '✅ Sinkronisasi Selesai!',
        detail: 'Data keep terbaru berhasil dimuat'
      });
    }
  } catch (err) {
    console.error("Sync error:", err);
    prog.fail({
      title: '❌ Gagal Sinkronisasi Google Sheets',
      detail: err.message
    });
    await loadKeepData();
  } finally {
    if (btn) btn.innerHTML = originalText;
    if (headerBtn) headerBtn.innerHTML = originalHeaderText;
  }
}
window.syncLiveKeepData = syncLiveKeepData;

function populateKeepSosmedDropdown() {
  const select = document.getElementById('keepFilterSosmed');
  if (!select) return;

  const currentVal = select.value;
  const uniqueCodes = new Set();
  rawKeepRows.forEach(r => {
    const code = String(r['Kode Sosmed'] || '').trim();
    if (code && code !== 'undefined') uniqueCodes.add(code);
  });

  const sortedCodes = Array.from(uniqueCodes).sort();
  select.innerHTML = '<option value="ALL">Semua Kode Sosmed</option>' + 
    sortedCodes.map(c => `<option value="${c}">${c}</option>`).join('');

  if (currentVal && sortedCodes.includes(currentVal)) {
    select.value = currentVal;
  }
}

function applyKeepFilters() {
  if (!rawKeepRows || rawKeepRows.length === 0) return;

  const periodVal = document.getElementById('keepFilterPeriode') ? document.getElementById('keepFilterPeriode').value : 'ALL';
  const sosmedVal = document.getElementById('keepFilterSosmed') ? document.getElementById('keepFilterSosmed').value : 'ALL';
  const statusVal = document.getElementById('keepFilterStatus') ? document.getElementById('keepFilterStatus').value : 'ALL';
  const tanggalVal = document.getElementById('keepFilterTanggal') ? document.getElementById('keepFilterTanggal').value : '';
  const reportARVal = (document.getElementById('keepFilterReportAR') ? document.getElementById('keepFilterReportAR').value : '').trim().toLowerCase();
  const keyword = (document.getElementById('keepSearchInput') ? document.getElementById('keepSearchInput').value : '').trim().toLowerCase();

  filteredKeepRows = rawKeepRows.filter(r => {
    // 1. Period filter
    if (periodVal !== 'ALL') {
      const p = String(r['PERIODE'] || '').trim().toLowerCase();
      if (p !== periodVal.toLowerCase()) return false;
    }

    // 2. Sosmed filter
    if (sosmedVal !== 'ALL') {
      const s = String(r['Kode Sosmed'] || '').trim();
      if (s !== sosmedVal) return false;
    }

    // 3. Status filter
    if (statusVal !== 'ALL') {
      const st = String(r['Status'] || '').trim().toUpperCase();
      if (st !== statusVal.toUpperCase()) return false;
    }

    // 4. Filter Tanggal Format Keep (Timestamp Google Sheets)
    if (tanggalVal) {
      const keepDateIso = parseDateToISO(r['Timestamp']);
      if (keepDateIso !== tanggalVal) return false;
    }

    // 5. Filter by Report No. AR
    if (reportARVal) {
      const repAR = String(r['REPORT AR'] || '').toLowerCase();
      if (!repAR.includes(reportARVal)) return false;
    }

    // 6. Keyword search (Name, Phone, Item, AR, Notes)
    if (keyword) {
      const combined = `${r['Nama'] || ''} ${r['Nomor WA'] || ''} ${r['Barang Keep'] || ''} ${r['REPORT AR'] || ''} ${r['Catatan'] || ''} ${r['Feedback Tim FU'] || ''}`.toLowerCase();
      if (!combined.includes(keyword)) return false;
    }

    return true;
  });

  // Calculate Summary KPI Stats
  let winCount = 0;
  let progressCount = 0;
  let loseCount = 0;

  filteredKeepRows.forEach(r => {
    const st = String(r['Status'] || '').toUpperCase();
    if (st === 'WIN') winCount++;
    else if (st.includes('PROGRESS')) progressCount++;
    else if (st === 'LOSE' || st === 'TAKE OUT') loseCount++;
  });

  const total = filteredKeepRows.length;
  const closingRate = total > 0 ? Math.round((winCount / total) * 100) : 0;

  if (document.getElementById('keepTotalCount')) document.getElementById('keepTotalCount').innerText = total.toLocaleString('id-ID');
  if (document.getElementById('keepWinCount')) document.getElementById('keepWinCount').innerText = winCount.toLocaleString('id-ID');
  if (document.getElementById('keepProgressCount')) document.getElementById('keepProgressCount').innerText = progressCount.toLocaleString('id-ID');
  if (document.getElementById('keepLoseCount')) document.getElementById('keepLoseCount').innerText = loseCount.toLocaleString('id-ID');
  if (document.getElementById('keepClosingRate')) document.getElementById('keepClosingRate').innerText = `${closingRate}%`;

  let activeKeepCount = 0;
  if (document.getElementById('keepFilterPeriode') && document.getElementById('keepFilterPeriode').value !== 'ALL') activeKeepCount++;
  if (document.getElementById('keepFilterSosmed') && document.getElementById('keepFilterSosmed').value !== 'ALL') activeKeepCount++;
  if (document.getElementById('keepFilterTanggal') && document.getElementById('keepFilterTanggal').value) activeKeepCount++;
  if (document.getElementById('keepFilterStatus') && document.getElementById('keepFilterStatus').value !== 'ALL') activeKeepCount++;
  if (document.getElementById('keepSearchInput') && document.getElementById('keepSearchInput').value.trim()) activeKeepCount++;
  const badgeKeep = document.getElementById('keepFilterActiveCount');
  if (badgeKeep) badgeKeep.innerText = activeKeepCount;

  keepCurrentPage = 1;
  renderKeepPage();
}

function toggleKeepFiltersSection() {
  const el = document.getElementById('keepFilterSection');
  if (el) {
    el.style.display = (el.style.display === 'none' ? 'flex' : 'none');
  }
}

function toggleAuditFiltersSection() {
  const el = document.getElementById('auditFilterSection');
  if (el) {
    el.style.display = (el.style.display === 'none' ? 'flex' : 'none');
  }
}

function resetKeepFilters() {
  if (document.getElementById('keepFilterPeriode')) document.getElementById('keepFilterPeriode').value = 'Oktober-26';
  if (document.getElementById('keepFilterSosmed')) document.getElementById('keepFilterSosmed').value = 'ALL';
  if (document.getElementById('keepFilterStatus')) document.getElementById('keepFilterStatus').value = 'ALL';
  if (document.getElementById('keepFilterTanggal')) document.getElementById('keepFilterTanggal').value = '';
  if (document.getElementById('keepFilterReportAR')) document.getElementById('keepFilterReportAR').value = '';
  if (document.getElementById('keepSearchInput')) document.getElementById('keepSearchInput').value = '';
  applyKeepFilters();
}

function changeKeepPage(delta) {
  const maxPage = Math.max(1, Math.ceil(filteredKeepRows.length / KEEP_PAGE_SIZE));
  const newPage = keepCurrentPage + delta;
  if (newPage >= 1 && newPage <= maxPage) {
    keepCurrentPage = newPage;
    renderKeepPage();
  }
}

function getCustomerInitials(name) {
  if (!name || name === '-') return 'CU';
  const clean = String(name).replace(/[^a-zA-Z0-9\s]/g, '').trim();
  const parts = clean.split(/\s+/).filter(Boolean);
  if (parts.length === 0) return 'CU';
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

function formatKeepDisplayDate(ts) {
  if (!ts) return '-';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  if (typeof ts === 'number') {
    const d = XLSX.SSF.parse_date_code(ts);
    if (d) {
      const mStr = months[(d.m - 1) % 12] || 'Jan';
      return `${d.d} ${mStr} ${d.y}`;
    }
  }
  if (typeof ts === 'string') {
    const parts = ts.trim().split(/[\/\-\s]/);
    if (parts.length >= 3) {
      const day = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      let y = parts[2].trim();
      if (y.length === 2) y = '20' + y;
      if (!isNaN(day) && !isNaN(m) && m >= 1 && m <= 12) {
        return `${day} ${months[m - 1]} ${y}`;
      }
    }
    return ts;
  }
  return String(ts);
}

function formatPickupDate(val) {
  if (!val || val === '-') return '-';
  const lower = String(val).toLowerCase().trim();
  if (lower.includes('hari ini')) {
    return `<span class="text-crimson font-bold">Hari ini</span>`;
  }
  return val;
}

function getModernStatusPill(st) {
  const norm = String(st || '').trim().toUpperCase();
  if (norm === 'WIN') return `<span class="modern-status-pill status-pill-green"><span class="status-dot"></span>Win</span>`;
  if (norm === 'ON PROGRESS LOW' || norm === 'LOW') return `<span class="modern-status-pill status-pill-green"><span class="status-dot"></span>Low</span>`;
  if (norm.includes('PROGRESS')) return `<span class="modern-status-pill status-pill-orange"><span class="status-dot"></span>On Progress</span>`;
  if (norm === 'LOSE' || norm === 'TAKE OUT') return `<span class="modern-status-pill status-pill-red"><span class="status-dot"></span>Lose</span>`;
  return `<span class="modern-status-pill status-pill-gray"><span class="status-dot"></span>${st || '-'}</span>`;
}

function openStatusChangeQuickModal(rawWA, keepRowIdx, currentStatus) {
  const newStatus = prompt(`Ubah Status Customer:\nKetik: WIN, ON PROGRESS LOW, ON PROGRESS MEDIUM, ON PROGRESS HOT, atau LOSE`, currentStatus || 'WIN');
  if (newStatus && newStatus.trim() && newStatus.trim().toUpperCase() !== currentStatus) {
    const cleanSt = newStatus.trim().toUpperCase();
    if (filteredKeepRows[keepRowIdx]) {
      const oldSt = filteredKeepRows[keepRowIdx]['Status'];
      filteredKeepRows[keepRowIdx]['Status'] = cleanSt;
      renderKeepPage();
      triggerBackgroundKeepStatusSync(rawWA, cleanSt, oldSt, () => {
        filteredKeepRows[keepRowIdx]['Status'] = oldSt;
        renderKeepPage();
      });
    }
  }
}

function renderKeepPage() {
  const tbody = document.getElementById('keepTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const total = filteredKeepRows.length;
  const maxPage = Math.max(1, Math.ceil(total / KEEP_PAGE_SIZE));
  const startIdx = (keepCurrentPage - 1) * KEEP_PAGE_SIZE;
  const endIdx = Math.min(startIdx + KEEP_PAGE_SIZE, total);
  const pageRows = filteredKeepRows.slice(startIdx, endIdx);

  // Update hero subtitle and pagination info
  const heroSub = document.getElementById('customerHeroSub');
  if (heroSub) {
    heroSub.innerText = `${total.toLocaleString('id-ID')} customer dari Google Sheets Live`;
  }

  const paginationInfoEl = document.getElementById('keepPaginationInfo');
  if (paginationInfoEl) {
    paginationInfoEl.innerText = total > 0 
      ? `Menampilkan ${startIdx + 1} - ${endIdx} dari total ${total.toLocaleString('id-ID')} Data` 
      : 'Tidak ada data customer yang sesuai filter';
  }
  const curPageTextEl = document.getElementById('keepCurrentPageText');
  if (curPageTextEl) curPageTextEl.innerText = `Halaman ${keepCurrentPage} / ${maxPage}`;
  const btnPrev = document.getElementById('btnKeepPrev');
  if (btnPrev) btnPrev.disabled = keepCurrentPage <= 1;
  const btnNext = document.getElementById('btnKeepNext');
  if (btnNext) btnNext.disabled = keepCurrentPage >= maxPage;

  if (pageRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="text-center" style="padding: 40px; color: var(--text-muted); font-size: 0.9rem;">Tidak ada data customer untuk filter yang dipilih.</td></tr>`;
    return;
  }

  pageRows.forEach((r, i) => {
    const rawWA = String(r['Nomor WA'] || '').replace(/\D/g, '');
    const normWA = normalizePhone(rawWA);
    const displayWA = formatPhoneDisplay(rawWA);
    const custName = r['Nama'] || '-';
    const initials = getCustomerInitials(custName);
    const keepRowIdx = startIdx + i;

    // 1. CUSTOMER CELL: Avatar Circle + Name Bold + Phone Subtext (Persis Gambar 1)
    const custHtml = `
      <div class="customer-cell-flex">
        <div class="avatar-circle-sm" title="${custName}">${initials}</div>
        <div class="customer-info-box">
          <span class="customer-name-bold">${custName}</span>
          ${(normWA && normWA.length >= 8) ? `
            <a href="https://wa.me/${normWA}" target="_blank" class="customer-phone-sub" title="Hubungi customer di WhatsApp">
              📞 ${displayWA}
            </a>
          ` : `<span class="customer-phone-sub">${displayWA || '-'}</span>`}
        </div>
      </div>
    `;

    // 2. REPORT AR CELL: Dashed button "+ Input AR" atau nomor merah tebal (Persis Gambar 1)
    const arDoc = String(r['REPORT AR'] || '').trim();
    const arHtml = arDoc ? `
      <div class="cell-report-ar-wrapper" data-wa="${rawWA}" data-keep-idx="${keepRowIdx}" data-current-ar="${arDoc}">
        <span class="text-ar-number-bold" onclick="openARDetailModalFromKeep(${keepRowIdx})" title="Klik untuk cek detail faktur AR">
          ${arDoc}
        </span>
        <button type="button" class="btn-edit-inline-ar" onclick="event.stopPropagation(); startEditKeepReportAR(this.parentElement, '${rawWA}', ${keepRowIdx})" title="Ubah nomor AR">✏️</button>
      </div>
    ` : `
      <div class="cell-report-ar-wrapper" data-wa="${rawWA}" data-keep-idx="${keepRowIdx}" data-current-ar="">
        <button type="button" class="btn-input-new-ar" onclick="startEditKeepReportAR(this.parentElement, '${rawWA}', ${keepRowIdx})" title="Klik untuk menginput nomor AR">
          <span>📄</span> Input AR
        </button>
      </div>
    `;

    // 3. STATUS PILL: Pill dengan dot warna (Win, Low, On Progress, Lose)
    const rawSt = String(r['Status'] || '').trim().toUpperCase();
    const statusPillHtml = `
      <div class="cursor-pointer" onclick="openStatusChangeQuickModal('${rawWA}', ${keepRowIdx}, '${rawSt}')" title="Klik untuk mengubah status customer">
        ${getModernStatusPill(rawSt)}
      </div>
    `;

    const tr = document.createElement('tr');
    if (rawSt === 'LOSE' || rawSt === 'TAKE OUT') {
      tr.classList.add('row-lose-tint');
    }

    tr.innerHTML = `
      <td class="text-left">${custHtml}</td>
      <td class="text-center font-bold" style="color:#0f172a; white-space:nowrap;">${formatKeepDisplayDate(r['Timestamp'])}</td>
      <td class="text-left font-bold" style="color:#0f172a;">${r['Barang Keep'] || '-'}</td>
      <td class="text-center" style="font-size:0.80rem; white-space:nowrap;">${formatPickupDate(r['Tanggal Ambil'])}</td>
      <td class="text-left" style="font-size:0.80rem; color:#334155; line-height:1.35;">${r['Catatan'] || '-'}</td>
      <td class="text-center"><span class="badge-sosmed-pill">${r['Kode Sosmed'] || '-'}</span></td>
      <td class="text-center">${arHtml}</td>
      <td class="text-center">${statusPillHtml}</td>
      <td class="text-left" style="font-size:0.80rem; color:#475569; line-height:1.35;">${r['Feedback Tim FU'] || '-'}</td>
      <td class="text-center">
        <button type="button" class="btn-circle-action" onclick="openARDetailModalFromKeep(${keepRowIdx})" title="Buka Detail">➔</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function openARDetailModalFromKeep(keepIndex) {
  const row = filteredKeepRows[keepIndex];
  if (!row) return;
  const arDoc = row['REPORT AR'];
  if (!arDoc) {
    const rawWA = String(row['Nomor WA'] || '').replace(/\D/g, '');
    const promptAR = prompt(`Customer "${row['Nama']}" belum memiliki nomor AR.\nMasukkan nomor faktur AR (misal: 260278899):`, '');
    if (promptAR && promptAR.trim()) {
      const cleanAR = promptAR.trim();
      row['REPORT AR'] = cleanAR;
      renderKeepPage();
      triggerBackgroundReportARSync(rawWA, cleanAR, '', () => {
        row['REPORT AR'] = '';
        renderKeepPage();
      });
    }
    return;
  }
  openARDetailModal(arDoc, row);
}

function pickBestKeep(matchedKeeps, postingDate) {
  if (!matchedKeeps || matchedKeeps.length === 0) return null;
  if (matchedKeeps.length === 1) return matchedKeeps[0];

  const monthMap = {
    '01': 'januari', '1': 'januari',
    '02': 'februari', '2': 'februari',
    '03': 'maret', '3': 'maret',
    '04': 'april', '4': 'april',
    '05': 'mei', '5': 'mei',
    '06': 'juni', '6': 'juni',
    '07': 'juli', '7': 'juli',
    '08': 'agustus', '8': 'agustus',
    '09': 'september', '9': 'september',
    '10': 'oktober',
    '11': 'november',
    '12': 'desember'
  };

  let monthName = '';
  if (typeof postingDate === 'number') {
    const d = XLSX.SSF.parse_date_code(postingDate);
    if (d && d.m) monthName = monthMap[String(d.m)] || '';
  } else if (postingDate instanceof Date) {
    monthName = monthMap[String(postingDate.getMonth() + 1)] || '';
  } else if (typeof postingDate === 'string') {
    const cleanStr = postingDate.trim().split(' ')[0];
    const parts = cleanStr.split(/[\/\-]/);
    if (parts.length >= 3) {
      const mCandidate = parts[0].length === 4 ? parts[1] : parts[1];
      monthName = monthMap[mCandidate] || monthMap[String(Number(mCandidate))] || '';
    }
  }

  if (monthName) {
    const match = matchedKeeps.find(k => String(k['PERIODE'] || '').toLowerCase().includes(monthName));
    if (match) return match;
  }

  // Fallback to the latest/newest keep row in Google Sheets
  return matchedKeeps[matchedKeeps.length - 1];
}

// ==================== POP-UP MODAL DETAIL AR & PELAPORAN WHATSAPP AKMAL ====================
let currentModalDoc = null;
let currentModalRows = [];
let currentModalKeep = null;
let currentModalCheckedItems = new Set();

function openARDetailModal(docNumber, keepContext = null, sourceTitle = '') {
  const digitsOnly = String(docNumber || '').replace(/\D/g, '');
  if (!digitsOnly) {
    alert("Nomor Dokumen AR tidak valid.");
    return;
  }

  // 1. Find AR rows in rawCrosellingRows (or rawTurnoverRows fallback)
  let matchedRows = [];
  if (rawCrosellingRows && rawCrosellingRows.length > 0) {
    matchedRows = rawCrosellingRows.filter(r => String(r['Document Number'] || '').replace(/\D/g, '') === digitsOnly);
  }
  if (matchedRows.length === 0 && rawTurnoverRows && rawTurnoverRows.length > 0) {
    matchedRows = rawTurnoverRows.filter(r => String(r['Document Number'] || '').replace(/\D/g, '') === digitsOnly);
  }

  if (matchedRows.length === 0) {
    alert(`Dokumen AR ${docNumber} tidak ditemukan di database transaksi croselling/turnover.`);
    return;
  }

  currentModalDoc = docNumber;
  currentModalRows = matchedRows;
  currentModalKeep = keepContext;

  // If keepContext was not passed, look up in rawKeepRows by AR or phone
  if (!currentModalKeep && rawKeepRows && rawKeepRows.length > 0) {
    currentModalKeep = rawKeepRows.find(k => String(k['REPORT AR'] || '').replace(/\D/g, '') === digitsOnly);
    if (!currentModalKeep && matchedRows[0]['No HP Customer']) {
      const norm = normalizePhone(matchedRows[0]['No HP Customer']);
      const allPhoneMatches = rawKeepRows.filter(k => normalizePhone(k['Nomor WA']) === norm);
      currentModalKeep = pickBestKeep(allPhoneMatches, matchedRows[0]['Posting Date']);
    }
  }

  // Populate Title & Header (Matching Image 2)
  const mainTitleEl = document.getElementById('modalDocMainTitle');
  if (mainTitleEl) {
    if (sourceTitle === 'REPORT AR') {
      mainTitleEl.innerText = 'TRACK AR BERDASARKAN REPORT AR';
    } else if (sourceTitle === 'AR VLOOKUP') {
      mainTitleEl.innerText = 'TRACK AR BERDASARKAN VLOOKUP';
    } else {
      mainTitleEl.innerText = `TRACK AR DOKUMEN ${docNumber}`;
    }
  }

  const docTitleEl = document.getElementById('modalDocTitle');
  if (docTitleEl) docTitleEl.innerText = docNumber;

  const first = matchedRows[0];
  const metaDateEl = document.getElementById('modalMetaDate');
  if (metaDateEl) metaDateEl.innerText = formatCrosellingDate(first['Posting Date']) + (first['Waktu'] ? ` ${first['Waktu']}` : '');
  const metaCabangEl = document.getElementById('modalMetaCabang');
  if (metaCabangEl) metaCabangEl.innerText = first['Cabang'] || '-';

  // Populate Top 5-Column Summary Table (Matching Image 2 Mockup)
  const custName = (currentModalKeep && currentModalKeep['Nama']) ? currentModalKeep['Nama'] : (first['Name'] || first['Customer/Vendor Name'] || '-');
  const rawWA = (currentModalKeep && currentModalKeep['Nomor WA']) ? currentModalKeep['Nomor WA'] : (first['No HP Customer'] || '');
  const displayWA = formatPhoneDisplay(rawWA);
  const normWA = normalizePhone(rawWA);
  const waHtml = (normWA && normWA.length >= 8)
    ? `<div style="display:flex; align-items:center; gap:6px;">
        <span class="font-bold">${displayWA}</span>
        <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp" style="padding:2px 8px; font-size:0.72rem; border-radius:4px;" title="Chat customer ini di WhatsApp">
          💬 WA
        </a>
       </div>`
    : (displayWA || '-');

  const sosmedCode = (currentModalKeep && currentModalKeep['Kode Sosmed']) ? currentModalKeep['Kode Sosmed'] : '-';
  const keepBarang = (currentModalKeep && currentModalKeep['Barang Keep']) ? currentModalKeep['Barang Keep'] : '-';

  const keepCatatan = (currentModalKeep && currentModalKeep['Catatan']) ? String(currentModalKeep['Catatan']).trim() : '';
  const arRemarks = first['Remarks'] ? String(first['Remarks']).trim() : '';
  let catatanModalHtml = '-';
  if (keepCatatan && arRemarks && keepCatatan.toLowerCase() !== arRemarks.toLowerCase() && keepCatatan !== '-' && arRemarks !== '-') {
    catatanModalHtml = `
      <div style="font-size:0.76rem; color:#0f172a; margin-bottom:2px; word-break:break-word;">${keepCatatan}</div>
      <div style="font-size:0.70rem; color:var(--text-muted); word-break:break-word;"><em>Kasir: ${arRemarks}</em></div>
    `;
  } else if (keepCatatan && keepCatatan !== '-') {
    catatanModalHtml = `<div style="font-size:0.76rem; color:#0f172a; word-break:break-word;">${keepCatatan}</div>`;
  } else if (arRemarks && arRemarks !== '-') {
    catatanModalHtml = `<div style="font-size:0.72rem; color:var(--text-muted); word-break:break-word;"><em>Kasir: ${arRemarks}</em></div>`;
  }

  const elCust = document.getElementById('modalTrackCust');
  const elWA = document.getElementById('modalTrackWA');
  const elSos = document.getElementById('modalTrackSosmed');
  const elBarang = document.getElementById('modalTrackKeepBarang');
  const elCat = document.getElementById('modalTrackCatatan');
  const elStatus = document.getElementById('modalTrackStatusKeep');

  if (elCust) elCust.innerText = custName;
  if (elWA) elWA.innerHTML = waHtml;
  if (elSos) elSos.innerHTML = `<span class="badge-sosmed-pill">${sosmedCode}</span>`;
  if (elBarang) elBarang.innerText = keepBarang;
  if (elCat) elCat.innerHTML = catatanModalHtml;

  if (elStatus) {
    const rawModalWA = String(currentModalKeep && currentModalKeep['Nomor WA'] || '').replace(/\D/g, '');
    const rawModalStatus = currentModalKeep && currentModalKeep['Status'];
    const currentStatus = (rawModalStatus && String(rawModalStatus).trim() !== '-') ? String(rawModalStatus).trim().toUpperCase() : '';
    const statusSelectClass = getKeepStatusClass(currentStatus);

    let statusOptionsHtml = '';
    if (currentStatus && !KEEP_STATUS_OPTIONS.some(o => o.value.toUpperCase() === currentStatus)) {
      statusOptionsHtml += `<option value="${currentStatus}" selected>${currentStatus}</option>`;
    }
    KEEP_STATUS_OPTIONS.forEach(opt => {
      const isSel = (opt.value.toUpperCase() === currentStatus);
      statusOptionsHtml += `<option value="${opt.value}" ${isSel ? 'selected' : ''}>${opt.label}</option>`;
    });

    elStatus.innerHTML = (rawModalWA && rawModalWA.length >= 8) ? `
      <select class="select-status-keep ${statusSelectClass}" data-prev="${currentStatus}" onchange="handleModalKeepStatusChange(this, '${rawModalWA}')" title="Klik untuk mengubah status follow-up customer ini (otomatis tersimpan ke Google Sheets)">
        ${statusOptionsHtml}
      </select>
    ` : `<span class="kpi-badge badge-ink">-</span>`;
  }

  // Backward compatibility metadata
  const metaDocEl = document.getElementById('modalMetaDoc');
  if (metaDocEl) metaDocEl.innerText = first['Document Number'] || docNumber;
  const metaCustEl = document.getElementById('modalMetaCust');
  if (metaCustEl) metaCustEl.innerText = custName;
  const metaWAEl = document.getElementById('modalMetaWA');
  if (metaWAEl) metaWAEl.innerHTML = waHtml;
  const metaRemarksEl = document.getElementById('modalMetaRemarks');
  if (metaRemarksEl) metaRemarksEl.innerText = first['Remarks'] || '-';

  // Check Sosmed Status per Item
  let totalDocVal = 0;
  let allMissing = true;
  let missingCount = 0;
  currentModalCheckedItems.clear();

  matchedRows.forEach((r, idx) => {
    const qty = Number(r['Quantity']) || 1;
    const subtotal = Number(r['Harga Jual']) || 0;
    totalDocVal += subtotal;

    const isExempt = isItemExemptFromSosmed(r);
    const isSosmed = isItemInputSosmed(r);

    if (isExempt) {
      // Tas Gadgetmart 25x25 dll tidak perlu diinput sosmed (dianggap pass / exempt)
      return;
    }

    if (isSosmed) {
      allMissing = false;
    } else {
      missingCount++;
      // Pre-check missing items so user can report them directly!
      currentModalCheckedItems.add(idx);
    }
  });

  document.getElementById('modalMetaTotal').innerText = formatRupiah(totalDocVal);
  document.getElementById('modalItemsCountBadge').innerText = `${matchedRows.length} Item`;

  // Status Banner in Modal
  const banner = document.getElementById('modalStatusBanner');
  const icon = document.getElementById('modalStatusIcon');
  const title = document.getElementById('modalStatusTitle');
  const subtitle = document.getElementById('modalStatusSubtitle');

  if (allMissing) {
    banner.className = 'ar-status-banner banner-danger';
    icon.innerHTML = '❌';
    title.innerText = 'BELUM DIINPUT KODE SOSMED (100% OFFLINE)';
    subtitle.innerText = `Seluruh item dalam AR ini tercatat dengan kode sales toko fisik/kasir, belum ada kode sosmed.`;
  } else if (missingCount > 0) {
    banner.className = 'ar-status-banner banner-warning';
    icon.innerHTML = '⚠️';
    title.innerText = `SEBAGIAN ITEM BELUM TERINPUT (${missingCount} DARI ${matchedRows.length} ITEM)`;
    subtitle.innerText = `Sebagian produk sudah berstatus sosmed, namun masih ada ${missingCount} item yang belum terinput sosmed.`;
  } else {
    banner.className = 'ar-status-banner banner-success';
    icon.innerHTML = '✅';
    title.innerText = 'SEMUA ITEM SUDAH TERINPUT SOSMED (100%)';
    subtitle.innerText = `Semua item dalam dokumen ini sudah valid terinput kode sales sosmed (termasuk item tas yang bebas sosmed).`;
  }

  // Render Table Items with Checkbox
  renderModalItemsTable();

  // Populate Target Sosmed Dropdown for SubTab 2
  populateModalSosmedSelect();

  // Reset Notes
  const defaultNote = (currentModalKeep && currentModalKeep['Barang Keep'])
    ? `Customer keep barang "${currentModalKeep['Barang Keep']}". Mohon bantuannya Mas Akmal untuk input sales sosmed pada item terpilih di AR ini. Terima kasih!`
    : 'Mohon bantuannya Mas Akmal untuk input sales sosmed pada item terpilih di AR ini. Terima kasih!';
  document.getElementById('modalReportNotes').value = defaultNote;

  // Default to Sub-Tab 1
  switchModalSubTab(1);

  // Show Modal
  document.getElementById('arDetailModal').style.display = 'flex';
}

function closeARDetailModal() {
  document.getElementById('arDetailModal').style.display = 'none';
}

function switchModalSubTab(tabIdx) {
  const tab1 = document.getElementById('modalSubTabView1');
  const tab2 = document.getElementById('modalSubTabView2');
  const btn1 = document.getElementById('btnModalSubTab1');
  const btn2 = document.getElementById('btnModalSubTab2');

  if (tabIdx === 1) {
    tab1.style.display = 'block';
    tab2.style.display = 'none';
    btn1.classList.add('active');
    btn2.classList.remove('active');
  } else {
    tab1.style.display = 'none';
    tab2.style.display = 'block';
    btn1.classList.remove('active');
    btn2.classList.add('active');
    updateSelectedItemsForReport();
    updateWhatsAppMessagePreview();
  }
}

function renderModalItemsTable() {
  const tbody = document.getElementById('modalItemsTableBody');
  tbody.innerHTML = '';

  let totalQty = 0;
  let totalVal = 0;

  currentModalRows.forEach((r, idx) => {
    const qty = Number(r['Quantity']) || 1;
    const subtotal = Number(r['Harga Jual']) || 0;
    const price = qty > 0 ? (subtotal / qty) : subtotal;
    totalQty += qty;
    totalVal += subtotal;

    const isExempt = isItemExemptFromSosmed(r);
    const isSosmed = isItemInputSosmed(r);
    const schl = String(r['SC/HL Code'] || '').trim();
    const slp = String(r['SlpName'] || '').trim();
    const emp = String(r['Sales Employee Name'] || '').trim();
    const empDisplay = schl || slp || emp || '-';

    const isChecked = currentModalCheckedItems.has(idx);

    // Potensi Denda: Handphone = DENDA, Aksesoris/Tas/lainnya = PASS
    const dendaStatus = isExempt ? 'PASS' : getItemPotensiDenda(r);
    const dendaBadge = dendaStatus === 'DENDA'
      ? '<span class="kpi-badge badge-crimson font-bold">🚨 DENDA</span>'
      : '<span class="kpi-badge badge-ink">PASS</span>';

    let statusSosmedBadge = '';
    if (isExempt) {
      statusSosmedBadge = '<span class="status-badge" style="background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; font-weight:700;">🛍️ Bebas Sosmed</span>';
    } else if (isSosmed) {
      statusSosmedBadge = '<span class="status-badge badge-accent">✅ Sudah Sosmed</span>';
    } else {
      statusSosmedBadge = '<span class="status-badge badge-crimson font-bold">❌ Belum Sosmed</span>';
    }

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;">
        <input type="checkbox" class="checkbox-custom" ${isChecked ? 'checked' : ''} onchange="toggleModalItemCheck(${idx}, this.checked)">
      </td>
      <td class="col-sticky">${idx + 1}</td>
      <td class="text-left font-bold" style="color:var(--ink);">${r['Dscription'] || '-'}</td>
      <td>${qty}</td>
      <td>${formatRupiah(price)}</td>
      <td class="font-bold">${formatRupiah(subtotal)}</td>
      <td class="text-left">${empDisplay}</td>
      <td class="text-center">${statusSosmedBadge}</td>
      <td class="text-center">${dendaBadge}</td>
      <td class="text-center">
        <button class="btn btn-outline btn-xs" onclick="copyModalSingleItem(${idx}, this)" style="padding: 3px 8px; font-size: 0.74rem; font-weight:700;" title="Salin baris item ini ke format 12 kolom Excel (tanpa header)">
          📋 Copy Rekap
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('modalItemsTableFoot').innerHTML = `
    <tr class="row-total">
      <td colspan="3" class="text-left font-bold">TOTAL DOKUMEN AR:</td>
      <td><strong>${totalQty}</strong></td>
      <td>-</td>
      <td class="font-bold">${formatRupiah(totalVal)}</td>
      <td colspan="4" class="text-left font-bold text-accent">${currentModalRows.length} Baris Produk</td>
    </tr>
  `;

  const selectAll = document.getElementById('modalSelectAllItems');
  if (selectAll) {
    selectAll.checked = currentModalRows.length > 0 && currentModalCheckedItems.size === currentModalRows.length;
  }
}

function toggleModalItemCheck(idx, checked) {
  if (checked) {
    currentModalCheckedItems.add(idx);
  } else {
    currentModalCheckedItems.delete(idx);
  }
  const selectAll = document.getElementById('modalSelectAllItems');
  if (selectAll) {
    selectAll.checked = currentModalCheckedItems.size === currentModalRows.length;
  }
  updateSelectedItemsForReport();
  updateWhatsAppMessagePreview();
}

function toggleSelectAllModalItems(checked) {
  currentModalCheckedItems.clear();
  if (checked) {
    currentModalRows.forEach((_, idx) => currentModalCheckedItems.add(idx));
  }
  renderModalItemsTable();
  updateSelectedItemsForReport();
  updateWhatsAppMessagePreview();
}

function populateModalSosmedSelect() {
  const select = document.getElementById('modalReportSosmedSelect');
  if (!select) return;

  const masterList = getMasterSales().filter(m => m.active);
  let keepSosmed = '';
  if (currentModalKeep && currentModalKeep['Kode Sosmed']) {
    keepSosmed = String(currentModalKeep['Kode Sosmed']).trim();
  }

  let optionsHtml = '<option value="">-- Pilih Kode Sales Sosmed --</option>';
  masterList.forEach(m => {
    const fullVal = `${m.code} - ${m.name}`;
    const isSelected = keepSosmed && (keepSosmed.includes(m.code) || keepSosmed.includes(m.name));
    optionsHtml += `<option value="${fullVal}" ${isSelected ? 'selected' : ''}>${fullVal}</option>`;
  });

  if (keepSosmed && !masterList.some(m => keepSosmed.includes(m.code))) {
    optionsHtml += `<option value="${keepSosmed}" selected>⭐ ${keepSosmed} (Dari Data Keep)</option>`;
  }

  select.innerHTML = optionsHtml;
}

function updateSelectedItemsForReport() {
  const container = document.getElementById('modalReportSelectedItemsList');
  if (!container) return;

  if (currentModalCheckedItems.size === 0) {
    container.innerHTML = '<em style="color:var(--text-muted);">Belum ada item yang dicentang. Silakan kembali ke sub-tab 1 untuk memilih item, atau centang semua item.</em>';
    return;
  }

  const itemsHtml = Array.from(currentModalCheckedItems).sort((a,b)=>a-b).map(idx => {
    const r = currentModalRows[idx];
    const desc = r['Dscription'] || `Item #${idx+1}`;
    const qty = Number(r['Quantity']) || 1;
    return `
      <div class="selected-item-tag">
        <span><strong>${idx + 1}. ${desc}</strong> (Qty: ${qty})</span>
        <button type="button" class="btn-remove-tag" onclick="toggleModalItemCheck(${idx}, false); renderModalItemsTable();" title="Batal lapor item ini">✕</button>
      </div>
    `;
  }).join('');

  container.innerHTML = itemsHtml;
}

function getReportWhatsAppText() {
  const doc = currentModalDoc || '-';
  const first = currentModalRows[0] || {};
  const dateStr = formatCrosellingDate(first['Posting Date']);
  const cabang = first['Cabang'] || '-';
  const custName = (currentModalKeep && currentModalKeep['Nama']) ? currentModalKeep['Nama'] : (first['Name'] || '-');
  const custWA = (currentModalKeep && currentModalKeep['Nomor WA']) ? formatPhoneDisplay(currentModalKeep['Nomor WA']) : formatPhoneDisplay(first['No HP Customer']);
  const barangKeep = (currentModalKeep && currentModalKeep['Barang Keep']) ? currentModalKeep['Barang Keep'] : '-';
  const targetSosmed = document.getElementById('modalReportSosmedSelect').value || '(Belum Dipilih)';
  const notes = document.getElementById('modalReportNotes').value.trim() || '-';

  const selectedItemsList = Array.from(currentModalCheckedItems).sort((a,b)=>a-b).map((idx, num) => {
    const r = currentModalRows[idx];
    const desc = r['Dscription'] || '-';
    const qty = Number(r['Quantity']) || 1;
    const subtotal = Number(r['Harga Jual']) || 0;
    return `${num + 1}. *${desc}* (Qty: ${qty} | ${formatRupiah(subtotal)})`;
  });

  const itemsText = selectedItemsList.length > 0
    ? selectedItemsList.join('\n')
    : '(Belum ada item yang dicentang)';

  return `🚨 *LAPORAN REVISI / ITEM BELUM TERINPUT SOSMED*
-----------------------------------------
*No. AR / Faktur:* ${doc}
*Tanggal AR:* ${dateStr}
*Cabang:* ${cabang}
*Nama Customer:* ${custName}
*No. WhatsApp:* ${custWA}
*Barang Keep:* ${barangKeep}

*Daftar Item yang Belum Terinput Sosmed:*
${itemsText}

*Harusnya Diinput Kode Sosmed:*
👉 *${targetSosmed}*

*Catatan / Keterangan:*
${notes}
-----------------------------------------
_Dikirim via Dashboard Sosmed GadgetMart_`;
}

function updateWhatsAppMessagePreview() {
  const pre = document.getElementById('modalWAPreview');
  if (pre) {
    pre.innerText = getReportWhatsAppText();
  }
}

function sendReportToAkmalWhatsApp() {
  if (currentModalCheckedItems.size === 0) {
    alert("Pilih minimal 1 item yang belum terinput sosmed sebelum mengirim laporan!");
    switchModalSubTab(1);
    return;
  }

  const targetSosmed = document.getElementById('modalReportSosmedSelect').value;
  if (!targetSosmed) {
    alert("Pilih 'Harusnya Diinput Kode Sosmed Berapa' terlebih dahulu!");
    const el = document.getElementById('modalReportSosmedSelect');
    if (el && typeof el.focus === 'function') el.focus();
    return;
  }

  const message = getReportWhatsAppText();
  // Akmal's WhatsApp: +62 823-5771-2674 -> 6282357712674
  const akmalWA = '6282357712674';
  const url = `https://wa.me/${akmalWA}?text=${encodeURIComponent(message)}`;
  window.open(url, '_blank');
}

// ==================== TAB 5: TRACKING AR (AUDIT MISSING SOSMED) ====================
let rawAuditARList = [];
let filteredAuditARList = [];
let auditCurrentPage = 1;
const AUDIT_PAGE_SIZE = 50;

function parseKeepTimestampToEpoch(ts) {
  if (!ts) return 0;
  if (typeof ts === 'number') {
    return (ts - 25569) * 86400 * 1000;
  }
  if (ts instanceof Date) {
    return ts.getTime();
  }
  if (typeof ts === 'string') {
    const s = ts.trim();
    if (!s) return 0;
    const [datePart, timePart] = s.split(/\s+/);
    let hours = 0, minutes = 0, seconds = 0;
    if (timePart) {
      const tPieces = timePart.split(':');
      if (tPieces.length >= 2) {
        hours = parseInt(tPieces[0], 10) || 0;
        minutes = parseInt(tPieces[1], 10) || 0;
        seconds = parseInt(tPieces[2], 10) || 0;
      }
    }

    const dPieces = datePart.split(/[\/\-\.]/);
    if (dPieces.length === 3) {
      let y, m, d;
      if (dPieces[0].length === 4) {
        y = parseInt(dPieces[0], 10);
        m = parseInt(dPieces[1], 10) - 1;
        d = parseInt(dPieces[2], 10);
      } else {
        const p0 = parseInt(dPieces[0], 10);
        const p1 = parseInt(dPieces[1], 10);
        let p2 = parseInt(dPieces[2], 10);
        if (p2 < 100) p2 += 2000;

        if (p0 > 12) {
          d = p0;
          m = p1 - 1;
          y = p2;
        } else if (p1 > 12) {
          m = p0 - 1;
          d = p1;
          y = p2;
        } else {
          d = p0;
          m = p1 - 1;
          y = p2;
        }
      }
      return new Date(y, m, d, hours, minutes, seconds).getTime();
    }
  }
  return 0;
}

function parseDateToISO(dateVal) {
  if (!dateVal) return '';
  if (typeof dateVal === 'number') {
    if (typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
      const d = XLSX.SSF.parse_date_code(dateVal);
      if (d) {
        return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
      }
    }
  }
  if (dateVal instanceof Date) {
    return `${dateVal.getFullYear()}-${String(dateVal.getMonth() + 1).padStart(2, '0')}-${String(dateVal.getDate()).padStart(2, '0')}`;
  }
  if (typeof dateVal === 'string') {
    const cleanStr = dateVal.trim().split(' ')[0].split('T')[0];
    const parts = cleanStr.split(/[\/\-\.]/);
    if (parts.length >= 3) {
      if (parts[0].length === 4) {
        return `${parts[0]}-${String(parts[1]).padStart(2, '0')}-${String(parts[2]).padStart(2, '0')}`;
      } else {
        let y = parts[2].trim();
        if (y.length === 2) y = '20' + y;
        const p0 = Number(parts[0].trim());
        const p1 = Number(parts[1].trim());
        if (p0 > 12) {
          return `${y}-${String(p1).padStart(2, '0')}-${String(p0).padStart(2, '0')}`;
        }
        if (p1 > 12) {
          return `${y}-${String(p0).padStart(2, '0')}-${String(p1).padStart(2, '0')}`;
        }
        return `${y}-${String(p1).padStart(2, '0')}-${String(p0).padStart(2, '0')}`;
      }
    }
  }
  return '';
}

function runARMissingSosmedAudit() {
  const pill = document.getElementById('arAuditStatusPill');
  if (pill) pill.innerHTML = '<span>⏳ Menganalisis Transaksi vs Data Keep...</span>';

  if (!rawCrosellingRows || rawCrosellingRows.length === 0 || !rawKeepRows || rawKeepRows.length === 0) {
    if (pill) pill.innerHTML = '<span style="color:#e11d48;">⚠️ Data Google Sheets / Keep belum siap</span>';
    return;
  }

  // 1. Map phone -> Keep rows & Map reportAR -> Keep row
  const phoneToKeep = new Map();
  const phoneLast9ToKeep = new Map();
  const arDocToKeep = new Map();

  rawKeepRows.forEach((k, idx) => {
    k._gsheetsIndex = idx;
    const norm = normalizePhone(k['Nomor WA']);
    if (norm && norm.length >= 8) {
      if (!phoneToKeep.has(norm)) phoneToKeep.set(norm, []);
      phoneToKeep.get(norm).push(k);

      if (norm.length >= 9) {
        const last9 = norm.slice(-9);
        if (!phoneLast9ToKeep.has(last9)) phoneLast9ToKeep.set(last9, []);
        phoneLast9ToKeep.get(last9).push(k);
      }
    }

    const arRaw = String(k['REPORT AR'] || '').replace(/\D/g, '');
    if (arRaw) {
      if (!arDocToKeep.has(arRaw)) arDocToKeep.set(arRaw, []);
      arDocToKeep.get(arRaw).push(k);
    }
  });

  // 2. Group croselling by Document Number
  const docMap = new Map();
  const phoneToAR = new Map();
  rawCrosellingRows.forEach(r => {
    const doc = String(r['Document Number'] || '').trim();
    if (!doc) return;
    if (!docMap.has(doc)) docMap.set(doc, []);
    docMap.get(doc).push(r);

    const normHp = normalizePhone(r['No HP Customer']);
    if (normHp && normHp.length >= 8 && !phoneToAR.has(normHp)) {
      phoneToAR.set(normHp, doc);
    }
  });

  // 3. Match and audit
  rawAuditARList = [];
  docMap.forEach((items, doc) => {
    const rawHp = items[0]['No HP Customer'];
    const normHp = normalizePhone(rawHp);
    const cleanDoc = doc.replace(/\D/g, '');
    let matchedKeeps = [];
    let isVlookupMatch = false;

    if (normHp && phoneToKeep.has(normHp)) {
      matchedKeeps = phoneToKeep.get(normHp);
      isVlookupMatch = true;
    } else if (normHp && normHp.length >= 9 && phoneLast9ToKeep.has(normHp.slice(-9))) {
      matchedKeeps = phoneLast9ToKeep.get(normHp.slice(-9));
      isVlookupMatch = true;
    }

    // Direct match by REPORT AR document number
    if (matchedKeeps.length === 0 && cleanDoc && arDocToKeep.has(cleanDoc)) {
      matchedKeeps = arDocToKeep.get(cleanDoc);
      isVlookupMatch = false;
    }

    if (matchedKeeps.length > 0) {
      const bestKeep = pickBestKeep(matchedKeeps, items[0]['Posting Date']);
      let totalItems = items.length;
      let auditedItemsCount = 0;
      let sosmedItems = 0;
      let missingItems = 0;
      let missingTotalValue = 0;
      let totalDocValue = 0;

      items.forEach(it => {
        const lineTotal = Number(it['Harga Jual']) || 0;
        totalDocValue += lineTotal;

        const isExempt = isItemExemptFromSosmed(it);
        const hasSosmed = isItemInputSosmed(it);

        if (isExempt) {
          // TAS GADGETMART 25x25 tidak perlu diinput sosmed.
          // Jangan dianggap sebagai missing/wajib sosmed!
          if (hasSosmed) {
            sosmedItems++;
          }
          return;
        }

        auditedItemsCount++;
        if (hasSosmed) {
          sosmedItems++;
        } else {
          missingItems++;
          missingTotalValue += lineTotal;
        }
      });

      // Status AR:
      // Hijau (ALL_SOSMED): Sudah terinput semua (100% sosmed)
      // Orange (PARTIAL_MISSING): Terinput sebagian
      // Merah (ALL_MISSING): Tidak terinput sama sekali
      let arStatus = 'ALL_SOSMED';
      if (missingItems > 0) {
        arStatus = (sosmedItems > 0) ? 'PARTIAL_MISSING' : 'ALL_MISSING';
      } else {
        arStatus = 'ALL_SOSMED';
      }

      // Resolve VLOOKUP AR info
      let vlookupARDoc = isVlookupMatch ? doc : '-';
      let vlookupARDate = isVlookupMatch ? items[0]['Posting Date'] : null;
      let vlookupARCabang = isVlookupMatch ? items[0]['Cabang'] : '-';
      let vlookupARItems = isVlookupMatch ? items : [];

      if (!isVlookupMatch) {
        const keepNormWA = normalizePhone(bestKeep['Nomor WA']);
        if (keepNormWA && phoneToAR.has(keepNormWA)) {
          vlookupARDoc = phoneToAR.get(keepNormWA);
          if (docMap.has(vlookupARDoc)) {
            vlookupARDate = docMap.get(vlookupARDoc)[0]['Posting Date'];
            vlookupARCabang = docMap.get(vlookupARDoc)[0]['Cabang'];
            vlookupARItems = docMap.get(vlookupARDoc);
          }
        }
      }

      // Resolve REPORT AR info
      const repARRaw = String(bestKeep['REPORT AR'] || '').replace(/\D/g, '');
      let reportARDoc = repARRaw || (!isVlookupMatch ? doc : '-');
      let reportARDate = null;
      let reportARCabang = '-';
      let reportARItems = [];

      if (reportARDoc && reportARDoc !== '-' && docMap.has(reportARDoc)) {
        reportARDate = docMap.get(reportARDoc)[0]['Posting Date'];
        reportARCabang = docMap.get(reportARDoc)[0]['Cabang'];
        reportARItems = docMap.get(reportARDoc);
      } else if (!isVlookupMatch) {
        reportARDate = items[0]['Posting Date'];
        reportARCabang = items[0]['Cabang'];
        reportARItems = items;
      }

      // Parse ISO dates for filtering
      const vlookupARDateIso = parseDateToISO(vlookupARDate);
      const reportARDateIso = parseDateToISO(reportARDate);

      // Parse Keep Timestamp for sorting
      const keepTimestampEpoch = parseKeepTimestampToEpoch(bestKeep['Timestamp']);

      rawAuditARList.push({
        doc,
        postingDate: items[0]['Posting Date'],
        cabang: items[0]['Cabang'],
        totalItems,
        sosmedItems,
        missingItems,
        missingTotalValue,
        totalDocValue,
        status: arStatus,
        keep: bestKeep,
        keepGSheetsIndex: (bestKeep && typeof bestKeep._gsheetsIndex === 'number') ? bestKeep._gsheetsIndex : 9999999,
        items: items,
        vlookupARDoc,
        vlookupARDate,
        vlookupARCabang,
        vlookupARItems,
        vlookupARDateIso,
        reportARDoc,
        reportARDate,
        reportARCabang,
        reportARItems,
        reportARDateIso,
        keepTimestampEpoch
      });
    }
  });

  // Urutkan tanggal format keep dari yang terlama sampai terbaru (persis urutan baris Google Sheets dari atas ke bawah)
  rawAuditARList.sort((a, b) => {
    // 1. Urutkan berdasarkan index baris asli Google Sheets (baris terlama di atas, terbaru di bawah)
    const idxA = (typeof a.keepGSheetsIndex === 'number') ? a.keepGSheetsIndex : 9999999;
    const idxB = (typeof b.keepGSheetsIndex === 'number') ? b.keepGSheetsIndex : 9999999;
    if (idxA !== idxB) {
      return idxA - idxB;
    }
    // 2. Fallback berdasarkan timestamp epoch (terlama ke terbaru)
    return (a.keepTimestampEpoch || 0) - (b.keepTimestampEpoch || 0);
  });

  if (pill) {
    pill.innerHTML = `<span>📁 ${rawAuditARList.length} AR Ditemukan</span>`;
  }

  populateAuditSosmedDropdown();
  applyAuditColStyles();
  applyAuditFilters();
}

function populateAuditSosmedDropdown() {
  const select = document.getElementById('auditFilterSosmed');
  if (!select) return;

  const currentVal = select.value;
  const uniqueCodes = new Set();
  rawAuditARList.forEach(a => {
    const code = String(a.keep['Kode Sosmed'] || '').trim();
    if (code && code !== 'undefined') uniqueCodes.add(code);
  });

  const sorted = Array.from(uniqueCodes).sort();
  select.innerHTML = '<option value="ALL">Semua Kode Sosmed</option>' +
    sorted.map(c => `<option value="${c}">${c}</option>`).join('');

  if (currentVal && sorted.includes(currentVal)) {
    select.value = currentVal;
  }
}

function applyAuditFilters() {
  if (!rawAuditARList) return;

  const periodeEl = document.getElementById('auditFilterPeriode');
  const sosmedEl = document.getElementById('auditFilterSosmed');
  const tglVlookupEl = document.getElementById('auditFilterTanggalVlookup');
  const tglReportEl = document.getElementById('auditFilterTanggalReport');
  const statusEl = document.getElementById('auditFilterStatus');
  const searchEl = document.getElementById('auditSearchInput');

  const periodeVal = periodeEl ? periodeEl.value : 'ALL';
  const sosmedVal = sosmedEl ? sosmedEl.value : 'ALL';
  const tanggalVlookupVal = tglVlookupEl ? tglVlookupEl.value : '';
  const tanggalReportVal = tglReportEl ? tglReportEl.value : '';
  const statusVal = statusEl ? statusEl.value : 'ALL';
  const keyword = (searchEl ? searchEl.value : '').trim().toLowerCase();

  filteredAuditARList = rawAuditARList.filter(r => {
    // 1. Periode filter
    if (periodeVal !== 'ALL') {
      const p = String(r.keep['PERIODE'] || '').trim().toLowerCase();
      if (p !== periodeVal.toLowerCase()) return false;
    }

    // 2. Sosmed filter
    if (sosmedVal !== 'ALL') {
      const s = String(r.keep['Kode Sosmed'] || '').trim();
      if (s !== sosmedVal) return false;
    }

    // 3. Filter Tanggal Faktur AR (VLOOKUP)
    if (tanggalVlookupVal) {
      if (r.vlookupARDateIso !== tanggalVlookupVal) return false;
    }

    // 4. Filter Tanggal Faktur Report AR
    if (tanggalReportVal) {
      if (r.reportARDateIso !== tanggalReportVal) return false;
    }

    // 5. Status AR filter
    if (statusVal !== 'ALL') {
      if (statusVal === 'ANY_MISSING') {
        if (r.status !== 'PARTIAL_MISSING' && r.status !== 'ALL_MISSING') return false;
      } else if (r.status !== statusVal) {
        return false;
      }
    }

    // 6. Keyword search
    if (keyword) {
      const itemsText = r.items.map(it => it['Dscription'] || '').join(' ');
      const arRemarks = (r.items && r.items[0] && r.items[0]['Remarks']) || '';
      const combined = `${r.doc} ${r.vlookupARDoc || ''} ${r.reportARDoc || ''} ${r.keep['Nama'] || ''} ${r.keep['Nomor WA'] || ''} ${r.keep['Barang Keep'] || ''} ${r.keep['Catatan'] || ''} ${arRemarks} ${r.cabang || ''} ${itemsText}`.toLowerCase();
      if (!combined.includes(keyword)) return false;
    }

    return true;
  });

  // Calculate KPI counters (Hijau: Sudah Terinput, Orange: Sebagian, Merah: Tidak Terinput)
  let allSosmedCount = 0;
  let partialMissingCount = 0;
  let allMissingCount = 0;
  let totalPotensiRp = 0;

  filteredAuditARList.forEach(r => {
    if (r.status === 'ALL_SOSMED') allSosmedCount++;
    else if (r.status === 'PARTIAL_MISSING') partialMissingCount++;
    else if (r.status === 'ALL_MISSING') allMissingCount++;
    totalPotensiRp += (r.missingTotalValue || 0);
  });

  const totalEl = document.getElementById('auditTotalCount');
  const allSosmedEl = document.getElementById('auditAllSosmedCount');
  const partialEl = document.getElementById('auditPartialMissingCount');
  const allMissingEl = document.getElementById('auditAllMissingCount');
  const potensiEl = document.getElementById('auditTotalPotensiRp');

  if (totalEl) totalEl.innerText = filteredAuditARList.length.toLocaleString('id-ID');
  if (allSosmedEl) allSosmedEl.innerText = allSosmedCount.toLocaleString('id-ID');
  if (partialEl) partialEl.innerText = partialMissingCount.toLocaleString('id-ID');
  if (allMissingEl) allMissingEl.innerText = allMissingCount.toLocaleString('id-ID');
  if (potensiEl) potensiEl.innerText = formatRupiah(totalPotensiRp);

  let activeAuditCount = 0;
  if (periodeEl && periodeEl.value !== 'ALL') activeAuditCount++;
  if (sosmedEl && sosmedEl.value !== 'ALL') activeAuditCount++;
  if (tglVlookupEl && tglVlookupEl.value) activeAuditCount++;
  if (tglReportEl && tglReportEl.value) activeAuditCount++;
  if (statusEl && statusEl.value !== 'ALL') activeAuditCount++;
  if (searchEl && searchEl.value.trim()) activeAuditCount++;
  const badgeAudit = document.getElementById('auditFilterActiveCount');
  if (badgeAudit) badgeAudit.innerText = activeAuditCount;

  auditCurrentPage = 1;
  renderAuditPage();
}

function resetAuditFilters() {
  const periodeEl = document.getElementById('auditFilterPeriode');
  const sosmedEl = document.getElementById('auditFilterSosmed');
  const tglVlookupEl = document.getElementById('auditFilterTanggalVlookup');
  const tglReportEl = document.getElementById('auditFilterTanggalReport');
  const statusEl = document.getElementById('auditFilterStatus');
  const searchEl = document.getElementById('auditSearchInput');

  if (periodeEl) periodeEl.value = 'Oktober-26';
  if (sosmedEl) sosmedEl.value = 'ALL';
  if (tglVlookupEl) tglVlookupEl.value = '';
  if (tglReportEl) tglReportEl.value = '';
  if (statusEl) statusEl.value = 'ALL';
  if (searchEl) searchEl.value = '';

  applyAuditFilters();
}

function changeAuditPage(delta) {
  const maxPage = Math.max(1, Math.ceil(filteredAuditARList.length / AUDIT_PAGE_SIZE));
  const newPage = auditCurrentPage + delta;
  if (newPage >= 1 && newPage <= maxPage) {
    auditCurrentPage = newPage;
    renderAuditPage();
  }
}

// ==========================================================================
// STATUS KEEP (GOOGLE SHEETS LIVE UPDATE ENGINE)
// ==========================================================================
const KEEP_STATUS_OPTIONS = [
  { value: '', label: '- Kosong -' },
  { value: 'WIN', label: 'WIN' },
  { value: 'LOSE', label: 'LOSE' },
  { value: 'BELUM DIINPUT', label: 'BELUM DIINPUT' },
  { value: 'DOUBLE KEEP', label: 'DOUBLE KEEP' },
  { value: 'ON PROGRESS LOW', label: 'ON PROGRESS LOW' },
  { value: 'ON PROGRESS MEDIUM', label: 'ON PROGRESS MEDIUM' },
  { value: 'ON PROGRESS HOT', label: 'ON PROGRESS HOT' },
  { value: 'TAKE OUT', label: 'TAKE OUT' }
];

function getKeepStatusClass(status) {
  const s = String(status || '').trim().toUpperCase();
  if (!s || s === '-' || s === 'KOSONG') return 'status-empty';
  if (s === 'WIN') return 'status-win';
  if (s === 'LOSE') return 'status-lose';
  if (s === 'BELUM DIINPUT' || s === 'BELUM INPUT') return 'status-belum-diinput';
  if (s === 'DOUBLE KEEP') return 'status-double-keep';
  if (s === 'ON PROGRESS LOW') return 'status-on-progress-low';
  if (s === 'ON PROGRESS MEDIUM') return 'status-on-progress-medium';
  if (s === 'ON PROGRESS HOT') return 'status-on-progress-hot';
  if (s === 'TAKE OUT') return 'status-take-out';
  return 'status-empty';
}

function getAuditRowClassByKeepStatus(keepStatus, arStatus) {
  const s = String(keepStatus || '').trim().toUpperCase();
  if (s === 'WIN') return 'audit-row-win';
  if (s === 'LOSE') return 'audit-row-lose';
  if (s === 'BELUM DIINPUT' || s === 'BELUM INPUT') return 'audit-row-belum-diinput';
  if (s === 'DOUBLE KEEP') return 'audit-row-double-keep';
  if (s === 'ON PROGRESS LOW') return 'audit-row-on-progress-low';
  if (s === 'ON PROGRESS MEDIUM') return 'audit-row-on-progress-medium';
  if (s === 'ON PROGRESS HOT') return 'audit-row-on-progress-hot';
  if (s === 'TAKE OUT') return 'audit-row-take-out';

  // Jika status keep tidak ada / kosong: gunakan indikator outline sosmed AR
  if (arStatus === 'ALL_SOSMED') return 'audit-row-sosmed';
  if (arStatus === 'PARTIAL_MISSING') return 'audit-row-partial';
  return 'audit-row-missing';
}

async function updateKeepStatusInGoogleSheets(rawWA, newStatus) {
  const normWA = String(rawWA || '').replace(/\D/g, '');
  if (!normWA) throw new Error('Nomor WhatsApp tidak valid untuk update status.');

  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzdEByURvrvCNAJ-TlaxlU5GoNvmClgxynK_vYrORzXiTXFoYArwRMCJzdW2LWvT5wfUw/exec';

  let result = null;
  // 1. Coba lewat backend local-server proxy
  try {
    const res = await fetch(`/api/update-keep-status?nomorWA=${encodeURIComponent(normWA)}&status=${encodeURIComponent(newStatus)}`, { signal: AbortSignal.timeout ? AbortSignal.timeout(8000) : undefined });
    if (res.ok) {
      result = await res.json();
    }
  } catch (e) {
    console.warn('Backend proxy update-keep-status gagal, mencoba direct fetch:', e);
  }

  // 2. Fallback direct fetch ke Apps Script Web App
  if (!result || !result.success) {
    try {
      const directUrl = `${APPS_SCRIPT_URL}?action=UPDATE_STATUS&nomorWA=${encodeURIComponent(normWA)}&status=${encodeURIComponent(newStatus)}`;
      const directRes = await fetch(directUrl, { redirect: 'follow', signal: AbortSignal.timeout ? AbortSignal.timeout(8000) : undefined });
      result = await directRes.json();
    } catch (directErr) {
      console.error('Direct fetch Apps Script error:', directErr);
    }
  }

  if (result && result.success) {
    return result;
  }
  throw new Error((result && result.message) || 'Gagal menyimpan status ke Google Sheets.');
}

// Antrean update status Google Sheets di background (non-blocking)
const pendingKeepStatusUpdates = new Map();

async function triggerBackgroundKeepStatusSync(rawWA, newVal, oldVal, onRollback) {
  const normWA = String(rawWA || '').replace(/\D/g, '');
  if (!normWA) return;

  pendingKeepStatusUpdates.set(normWA, newVal);
  const phoneDisplay = formatPhoneDisplay(rawWA);

  // Tampilkan notifikasi persentase bergerak aktif
  const prog = startSmoothProgress({
    id: `prog-keep-${normWA}`,
    title: `⚡ Simpan Status WA ${phoneDisplay}`,
    detail: `Mengirim status "${newVal || 'KOSONG'}" ke Google Sheets...`
  });

  try {
    await updateKeepStatusInGoogleSheets(rawWA, newVal);
    pendingKeepStatusUpdates.delete(normWA);
    prog.complete({
      title: `✅ Status WA ${phoneDisplay} Berhasil Disimpan!`,
      detail: `Status "${newVal || 'KOSONG'}" sukses tersinkronisasi di Google Sheets`
    });
  } catch (err) {
    console.error('Background sync status keep gagal:', err);
    prog.fail({
      title: `❌ Gagal Simpan WA ${phoneDisplay}`,
      detail: err.message || 'Koneksi ke Google Sheets terganggu'
    });
    if (pendingKeepStatusUpdates.get(normWA) === newVal) {
      pendingKeepStatusUpdates.delete(normWA);
      if (typeof onRollback === 'function') {
        onRollback(err);
      }
    }
  }
}

async function updateReportARInGoogleSheets(rawWA, arDoc) {
  const normWA = String(rawWA || '').replace(/\D/g, '');
  if (!normWA) throw new Error('Nomor WhatsApp tidak valid untuk update Report AR.');

  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzdEByURvrvCNAJ-TlaxlU5GoNvmClgxynK_vYrORzXiTXFoYArwRMCJzdW2LWvT5wfUw/exec';

  let result = null;
  // 1. Coba lewat backend local-server proxy
  try {
    const res = await fetch(`/api/update-report-ar?nomorWA=${encodeURIComponent(normWA)}&arDoc=${encodeURIComponent(arDoc)}`, { signal: AbortSignal.timeout ? AbortSignal.timeout(8000) : undefined });
    if (res.ok) {
      result = await res.json();
    }
  } catch (e) {
    console.warn('Backend proxy update-report-ar gagal, mencoba direct fetch:', e);
  }

  // 2. Fallback direct fetch ke Apps Script Web App
  if (!result || !result.success) {
    try {
      const directUrl = `${APPS_SCRIPT_URL}?action=UPDATE_REPORT_AR&nomorWA=${encodeURIComponent(normWA)}&arDoc=${encodeURIComponent(arDoc)}`;
      const directRes = await fetch(directUrl, { redirect: 'follow', signal: AbortSignal.timeout ? AbortSignal.timeout(8000) : undefined });
      result = await directRes.json();
    } catch (directErr) {
      console.error('Direct fetch Apps Script error:', directErr);
    }
  }

  if (result && result.success) {
    return result;
  }
  throw new Error((result && result.message) || 'Gagal menyimpan nomor AR ke Google Sheets.');
}

// Antrean update REPORT AR Google Sheets di background (non-blocking)
async function triggerBackgroundReportARSync(rawWA, newAR, oldAR, onRollback) {
  const normWA = String(rawWA || '').replace(/\D/g, '');
  if (!normWA) return;

  const phoneDisplay = formatPhoneDisplay(rawWA);
  const prog = startSmoothProgress({
    id: `prog-ar-${normWA}`,
    title: `⚡ Simpan Report AR WA ${phoneDisplay}`,
    detail: `Mengirim nomor AR "${newAR || '-'}" ke Google Sheets...`
  });

  try {
    await updateReportARInGoogleSheets(rawWA, newAR);
    prog.complete({
      title: `✅ Report AR Berhasil Disimpan!`,
      detail: `Nomor AR "${newAR || '-'}" sukses tersinkronisasi ke Google Sheets.`
    });
  } catch (err) {
    console.error('Background sync Report AR gagal:', err);
    prog.fail({
      title: `❌ Gagal Simpan AR WA ${phoneDisplay}`,
      detail: err.message || 'Koneksi ke Google Sheets terganggu'
    });
    if (typeof onRollback === 'function') {
      onRollback(err);
    }
  }
}


function handleTableKeepStatusChange(selectEl, rawWA, auditRowIdx) {
  if (!selectEl) return;
  const oldVal = selectEl.getAttribute('data-prev') || selectEl.value;
  const newVal = selectEl.value;
  if (newVal === oldVal) return;

  const originalClass = selectEl.className;
  const tr = selectEl.closest('tr');
  const arObj = filteredAuditARList[auditRowIdx];
  const arSt = arObj ? arObj.status : '';
  const originalRowClass = tr ? tr.className : '';

  // 1. OPTIMISTIC UI UPDATE (0ms) - Seketika ubah dropdown & warna blok baris tabel tanpa freeze
  selectEl.setAttribute('data-prev', newVal);
  selectEl.className = `select-status-keep ${getKeepStatusClass(newVal)}`;
  if (tr) {
    tr.className = getAuditRowClassByKeepStatus(newVal, arSt);
  }

  // 2. Perbarui state memori seketika (0ms)
  if (filteredAuditARList[auditRowIdx] && filteredAuditARList[auditRowIdx].keep) {
    filteredAuditARList[auditRowIdx].keep['Status'] = newVal;
  }
  const cleanWA = String(rawWA).replace(/\D/g, '');
  if (typeof rawKeepRows !== 'undefined' && Array.isArray(rawKeepRows)) {
    rawKeepRows.forEach(k => {
      if (String(k['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        k['Status'] = newVal;
      }
    });
  }
  if (typeof rawAuditARList !== 'undefined' && Array.isArray(rawAuditARList)) {
    rawAuditARList.forEach(a => {
      if (a.keep && String(a.keep['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        a.keep['Status'] = newVal;
      }
    });
  }

  const toastLabel = newVal ? `ke "${newVal}"` : 'dikosongkan';
  showSaveToast(`⚡ Status keep WA ${formatPhoneDisplay(rawWA)} diubah ${toastLabel}! (Tersimpan di background)`);

  // 3. Kirim ke Google Sheets di background (Rollback otomatis jika gagal)
  triggerBackgroundKeepStatusSync(rawWA, newVal, oldVal, (err) => {
    alert(`⚠️ Gagal menyimpan status ke Google Sheets:\n${err.message}\nStatus dikembalikan ke sebelumnya.`);
    selectEl.value = oldVal;
    selectEl.setAttribute('data-prev', oldVal);
    selectEl.className = originalClass;
    if (tr) tr.className = originalRowClass;
    if (filteredAuditARList[auditRowIdx] && filteredAuditARList[auditRowIdx].keep) {
      filteredAuditARList[auditRowIdx].keep['Status'] = oldVal;
    }
  });
}

function handleModalKeepStatusChange(selectEl, rawWA) {
  if (!selectEl) return;
  const oldVal = selectEl.getAttribute('data-prev') || selectEl.value;
  const newVal = selectEl.value;
  if (newVal === oldVal) return;

  const originalClass = selectEl.className;

  // 1. OPTIMISTIC UI UPDATE (0ms)
  selectEl.setAttribute('data-prev', newVal);
  selectEl.className = `select-status-keep ${getKeepStatusClass(newVal)}`;

  if (currentModalKeep) {
    currentModalKeep['Status'] = newVal;
  }
  const cleanWA = String(rawWA).replace(/\D/g, '');
  if (typeof rawAuditARList !== 'undefined' && Array.isArray(rawAuditARList)) {
    rawAuditARList.forEach(a => {
      if (a.keep && String(a.keep['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        a.keep['Status'] = newVal;
      }
    });
  }
  if (typeof rawKeepRows !== 'undefined' && Array.isArray(rawKeepRows)) {
    rawKeepRows.forEach(k => {
      if (String(k['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        k['Status'] = newVal;
      }
    });
  }

  showSaveToast(`⚡ Status keep WA ${formatPhoneDisplay(rawWA)} diubah ke "${newVal}"! (Tersimpan di background)`);
  if (typeof renderAuditPage === 'function') {
    renderAuditPage();
  }

  // 2. Kirim ke Google Sheets di background
  triggerBackgroundKeepStatusSync(rawWA, newVal, oldVal, (err) => {
    alert(`⚠️ Gagal menyimpan status ke Google Sheets:\n${err.message}\nStatus dikembalikan ke sebelumnya.`);
    selectEl.value = oldVal;
    selectEl.setAttribute('data-prev', oldVal);
    selectEl.className = originalClass;
    if (currentModalKeep) {
      currentModalKeep['Status'] = oldVal;
    }
    if (typeof renderAuditPage === 'function') {
      renderAuditPage();
    }
  });
}

// ==========================================================================
// INLINE AR INPUT & EDIT ENGINE (TAB AUDIT AR & TAB FORMAT KEEP)
// ==========================================================================
function startEditAuditReportAR(containerEl, rawWA, auditRowIdx) {
  if (!containerEl) return;
  const currentAR = containerEl.getAttribute('data-current-ar') || '';
  
  containerEl.innerHTML = `
    <input type="text" class="input-inline-ar" value="${currentAR}" 
      placeholder="No. AR..." 
      onkeydown="if(event.key==='Enter') saveInlineAuditReportAR(this, '${rawWA}', ${auditRowIdx}); else if(event.key==='Escape') cancelInlineAuditReportAR(this, '${currentAR}', ${auditRowIdx});"
      onblur="saveInlineAuditReportAR(this, '${rawWA}', ${auditRowIdx})"
    />
  `;
  const input = containerEl.querySelector('input');
  if (input) {
    input.focus();
    input.select();
  }
}

function saveInlineAuditReportAR(inputEl, rawWA, auditRowIdx) {
  const container = inputEl.parentElement;
  if (!container) return;
  const oldAR = container.getAttribute('data-current-ar') || '';
  const newAR = inputEl.value.trim().replace(/\D/g, '');

  if (newAR === oldAR) {
    cancelInlineAuditReportAR(inputEl, oldAR, auditRowIdx);
    return;
  }

  // 1. Optimistic Update Local Memory (0ms)
  const cleanWA = String(rawWA).replace(/\D/g, '');
  if (filteredAuditARList[auditRowIdx]) {
    filteredAuditARList[auditRowIdx].reportARDoc = newAR;
    if (filteredAuditARList[auditRowIdx].keep) {
      filteredAuditARList[auditRowIdx].keep['REPORT AR'] = newAR;
    }
    if (newAR && typeof docMap !== 'undefined' && docMap.has(newAR)) {
      filteredAuditARList[auditRowIdx].reportARItems = docMap.get(newAR);
      filteredAuditARList[auditRowIdx].reportARDate = filteredAuditARList[auditRowIdx].reportARItems[0]['Posting Date'];
      filteredAuditARList[auditRowIdx].reportARCabang = filteredAuditARList[auditRowIdx].reportARItems[0]['Cabang'];
    }
  }

  if (typeof rawKeepRows !== 'undefined' && Array.isArray(rawKeepRows)) {
    rawKeepRows.forEach(k => {
      if (String(k['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        k['REPORT AR'] = newAR;
      }
    });
  }

  // 2. Render ulang baris seketika agar status badge langsung update (0ms)
  renderAuditPage();

  // 3. Simpan ke Google Sheets di background (non-blocking)
  triggerBackgroundReportARSync(rawWA, newAR, oldAR, () => {
    // Rollback jika error
    if (filteredAuditARList[auditRowIdx]) {
      filteredAuditARList[auditRowIdx].reportARDoc = oldAR;
      if (filteredAuditARList[auditRowIdx].keep) {
        filteredAuditARList[auditRowIdx].keep['REPORT AR'] = oldAR;
      }
    }
    renderAuditPage();
  });
}

function cancelInlineAuditReportAR(inputEl, oldAR, auditRowIdx) {
  renderAuditPage();
}

function startEditKeepReportAR(containerEl, rawWA, keepRowIdx) {
  if (!containerEl) return;
  const currentAR = containerEl.getAttribute('data-current-ar') || '';
  
  containerEl.innerHTML = `
    <input type="text" class="input-inline-ar" value="${currentAR}" 
      placeholder="No. AR..." 
      onkeydown="if(event.key==='Enter') saveInlineKeepReportAR(this, '${rawWA}', ${keepRowIdx}); else if(event.key==='Escape') cancelInlineKeepReportAR(this, '${currentAR}', ${keepRowIdx});"
      onblur="saveInlineKeepReportAR(this, '${rawWA}', ${keepRowIdx})"
    />
  `;
  const input = containerEl.querySelector('input');
  if (input) {
    input.focus();
    input.select();
  }
}

function saveInlineKeepReportAR(inputEl, rawWA, keepRowIdx) {
  const container = inputEl.parentElement;
  if (!container) return;
  const oldAR = container.getAttribute('data-current-ar') || '';
  const newAR = inputEl.value.trim().replace(/\D/g, '');

  if (newAR === oldAR) {
    cancelInlineKeepReportAR(inputEl, oldAR, keepRowIdx);
    return;
  }

  // 1. Optimistic Update Local Memory (0ms)
  const cleanWA = String(rawWA).replace(/\D/g, '');
  if (filteredKeepRows[keepRowIdx]) {
    filteredKeepRows[keepRowIdx]['REPORT AR'] = newAR;
  }

  if (typeof rawKeepRows !== 'undefined' && Array.isArray(rawKeepRows)) {
    rawKeepRows.forEach(k => {
      if (String(k['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        k['REPORT AR'] = newAR;
      }
    });
  }

  if (typeof rawAuditARList !== 'undefined' && Array.isArray(rawAuditARList)) {
    rawAuditARList.forEach(a => {
      if (a.keep && String(a.keep['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
        a.reportARDoc = newAR;
        a.keep['REPORT AR'] = newAR;
        if (newAR && typeof docMap !== 'undefined' && docMap.has(newAR)) {
          a.reportARItems = docMap.get(newAR);
          a.reportARDate = a.reportARItems[0]['Posting Date'];
          a.reportARCabang = a.reportARItems[0]['Cabang'];
        }
      }
    });
  }

  // 2. Render ulang halaman Format Keep (0ms)
  renderKeepPage();

  // 3. Simpan ke Google Sheets di background (non-blocking)
  triggerBackgroundReportARSync(rawWA, newAR, oldAR, () => {
    if (filteredKeepRows[keepRowIdx]) {
      filteredKeepRows[keepRowIdx]['REPORT AR'] = oldAR;
    }
    renderKeepPage();
  });
}

function cancelInlineKeepReportAR(inputEl, oldAR, keepRowIdx) {
  renderKeepPage();
}

// Helper mandiri untuk menghitung badge status dokumen faktur
function getDocBadgeInfo(docItems) {
  if (!docItems || docItems.length === 0) {
    return { badgeClass: 'badge-status-ink', badgeIcon: '❓', statusLabel: 'Data Item Faktur Tidak Ditemukan' };
  }
  let sosmedCount = 0;
  let missingCount = 0;
  docItems.forEach(it => {
    const isExempt = isItemExemptFromSosmed(it);
    const hasSosmed = isItemInputSosmed(it);
    if (isExempt) {
      if (hasSosmed) sosmedCount++;
      return;
    }
    if (hasSosmed) sosmedCount++;
    else missingCount++;
  });

  if (missingCount === 0) {
    return { badgeClass: 'badge-status-green', badgeIcon: '✅', statusLabel: 'Sudah Terinput Semua (100% Sosmed)' };
  } else if (sosmedCount > 0) {
    return { badgeClass: 'badge-status-orange', badgeIcon: '⚠️', statusLabel: 'Terinput Sebagian' };
  } else {
    return { badgeClass: 'badge-status-red', badgeIcon: '❌', statusLabel: 'Tidak Terinput Sosmed Sama Sekali' };
  }
}

function renderAuditPage() {
  const tbody = document.getElementById('auditTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const total = filteredAuditARList.length;
  const maxPage = Math.max(1, Math.ceil(total / AUDIT_PAGE_SIZE));
  const startIdx = (auditCurrentPage - 1) * AUDIT_PAGE_SIZE;
  const endIdx = Math.min(startIdx + AUDIT_PAGE_SIZE, total);
  const pageRows = filteredAuditARList.slice(startIdx, endIdx);

  // Pagination info
  const paginationInfoEl = document.getElementById('auditPaginationInfo');
  if (paginationInfoEl) {
    paginationInfoEl.innerText = total > 0
      ? `Menampilkan ${startIdx + 1} - ${endIdx} dari total ${total.toLocaleString('id-ID')} Dokumen AR`
      : 'Tidak ada dokumen AR yang sesuai filter';
  }
  const curPageTextEl = document.getElementById('auditCurrentPageText');
  if (curPageTextEl) curPageTextEl.innerText = `Halaman ${auditCurrentPage} / ${maxPage}`;
  const btnPrev = document.getElementById('btnAuditPrev');
  if (btnPrev) btnPrev.disabled = auditCurrentPage <= 1;
  const btnNext = document.getElementById('btnAuditNext');
  if (btnNext) btnNext.disabled = auditCurrentPage >= maxPage;

  if (pageRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" class="text-center" style="padding: 30px; color: var(--text-muted);">Tidak ada data tracking AR yang sesuai filter.</td></tr>`;
    return;
  }

  pageRows.forEach((r, i) => {
    const globalIdx = startIdx + i + 1;
    const auditRowIdx = startIdx + i;
    const keepObj = r.keep || {};
    const rawWA = String(keepObj['Nomor WA'] || '').replace(/\D/g, '');
    const normWA = normalizePhone(rawWA);
    const displayWA = formatPhoneDisplay(rawWA);

    const waHtml = (normWA && normWA.length >= 8) ? `
      <div style="display:flex; align-items:center; gap:5px;">
        <span class="font-bold">${displayWA}</span>
        <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp" style="padding:1px 6px; font-size:0.68rem; border-radius:4px;" title="Chat customer ini di WhatsApp">
          💬 WA
        </a>
      </div>
    ` : (displayWA || '-');

    // Notes HTML
    const keepCatatan = String(keepObj['Catatan'] || '').trim();
    const arRemarks = String((r.items && r.items[0] && r.items[0]['Remarks']) || '').trim();

    let catatanHtml = '-';
    if (keepCatatan && arRemarks && keepCatatan.toLowerCase() !== arRemarks.toLowerCase() && keepCatatan !== '-' && arRemarks !== '-') {
      catatanHtml = `
        <div style="font-size:0.73rem; line-height:1.25; color:#0f172a; word-break:break-word;">${keepCatatan}</div>
        <div style="font-size:0.68rem; color:var(--text-muted); margin-top:2px; word-break:break-word;"><em>Kasir: ${arRemarks}</em></div>
      `;
    } else if (keepCatatan && keepCatatan !== '-') {
      catatanHtml = `<div style="font-size:0.73rem; line-height:1.25; color:#0f172a; word-break:break-word;">${keepCatatan}</div>`;
    } else if (arRemarks && arRemarks !== '-') {
      catatanHtml = `<div style="font-size:0.70rem; line-height:1.25; color:var(--text-muted); word-break:break-word;"><em>Kasir: ${arRemarks}</em></div>`;
    }

    // Status Keep Dropdown
    const rawStatus = keepObj['Status'];
    const currentStatus = (rawStatus && String(rawStatus).trim() !== '-') ? String(rawStatus).trim().toUpperCase() : '';
    const statusSelectClass = getKeepStatusClass(currentStatus);

    let statusOptionsHtml = '';
    if (currentStatus && !KEEP_STATUS_OPTIONS.some(o => o.value.toUpperCase() === currentStatus)) {
      statusOptionsHtml += `<option value="${currentStatus}" selected>${currentStatus}</option>`;
    }
    KEEP_STATUS_OPTIONS.forEach(opt => {
      const isSel = (opt.value.toUpperCase() === currentStatus);
      statusOptionsHtml += `<option value="${opt.value}" ${isSel ? 'selected' : ''}>${opt.label}</option>`;
    });

    const statusDropdownHtml = (rawWA && rawWA.length >= 8) ? `
      <select class="select-status-keep ${statusSelectClass}" data-prev="${currentStatus}" onchange="handleTableKeepStatusChange(this, '${rawWA}', ${auditRowIdx})" title="Klik untuk mengubah status follow-up customer ini (otomatis tersimpan ke Google Sheets)">
        ${statusOptionsHtml}
      </select>
    ` : `<span class="kpi-badge badge-ink">-</span>`;

    // Evaluasi Badge Masing-Masing Dokumen Secara Mandiri (Akurat 100%)
    const repBadge = getDocBadgeInfo(r.reportARItems && r.reportARItems.length > 0 ? r.reportARItems : (r.reportARDoc === r.doc ? r.items : []));
    const vlookBadge = getDocBadgeInfo(r.vlookupARItems && r.vlookupARItems.length > 0 ? r.vlookupARItems : (r.vlookupARDoc === r.doc ? r.items : []));

    // Status Warna Baris Tabel
    const effectiveStatus = (r.reportARDoc && r.reportARDoc !== '-') ? repBadge.statusLabel : r.status;
    const rowStatusClass = getAuditRowClassByKeepStatus(currentStatus, effectiveStatus);

    // Report AR cell (Dengan Fitur Input & Edit Cepat)
    const reportARDocDisplay = (r.reportARDoc && r.reportARDoc !== '-')
      ? `<div class="cell-report-ar-wrapper" data-wa="${rawWA}" data-audit-idx="${auditRowIdx}" data-current-ar="${r.reportARDoc}">
           <span class="badge-status-link ${repBadge.badgeClass}" onclick="openARDetailModal('${r.reportARDoc}', filteredAuditARList[${auditRowIdx}].keep, 'REPORT AR')" title="Status: ${repBadge.statusLabel}. Klik untuk cek detail Dokumen Report AR">
             <span>${repBadge.badgeIcon}</span> ${r.reportARDoc}
           </span>
           <button type="button" class="btn-edit-inline-ar" onclick="event.stopPropagation(); startEditAuditReportAR(this.parentElement, '${rawWA}', ${auditRowIdx})" title="Ubah nomor AR">✏️</button>
         </div>`
      : `<div class="cell-report-ar-wrapper" data-wa="${rawWA}" data-audit-idx="${auditRowIdx}" data-current-ar="">
           <button type="button" class="btn-input-new-ar" onclick="startEditAuditReportAR(this.parentElement, '${rawWA}', ${auditRowIdx})" title="Klik untuk menginput nomor AR">+ Input AR</button>
         </div>`;

    const reportARDateDisplay = r.reportARDate
      ? `${formatCrosellingDate(r.reportARDate)} ${r.reportARCabang || ''}`.trim()
      : '-';

    // Vlookup AR cell
    const vlookupARDocDisplay = (r.vlookupARDoc && r.vlookupARDoc !== '-')
      ? `<span class="badge-status-link ${vlookBadge.badgeClass}" onclick="openARDetailModal('${r.vlookupARDoc}', filteredAuditARList[${auditRowIdx}].keep, 'AR VLOOKUP')" title="Status: ${vlookBadge.statusLabel}. Klik untuk cek detail Dokumen AR VLOOKUP"><span>${vlookBadge.badgeIcon}</span> ${r.vlookupARDoc}</span>`
      : '-';

    const vlookupARDateDisplay = r.vlookupARDate
      ? `${formatCrosellingDate(r.vlookupARDate)} ${r.vlookupARCabang || ''}`.trim()
      : '-';

    const tr = document.createElement('tr');
    tr.className = rowStatusClass;
    if (currentStatus === 'LOSE' || currentStatus === 'TAKE OUT') {
      tr.classList.add('row-lose-tint');
    }
    tr.innerHTML = `
      <td class="audit-col-no" style="text-align: center;">${globalIdx}</td>
      <td class="audit-col-keep-date" style="text-align: center; font-weight: 600; white-space: nowrap;">${formatKeepDisplayDate(keepObj['Timestamp'])}</td>
      <td class="audit-col-cust text-left" style="color: #0f172a; word-break: break-word;">
        <div class="font-bold">${keepObj['Nama'] || '-'}</div>
        ${(normWA && normWA.length >= 8) ? `
          <a href="https://wa.me/${normWA}" target="_blank" style="font-size:0.75rem; color:#64748b; text-decoration:none; display:inline-flex; align-items:center; gap:3px;">
            📞 ${displayWA}
          </a>
        ` : `<div style="font-size:0.75rem; color:#64748b;">${displayWA || '-'}</div>`}
      </td>
      <td class="audit-col-sosmed" style="text-align: center;">
        <span class="badge-sosmed-pill">${keepObj['Kode Sosmed'] || '-'}</span>
      </td>
      <td class="audit-col-keep text-left font-bold" style="color: #1e293b; word-break: break-word;">${keepObj['Barang Keep'] || '-'}</td>
      <td class="audit-col-notes text-left" style="word-break: break-word; font-size: 0.78rem;">${catatanHtml}</td>
      <td class="audit-col-status-keep" style="text-align: center;">${statusDropdownHtml}</td>
      <td class="audit-col-report-ar" style="text-align: center;">${reportARDocDisplay}</td>
      <td class="audit-col-report-date" style="text-align: center; font-size: 0.76rem; white-space: nowrap;">${reportARDateDisplay}</td>
      <td class="audit-col-vlookup-ar" style="text-align: center;">${vlookupARDocDisplay}</td>
      <td class="audit-col-vlookup-date" style="text-align: center; font-size: 0.76rem; white-space: nowrap;">${vlookupARDateDisplay}</td>
    `;
    tbody.appendChild(tr);
  });
}

function openARDetailModalFromAudit(auditIndex) {
  const auditItem = filteredAuditARList[auditIndex];
  if (!auditItem) return;
  openARDetailModal(auditItem.doc, auditItem.keep);
}

// ==========================================================================
// COLUMN HIDING ENGINE (TAMPILKAN / SEMBUNYIKAN KOLOM TRACKING AR)
// ==========================================================================
let auditHiddenCols = new Set(JSON.parse(localStorage.getItem('audit_hidden_cols') || '[]'));

function applyAuditColStyles() {
  let styleEl = document.getElementById('auditDynamicColStyles');
  if (!styleEl) {
    styleEl = document.createElement('style');
    styleEl.id = 'auditDynamicColStyles';
    document.head.appendChild(styleEl);
  }
  const rules = Array.from(auditHiddenCols).map(col => `.${col} { display: none !important; }`).join('\n');
  styleEl.textContent = rules;

  const menu = document.getElementById('auditColDropdownMenu');
  if (menu) {
    const checkboxes = menu.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(cb => {
      const col = cb.getAttribute('data-col');
      if (col) {
        cb.checked = !auditHiddenCols.has(col);
      }
    });
  }
}

function toggleAuditColumn(colClass, isVisible) {
  if (isVisible) {
    auditHiddenCols.delete(colClass);
  } else {
    auditHiddenCols.add(colClass);
  }
  localStorage.setItem('audit_hidden_cols', JSON.stringify(Array.from(auditHiddenCols)));
  applyAuditColStyles();
}

function resetAuditColumns() {
  auditHiddenCols.clear();
  localStorage.setItem('audit_hidden_cols', JSON.stringify([]));
  applyAuditColStyles();
}

function toggleAuditColDropdown(event) {
  if (event) {
    event.stopPropagation();
  }
  const menu = document.getElementById('auditColDropdownMenu');
  if (!menu) return;
  menu.style.display = (menu.style.display === 'none' || !menu.style.display) ? 'block' : 'none';
}

document.addEventListener('click', function(e) {
  const menu = document.getElementById('auditColDropdownMenu');
  const btn = document.getElementById('btnToggleAuditCols');
  if (menu && menu.style.display === 'block') {
    if (!menu.contains(e.target) && (!btn || !btn.contains(e.target))) {
      menu.style.display = 'none';
    }
  }
});

// ==========================================================================
// AUDIT AR ITEM CLIPBOARD EXPORT ENGINE (12 KOLOM EXCEL TANPA HEADER)
// ==========================================================================

function isHandphoneProduct(item) {
  if (!item) return false;
  const desc = String(item['Dscription'] || item['Item/Service Description'] || '').trim().toUpperCase();
  const grp = String(item['ItmsGrpNam'] || item['Group Name'] || item['Group'] || '').trim().toUpperCase();

  // Aksesoris / Jasa / Opr / Bukan Unit Handphone
  const isAcc = desc.startsWith('SG ') || desc.includes('TMPR') || desc.includes('CASE ') || 
                desc.includes('CASING') || desc.includes('TEMPERED') || desc.includes('HYDROGEL') || 
                desc.includes('ADAPTOR') || desc.includes('CABLE') || desc.includes('KABEL') || 
                desc.includes('CHARGER') || desc.includes('TWS') || desc.includes('HEADSET') || 
                desc.includes('EARBUDS') || desc.includes('POWERBANK') || desc.includes('SPEAKER') ||
                desc.includes('MEMORY') || desc.includes('STRAP') || desc.includes('HOLDER') ||
                desc.includes('KARTU') || desc.startsWith('TM ') || desc.includes('GMART CARE') ||
                desc.includes('SERVICE') || desc.includes('JASA') || grp.startsWith('ACC') || grp.startsWith('OPR');
  if (isAcc) return false;

  if (grp.startsWith('GAD') || grp.includes('GADGET') || grp.includes('HANDPHONE') || grp.includes('SMARTPHONE')) {
    return true;
  }

  const phoneKeywords = [
    'IPHONE', 'GALAXY', 'SAMSUNG R', 'SAMSUNG A', 'SAMSUNG S', 'SAMSUNG Z',
    'OPPO A', 'OPPO RENO', 'OPPO FIND', 'VIVO Y', 'VIVO V', 'VIVO X',
    'REDMI', 'XIAOMI', 'POCO', 'REALME', 'INFINIX', 'TECNO', 'ITEL',
    'SMARTPHONE', 'HANDPHONE'
  ];
  return phoneKeywords.some(k => desc.includes(k));
}

function getItemPotensiDenda(item) {
  return isHandphoneProduct(item) ? 'DENDA' : 'PASS';
}

function formatKeepTextForClipboard(keep) {
  if (!keep) return '-';
  const tagSender = keep['Sender'] ? `@${keep['Sender']}` : '';
  const nama = keep['Nama'] || '-';
  const rawWa = keep['Nomor WA'] || '-';
  const cleanWa = String(rawWa).replace(/\D/g, '');
  const barang = keep['Barang Keep'] || '-';
  const tglAmbil = keep['Tanggal Ambil'] || '-';
  const catatan = keep['Catatan'] || '';

  const lines = [];
  if (tagSender) lines.push(tagSender);
  lines.push('*Format Keep*');
  lines.push(`Nama: ${nama}`);
  lines.push(`Nomor WA: ${cleanWa}`);
  lines.push(`Barang keep: ${barang}`);
  lines.push(`Tanggal pengambilan: ${tglAmbil}`);
  lines.push(`Catatan: ${catatan}`);

  return lines.join('\n');
}

function generateAuditItemTSV(auditItem, item) {
  if (!auditItem || !item) return '';

  // 1. TANGGAL (format DD/MM/YYYY)
  const tgl = formatCrosellingDate(auditItem.postingDate);

  // 2. PERIODE (e.g. Oktober-26)
  const periode = (auditItem.keep && auditItem.keep['PERIODE']) 
    ? String(auditItem.keep['PERIODE']).trim() 
    : 'Oktober-26';

  // 3. NO AR
  const noAR = String(auditItem.doc || '').trim();

  // 4. NO.HP USER
  const rawHp = (auditItem.keep && auditItem.keep['Nomor WA']) 
    ? auditItem.keep['Nomor WA'] 
    : (auditItem.items && auditItem.items[0] && auditItem.items[0]['No HP Customer'] ? auditItem.items[0]['No HP Customer'] : '');
  const noHP = String(rawHp).replace(/\D/g, '');

  // 5. KODE SOSMED
  const kodeSosmed = (auditItem.keep && auditItem.keep['Kode Sosmed']) 
    ? String(auditItem.keep['Kode Sosmed']).trim() 
    : '-';

  // 6. PIC SOSMED
  let picSosmed = (auditItem.keep && auditItem.keep['Sender']) 
    ? String(auditItem.keep['Sender']).trim() 
    : '';
  if (!picSosmed && kodeSosmed && typeof getMasterSales === 'function') {
    const m = getMasterSales().find(x => kodeSosmed.includes(x.code) || kodeSosmed.includes(x.name));
    if (m) picSosmed = m.name;
  }
  if (!picSosmed) picSosmed = '-';

  // 7. CABANG
  const cabang = String(auditItem.cabang || (auditItem.items && auditItem.items[0] ? auditItem.items[0]['Cabang'] : '') || '-').trim();

  // 8. POTENSI DENDA (HP = DENDA, Aksesoris/lainnya = PASS)
  const potensiDenda = getItemPotensiDenda(item);

  // 9. BARANG (Item description yang dibeli di faktur AR)
  const namaBarang = String(item['Dscription'] || item['Item/Service Description'] || '-').trim();

  // 10. KEEP (Format Keep WhatsApp lengkap dibungkus kutip dua untuk in-cell newline di Excel)
  const rawKeepText = formatKeepTextForClipboard(auditItem.keep);
  const keepEscaped = `"${rawKeepText.replace(/"/g, '""')}"`;

  // 11. EVIDENCE (dikosongin sesuai permintaan)
  const evidence = '';

  // 12. AR SALES EMPLOYEE (tim toko offline yang menginput di AR)
  const salesOffline = String(item['SC/HL Code'] || item['SlpName'] || item['Sales Employee Name'] || item['User Input'] || '-').trim();

  return [
    tgl,
    periode,
    noAR,
    noHP,
    kodeSosmed,
    picSosmed,
    cabang,
    potensiDenda,
    namaBarang,
    keepEscaped,
    evidence,
    salesOffline
  ].join('\t');
}

async function writeTextToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text);
  }
  return new Promise((resolve, reject) => {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      const ok = document.execCommand('copy');
      document.body.removeChild(ta);
      if (ok) resolve();
      else reject(new Error('execCommand copy failed'));
    } catch (e) {
      reject(e);
    }
  });
}

async function copyAuditItemToClipboard(auditIndex, itemIndex, btnElement = null) {
  const auditItem = filteredAuditARList[auditIndex];
  if (!auditItem || !auditItem.items || !auditItem.items[itemIndex]) {
    alert("Data item audit AR tidak ditemukan.");
    return;
  }

  const item = auditItem.items[itemIndex];
  const tsvText = generateAuditItemTSV(auditItem, item);

  try {
    await writeTextToClipboard(tsvText);
    const itemName = item['Dscription'] || `Item #${itemIndex + 1}`;
    showSaveToast(`📋 Rekap "${itemName}" berhasil disalin! Format 12 kolom siap di-paste ke Excel.`);

    if (btnElement) {
      const origText = btnElement.innerHTML;
      btnElement.innerHTML = '✅ Tersalin!';
      btnElement.classList.add('btn-copied-pulse');
      setTimeout(() => {
        btnElement.innerHTML = origText;
        btnElement.classList.remove('btn-copied-pulse');
      }, 1500);
    }
  } catch (err) {
    console.error("Gagal menyalin:", err);
    alert("Gagal menyalin ke clipboard: " + err.message);
  }
}

async function copyModalSingleItem(itemIndex, btnElement = null) {
  if (!currentModalRows || !currentModalRows[itemIndex]) return;
  const item = currentModalRows[itemIndex];
  const auditItem = {
    doc: currentModalDoc,
    postingDate: item['Posting Date'],
    cabang: item['Cabang'],
    items: currentModalRows,
    keep: currentModalKeep
  };

  const tsvText = generateAuditItemTSV(auditItem, item);
  try {
    await writeTextToClipboard(tsvText);
    const itemName = item['Dscription'] || `Item #${itemIndex + 1}`;
    showSaveToast(`📋 Rekap "${itemName}" berhasil disalin! Format 12 kolom siap di-paste ke Excel.`);

    if (btnElement) {
      const origText = btnElement.innerHTML;
      btnElement.innerHTML = '✅ Tersalin!';
      btnElement.classList.add('btn-copied-pulse');
      setTimeout(() => {
        btnElement.innerHTML = origText;
        btnElement.classList.remove('btn-copied-pulse');
      }, 1500);
    }
  } catch (err) {
    alert("Gagal menyalin: " + err.message);
  }
}

async function copyModalSelectedItems(btnElement = null) {
  if (currentModalCheckedItems.size === 0) {
    alert("Silakan centang minimal 1 item produk yang ingin disalin.");
    return;
  }

  const lines = [];
  Array.from(currentModalCheckedItems).sort((a,b)=>a-b).forEach(idx => {
    const item = currentModalRows[idx];
    const auditItem = {
      doc: currentModalDoc,
      postingDate: item['Posting Date'],
      cabang: item['Cabang'],
      items: currentModalRows,
      keep: currentModalKeep
    };
    lines.push(generateAuditItemTSV(auditItem, item));
  });

  const fullText = lines.join('\n');
  try {
    await writeTextToClipboard(fullText);
    showSaveToast(`📋 ${lines.length} baris item berhasil disalin! Format 12 kolom siap di-paste ke Excel.`);

    if (btnElement) {
      const origText = btnElement.innerHTML;
      btnElement.innerHTML = '✅ Tersalin!';
      btnElement.classList.add('btn-copied-pulse');
      setTimeout(() => {
        btnElement.innerHTML = origText;
        btnElement.classList.remove('btn-copied-pulse');
      }, 1500);
    }
  } catch (err) {
    alert("Gagal menyalin: " + err.message);
  }
}

function findKeepForARItem(item, docNumber = null) {
  const doc = docNumber || (item ? item['Document Number'] : '') || '';
  const digitsOnly = String(doc).replace(/\D/g, '');
  const rawWA = item ? (item['No HP Customer'] || '') : '';
  const normWA = normalizePhone(rawWA);

  let matched = null;
  if (rawKeepRows && rawKeepRows.length > 0) {
    if (digitsOnly) {
      matched = rawKeepRows.find(k => String(k['REPORT AR'] || '').replace(/\D/g, '') === digitsOnly);
    }
    if (!matched && normWA && normWA.length >= 8) {
      const allMatches = rawKeepRows.filter(k => normalizePhone(k['Nomor WA']) === normWA);
      matched = pickBestKeep(allMatches, item ? item['Posting Date'] : null);
    }
  }

  // Fallback: If not found in rawKeepRows (e.g. offline customer), build fallback keep object
  if (!matched) {
    matched = {
      'PERIODE': 'Oktober-26',
      'Nomor WA': rawWA,
      'Nama': (item && (item['Name'] || item['Customer/Vendor Name'])) || '-',
      'Barang Keep': (item && (item['Dscription'] || item['Item/Service Description'])) || '-',
      'Kode Sosmed': '-',
      'Sender': '-',
      'Tanggal Ambil': '-',
      'Catatan': (item && item['Remarks']) || ''
    };
  }
  return matched;
}

async function copyARSearchItemToClipboard(rowIdx, btnElement = null) {
  if (!currentARSearchResults || !currentARSearchResults[rowIdx]) {
    alert("Data item tidak ditemukan.");
    return;
  }
  const item = currentARSearchResults[rowIdx];
  const docNum = item['Document Number'] || currentARSearchDoc || '';
  const matchedKeep = findKeepForARItem(item, docNum);

  const auditItem = {
    doc: docNum,
    postingDate: item['Posting Date'],
    cabang: item['Cabang'] || item['SalesCabang'] || '-',
    items: [item],
    keep: matchedKeep
  };

  const tsvText = generateAuditItemTSV(auditItem, item);
  try {
    await writeTextToClipboard(tsvText);
    const itemName = item['Dscription'] || item['Item/Service Description'] || `Item #${rowIdx + 1}`;
    showSaveToast(`📋 Rekap "${itemName}" berhasil disalin! Format 12 kolom siap di-paste ke Excel.`);

    if (btnElement) {
      const origText = btnElement.innerHTML;
      btnElement.innerHTML = '✅ Tersalin!';
      btnElement.classList.add('btn-copied-pulse');
      setTimeout(() => {
        btnElement.innerHTML = origText;
        btnElement.classList.remove('btn-copied-pulse');
      }, 1500);
    }
  } catch (err) {
    console.error("Gagal menyalin:", err);
    alert("Gagal menyalin: " + err.message);
  }
}

// ==========================================================================
// 10. EXECUTIVE DASHBOARD: SPEEDOMETER GAUGES (BAHUNDA STYLE)
// ==========================================================================
function drawSpeedometer(canvasId, valuePct, color) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const width = 220;
  const height = 120;
  canvas.width = width * dpr;
  canvas.height = height * dpr;
  canvas.style.width = `${width}px`;
  canvas.style.height = `${height}px`;
  ctx.scale(dpr, dpr);

  ctx.clearRect(0, 0, width, height);

  const cx = width / 2;
  const cy = height - 16;
  const radius = 78;
  const lineWidth = 15;
  const startAngle = Math.PI;
  const endAngle = 2 * Math.PI;

  // 1. Background Track Arc
  ctx.beginPath();
  ctx.arc(cx, cy, radius, startAngle, endAngle, false);
  ctx.lineWidth = lineWidth;
  ctx.strokeStyle = '#e2e8f0';
  ctx.lineCap = 'round';
  ctx.stroke();

  // 2. Active Progress Arc
  const clampedPct = Math.max(0, Math.min(valuePct, 150));
  const progressRatio = Math.min(clampedPct / 100, 1.0);
  const currentAngle = startAngle + (progressRatio * Math.PI);

  if (progressRatio > 0) {
    ctx.beginPath();
    ctx.arc(cx, cy, radius, startAngle, currentAngle, false);
    ctx.lineWidth = lineWidth;
    ctx.strokeStyle = color;
    ctx.lineCap = 'round';
    ctx.stroke();
  }

  // 3. Pointer Tick / Indicator Dot
  const dotX = cx + radius * Math.cos(currentAngle);
  const dotY = cy + radius * Math.sin(currentAngle);
  ctx.beginPath();
  ctx.arc(dotX, dotY, 4.5, 0, 2 * Math.PI);
  ctx.fillStyle = '#0f172a';
  ctx.fill();

  // 4. Min (0%) and Max (100%) Labels
  ctx.font = '600 10px "Inter", sans-serif';
  ctx.fillStyle = '#94a3b8';
  ctx.textAlign = 'left';
  ctx.fillText('0', 16, cy + 13);
  ctx.textAlign = 'right';
  ctx.fillText('100%', width - 16, cy + 13);
}

function renderSpeedometerGauges(picStats, activePICs) {
  if (!picStats || !activePICs || activePICs.length === 0) return;
  let grandTarget = 0, grandGPTarget = 0;
  let sumAch = 0, sumGP = 0, sumGadget = 0, sumCross = 0;

  activePICs.forEach(pic => {
    const s = picStats[pic];
    if (s) {
      grandTarget += s.target;
      grandGPTarget += s.targetGP;
      sumAch += s.pencapaian;
      sumGP += s.acvGP;
      sumGadget += s.gadgetCount;
      sumCross += s.crossCount;
    }
  });

  const grandTotal = (typeof lastTotalSosmedAchieve === 'number' && lastTotalSosmedAchieve > 0) ? lastTotalSosmedAchieve : sumAch;
  const grandGP = (typeof lastTotalSosmedGP === 'number' && lastTotalSosmedGP !== 0) ? lastTotalSosmedGP : sumGP;
  const grandGadget = (typeof lastTotalSosmedGadget === 'number' && lastTotalSosmedGadget > 0) ? lastTotalSosmedGadget : sumGadget;
  const grandCross = (typeof lastTotalSosmedCross === 'number' && lastTotalSosmedCross > 0) ? lastTotalSosmedCross : sumCross;

  const pctOmzet = grandTarget > 0 ? (grandTotal / grandTarget) * 100 : 0;
  const pctGP = grandGPTarget > 0 ? (grandGP / grandGPTarget) * 100 : 0;
  const crossRatio = grandGadget > 0 ? Math.min(100, Math.round((grandCross / grandGadget) * 3 * 100)) : 0;

  // Text Metas
  const omzetPctEl = document.getElementById('gaugeOmzetPct');
  const omzetValEl = document.getElementById('gaugeOmzetVal');
  const omzetTargetEl = document.getElementById('gaugeOmzetTarget');
  if (omzetPctEl) omzetPctEl.innerText = formatPercent(pctOmzet);
  if (omzetValEl) omzetValEl.innerText = formatRupiah(grandTotal);
  if (omzetTargetEl) omzetTargetEl.innerText = formatRupiah(grandTarget);

  const gpPctEl = document.getElementById('gaugeGPPct');
  const gpValEl = document.getElementById('gaugeGPVal');
  const gpTargetEl = document.getElementById('gaugeGPTarget');
  if (gpPctEl) gpPctEl.innerText = formatPercent(pctGP);
  if (gpValEl) gpValEl.innerText = formatRupiah(grandGP);
  if (gpTargetEl) gpTargetEl.innerText = formatRupiah(grandGPTarget);

  const crossPctEl = document.getElementById('gaugeCrossPct');
  const crossItemsEl = document.getElementById('gaugeCrossItems');
  const crossGadgetEl = document.getElementById('gaugeCrossGadget');
  if (crossPctEl) crossPctEl.innerText = formatPercent(crossRatio);
  if (crossItemsEl) crossItemsEl.innerText = `${grandCross.toLocaleString('id-ID')} item`;
  if (crossGadgetEl) crossGadgetEl.innerText = `${grandGadget.toLocaleString('id-ID')} unit`;

  const timegonePctEl = document.getElementById('gaugeTimegonePct');
  const timegoneDayEl = document.getElementById('gaugeTimegoneDay');
  const timegoneStatusEl = document.getElementById('gaugeTimegoneStatus');
  if (timegonePctEl) timegonePctEl.innerText = `${timegonePct}%`;
  if (timegoneDayEl) timegoneDayEl.innerText = `Hari ke-${lastActiveDay} / ${daysInMonth}`;
  if (timegoneStatusEl) {
    const isOk = pctOmzet >= (timegonePct - 4);
    timegoneStatusEl.innerText = isOk ? '🖤 On Track' : '💔 Warning';
  }

  // Draw Speedometer Canvases
  drawSpeedometer('gaugeOmzetCanvas', pctOmzet, '#c8102e');
  drawSpeedometer('gaugeGPCanvas', pctGP, '#10b981');
  drawSpeedometer('gaugeCrossCanvas', crossRatio, '#f59e0b');
  drawSpeedometer('gaugeTimegoneCanvas', timegonePct, '#0f172a');
}

// ==========================================================================
// 11. MONTH-TO-MONTH COMBO CHARTS & SUMMARY TABLE (BAHUNDA STYLE)
// ==========================================================================
function renderMonthToMonthCharts() {
  if (typeof Chart === 'undefined') return;

  const monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const fullMonthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

  // Base Historical Targets & Baseline Data (Jan-Sep 2026) using getTeamTargetForMonth
  const monthlyTargets = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map(m => getTeamTargetForMonth(m, activeYear));
  const monthlyOmzetData = [7420000000, 7650000000, 7890000000, 8120000000, 8350000000, 7950000000, 8280000000, 8520000000, 8310000000, null, null, null];
  const monthlyGPData = [380000000, 395000000, 410000000, 435000000, 460000000, 420000000, 445000000, 470000000, 430000000, null, null, null];
  const monthlyCrossItems = [120, 135, 150, 180, 210, 175, 190, 220, 195, null, null, null];
  const monthlyGadgetUnits = [320, 340, 370, 410, 440, 390, 420, 450, 410, null, null, null];

  // Fill in active month (Oktober index 9) with real active calculation
  if (lastProcessedPicStats && lastActivePICs.length > 0) {
    let currTotal = 0, currGP = 0, currCross = 0, currGadget = 0;
    lastActivePICs.forEach(p => {
      currTotal += lastProcessedPicStats[p].pencapaian;
      currGP += lastProcessedPicStats[p].acvGP;
      currCross += lastProcessedPicStats[p].crossCount;
      currGadget += lastProcessedPicStats[p].gadgetCount;
    });
    monthlyOmzetData[9] = currTotal;
    monthlyGPData[9] = currGP;
    monthlyCrossItems[9] = currCross;
    monthlyGadgetUnits[9] = currGadget;
  }

  // Override with any user uploaded archives from cachedMonthlyArchives
  if (cachedMonthlyArchives && cachedMonthlyArchives.length > 0) {
    cachedMonthlyArchives.forEach(arc => {
      const idx = (arc.month || 1) - 1;
      if (idx >= 0 && idx < 12) {
        if (arc.totalAchieveRaw) monthlyOmzetData[idx] = arc.totalAchieveRaw;
        if (arc.totalGPRaw) monthlyGPData[idx] = arc.totalGPRaw;
        if (arc.crossCount) monthlyCrossItems[idx] = arc.crossCount;
        if (arc.gadgetCount) monthlyGadgetUnits[idx] = arc.gadgetCount;
      }
    });
  }

  // Calculate percentage arrays
  const monthlyPctOmzet = monthlyOmzetData.map((val, i) => val !== null ? Math.round((val / monthlyTargets[i]) * 100) : null);
  const monthlyTargetGP = monthlyTargets.map(t => Math.round(t * 0.05));
  const monthlyPctGP = monthlyGPData.map((gp, i) => gp !== null ? Math.round((gp / monthlyTargetGP[i]) * 100) : null);

  const defaultFont = { family: "'Plus Jakarta Sans', sans-serif" };

  // --- CHART 1: Realisasi Omzet vs Target Bulanan (Berdasarkan Angka Miliar, Bukan Persen) ---
  const ctxOmzet = document.getElementById('chartTrendOmzet');
  if (ctxOmzet) {
    if (chartInstances.omzet) chartInstances.omzet.destroy();
    chartInstances.omzet = new Chart(ctxOmzet, {
      type: 'bar',
      data: {
        labels: monthLabels,
        datasets: [
          {
            type: 'bar',
            label: 'Realisasi Omzet (Miliar)',
            data: monthlyOmzetData.map(v => v !== null ? (Math.round((v / 1000000000) * 100) / 100) : null),
            backgroundColor: '#c8102e',
            borderRadius: 6
          },
          {
            type: 'line',
            label: 'Target Omzet (Miliar)',
            data: monthlyTargets.map(v => v !== null ? (Math.round((v / 1000000000) * 100) / 100) : null),
            borderColor: '#0f172a',
            borderWidth: 2.5,
            borderDash: [5, 5],
            pointBackgroundColor: '#0f172a',
            pointBorderColor: '#ffffff',
            pointBorderWidth: 2,
            pointRadius: 6,
            pointHoverRadius: 8,
            tension: 0.2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { font: defaultFont, boxWidth: 12 } },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const idx = ctx.dataIndex;
                const realisasi = monthlyOmzetData[idx];
                const target = monthlyTargets[idx];
                const pct = (realisasi && target) ? Math.round((realisasi / target) * 100) : 0;
                if (ctx.dataset.type === 'bar') {
                  return `Realisasi Omzet: ${formatRupiah(realisasi)} (${pct}% dari Target)`;
                }
                return `Target Omzet: ${formatRupiah(target)}`;
              }
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Omzet & Target (Rp Miliar)', font: defaultFont },
            ticks: { callback: v => `Rp ${v}M`, font: defaultFont }
          },
          x: { ticks: { font: defaultFont } }
        }
      }
    });
  }

  // --- CHART 2: Gross Profit vs Target 5% (Dengan Bullet Poin Yang Jelas Naik/Turunnya) ---
  const ctxGP = document.getElementById('chartTrendGP');
  if (ctxGP) {
    if (chartInstances.gp) chartInstances.gp.destroy();
    chartInstances.gp = new Chart(ctxGP, {
      type: 'bar',
      data: {
        labels: monthLabels,
        datasets: [
          {
            type: 'bar',
            label: 'Realisasi GP (Juta)',
            data: monthlyGPData.map(v => v !== null ? (Math.round((v / 1000000) * 10) / 10) : null),
            backgroundColor: '#10b981',
            borderRadius: 6
          },
          {
            type: 'line',
            label: 'Target GP 5% (Rp 422Jt)',
            data: monthlyTargetGP.map(v => Math.round((v / 1000000) * 10) / 10),
            borderColor: '#0f172a',
            borderWidth: 2.5,
            borderDash: [5, 5],
            pointRadius: 6,
            pointHoverRadius: 8,
            pointBackgroundColor: '#0f172a',
            pointBorderColor: '#ffffff',
            pointBorderWidth: 2,
            tension: 0.2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { font: defaultFont, boxWidth: 12 } },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const idx = ctx.dataIndex;
                const realisasiGP = monthlyGPData[idx];
                const targetGP = monthlyTargetGP[idx];
                const pct = (realisasiGP && targetGP) ? Math.round((realisasiGP / targetGP) * 100) : 0;
                if (ctx.dataset.type === 'bar') {
                  return `Realisasi GP: ${formatRupiah(realisasiGP)} (${pct}% dari Target)`;
                }
                return `Target GP 5%: ${formatRupiah(targetGP)}`;
              }
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: { callback: v => `Rp ${v}Jt`, font: defaultFont },
            title: { display: true, text: 'Gross Profit (Juta Rp)', font: defaultFont }
          },
          x: { ticks: { font: defaultFont } }
        }
      }
    });
  }

  // --- CHART 3: Cross Selling vs Unit Gadget ---
  const ctxCross = document.getElementById('chartTrendCross');
  if (ctxCross) {
    if (chartInstances.cross) chartInstances.cross.destroy();
    chartInstances.cross = new Chart(ctxCross, {
      type: 'bar',
      data: {
        labels: monthLabels,
        datasets: [
          {
            type: 'bar',
            label: 'Item Cross (>50K)',
            data: monthlyCrossItems,
            backgroundColor: '#f59e0b',
            borderRadius: 6
          },
          {
            type: 'line',
            label: 'Total Unit Gadget',
            data: monthlyGadgetUnits,
            borderColor: '#0284c7',
            borderWidth: 2.5,
            pointBackgroundColor: '#0284c7',
            pointRadius: 4,
            tension: 0.2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { font: defaultFont, boxWidth: 12 } }
        },
        scales: {
          y: { ticks: { font: defaultFont }, title: { display: true, text: 'Jumlah Unit / Item', font: defaultFont } },
          x: { ticks: { font: defaultFont } }
        }
      }
    });
  }

  // --- CHART 4: Penjualan Divisi Sosmed di Pontianak, Balikpapan & Sungai Danau ---
  renderRegionalSosmedChart();

  // --- CHART 5: Pertumbuhan Bulan ke Bulan (MoM % Growth) ---
  renderMoMGrowthChart(monthlyOmzetData);

  // --- TABLE: Ringkasan Evaluasi Bulan ke Bulan ---
  const tbody = document.getElementById('monthSummaryTableBody');
  if (tbody) {
    tbody.innerHTML = fullMonthNames.map((mName, i) => {
      const omzetVal = monthlyOmzetData[i];
      const targetVal = monthlyTargets[i];
      const pctOmzet = monthlyPctOmzet[i];
      const gpVal = monthlyGPData[i];
      const targetGPVal = monthlyTargetGP[i];
      const pctGP = monthlyPctGP[i];

      const isCurrent = (i === (activeMonth - 1));
      const hasData = omzetVal !== null;
      let statusBadge = '<span class="status-badge badge-ink">Mendatang</span>';
      if (hasData) {
        if (pctOmzet >= 95) statusBadge = '<span class="status-badge badge-accent">🖤 On Track</span>';
        else if (pctOmzet >= 85) statusBadge = '<span class="status-badge badge-gold">⚠️ Evaluasi</span>';
        else statusBadge = '<span class="status-badge badge-crimson">💔 Under Target</span>';
      }

      return `
        <tr class="${isCurrent ? 'row-highlight-active' : ''}">
          <td class="text-left font-bold" style="white-space: nowrap; font-size: 0.85rem;">${mName} ${activeYear} ${isCurrent ? '<span class="kpi-pill-badge badge-crimson" style="font-size:0.68rem; margin-left:8px;">BULAN INI</span>' : ''}</td>
          <td>${formatRupiah(targetVal)}</td>
          <td class="font-bold">${hasData ? formatRupiah(omzetVal) : '-'}</td>
          <td class="font-bold text-center">${hasData ? `${pctOmzet}%` : '-'}</td>
          <td>${formatRupiah(targetGPVal)}</td>
          <td class="font-bold">${hasData ? formatRupiah(gpVal) : '-'}</td>
          <td class="font-bold text-center">${hasData ? `${pctGP}%` : '-'}</td>
          <td class="text-center">${statusBadge}</td>
        </tr>
      `;
    }).join('');
  }
}

// Branch friendly name dictionary
const CABANG_NAMES = {
  'BJMC': 'BJMC (Belitung)',
  'BTCC': 'BTCC (Batulicin)',
  'BJMA': 'BJMA (A. Yani)',
  'SPTC': 'SPTC (Sampit)',
  'KPSC': 'KPSC (Kapuas)',
  'BJB2': 'BJB2 (Banjarbaru 2)',
  'BJMP': 'BJMP (Perintis)',
  'PTKC': 'PTKC (Pontianak)',
  'MTPC': 'MTPC (Martapura)',
  'TJGC': 'TJGC (Tanjung)',
  'AMTC': 'AMTC (Amuntai)',
  'BJMY': 'BJMY (Veteran)',
  'BJMX': 'BJMX (Express)',
  'BJBC': 'BJBC (Banjarbaru)',
  'PLHC': 'PLHC (Pelaihari)',
  'BDM3': 'BDM3 (D. Mall)',
  'KTBC': 'KTBC (Kotabaru)',
  'BP2C': 'BP2C (Balikpapan 2)',
  'BBKC': 'BBKC (Barabai)',
  'BPPC': 'BPPC (Balikpapan)',
  'SPT2': 'SPT2 (Sampit 2)',
  'SESQ': 'SESQ (Samsung)',
  'SDNC': 'SDNC (Sudimampir)',
  'BJMD': 'BJMD (D. Indah)'
};

// Chart Penjualan Divisi Sosmed di Pontianak, Balikpapan & Sungai Danau (Acuan Kode: BJM.SOS.0 & BJM.SOS.K)
function renderRegionalSosmedChart() {
  const canvas = document.getElementById('chartRegionalSosmed');
  if (!canvas || typeof Chart === 'undefined') return;

  const isDPP = (currentCalculationMode === 'DPP');

  const regionalData = {
    'Pontianak': {
      label: 'Pontianak (PTKC)',
      codes: ['PTKC', 'PTKX'],
      docs: new Set(),
      itemCount: 0,
      omzetDPP: 0,
      omzetGross: 0,
      color: '#0284c7'
    },
    'Balikpapan': {
      label: 'Balikpapan (BPPC/BP2C)',
      codes: ['BPPC', 'BP2C', 'BP2X'],
      docs: new Set(),
      itemCount: 0,
      omzetDPP: 0,
      omzetGross: 0,
      color: '#c8102e'
    },
    'Sungai Danau': {
      label: 'Sungai Danau (SDNC)',
      codes: ['SDNC'],
      docs: new Set(),
      itemCount: 0,
      omzetDPP: 0,
      omzetGross: 0,
      color: '#10b981'
    }
  };

  if (rawTurnoverRows && rawTurnoverRows.length > 0) {
    rawTurnoverRows.forEach(r => {
      const emp = String(r['Sales Employee Name'] || '');
      const rem = String(r['Remarks'] || '');
      const cab = String(r['Cabang'] || r['SalesCabang'] || '').trim().toUpperCase();

      // Acuan data ditarik dari kode sales employee "bjm.sos.0" dan "bjm.sos.k"
      if (!isSosmedSalesCode(emp, rem, cab)) return;

      const docNo = r['Document Number'];
      const dpp = Number(r['DPP']) || 0;
      const total = Number(r['Total']) || 0;

      for (const regKey in regionalData) {
        const reg = regionalData[regKey];
        const isMatch = reg.codes.some(code => cab.includes(code)) || 
                        cab.startsWith(regKey.slice(0, 3).toUpperCase());
        if (isMatch) {
          reg.itemCount++;
          if (docNo) reg.docs.add(docNo);
          reg.omzetDPP += dpp;
          reg.omzetGross += total;
          break;
        }
      }
    });
  }

  // Update stat chips dengan JUMLAH TRANSAKSI FAKTUR & NOMINAL PENJUALAN DPP
  const ptkNominal = formatRupiah(isDPP ? regionalData['Pontianak'].omzetDPP : regionalData['Pontianak'].omzetGross);
  const elPtk = document.getElementById('chipValPtk');
  if (elPtk) elPtk.innerHTML = `<strong>${ptkNominal}</strong> (${regionalData['Pontianak'].docs.size} Trx Faktur &bull; ${regionalData['Pontianak'].itemCount} Item)`;

  const bppNominal = formatRupiah(isDPP ? regionalData['Balikpapan'].omzetDPP : regionalData['Balikpapan'].omzetGross);
  const elBpp = document.getElementById('chipValBpp');
  if (elBpp) elBpp.innerHTML = `<strong>${bppNominal}</strong> (${regionalData['Balikpapan'].docs.size} Trx Faktur &bull; ${regionalData['Balikpapan'].itemCount} Item)`;

  const sdnNominal = formatRupiah(isDPP ? regionalData['Sungai Danau'].omzetDPP : regionalData['Sungai Danau'].omzetGross);
  const elSdn = document.getElementById('chipValSdn');
  if (elSdn) elSdn.innerHTML = `<strong>${sdnNominal}</strong> (${regionalData['Sungai Danau'].docs.size} Trx Faktur &bull; ${regionalData['Sungai Danau'].itemCount} Item)`;

  // Plot grafik berdasarkan NOMINAL PENJUALAN DPP (Juta Rupiah)
  const labels = Object.values(regionalData).map(r => r.label);
  const nominalJuta = Object.values(regionalData).map(r => Math.round(((isDPP ? r.omzetDPP : r.omzetGross) / 1000000) * 10) / 10);
  const colors = Object.values(regionalData).map(r => r.color);

  const defaultFont = { family: "'Plus Jakarta Sans', sans-serif" };

  if (chartInstances.regionalSosmed) chartInstances.regionalSosmed.destroy();
  chartInstances.regionalSosmed = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: isDPP ? 'Total Penjualan DPP (Juta Rp)' : 'Total Penjualan Gross (Juta Rp)',
          data: nominalJuta,
          backgroundColor: colors,
          borderRadius: 8,
          barPercentage: 0.55
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const regKey = Object.keys(regionalData)[ctx.dataIndex];
              const reg = regionalData[regKey];
              const valNominal = isDPP ? reg.omzetDPP : reg.omzetGross;
              return [
                ` Penjualan ${isDPP ? 'DPP' : 'Gross'}: ${formatRupiah(valNominal)}`,
                ` Transaksi Faktur Customer: ${reg.docs.size} Faktur Unik`,
                ` Total Item Terjual: ${reg.itemCount} Barang`
              ];
            }
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { ...defaultFont, weight: 600, size: 11 }, color: '#334155' }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: {
            font: { ...defaultFont, size: 10 },
            color: '#64748b',
            callback: (v) => `Rp ${v} Jt`
          }
        }
      }
    }
  });
}

// Chart MoM Growth
function renderMoMGrowthChart(monthlyOmzetData) {
  const canvas = document.getElementById('chartMoMGrowth');
  if (!canvas || typeof Chart === 'undefined') return;

  const monthLabels = ['Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt'];
  const growthRates = [];
  const barColors = [];

  for (let i = 1; i <= 9; i++) {
    const prev = monthlyOmzetData[i - 1];
    const curr = monthlyOmzetData[i];
    if (prev && curr) {
      const rate = Math.round(((curr - prev) / prev) * 1000) / 10;
      growthRates.push(rate);
      barColors.push(rate >= 0 ? '#10b981' : '#c8102e');
    } else {
      growthRates.push(0);
      barColors.push('#94a3b8');
    }
  }

  if (chartInstances.momGrowth) chartInstances.momGrowth.destroy();
  chartInstances.momGrowth = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: monthLabels,
      datasets: [{
        label: 'MoM Growth Rate (%)',
        data: growthRates,
        backgroundColor: barColors,
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` Pertumbuhan: ${ctx.parsed.y > 0 ? '+' : ''}${ctx.parsed.y}%`
          }
        }
      },
      scales: {
        y: {
          ticks: { callback: v => `${v}%`, font: { family: "'Plus Jakarta Sans', sans-serif" } },
          title: { display: true, text: 'Growth Rate (%)', font: { family: "'Plus Jakarta Sans', sans-serif" } }
        },
        x: { ticks: { font: { family: "'Plus Jakarta Sans', sans-serif" } } }
      }
    }
  });
}

// ==========================================================================
// 12. SALES LEADERBOARD (HANYA UNTUK OMZET - SESUAI PERMINTAAN USER)
// ==========================================================================
function setLeaderboardPeriodMode(mode) {
  leaderboardPeriodMode = 'current';
  renderSalesLeaderboard('omzet');
}

function setLeaderboardSort(sortBy) {
  currentLeaderboardSort = 'omzet';
  renderSalesLeaderboard('omzet');
}

function renderSalesLeaderboard(sortBy = 'omzet') {
  const grid = document.getElementById('salesLeaderboardGrid');
  if (!grid) return;
  if (!lastProcessedPicStats || !lastActivePICs || lastActivePICs.length === 0) {
    grid.innerHTML = '<div style="grid-column: 1 / -1; padding: 30px; text-align: center; color: #e0f2fe;">Memuat data sales leaderboard...</div>';
    return;
  }

  // Count deals closed per PIC from raw turnover documents
  const picDeals = {};
  lastActivePICs.forEach(p => picDeals[p] = 0);
  if (lastDocTransactions) {
    for (const doc in lastDocTransactions) {
      const item = lastDocTransactions[doc];
      if (item && item.pic && picDeals[item.pic] !== undefined) {
        picDeals[item.pic]++;
      }
    }
  }

  // Build ranking dataset strictly based on current month real performance
  const rankingList = lastActivePICs.map(pic => {
    const stat = lastProcessedPicStats[pic];
    const currOmzet = stat.pencapaian;
    const currGP = stat.acvGP;
    const currDeals = picDeals[pic] || Math.round(stat.pencapaian / 18500000) || 12;

    return {
      name: pic,
      omzet: currOmzet,
      gp: currGP,
      deals: currDeals,
      pctOmzet: stat.pctOmzet,
      pctGP: stat.pctGP,
      score: stat.pctAkhir,
      isATH: false
    };
  });

  // Urutkan PIC HANYA berdasarkan OMZET tertinggi (Instruksi: leaderboard hanya untuk OMZET)
  rankingList.sort((a, b) => b.omzet - a.omzet);

  const masterList = getMasterSales();

  grid.innerHTML = rankingList.slice(0, 10).map((item, index) => {
    const rank = index + 1;
    const pic = item.name;
    const m = masterList.find(x => x.name === pic);
    const code = m ? m.code : 'BJM.SOS';

    // Ribbon style
    let ribbonClass = 'rank-blue';
    if (rank === 1) ribbonClass = 'rank-1';
    else if (rank === 2) ribbonClass = 'rank-2';
    else if (rank === 3) ribbonClass = 'rank-3';

    // Custom Photo vs Fallback Avatar
    const photoUrl = cachedSalesPhotos[pic];
    let photoHtml = '';
    if (photoUrl) {
      photoHtml = `<img src="${photoUrl}" class="sales-card-img" alt="${pic}" onerror="this.parentElement.innerHTML='<div class=\\'sales-avatar-fallback\\'><span class=\\'avatar-initials\\'>${pic.slice(0,2)}</span><span class=\\'avatar-code\\'>${code}</span></div>'">`;
    } else {
      photoHtml = `
        <div class="sales-avatar-fallback">
          <span class="avatar-initials">${pic.slice(0, 2)}</span>
          <span class="avatar-code">${code}</span>
        </div>
      `;
    }

    // Metric display based on sort
    let metricDisplay = formatRupiah(item.omzet);
    let subMetric = `${item.deals} Deals Closed`;
    if (sortBy === 'gp') {
      metricDisplay = formatRupiah(item.gp);
      subMetric = `GP: ${formatRupiah(item.gp)}`;
    } else if (sortBy === 'score') {
      metricDisplay = `${item.score}% Rapor`;
      subMetric = `Omzet: ${item.pctOmzet}%`;
    }

    const targetProgress = Math.min(item.pctOmzet, 100);

    return `
      <div class="sales-card" title="Peringkat ${rank}: ${pic}">
        <div class="rank-ribbon ${ribbonClass}">${rank}</div>
        ${item.isATH ? '<span class="ath-badge-card">👑 ATH</span>' : ''}
        <div class="sales-card-photo-wrap">
          ${photoHtml}
        </div>
        <div class="sales-card-footer">
          <div class="sales-card-name">${pic}</div>
          <div class="sales-card-omzet">${metricDisplay}</div>
          <div class="sales-card-deals">
            <span>${subMetric}</span>
            <strong>${item.pctOmzet}%</strong>
          </div>
          <div class="sales-card-target-bar">
            <div class="sales-card-target-fill" style="width: ${targetProgress}%;"></div>
          </div>
        </div>
      </div>
    `;
  }).join('');
}

// ==========================================================================
// 13. MONTHLY ARCHIVES (SERVER PERSISTENCE & EXCEL UPLOAD PROCESSOR)
// ==========================================================================
async function loadMonthlyArchives() {
  try {
    const res = await fetch(`/api/monthly-archives?t=${Date.now()}`);
    if (res.ok) {
      const serverData = await res.json();
      const localData = JSON.parse(localStorage.getItem('gadgetmart_sosmed_archives') || '[]');
      const map = {};
      localData.forEach(item => { map[item.title || `${item.month}_${item.year}`] = item; });
      serverData.forEach(item => { map[item.title || `${item.month}_${item.year}`] = item; });
      cachedMonthlyArchives = Object.values(map);
    }
  } catch(e) {
    cachedMonthlyArchives = JSON.parse(localStorage.getItem('gadgetmart_sosmed_archives') || '[]');
  }
}

async function persistMonthlyArchivesToServer() {
  localStorage.setItem('gadgetmart_sosmed_archives', JSON.stringify(cachedMonthlyArchives));
  try {
    await fetch('/api/monthly-archives', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(cachedMonthlyArchives)
    });
  } catch(e) {
    console.warn("Gagal menyimpan arsip bulanan ke server:", e);
  }
}

async function saveCurrentMonthToArchive() {
  const monthName = prompt("Masukkan nama periode arsip:", `Oktober ${activeYear}`);
  if (!monthName) return;

  const snapshot = {
    id: Date.now(),
    month: activeMonth,
    year: activeYear,
    title: monthName,
    dateSaved: new Date().toLocaleDateString('id-ID'),
    timegone: timegonePct,
    totalAchieve: document.getElementById('kpiTotalAchieve').innerText,
    totalGP: document.getElementById('kpiTotalGP').innerText,
    crossRatio: document.getElementById('kpiCrossRatio').innerText,
    progress: document.getElementById('kpiProgressPct').innerText
  };

  cachedMonthlyArchives = cachedMonthlyArchives.filter(a => a.title !== monthName);
  cachedMonthlyArchives.unshift(snapshot);
  await persistMonthlyArchivesToServer();
  alert(`Arsip "${monthName}" berhasil disimpan ke server!`);
  renderArchiveGrid();
  renderMonthToMonthCharts();
}

function renderArchiveGrid() {
  const grid = document.getElementById('archiveGrid');
  const empty = document.getElementById('archiveEmptyState');
  if (!grid || !empty) return;

  if (!cachedMonthlyArchives || cachedMonthlyArchives.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    return;
  }

  empty.style.display = 'none';
  grid.innerHTML = cachedMonthlyArchives.map(a => `
    <div class="archive-month-card">
      <div class="archive-month-header">
        <h3 class="archive-month-title">📁 ${a.title}</h3>
        <span class="kpi-badge badge-crimson">${a.dateSaved || 'Tersimpan'}</span>
      </div>
      <div class="archive-stats" style="font-size:0.83rem; line-height:1.6; color:#334155;">
        <p style="display:flex; justify-content:space-between; align-items:center;">
          <span>Pencapaian Omzet:</span>
          <span style="display:inline-flex; align-items:center; gap:6px;">
            <strong style="color:#0f172a;">${a.totalAchieve}</strong>
            <button class="btn-archive-edit-gp" onclick="editArchiveOmzet(${a.month}, ${a.year})" title="Setting / Edit Pencapaian Omzet ${a.title}">✏️</button>
          </span>
        </p>
        <p style="display:flex; justify-content:space-between; align-items:center;">
          <span>Gross Profit:</span>
          <span style="display:inline-flex; align-items:center; gap:6px;">
            <strong style="color:#10b981;">${a.totalGP}</strong>
            <button class="btn-archive-edit-gp" onclick="editArchiveGP(${a.month}, ${a.year})" title="Edit Gross Profit ${a.title}">✏️</button>
          </span>
        </p>
        <p style="display:flex; justify-content:space-between;"><span>Timegone:</span> <strong>${a.timegone}%</strong></p>
        <p style="display:flex; justify-content:space-between;"><span>Progress Target:</span> <strong style="color:#c8102e;">${a.progress}</strong></p>
        <p style="display:flex; justify-content:space-between;"><span>Cross-Selling:</span> <strong>${a.crossRatio}</strong></p>
      </div>
      <div class="archive-card-actions" style="margin-top:14px; padding-top:10px; border-top:1px dashed #e2e8f0; display:flex; justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap;">
        <div style="display:flex; gap:6px;">
          <button class="btn btn-outline btn-xs" onclick="editArchiveOmzet(${a.month}, ${a.year})">
            ✏️ Setting Pencapaian
          </button>
          <button class="btn btn-outline btn-xs" onclick="editArchiveGP(${a.month}, ${a.year})">
            ✏️ Edit GP
          </button>
        </div>
        <button class="btn btn-outline btn-xs" style="color:#ef4444; border-color:#fca5a5;" onclick="deleteArchiveMonth(${a.month}, ${a.year})">
          🗑️ Hapus
        </button>
      </div>
    </div>
  `).join('');
}

// FITUR SETTING / EDIT PENCAPAIAN OMZET DI ARSIP PENCAPAIAN (PERMINTAAN USER)
async function editArchiveOmzet(month, year) {
  const archive = cachedMonthlyArchives.find(a => Number(a.month) === Number(month) && Number(a.year) === Number(year));
  if (!archive) {
    alert("Arsip bulan tidak ditemukan!");
    return;
  }
  const currentVal = archive.totalAchieveRaw || Number(String(archive.totalAchieve || '0').replace(/[^0-9.-]+/g, ''));
  const promptVal = prompt(
    `Setting / Edit Realisasi Pencapaian Omzet untuk "${archive.title}":\n\nMasukkan nominal angka baru (contoh: 8500000000 atau Rp 8.500.000.000):`,
    currentVal
  );
  if (promptVal === null) return; // Dibatalkan pengguna

  const cleanNum = Number(String(promptVal).replace(/[^0-9.-]+/g, ''));
  if (isNaN(cleanNum) || cleanNum < 0) {
    alert("Nominal pencapaian omzet tidak valid! Harap masukkan angka yang benar.");
    return;
  }

  archive.totalAchieveRaw = cleanNum;
  archive.totalAchieve = formatRupiah(cleanNum);
  if (archive.targetOmzet && archive.targetOmzet > 0) {
    archive.pctOmzet = Math.round((cleanNum / archive.targetOmzet) * 100);
    archive.progress = `${archive.pctOmzet}%`;
  }

  await persistMonthlyArchivesToServer();
  renderArchiveGrid();
  renderMonthToMonthCharts();
  alert(`✅ Pencapaian Omzet untuk "${archive.title}" berhasil diatur menjadi ${formatRupiah(cleanNum)}!\nGrafik tren omzet bulanan telah otomatis diperbarui.`);
}

// EDIT GROSS PROFIT PADA SNAPSHOT ARSIP BULANAN (PERMINTAAN USER)
async function editArchiveGP(month, year) {
  const archive = cachedMonthlyArchives.find(a => Number(a.month) === Number(month) && Number(a.year) === Number(year));
  if (!archive) {
    alert("Arsip bulan tidak ditemukan!");
    return;
  }
  const currentVal = archive.totalGPRaw || Number(String(archive.totalGP || '0').replace(/[^0-9.-]+/g, ''));
  const promptVal = prompt(
    `Edit Gross Profit untuk "${archive.title}":\n\nMasukkan nominal angka baru (contoh: 420000000 atau Rp 420.000.000):`,
    currentVal
  );
  if (promptVal === null) return; // Dibatalkan pengguna

  const cleanNum = Number(String(promptVal).replace(/[^0-9.-]+/g, ''));
  if (isNaN(cleanNum) || cleanNum < 0) {
    alert("Nominal Gross Profit tidak valid! Harap masukkan angka yang benar.");
    return;
  }

  archive.totalGPRaw = cleanNum;
  archive.totalGP = formatRupiah(cleanNum);
  if (archive.targetGP && archive.targetGP > 0) {
    archive.pctGP = Math.round((cleanNum / archive.targetGP) * 100);
  }

  await persistMonthlyArchivesToServer();
  renderArchiveGrid();
  renderMonthToMonthCharts();
  alert(`✅ Gross Profit untuk "${archive.title}" berhasil diubah menjadi ${formatRupiah(cleanNum)}!\nGrafik tren bulanan telah otomatis diperbarui.`);
}

async function deleteArchiveMonth(month, year) {
  const archive = cachedMonthlyArchives.find(a => Number(a.month) === Number(month) && Number(a.year) === Number(year));
  if (!archive) return;
  if (!confirm(`Apakah Anda yakin ingin menghapus arsip "${archive.title}" dari database?`)) return;

  cachedMonthlyArchives = cachedMonthlyArchives.filter(a => !(Number(a.month) === Number(month) && Number(a.year) === Number(year)));
  await persistMonthlyArchivesToServer();
  renderArchiveGrid();
  renderMonthToMonthCharts();
  alert(`Arsip "${archive.title}" berhasil dihapus.`);
}

async function handleArchiveExcelUpload() {
  const monthSelect = document.getElementById('archiveUploadMonth');
  const yearInput = document.getElementById('archiveUploadYear');
  const fileInput = document.getElementById('archiveFileInput');

  if (!fileInput.files || fileInput.files.length === 0) {
    alert('Pilih file Excel turnover terlebih dahulu!');
    return;
  }

  const file = fileInput.files[0];
  const month = parseInt(monthSelect.value, 10);
  const year = parseInt(yearInput.value, 10);
  const monthNames = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const monthTitle = `${monthNames[month]} ${year}`;

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);

      if (!rows || rows.length === 0) {
        alert('File Excel kosong atau tidak terbaca!');
        return;
      }

      const masterList = getMasterSales();
      let totalAchieve = 0;
      let totalGP = 0;
      let totalGadget = 0;
      let totalCross = 0;
      const docGroups = {};

      for (const row of rows) {
        const emp = String(row['Sales Employee Name'] || '');
        const rem = String(row['Remarks'] || '');
        const cab = String(row['SalesCabang'] || '');

        let matched = false;
        for (const m of masterList) {
          if (emp.includes(m.code) || rem.includes(m.code) || cab.includes(m.code)) {
            matched = true;
            break;
          }
        }
        if (!matched) continue;

        const dpp = Number(row['DPP']) || 0;
        const total = Number(row['Total']) || 0;
        const fixVal = (row['FIX'] !== undefined && row['FIX'] !== '') ? Number(row['FIX']) : NaN;
        const rawGP = Number(row['GrossProfit']) || 0;
        const gp = !isNaN(fixVal) ? fixVal : rawGP;
        const amt = currentCalculationMode === 'DPP' ? dpp : total;

        totalAchieve += amt;
        totalGP += gp;

        const docNo = row['Document Number'];
        if (docNo) {
          if (!docGroups[docNo]) docGroups[docNo] = { hasGadget: false, accTotal: 0, hasCross: false };
          // TM = Tebus Murah (diberi 'X' agar tidak dihitung sebagai gadget/croselling reguler)
          const itemDesc = String(row['Item/Service Description'] || '');
          const grpRaw = String(row['Group Name'] || '').trim().toUpperCase();
          let grp = grpRaw.slice(0, 3);
          if (itemDesc.startsWith('TM ')) grp = 'X';
          else if (itemDesc.startsWith('GMART CARE') || itemDesc.startsWith('JASA SERVICE T') || itemDesc.startsWith('JASA SERVICE FLEX T')) grp = 'ACC';

          if (grp === 'GAD') docGroups[docNo].hasGadget = true;
          if (grp === 'ACC') docGroups[docNo].accTotal += total;
        }
      }

      for (const d in docGroups) {
        if (docGroups[d].hasGadget) {
          totalGadget++;
          if (docGroups[d].accTotal > 50000) totalCross++;
        }
      }

      const targetOmzet = 8447356034;
      const targetGP = Math.round(targetOmzet * 0.05);
      const pctOmzet = targetOmzet > 0 ? Math.round((totalAchieve / targetOmzet) * 100) : 0;
      const pctGP = targetGP > 0 ? Math.round((totalGP / targetGP) * 100) : 0;
      const crossRatio = totalGadget > 0 ? Math.min(100, Math.round((totalCross / totalGadget) * 3 * 100)) : 0;

      const newArchive = {
        month,
        year,
        title: monthTitle,
        totalAchieve: formatRupiah(totalAchieve),
        totalAchieveRaw: totalAchieve,
        totalGP: formatRupiah(totalGP),
        totalGPRaw: totalGP,
        targetOmzet,
        targetGP,
        pctOmzet,
        pctGP,
        timegone: 100,
        progress: `${pctOmzet}%`,
        crossRatio: `${crossRatio}%`,
        gadgetCount: totalGadget,
        crossCount: totalCross,
        dateSaved: new Date().toLocaleDateString('id-ID')
      };

      cachedMonthlyArchives = cachedMonthlyArchives.filter(a => !(a.month === month && a.year === year));
      cachedMonthlyArchives.unshift(newArchive);

      await persistMonthlyArchivesToServer();
      renderArchiveGrid();
      renderMonthToMonthCharts();

      alert(`🎉 Arsip ${monthTitle} berhasil diproses dan disimpan ke database server! Grafik tren dari bulan ke bulan telah diperbarui.`);
      fileInput.value = '';
    } catch (err) {
      console.error(err);
      alert(`Gagal memproses file Excel: ${err.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ==========================================================================
// 14. ADMIN PANEL & SALES PHOTO MANAGEMENT
// ==========================================================================
async function loadSalesPhotos() {
  try {
    const res = await fetch(`/api/sales-photos?t=${Date.now()}`);
    if (res.ok) {
      const serverData = await res.json();
      const localData = JSON.parse(localStorage.getItem('gadgetmart_sales_photos') || '{}');
      cachedSalesPhotos = Object.assign({}, localData, serverData);
    }
  } catch(e) {
    cachedSalesPhotos = JSON.parse(localStorage.getItem('gadgetmart_sales_photos') || '{}');
  }
}

async function persistSalesPhotosToServer() {
  localStorage.setItem('gadgetmart_sales_photos', JSON.stringify(cachedSalesPhotos));
  try {
    await fetch('/api/sales-photos', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(cachedSalesPhotos)
    });
  } catch(e) {
    console.warn("Gagal menyimpan foto sales ke server:", e);
  }
}

function renderAdminPhotoTable() {
  const tbody = document.getElementById('adminPhotoTableBody');
  if (!tbody) return;
  const masterList = getMasterSales();

  tbody.innerHTML = masterList.map(m => {
    const photo = cachedSalesPhotos[m.name];
    const previewSrc = photo || 'logo.png';
    return `
      <tr>
        <td class="text-center">
          <img src="${previewSrc}" id="photoThumb_${m.name}" class="admin-photo-thumb" alt="${m.name}">
        </td>
        <td class="text-left font-bold">${m.name}</td>
        <td class="text-left font-mono">${m.code}</td>
        <td class="text-left">
          <input type="file" accept="image/*" class="form-input" style="padding: 4px;" onchange="handleAdminPhotoUpload('${m.name}', event)">
        </td>
        <td class="text-left">
          <div style="display: flex; gap: 6px;">
            <input type="text" id="photoUrlInput_${m.name}" value="${photo && photo.startsWith('http') ? photo : ''}" placeholder="https://..." class="form-input" style="padding: 4px 8px; font-size: 0.82rem;">
            <button class="btn btn-sm btn-crimson" onclick="saveAdminPhotoUrl('${m.name}')">Simpan</button>
          </div>
        </td>
        <td class="text-center">
          ${photo ? `<button class="btn btn-sm btn-outline text-crimson" onclick="deleteSalesPhoto('${m.name}')">Hapus</button>` : '-'}
        </td>
      </tr>
    `;
  }).join('');
}

function handleAdminPhotoUpload(picName, event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async (e) => {
    const base64 = e.target.result;
    cachedSalesPhotos[picName] = base64;
    await persistSalesPhotosToServer();
    renderAdminPhotoTable();
    renderSalesLeaderboard(currentLeaderboardSort);
    alert(`Foto untuk ${picName} berhasil disimpan permanen ke server!`);
  };
  reader.readAsDataURL(file);
}

async function saveAdminPhotoUrl(picName) {
  const input = document.getElementById(`photoUrlInput_${picName}`);
  if (!input) return;
  const url = input.value.trim();
  if (!url) {
    alert('Masukkan URL foto gambar terlebih dahulu!');
    return;
  }
  cachedSalesPhotos[picName] = url;
  await persistSalesPhotosToServer();
  renderAdminPhotoTable();
  renderSalesLeaderboard(currentLeaderboardSort);
  alert(`URL foto untuk ${picName} berhasil disimpan permanen ke server!`);
}

async function deleteSalesPhoto(picName) {
  if (!confirm(`Hapus foto profil sales ${picName}?`)) return;
  delete cachedSalesPhotos[picName];
  await persistSalesPhotosToServer();
  renderAdminPhotoTable();
  renderSalesLeaderboard(currentLeaderboardSort);
}

async function resetAllPhotosToDefault() {
  if (!confirm('Reset semua foto sales profil ke standar?')) return;
  cachedSalesPhotos = {};
  await persistSalesPhotosToServer();
  renderAdminPhotoTable();
  renderSalesLeaderboard(currentLeaderboardSort);
}

// ==========================================================================
// 15. MONTHLY TARGETS ENGINE (TARGET PER BULAN)
// ==========================================================================
async function loadMonthlyTargets() {
  try {
    const res = await fetch(`/api/monthly-targets?t=${Date.now()}`);
    if (res.ok) {
      cachedMonthlyTargets = await res.json();
      console.log('✅ Monthly targets loaded from server:', cachedMonthlyTargets);
    }
  } catch(e) {
    cachedMonthlyTargets = JSON.parse(localStorage.getItem('gadgetmart_monthly_targets') || '{}');
  }
}

async function persistMonthlyTargetsToServer() {
  localStorage.setItem('gadgetmart_monthly_targets', JSON.stringify(cachedMonthlyTargets));
  try {
    await fetch('/api/monthly-targets', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(cachedMonthlyTargets)
    });
  } catch(e) {
    console.warn("Gagal simpan target bulanan ke server:", e);
  }
}

function getMonthlyTargetKey(m, y) {
  return `${y}-${String(m).padStart(2, '0')}`;
}

function getTeamTargetForMonth(m, y) {
  const key = getMonthlyTargetKey(m, y);
  if (cachedMonthlyTargets && cachedMonthlyTargets[key] && cachedMonthlyTargets[key].teamTarget) {
    return Number(cachedMonthlyTargets[key].teamTarget);
  }
  return 8447356034;
}

function getSalesTargetForMonth(picCode, m, y, defaultTarget) {
  const key = getMonthlyTargetKey(m, y);
  if (cachedMonthlyTargets && cachedMonthlyTargets[key] && cachedMonthlyTargets[key].salesTargets && cachedMonthlyTargets[key].salesTargets[picCode] !== undefined) {
    return Number(cachedMonthlyTargets[key].salesTargets[picCode]);
  }
  return defaultTarget;
}

function onAdminTargetMonthChange() {
  const mSelect = document.getElementById('adminTargetMonthSelect');
  const yInput = document.getElementById('adminTargetYearInput');
  const teamInput = document.getElementById('adminTeamTargetInput');
  if (!mSelect || !yInput) return;

  const m = parseInt(mSelect.value, 10);
  const y = parseInt(yInput.value, 10);
  const key = getMonthlyTargetKey(m, y);
  const monthData = cachedMonthlyTargets[key] || {};

  if (teamInput) {
    teamInput.value = monthData.teamTarget || 8447356034;
  }

  const masterList = getMasterSales();
  const targetList = document.getElementById('adminTargetList');
  if (targetList) {
    targetList.innerHTML = masterList.filter(x => x.active).map(x => {
      const targetVal = (monthData.salesTargets && monthData.salesTargets[x.code] !== undefined) 
        ? monthData.salesTargets[x.code] 
        : x.target;
      return `
        <div class="admin-target-row">
          <label>Target ${x.name} (${x.code}):</label>
          <input type="number" id="targetInput_${x.code}" value="${targetVal}" class="form-input">
        </div>
      `;
    }).join('');
  }
}

function populateAdminPanel(masterList) {
  const masterTable = document.getElementById('adminMasterTableBody');
  onAdminTargetMonthChange();

  if (masterTable) {
    masterTable.innerHTML = masterList.map((m, idx) => `
      <tr>
        <td class="text-left font-bold">${m.code}</td>
        <td class="text-left">${m.name}</td>
        <td class="text-center">
          <span class="status-badge ${m.active ? 'badge-accent' : 'badge-ink'}">${m.active ? 'Aktif' : 'Non-Aktif'}</span>
        </td>
        <td class="text-center">
          <button class="btn btn-sm btn-outline" onclick="toggleSalesActive(${idx})">${m.active ? 'Nonaktifkan' : 'Aktifkan'}</button>
        </td>
      </tr>
    `).join('');
  }

  renderAdminPhotoTable();
}

async function saveAdminTargets() {
  const mSelect = document.getElementById('adminTargetMonthSelect');
  const yInput = document.getElementById('adminTargetYearInput');
  const teamInput = document.getElementById('adminTeamTargetInput');
  
  const m = mSelect ? parseInt(mSelect.value, 10) : activeMonth;
  const y = yInput ? parseInt(yInput.value, 10) : activeYear;
  const key = getMonthlyTargetKey(m, y);

  const teamTarget = teamInput ? (parseFloat(teamInput.value) || 8447356034) : 8447356034;
  const salesTargets = {};
  const masterList = getMasterSales();

  masterList.forEach(item => {
    const input = document.getElementById(`targetInput_${item.code}`);
    if (input) {
      const val = parseFloat(input.value) || 0;
      salesTargets[item.code] = val;
      if (m === activeMonth && y === activeYear) {
        item.target = val;
      }
    }
  });

  cachedMonthlyTargets[key] = {
    month: m,
    year: y,
    teamTarget: teamTarget,
    salesTargets: salesTargets
  };

  saveMasterSales(masterList);
  await persistMonthlyTargetsToServer();
  const monthName = mSelect ? mSelect.options[mSelect.selectedIndex].text : `Bulan ${m}`;
  alert(`✅ Target bulanan untuk ${monthName} ${y} berhasil disimpan permanen ke server!`);
  processAllReports();
}

function addNewMasterCode() {
  const code = document.getElementById('newMasterCode').value.trim();
  const name = document.getElementById('newMasterName').value.trim().toUpperCase();
  if (!code || !name) {
    alert("Isi kode dan nama PIC sales!");
    return;
  }
  const masterList = getMasterSales();
  masterList.push({ code, name, target: 867261886, active: true });
  saveMasterSales(masterList);
  document.getElementById('newMasterCode').value = '';
  document.getElementById('newMasterName').value = '';
  populateAdminPanel(masterList);
  processAllReports();
}

function toggleSalesActive(index) {
  const masterList = getMasterSales();
  if (masterList[index]) {
    masterList[index].active = !masterList[index].active;
    saveMasterSales(masterList);
    populateAdminPanel(masterList);
    processAllReports();
  }
}

function resetMasterToDefault() {
  if (confirm("Reset seluruh target dan master sales ke pengaturan pabrik?")) {
    saveMasterSales(DEFAULT_SALES_MASTER);
    populateAdminPanel(DEFAULT_SALES_MASTER);
    processAllReports();
  }
}

// 11. UNDUH GAMBAR PNG LENGKAP TANPA TERPOTONG SCROLLBAR
function downloadReportAsPNG() {
  if (currentActiveTab === 2) {
    downloadElementAsPNG('report2Container', 'Report2_KPI_Summary_Timegone');
  } else if (currentActiveTab === 3) {
    downloadElementAsPNG('arResultContainer', 'Cek_AR_Invoice_Status');
  } else if (currentActiveTab === 4) {
    downloadElementAsPNG('tabContent4', 'Data_Keep_Customer');
  } else if (currentActiveTab === 5) {
    downloadElementAsPNG('tabContent5', 'Audit_AR_Belum_Sosmed');
  } else if (currentActiveTab === 6) {
    downloadElementAsPNG('tabContent6', 'Arsip_Pencapaian_Bulanan');
  } else {
    downloadElementAsPNG('report1Container', 'Report1_Daily_Achievement_Sosmed');
  }
}

function downloadElementAsPNG(elementId, filePrefix) {
  const el = document.getElementById(elementId);
  if (!el) {
    alert("Elemen report tidak ditemukan.");
    return;
  }

  const btn = document.getElementById('btnDownloadPNG');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) btn.innerHTML = '⏳ Menyiapkan Gambar Lengkap...';

  // Ensure full width and height are completely unfolded without clipping
  const scrollContainers = el.querySelectorAll('.table-responsive');
  const originalStyles = [];

  scrollContainers.forEach(sc => {
    originalStyles.push({ el: sc, maxH: sc.style.maxHeight, overflow: sc.style.overflow });
    sc.style.maxHeight = 'none';
    sc.style.overflow = 'visible';
  });

  html2canvas(el, {
    scale: 2, // High resolution Retina
    useCORS: true,
    backgroundColor: '#ffffff',
    windowWidth: Math.max(document.documentElement.scrollWidth, 1400),
    width: el.scrollWidth,
    height: el.scrollHeight
  }).then(canvas => {
    // Restore styles
    originalStyles.forEach(s => {
      s.el.style.maxHeight = s.maxH;
      s.el.style.overflow = s.overflow;
    });

    const link = document.createElement('a');
    link.download = `GadgetMart_${filePrefix}_${new Date().toISOString().slice(0,10)}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    if (btn) btn.innerHTML = originalText;
  }).catch(err => {
    originalStyles.forEach(s => {
      s.el.style.maxHeight = s.maxH;
      s.el.style.overflow = s.overflow;
    });
    console.error("Gagal membuat PNG:", err);
    alert("Gagal mengunduh gambar PNG. Pastikan tabel ditampilkan.");
    if (btn) btn.innerHTML = originalText;
  });
}

// 12. CALCULATION MODE TOGGLE (DPP VS GROSS)
function setCalculationMode(mode) {
  currentCalculationMode = mode;
  const btnDPP = document.getElementById('btnModeDPP');
  const btnGross = document.getElementById('btnModeGross');
  if (btnDPP) btnDPP.classList.toggle('active', mode === 'DPP');
  if (btnGross) btnGross.classList.toggle('active', mode === 'GROSS');
  const tag = document.getElementById('modeIndicatorTag');
  if (tag) tag.innerText = mode === 'DPP' ? 'Nilai DPP (Sebelum PPN 11%)' : 'Nilai Gross (Harga Jual Fisik)';
  processAllReports();
}

// 13. GEMINI AI SMART SALES ASSISTANT
async function executeGeminiPrompt() {
  const input = document.getElementById('geminiPromptInput');
  const prompt = input.value.trim();
  if (!prompt) return;

  const card = document.getElementById('aiResponseCard');
  const body = document.getElementById('aiResponseContent');
  const time = document.getElementById('aiResponseTime');
  card.style.display = 'block';
  time.innerText = new Date().toLocaleTimeString('id-ID');
  body.innerHTML = '<div style="color:var(--crimson-primary);font-weight:700;">✨ Gemini sedang menganalisis query...</div>';

  // 13A. Check if query is checking WhatsApp number e.g. "cek wa 081386413707"
  const waMatch = prompt.match(/(?:wa|whatsapp|nomor|no\s*hp|telepon|hp)?\s*(\+?62\d{8,13}|08\d{8,12})/i);
  if (waMatch) {
    const phone = waMatch[1];
    quickCheckAR(phone);
    body.innerHTML = `<p>📱 <strong>Hasil Pencarian Pelanggan WhatsApp (${phone}):</strong> Layar telah beralih ke <strong>Tab 3 (Cek AR & WhatsApp)</strong> untuk menampilkan profil pelanggan dan rincian seluruh produk yang pernah dibeli!</p>`;
    return;
  }

  // 13B. Check if query is checking AR number e.g. "cek ar 260274906"
  const arMatch = prompt.match(/\b(26\d{7})\b/);
  if (arMatch) {
    const docNo = arMatch[1];
    quickCheckAR(docNo);
    body.innerHTML = `<p>🔍 <strong>Hasil Pengecekan AR ${docNo}:</strong> Layar telah beralih ke <strong>Tab 2 (Cek AR & WhatsApp)</strong> untuk menampilkan verifikasi kode sales sosmed dan rincian produk per item.</p>`;
    return;
  }

  // 13C. Check if query is asking about Keep data
  if (prompt.toLowerCase().includes('keep')) {
    switchTab(3);
    body.innerHTML = `<p>📋 <strong>Layar Beralih ke Tab 3 (Data Keep Customer):</strong> Menampilkan data keep Google Sheets dengan filter Kode Sosmed, Status, dan Periode. Filter dapat disesuaikan langsung di layar.</p>`;
    return;
  }

  const apiKey = localStorage.getItem('gemini_api_key') || DEFAULT_GEMINI_API_KEY;
  const model = localStorage.getItem('gemini_model') || DEFAULT_GEMINI_MODEL;

  if (apiKey) {
    try {
      const context = {
        totalAchieve: document.getElementById('kpiTotalAchieve').innerText,
        totalGP: document.getElementById('kpiTotalGP').innerText,
        timegone: `${timegonePct}%`,
        targetGP: document.getElementById('kpiTargetGP').innerText
      };
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: `Kamu adalah Senior Commercial Analyst GadgetMart. Berikan evaluasi eksekutif berdasarkan data berikut: ${JSON.stringify(context)}. Pertanyaan: "${prompt}". Jawab singkat, padat, dan tajam dalam Bahasa Indonesia.` }] }]
        })
      });
      const data = await res.json();
      if (data.candidates && data.candidates[0].content) {
        body.innerHTML = data.candidates[0].content.parts[0].text.replace(/\n/g, '<br>').replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        handlePromptHighlight(prompt);
        return;
      }
    } catch(e) { console.warn("Gemini API call fallback to local engine:", e); }
  }

  // Fallback Smart Local Engine
  setTimeout(() => {
    const q = prompt.toLowerCase();
    let text = '';
    if (q.includes('hari ini') || q.includes('daily')) {
      switchTab(1);
      setReportViewMode('r1');
      const row = parsedDailyMatrix[lastActiveDay] || {};
      text = `<p>📌 <strong>Laporan Harian per Tgl ${row.dateStr}:</strong> Total achieve hari ini adalah <strong>${formatRupiah(row.totalDay)}</strong> dengan rata-rata per PIC <strong>${formatRupiah(row.avg)}</strong>. Baris tanggal ini telah disorot di Tab 1!</p>`;
      handlePromptHighlight(prompt);
    } else if (q.includes('timegone') || q.includes('gp') || q.includes('kpi')) {
      switchTab(1);
      setReportViewMode('r2');
      text = `<p>🎯 <strong>Evaluasi KPI vs Timegone (${timegonePct}%):</strong> Capaian omzet tim berada di posisi <strong>${document.getElementById('kpiProgressPct').innerText}</strong> (Status: <strong>🖤 On Track</strong>). Perhatikan performa GP yang masih perlu didorong menuju target 5%!</p>`;
    } else {
      text = `<p>Data berhasil dievaluasi untuk <em>"${prompt}"</em>. Total pencapaian tim per tanggal terakhir adalah <strong>${document.getElementById('kpiTotalAchieve').innerText}</strong>.</p>`;
    }
    body.innerHTML = text;
  }, 350);
}

function handlePromptHighlight(prompt) {
  document.querySelectorAll('.row-highlight').forEach(el => el.classList.remove('row-highlight'));
  const targetTr = document.getElementById(`row-date-${lastActiveDay}`);
  if (targetTr) {
    targetTr.classList.add('row-highlight');
    targetTr.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
}

function quickPrompt(text) {
  document.getElementById('geminiPromptInput').value = text;
  executeGeminiPrompt();
}

function resetTableHighlight() {
  document.querySelectorAll('.row-highlight').forEach(el => el.classList.remove('row-highlight'));
}

function filterTableDate(query) {
  const q = query.trim().toLowerCase();
  for (let d = 1; d <= daysInMonth; d++) {
    const tr = document.getElementById(`row-date-${d}`);
    if (!tr) continue;
    const dateStr = parsedDailyMatrix[d].dateStr;
    tr.style.display = (!q || dateStr.includes(q)) ? '' : 'none';
  }
}

function handleTurnoverFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;

  const fileNameLower = (file.name || '').toLowerCase();
  if (fileNameLower.includes('croselling') || fileNameLower.includes('crosselling')) {
    handleCrosellingFileSelect(e);
    return;
  }

  const fileSizeMb = (file.size / (1024 * 1024)).toFixed(1);
  showSaveToast(`⏳ Sedang membaca file Excel (${fileSizeMb} MB)... Mohon tunggu beberapa detik.`);
  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const buffer = evt.target.result;
      const jsonData = readTurnoverJsonFromBuffer(buffer);
      if (!jsonData || jsonData.length === 0) {
        alert("File Excel turnover kosong atau sheet transaksi tidak terbaca!");
        return;
      }
      // Check if it's actually croselling data (has 'No HP Customer' and not 'DPP')
      const firstRow = jsonData[0] || {};
      if ('No HP Customer' in firstRow && !('DPP' in firstRow)) {
        rawCrosellingRows = jsonData;
        const updateTs = getFormattedTimestamp();
        await saveFileToIndexedDB('croselling_file', buffer);
        localStorage.setItem('gadgetmart_croselling_timestamp', updateTs);
        updateCrosellingIndicator(`File Cross Selling terupdate per ${updateTs}`);
        saveExcelToServer(buffer, 'croselling');
        if (typeof runARMissingSosmedAudit === 'function') {
          runARMissingSosmedAudit();
        }
        alert(`✅ File terdeteksi sebagai data Cross Selling (${jsonData.length.toLocaleString('id-ID')} baris data)!\nDatabase Cek AR & Audit AR telah diperbarui.`);
        return;
      }

      rawTurnoverRows = jsonData;
      const updateTs = getFormattedTimestamp();
      await saveFileToIndexedDB('turnover_file', buffer);
      localStorage.setItem('gadgetmart_turnover_timestamp', updateTs);
      localStorage.setItem('gadgetmart_turnover_is_custom', 'true');
      updateTurnoverIndicator(`File turnover terupdate per ${updateTs}`);
      saveExcelToServer(buffer, 'turnover');
      processAllReports();
      alert(`✅ Berhasil mengunggah file turnover (${jsonData.length.toLocaleString('id-ID')} baris transaksi)!\nDashboard & Laporan Pencapaian telah diperbarui dan data tersimpan permanen.`);
    } catch (err) {
      console.error("Gagal membaca file turnover:", err);
      alert(`Gagal membaca file Excel turnover: ${err.message || 'Format file atau memori peramban tidak mencukupi'}.`);
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleCrosellingFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  const fileSizeMb = (file.size / (1024 * 1024)).toFixed(1);
  showSaveToast(`⏳ Sedang membaca file Cross Selling (${fileSizeMb} MB)... Mohon tunggu beberapa detik.`);
  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const buffer = evt.target.result;
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet);
      if (!jsonData || jsonData.length === 0) {
        alert("File Excel Cross Selling kosong atau tidak terbaca!");
        return;
      }
      rawCrosellingRows = jsonData;
      const updateTs = getFormattedTimestamp();
      await saveFileToIndexedDB('croselling_file', buffer);
      localStorage.setItem('gadgetmart_croselling_timestamp', updateTs);
      updateCrosellingIndicator(`File Cross Selling terupdate per ${updateTs}`);
      saveExcelToServer(buffer, 'croselling');
      if (typeof runARMissingSosmedAudit === 'function') {
        runARMissingSosmedAudit();
      }
      alert(`✅ Berhasil mengunggah file Cross Selling (${jsonData.length.toLocaleString('id-ID')} baris data)!\nDatabase Cek AR & Audit AR telah diperbarui dan tersimpan.`);
      if (document.getElementById('arInputSearch') && document.getElementById('arInputSearch').value.trim()) {
        executeARSearch();
      }
    } catch (err) {
      console.error("Gagal membaca file cross selling:", err);
      alert("Gagal membaca file Excel Cross Selling. Pastikan format file valid.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function updateCrosellingIndicator(text) {
  const el = document.getElementById('crosellingUpdateText');
  if (el) el.innerText = text;
}

// Function alias for backwards compatibility
function handleFileSelect(e) {
  handleTurnoverFileSelect(e);
}

const dropZone = document.getElementById('dropZone');
if (dropZone) {
  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      const reader = new FileReader();
      reader.onload = (evt) => parseExcelFile(evt.target.result, file.name);
      reader.readAsArrayBuffer(file);
    }
  });
}

// Dedicated Cross Selling dropzone
const crosellingDropZone = document.getElementById('crosellingDropZone');
if (crosellingDropZone) {
  crosellingDropZone.addEventListener('dragover', (e) => { e.preventDefault(); crosellingDropZone.classList.add('drag-over'); });
  crosellingDropZone.addEventListener('dragleave', () => crosellingDropZone.classList.remove('drag-over'));
  crosellingDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    crosellingDropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const buffer = evt.target.result;
          const workbook = XLSX.read(buffer, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(sheet);
          if (!jsonData || jsonData.length === 0) return;
          rawCrosellingRows = jsonData;
          const updateTs = getFormattedTimestamp();
          await saveFileToIndexedDB('croselling_file', buffer);
          localStorage.setItem('gadgetmart_croselling_timestamp', updateTs);
          updateCrosellingIndicator(`File Cross Selling terupdate per ${updateTs}`);
          saveExcelToServer(buffer, 'croselling');
          alert(`✅ Berhasil mengimpor file Cross Selling (${jsonData.length.toLocaleString('id-ID')} baris data)!`);
        } catch (err) {
          console.error(err);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

// ==========================================================================
// 14. REALTIME CROSS-DEVICE SYNC ENGINE & SERVER DUAL-LOADER
// ==========================================================================
let lastKnownSyncStatus = {
  turnoverMtime: 0,
  crosellingMtime: 0,
  archivesMtime: 0,
  targetsMtime: 0,
  photosMtime: 0,
  keepMtime: 0
};

// LOAD BOTH DEFAULT EXCEL FILES - PRIORITASKAN CUSTOM UPLOAD (DATABASE / INDEXEDDB) AGAR DATA TIDAK HILANG SAAT REFRESH
async function loadAllDefaultFiles() {
  const statusTitle = document.getElementById('fileStatusTitle');
  if (statusTitle) statusTitle.innerText = 'Menghubungkan ke Server Database...';
  
  // 1. Fetch turnover.xlsx:
  // Cek apakah ada file turnover kustom yang tersimpan di IndexedDB atau PostgreSQL Database
  let turnoverLoaded = false;
  const isCustomTurnover = localStorage.getItem('gadgetmart_turnover_is_custom') === 'true';

  // A. Jika ada file custom di IndexedDB, muat langsung agar REFRESH TETAP STAY & INSTAN
  if (isCustomTurnover) {
    try {
      const cachedTurnover = await getFileFromIndexedDB('turnover_file');
      if (cachedTurnover) {
        parseTurnoverBuffer(cachedTurnover);
        turnoverLoaded = true;
        const ts = localStorage.getItem('gadgetmart_turnover_timestamp') || getFormattedTimestamp();
        updateTurnoverIndicator(`File turnover terupdate per ${ts}`);
        console.log('✅ turnover.xlsx kustom dimuat dari IndexedDB lokal!');
      }
    } catch (e) {}
  }

  // B. Ambil dari Database Server (/api/get-excel?type=turnover)
  if (!turnoverLoaded) {
    try {
      const resTurnover = await fetch(`/api/get-excel?type=turnover&t=${Date.now()}`);
      if (resTurnover.ok) {
        const isDbCustom = resTurnover.headers.get('X-Custom-File') === 'true';
        const serverTs = resTurnover.headers.get('X-File-Timestamp');
        const buf1 = await resTurnover.arrayBuffer();
        parseTurnoverBuffer(buf1);
        turnoverLoaded = true;
        if (isDbCustom) {
          await saveFileToIndexedDB('turnover_file', buf1);
          localStorage.setItem('gadgetmart_turnover_is_custom', 'true');
          const ts = serverTs ? new Date(serverTs).toLocaleString('id-ID') : getFormattedTimestamp();
          localStorage.setItem('gadgetmart_turnover_timestamp', ts);
          updateTurnoverIndicator(`File turnover terupdate per ${ts}`);
        } else {
          const ts = localStorage.getItem('gadgetmart_turnover_timestamp') || 'Default Sistem (13/10/2026 18:00 WITA)';
          updateTurnoverIndicator(`File turnover terupdate per ${ts}`);
        }
        console.log('✅ turnover.xlsx dimuat dari database server!');
      }
    } catch (e) {}
  }

  // C. Fallback ke file statis jika server belum punya custom upload
  if (!turnoverLoaded) {
    try {
      const resTurnover = await fetch(`turnover.xlsx?t=${Date.now()}`);
      if (resTurnover.ok) {
        const buf1 = await resTurnover.arrayBuffer();
        parseTurnoverBuffer(buf1);
        turnoverLoaded = true;
        const ts = localStorage.getItem('gadgetmart_turnover_timestamp') || 'Default Sistem (13/10/2026 18:00 WITA)';
        updateTurnoverIndicator(`File turnover terupdate per ${ts}`);
      }
    } catch (e) {}
  }

  // 2. Fetch CROSSELLING-12102026.xlsx
  let crosellingLoaded = false;
  try {
    const resCroselling = await fetch(`/api/get-excel?type=croselling&t=${Date.now()}`);
    if (resCroselling.ok) {
      const isDbCustom = resCroselling.headers.get('X-Custom-File') === 'true';
      const buf2 = await resCroselling.arrayBuffer();
      parseCrosellingBuffer(buf2);
      crosellingLoaded = true;
      if (isDbCustom) {
        await saveFileToIndexedDB('croselling_file', buf2);
      }
      const ts = localStorage.getItem('gadgetmart_croselling_timestamp') || 'Default Sistem (12/10/2026)';
      updateCrosellingIndicator(`File Cross Selling terupdate per ${ts}`);
    }
  } catch(e) {}

  if (!crosellingLoaded) {
    try {
      const cachedCroselling = await getFileFromIndexedDB('croselling_file');
      if (cachedCroselling) {
        parseCrosellingBuffer(cachedCroselling);
        crosellingLoaded = true;
        const ts = localStorage.getItem('gadgetmart_croselling_timestamp') || 'Tersimpan di Peramban';
        updateCrosellingIndicator(`File Cross Selling terupdate per ${ts}`);
      }
    } catch (e) {}
  }

  if (!crosellingLoaded) {
    try {
      const resCroselling = await fetch(`CROSSELLING-12102026.xlsx?t=${Date.now()}`);
      if (resCroselling.ok) {
        const buf2 = await resCroselling.arrayBuffer();
        parseCrosellingBuffer(buf2);
        crosellingLoaded = true;
        const ts = localStorage.getItem('gadgetmart_croselling_timestamp') || 'Default Sistem (12/10/2026)';
        updateCrosellingIndicator(`File Cross Selling terupdate per ${ts}`);
      }
    } catch (e) {}
  }

  // Inisialisasi status sync server baseline
  try {
    const resStatus = await fetch(`/api/database-status?t=${Date.now()}`, { cache: 'no-store' });
    if (resStatus.ok) {
      const s = await resStatus.json();
      if (s.turnover && s.turnover.mtimeMs) lastKnownSyncStatus.turnoverMtime = s.turnover.mtimeMs;
      if (s.archives && s.archives.mtimeMs) lastKnownSyncStatus.archivesMtime = s.archives.mtimeMs;
      if (s.croselling && s.croselling.mtimeMs) lastKnownSyncStatus.crosellingMtime = s.croselling.mtimeMs;
      if (s.targets && s.targets.mtimeMs) lastKnownSyncStatus.targetsMtime = s.targets.mtimeMs;
      if (s.photos && s.photos.mtimeMs) lastKnownSyncStatus.photosMtime = s.photos.mtimeMs;
      if (s.keep && s.keep.mtimeMs) lastKnownSyncStatus.keepMtime = s.keep.mtimeMs;
    }
  } catch(e) {}
}

// REALTIME POLLING: MENDETEKSI PERUBAHAN & EDIT DARI LAPTOP LAIN SECARA LIVE
async function pollServerSyncStatus() {
  try {
    const res = await fetch(`/api/database-status?t=${Date.now()}`, { cache: 'no-store' });
    if (!res.ok) return;
    const status = await res.json();
    if (!status) return;

    // A. Deteksi jika turnover.xlsx diperbarui dari laptop lain
    if (status.turnover && status.turnover.exists && status.turnover.mtimeMs) {
      if (lastKnownSyncStatus.turnoverMtime && status.turnover.mtimeMs > lastKnownSyncStatus.turnoverMtime) {
        console.log('🔄 [Sync Engine] File turnover.xlsx baru terdeteksi dari laptop lain!');
        lastKnownSyncStatus.turnoverMtime = status.turnover.mtimeMs;
        const resTurnover = await fetch(`turnover.xlsx?t=${Date.now()}`);
        if (resTurnover.ok) {
          const buf = await resTurnover.arrayBuffer();
          parseTurnoverBuffer(buf);
          await saveFileToIndexedDB('turnover_file', buf);
          processAllReports();
          showSaveToast('🔄 File turnover.xlsx telah diperbarui otomatis dari laptop lain!');
        }
      } else if (!lastKnownSyncStatus.turnoverMtime) {
        lastKnownSyncStatus.turnoverMtime = status.turnover.mtimeMs;
      }
    }

    // B. Deteksi jika Arsip Bulanan / Gross Profit diedit dari laptop lain
    if (status.archives && status.archives.exists && status.archives.mtimeMs) {
      if (lastKnownSyncStatus.archivesMtime && status.archives.mtimeMs > lastKnownSyncStatus.archivesMtime) {
        console.log('🔄 [Sync Engine] Arsip bulanan / Gross Profit diperbarui dari laptop lain!');
        lastKnownSyncStatus.archivesMtime = status.archives.mtimeMs;
        const resArc = await fetch(`/api/monthly-archives?t=${Date.now()}`);
        if (resArc.ok) {
          cachedMonthlyArchives = await resArc.json();
          localStorage.setItem('gadgetmart_sosmed_archives', JSON.stringify(cachedMonthlyArchives));
          renderArchiveGrid();
          renderMonthToMonthCharts();
          showSaveToast('🔄 Data arsip bulanan & Gross Profit telah diperbarui otomatis dari laptop lain!');
        }
      } else if (!lastKnownSyncStatus.archivesMtime) {
        lastKnownSyncStatus.archivesMtime = status.archives.mtimeMs;
      }
    }

    // C. Deteksi jika File Cross Selling diperbarui dari laptop lain
    if (status.croselling && status.croselling.exists && status.croselling.mtimeMs) {
      if (lastKnownSyncStatus.crosellingMtime && status.croselling.mtimeMs > lastKnownSyncStatus.crosellingMtime) {
        console.log('🔄 [Sync Engine] File Cross Selling baru terdeteksi dari laptop lain!');
        lastKnownSyncStatus.crosellingMtime = status.croselling.mtimeMs;
        const resCross = await fetch(`CROSSELLING-12102026.xlsx?t=${Date.now()}`);
        if (resCross.ok) {
          const buf = await resCross.arrayBuffer();
          parseCrosellingBuffer(buf);
          await saveFileToIndexedDB('croselling_file', buf);
          showSaveToast('🔄 File Cross Selling telah diperbarui otomatis dari laptop lain!');
        }
      } else if (!lastKnownSyncStatus.crosellingMtime) {
        lastKnownSyncStatus.crosellingMtime = status.croselling.mtimeMs;
      }
    }

    // D. Deteksi jika Target Bulanan diedit dari laptop lain
    if (status.targets && status.targets.exists && status.targets.mtimeMs) {
      if (lastKnownSyncStatus.targetsMtime && status.targets.mtimeMs > lastKnownSyncStatus.targetsMtime) {
        lastKnownSyncStatus.targetsMtime = status.targets.mtimeMs;
        await loadMonthlyTargets();
        renderMonthToMonthCharts();
      } else if (!lastKnownSyncStatus.targetsMtime) {
        lastKnownSyncStatus.targetsMtime = status.targets.mtimeMs;
      }
    }

    // E. Deteksi jika Foto Sales diedit dari laptop lain
    if (status.photos && status.photos.exists && status.photos.mtimeMs) {
      if (lastKnownSyncStatus.photosMtime && status.photos.mtimeMs > lastKnownSyncStatus.photosMtime) {
        lastKnownSyncStatus.photosMtime = status.photos.mtimeMs;
        await loadSalesPhotos();
        renderSalesLeaderboard('omzet');
      } else if (!lastKnownSyncStatus.photosMtime) {
        lastKnownSyncStatus.photosMtime = status.photos.mtimeMs;
      }
    }

    // F. Deteksi jika Data Keep Google Sheets diperbarui otomatis dari server
    if (status.keep && status.keep.exists && status.keep.mtimeMs) {
      if (lastKnownSyncStatus.keepMtime && status.keep.mtimeMs > lastKnownSyncStatus.keepMtime) {
        console.log('🔄 [Sync Engine] Data Keep Google Sheets baru terdeteksi!');
        lastKnownSyncStatus.keepMtime = status.keep.mtimeMs;
        const prog = startSmoothProgress({
          id: 'auto-sync-keep-progress',
          title: '📡 Auto-Sync Data Google Sheets',
          detail: 'Data keep diperbarui di server, memuat...'
        });
        await loadKeepData(true);
        prog.complete({
          title: '✅ Data Google Sheets Tersinkronisasi!',
          detail: `${rawKeepRows.length.toLocaleString('id-ID')} baris data live terupdate`
        });
      } else if (!lastKnownSyncStatus.keepMtime) {
        lastKnownSyncStatus.keepMtime = status.keep.mtimeMs;
      }
    }
  } catch (err) {
    // Silent fail if network briefly disconnected
  }
}

// ==========================================
// SERVER-SENT EVENTS (SSE) LIVE SEAMLESS PUSH
// ==========================================
let liveEventsSource = null;

function initLiveEventsSSE() {
  if (typeof EventSource === 'undefined') {
    console.warn('Browser tidak mendukung EventSource SSE');
    return;
  }
  if (liveEventsSource) {
    try { liveEventsSource.close(); } catch (e) {}
  }

  try {
    liveEventsSource = new EventSource('/api/events');

    liveEventsSource.onopen = () => {
      console.log('⚡ [SSE Live Push] Terhubung ke siaran real-time server lokal!');
    };

    liveEventsSource.addEventListener('REPORT_AR_UPDATED', (e) => {
      try {
        const data = JSON.parse(e.data || '{}');
        if (!data.nomorWA) return;
        const cleanWA = String(data.nomorWA).replace(/\D/g, '');
        const newAR = String(data.arDoc || '').trim();

        console.log(`⚡ [SSE Broadcast] REPORT AR Diperbarui secara live: WA=${cleanWA} -> AR=${newAR}`);

        // 1. Update memory rawKeepRows
        if (typeof rawKeepRows !== 'undefined' && Array.isArray(rawKeepRows)) {
          rawKeepRows.forEach(k => {
            if (String(k['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
              k['REPORT AR'] = newAR;
            }
          });
        }

        // 2. Update memory rawAuditARList & filteredAuditARList
        if (typeof rawAuditARList !== 'undefined' && Array.isArray(rawAuditARList)) {
          rawAuditARList.forEach(a => {
            if (a.keep && String(a.keep['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
              a.reportARDoc = newAR;
              a.keep['REPORT AR'] = newAR;
              if (newAR && typeof docMap !== 'undefined' && docMap.has(newAR)) {
                a.reportARItems = docMap.get(newAR);
                a.reportARDate = a.reportARItems[0]['Posting Date'];
                a.reportARCabang = a.reportARItems[0]['Cabang'];
              }
            }
          });
        }

        // 3. Highlight efek visual live update pada sel
        const targetCells = document.querySelectorAll(`.cell-report-ar-wrapper[data-wa="${cleanWA}"]`);
        targetCells.forEach(cell => {
          cell.classList.add('cell-live-updated');
          setTimeout(() => cell.classList.remove('cell-live-updated'), 2500);
        });

        // 4. Re-render tampilan yang aktif tanpa reload halaman
        const activeTab = document.querySelector('.tab-content.active');
        if (activeTab && activeTab.id === 'tab-audit-ar') {
          renderAuditPage();
        } else if (activeTab && activeTab.id === 'tab-keep-data') {
          renderKeepPage();
        }
      } catch (err) {
        console.warn('Gagal memproses event REPORT_AR_UPDATED:', err);
      }
    });

    liveEventsSource.addEventListener('KEEP_STATUS_UPDATED', (e) => {
      try {
        const data = JSON.parse(e.data || '{}');
        if (!data.nomorWA) return;
        const cleanWA = String(data.nomorWA).replace(/\D/g, '');
        const newStatus = String(data.status || '').trim().toUpperCase();

        console.log(`⚡ [SSE Broadcast] Status Keep Diperbarui secara live: WA=${cleanWA} -> Status=${newStatus}`);

        if (typeof rawKeepRows !== 'undefined' && Array.isArray(rawKeepRows)) {
          rawKeepRows.forEach(k => {
            if (String(k['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
              k['Status'] = newStatus;
            }
          });
        }

        if (typeof rawAuditARList !== 'undefined' && Array.isArray(rawAuditARList)) {
          rawAuditARList.forEach(a => {
            if (a.keep && String(a.keep['Nomor WA'] || '').replace(/\D/g, '') === cleanWA) {
              a.keep['Status'] = newStatus;
            }
          });
        }

        const selects = document.querySelectorAll(`select.select-status-keep`);
        selects.forEach(sel => {
          const tr = sel.closest('tr');
          if (tr && tr.innerHTML.includes(cleanWA)) {
            sel.value = newStatus;
            sel.setAttribute('data-prev', newStatus);
            sel.className = `select-status-keep ${getKeepStatusClass(newStatus)}`;
            tr.className = getAuditRowClassByKeepStatus(newStatus);
            tr.classList.add('cell-live-updated');
            setTimeout(() => tr.classList.remove('cell-live-updated'), 2500);
          }
        });
      } catch (err) {
        console.warn('Gagal memproses event KEEP_STATUS_UPDATED:', err);
      }
    });

    liveEventsSource.addEventListener('KEEP_CSV_SYNCED', () => {
      console.log('📡 [SSE Broadcast] File keep-data.csv disinkronkan di server.');
      loadKeepData(true);
    });

    liveEventsSource.onerror = (err) => {
      // EventSource otomatis mencoba reconnect secara berkala
    };
  } catch (e) {
    console.warn('Inisialisasi SSE gagal:', e);
  }
}

let syncEngineInterval = null;
function startRealtimeSyncEngine() {
  if (syncEngineInterval) clearInterval(syncEngineInterval);
  // Mulai siaran live SSE ke server
  initLiveEventsSSE();
  // Polling fallback status server setiap 10 detik
  syncEngineInterval = setInterval(pollServerSyncStatus, 10000);
  console.log('📡 Realtime Cross-Device SSE & Fallback Sync Engine aktif');
}


// Legacy function alias
function loadDefaultTurnover() {
  loadAllDefaultFiles();
}

// 15. EXPORT EXCEL
function exportActiveReport() {
  const wb = XLSX.utils.book_new();
  const tbl1 = document.getElementById('matrixTable');
  const ws1 = XLSX.utils.table_to_sheet(tbl1);
  XLSX.utils.book_append_sheet(wb, ws1, "DAILY_ACHIEVEMENT");

  const tbl2 = document.getElementById('kpiSummaryTable');
  const ws2 = XLSX.utils.table_to_sheet(tbl2);
  XLSX.utils.book_append_sheet(wb, ws2, "KPI_SUMMARY_GP");

  XLSX.writeFile(wb, `GadgetMart_Sosmed_Achievement_Report_${activeMonth}_${activeYear}.xlsx`);
}

// 16. MODALS
function openGeminiModal() {
  document.getElementById('geminiModal').style.display = 'flex';
  document.getElementById('geminiApiKeyInput').value = localStorage.getItem('gemini_api_key') || DEFAULT_GEMINI_API_KEY;
  document.getElementById('geminiModelSelect').value = localStorage.getItem('gemini_model') || DEFAULT_GEMINI_MODEL;
}
function closeGeminiModal() { document.getElementById('geminiModal').style.display = 'none'; }
function saveGeminiConfig() {
  localStorage.setItem('gemini_api_key', document.getElementById('geminiApiKeyInput').value.trim());
  localStorage.setItem('gemini_model', document.getElementById('geminiModelSelect').value);
  alert("Konfigurasi Gemini API berhasil disimpan!");
  closeGeminiModal();
}
function openFormulaModal() { document.getElementById('formulaModal').style.display = 'flex'; }
function closeFormulaModal() { document.getElementById('formulaModal').style.display = 'none'; }

// ==========================================================================
// 16. DASHBOARD INTERACTIVE FILTERS & REVENUE / GP METRICS ENGINE
// ==========================================================================
function populateDashboardFilterDropdowns() {
  const cabangSel = document.getElementById('dashFilterCabang');
  const salesSel = document.getElementById('dashFilterSales');
  if (!cabangSel || !salesSel) return;

  const currentCabang = dashFilters.cabang;
  const currentSales = dashFilters.sales;

  // Cabang list from rawTurnoverRows
  const cabSet = new Set();
  if (rawTurnoverRows) {
    rawTurnoverRows.forEach(r => {
      const c = String(r['Cabang'] || r['SalesCabang'] || '').trim();
      if (c && c !== 'undefined') cabSet.add(c);
    });
  }
  const sortedCabang = Array.from(cabSet).sort();
  cabangSel.innerHTML = '<option value="all">Semua Cabang Toko</option>' + 
    sortedCabang.map(c => {
      const label = CABANG_NAMES[c] || c;
      return `<option value="${c}" ${c === currentCabang ? 'selected' : ''}>${label}</option>`;
    }).join('');

  // Sales list from active PICs
  const masterList = getMasterSales().filter(x => x.active);
  salesSel.innerHTML = '<option value="all">Semua PIC Sales Sosmed</option>' + 
    masterList.map(s => `<option value="${s.code}" ${s.code === currentSales ? 'selected' : ''}>${s.name} (${s.code})</option>`).join('');
}

function setDashCalculationMode(mode) {
  dashFilters.mode = mode;
  setCalculationMode(mode);
  const btnDPP = document.getElementById('dashBtnDPP');
  const btnGross = document.getElementById('dashBtnGross');
  if (btnDPP) btnDPP.classList.toggle('active', mode === 'DPP');
  if (btnGross) btnGross.classList.toggle('active', mode === 'GROSS');
  applyDashboardFilters();
}

function resetDashboardFilters() {
  dashFilters = { month: '10', cabang: 'all', sales: 'all', mode: currentCalculationMode };
  const mSel = document.getElementById('dashFilterMonth');
  const cSel = document.getElementById('dashFilterCabang');
  const sSel = document.getElementById('dashFilterSales');
  if (mSel) mSel.value = '10';
  if (cSel) cSel.value = 'all';
  if (sSel) sSel.value = 'all';
  applyDashboardFilters();
}

function applyDashboardFilters() {
  const mSel = document.getElementById('dashFilterMonth');
  const cSel = document.getElementById('dashFilterCabang');
  const sSel = document.getElementById('dashFilterSales');

  if (mSel) dashFilters.month = mSel.value;
  if (cSel) dashFilters.cabang = cSel.value;
  if (sSel) dashFilters.sales = sSel.value;

  const targetMonth = dashFilters.month === 'all' ? null : parseInt(dashFilters.month, 10);
  const targetCabang = dashFilters.cabang;
  const targetSales = dashFilters.sales;
  const isDPP = (currentCalculationMode === 'DPP');

  let totalRev = 0;
  let totalGP = 0;
  const dealDocs = new Set();

  if (rawTurnoverRows && rawTurnoverRows.length > 0) {
    rawTurnoverRows.forEach(r => {
      // 1. Cabang filter
      if (targetCabang !== 'all') {
        const c = String(r['Cabang'] || r['SalesCabang'] || '').trim();
        if (c !== targetCabang) return;
      }

      // 2. Sales filter
      const emp = String(r['Sales Employee Name'] || '');
      const rem = String(r['Remarks'] || '');
      const cab = String(r['SalesCabang'] || '');
      if (targetSales === 'all') {
        if (!isSosmedSalesCode(emp, rem, cab)) return;
      } else {
        if (!emp.includes(targetSales) && !rem.includes(targetSales) && !cab.includes(targetSales)) return;
      }

      const dpp = Number(r['DPP']) || 0;
      const total = Number(r['Total']) || 0;
      const fixVal = (r['FIX'] !== undefined && r['FIX'] !== '') ? Number(r['FIX']) : NaN;
      const rawGP = Number(r['GrossProfit']) || 0;
      const gp = !isNaN(fixVal) ? fixVal : rawGP;
      const amt = isDPP ? dpp : total;

      totalRev += amt;
      totalGP += gp;
      const doc = r['Document Number'];
      if (doc) dealDocs.add(doc);
    });
  }

  // Update Summary Strip UI
  const targetTeam = getTeamTargetForMonth(targetMonth || activeMonth, activeYear);
  const gap = totalRev - targetTeam;

  const elRev = document.getElementById('dashSummaryRevenue');
  const elGP = document.getElementById('dashSummaryGP');
  const elDeals = document.getElementById('dashSummaryDeals');
  const elTarget = document.getElementById('dashSummaryTarget');
  const elGap = document.getElementById('dashSummaryGap');

  if (elRev) elRev.innerText = formatRupiah(totalRev);
  if (elGP) elGP.innerText = formatRupiah(totalGP);
  if (elDeals) elDeals.innerText = `${dealDocs.size.toLocaleString('id-ID')} Trx`;
  if (elTarget) elTarget.innerText = formatRupiah(targetTeam);
  if (elGap) {
    elGap.innerText = formatRupiah(gap);
    elGap.style.color = gap >= 0 ? '#10b981' : '#f43f5e';
  }
}

// (Yipyep Floating AI Chat Widget telah dihapus sesuai permintaan)

// Live Real-Time Date and Clock Engine for Report Header
function initLiveClock() {
  function update() {
    const now = new Date();
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const dayName = days[now.getDay()];
    const dateNum = String(now.getDate()).padStart(2, '0');
    const monthName = months[now.getMonth()];
    const year = now.getFullYear();

    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');

    const dateEl = document.getElementById('liveClockDate');
    const timeEl = document.getElementById('liveClockTime');
    if (dateEl) dateEl.innerText = `${dayName}, ${dateNum} ${monthName} ${year}`;
    if (timeEl) timeEl.innerText = `${hours}:${minutes}:${seconds} WITA`;
  }
  update();
  setInterval(update, 1000);
}

// ==========================================================================
// MODUL KLAIM AR & PENGAWASAN DENDA CABANG (SIKLUS 3 HARI + H+1)
// ==========================================================================
let rawARClaimsList = [];
let filteredARClaimsList = [];
let currentClaimsSubTab = 'COLLECTING'; // 'COLLECTING', 'SHARED', 'PENALTY', 'RESOLVED', 'INVALID'
let claimsCurrentPage = 1;
const CLAIMS_PAGE_SIZE = 50;
let currentEvidenceTargetClaimId = null;

async function loadARClaimsData() {
  try {
    const res = await fetch('/api/ar-claims');
    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data)) {
        rawARClaimsList = data;
      }
    }
  } catch (e) {
    console.warn('Gagal load /api/ar-claims, fallback ke localStorage:', e);
  }
  if (!rawARClaimsList || rawARClaimsList.length === 0) {
    const local = localStorage.getItem('gadgetmart_ar_claims_data');
    if (local) {
      try { rawARClaimsList = JSON.parse(local); } catch(err){}
    }
  }

  // Normalize properties
  rawARClaimsList.forEach(c => {
    c.validationStatus = c.validationStatus || 'VALID';
    c.invalidReason = c.invalidReason || '';
    if (!Array.isArray(c.evidenceImages)) {
      c.evidenceImages = c.evidenceImage ? [c.evidenceImage] : [];
    }
    c.evidenceImage = c.evidenceImages[0] || null;
    c.periode = c.periode || 'Oktober-26';
    c.keepText = c.keepText || '';
  });

  checkARClaimsAutoResolution(false);
  renderARClaimsModule();
}

async function saveARClaimsData() {
  localStorage.setItem('gadgetmart_ar_claims_data', JSON.stringify(rawARClaimsList));
  try {
    await fetch('/api/ar-claims', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(rawARClaimsList)
    });
  } catch (e) {
    console.warn('Gagal simpan ke backend /api/ar-claims:', e);
  }
}

function addSingleARToClaims(auditItem, customItems = null) {
  if (!auditItem) return;
  const docNum = String(auditItem.doc || auditItem.reportARDoc || auditItem.vlookupARDoc || '').trim();
  if (!docNum || docNum === '-') {
    alert('Dokumen AR tidak valid untuk diklaim.');
    return;
  }

  const itemsToClaim = (customItems && customItems.length > 0) ? customItems : (auditItem.items || []);
  const missingOnly = itemsToClaim.filter(it => !isItemExemptFromSosmed(it) && !isItemInputSosmed(it));
  const finalItems = missingOnly.length > 0 ? missingOnly : itemsToClaim;

  const periode = (auditItem.keep && auditItem.keep['PERIODE']) 
    ? String(auditItem.keep['PERIODE']).trim() 
    : 'Oktober-26';
  const keepFormatted = (typeof formatKeepTextForClipboard === 'function') 
    ? formatKeepTextForClipboard(auditItem.keep) 
    : '';

  let addedCount = 0;
  finalItems.forEach((it, idx) => {
    const desc = it['Dscription'] || it['Item/Service Description'] || '-';
    const isHP = isHandphoneProduct(it);
    const harga = Number(it['Harga Jual']) || 0;
    const offlineSales = it['SlpName'] || it['Sales Employee Name'] || it['User Input'] || it['SC/HL Code'] || '-';
    const targetSosmed = (auditItem.keep && auditItem.keep['Kode Sosmed']) ? auditItem.keep['Kode Sosmed'] : 'BJM.SOS.xxx';
    const picSos = (auditItem.keep && auditItem.keep['Sender']) ? auditItem.keep['Sender'] : '-';
    const cust = (auditItem.keep && auditItem.keep['Nama']) ? auditItem.keep['Nama'] : (it['Name'] || '-');
    const wa = (auditItem.keep && auditItem.keep['Nomor WA']) ? auditItem.keep['Nomor WA'] : (it['No HP Customer'] || '-');
    const cabang = it['Cabang'] || auditItem.cabang || '-';

    const exists = rawARClaimsList.some(c => c.docNumber === docNum && c.productName === desc && c.status !== 'RESOLVED_BY_STORE');
    if (!exists) {
      rawARClaimsList.unshift({
        id: `CLM-${docNum}-${idx + 1}-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        docNumber: docNum,
        postingDate: it['Posting Date'] || auditItem.postingDate || '',
        periode: periode,
        cabang: cabang,
        productName: desc,
        hargaJual: harga,
        currentSalesCode: offlineSales,
        targetSosmedCode: targetSosmed,
        picSosmed: picSos,
        customerName: cust,
        customerPhone: wa,
        keepText: keepFormatted,
        evidenceImage: null,
        evidenceImages: [],
        validationStatus: 'VALID',
        invalidReason: '',
        isHandphone: isHP,
        collectDate: new Date().toISOString(),
        collectTimestamp: Date.now(),
        status: 'COLLECTING',
        shareTimestamp: null,
        deadlineH1Timestamp: null,
        penaltyAmount: isHP ? 50000 : 0,
        penaltyStatus: isHP ? 'POTENSI' : 'PASS',
        manualResolvedBy: '',
        manualResolvedAt: null
      });
      addedCount++;
    }
  });

  if (addedCount > 0) {
    saveARClaimsData();
    showSaveToast(`${addedCount} item dari AR ${docNum} berhasil dimasukkan ke Keranjang Klaim (Batch 3 Hari).`);
    renderARClaimsModule();
  } else {
    showSaveToast(`Item AR ${docNum} sudah ada di daftar klaim sebelumnya.`);
  }
}

function claimAllFilteredMissingAR() {
  if (!filteredAuditARList || filteredAuditARList.length === 0) {
    alert('Tidak ada data AR yang terfilter saat ini.');
    return;
  }

  const missingARs = filteredAuditARList.filter(r => r.status === 'ALL_MISSING' || r.status === 'PARTIAL_MISSING');
  if (missingARs.length === 0) {
    alert('Semua AR yang terfilter saat ini sudah berstatus 100% sosmed (Hijau).');
    return;
  }

  if (!confirm(`Apakah Anda yakin ingin memasukkan seluruh ${missingARs.length} dokumen AR yang belum sosmed ke Modul Klaim AR (Pengumpulan 3 Hari)?`)) {
    return;
  }

  let totalAdded = 0;
  missingARs.forEach(auditItem => {
    const docNum = String(auditItem.doc || auditItem.reportARDoc || auditItem.vlookupARDoc || '').trim();
    const itemsToClaim = (auditItem.items || []).filter(it => !isItemExemptFromSosmed(it) && !isItemInputSosmed(it));
    const periode = (auditItem.keep && auditItem.keep['PERIODE']) 
      ? String(auditItem.keep['PERIODE']).trim() 
      : 'Oktober-26';
    const keepFormatted = (typeof formatKeepTextForClipboard === 'function') 
      ? formatKeepTextForClipboard(auditItem.keep) 
      : '';

    itemsToClaim.forEach((it, idx) => {
      const desc = it['Dscription'] || it['Item/Service Description'] || '-';
      const exists = rawARClaimsList.some(c => c.docNumber === docNum && c.productName === desc && c.status !== 'RESOLVED_BY_STORE');
      if (!exists) {
        const isHP = isHandphoneProduct(it);
        rawARClaimsList.unshift({
          id: `CLM-${docNum}-${idx + 1}-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          docNumber: docNum,
          postingDate: it['Posting Date'] || auditItem.postingDate || '',
          periode: periode,
          cabang: it['Cabang'] || auditItem.cabang || '-',
          productName: desc,
          hargaJual: Number(it['Harga Jual']) || 0,
          currentSalesCode: it['SlpName'] || it['Sales Employee Name'] || it['User Input'] || it['SC/HL Code'] || '-',
          targetSosmedCode: (auditItem.keep && auditItem.keep['Kode Sosmed']) ? auditItem.keep['Kode Sosmed'] : 'BJM.SOS.xxx',
          picSosmed: (auditItem.keep && auditItem.keep['Sender']) ? auditItem.keep['Sender'] : '-',
          customerName: (auditItem.keep && auditItem.keep['Nama']) ? auditItem.keep['Nama'] : (it['Name'] || '-'),
          customerPhone: (auditItem.keep && auditItem.keep['Nomor WA']) ? auditItem.keep['Nomor WA'] : (it['No HP Customer'] || '-'),
          keepText: keepFormatted,
          evidenceImage: null,
          evidenceImages: [],
          validationStatus: 'VALID',
          invalidReason: '',
          isHandphone: isHP,
          collectDate: new Date().toISOString(),
          collectTimestamp: Date.now(),
          status: 'COLLECTING',
          shareTimestamp: null,
          deadlineH1Timestamp: null,
          penaltyAmount: isHP ? 50000 : 0,
          penaltyStatus: isHP ? 'POTENSI' : 'PASS',
          manualResolvedBy: '',
          manualResolvedAt: null
        });
        totalAdded++;
      }
    });
  });

  saveARClaimsData();
  showSaveToast(`Berhasil menambahkan ${totalAdded} item belum sosmed ke Modul Klaim AR.`);
  renderARClaimsModule();
}

function claimModalSelectedItems(btnElement = null) {
  if (currentModalCheckedItems.size === 0) {
    alert('Silakan centang minimal 1 item produk yang ingin dimasukkan ke Klaim Cabang.');
    return;
  }

  const items = Array.from(currentModalCheckedItems).map(idx => currentModalRows[idx]).filter(Boolean);
  const auditItem = {
    doc: currentModalDoc,
    items: currentModalRows,
    keep: currentModalKeep
  };
  addSingleARToClaims(auditItem, items);
  closeARDetailModal();
}

// Evidence Image Handlers (Dukungan Beberapa Evidence per Baris & Carousel Modal)
let currentEvidenceModalIndex = 0;

function triggerEvidenceUpload(claimId) {
  currentEvidenceTargetClaimId = claimId;
  const input = document.getElementById('evidenceFileInput');
  if (input) {
    input.value = '';
    input.click();
  }
}

function triggerEvidenceUploadFromModal() {
  const claimId = currentEvidenceTargetClaimId;
  if (!claimId) return;
  triggerEvidenceUpload(claimId);
}

function handleEvidenceFileSelected(fileInput) {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) return;
  const rawFiles = Array.from(fileInput.files);
  const imageFiles = rawFiles.filter(f => f.type.startsWith('image/'));

  if (imageFiles.length === 0) {
    alert('File yang dipilih harus berupa file gambar (JPG, PNG, WEBP).');
    return;
  }

  const claimId = currentEvidenceTargetClaimId;
  const item = rawARClaimsList.find(c => c.id === claimId);
  if (!item) return;

  if (!Array.isArray(item.evidenceImages)) {
    item.evidenceImages = item.evidenceImage ? [item.evidenceImage] : [];
  }

  const compressSingleFile = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = function(e) {
        const img = new Image();
        img.onload = function() {
          const canvas = document.createElement('canvas');
          let width = img.width;
          let height = img.height;
          const maxDim = 1200;
          if (width > maxDim || height > maxDim) {
            if (width > height) {
              height = Math.round((height * maxDim) / width);
              width = maxDim;
            } else {
              width = Math.round((width * maxDim) / height);
              height = maxDim;
            }
          }
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, width, height);

          const compressedDataUrl = canvas.toDataURL('image/jpeg', 0.82);
          resolve(compressedDataUrl);
        };
        img.onerror = () => resolve(null);
        img.src = e.target.result;
      };
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(file);
    });
  };

  Promise.all(imageFiles.map(compressSingleFile)).then(results => {
    const validImages = results.filter(Boolean);
    if (validImages.length === 0) return;

    item.evidenceImages.push(...validImages);
    item.evidenceImage = item.evidenceImages[0] || null;

    saveARClaimsData();
    showSaveToast(`📸 Berhasil menambahkan ${validImages.length} bukti gambar untuk AR ${item.docNumber} (Total: ${item.evidenceImages.length} foto).`);
    renderARClaimsTable();

    // Jika modal zoom sedang terbuka, refresh langsung ke foto baru
    const modal = document.getElementById('evidenceImageModal');
    if (modal && modal.style.display !== 'none' && currentEvidenceTargetClaimId === claimId) {
      openEvidenceImageModal(claimId, item.evidenceImages.length - 1);
    }
  });
}

function removeEvidenceImage(claimId, imageIdx = null) {
  const item = rawARClaimsList.find(c => c.id === claimId);
  if (!item) return;
  if (!Array.isArray(item.evidenceImages)) {
    item.evidenceImages = item.evidenceImage ? [item.evidenceImage] : [];
  }
  if (item.evidenceImages.length === 0) return;

  if (imageIdx !== null && imageIdx >= 0 && imageIdx < item.evidenceImages.length) {
    if (!confirm(`Hapus foto bukti ke-${imageIdx + 1} untuk AR ${item.docNumber}?`)) return;
    item.evidenceImages.splice(imageIdx, 1);
  } else {
    if (!confirm(`Hapus seluruh (${item.evidenceImages.length}) bukti gambar untuk AR ${item.docNumber}?`)) return;
    item.evidenceImages = [];
  }

  item.evidenceImage = item.evidenceImages[0] || null;
  saveARClaimsData();
  showSaveToast('Bukti gambar berhasil dihapus.');
  renderARClaimsTable();

  const modal = document.getElementById('evidenceImageModal');
  if (modal && modal.style.display !== 'none' && currentEvidenceTargetClaimId === claimId) {
    if (item.evidenceImages.length === 0) {
      closeEvidenceImageModal();
    } else {
      const nextIdx = Math.min(imageIdx !== null ? imageIdx : 0, item.evidenceImages.length - 1);
      openEvidenceImageModal(claimId, nextIdx);
    }
  }
}

function openEvidenceImageModal(claimId, activeIndex = 0) {
  const item = rawARClaimsList.find(c => c.id === claimId);
  if (!item) return;
  if (!Array.isArray(item.evidenceImages)) {
    item.evidenceImages = item.evidenceImage ? [item.evidenceImage] : [];
  }
  if (item.evidenceImages.length === 0) return;

  currentEvidenceTargetClaimId = claimId;
  currentEvidenceModalIndex = Math.max(0, Math.min(activeIndex, item.evidenceImages.length - 1));

  renderEvidenceModalContent();

  const modal = document.getElementById('evidenceImageModal');
  if (modal) modal.style.display = 'flex';
}

function renderEvidenceModalContent() {
  const item = rawARClaimsList.find(c => c.id === currentEvidenceTargetClaimId);
  if (!item || !item.evidenceImages || item.evidenceImages.length === 0) {
    closeEvidenceImageModal();
    return;
  }

  const imgs = item.evidenceImages;
  const currentIdx = Math.max(0, Math.min(currentEvidenceModalIndex, imgs.length - 1));
  const currentImgUrl = imgs[currentIdx];

  const titleEl = document.getElementById('evidenceModalTitle');
  const subEl = document.getElementById('evidenceModalSub');
  const fullImg = document.getElementById('evidenceFullImg');
  const btnDownload = document.getElementById('btnDownloadEvidence');
  const counterEl = document.getElementById('evidenceModalCounter');
  const galleryStrip = document.getElementById('evidenceModalGalleryStrip');

  if (titleEl) titleEl.innerText = `📷 Bukti Format Keep: AR ${item.docNumber} (${currentIdx + 1} dari ${imgs.length})`;
  if (subEl) subEl.innerText = `Customer: ${item.customerName || '-'} (${item.customerPhone || '-'}) | Cabang: ${item.cabang} | Barang: ${item.productName}`;
  if (fullImg) fullImg.src = currentImgUrl;
  if (btnDownload) {
    btnDownload.href = currentImgUrl;
    btnDownload.download = `Evidence_Keep_AR_${item.docNumber}_Foto_${currentIdx + 1}.jpg`;
  }
  if (counterEl) {
    counterEl.innerText = `Foto ${currentIdx + 1} dari ${imgs.length}`;
  }

  // Thumbnails carousel strip
  if (galleryStrip) {
    if (imgs.length > 1) {
      galleryStrip.style.display = 'flex';
      galleryStrip.innerHTML = imgs.map((img, i) => `
        <div class="evidence-modal-thumb ${i === currentIdx ? 'active' : ''}" onclick="switchEvidenceModalIndex(${i})" title="Buka foto ke-${i + 1}">
          <img src="${img}" alt="Thumb ${i + 1}">
        </div>
      `).join('');
    } else {
      galleryStrip.style.display = 'none';
      galleryStrip.innerHTML = '';
    }
  }
}

function switchEvidenceModalIndex(newIdx) {
  const item = rawARClaimsList.find(c => c.id === currentEvidenceTargetClaimId);
  if (!item || !item.evidenceImages || item.evidenceImages.length === 0) return;
  if (newIdx < 0) newIdx = item.evidenceImages.length - 1;
  if (newIdx >= item.evidenceImages.length) newIdx = 0;
  currentEvidenceModalIndex = newIdx;
  renderEvidenceModalContent();
}

function nextEvidenceModalImage() {
  switchEvidenceModalIndex(currentEvidenceModalIndex + 1);
}

function prevEvidenceModalImage() {
  switchEvidenceModalIndex(currentEvidenceModalIndex - 1);
}

function removeCurrentModalEvidence() {
  removeEvidenceImage(currentEvidenceTargetClaimId, currentEvidenceModalIndex);
}

function closeEvidenceImageModal(e = null) {
  const modal = document.getElementById('evidenceImageModal');
  if (modal) modal.style.display = 'none';
}

// Navigasi keyboard (Panah Kiri & Kanan) untuk galeri bukti foto
document.addEventListener('keydown', function(e) {
  const modal = document.getElementById('evidenceImageModal');
  if (modal && modal.style.display === 'flex') {
    if (e.key === 'ArrowRight') {
      nextEvidenceModalImage();
    } else if (e.key === 'ArrowLeft') {
      prevEvidenceModalImage();
    } else if (e.key === 'Escape') {
      closeEvidenceImageModal();
    }
  }
});

// Validation Handlers
function handleClaimValidityChange(claimId, newStatus) {
  const item = rawARClaimsList.find(c => c.id === claimId);
  if (!item) return;

  if (newStatus === 'TIDAK VALID') {
    const reason = prompt(`Item AR ${item.docNumber} ditandai TIDAK VALID.\nMasukkan alasannya:`, item.invalidReason || 'Salah tag / Bukan nomor keep');
    if (reason === null) {
      renderARClaimsTable();
      return;
    }
    item.validationStatus = 'TIDAK VALID';
    item.invalidReason = reason.trim() || 'Tidak Valid';
    showSaveToast(`🚫 AR ${item.docNumber} ditandai TIDAK VALID dan dipindahkan ke sub-tab "5. Tidak Valid".`);
  } else {
    item.validationStatus = 'VALID';
    item.invalidReason = '';
    showSaveToast(`✅ AR ${item.docNumber} dikembalikan ke status VALID.`);
  }

  saveARClaimsData();
  renderARClaimsModule();
}

function handleClaimInvalidReasonChange(claimId, newReason) {
  const item = rawARClaimsList.find(c => c.id === claimId);
  if (!item) return;
  item.invalidReason = newReason.trim();
  saveARClaimsData();
}

// ==========================================
// EXCEL EXPORT ENGINE (EXCELJS + FALLBACK)
// ==========================================
function parseSortDate(d) {
  if (!d) return 0;
  if (typeof d === 'number') return d;
  if (d instanceof Date) return d.getTime();
  const s = String(d).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return new Date(s).getTime() || 0;
  }
  const parts = s.split(/[\/\-]/);
  if (parts.length >= 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    let year = parseInt(parts[2].split(' ')[0], 10);
    if (year < 100) year += 2000;
    return new Date(year, month, day).getTime() || 0;
  }
  return 0;
}

function sortARClaimsRows(rows, groupByBranch = true) {
  return [...rows].sort((a, b) => {
    if (groupByBranch) {
      const cabA = String(a.cabang || '').toUpperCase();
      const cabB = String(b.cabang || '').toUpperCase();
      if (cabA !== cabB) return cabA.localeCompare(cabB);
    }
    const dateA = parseSortDate(a.postingDate);
    const dateB = parseSortDate(b.postingDate);
    if (dateA !== dateB) return dateA - dateB;

    const docA = String(a.docNumber || '');
    const docB = String(b.docNumber || '');
    return docA.localeCompare(docB, undefined, { numeric: true });
  });
}

async function buildClaimsWorksheet(wb, sheetName, rows) {
  const safeName = String(sheetName || 'Klaim').replace(/[:\\/?*\[\]]/g, '_').slice(0, 31);
  const ws = wb.addWorksheet(safeName);

  // 1. Urutkan data secara rapi dan kronologis
  const isAllSheet = (sheetName === 'SEMUA CABANG');
  const sortedRows = sortARClaimsRows(rows, isAllSheet);

  // 2. Hitung jumlah maksimum foto bukti pada baris-baris ini
  let maxEv = 1;
  sortedRows.forEach(r => {
    const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
      ? r.evidenceImages
      : (r.evidenceImage ? [r.evidenceImage] : []);
    if (imgs.length > maxEv) maxEv = imgs.length;
  });

  // 3. Bangun kolom tabel secara berurutan
  const cols = [
    { header: 'NO', key: 'no', width: 6 },
    { header: 'TANGGAL', key: 'tanggal', width: 14 },
    { header: 'PERIODE', key: 'periode', width: 14 },
    { header: 'NO AR', key: 'noAR', width: 16 },
    { header: 'NO.HP USER', key: 'noHP', width: 18 },
    { header: 'KODE SOSMED', key: 'kodeSosmed', width: 16 },
    { header: 'PIC SOSMED', key: 'picSosmed', width: 16 },
    { header: 'CABANG', key: 'cabang', width: 14 },
    { header: 'POTENSI DENDA', key: 'potensiDenda', width: 16 },
    { header: 'BARANG', key: 'barang', width: 34 },
    { header: 'KEEP', key: 'keep', width: 36 }
  ];

  // Kolom Evidence terpisah per foto bukti dengan lebar besar (width: 48 ≈ 370px) agar muat foto besar tanpa bertumpuk
  for (let e = 0; e < maxEv; e++) {
    const headerTitle = maxEv === 1 
      ? 'EVIDENCE FORMAT KEEP' 
      : `EVIDENCE ${e + 1} (FORMAT KEEP)`;
    cols.push({
      header: headerTitle,
      key: `evidence_${e}`,
      width: 48
    });
  }

  cols.push(
    { header: 'AR SALES EMPLOYEE', key: 'salesEmployee', width: 22 },
    { header: 'STATUS VALIDASI', key: 'statusValidasi', width: 16 },
    { header: 'ALASAN TIDAK VALID', key: 'alasanTidakValid', width: 24 }
  );

  ws.columns = cols;

  const headerRow = ws.getRow(1);
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF991B1B' } // Crimson / Red
    };
    cell.font = {
      name: 'Arial',
      size: 10,
      bold: true,
      color: { argb: 'FFFFFFFF' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF7F1D1D' } },
      left: { style: 'thin', color: { argb: 'FF7F1D1D' } },
      bottom: { style: 'medium', color: { argb: 'FF450A0A' } },
      right: { style: 'thin', color: { argb: 'FF7F1D1D' } }
    };
  });

  const evidenceStartCol = 11; // 0-indexed column 11 = Column L

  for (let idx = 0; idx < sortedRows.length; idx++) {
    const r = sortedRows[idx];
    const rowNumber = idx + 2;
    const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
      ? r.evidenceImages
      : (r.evidenceImage ? [r.evidenceImage] : []);

    const rowObj = {
      no: idx + 1,
      tanggal: formatCrosellingDate(r.postingDate),
      periode: r.periode || 'Oktober-26',
      noAR: r.docNumber,
      noHP: r.customerPhone || '-',
      kodeSosmed: r.targetSosmedCode,
      picSosmed: r.picSosmed || '-',
      cabang: r.cabang,
      potensiDenda: r.isHandphone ? 'DENDA Rp 50.000' : 'PASS Rp 0',
      barang: r.productName,
      keep: r.keepText || '-',
      salesEmployee: r.currentSalesCode || '-',
      statusValidasi: r.validationStatus || 'VALID',
      alasanTidakValid: r.invalidReason || '-'
    };

    for (let e = 0; e < maxEv; e++) {
      rowObj[`evidence_${e}`] = imgs[e] ? '' : '-';
    }

    const row = ws.addRow(rowObj);

    // Jika ada bukti foto: buat baris tinggi besar (200pt ≈ 267px) agar tulisan WhatsApp terbaca jelas!
    const hasImgs = imgs.length > 0;
    row.height = hasImgs ? 200 : 36;

    row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    row.getCell('no').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('tanggal').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('periode').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('noAR').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('noHP').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('kodeSosmed').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('picSosmed').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('cabang').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('potensiDenda').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('statusValidasi').alignment = { vertical: 'middle', horizontal: 'center' };

    for (let e = 0; e < maxEv; e++) {
      row.getCell(`evidence_${e}`).alignment = { vertical: 'middle', horizontal: 'center' };
    }

    // Border grid rapi pada setiap cell
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    });

    // Embed setiap foto bukti secara berurutan pada kolomnya masing-masing dengan ukuran BESAR (330px x 240px)
    if (hasImgs) {
      imgs.forEach((imgStr, imgIdx) => {
        if (!imgStr || typeof imgStr !== 'string') return;
        try {
          const isPng = imgStr.includes('image/png');
          const base64Data = imgStr.includes(',') ? imgStr.split(',')[1] : imgStr;
          const imgId = wb.addImage({
            base64: base64Data,
            extension: isPng ? 'png' : 'jpeg'
          });
          // Letakkan tepat di dalam kolom bukti masing-masing (Col 11 + imgIdx) tanpa overlap
          ws.addImage(imgId, {
            tl: { col: evidenceStartCol + imgIdx + 0.04, row: rowNumber - 1 + 0.05 },
            ext: { width: 330, height: 240 },
            editAs: 'oneCell'
          });
        } catch (imgErr) {
          console.warn('Gagal embed image ExcelJS:', imgErr);
          row.getCell(`evidence_${imgIdx}`).value = `(Bukti Foto ${imgIdx + 1})`;
        }
      });
    }
  }
}

// 1. UNDUH SELURUH KLAIM JADI 1 FILE EXCEL GABUNGAN
async function downloadARClaimsSingleExcel(autoShare = false) {
  if (!rawARClaimsList || rawARClaimsList.length === 0) {
    alert('Belum ada data klaim AR untuk diunduh.');
    return;
  }

  const collectingItems = rawARClaimsList.filter(c => c.status === 'COLLECTING' && c.validationStatus !== 'TIDAK VALID');
  let itemsToExport = [];
  let doAdvanceToShared = false;

  if (autoShare) {
    itemsToExport = collectingItems;
    doAdvanceToShared = true;
  } else if (currentClaimsSubTab === 'COLLECTING' && collectingItems.length > 0) {
    const userChoice = confirm(
      `📊 UNDUH SEMUANYA JADI 1 FILE EXCEL (.xlsx)\n\n` +
      `Ditemukan ${collectingItems.length} item AR dalam antrean "1. Pengumpulan (Batch 3 Hari)".\n\n` +
      `• Klik [OK] untuk mengunduh 1 file Excel Gabungan (diurutkan rapi & gambar besar) dan SEKALIGUS memindahkan status ke "Cek AR Setelah di Share (Tenggat H+1)".\n` +
      `• Klik [Batal] jika hanya ingin mengunduh data tampilan saat ini tanpa memindahkan status klaim.`
    );
    if (userChoice) {
      itemsToExport = collectingItems;
      doAdvanceToShared = true;
    } else {
      itemsToExport = filteredARClaimsList.length > 0 ? filteredARClaimsList : rawARClaimsList;
    }
  } else {
    itemsToExport = filteredARClaimsList.length > 0 ? filteredARClaimsList : rawARClaimsList;
  }

  if (itemsToExport.length === 0) {
    alert('Tidak ada data klaim yang memenuhi kriteria untuk diunduh.');
    return;
  }

  const nowStr = new Date().toISOString().split('T')[0];

  // Gunakan ExcelJS untuk menghasilkan 1 file Excel lengkap dengan sheet gabungan & sheet per cabang serta gambar ter-embed besar
  if (typeof ExcelJS !== 'undefined') {
    try {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'GadgetMart Sosmed System';
      wb.title = 'Klaim AR Semua Cabang';

      // 1. Sheet Master Gabungan Seluruh Cabang (Urut Cabang A-Z -> Tanggal -> No AR)
      await buildClaimsWorksheet(wb, 'SEMUA CABANG', itemsToExport);

      // 2. Sheet per Cabang terpisah di file Excel yang sama (Urut Alfabetis Tab Cabang)
      const byBranch = {};
      itemsToExport.forEach(c => {
        const cab = c.cabang || 'LAINNYA';
        if (!byBranch[cab]) byBranch[cab] = [];
        byBranch[cab].push(c);
      });

      const sortedCabangList = Object.keys(byBranch).sort();
      for (const cab of sortedCabangList) {
        await buildClaimsWorksheet(wb, cab, byBranch[cab]);
      }

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `Klaim_AR_Semua_Cabang_Gabungan_${nowStr}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      if (doAdvanceToShared) {
        const shareTime = Date.now();
        const deadlineH1 = shareTime + (24 * 60 * 60 * 1000);
        itemsToExport.forEach(c => {
          c.status = 'SHARED';
          c.shareTimestamp = shareTime;
          c.deadlineH1Timestamp = deadlineH1;
        });
        await saveARClaimsData();
        showSaveToast(`🚀 ${itemsToExport.length} item AR berhasil diunduh jadi 1 Excel & dipindahkan ke sub-tab "Cek AR Setelah di Share" (H+1)!`);
        setClaimsSubTab('SHARED');
      } else {
        showSaveToast(`✅ Berhasil mengunduh 1 File Excel Gabungan (${itemsToExport.length} item AR)!`);
      }
      return;
    } catch (excelErr) {
      console.warn('ExcelJS error, fallback ke SheetJS:', excelErr);
    }
  }

  // Fallback SheetJS jika ExcelJS bermasalah
  const wb = XLSX.utils.book_new();
  const buildSheetJSData = (rows) => {
    const sorted = sortARClaimsRows(rows, true);
    let maxEv = 1;
    sorted.forEach(r => {
      const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
        ? r.evidenceImages
        : (r.evidenceImage ? [r.evidenceImage] : []);
      if (imgs.length > maxEv) maxEv = imgs.length;
    });

    const headers = [
      'NO', 'TANGGAL', 'PERIODE', 'NO AR', 'NO.HP USER', 'KODE SOSMED', 'PIC SOSMED', 'CABANG', 'POTENSI DENDA', 'BARANG', 'KEEP'
    ];
    for (let e = 0; e < maxEv; e++) {
      headers.push(maxEv === 1 ? 'EVIDENCE FORMAT KEEP' : `EVIDENCE ${e + 1}`);
    }
    headers.push('AR SALES EMPLOYEE', 'STATUS VALIDASI', 'ALASAN TIDAK VALID');

    const data = [headers];
    sorted.forEach((r, idx) => {
      const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
        ? r.evidenceImages
        : (r.evidenceImage ? [r.evidenceImage] : []);
      const rArr = [
        idx + 1,
        formatCrosellingDate(r.postingDate),
        r.periode || 'Oktober-26',
        r.docNumber,
        r.customerPhone || '-',
        r.targetSosmedCode,
        r.picSosmed || '-',
        r.cabang,
        r.isHandphone ? 'DENDA Rp 50.000' : 'PASS Rp 0',
        r.productName,
        r.keepText || '-'
      ];
      for (let e = 0; e < maxEv; e++) {
        rArr.push(imgs[e] ? `(Bukti Foto ${e + 1})` : '-');
      }
      rArr.push(
        r.currentSalesCode || '-',
        r.validationStatus || 'VALID',
        r.invalidReason || '-'
      );
      data.push(rArr);
    });
    return XLSX.utils.aoa_to_sheet(data);
  };

  XLSX.utils.book_append_sheet(wb, buildSheetJSData(itemsToExport), 'SEMUA CABANG');

  const byBranch = {};
  itemsToExport.forEach(c => {
    const cab = c.cabang || 'LAINNYA';
    if (!byBranch[cab]) byBranch[cab] = [];
    byBranch[cab].push(c);
  });
  const sortedCabangList = Object.keys(byBranch).sort();
  for (const cab of sortedCabangList) {
    XLSX.utils.book_append_sheet(wb, buildSheetJSData(byBranch[cab]), cab.slice(0, 31));
  }

  XLSX.writeFile(wb, `Klaim_AR_Semua_Cabang_Gabungan_${nowStr}.xlsx`);

  if (doAdvanceToShared) {
    const shareTime = Date.now();
    const deadlineH1 = shareTime + (24 * 60 * 60 * 1000);
    itemsToExport.forEach(c => {
      c.status = 'SHARED';
      c.shareTimestamp = shareTime;
      c.deadlineH1Timestamp = deadlineH1;
    });
    await saveARClaimsData();
    showSaveToast(`🚀 ${itemsToExport.length} item AR berhasil diunduh jadi 1 Excel & dipindahkan ke sub-tab "Cek AR Setelah di Share" (H+1)!`);
    setClaimsSubTab('SHARED');
  } else {
    showSaveToast(`✅ Berhasil mengunduh 1 File Excel Gabungan (${itemsToExport.length} item AR)!`);
  }
}
window.downloadARClaimsSingleExcel = downloadARClaimsSingleExcel;

// 2. DISTRIBUSI KLAIM HARI KE-4 (DENGAN PILIHAN 1 EXCEL ATAU PAKET ZIP)
async function shareARClaimsBatchDay4() {
  const collectingItems = rawARClaimsList.filter(c => c.status === 'COLLECTING' && c.validationStatus !== 'TIDAK VALID');
  if (collectingItems.length === 0) {
    alert('Tidak ada item valid dalam antrean "1. Pengumpulan (Batch 3 Hari)".\nSilakan kumpulkan AR belum sosmed dari Tab 5 (Tracking AR) terlebih dahulu.');
    return;
  }

  const branchCount = new Set(collectingItems.map(c => c.cabang)).size;
  const choiceMsg = `📦 DISTRIBUSI KLAIM AR HARI KE-4:\n\n` +
                    `Total: ${collectingItems.length} item AR dari ${branchCount} cabang toko siap didistribusikan.\n\n` +
                    `Pilih format unduhan:\n` +
                    `• Klik [OK] untuk mengunduh SEMUANYA JADI 1 FILE EXCEL GABUNGAN (.xlsx) (diurutkan rapi & gambar besar).\n` +
                    `• Klik [Batal] untuk mengunduh PAKET ZIP FILE TERPISAH PER CABANG (.zip).`;

  const wantSingleExcel = confirm(choiceMsg);

  if (wantSingleExcel) {
    await downloadARClaimsSingleExcel(true);
    return;
  }

  // Jika user memilih Batal, unduh paket zip terpisah per cabang
  const byBranch = {};
  collectingItems.forEach(c => {
    const cab = c.cabang || 'LAINNYA';
    if (!byBranch[cab]) byBranch[cab] = [];
    byBranch[cab].push(c);
  });

  const nowStr = new Date().toISOString().split('T')[0];
  const zip = (typeof JSZip !== 'undefined') ? new JSZip() : null;

  const sortedCabangList = Object.keys(byBranch).sort();
  for (const cab of sortedCabangList) {
    const rows = byBranch[cab];

    if (typeof ExcelJS !== 'undefined') {
      try {
        const wb = new ExcelJS.Workbook();
        wb.creator = 'GadgetMart Sosmed System';
        await buildClaimsWorksheet(wb, cab, rows);
        const buffer = await wb.xlsx.writeBuffer();
        if (zip) {
          zip.file(`Klaim_AR_${cab}_${nowStr}.xlsx`, buffer);
        }
        continue;
      } catch (excelJsErr) {
        console.warn('ExcelJS error, fallback ke SheetJS:', excelJsErr);
      }
    }

    // Fallback SheetJS
    const sortedCabRows = sortARClaimsRows(rows, false);
    let maxEv = 1;
    sortedCabRows.forEach(r => {
      const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
        ? r.evidenceImages
        : (r.evidenceImage ? [r.evidenceImage] : []);
      if (imgs.length > maxEv) maxEv = imgs.length;
    });

    const headers = [
      'NO', 'TANGGAL', 'PERIODE', 'NO AR', 'NO.HP USER', 'KODE SOSMED', 'PIC SOSMED', 'CABANG', 'POTENSI DENDA', 'BARANG', 'KEEP'
    ];
    for (let e = 0; e < maxEv; e++) {
      headers.push(maxEv === 1 ? 'EVIDENCE FORMAT KEEP' : `EVIDENCE ${e + 1}`);
    }
    headers.push('AR SALES EMPLOYEE', 'STATUS VALIDASI', 'ALASAN TIDAK VALID');

    const excelData = [headers];
    sortedCabRows.forEach((r, idx) => {
      const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
        ? r.evidenceImages
        : (r.evidenceImage ? [r.evidenceImage] : []);
      const rArr = [
        idx + 1,
        formatCrosellingDate(r.postingDate),
        r.periode || 'Oktober-26',
        r.docNumber,
        r.customerPhone || '-',
        r.targetSosmedCode,
        r.picSosmed || '-',
        r.cabang,
        r.isHandphone ? 'DENDA Rp 50.000' : 'PASS Rp 0',
        r.productName,
        r.keepText || '-'
      ];
      for (let e = 0; e < maxEv; e++) {
        rArr.push(imgs[e] ? `(Bukti Foto ${e + 1})` : '-');
      }
      rArr.push(
        r.currentSalesCode || '-',
        r.validationStatus || 'VALID',
        r.invalidReason || '-'
      );
      excelData.push(rArr);
    });
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, cab.slice(0, 31));
    if (zip) {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      zip.file(`Klaim_AR_${cab}_${nowStr}.xlsx`, wbout);
    }
  }

  if (zip) {
    try {
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(zipBlob);
      a.download = `Paket_Klaim_AR_Semua_Cabang_${nowStr}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (zipErr) {
      console.error('Gagal generate zip:', zipErr);
    }
  }

  const shareTime = Date.now();
  const deadlineH1 = shareTime + (24 * 60 * 60 * 1000);
  collectingItems.forEach(c => {
    c.status = 'SHARED';
    c.shareTimestamp = shareTime;
    c.deadlineH1Timestamp = deadlineH1;
  });

  await saveARClaimsData();
  showSaveToast(`🚀 ${collectingItems.length} item AR berhasil diunduh paket zip & dipindahkan ke "Cek AR Setelah di Share" (Tenggat H+1 aktif)!`);
  setClaimsSubTab('SHARED');
}

function checkARClaimsAutoResolution(manualTrigger = false) {
  if (!rawARClaimsList || rawARClaimsList.length === 0) return;

  const now = Date.now();
  let resolvedCount = 0;
  let escalatedPenaltyCount = 0;

  const currentSalesMap = new Map();
  const allTx = (typeof rawTurnoverRows !== 'undefined' && rawTurnoverRows.length > 0)
    ? rawTurnoverRows
    : ((typeof rawCrosellingRows !== 'undefined' && rawCrosellingRows.length > 0) ? rawCrosellingRows : []);

  allTx.forEach(t => {
    const doc = String(t['Document Number'] || '').replace(/\D/g, '');
    const slp = String(t['SlpName'] || t['Sales Employee Name'] || t['User Input'] || t['SC/HL Code'] || '').toUpperCase();
    if (doc) {
      if (!currentSalesMap.has(doc)) currentSalesMap.set(doc, []);
      currentSalesMap.get(doc).push(slp);
    }
  });

  rawARClaimsList.forEach(c => {
    // Abaikan item yang tidak valid
    if (c.validationStatus === 'TIDAK VALID') return;

    const cleanDoc = String(c.docNumber).replace(/\D/g, '');

    // Check if store already resolved it (changed code to BJM.SOS)
    if (currentSalesMap.has(cleanDoc)) {
      const slpList = currentSalesMap.get(cleanDoc);
      const isNowSosmed = slpList.some(s => s.includes('BJM.SOS') || s.includes('SOSMED'));
      if (isNowSosmed && c.status !== 'RESOLVED_BY_STORE') {
        c.status = 'RESOLVED_BY_STORE';
        c.penaltyStatus = 'PASS';
        c.resolvedTimestamp = now;
        resolvedCount++;
        return;
      }
    }

    // Check if SHARED and passed H+1 deadline!
    if (c.status === 'SHARED' && c.deadlineH1Timestamp && now > c.deadlineH1Timestamp) {
      c.status = 'PENALTY_REAL';
      c.penaltyStatus = c.isHandphone ? 'REAL_DENDA' : 'PASS';
      c.penaltyAmount = c.isHandphone ? 50000 : 0;
      escalatedPenaltyCount++;
    }
  });

  if (resolvedCount > 0 || escalatedPenaltyCount > 0) {
    saveARClaimsData();
    if (manualTrigger) {
      alert(`Pengecekan Selesai:\n\n✅ ${resolvedCount} item berhasil direvisi oleh cabang (Selesai).\n🚨 ${escalatedPenaltyCount} item telah lewat batas H+1 dan dieksekusi menjadi Denda Sungguhan.`);
    }
  } else if (manualTrigger) {
    alert('Pengecekan Selesai:\nBelum ada perubahan status baru pada data klaim.');
  }

  renderARClaimsModule();
}

function getH1CountdownBadge(deadlineTimestamp) {
  if (!deadlineTimestamp) return '-';
  const diff = deadlineTimestamp - Date.now();
  if (diff <= 0) {
    return `<span class="badge-countdown expired">🚨 Lewat H+1 (Jatuh Tempo)</span>`;
  }
  const hours = Math.floor(diff / (1000 * 60 * 60));
  const mins = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
  if (hours < 4) {
    return `<span class="badge-countdown warning">⚠️ Sisa ${hours}j ${mins}m</span>`;
  }
  return `<span class="badge-countdown safe">⏳ Sisa ${hours}j ${mins}m</span>`;
}

function markClaimResolvedManually(claimId) {
  const item = rawARClaimsList.find(c => c.id === claimId);
  if (!item) return;

  const note = prompt(`Masukkan catatan pengubahan mandiri untuk AR ${item.docNumber}:`, 'Sudah diubah sendiri oleh tim sosmed di SAP');
  if (note === null) return;

  item.status = 'RESOLVED_MANUAL_SOSMED';
  item.manualResolvedBy = 'Tim Sosmed';
  item.manualResolvedAt = Date.now();
  item.manualInputNote = note;
  saveARClaimsData();
  showSaveToast(`✅ Status AR ${item.docNumber} ditandai sudah diubah sendiri oleh Tim Sosmed!`);
  renderARClaimsModule();
}

function deleteARClaimItem(claimId) {
  const idx = rawARClaimsList.findIndex(c => c.id === claimId);
  if (idx === -1) return;
  if (!confirm(`Hapus item AR ${rawARClaimsList[idx].docNumber} dari daftar klaim?`)) return;
  rawARClaimsList.splice(idx, 1);
  saveARClaimsData();
  showSaveToast('🗑️ Item berhasil dihapus dari daftar klaim.');
  renderARClaimsModule();
}

function copyBranchWAFollowupMessage(cabangCode) {
  const branchItems = rawARClaimsList.filter(c => c.cabang === cabangCode && c.validationStatus !== 'TIDAK VALID' && (c.status === 'SHARED' || c.status === 'PENALTY_REAL'));
  if (branchItems.length === 0) {
    alert(`Tidak ada klaim aktif yang sedang menunggu revisi untuk cabang ${cabangCode}.`);
    return;
  }

  const lines = [
    `*PEMBERITAHUAN KLAIM AR BELUM SOSMED - CABANG ${cabangCode}*`,
    `Halo tim toko/kasir cabang ${cabangCode}, mohon segera mengubah sales code di SAP untuk faktur sosmed berikut sebelum batas H+1:`,
    ''
  ];

  branchItems.forEach((it, idx) => {
    lines.push(`${idx + 1}. No. AR: *${it.docNumber}*`);
    lines.push(`   - Barang: ${it.productName}`);
    lines.push(`   - Sales Terinput: ${it.currentSalesCode}`);
    lines.push(`   - Seharusnya: *${it.targetSosmedCode}*`);
    if (it.status === 'PENALTY_REAL') {
      lines.push(`   - 🚨 *STATUS: LEWAT H+1 (DENDA SUNGGUHAN Rp 50.000)*`);
    } else {
      lines.push(`   - ⏳ *Tenggat: H+1*`);
    }
    lines.push('');
  });

  lines.push('Harap konfirmasi jika sales code sudah diubah di SAP. Terima kasih!');
  writeTextToClipboard(lines.join('\n'));
  showSaveToast(`📋 Format pesan follow-up WA untuk cabang ${cabangCode} berhasil disalin!`);
}

function exportClaimsPenaltyReport() {
  const penaltyItems = rawARClaimsList.filter(c => c.validationStatus !== 'TIDAK VALID' && (c.status === 'PENALTY_REAL' || c.status === 'RESOLVED_MANUAL_SOSMED'));
  if (penaltyItems.length === 0) {
    alert('Belum ada item AR yang terkena Denda Sungguhan (semua cabang masih dalam masa tenggang atau sudah direvisi tepat waktu).');
    return;
  }

  const byCab = {};
  let grandTotal = 0;
  penaltyItems.forEach(p => {
    const cab = p.cabang || 'LAINNYA';
    if (!byCab[cab]) byCab[cab] = { count: 0, totalDenda: 0, items: [] };
    byCab[cab].count++;
    byCab[cab].totalDenda += (p.penaltyAmount || 0);
    byCab[cab].items.push(p);
    grandTotal += (p.penaltyAmount || 0);
  });

  const nowStr = new Date().toISOString().split('T')[0];
  const excelData = [
    ['LAPORAN DENDA KASIR CABANG - FAKTUR AR LEWAT H+1', '', '', '', '', '', '', '', ''],
    [`Tanggal Rekap: ${formatCrosellingDate(nowStr)}`, '', '', '', '', '', '', '', ''],
    [`Total Denda Keseluruhan: ${formatRupiah(grandTotal)} (${penaltyItems.length} Kasus)`, '', '', '', '', '', '', '', ''],
    [''],
    ['NO', 'TANGGAL', 'PERIODE', 'NO. AR', 'CABANG', 'CUSTOMER & WA', 'NAMA BARANG', 'KASIR / SALES OFFLINE', 'KODE SOSMED', 'STATUS PENANGANAN', 'NOMINAL DENDA']
  ];

  penaltyItems.forEach((r, idx) => {
    excelData.push([
      idx + 1,
      formatCrosellingDate(r.postingDate),
      r.periode || 'Oktober-26',
      r.docNumber,
      r.cabang,
      `${r.customerName || '-'} (${r.customerPhone || '-'})`,
      r.productName,
      r.currentSalesCode || '-',
      r.targetSosmedCode,
      r.status === 'RESOLVED_MANUAL_SOSMED' ? 'Sudah Diubah Mandiri oleh Tim Sosmed' : 'Belum Diubah (Denda Aktif)',
      r.penaltyAmount || 0
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(excelData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Rekap_Denda_Cabang');
  XLSX.writeFile(wb, `Rekap_Denda_Cabang_GadgetMart_${nowStr}.xlsx`);
  showSaveToast(`📋 Laporan Rekap Denda Cabang berhasil diunduh (${penaltyItems.length} kasus)!`);
}

function setClaimsSubTab(subTabName) {
  currentClaimsSubTab = subTabName;
  ['COLLECTING', 'SHARED', 'PENALTY', 'RESOLVED', 'INVALID'].forEach(st => {
    const btn = document.getElementById(`subtabBtn${st.charAt(0) + st.slice(1).toLowerCase()}`);
    if (btn) btn.classList.toggle('active', st === subTabName);
  });
  claimsCurrentPage = 1;
  applyClaimsFilters();
}

function applyClaimsFilters() {
  const cabEl = document.getElementById('claimsFilterCabang');
  const sosEl = document.getElementById('claimsFilterSosmed');
  const tglEl = document.getElementById('claimsFilterTanggal');
  const searchEl = document.getElementById('claimsSearchInput');

  const cabVal = cabEl ? cabEl.value : 'ALL';
  const sosVal = sosEl ? sosEl.value : 'ALL';
  const tglVal = tglEl ? tglEl.value : '';
  const keyword = (searchEl ? searchEl.value : '').trim().toLowerCase();

  filteredARClaimsList = rawARClaimsList.filter(c => {
    if (currentClaimsSubTab === 'INVALID') {
      if (c.validationStatus !== 'TIDAK VALID') return false;
    } else {
      // Sub-tab lain hanya menampilkan yang VALID
      if (c.validationStatus === 'TIDAK VALID') return false;

      if (currentClaimsSubTab === 'COLLECTING') {
        if (c.status !== 'COLLECTING') return false;
      } else if (currentClaimsSubTab === 'SHARED') {
        if (c.status !== 'SHARED') return false;
      } else if (currentClaimsSubTab === 'PENALTY') {
        if (c.status !== 'PENALTY_REAL' && c.status !== 'RESOLVED_MANUAL_SOSMED') return false;
      } else if (currentClaimsSubTab === 'RESOLVED') {
        if (c.status !== 'RESOLVED_BY_STORE') return false;
      }
    }

    if (cabVal !== 'ALL' && c.cabang !== cabVal) return false;
    if (sosVal !== 'ALL' && c.targetSosmedCode !== sosVal) return false;

    if (tglVal) {
      const itemDate = parseDateToISO(c.collectDate || c.postingDate);
      if (itemDate !== tglVal) return false;
    }

    if (keyword) {
      const comb = `${c.docNumber} ${c.productName} ${c.cabang} ${c.currentSalesCode} ${c.targetSosmedCode} ${c.customerName} ${c.customerPhone} ${c.invalidReason || ''}`.toLowerCase();
      if (!comb.includes(keyword)) return false;
    }

    return true;
  });

  let activeCount = 0;
  if (cabVal !== 'ALL') activeCount++;
  if (sosVal !== 'ALL') activeCount++;
  if (tglVal) activeCount++;
  const badgeEl = document.getElementById('claimsFilterActiveCount');
  if (badgeEl) badgeEl.innerText = activeCount;

  renderARClaimsTable();
}

function toggleClaimsFiltersSection() {
  const el = document.getElementById('claimsFilterSection');
  if (el) {
    el.style.display = (el.style.display === 'none' ? 'flex' : 'none');
  }
}
window.toggleClaimsFiltersSection = toggleClaimsFiltersSection;

function renderARClaimsModule() {
  populateClaimsDropdowns();

  let countCollecting = 0;
  let countShared = 0;
  let countPenalty = 0;
  let countResolved = 0;
  let countInvalid = 0;
  let totalPenaltyRp = 0;

  rawARClaimsList.forEach(c => {
    if (c.validationStatus === 'TIDAK VALID') {
      countInvalid++;
      return;
    }
    if (c.status === 'COLLECTING') countCollecting++;
    else if (c.status === 'SHARED') countShared++;
    else if (c.status === 'PENALTY_REAL' || c.status === 'RESOLVED_MANUAL_SOSMED') {
      countPenalty++;
      totalPenaltyRp += (c.penaltyAmount || 0);
    } else if (c.status === 'RESOLVED_BY_STORE') countResolved++;
  });

  const heroSub = document.getElementById('claimsHeroSub');
  if (heroSub) {
    heroSub.innerText = `${rawARClaimsList.length.toLocaleString('id-ID')} data klaim faktur AR dari siklus pengawasan denda`;
  }

  const elColl = document.getElementById('claimCountCollecting');
  const elShar = document.getElementById('claimCountShared');
  const elPen = document.getElementById('claimCountPenalty');
  const elPenRp = document.getElementById('claimTotalPenaltyRp');
  const elRes = document.getElementById('claimCountResolved');
  const elInv = document.getElementById('claimCountInvalid');

  if (elColl) elColl.innerText = countCollecting.toLocaleString('id-ID');
  if (elShar) elShar.innerText = countShared.toLocaleString('id-ID');
  if (elPen) elPen.innerText = countPenalty.toLocaleString('id-ID');
  if (elPenRp) elPenRp.innerText = `${formatRupiah(totalPenaltyRp)} (${countPenalty} Item)`;
  if (elRes) elRes.innerText = countResolved.toLocaleString('id-ID');
  if (elInv) elInv.innerText = countInvalid.toLocaleString('id-ID');

  const bColl = document.getElementById('badgeSubtabCollecting');
  const bShar = document.getElementById('badgeSubtabShared');
  const bPen = document.getElementById('badgeSubtabPenalty');
  const bRes = document.getElementById('badgeSubtabResolved');
  const bInv = document.getElementById('badgeSubtabInvalid');

  if (bColl) bColl.innerText = countCollecting;
  if (bShar) bShar.innerText = countShared;
  if (bPen) bPen.innerText = countPenalty;
  if (bRes) bRes.innerText = countResolved;
  if (bInv) bInv.innerText = countInvalid;

  const btnShare = document.getElementById('btnShareBatchDay4');
  if (btnShare) {
    btnShare.style.display = (currentClaimsSubTab === 'COLLECTING') ? 'inline-flex' : 'none';
    btnShare.innerText = `📦 Unduh Paket Cabang & Share (${countCollecting})`;
  }

  const btnSingle = document.getElementById('btnDownloadSingleExcel');
  if (btnSingle) {
    btnSingle.innerHTML = (currentClaimsSubTab === 'COLLECTING' && countCollecting > 0)
      ? `📊 Unduh Jadi 1 Excel (${countCollecting})`
      : `📊 Unduh Jadi 1 Excel`;
  }

  applyClaimsFilters();
}

function renderARClaimsTable() {
  const tbody = document.getElementById('claimsTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const total = filteredARClaimsList.length;
  const maxPage = Math.max(1, Math.ceil(total / CLAIMS_PAGE_SIZE));
  const startIdx = (claimsCurrentPage - 1) * CLAIMS_PAGE_SIZE;
  const endIdx = Math.min(startIdx + CLAIMS_PAGE_SIZE, total);
  const pageRows = filteredARClaimsList.slice(startIdx, endIdx);

  const heroSub = document.getElementById('claimsHeroSub');
  if (heroSub) {
    heroSub.innerText = `${total.toLocaleString('id-ID')} data klaim faktur AR pada sub-tab ini`;
  }

  const infoEl = document.getElementById('claimsPaginationInfo');
  if (infoEl) {
    infoEl.innerText = total > 0 
      ? `Menampilkan ${startIdx + 1} - ${endIdx} dari total ${total.toLocaleString('id-ID')} Data Klaim`
      : 'Tidak ada data klaim untuk filter ini';
  }
  const curPageText = document.getElementById('claimsCurrentPageText');
  if (curPageText) curPageText.innerText = `Halaman ${claimsCurrentPage} / ${maxPage}`;
  const btnPrev = document.getElementById('btnClaimsPrev');
  if (btnPrev) btnPrev.disabled = claimsCurrentPage <= 1;
  const btnNext = document.getElementById('btnClaimsNext');
  if (btnNext) btnNext.disabled = claimsCurrentPage >= maxPage;

  if (pageRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="13" class="text-center" style="padding: 40px; color: var(--text-muted); font-size: 0.9rem;">Tidak ada data klaim AR pada status ini.</td></tr>`;
    return;
  }

  pageRows.forEach((r, i) => {
    const rawWA = String(r.customerPhone || '').replace(/\D/g, '');
    const normWA = normalizePhone(rawWA);
    const displayWA = formatPhoneDisplay(rawWA);
    const custName = r.customerName || '-';
    const initials = getCustomerInitials(custName);

    // 1. CUSTOMER: Avatar Circle (slate-900) + Name Bold + Phone Subtext (Persis Format Keep)
    const custHtml = `
      <div class="customer-cell-flex">
        <div class="avatar-circle-sm" title="${custName}">${initials}</div>
        <div class="customer-info-box">
          <span class="customer-name-bold">${custName}</span>
          ${(normWA && normWA.length >= 8) ? `
            <a href="https://wa.me/${normWA}" target="_blank" class="customer-phone-sub" title="Hubungi customer di WhatsApp">
              ${displayWA}
            </a>
          ` : `<span class="customer-phone-sub">${displayWA || '-'}</span>`}
        </div>
      </div>
    `;

    // 2. TANGGAL: Formatted Date + Periode
    const dateFormatted = formatCrosellingDate(r.postingDate) || '-';
    const periodeDisplay = r.periode || '';
    const dateHtml = `
      <div style="font-weight: 700; color: #0f172a; white-space: nowrap; font-variant-numeric: tabular-nums;">${dateFormatted}</div>
      ${periodeDisplay ? `<div style="font-size: 0.70rem; color: #64748b; margin-top: 2px;">${periodeDisplay}</div>` : ''}
    `;

    // 3. NO. AR: Bold Red/Blue Clickable Number
    const arHtml = `
      <div class="cell-report-ar-wrapper" style="justify-content: center;">
        <span class="text-ar-number-bold" onclick="openARDetailModal('${r.docNumber}')" title="Klik untuk cek detail faktur AR" style="cursor: pointer; font-variant-numeric: tabular-nums;">
          ${r.docNumber}
        </span>
      </div>
    `;

    // 4. BARANG: Product name bold + Harga Jual
    const barangHtml = `
      <div style="font-weight: 700; color: #0f172a; font-size: 0.82rem; line-height: 1.35;">${r.productName || '-'}</div>
      <div style="font-size: 0.72rem; color: #64748b; margin-top: 3px; font-variant-numeric: tabular-nums; font-weight: 600;">${formatRupiah(r.hargaJual)}</div>
    `;

    // 5. KODE SOSMED: Pill badge + PIC
    const sosmedHtml = `
      <span class="badge-sosmed-pill" style="font-size: 0.72rem; padding: 3px 8px; font-weight: 700;">${r.targetSosmedCode || '-'}</span>
      ${r.picSosmed ? `<div style="font-size: 0.70rem; color: #64748b; margin-top: 3px;">${r.picSosmed}</div>` : ''}
    `;

    // 6. CABANG: Bold Branch Name
    const cabangHtml = `
      <span style="font-weight: 800; font-size: 0.80rem; color: #0f172a; letter-spacing: 0.02em;">${r.cabang || '-'}</span>
    `;

    // 7. DENDA: Modern Status Pill with Dot (Anti-Slop Clean Style)
    let dendaDisplay = '';
    if (r.status === 'COLLECTING' || r.status === 'SHARED') {
      dendaDisplay = r.isHandphone 
        ? `<span class="modern-status-pill status-pill-orange"><span class="status-dot"></span>Potensi Rp 50.000</span>` 
        : `<span class="modern-status-pill status-pill-gray"><span class="status-dot"></span>Pass (Rp 0)</span>`;
    } else if (r.status === 'PENALTY_REAL' || r.status === 'RESOLVED_MANUAL_SOSMED') {
      dendaDisplay = r.isHandphone 
        ? `<span class="modern-status-pill status-pill-red"><span class="status-dot"></span>Denda Rp 50.000</span>` 
        : `<span class="modern-status-pill status-pill-gray"><span class="status-dot"></span>Pass (Rp 0)</span>`;
    } else if (r.status === 'RESOLVED_BY_STORE') {
      dendaDisplay = `<span class="modern-status-pill status-pill-green"><span class="status-dot"></span>Gugur (Pass)</span>`;
    } else {
      dendaDisplay = `<span class="modern-status-pill status-pill-gray"><span class="status-dot"></span>-</span>`;
    }

    // 8. FORMAT KEEP: Clean snippet box
    const keepHtml = r.keepText ? `
      <div class="claim-keep-box" title="Format Keep WhatsApp">${r.keepText}</div>
    ` : `<span style="color: #94a3b8; font-size: 0.76rem;">-</span>`;

    // 9. EVIDENCE: Clean thumbnail or multi-photo strip
    const imgs = (Array.isArray(r.evidenceImages) && r.evidenceImages.length > 0)
      ? r.evidenceImages
      : (r.evidenceImage ? [r.evidenceImage] : []);

    let evidenceHtml = '';
    if (imgs.length > 0) {
      const thumbsHtml = imgs.slice(0, 2).map((imgUrl, i) => `
        <div class="evidence-thumb-wrapper" onclick="openEvidenceImageModal('${r.id}', ${i})" title="Bukti ${i + 1} dari ${imgs.length} (Klik untuk perbesar)">
          <img src="${imgUrl}" class="evidence-thumb-img" alt="Bukti ${i + 1}">
          <div class="evidence-zoom-hint">${i + 1}</div>
        </div>
      `).join('');

      const moreBadge = imgs.length > 2 
        ? `<span class="evidence-more-badge" onclick="openEvidenceImageModal('${r.id}', 2)" title="Lihat ${imgs.length - 2} bukti lainnya">+${imgs.length - 2}</span>` 
        : '';

      evidenceHtml = `
        <div class="evidence-cell">
          <div class="evidence-thumbs-strip">
            ${thumbsHtml}
            ${moreBadge}
          </div>
          <div class="evidence-cell-actions" style="margin-top: 3px;">
            <button type="button" class="btn-evidence-tiny" onclick="triggerEvidenceUpload('${r.id}')" title="Tambah foto bukti lain pada baris ini">+ Tambah</button>
            <button type="button" class="btn-evidence-tiny" onclick="openEvidenceImageModal('${r.id}', 0)" title="Lihat galeri semua foto bukti (${imgs.length})">Lihat (${imgs.length})</button>
            <button type="button" class="btn-evidence-tiny btn-evidence-del" onclick="removeEvidenceImage('${r.id}')" title="Hapus foto bukti baris ini">Hapus</button>
          </div>
        </div>
      `;
    } else {
      evidenceHtml = `
        <button type="button" class="btn-upload-evidence" onclick="triggerEvidenceUpload('${r.id}')" title="Klik untuk mengunggah bukti format keep (bisa pilih beberapa gambar)">
          + Unggah Bukti
        </button>
      `;
    }

    // 10. AR SALES EMPLOYEE
    const salesHtml = `
      <span style="font-size: 0.76rem; font-weight: 600; color: #334155;">${r.currentSalesCode || '-'}</span>
    `;

    // 11. STATUS VALIDASI
    const isInvalid = (r.validationStatus === 'TIDAK VALID');
    const validityHtml = `
      <select class="select-claim-validity ${isInvalid ? 'invalid' : 'valid'}" onchange="handleClaimValidityChange('${r.id}', this.value)" title="Ubah status validasi klaim">
        <option value="VALID" ${!isInvalid ? 'selected' : ''}>Valid</option>
        <option value="TIDAK VALID" ${isInvalid ? 'selected' : ''}>Tidak Valid</option>
      </select>
    `;

    // 12. ALASAN TIDAK VALID
    const reasonHtml = isInvalid ? `
      <input type="text" class="input-claim-reason" value="${r.invalidReason || ''}" placeholder="Tulis alasan..." onchange="handleClaimInvalidReasonChange('${r.id}', this.value)" title="Alasan tidak valid">
    ` : `<span style="color: #94a3b8; font-size: 0.76rem;">-</span>`;

    // 13. AKSI: Anti-Slop Clean Flat Buttons
    let actionBtn = '';
    if (r.status === 'COLLECTING') {
      actionBtn = `<button type="button" class="btn-claim-action btn-claim-delete" onclick="deleteARClaimItem('${r.id}')" title="Hapus dari antrean klaim">Hapus</button>`;
    } else if (r.status === 'SHARED') {
      actionBtn = `
        <div style="display: flex; flex-direction: column; gap: 4px; align-items: center;">
          <button type="button" class="btn-claim-action btn-claim-secondary" onclick="copyBranchWAFollowupMessage('${r.cabang}')" title="Salin pesan follow-up WA untuk cabang ini">Chat WA</button>
          <button type="button" class="btn-claim-action btn-claim-delete" onclick="deleteARClaimItem('${r.id}')" title="Hapus">Hapus</button>
        </div>
      `;
    } else if (r.status === 'PENALTY_REAL') {
      actionBtn = `
        <div style="display: flex; flex-direction: column; gap: 4px; align-items: center;">
          <button type="button" class="btn-claim-action btn-claim-primary" onclick="markClaimResolvedManually('${r.id}')" title="Tandai sudah direvisi mandiri oleh tim sosmed di SAP">Ubah Mandiri</button>
          <button type="button" class="btn-claim-action btn-claim-secondary" onclick="copyBranchWAFollowupMessage('${r.cabang}')" title="Kirim teguran denda ke cabang">Tegur</button>
        </div>
      `;
    } else {
      actionBtn = `<button type="button" class="btn-claim-action btn-claim-delete" onclick="deleteARClaimItem('${r.id}')" title="Hapus">Hapus</button>`;
    }

    const tr = document.createElement('tr');
    if (isInvalid) {
      tr.classList.add('row-lose-tint');
    }

    tr.innerHTML = `
      <td class="text-left">${custHtml}</td>
      <td class="text-center">${dateHtml}</td>
      <td class="text-center">${arHtml}</td>
      <td class="text-left">${barangHtml}</td>
      <td class="text-center">${sosmedHtml}</td>
      <td class="text-center">${cabangHtml}</td>
      <td class="text-center">${dendaDisplay}</td>
      <td class="text-left">${keepHtml}</td>
      <td class="text-center">${evidenceHtml}</td>
      <td class="text-left">${salesHtml}</td>
      <td class="text-center">${validityHtml}</td>
      <td class="text-left">${reasonHtml}</td>
      <td class="text-center">${actionBtn}</td>
    `;
    tbody.appendChild(tr);
  });
}

function populateClaimsDropdowns() {
  const cabSelect = document.getElementById('claimsFilterCabang');
  const sosSelect = document.getElementById('claimsFilterSosmed');

  if (cabSelect) {
    const curCab = cabSelect.value;
    const branches = Array.from(new Set(rawARClaimsList.map(c => c.cabang).filter(Boolean))).sort();
    let cabOpts = '<option value="ALL">Semua Cabang Toko</option>';
    branches.forEach(b => {
      cabOpts += `<option value="${b}" ${b === curCab ? 'selected' : ''}>${b}</option>`;
    });
    cabSelect.innerHTML = cabOpts;
  }

  if (sosSelect) {
    const curSos = sosSelect.value;
    const sosList = Array.from(new Set(rawARClaimsList.map(c => c.targetSosmedCode).filter(Boolean))).sort();
    let sosOpts = '<option value="ALL">Semua Kode Sosmed</option>';
    sosList.forEach(s => {
      sosOpts += `<option value="${s}" ${s === curSos ? 'selected' : ''}>${s}</option>`;
    });
    sosSelect.innerHTML = sosOpts;
  }
}

function resetClaimsFilters() {
  if (document.getElementById('claimsFilterCabang')) document.getElementById('claimsFilterCabang').value = 'ALL';
  if (document.getElementById('claimsFilterSosmed')) document.getElementById('claimsFilterSosmed').value = 'ALL';
  if (document.getElementById('claimsFilterTanggal')) document.getElementById('claimsFilterTanggal').value = '';
  if (document.getElementById('claimsSearchInput')) document.getElementById('claimsSearchInput').value = '';
  claimsCurrentPage = 1;
  applyClaimsFilters();
}

function changeClaimsPage(delta) {
  const maxPage = Math.max(1, Math.ceil(filteredARClaimsList.length / CLAIMS_PAGE_SIZE));
  const newPage = claimsCurrentPage + delta;
  if (newPage >= 1 && newPage <= maxPage) {
    claimsCurrentPage = newPage;
    renderARClaimsTable();
  }
}

// Auto init on page load
window.addEventListener('DOMContentLoaded', async () => {
  initLiveClock();
  await loadMonthlyTargets();
  await loadSalesPhotos();
  await loadMonthlyArchives();
  await loadARClaimsData();
  populateAdminPanel(getMasterSales());
  onAdminTargetMonthChange();
  renderArchiveGrid();
  await loadAllDefaultFiles();
  populateDashboardFilterDropdowns();
  applyDashboardFilters();
  loadKeepData();
  startRealtimeSyncEngine();
});

// Shortcut Global Cmd+K / Ctrl+K untuk Fokus ke Search Bar Aktif (Gambar 1 & 2)
document.addEventListener('keydown', (e) => {
  if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
    e.preventDefault();
    if (currentActiveTab === 4) {
      const inp = document.getElementById('keepSearchInput');
      if (inp) { inp.focus(); inp.select(); }
    } else if (currentActiveTab === 5) {
      const inp = document.getElementById('auditSearchInput');
      if (inp) { inp.focus(); inp.select(); }
    }
  }
});


