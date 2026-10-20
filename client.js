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

function toggleWorkspaceMenu(e) {
  if (e) e.stopPropagation();
  const dropdown = document.getElementById('workspaceNavDropdown');
  if (dropdown) dropdown.classList.toggle('show');
}

function closeWorkspaceMenu() {
  const dropdown = document.getElementById('workspaceNavDropdown');
  if (dropdown) dropdown.classList.remove('show');
}

// Close workspace dropdown menu when clicking anywhere outside
document.addEventListener('click', function(e) {
  const wrapper = document.querySelector('.header-menu-wrapper');
  const dropdown = document.getElementById('workspaceNavDropdown');
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
  for (let i = 1; i <= 7; i++) {
    const btn = document.getElementById(`tabBtn${i}`);
    const content = document.getElementById(`tabContent${i}`);
    if (btn) btn.classList.toggle('active', i === tabIndex);
    if (content) content.classList.toggle('active', i === tabIndex);
  }

  closeWorkspaceMenu();

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
  if (tabIndex === 6) renderArchiveGrid();
  if (tabIndex === 7) {
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
  }, 5000);
}

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

async function loadKeepData() {
  const pill = document.getElementById('keepStatusPill');
  if (pill) pill.innerHTML = '<span>⏳ Membaca Data Keep Google Sheets...</span>';

  try {
    const res = await fetch(`keep-data.csv?t=${Date.now()}`);
    if (!res.ok) throw new Error("File keep-data.csv tidak ditemukan");
    const csvText = await res.text();
    const wb = XLSX.read(csvText, { type: 'string' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(ws);

    if (!jsonData || jsonData.length === 0) return;
    rawKeepRows = jsonData;
    console.log(`Keep data loaded: ${rawKeepRows.length} rows`);

    if (pill) {
      pill.innerHTML = `<span>📁 Sumber: Google Sheets (${rawKeepRows.length.toLocaleString('id-ID')} Baris Data)</span>`;
    }

    // Populate Sosmed filter dropdown
    populateKeepSosmedDropdown();
    // Apply default filters (Oktober-26)
    applyKeepFilters();

    // Trigger Audit if Croselling data is already loaded
    if (typeof rawCrosellingRows !== 'undefined' && rawCrosellingRows.length > 0 && typeof runARMissingSosmedAudit === 'function') {
      runARMissingSosmedAudit();
    }
  } catch (err) {
    console.warn("Gagal memuat keep-data.csv:", err);
    if (pill) pill.innerHTML = '<span style="color:#e11d48;">❌ Gagal memuat data keep</span>';
  }
}

function parseCSVLine(text) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (inQuotes && text[i+1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      result.push(cur.trim());
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur.trim());
  return result;
}
window.parseCSVLine = parseCSVLine;

function parseKeepCSV(csvText) {
  try {
    if (!csvText || typeof csvText !== 'string') return;
    let jsonData = [];
    if (typeof XLSX !== 'undefined' && XLSX.read) {
      try {
        const wb = XLSX.read(csvText, { type: 'string' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        jsonData = XLSX.utils.sheet_to_json(ws);
      } catch (xlsxErr) {
        console.warn("SheetJS parse warning, using fallback CSV parser:", xlsxErr);
      }
    }

    if (!jsonData || jsonData.length === 0) {
      const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
      if (lines.length > 1) {
        const headers = parseCSVLine(lines[0]);
        for (let i = 1; i < lines.length; i++) {
          const vals = parseCSVLine(lines[i]);
          const row = {};
          headers.forEach((h, idx) => {
            row[h] = vals[idx] !== undefined ? vals[idx] : '';
          });
          jsonData.push(row);
        }
      }
    }

    if (!jsonData || jsonData.length === 0) return;
    rawKeepRows = jsonData;
    console.log(`Keep CSV parsed successfully: ${rawKeepRows.length} rows`);
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

function renderKeepTable() {
  applyKeepFilters();
}
window.renderKeepTable = renderKeepTable;

async function syncLiveKeepData() {
  const btn = document.getElementById('btnSyncKeepData');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) btn.innerHTML = '⏳ Menghubungi Google Sheets...';

  try {
    let csvData = null;
    try {
      const res = await fetch('/api/sync-keep');
      const text = await res.text();
      if (text.trim().startsWith('{')) {
        const data = JSON.parse(text);
        if (data.success && data.csvText) {
          csvData = data.csvText;
        }
      }
    } catch (e) {
      console.warn("Backend sync-keep warning:", e);
    }

    // Direct browser fallback to Google Sheets CSV export
    if (!csvData) {
      console.log("Fetching Google Sheets directly from browser...");
      const directRes = await fetch('https://docs.google.com/spreadsheets/d/1V0pIa0NZCprJqRwnSnikFaLZ-dd50nY0uE3bjVYsUe0/export?format=csv&gid=1014933310');
      if (directRes.ok) {
        csvData = await directRes.text();
      }
    }

    if (csvData) {
      parseKeepCSV(csvData);
      try {
        localStorage.setItem('gadgetmart_sosmed_keep_cache', csvData);
      } catch (storageErr) {
        console.warn("Storage quota exceeded for localStorage keep cache:", storageErr);
      }
      applyKeepFilters();
      alert(`✅ Data Keep berhasil disinkronisasi langsung dari Google Sheets!\nTotal ${rawKeepRows.length.toLocaleString('id-ID')} baris data berhasil dimuat.`);
    } else {
      await loadKeepData();
      alert("✅ Sinkronisasi selesai, memuat data keep terbaru.");
    }
  } catch (err) {
    console.error("Sync error:", err);
    alert(`Sinkronisasi Google Sheets gagal: ${err.message}. Mencoba membaca data cache lokal.`);
    await loadKeepData();
  } finally {
    if (btn) btn.innerHTML = originalText;
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

  const periodVal = document.getElementById('keepFilterPeriode').value;
  const sosmedVal = document.getElementById('keepFilterSosmed').value;
  const statusVal = document.getElementById('keepFilterStatus').value;
  const keyword = (document.getElementById('keepSearchInput').value || '').trim().toLowerCase();

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

    // 4. Keyword search (Name, Phone, Item, AR, Notes)
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

  document.getElementById('keepTotalCount').innerText = total.toLocaleString('id-ID');
  document.getElementById('keepWinCount').innerText = winCount.toLocaleString('id-ID');
  document.getElementById('keepProgressCount').innerText = progressCount.toLocaleString('id-ID');
  document.getElementById('keepLoseCount').innerText = loseCount.toLocaleString('id-ID');
  document.getElementById('keepClosingRate').innerText = `${closingRate}%`;

  keepCurrentPage = 1;
  renderKeepPage();
}

function resetKeepFilters() {
  document.getElementById('keepFilterPeriode').value = 'Oktober-26';
  document.getElementById('keepFilterSosmed').value = 'ALL';
  document.getElementById('keepFilterStatus').value = 'ALL';
  document.getElementById('keepSearchInput').value = '';
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

function formatKeepTimestamp(ts) {
  if (!ts) return '-';
  if (typeof ts === 'number') {
    const d = XLSX.SSF.parse_date_code(ts);
    if (d) {
      const timeStr = (d.H || d.M) ? ` ${String(d.H).padStart(2,'0')}:${String(d.M).padStart(2,'0')}` : '';
      return `${String(d.d).padStart(2, '0')}/${String(d.m).padStart(2, '0')}/${d.y}${timeStr}`;
    }
  }
  if (typeof ts === 'string') {
    const parts = ts.split(/[\/\-\s]/);
    if (parts.length >= 3) return ts;
  }
  return String(ts);
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

  // Update pagination info
  document.getElementById('keepPaginationInfo').innerText = total > 0 
    ? `Menampilkan ${startIdx + 1} - ${endIdx} dari total ${total.toLocaleString('id-ID')} Data` 
    : 'Tidak ada data yang sesuai filter';
  document.getElementById('keepCurrentPageText').innerText = `Halaman ${keepCurrentPage} / ${maxPage}`;
  document.getElementById('btnKeepPrev').disabled = keepCurrentPage <= 1;
  document.getElementById('btnKeepNext').disabled = keepCurrentPage >= maxPage;

  if (pageRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="11" class="text-center" style="padding: 30px; color: var(--text-muted);">Tidak ada data keep untuk filter yang dipilih.</td></tr>`;
    return;
  }

  pageRows.forEach((r, i) => {
    const globalIdx = startIdx + i + 1;
    const rawWA = String(r['Nomor WA'] || '').replace(/\D/g, '');
    const normWA = normalizePhone(rawWA);
    const displayWA = formatPhoneDisplay(rawWA);

    // WA Button
    const waHtml = (normWA && normWA.length >= 8) ? `
      <div style="display:flex; align-items:center; gap:6px;">
        <span class="font-bold">${displayWA}</span>
        <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp" style="padding:2px 7px; font-size:0.7rem;" title="Chat customer ini di WhatsApp">
          💬 WA
        </a>
      </div>
    ` : (displayWA || '-');

    // AR Link (Buka Pop-up AR Modal langsung, tidak dialihkan ke Tab 3)
    const arDoc = r['REPORT AR'];
    const keepRowIdx = startIdx + i;
    const arHtml = arDoc ? `
      <span class="font-bold text-crimson cursor-pointer" onclick="openARDetailModalFromKeep(${keepRowIdx})" title="Klik untuk membuka pop-up AR & form lapor Akmal">
        📄 ${arDoc}
      </span>
    ` : '<span style="color:var(--text-dim);">-</span>';

    // Status Badge
    const st = String(r['Status'] || '').trim().toUpperCase();
    let statusClass = 'badge-ink';
    if (st === 'WIN') statusClass = 'badge-status-win';
    else if (st === 'ON PROGRESS HOT') statusClass = 'badge-status-hot';
    else if (st === 'ON PROGRESS MEDIUM') statusClass = 'badge-status-med';
    else if (st === 'ON PROGRESS LOW') statusClass = 'badge-status-low';
    else if (st === 'LOSE' || st === 'TAKE OUT') statusClass = 'badge-status-lose';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-sticky">${globalIdx}</td>
      <td class="text-left font-bold">${formatKeepTimestamp(r['Timestamp'])}</td>
      <td class="text-left font-bold">${r['Nama'] || '-'}</td>
      <td class="text-left">${waHtml}</td>
      <td class="text-left font-bold" style="color:var(--ink);">${r['Barang Keep'] || '-'}</td>
      <td class="text-left" style="color:var(--ink-secondary); font-size:0.75rem;">${r['Tanggal Ambil'] || '-'}</td>
      <td class="text-left" style="font-size:0.75rem; white-space:normal; line-height:1.3;">${r['Catatan'] || '-'}</td>
      <td class="text-left"><span class="kpi-badge badge-crimson font-bold">${r['Kode Sosmed'] || '-'}</span></td>
      <td class="text-center">${arHtml}</td>
      <td class="text-center"><span class="kpi-badge ${statusClass} font-bold">${r['Status'] || '-'}</span></td>
      <td class="text-left" style="font-size:0.75rem; white-space:normal; line-height:1.3;">${r['Feedback Tim FU'] || '-'}</td>
    `;
    tbody.appendChild(tr);
  });
}

function openARDetailModalFromKeep(keepIndex) {
  const row = filteredKeepRows[keepIndex];
  if (!row) return;
  const arDoc = row['REPORT AR'];
  if (!arDoc) {
    alert("Baris data keep ini tidak memiliki nomor REPORT AR.");
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

  if (elCust) elCust.innerText = custName;
  if (elWA) elWA.innerHTML = waHtml;
  if (elSos) elSos.innerHTML = `<span class="badge-sosmed-pill">${sosmedCode}</span>`;
  if (elBarang) elBarang.innerText = keepBarang;
  if (elCat) elCat.innerHTML = catatanModalHtml;

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

    const isSosmed = isItemInputSosmed(r);
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
    subtitle.innerText = `Seluruh ${matchedRows.length} item dalam AR ini tercatat dengan kode sales toko fisik/kasir, belum ada kode sosmed.`;
  } else if (missingCount > 0) {
    banner.className = 'ar-status-banner banner-warning';
    icon.innerHTML = '⚠️';
    title.innerText = `SEBAGIAN ITEM BELUM TERINPUT (${missingCount} DARI ${matchedRows.length} ITEM)`;
    subtitle.innerText = `Sebagian produk sudah berstatus sosmed, namun masih ada ${missingCount} item yang belum terinput sosmed.`;
  } else {
    banner.className = 'ar-status-banner banner-success';
    icon.innerHTML = '✅';
    title.innerText = 'SEMUA ITEM SUDAH TERINPUT SOSMED';
    subtitle.innerText = `Semua ${matchedRows.length} item dalam dokumen ini sudah valid terinput kode sales sosmed.`;
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

    const isSosmed = isItemInputSosmed(r);
    const schl = String(r['SC/HL Code'] || '').trim();
    const slp = String(r['SlpName'] || '').trim();
    const emp = String(r['Sales Employee Name'] || '').trim();
    const empDisplay = schl || slp || emp || '-';

    const isChecked = currentModalCheckedItems.has(idx);

    // Potensi Denda: Handphone = DENDA, Aksesoris/lainnya = PASS
    const dendaStatus = getItemPotensiDenda(r);
    const dendaBadge = dendaStatus === 'DENDA'
      ? '<span class="kpi-badge badge-crimson font-bold">🚨 DENDA</span>'
      : '<span class="kpi-badge badge-ink">PASS</span>';

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
      <td class="text-center">
        ${isSosmed
          ? '<span class="status-badge badge-accent">✅ Sudah Sosmed</span>'
          : '<span class="status-badge badge-crimson font-bold">❌ Belum Sosmed</span>'}
      </td>
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

  rawKeepRows.forEach(k => {
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
      let sosmedItems = 0;
      let missingItems = 0;
      let missingTotalValue = 0;
      let totalDocValue = 0;

      items.forEach(it => {
        const lineTotal = Number(it['Harga Jual']) || 0;
        totalDocValue += lineTotal;

        if (isItemInputSosmed(it)) {
          sosmedItems++;
        } else {
          missingItems++;
          missingTotalValue += lineTotal;
        }
      });

      if (missingItems > 0) {
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
          status: sosmedItems === 0 ? 'ALL_MISSING' : 'PARTIAL_MISSING',
          keep: bestKeep,
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
    }
  });

  // Sort descending by Timestamp Format Keep!
  rawAuditARList.sort((a, b) => (b.keepTimestampEpoch || 0) - (a.keepTimestampEpoch || 0));

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
      if (r.status !== statusVal) return false;
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

  // Calculate KPI counters
  let allMissingCount = 0;
  let partialMissingCount = 0;
  let totalPotensiRp = 0;

  filteredAuditARList.forEach(r => {
    if (r.status === 'ALL_MISSING') allMissingCount++;
    else if (r.status === 'PARTIAL_MISSING') partialMissingCount++;
    totalPotensiRp += (r.missingTotalValue || 0);
  });

  const totalEl = document.getElementById('auditTotalCount');
  const allEl = document.getElementById('auditAllMissingCount');
  const partialEl = document.getElementById('auditPartialMissingCount');
  const potensiEl = document.getElementById('auditTotalPotensiRp');

  if (totalEl) totalEl.innerText = filteredAuditARList.length.toLocaleString('id-ID');
  if (allEl) allEl.innerText = allMissingCount.toLocaleString('id-ID');
  if (partialEl) partialEl.innerText = partialMissingCount.toLocaleString('id-ID');
  if (potensiEl) potensiEl.innerText = formatRupiah(totalPotensiRp);

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
    tbody.innerHTML = `<tr><td colspan="11" class="text-center" style="padding: 30px; color: var(--text-muted);">Tidak ada data tracking AR yang sesuai filter.</td></tr>`;
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

    // Notes HTML (Keep Catatan and/or Cashier Remarks)
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

    // Report AR cell
    const reportARDocDisplay = (r.reportARDoc && r.reportARDoc !== '-')
      ? `<span class="cursor-pointer font-bold text-crimson" onclick="openARDetailModal('${r.reportARDoc}', filteredAuditARList[${auditRowIdx}].keep, 'REPORT AR')" title="Klik untuk cek detail Dokumen Report AR">📄 ${r.reportARDoc}</span>`
      : '-';

    const reportARDateDisplay = r.reportARDate
      ? `${formatCrosellingDate(r.reportARDate)} ${r.reportARCabang || ''}`.trim()
      : '-';

    // Vlookup AR cell
    const vlookupARDocDisplay = (r.vlookupARDoc && r.vlookupARDoc !== '-')
      ? `<span class="cursor-pointer font-bold text-crimson" onclick="openARDetailModal('${r.vlookupARDoc}', filteredAuditARList[${auditRowIdx}].keep, 'AR VLOOKUP')" title="Klik untuk cek detail Dokumen AR VLOOKUP">📄 ${r.vlookupARDoc}</span>`
      : '-';

    const vlookupARDateDisplay = r.vlookupARDate
      ? `${formatCrosellingDate(r.vlookupARDate)} ${r.vlookupARCabang || ''}`.trim()
      : '-';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="audit-col-no" style="text-align: center;">${globalIdx}</td>
      <td class="audit-col-keep-date" style="text-align: center; font-weight: 600;">${formatKeepTimestamp(keepObj['Timestamp'])}</td>
      <td class="audit-col-cust font-bold" style="color: #0f172a;">${keepObj['Nama'] || '-'}</td>
      <td class="audit-col-wa">${waHtml}</td>
      <td class="audit-col-sosmed" style="text-align: center;">
        <span class="badge-sosmed-pill">${keepObj['Kode Sosmed'] || '-'}</span>
      </td>
      <td class="audit-col-keep font-bold" style="color: #1e293b;">${keepObj['Barang Keep'] || '-'}</td>
      <td class="audit-col-notes">${catatanHtml}</td>
      <td class="audit-col-report-ar" style="text-align: center;">${reportARDocDisplay}</td>
      <td class="audit-col-report-date" style="text-align: center; font-size: 0.73rem;">${reportARDateDisplay}</td>
      <td class="audit-col-vlookup-ar" style="text-align: center;">${vlookupARDocDisplay}</td>
      <td class="audit-col-vlookup-date" style="text-align: center; font-size: 0.73rem;">${vlookupARDateDisplay}</td>
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
  } catch (err) {
    // Silent fail if network briefly disconnected
  }
}

let syncEngineInterval = null;
function startRealtimeSyncEngine() {
  if (syncEngineInterval) clearInterval(syncEngineInterval);
  // Polling status server setiap 5 detik
  syncEngineInterval = setInterval(pollServerSyncStatus, 5000);
  console.log('📡 Realtime Cross-Device Sync Engine aktif (interval 5s)');
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

// ==========================================================================
// 17. TANYA YIPYEP FLOATING WIDGET ENGINE (CHAT AI ASISTEN KUCING ROBOT)
// ==========================================================================
function toggleYipyepChat() {
  const win = document.getElementById('yipyepChatWindow');
  if (!win) return;
  const isHidden = win.style.display === 'none' || !win.style.display;
  win.style.display = isHidden ? 'flex' : 'none';
  if (isHidden) {
    const input = document.getElementById('yipyepChatInput');
    if (input) input.focus();
  }
}

function quickAskYipyep(question) {
  const input = document.getElementById('yipyepChatInput');
  if (input) {
    input.value = question;
    sendYipyepChat();
  }
}

function formatMarkdownForChat(mdText) {
  if (!mdText) return '';
  return mdText
    .replace(/\r\n/g, '\n')
    .replace(/^### (.*$)/gim, '<strong style="display:block; margin:6px 0 2px 0;">$1</strong>')
    .replace(/^## (.*$)/gim, '<strong style="display:block; margin:8px 0 3px 0; font-size:1.05em;">$1</strong>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>')
    .replace(/^\* (.*$)/gim, '• $1')
    .replace(/^- (.*$)/gim, '• $1')
    .replace(/\n\n/g, '<br><br>')
    .replace(/\n/g, '<br>');
}

function getYipyepAIContext() {
  try {
    const monthNames = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];

    // 1. Database Arsip Bulanan Konsolidasi (Januari - Desember 2026)
    const arsipBulanKeBulan = monthNames.map((mName, i) => {
      const monthNum = i + 1;
      const arc = (typeof cachedMonthlyArchives !== 'undefined' && Array.isArray(cachedMonthlyArchives))
        ? cachedMonthlyArchives.find(a => Number(a.month) === monthNum && (Number(a.year) === activeYear || !a.year))
        : null;

      if (arc) {
        return {
          bulan: `${mName} ${activeYear}`,
          status: 'Tersimpan di Arsip',
          targetOmzet: arc.targetOmzet || 8447356034,
          realisasiOmzet: arc.totalAchieve || (arc.totalAchieveRaw ? formatRupiah(arc.totalAchieveRaw) : '-'),
          persenOmzet: arc.pctOmzet ? `${arc.pctOmzet}%` : (arc.progress || '-'),
          targetGP: arc.targetGP || 422367802,
          realisasiGP: arc.totalGP || (arc.totalGPRaw ? formatRupiah(arc.totalGPRaw) : '-'),
          persenGP: arc.pctGP ? `${arc.pctGP}%` : '-',
          crossSellingRatio: arc.crossRatio || '-'
        };
      } else if (monthNum === activeMonth) {
        const totAch = document.getElementById('kpiTotalAchieve') ? document.getElementById('kpiTotalAchieve').innerText : 'Rp 3.279.978.057';
        const totGP = document.getElementById('kpiTotalGP') ? document.getElementById('kpiTotalGP').innerText : 'Rp 71.955.484';
        const prog = document.getElementById('kpiProgressPct') ? document.getElementById('kpiProgressPct').innerText : '39%';
        return {
          bulan: `${mName} ${activeYear} (Bulan Berjalan / Live)`,
          status: 'Sedang Berjalan (Live)',
          targetOmzet: 8447356034,
          realisasiOmzet: totAch,
          persenOmzet: prog,
          targetGP: 422367802,
          realisasiGP: totGP,
          timegone: typeof timegonePct !== 'undefined' ? `${timegonePct}%` : '42%'
        };
      } else {
        return {
          bulan: `${mName} ${activeYear}`,
          status: 'Mendatang (Belum Ada Transaksi)',
          targetOmzet: 8447356034,
          realisasiOmzet: '-'
        };
      }
    });

    // 2. Data Penjualan Cabang (Live dari turnover bulan aktif)
    const cabangMap = {};
    if (typeof rawTurnoverRows !== 'undefined' && Array.isArray(rawTurnoverRows) && rawTurnoverRows.length > 0) {
      rawTurnoverRows.forEach(r => {
        const cab = String(r['Cabang'] || r['SalesCabang'] || 'LAINNYA').trim().toUpperCase();
        const dpp = Number(r['DPP']) || 0;
        const total = Number(r['Total']) || 0;
        const amt = (typeof currentCalculationMode !== 'undefined' && currentCalculationMode === 'DPP') ? dpp : total;
        cabangMap[cab] = (cabangMap[cab] || 0) + amt;
      });
    }
    const cabangList = Object.keys(cabangMap)
      .sort((a,b) => cabangMap[b] - cabangMap[a])
      .map(c => ({
        kodeCabang: c,
        namaCabang: (typeof CABANG_NAMES !== 'undefined' && CABANG_NAMES[c]) ? CABANG_NAMES[c] : c,
        totalOmzet: formatRupiah(cabangMap[c]),
        nominalRaw: cabangMap[c]
      }));

    // 3. Leaderboard Sales Teratas
    let salesLeaderboard = [];
    if (typeof lastActivePICs !== 'undefined' && Array.isArray(lastActivePICs) && typeof lastProcessedPicStats !== 'undefined' && lastProcessedPicStats) {
      salesLeaderboard = lastActivePICs.map((pic, idx) => {
        const st = lastProcessedPicStats[pic] || {};
        return {
          ranking: idx + 1,
          namaSales: pic,
          omzet: formatRupiah(st.pencapaian || 0),
          persenTarget: `${Math.round(st.pctAchieve || 0)}%`,
          grossProfit: formatRupiah(st.acvGP || 0)
        };
      });
    }

    // 4. Data Keep Customer
    const keepSummary = {
      totalKeep: (typeof filteredKeepRows !== 'undefined' && filteredKeepRows) ? filteredKeepRows.length : 0,
      closingRate: document.getElementById('keepClosingRate') ? document.getElementById('keepClosingRate').innerText : '0%',
      win: document.getElementById('keepWinCount') ? document.getElementById('keepWinCount').innerText : '0',
      progress: document.getElementById('keepProgressCount') ? document.getElementById('keepProgressCount').innerText : '0',
      lose: document.getElementById('keepLoseCount') ? document.getElementById('keepLoseCount').innerText : '0'
    };

    return {
      arsipBulananLengkap: arsipBulanKeBulan,
      penjualanCabangOktober: cabangList,
      leaderboardSales: salesLeaderboard,
      ringkasanKeep: keepSummary
    };
  } catch (err) {
    console.warn("getYipyepAIContext error fallback:", err);
    return {
      arsipBulananLengkap: [],
      penjualanCabangOktober: [],
      leaderboardSales: [],
      ringkasanKeep: {}
    };
  }
}

async function sendYipyepChat() {
  const input = document.getElementById('yipyepChatInput');
  const container = document.getElementById('yipyepMessagesBody');
  if (!input || !container) return;

  const prompt = input.value.trim();
  if (!prompt) return;

  const timeStr = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
  const userHtml = `
    <div class="yipyep-msg-bubble yipyep-user">
      <div class="yipyep-bubble-text">${prompt}</div>
      <div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; text-align:right; margin-top:4px;">${timeStr}</div>
    </div>
  `;
  container.insertAdjacentHTML('beforeend', userHtml);
  input.value = '';
  container.scrollTop = container.scrollHeight;

  const typingId = `typing_${Date.now()}`;
  container.insertAdjacentHTML('beforeend', `
    <div class="yipyep-msg-bubble yipyep-bot" id="${typingId}">
      <div class="yipyep-bubble-text"><span class="yipyep-typing-dots"><span>.</span><span>.</span><span>.</span> Yipyep sedang menghitung data</span></div>
    </div>
  `);
  container.scrollTop = container.scrollHeight;

  try {
    // Shortcuts: WA number, AR number, Keep
    const waMatch = prompt.match(/(?:wa|whatsapp|nomor|no\s*hp|telepon|hp)?\s*(\+?62\d{8,13}|08\d{8,12})/i);
    const arMatch = prompt.match(/\b(26\d{7})\b/);
    const isKeep = prompt.toLowerCase().includes('keep');

    if (waMatch) {
      const phone = waMatch[1];
      if (typeof quickCheckAR === 'function') quickCheckAR(phone);
      const typingEl = document.getElementById(typingId);
      if (typingEl) {
        typingEl.innerHTML = `<div class="yipyep-bubble-text">📱 <strong>Hasil Cek WA (${phone}):</strong> Aku sudah alihkan ke <strong>Tab 3 (Cek AR & WhatsApp)</strong> untuk melihat data pelanggan dan transaksi!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
        container.scrollTop = container.scrollHeight;
      }
      return;
    }

    if (arMatch) {
      const docNo = arMatch[1];
      if (typeof quickCheckAR === 'function') quickCheckAR(docNo);
      const typingEl = document.getElementById(typingId);
      if (typingEl) {
        typingEl.innerHTML = `<div class="yipyep-bubble-text">🔍 <strong>Hasil Cek AR ${docNo}:</strong> Aku sudah buka rincian faktur di <strong>Tab 3</strong> untuk memverifikasi item sosmed!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
        container.scrollTop = container.scrollHeight;
      }
      return;
    }

    if (isKeep) {
      if (typeof switchTab === 'function') switchTab(4);
      const typingEl = document.getElementById(typingId);
      if (typingEl) {
        typingEl.innerHTML = `<div class="yipyep-bubble-text">📋 <strong>Data Keep:</strong> Aku sudah membawamu ke <strong>Tab 4 (Data Keep Customer)</strong> untuk mengecek status WIN/Progress/Lose!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
        container.scrollTop = container.scrollHeight;
      }
      return;
    }

    // Context & System Instruction
    const dbContext = getYipyepAIContext();
    const apiKey = localStorage.getItem('gemini_api_key') || DEFAULT_GEMINI_API_KEY;
    let userModel = localStorage.getItem('gemini_model') || DEFAULT_GEMINI_MODEL;
    if (userModel === 'gemini-flash-latest' || userModel === 'gemini-2.5-flash') {
      userModel = 'gemini-3.5-flash-lite';
    }

    const systemInstruction = `Kamu adalah Yipyep 🐱✨, asisten AI analitik robot kucing pintar dan ramah untuk Divisi Sosmed GadgetMart.
Kamu memiliki akses langsung ke DATABASE ARSIP BULANAN (Januari - Desember 2026), DATA PENJUALAN PER CABANG TOKO, LEADERBOARD SALES, dan DATA KEEP CUSTOMER.

DATA SISTEM REAL-TIME:
${JSON.stringify(dbContext)}

PANDUAN MENJAWAB:
1. Jawab pertanyaan user dengan ramah, cerdas, akurat, dan percaya diri.
2. Jika user meminta rincian atau memecah omzet/pencapaian per bulan (misal dari Januari sampai Desember), tampilkan rincian setiap bulan secara jelas dan runtut berdasarkan data 'arsipBulananLengkap'.
3. Jika user menanyakan cabang tertentu (misal BJMC, BTCC, dll), berikan informasi data cabang yang ada (termasuk kontribusi cabang di bulan berjalan Oktober atau perbandingan dengan total).
4. Gunakan pemformatan Markdown yang bersih: bullet points (*), teks tebal (**teks**), atau tabel mini bila diperlukan agar sangat mudah dibaca.
5. Bila user menyapa atau bertanya umum, jawab dengan gaya kucing robot cerdas yang menyenangkan.`;

    let aiSuccess = false;

    // STEP 1: Try Backend Serverless Route /api/chat-ai (Ultra-fast, zero CORS, Node.js connection)
    try {
      const serverRes = await fetch('/api/chat-ai', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt: prompt,
          systemInstruction: systemInstruction,
          apiKey: apiKey,
          model: userModel
        })
      });
      if (serverRes.ok) {
        const data = await serverRes.json();
        if (data && data.success && data.text) {
          const formatted = formatMarkdownForChat(data.text);
          const typingEl = document.getElementById(typingId);
          if (typingEl) {
            typingEl.innerHTML = `<div class="yipyep-bubble-text">🐱 <strong>Yipyep:</strong><br>${formatted}</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
            container.scrollTop = container.scrollHeight;
          }
          aiSuccess = true;
          return;
        }
      }
    } catch (serverErr) {
      console.warn("Backend /api/chat-ai call failed, trying direct browser fetch:", serverErr);
    }

    // STEP 2: Direct browser fetch fallback (if backend unavailable)
    if (!aiSuccess && apiKey) {
      const modelsToTry = [userModel, 'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-3.8-flash'];
      const uniqueModels = Array.from(new Set(modelsToTry.filter(Boolean)));

      for (const m of uniqueModels) {
        try {
          const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{ parts: [{ text: `${systemInstruction}\n\nPertanyaan User: "${prompt}"` }] }]
            })
          });
          const data = await res.json();
          if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) {
            const rawText = data.candidates[0].content.parts.map(p => p.text || '').join('');
            const formatted = formatMarkdownForChat(rawText);
            const typingEl = document.getElementById(typingId);
            if (typingEl) {
              typingEl.innerHTML = `<div class="yipyep-bubble-text">🐱 <strong>Yipyep:</strong><br>${formatted}</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
              container.scrollTop = container.scrollHeight;
            }
            aiSuccess = true;
            return;
          }
        } catch(e) {
          console.warn(`Direct Gemini call with ${m} failed:`, e);
        }
      }
    }

    // STEP 3: Fallback Local Database Calculation Engine
    if (!aiSuccess) {
      const q = prompt.toLowerCase();
      let reply = '';

      const isMonthRangeQuery = q.includes('januari') || q.includes('desember') || q.includes('bulan') || q.includes('arsip') || q.includes('pecah');
      const matchedCabang = Object.keys(CABANG_NAMES).find(c => q.includes(c.toLowerCase()));

      if (isMonthRangeQuery && matchedCabang) {
        const cabInfo = (dbContext.penjualanCabangOktober || []).find(c => c.kodeCabang === matchedCabang);
        const cabName = CABANG_NAMES[matchedCabang] || matchedCabang;
        const cabOktoberVal = cabInfo ? cabInfo.totalOmzet : 'Rp 0';

        const lines = (dbContext.arsipBulananLengkap || []).map(m => {
          return `• <strong>${m.bulan}:</strong> ${m.realisasiOmzet} <em>(${m.status})</em>`;
        }).join('<br>');

        reply = `Meong! 🐾 Berikut adalah rincian omzet Divisi Sosmed dari <strong>Database Arsip Bulanan (Januari - Desember 2026)</strong>:<br><br>${lines}<br><br>🏢 <strong>Khusus Cabang ${cabName}:</strong> Pada transaksi bulan berjalan (Oktober 2026), cabang ini mencatat omzet sebesar <strong>${cabOktoberVal}</strong>!`;
      } else if (isMonthRangeQuery) {
        const lines = (dbContext.arsipBulananLengkap || []).map(m => {
          return `• <strong>${m.bulan}:</strong> ${m.realisasiOmzet} (${m.persenOmzet || m.status})`;
        }).join('<br>');
        reply = `Meong! 🐾 Ini data rekapitulasi omzet dari <strong>Database Arsip Bulanan 2026</strong>:<br><br>${lines}`;
      } else if (matchedCabang) {
        const cabInfo = (dbContext.penjualanCabangOktober || []).find(c => c.kodeCabang === matchedCabang);
        const cabName = CABANG_NAMES[matchedCabang] || matchedCabang;
        const val = cabInfo ? cabInfo.totalOmzet : 'Rp 0';
        reply = `🏢 <strong>Cabang ${cabName}:</strong> Total omzet sosmed yang tercatat di bulan berjalan mencapai <strong>${val}</strong>!`;
      } else if (q.includes('omzet') || q.includes('revenue') || q.includes('capaian')) {
        const totalAchieve = document.getElementById('kpiTotalAchieve') ? document.getElementById('kpiTotalAchieve').innerText : 'Rp 3.279.978.057';
        const progressPct = document.getElementById('kpiProgressPct') ? document.getElementById('kpiProgressPct').innerText : '39%';
        reply = `Meong! 🐾 Total pencapaian omzet tim sosmed bulan ini adalah <strong>${totalAchieve}</strong> (${progressPct} dari target). Status capaian berada di posisi <strong>🖤 On Track</strong>!`;
      } else if (q.includes('gp') || q.includes('gross') || q.includes('profit')) {
        const totalGP = document.getElementById('kpiTotalGP') ? document.getElementById('kpiTotalGP').innerText : 'Rp 71.955.484';
        reply = `Gross Profit tim yang telah dikumpulkan mencapai <strong>${totalGP}</strong>. Ingat target GP kita adalah 5%, pantau terus margin produk ya! 📈`;
      } else if (q.includes('cabang')) {
        const top3 = (dbContext.penjualanCabangOktober || []).slice(0, 3).map((c, i) => `<strong>${i+1}. ${c.namaCabang}</strong> (${c.totalOmzet})`).join(', ');
        reply = `🏢 <strong>Top Cabang Sosmed:</strong> ${top3}! 🚀`;
      } else {
        const totalAchieve = document.getElementById('kpiTotalAchieve') ? document.getElementById('kpiTotalAchieve').innerText : 'Rp 3.279.978.057';
        reply = `Halo! Aku <strong>Yipyep</strong> 🐱✨. Database arsip bulanan (Januari - Desember) dan data transaksi live sudah terhubung penuh. Tanyakan padaku rincian omzet, data cabang toko, atau performa sales!`;
      }

      const typingEl = document.getElementById(typingId);
      if (typingEl) {
        typingEl.innerHTML = `<div class="yipyep-bubble-text">🐱 <strong>Yipyep:</strong><br>${reply}</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
        container.scrollTop = container.scrollHeight;
      }
    }
  } catch (globalErr) {
    console.error("Critical Yipyep chat error:", globalErr);
    const typingEl = document.getElementById(typingId);
    if (typingEl) {
      typingEl.innerHTML = `<div class="yipyep-bubble-text">🐱 <strong>Yipyep:</strong><br>Meong! 🐾 Aku sudah membaca database sistem. Saat ini total pencapaian omzet tim sosmed mencapai <strong>${document.getElementById('kpiTotalAchieve') ? document.getElementById('kpiTotalAchieve').innerText : 'Rp 3.279.978.057'}</strong>. Silakan tanyakan data spesifik seputar cabang, arsip bulanan, atau faktur AR!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
      container.scrollTop = container.scrollHeight;
    }
  }
}

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

// Auto init on page load
window.addEventListener('DOMContentLoaded', async () => {
  initLiveClock();
  await loadMonthlyTargets();
  await loadSalesPhotos();
  await loadMonthlyArchives();
  populateAdminPanel(getMasterSales());
  onAdminTargetMonthChange();
  renderArchiveGrid();
  await loadAllDefaultFiles();
  populateDashboardFilterDropdowns();
  applyDashboardFilters();
  loadKeepData();
  startRealtimeSyncEngine();
});


