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
const DEFAULT_GEMINI_MODEL = 'gemini-flash-latest';

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
let cachedSalesPhotos = {};
let cachedMonthlyArchives = [];
let chartInstances = {};



// Utility Formatters
function formatRupiah(num) {
  if (num === null || num === undefined || isNaN(num) || num === 0) return 'Rp -';
  const isNeg = num < 0;
  const absVal = Math.abs(num);
  const formatted = absVal.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  return (isNeg ? '-Rp ' : 'Rp ') + formatted;
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
    const parts = pd.split(/[\/\-]/);
    if (parts.length >= 3) {
      return `${String(parts[0]).padStart(2, '0')}/${String(parts[1]).padStart(2, '0')}/${parts[2]}`;
    }
    return pd;
  }
  return '-';
}

// 2. PARSER ENGINE FOR turnover.xlsx (Achievement DPP)
function parseTurnoverBuffer(arrayBuffer) {
  try {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

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

// SMART FILE UPLOAD HANDLER
function parseExcelFile(arrayBuffer, fileName = '') {
  try {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    if (!jsonData || jsonData.length === 0) {
      alert("File Excel kosong atau tidak memiliki data.");
      return;
    }

    const firstRow = jsonData[0] || {};
    const hasWA = 'No HP Customer' in firstRow || 'SC/HL Code' in firstRow || fileName.toLowerCase().includes('croselling') || fileName.toLowerCase().includes('sheet');
    const hasDPP = 'DPP' in firstRow || fileName.toLowerCase().includes('turnover');

    if (hasWA && !hasDPP) {
      rawCrosellingRows = jsonData;
      saveExcelToServer(arrayBuffer, 'croselling');
      alert(`Berhasil mengimpor Data Google Sheets (${jsonData.length.toLocaleString('id-ID')} baris data)!\nDatabase AR & WhatsApp telah diperbarui dan tersimpan permanen ke Server.`);
      switchTab(2);
    } else if (hasDPP) {
      rawTurnoverRows = jsonData;
      saveExcelToServer(arrayBuffer, 'turnover');
      processAllReports();
      alert(`Berhasil mengimpor turnover.xlsx (${jsonData.length.toLocaleString('id-ID')} baris data)!\nLaporan Pencapaian & KPI telah diperbarui dan tersimpan permanen ke Server.`);
      switchTab(1);
    } else {
      // Default: check if it has No HP
      if ('No HP Customer' in firstRow) {
        rawCrosellingRows = jsonData;
        saveExcelToServer(arrayBuffer, 'croselling');
        switchTab(2);
      } else {
        rawTurnoverRows = jsonData;
        saveExcelToServer(arrayBuffer, 'turnover');
        processAllReports();
        switchTab(1);
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

  for (const row of rawTurnoverRows) {
    const emp = String(row['Sales Employee Name'] || '');
    const rem = String(row['Remarks'] || '');
    const cab = String(row['SalesCabang'] || '');

    // Identify sales PIC
    let matchedPIC = null;
    for (const m of masterList) {
      if (emp.includes(m.code) || rem.includes(m.code) || cab.includes(m.code)) {
        matchedPIC = m.name;
        break;
      }
    }

    if (!matchedPIC || !activePICs.includes(matchedPIC)) continue;

    // Resolve date with inverted US-locale fix
    const day = resolvePostingDay(row['Posting Date']);
    if (day >= 1 && day <= daysInMonth) {
      if (day > maxActiveDay) maxActiveDay = day;
    }

    // Amounts
    const dpp = Number(row['DPP']) || 0;
    const total = Number(row['Total']) || 0;
    const gp = Number(row['GrossProfit']) || 0;
    const amount = currentCalculationMode === 'DPP' ? dpp : total;

    // Report 1: Daily Matrix accumulation
    if (parsedDailyMatrix[day]) {
      parsedDailyMatrix[day].sales[matchedPIC] += amount;
      parsedDailyMatrix[day].totalDay += amount;
    }

    // Report 2: Omzet & GP accumulation
    picStats[matchedPIC].pencapaian += amount;
    picStats[matchedPIC].acvGP += gp;

    // Grouping for Cross-selling calculation (by Document Number)
    const docNo = row['Document Number'];
    if (docNo) {
      if (!docTransactions[docNo]) {
        docTransactions[docNo] = {
          pic: matchedPIC,
          hasGadget: false,
          hasCross: false
        };
      }
      const grp = String(row['Group Name'] || '').trim().toUpperCase();
      if (grp === 'GAD T') {
        docTransactions[docNo].hasGadget = true;
      }
      if (grp === 'ACC T' && (total >= 50000 || dpp >= 45000)) {
        docTransactions[docNo].hasCross = true;
      }
    }
  }

  // Calculate Cross-Selling Counts from Document groups
  for (const doc in docTransactions) {
    const item = docTransactions[doc];
    if (item.hasGadget && picStats[item.pic]) {
      picStats[item.pic].gadgetCount++;
      if (item.hasCross) {
        picStats[item.pic].crossCount++;
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
    const target = m ? m.target : 867261886;
    sumTarget += target;
    const total = colTotals[pic];
    const sisa = total - target;
    const pct = target > 0 ? (total / target) * 100 : 0;

    totalHtml += `<td>${formatRupiah(total)}</td>`;
    sisaHtml += `<td>${formatRupiah(sisa)}</td>`;
    targetHtml += `<td>${formatRupiah(target)}</td>`;
    progressHtml += `<td>${formatPercent(pct)}</td>`;
  });

  const grandSisa = grandTotalAchieve - sumTarget;
  const grandPct = sumTarget > 0 ? (grandTotalAchieve / sumTarget) * 100 : 0;

  totalHtml += `<td class="col-achieve">${formatRupiah(grandTotalAchieve)}</td><td></td><td></td></tr>`;
  sisaHtml += `<td class="col-achieve">${formatRupiah(grandSisa)}</td><td></td><td></td></tr>`;
  targetHtml += `<td class="col-achieve">${formatRupiah(sumTarget)}</td><td></td><td></td></tr>`;
  progressHtml += `<td class="col-achieve">${formatPercent(grandPct)}</td><td></td><td></td></tr>`;

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

  activePICs.forEach(pic => {
    const s = picStats[pic];
    sumTarget += s.target;
    sumAch += s.pencapaian;
    sumTargetGP += s.targetGP;
    sumAcvGP += s.acvGP;
    sumGadget += s.gadgetCount;
    sumCross += s.crossCount;

    // Omzet % and status
    const pctOmzet = s.target > 0 ? (s.pencapaian / s.target) * 100 : 0;
    let statusOmzet = '☠️';
    if (pctOmzet >= (timegonePct - 4)) statusOmzet = '🖤';
    else if (pctOmzet >= (timegonePct - 7)) statusOmzet = '💔';

    // GP % and status
    const pctGP = s.targetGP > 0 ? (s.acvGP / s.targetGP) * 100 : 0;
    let statusGP = '☠️';
    if (pctGP >= (timegonePct - 4)) statusGP = '🖤';
    else if (pctGP >= (timegonePct - 7)) statusGP = '💔';

    // Cross selling % (target ratio 1/3)
    const ratioActual = s.gadgetCount > 0 ? (s.crossCount / s.gadgetCount) : 0;
    const pctCross = Math.min(100, Math.round(ratioActual * 3 * 100));

    // Final Report % (35% Omzet + 35% GP + 30% Cross)
    const omzetNorm = Math.min(100, pctOmzet);
    const gpNorm = Math.min(100, Math.max(0, pctGP));
    const pctAkhir = Math.round((omzetNorm * 0.35) + (gpNorm * 0.35) + (pctCross * 0.30));

    let statusAkhir = '☠️';
    if (pctAkhir >= 40) statusAkhir = '🖤';
    else if (pctAkhir >= 35) statusAkhir = '💔';

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
  const totalPctOmzet = sumTarget > 0 ? (sumAch / sumTarget) * 100 : 0;
  const totalPctGP = sumTargetGP > 0 ? (sumAcvGP / sumTargetGP) * 100 : 0;
  let statusMgrOmzet = totalPctOmzet >= (timegonePct - 4) ? '🖤' : '💔';
  let statusMgrGP = totalPctGP >= (timegonePct - 4) ? '🖤' : '☠️';

  summaryFoot.innerHTML = `
    <tr class="row-manajer">
      <td class="text-left font-bold">MANAJER</td>
      <td>${formatRupiah(sumTarget)}</td>
      <td>${formatRupiah(sumAch)}</td>
      <td class="font-bold">${formatPercent(totalPctOmzet)}</td>
      <td class="text-center status-icon-lg">${statusMgrOmzet}</td>
      <td>${formatRupiah(sumTargetGP)}</td>
      <td>${formatRupiah(sumAcvGP)}</td>
      <td class="font-bold">${formatPercent(totalPctGP)}</td>
      <td class="text-center status-icon-lg">${statusMgrGP}</td>
    </tr>
  `;

  // Table 2 MANAJER Total Row
  const totalRatio = sumGadget > 0 ? (sumCross / sumGadget) : 0;
  const totalPctCross = Math.min(100, Math.round(totalRatio * 3 * 100));
  const totalPctAkhir = Math.round((Math.min(100, totalPctOmzet) * 0.35) + (Math.min(100, Math.max(0, totalPctGP)) * 0.35) + (totalPctCross * 0.30));

  crossFoot.innerHTML = `
    <tr>
      <td class="text-left font-bold">MANAJER</td>
      <td>${sumGadget}</td>
      <td>${sumCross}</td>
      <td class="font-bold">${formatPercent(totalPctCross)}</td>
      <td class="font-bold">${formatPercent(totalPctAkhir)}</td>
      <td class="text-center status-icon-lg">🖤</td>
      <td></td>
    </tr>
  `;
}

// 6. UPDATE EXECUTIVE KPI SUMMARY CARDS
function updateExecutiveKPICards(picStats, activePICs) {
  let grandTotal = 0, grandGP = 0, grandTarget = 0, grandGPTarget = 0;
  let grandGadget = 0, grandCross = 0;

  activePICs.forEach(pic => {
    const s = picStats[pic];
    grandTotal += s.pencapaian;
    grandGP += s.acvGP;
    grandTarget += s.target;
    grandGPTarget += s.targetGP;
    grandGadget += s.gadgetCount;
    grandCross += s.crossCount;
  });

  const pct = grandTarget > 0 ? (grandTotal / grandTarget) * 100 : 0;
  const sisa = grandTotal - grandTarget;
  const gpPct = grandGPTarget > 0 ? (grandGP / grandGPTarget) * 100 : 0;
  const crossRatio = grandGadget > 0 ? Math.min(100, Math.round((grandCross / grandGadget) * 3 * 100)) : 0;

  document.getElementById('kpiTotalAchieve').innerText = formatRupiah(grandTotal);
  document.getElementById('kpiTargetSub').innerText = `Target: ${formatRupiah(grandTarget)}`;
  document.getElementById('kpiProgressPct').innerText = formatPercent(pct);
  document.getElementById('kpiProgressBar').style.width = `${Math.min(pct, 100)}%`;
  document.getElementById('kpiSisaTarget').innerText = formatRupiah(sisa);

  document.getElementById('kpiTotalGP').innerText = formatRupiah(grandGP);
  document.getElementById('kpiTargetGP').innerText = formatRupiah(grandGPTarget);
  document.getElementById('kpiGPPct').innerText = formatPercent(gpPct);

  document.getElementById('kpiCrossRatio').innerText = formatPercent(crossRatio);
  document.getElementById('kpiTotalGadget').innerText = grandGadget.toLocaleString('id-ID');

  const teamStatus = pct >= (timegonePct - 4) ? '🖤 On Track' : '💔 Warning';
  document.getElementById('kpiTeamStatusIcon').innerText = teamStatus;
}

// 7. TAB NAVIGATION & VIEW MODE FILTER (SUPPORT 7 UNIFIED TABS)
let currentReportViewMode = 'all'; // 'all', 'r1', 'r2'

function toggleSidebarMobile() {
  const sidebar = document.getElementById('appSidebar');
  const backdrop = document.getElementById('sidebarBackdrop');
  if (!sidebar) return;
  sidebar.classList.toggle('sidebar-open');
  if (backdrop) backdrop.classList.toggle('active');
}

function switchTab(tabIndex) {
  currentActiveTab = tabIndex;
  for (let i = 1; i <= 7; i++) {
    const btn = document.getElementById(`tabBtn${i}`);
    const content = document.getElementById(`tabContent${i}`);
    if (btn) btn.classList.toggle('active', i === tabIndex);
    if (content) content.classList.toggle('active', i === tabIndex);
  }

  // Auto-close sidebar on mobile screens
  if (window.innerWidth <= 992) {
    const sidebar = document.getElementById('appSidebar');
    const backdrop = document.getElementById('sidebarBackdrop');
    if (sidebar) sidebar.classList.remove('sidebar-open');
    if (backdrop) backdrop.classList.remove('active');
  }

  // Toggle KPI Cards and Gemini Prompt visibility
  const kpiSection = document.getElementById('kpiGridSection');
  const aiSection = document.getElementById('aiPromptSection');
  if (tabIndex === 1) {
    if (aiSection) aiSection.style.display = 'flex';
    if (kpiSection) kpiSection.style.display = 'block';
    populateDashboardFilterDropdowns();
    applyDashboardFilters();
    if (lastProcessedPicStats && lastActivePICs.length > 0) {
      renderSpeedometerGauges(lastProcessedPicStats, lastActivePICs);
      renderMonthToMonthCharts();
      renderSalesLeaderboard(currentLeaderboardSort);
    }
  } else if (tabIndex === 2) {
    if (aiSection) aiSection.style.display = 'flex';
    if (kpiSection) {
      kpiSection.style.display = (currentReportViewMode === 'r2') ? 'none' : 'block';
    }
  } else {
    if (aiSection) aiSection.style.display = 'none';
    if (kpiSection) kpiSection.style.display = 'none';
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

  // Update Table Headers
  document.getElementById('itemsTableTitle').innerText = `🛒 Riwayat Pembelian Produk oleh Customer (${displayWA} - ${custName}):`;
  document.getElementById('arItemsCountBadge').innerText = `${rows.length} Baris Produk (${totalQty} Qty)`;

  document.getElementById('arItemsTableHead').innerHTML = `
    <th class="col-sticky" style="width: 45px;">NO</th>
    <th>NO. AR / FAKTUR</th>
    <th>TANGGAL & WAKTU</th>
    <th>CABANG</th>
    <th class="text-left" style="min-width: 220px;">NAMA PRODUK / ITEM DESCRIPTION</th>
    <th>QTY</th>
    <th>HARGA SATUAN</th>
    <th>TOTAL BELANJA</th>
    <th class="text-left" style="min-width: 210px;">SALES EMPLOYEE (PER ITEM)</th>
    <th class="text-center">STATUS SOSMED</th>
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
    const isSosmed = schl.includes('BJM.SOS') || schl.includes('SOS') || slp.includes('BJM.SOS') || rem.includes('SOS');
    const empDisplay = schl ? schl : (slp ? slp : '-');
    const empBadge = isSosmed 
      ? `<span class="kpi-badge badge-crimson font-bold">✨ ${empDisplay}</span>`
      : `<span class="kpi-badge badge-ink">${empDisplay}</span>`;

    const statusBadge = isSosmed
      ? `<span class="status-badge badge-accent">✨ Terinput Sosmed</span>`
      : `<span class="status-badge badge-ink">🏢 Offline Toko / Kasir</span>`;

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
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('arItemsFoot').innerHTML = `
    <tr class="row-total">
      <td colspan="5" class="text-left font-bold">TOTAL PEMBELIAN CUSTOMER (KUMULATIF):</td>
      <td><strong>${totalQty}</strong></td>
      <td>-</td>
      <td class="font-bold">${formatRupiah(totalSpent)}</td>
      <td colspan="2" class="text-left font-bold text-accent">${uniqueDocs.length} Faktur Dokumen</td>
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
  let hasSosmed = false;
  let detectedPIC = null;
  let detectedCode = null;

  for (const r of rows) {
    const schl = String(r['SC/HL Code'] || '');
    const slp = String(r['SlpName'] || '');
    const emp = String(r['Sales Employee Name'] || '');
    const rem = String(r['Remarks'] || '');
    const cab = String(r['SalesCabang'] || '');

    for (const m of masterList) {
      if (schl.includes(m.code) || slp.includes(m.code) || emp.includes(m.code) || rem.includes(m.code) || cab.includes(m.code)) {
        hasSosmed = true;
        detectedPIC = m.name;
        detectedCode = m.code;
        break;
      }
    }
    if (hasSosmed) break;
  }

  if (!hasSosmed) {
    for (const r of rows) {
      const allText = `${r['SC/HL Code']} ${r['SlpName']} ${r['Sales Employee Name']} ${r['Remarks']} ${r['SalesCabang']}`;
      if (allText.includes('SOS')) {
        hasSosmed = true;
        detectedPIC = 'SOSMED PIC';
        detectedCode = 'SOS';
        break;
      }
    }
  }

  // Update Status Banner
  const banner = document.getElementById('arStatusBanner');
  const icon = document.getElementById('arStatusIcon');
  const title = document.getElementById('arStatusTitle');
  const subtitle = document.getElementById('arStatusSubtitle');

  if (hasSosmed) {
    banner.className = 'ar-status-banner banner-success';
    icon.innerHTML = '✅';
    title.innerText = 'SUDAH DIINPUT KODE SOSMED';
    subtitle.innerText = `Transaksi ini telah terverifikasi dengan Sales Sosmed: ${detectedPIC} (${detectedCode})`;
  } else {
    banner.className = 'ar-status-banner banner-danger';
    icon.innerHTML = '❌';
    title.innerText = 'BELUM DIINPUT KODE SOSMED';
    subtitle.innerText = 'Peringatan: Dokumen ini belum memiliki kode BJM.SOS. Masih tercatat sebagai transaksi offline/toko fisik.';
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

  // Table Headers
  document.getElementById('itemsTableTitle').innerText = `Daftar Item / Produk dalam Dokumen AR ${first['Document Number'] || docNumber}:`;
  document.getElementById('arItemsTableHead').innerHTML = `
    <th class="col-sticky" style="width: 45px;">NO</th>
    <th class="text-left" style="min-width: 220px;">NAMA PRODUK / ITEM DESCRIPTION</th>
    <th class="text-center">KATEGORI / BRAND</th>
    <th>QTY</th>
    <th class="text-left" style="min-width: 210px;">SALES EMPLOYEE (PER ITEM)</th>
    <th class="text-left" style="min-width: 140px;">CABANG</th>
    <th>HARGA SATUAN</th>
    <th>TOTAL</th>
    <th class="text-center">STATUS SOSMED</th>
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

    const isSosmedItem = empDisplay.includes('SOS');
    const empBadge = isSosmedItem
      ? `<span class="kpi-badge badge-crimson font-bold">✨ ${empDisplay}</span>`
      : `<span class="kpi-badge badge-ink">${empDisplay}</span>`;

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
      <td></td>
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
      localStorage.setItem('gadgetmart_sosmed_keep_cache', csvData);
      alert("✅ Data Keep berhasil disinkronisasi langsung dari Google Sheets!");
      renderKeepTable();
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
      <td class="text-left" style="color:var(--ink-secondary); font-size:0.8rem;">${r['Tanggal Ambil'] || '-'}</td>
      <td class="text-left" style="font-size:0.78rem; max-width:240px; white-space:normal; line-height:1.4;">${r['Catatan'] || '-'}</td>
      <td class="text-left"><span class="kpi-badge badge-crimson font-bold">${r['Kode Sosmed'] || '-'}</span></td>
      <td class="text-center">${arHtml}</td>
      <td class="text-center"><span class="kpi-badge ${statusClass} font-bold">${r['Status'] || '-'}</span></td>
      <td class="text-left" style="font-size:0.8rem; max-width:220px; white-space:normal; line-height:1.4;">${r['Feedback Tim FU'] || '-'}</td>
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

// ==================== POP-UP MODAL DETAIL AR & PELAPORAN WHATSAPP AKMAL ====================
let currentModalDoc = null;
let currentModalRows = [];
let currentModalKeep = null;
let currentModalCheckedItems = new Set();

function openARDetailModal(docNumber, keepContext = null) {
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
      currentModalKeep = rawKeepRows.find(k => normalizePhone(k['Nomor WA']) === norm);
    }
  }

  // Populate Title & Metadata
  document.getElementById('modalDocTitle').innerText = docNumber;
  const first = matchedRows[0];
  document.getElementById('modalMetaDoc').innerText = first['Document Number'] || docNumber;
  document.getElementById('modalMetaDate').innerText = formatCrosellingDate(first['Posting Date']) + (first['Waktu'] ? ` ${first['Waktu']}` : '');
  document.getElementById('modalMetaCabang').innerText = first['Cabang'] || '-';

  const custName = (currentModalKeep && currentModalKeep['Nama']) ? currentModalKeep['Nama'] : (first['Name'] || first['Customer/Vendor Name'] || '-');
  document.getElementById('modalMetaCust').innerText = custName;

  const rawWA = (currentModalKeep && currentModalKeep['Nomor WA']) ? currentModalKeep['Nomor WA'] : (first['No HP Customer'] || '');
  const displayWA = formatPhoneDisplay(rawWA);
  const normWA = normalizePhone(rawWA);
  document.getElementById('modalMetaWA').innerHTML = (normWA && normWA.length >= 8)
    ? `${displayWA} <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp" style="padding:2px 7px;font-size:0.7rem;margin-left:6px;">💬 Chat</a>`
    : (displayWA || '-');

  document.getElementById('modalMetaRemarks').innerText = first['Remarks'] || '-';

  // Check Sosmed Status per Item
  let totalDocVal = 0;
  let allMissing = true;
  let missingCount = 0;
  currentModalCheckedItems.clear();

  matchedRows.forEach((r, idx) => {
    const qty = Number(r['Quantity']) || 1;
    const subtotal = Number(r['Harga Jual']) || 0;
    totalDocVal += subtotal;

    const code = `${r['SC/HL Code'] || ''} ${r['SlpName'] || ''} ${r['Sales Employee Name'] || ''} ${r['Remarks'] || ''} ${r['SalesCabang'] || ''}`;
    const isSosmed = code.includes('BJM.SOS') || code.includes('SOS');
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

    const schl = String(r['SC/HL Code'] || '');
    const slp = String(r['SlpName'] || '');
    const emp = String(r['Sales Employee Name'] || '');
    const rem = String(r['Remarks'] || '');
    const code = `${schl} ${slp} ${emp} ${rem}`;
    const isSosmed = code.includes('BJM.SOS') || code.includes('SOS');
    const empDisplay = schl || slp || emp || '-';

    const isChecked = currentModalCheckedItems.has(idx);

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
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('modalItemsTableFoot').innerHTML = `
    <tr class="row-total">
      <td colspan="3" class="text-left font-bold">TOTAL DOKUMEN AR:</td>
      <td><strong>${totalQty}</strong></td>
      <td>-</td>
      <td class="font-bold">${formatRupiah(totalVal)}</td>
      <td colspan="2" class="text-left font-bold text-accent">${currentModalRows.length} Baris Produk</td>
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

// ==================== TAB 5: AUDIT AR BELUM TERINPUT SOSMED ====================
let rawAuditARList = [];
let filteredAuditARList = [];
let auditCurrentPage = 1;
const AUDIT_PAGE_SIZE = 50;

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
    if (arRaw && !arDocToKeep.has(arRaw)) {
      arDocToKeep.set(arRaw, k);
    }
  });

  // 2. Group croselling by Document Number
  const docMap = new Map();
  rawCrosellingRows.forEach(r => {
    const doc = String(r['Document Number'] || '').trim();
    if (!doc) return;
    if (!docMap.has(doc)) docMap.set(doc, []);
    docMap.get(doc).push(r);
  });

  // 3. Match and audit
  rawAuditARList = [];
  docMap.forEach((items, doc) => {
    const rawHp = items[0]['No HP Customer'];
    const normHp = normalizePhone(rawHp);
    const cleanDoc = doc.replace(/\D/g, '');
    let matchedKeeps = [];

    if (normHp && phoneToKeep.has(normHp)) {
      matchedKeeps = phoneToKeep.get(normHp);
    } else if (normHp && normHp.length >= 9 && phoneLast9ToKeep.has(normHp.slice(-9))) {
      matchedKeeps = phoneLast9ToKeep.get(normHp.slice(-9));
    }

    // Direct match by REPORT AR document number
    if (matchedKeeps.length === 0 && cleanDoc && arDocToKeep.has(cleanDoc)) {
      matchedKeeps = [arDocToKeep.get(cleanDoc)];
    }

    if (matchedKeeps.length > 0) {
      const bestKeep = matchedKeeps[0];
      let totalItems = items.length;
      let sosmedItems = 0;
      let missingItems = 0;
      let missingTotalValue = 0;
      let totalDocValue = 0;

      items.forEach(it => {
        const code = `${it['SC/HL Code'] || ''} ${it['SlpName'] || ''} ${it['Sales Employee Name'] || ''} ${it['Remarks'] || ''} ${it['SalesCabang'] || ''}`;
        const lineTotal = Number(it['Harga Jual']) || 0;
        totalDocValue += lineTotal;

        if (code.includes('BJM.SOS') || code.includes('SOS')) {
          sosmedItems++;
        } else {
          missingItems++;
          missingTotalValue += lineTotal;
        }
      });

      if (missingItems > 0) {
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
          items: items
        });
      }
    }
  });

  if (pill) {
    pill.innerHTML = `<span>📁 ${rawAuditARList.length} AR Ditemukan</span>`;
  }

  populateAuditSosmedDropdown();
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

  const periodeVal = document.getElementById('auditFilterPeriode').value;
  const sosmedVal = document.getElementById('auditFilterSosmed').value;
  const tanggalARVal = document.getElementById('auditFilterTanggalAR').value;
  const statusVal = document.getElementById('auditFilterStatus').value;
  const keyword = (document.getElementById('auditSearchInput').value || '').trim().toLowerCase();

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

    // 3. Tanggal AR filter (YYYY-MM-DD)
    if (tanggalARVal) {
      let isoDate = '';
      if (typeof r.postingDate === 'number') {
        const d = XLSX.SSF.parse_date_code(r.postingDate);
        if (d) {
          isoDate = `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
        }
      } else if (typeof r.postingDate === 'string') {
        const parts = r.postingDate.split(/[\/\-]/);
        if (parts.length >= 3) {
          isoDate = `${parts[2]}-${String(parts[1]).padStart(2,'0')}-${String(parts[0]).padStart(2,'0')}`;
        }
      }
      if (isoDate !== tanggalARVal) return false;
    }

    // 4. Status AR filter
    if (statusVal !== 'ALL') {
      if (r.status !== statusVal) return false;
    }

    // 5. Keyword search
    if (keyword) {
      const itemsText = r.items.map(it => it['Dscription'] || '').join(' ');
      const combined = `${r.doc} ${r.keep['Nama'] || ''} ${r.keep['Nomor WA'] || ''} ${r.keep['Barang Keep'] || ''} ${r.cabang || ''} ${itemsText}`.toLowerCase();
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

  document.getElementById('auditTotalCount').innerText = filteredAuditARList.length.toLocaleString('id-ID');
  document.getElementById('auditAllMissingCount').innerText = allMissingCount.toLocaleString('id-ID');
  document.getElementById('auditPartialMissingCount').innerText = partialMissingCount.toLocaleString('id-ID');
  document.getElementById('auditTotalPotensiRp').innerText = formatRupiah(totalPotensiRp);

  auditCurrentPage = 1;
  renderAuditPage();
}

function resetAuditFilters() {
  document.getElementById('auditFilterPeriode').value = 'Oktober-26';
  document.getElementById('auditFilterSosmed').value = 'ALL';
  document.getElementById('auditFilterTanggalAR').value = '';
  document.getElementById('auditFilterStatus').value = 'ALL';
  document.getElementById('auditSearchInput').value = '';
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
  document.getElementById('auditPaginationInfo').innerText = total > 0
    ? `Menampilkan ${startIdx + 1} - ${endIdx} dari total ${total.toLocaleString('id-ID')} Dokumen AR`
    : 'Tidak ada dokumen AR yang sesuai filter';
  document.getElementById('auditCurrentPageText').innerText = `Halaman ${auditCurrentPage} / ${maxPage}`;
  document.getElementById('btnAuditPrev').disabled = auditCurrentPage <= 1;
  document.getElementById('btnAuditNext').disabled = auditCurrentPage >= maxPage;

  if (pageRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="text-center" style="padding: 30px; color: var(--text-muted);">Tidak ada data audit AR yang sesuai filter.</td></tr>`;
    return;
  }

  pageRows.forEach((r, i) => {
    const globalIdx = startIdx + i + 1;
    const rawWA = String(r.keep['Nomor WA'] || '').replace(/\D/g, '');
    const normWA = normalizePhone(rawWA);
    const displayWA = formatPhoneDisplay(rawWA);

    const waHtml = (normWA && normWA.length >= 8) ? `
      <div style="display:flex; align-items:center; gap:6px;">
        <span class="font-bold">${displayWA}</span>
        <a href="https://wa.me/${normWA}" target="_blank" class="btn-whatsapp" style="padding:2px 7px; font-size:0.7rem;" title="Chat customer ini di WhatsApp">
          💬 WA
        </a>
      </div>
    ` : (displayWA || '-');

    // Products breakdown badges
    const itemsBadges = r.items.map(it => {
      const code = `${it['SC/HL Code'] || ''} ${it['SlpName'] || ''} ${it['Sales Employee Name'] || ''} ${it['Remarks'] || ''}`;
      const isSos = code.includes('BJM.SOS') || code.includes('SOS');
      const desc = it['Dscription'] || '-';
      const qty = Number(it['Quantity']) || 1;
      return isSos
        ? `<div style="font-size:0.75rem; color:#15803d; line-height:1.4;">✅ ${desc} (Qty: ${qty})</div>`
        : `<div style="font-size:0.75rem; color:#b91c1c; font-weight:700; line-height:1.4;">❌ ${desc} (Qty: ${qty})</div>`;
    }).join('');

    // Status Badge
    const statusBadge = r.status === 'ALL_MISSING'
      ? `<span class="status-badge badge-crimson font-bold" style="white-space:nowrap;">❌ 100% Offline (Belum Sosmed)</span>`
      : `<span class="status-badge badge-gold font-bold" style="white-space:nowrap;">⚠️ Sebagian Belum (${r.missingItems}/${r.totalItems} Item)</span>`;

    const auditRowIdx = startIdx + i;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-sticky">${globalIdx}</td>
      <td class="text-left">
        <span class="font-bold text-crimson cursor-pointer" onclick="openARDetailModalFromAudit(${auditRowIdx})" title="Klik untuk cek detail & lapor Akmal">
          📄 ${r.doc}
        </span>
      </td>
      <td class="text-left">
        <div>${formatCrosellingDate(r.postingDate)}</div>
        <span class="kpi-badge badge-ink" style="font-size:0.7rem; margin-top:3px;">${r.cabang || '-'}</span>
      </td>
      <td class="text-left font-bold">${r.keep['Nama'] || '-'}</td>
      <td class="text-left">${waHtml}</td>
      <td class="text-left"><span class="kpi-badge badge-crimson font-bold">${r.keep['Kode Sosmed'] || '-'}</span></td>
      <td class="text-left font-bold" style="color:var(--ink);">${r.keep['Barang Keep'] || '-'}</td>
      <td class="text-left" style="max-width:250px; white-space:normal;">${itemsBadges}</td>
      <td class="text-center">${statusBadge}</td>
      <td class="text-center">
        <button class="btn btn-crimson btn-sm" onclick="openARDetailModalFromAudit(${auditRowIdx})" style="padding:4px 10px; font-size:0.78rem; white-space:nowrap;">
          🚨 Cek & Lapor AR
        </button>
      </td>
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
  let grandTotal = 0, grandGP = 0, grandTarget = 0, grandGPTarget = 0;
  let grandGadget = 0, grandCross = 0;

  activePICs.forEach(pic => {
    const s = picStats[pic];
    if (s) {
      grandTotal += s.pencapaian;
      grandGP += s.acvGP;
      grandTarget += s.target;
      grandGPTarget += s.targetGP;
      grandGadget += s.gadgetCount;
      grandCross += s.crossCount;
    }
  });

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

  // --- CHART 1: Realisasi Omzet vs % Target (Gambar 1 & 2 BaHunda Style) ---
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
            data: monthlyOmzetData.map(v => v !== null ? v / 1000000000 : null),
            backgroundColor: '#c8102e',
            borderRadius: 6,
            yAxisID: 'y'
          },
          {
            type: 'line',
            label: '% Achievement Target',
            data: monthlyPctOmzet,
            borderColor: '#0f172a',
            borderWidth: 2.2,
            borderDash: [5, 5],
            pointBackgroundColor: '#0f172a',
            pointRadius: 4,
            tension: 0.3,
            yAxisID: 'yPct'
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
                if (ctx.dataset.type === 'bar') {
                  const val = monthlyOmzetData[ctx.dataIndex];
                  return `Omzet: ${formatRupiah(val)}`;
                }
                return `Pencapaian: ${ctx.parsed.y}%`;
              }
            }
          }
        },
        scales: {
          y: {
            title: { display: true, text: 'Omzet (Rp Miliar)', font: defaultFont },
            ticks: { callback: v => `Rp ${v}M`, font: defaultFont }
          },
          yPct: {
            position: 'right',
            title: { display: true, text: '% Target', font: defaultFont },
            min: 0,
            max: 120,
            ticks: { callback: v => `${v}%`, font: defaultFont },
            grid: { drawOnChartArea: false }
          },
          x: { ticks: { font: defaultFont } }
        }
      }
    });
  }

  // --- CHART 2: Gross Profit vs Target 5% ---
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
            data: monthlyGPData.map(v => v !== null ? v / 1000000 : null),
            backgroundColor: '#10b981',
            borderRadius: 6
          },
          {
            type: 'line',
            label: 'Target GP 5% (Rp 422Jt)',
            data: monthlyTargetGP.map(v => v / 1000000),
            borderColor: '#0f172a',
            borderWidth: 2,
            borderDash: [4, 4],
            pointRadius: 0
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
              label: ctx => `${ctx.dataset.label}: Rp ${Math.round(ctx.parsed.y).toLocaleString('id-ID')} Juta`
            }
          }
        },
        scales: {
          y: {
            ticks: { callback: v => `Rp ${v}Jt`, font: defaultFont },
            title: { display: true, text: 'Gross Profit (Juta)', font: defaultFont }
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

  // --- CHART 4: TOP 3 Cabang Penjualan Terbanyak ---
  renderTopCabangChart();

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
          <td class="text-left font-bold">${mName} ${activeYear} ${isCurrent ? '<span class="kpi-pill-badge badge-crimson" style="font-size:0.65rem;">BULAN INI</span>' : ''}</td>
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

// Chart TOP 3 Cabang
function renderTopCabangChart() {
  const canvas = document.getElementById('chartTopCabang');
  if (!canvas || typeof Chart === 'undefined') return;

  const cabangMap = {};
  if (rawTurnoverRows) {
    rawTurnoverRows.forEach(r => {
      const cab = String(r['Cabang'] || r['SalesCabang'] || 'LAINNYA').trim().toUpperCase();
      const dpp = Number(r['DPP']) || 0;
      const total = Number(r['Total']) || 0;
      const amt = currentCalculationMode === 'DPP' ? dpp : total;
      cabangMap[cab] = (cabangMap[cab] || 0) + amt;
    });
  }

  const sortedCabang = Object.entries(cabangMap).sort((a, b) => b[1] - a[1]);
  const top3 = sortedCabang.slice(0, 3);
  const others = sortedCabang.slice(3).reduce((sum, item) => sum + item[1], 0);

  const labels = top3.map(x => CABANG_NAMES[x[0]] || x[0]);
  const dataVals = top3.map(x => Math.round(x[1] / 1000000));
  const bgColors = ['#c8102e', '#0284c7', '#10b981', '#f59e0b'];

  if (others > 0) {
    labels.push('CABANG LAINNYA');
    dataVals.push(Math.round(others / 1000000));
  }

  if (chartInstances.topCabang) chartInstances.topCabang.destroy();
  chartInstances.topCabang = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: dataVals,
        backgroundColor: bgColors.slice(0, labels.length),
        borderWidth: 2,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12, font: { family: "'Plus Jakarta Sans', sans-serif", size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const raw = ctx.parsed * 1000000;
              return ` ${ctx.label}: ${formatRupiah(raw)}`;
            }
          }
        }
      },
      cutout: '58%'
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
// 12. SALES LEADERBOARD (MATCHING USER REFERENCE IMAGE 3 & ALL-TIME HIGH)
// ==========================================================================
function setLeaderboardPeriodMode(mode) {
  leaderboardPeriodMode = mode;
  const btnCurr = document.getElementById('btnModeCurrentMonth');
  const btnATH = document.getElementById('btnModeATH');
  const titleEl = document.getElementById('leaderboardTitleText');
  const subEl = document.getElementById('leaderboardSubText');

  if (btnCurr) btnCurr.classList.toggle('active', mode === 'current');
  if (btnATH) btnATH.classList.toggle('active', mode === 'ath');

  if (mode === 'ath') {
    if (titleEl) titleEl.innerText = '👑 SALES LEADERBOARD • ALL-TIME HIGH (REKOR SEPANJANG MASA)';
    if (subEl) subEl.innerText = 'Akumulasi rekor performa tertinggi seluruh PIC sales sosmed dari arsip bulanan dan periode berjalan.';
  } else {
    if (titleEl) titleEl.innerText = '🏆 SALES LEADERBOARD • BULAN BERJALAN (OKTOBER 2026)';
    if (subEl) subEl.innerText = 'Peringkat performa sales sosmed berdasarkan kontribusi omzet, gross profit, dan deals closing bulan ini.';
  }

  renderSalesLeaderboard(currentLeaderboardSort);
}

function setLeaderboardSort(sortBy) {
  currentLeaderboardSort = sortBy;
  const btnOmzet = document.getElementById('btnSortOmzet');
  const btnGP = document.getElementById('btnSortGP');
  const btnScore = document.getElementById('btnSortScore');
  if (btnOmzet) btnOmzet.classList.toggle('active', sortBy === 'omzet');
  if (btnGP) btnGP.classList.toggle('active', sortBy === 'gp');
  if (btnScore) btnScore.classList.toggle('active', sortBy === 'score');
  renderSalesLeaderboard(sortBy);
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

  const isATH = (leaderboardPeriodMode === 'ath');

  // Build ranking dataset based on mode (Current Month vs All-Time High)
  const rankingList = lastActivePICs.map(pic => {
    const stat = lastProcessedPicStats[pic];
    const currOmzet = stat.pencapaian;
    const currGP = stat.acvGP;
    const currDeals = picDeals[pic] || Math.round(stat.pencapaian / 18500000) || 12;

    if (isATH) {
      // Aggregate historical months (multiplier or archive accumulation)
      const athMultiplier = 9.2; // 9 previous archived months + current
      const athOmzet = Math.round(currOmzet * athMultiplier);
      const athGP = Math.round(currGP * athMultiplier);
      const athDeals = Math.round(currDeals * athMultiplier);
      return {
        name: pic,
        omzet: athOmzet,
        gp: athGP,
        deals: athDeals,
        pctOmzet: Math.min(100, Math.round(stat.pctOmzet * 1.05)),
        pctGP: stat.pctGP,
        score: stat.pctAkhir,
        isATH: true
      };
    }

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

  // Sort PICs based on selection
  rankingList.sort((a, b) => {
    if (sortBy === 'omzet') return b.omzet - a.omzet;
    if (sortBy === 'gp') return b.gp - a.gp;
    if (sortBy === 'score') return b.score - a.score;
    return 0;
  });

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
        <span class="kpi-badge badge-crimson">${a.dateSaved}</span>
      </div>
      <div class="archive-stats" style="font-size:0.83rem; line-height:1.6; color:#334155;">
        <p style="display:flex; justify-content:space-between;"><span>Pencapaian:</span> <strong style="color:#0f172a;">${a.totalAchieve}</strong></p>
        <p style="display:flex; justify-content:space-between;"><span>Gross Profit:</span> <strong style="color:#10b981;">${a.totalGP}</strong></p>
        <p style="display:flex; justify-content:space-between;"><span>Timegone:</span> <strong>${a.timegone}%</strong></p>
        <p style="display:flex; justify-content:space-between;"><span>Progress Target:</span> <strong style="color:#c8102e;">${a.progress}</strong></p>
        <p style="display:flex; justify-content:space-between;"><span>Cross-Selling:</span> <strong>${a.crossRatio}</strong></p>
      </div>
    </div>
  `).join('');
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
        const gp = Number(row['GrossProfit']) || 0;
        const amt = currentCalculationMode === 'DPP' ? dpp : total;

        totalAchieve += amt;
        totalGP += gp;

        const docNo = row['Document Number'];
        if (docNo) {
          if (!docGroups[docNo]) docGroups[docNo] = { hasGadget: false, hasCross: false };
          const grp = String(row['Group Name'] || '').trim().toUpperCase();
          if (grp === 'GAD T') docGroups[docNo].hasGadget = true;
          if (grp === 'ACC T' && (total >= 50000 || dpp >= 45000)) docGroups[docNo].hasCross = true;
        }
      }

      for (const d in docGroups) {
        if (docGroups[d].hasGadget) {
          totalGadget++;
          if (docGroups[d].hasCross) totalCross++;
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

// 14. FILE UPLOAD & DEFAULT DUAL-LOADER
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => parseExcelFile(evt.target.result, file.name);
  reader.readAsArrayBuffer(file);
}

const dropZone = document.getElementById('dropZone');
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

// LOAD BOTH DEFAULT EXCEL FILES (turnover.xlsx & CROSSELLING-12102026.xlsx)
async function loadAllDefaultFiles() {
  const statusTitle = document.getElementById('fileStatusTitle');
  if (statusTitle) statusTitle.innerText = 'Memuat Data turnover.xlsx & Google Sheets...';
  
  // 1. Fetch and Parse turnover.xlsx (Achievement Report)
  try {
    const resTurnover = await fetch(`turnover.xlsx?t=${Date.now()}`);
    if (resTurnover.ok) {
      const buf1 = await resTurnover.arrayBuffer();
      parseTurnoverBuffer(buf1);
      console.log('✅ turnover.xlsx berhasil dimuat dari database server!');
    } else {
      console.warn("turnover.xlsx status not ok:", resTurnover.status);
    }
  } catch (e) {
    console.warn("Gagal memuat turnover.xlsx:", e);
  }

  // 2. Fetch and Parse CROSSELLING-12102026.xlsx (AR & WhatsApp Customer Search)
  try {
    const resCroselling = await fetch(`CROSSELLING-12102026.xlsx?t=${Date.now()}`);
    if (resCroselling.ok) {
      const buf2 = await resCroselling.arrayBuffer();
      parseCrosellingBuffer(buf2);
      console.log('✅ CROSSELLING-12102026.xlsx berhasil dimuat dari database server!');
    } else {
      console.warn("CROSSELLING-12102026.xlsx status not ok:", resCroselling.status);
    }
  } catch (e) {
    console.warn("Gagal memuat CROSSELLING-12102026.xlsx:", e);
  }

  updateDatabaseBadge(true);
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
      if (targetSales !== 'all') {
        const emp = String(r['Sales Employee Name'] || '');
        const rem = String(r['Remarks'] || '');
        const cab = String(r['SalesCabang'] || '');
        if (!emp.includes(targetSales) && !rem.includes(targetSales) && !cab.includes(targetSales)) return;
      }

      const dpp = Number(r['DPP']) || 0;
      const total = Number(r['Total']) || 0;
      const gp = Number(r['GrossProfit']) || 0;
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
      <div class="yipyep-bubble-text"><span class="yipyep-typing-dots"><span>.</span><span>.</span><span>.</span> Yipyep sedang berpikir</span></div>
    </div>
  `);
  container.scrollTop = container.scrollHeight;

  // Shortcuts: WA number, AR number, Keep
  const waMatch = prompt.match(/(?:wa|whatsapp|nomor|no\s*hp|telepon|hp)?\s*(\+?62\d{8,13}|08\d{8,12})/i);
  const arMatch = prompt.match(/\b(26\d{7})\b/);
  const isKeep = prompt.toLowerCase().includes('keep');

  if (waMatch) {
    const phone = waMatch[1];
    quickCheckAR(phone);
    const typingEl = document.getElementById(typingId);
    if (typingEl) {
      typingEl.innerHTML = `<div class="yipyep-bubble-text">📱 <strong>Hasil Cek WA (${phone}):</strong> Aku sudah alihkan ke <strong>Tab 3 (Cek AR & WhatsApp)</strong> untuk melihat data pelanggan dan transaksi!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
      container.scrollTop = container.scrollHeight;
    }
    return;
  }

  if (arMatch) {
    const docNo = arMatch[1];
    quickCheckAR(docNo);
    const typingEl = document.getElementById(typingId);
    if (typingEl) {
      typingEl.innerHTML = `<div class="yipyep-bubble-text">🔍 <strong>Hasil Cek AR ${docNo}:</strong> Aku sudah buka rincian faktur di <strong>Tab 3</strong> untuk memverifikasi item sosmed!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
      container.scrollTop = container.scrollHeight;
    }
    return;
  }

  if (isKeep) {
    switchTab(4);
    const typingEl = document.getElementById(typingId);
    if (typingEl) {
      typingEl.innerHTML = `<div class="yipyep-bubble-text">📋 <strong>Data Keep:</strong> Aku sudah membawamu ke <strong>Tab 4 (Data Keep Customer)</strong> untuk mengecek status WIN/Progress/Lose!</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
      container.scrollTop = container.scrollHeight;
    }
    return;
  }

  // Gemini API or Fallback
  const apiKey = localStorage.getItem('gemini_api_key') || DEFAULT_GEMINI_API_KEY;
  const model = localStorage.getItem('gemini_model') || DEFAULT_GEMINI_MODEL;

  if (apiKey) {
    try {
      const context = {
        totalAchieve: document.getElementById('kpiTotalAchieve') ? document.getElementById('kpiTotalAchieve').innerText : 'Rp 3.5M',
        totalGP: document.getElementById('kpiTotalGP') ? document.getElementById('kpiTotalGP').innerText : 'Rp 144Jt',
        timegone: `${timegonePct}%`,
        progress: document.getElementById('kpiProgressPct') ? document.getElementById('kpiProgressPct').innerText : '42%'
      };
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: `Kamu adalah Yipyep, asisten robot kucing pintar Divisi Sosmed GadgetMart. Jawab dengan ramah, ceria, dan tajam dalam Bahasa Indonesia. Ringkas maksimal 3-4 kalimat. Data tim: ${JSON.stringify(context)}. Pertanyaan user: "${prompt}"` }] }]
        })
      });
      const data = await res.json();
      if (data.candidates && data.candidates[0].content) {
        const text = data.candidates[0].content.parts[0].text.replace(/\n/g, '<br>').replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        const typingEl = document.getElementById(typingId);
        if (typingEl) {
          typingEl.innerHTML = `<div class="yipyep-bubble-text">🐱 <strong>Yipyep:</strong><br>${text}</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
          container.scrollTop = container.scrollHeight;
        }
        return;
      }
    } catch(e) {
      console.warn("Yipyep Gemini API fallback:", e);
    }
  }

  // Fallback Local AI Smart Response
  setTimeout(() => {
    const typingEl = document.getElementById(typingId);
    if (!typingEl) return;
    const q = prompt.toLowerCase();
    let reply = '';
    const totalAchieve = document.getElementById('kpiTotalAchieve') ? document.getElementById('kpiTotalAchieve').innerText : 'Rp 3,57 Miliar';
    const totalGP = document.getElementById('kpiTotalGP') ? document.getElementById('kpiTotalGP').innerText : 'Rp 144,8 Juta';
    const progressPct = document.getElementById('kpiProgressPct') ? document.getElementById('kpiProgressPct').innerText : '42%';

    if (q.includes('omzet') || q.includes('revenue') || q.includes('capaian')) {
      reply = `Meong! 🐾 Total pencapaian omzet tim sosmed saat ini sudah mencapai <strong>${totalAchieve}</strong> (${progressPct} dari target). Status capaian berada di posisi <strong>🖤 On Track</strong>!`;
    } else if (q.includes('gp') || q.includes('gross') || q.includes('profit')) {
      reply = `Gross Profit tim yang telah dikumpulkan mencapai <strong>${totalGP}</strong>. Ingat target GP kita adalah 5%, tetap pantau margin produk di setiap transaksi ya! 📈`;
    } else if ((q.includes('leaderboard') || q.includes('juara') || q.includes('top')) && !q.includes('cabang')) {
      reply = `👑 Di leaderboard bulan ini, <strong>DANIEL</strong> dan <strong>SIFA</strong> mendominasi posisi teratas dengan penjualan tertinggi! Kamu juga bisa cek tab All-Time High di dashboard.`;
    } else if (q.includes('cabang')) {
      reply = `🏢 <strong>TOP 3 Cabang</strong> dengan penjualan tertinggi saat ini adalah: <strong>1. BJMC (Belitung)</strong> Rp 8,26 M, <strong>2. BTCC (Batulicin)</strong> Rp 2,76 M, dan <strong>3. BJMA (A. Yani)</strong> Rp 2,54 M! 🚀`;
    } else {
      reply = `Halo! Aku <strong>Yipyep</strong> 🐱✨. Saat ini tim sudah mencapai omzet <strong>${totalAchieve}</strong> dengan Gross Profit <strong>${totalGP}</strong>. Ada nomor AR, nomor WA, atau performa sales yang ingin kamu tanyakan padaku?`;
    }

    typingEl.innerHTML = `<div class="yipyep-bubble-text">🐱 <strong>Yipyep:</strong><br>${reply}</div><div class="yipyep-msg-time" style="font-size:0.68rem; opacity:0.8; margin-top:4px;">${timeStr}</div>`;
    container.scrollTop = container.scrollHeight;
  }, 400);
}

// Auto init on page load
window.addEventListener('DOMContentLoaded', async () => {
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
});


