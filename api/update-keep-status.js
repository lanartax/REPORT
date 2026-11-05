const fs = require('fs');
const path = require('path');

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbybKV-XtULQW6ct4GCNCCgk2f84AD4CVsSGCOOzweskJFaZtPk0jnEut8tJ80Aq1naoXA/exec';

function updateLocalCsv(nomorWA, newStatus, feedback = null, reportAR = null) {
  try {
    const csvPath = path.join(process.cwd(), 'keep-data.csv');
    if (!fs.existsSync(csvPath)) return;

    const content = fs.readFileSync(csvPath, 'utf8');
    const lines = content.split('\n');
    if (lines.length <= 1) return;

    const targetWA = String(nomorWA).replace(/\D/g, '');
    const targetLast9 = targetWA.slice(-9);

    let updated = false;
    for (let i = lines.length - 1; i >= 1; i--) {
      const line = lines[i];
      if (!line.trim()) continue;

      const cols = line.split(',');
      if (cols.length > 7) {
        const rowWA = String(cols[7] || '').replace(/\D/g, '');
        if (rowWA === targetWA || (rowWA.length >= 9 && rowWA.slice(-9) === targetLast9)) {
          while (cols.length < 23) cols.push('');
          if (reportAR !== null && reportAR !== undefined && String(reportAR).trim() !== '') {
            cols[13] = String(reportAR).trim();
            const today = new Date();
            cols[14] = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
          }
          if (newStatus !== null && newStatus !== undefined) cols[21] = String(newStatus).trim();
          if (feedback !== null && feedback !== undefined) cols[22] = String(feedback).trim();
          lines[i] = cols.join(',');
          updated = true;
          break;
        }
      }
    }

    if (updated) {
      fs.writeFileSync(csvPath, lines.join('\n'), 'utf8');
    }
  } catch (e) {
    // Abaikan jika read-only di serverless environment
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  try {
    let nomorWA = '';
    let status = '';
    let feedback = '';
    let reportAR = '';

    if (req.method === 'POST') {
      const body = req.body || {};
      nomorWA = body.nomorWA || '';
      status = body.status || '';
      feedback = body.feedback || body.feedbackTimFU || '';
      reportAR = body.reportAR || body.arDoc || '';
    } else {
      const urlParams = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
      nomorWA = urlParams.searchParams.get('nomorWA') || (req.query && req.query.nomorWA) || '';
      status = urlParams.searchParams.get('status') || (req.query && req.query.status) || '';
      feedback = urlParams.searchParams.get('feedback') || urlParams.searchParams.get('feedbackTimFU') || (req.query && (req.query.feedback || req.query.feedbackTimFU)) || '';
      reportAR = urlParams.searchParams.get('reportAR') || urlParams.searchParams.get('arDoc') || '';
    }

    if (!nomorWA) {
      return res.status(400).json({ success: false, message: 'Nomor WhatsApp tidak boleh kosong' });
    }

    const cleanWA = String(nomorWA).replace(/\D/g, '');
    updateLocalCsv(cleanWA, status, feedback, reportAR);

    let targetUrl = `${APPS_SCRIPT_URL}?action=UPDATE_STATUS&nomorWA=${encodeURIComponent(cleanWA)}&status=${encodeURIComponent(status || '')}`;
    if (feedback) {
      targetUrl += `&feedback=${encodeURIComponent(feedback)}`;
    }
    const gasRes = await fetch(targetUrl, { redirect: 'follow' });
    const gasJson = await gasRes.json();

    if (reportAR) {
      try {
        const arUrl = `${APPS_SCRIPT_URL}?action=UPDATE_REPORT_AR&nomorWA=${encodeURIComponent(cleanWA)}&arDoc=${encodeURIComponent(reportAR)}`;
        await fetch(arUrl, { redirect: 'follow' });
      } catch (arErr) {
        console.warn('Gagal update reportAR di Apps Script:', arErr);
      }
    }

    return res.status(200).json(gasJson);
  } catch (err) {
    console.error('API update-keep-status error:', err);
    return res.status(500).json({ success: false, message: err.message || 'Gagal terhubung ke Google Sheets' });
  }
};
