const fs = require('fs');
const path = require('path');

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzdEByURvrvCNAJ-TlaxlU5GoNvmClgxynK_vYrORzXiTXFoYArwRMCJzdW2LWvT5wfUw/exec';

function updateLocalCsv(nomorWA, arDoc) {
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
          while (cols.length < 15) cols.push('');
          cols[13] = String(arDoc || '').trim();
          if (arDoc) {
            const today = new Date();
            cols[14] = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
          }
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
    let arDoc = '';

    if (req.method === 'POST') {
      const body = req.body || {};
      nomorWA = body.nomorWA || '';
      arDoc = body.arDoc || body.reportAR || body.ar || '';
    } else {
      const urlParams = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
      nomorWA = urlParams.searchParams.get('nomorWA') || (req.query && req.query.nomorWA) || '';
      arDoc = urlParams.searchParams.get('arDoc') || (req.query && (req.query.arDoc || req.query.reportAR || req.query.ar)) || '';
    }

    if (!nomorWA) {
      return res.status(400).json({ success: false, message: 'Nomor WhatsApp tidak boleh kosong' });
    }

    const cleanWA = String(nomorWA).replace(/\D/g, '');
    updateLocalCsv(cleanWA, arDoc);

    const targetUrl = `${APPS_SCRIPT_URL}?action=UPDATE_REPORT_AR&nomorWA=${encodeURIComponent(cleanWA)}&arDoc=${encodeURIComponent(arDoc)}`;
    const gasRes = await fetch(targetUrl, { redirect: 'follow' });
    const gasJson = await gasRes.json();

    return res.status(200).json(gasJson);
  } catch (err) {
    console.error('API update-report-ar error:', err);
    return res.status(500).json({ success: false, message: err.message || 'Gagal terhubung ke Google Sheets' });
  }
};
