const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const PUBLIC_DIR = __dirname;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.json': 'application/json',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.csv': 'text/csv; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml'
};

const GOOGLE_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1V0pIa0NZCprJqRwnSnikFaLZ-dd50nY0uE3bjVYsUe0/export?format=csv&gid=1014933310';

const server = http.createServer(async (req, res) => {
  // Enable CORS for all devices in LAN
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  let reqPath = req.url.split('?')[0];

  // API Live Sync Endpoint for Google Sheets Keep Data
  if (reqPath === '/api/sync-keep') {
    try {
      console.log('🔄 Syncing Google Sheets Keep Data live from URL...');
      const response = await fetch(GOOGLE_SHEET_CSV_URL, { redirect: 'follow' });
      if (!response.ok) throw new Error(`Google Sheets responded with status ${response.status}`);
      const text = await response.text();
      const cachePath = path.join(PUBLIC_DIR, 'keep-data.csv');
      fs.writeFileSync(cachePath, text, 'utf8');
      console.log('✅ Google Sheets Keep Data synced successfully! Size:', (text.length / 1024 / 1024).toFixed(2), 'MB');
      
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: 'Data keep berhasil disinkronisasi langsung dari Google Sheets!', bytes: text.length }));
      return;
    } catch (err) {
      console.error('❌ Sync failed:', err.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: err.message }));
      return;
    }
  }

  // API Update Keep Status to Google Sheets via Google Apps Script Web App
  if (reqPath === '/api/update-keep-status') {
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbybKV-XtULQW6ct4GCNCCgk2f84AD4CVsSGCOOzweskJFaZtPk0jnEut8tJ80Aq1naoXA/exec';
    
    let nomorWA = '';
    let status = '';

    if (req.method === 'POST') {
      const chunks = [];
      req.on('data', chunk => chunks.push(chunk));
      req.on('end', async () => {
        try {
          const bodyStr = Buffer.concat(chunks).toString('utf8');
          const body = JSON.parse(bodyStr || '{}');
          nomorWA = body.nomorWA || '';
          status = body.status || '';

          if (!nomorWA || !status) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: false, error: 'nomorWA dan status harus diisi' }));
            return;
          }

          console.log(`📤 Mengirim update status ke Google Sheets: WA=${nomorWA}, Status=${status}`);
          const targetUrl = `${APPS_SCRIPT_URL}?action=UPDATE_STATUS&nomorWA=${encodeURIComponent(nomorWA)}&status=${encodeURIComponent(status)}`;
          const gasRes = await fetch(targetUrl, { redirect: 'follow' });
          const gasJson = await gasRes.json();
          console.log('📥 Respon Google Apps Script:', gasJson);

          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(gasJson));
        } catch (err) {
          console.error('❌ Gagal update status ke Apps Script:', err);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: err.message }));
        }
      });
      return;
    } else {
      // GET method support
      try {
        const urlParams = new URL(req.url, `http://${req.headers.host}`);
        nomorWA = urlParams.searchParams.get('nomorWA') || '';
        status = urlParams.searchParams.get('status') || '';

        if (!nomorWA || !status) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: 'nomorWA dan status harus diisi' }));
          return;
        }

        console.log(`📤 Mengirim update status (GET) ke Google Sheets: WA=${nomorWA}, Status=${status}`);
        const targetUrl = `${APPS_SCRIPT_URL}?action=UPDATE_STATUS&nomorWA=${encodeURIComponent(nomorWA)}&status=${encodeURIComponent(status)}`;
        const gasRes = await fetch(targetUrl, { redirect: 'follow' });
        const gasJson = await gasRes.json();
        console.log('📥 Respon Google Apps Script:', gasJson);

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(gasJson));
      } catch (err) {
        console.error('❌ Gagal update status ke Apps Script:', err);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: err.message }));
      }
      return;
    }
  }

  // API Save Excel File for LAN Persistence (Tersimpan permanen untuk semua laptop)
  if (reqPath === '/api/save-excel' && req.method === 'POST') {
    const urlParams = new URL(req.url, `http://${req.headers.host}`);
    const type = urlParams.searchParams.get('type');
    if (type !== 'turnover' && type !== 'croselling') {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: 'Tipe file tidak valid. Gunakan type=turnover atau type=croselling.' }));
      return;
    }

    const targetFileName = type === 'croselling' ? 'CROSSELLING-12102026.xlsx' : 'turnover.xlsx';
    const targetFilePath = path.join(PUBLIC_DIR, targetFileName);

    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try {
        const buffer = Buffer.concat(chunks);
        if (buffer.length < 500) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: 'Ukuran file terlalu kecil atau tidak valid.' }));
          return;
        }

        fs.writeFileSync(targetFilePath, buffer);
        console.log(`💾 File ${targetFileName} berhasil disimpan ke server (${(buffer.length / 1024 / 1024).toFixed(2)} MB)`);

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ 
          success: true, 
          message: `File ${targetFileName} berhasil disimpan ke server dan tersedia untuk seluruh device di LAN!`,
          fileName: targetFileName,
          size: buffer.length,
          savedAt: new Date().toISOString()
        }));
      } catch (err) {
        console.error('❌ Gagal menyimpan file ke disk:', err);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: err.message }));
      }
    });
    return;
  }

  // API Get Excel File (Ambil file turnover/croselling terbaru dari server)
  if (reqPath === '/api/get-excel') {
    const urlParams = new URL(req.url, `http://${req.headers.host}`);
    const type = urlParams.searchParams.get('type') || 'turnover';
    const targetFileName = type === 'croselling' ? 'CROSSELLING-12102026.xlsx' : 'turnover.xlsx';
    const targetFilePath = path.join(PUBLIC_DIR, targetFileName);

    if (req.method === 'DELETE') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, message: `File ${type} direset ke default.` }));
      return;
    }

    if (fs.existsSync(targetFilePath)) {
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${targetFileName}"`,
        'Cache-Control': 'no-store, no-cache, must-revalidate'
      });
      fs.createReadStream(targetFilePath).pipe(res);
    } else {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'File tidak ditemukan' }));
    }
    return;
  }

  // API Sales Photos (Get & Save)
  if (reqPath === '/api/sales-photos') {
    const photosFile = path.join(PUBLIC_DIR, 'sales_photos.json');
    if (req.method === 'GET') {
      if (fs.existsSync(photosFile)) {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        fs.createReadStream(photosFile).pipe(res);
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({}));
      }
      return;
    }
    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => body += chunk);
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          fs.writeFileSync(photosFile, JSON.stringify(parsed, null, 2), 'utf8');
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, message: 'Foto sales berhasil disimpan ke server!' }));
        } catch (err) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: err.message }));
        }
      });
      return;
    }
  }

  // API AR Claims Data (Get & Save for LAN & Multi-Device Persistence)
  if (reqPath === '/api/ar-claims') {
    const claimsFile = path.join(PUBLIC_DIR, 'ar_claims.json');
    if (req.method === 'GET') {
      if (fs.existsSync(claimsFile)) {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        fs.createReadStream(claimsFile).pipe(res);
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify([]));
      }
      return;
    }
    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => body += chunk);
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          fs.writeFileSync(claimsFile, JSON.stringify(parsed, null, 2), 'utf8');
          console.log(`💾 Data Klaim AR berhasil disimpan (${Array.isArray(parsed) ? parsed.length : 0} item)`);
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, count: Array.isArray(parsed) ? parsed.length : 0 }));
        } catch (err) {
          console.error('❌ Gagal menyimpan data klaim AR:', err);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: err.message }));
        }
      });
      return;
    }
  }

  // API Monthly Archives (Get & Save)
  if (reqPath === '/api/monthly-archives') {
    const archivesFile = path.join(PUBLIC_DIR, 'monthly_archives.json');
    if (req.method === 'GET') {
      if (fs.existsSync(archivesFile)) {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        fs.createReadStream(archivesFile).pipe(res);
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify([]));
      }
      return;
    }
    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => body += chunk);
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          fs.writeFileSync(archivesFile, JSON.stringify(parsed, null, 2), 'utf8');
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, message: 'Arsip bulanan berhasil disimpan ke server!' }));
        } catch (err) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: err.message }));
        }
      });
      return;
    }
  }

  // API Monthly Targets (Get & Save)
  if (reqPath === '/api/monthly-targets') {
    const targetsFile = path.join(PUBLIC_DIR, 'monthly_targets.json');
    if (req.method === 'GET') {
      if (fs.existsSync(targetsFile)) {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        fs.createReadStream(targetsFile).pipe(res);
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({}));
      }
      return;
    }
    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => body += chunk);
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          fs.writeFileSync(targetsFile, JSON.stringify(parsed, null, 2), 'utf8');
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, message: 'Target bulanan berhasil disimpan ke server!' }));
        } catch (err) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: err.message }));
        }
      });
      return;
    }
  }

  // API Database Status Check (Untuk Realtime Cross-Device Sync)
  if (reqPath === '/api/database-status') {
    const checkFile = (f) => {
      try {
        const p = path.join(PUBLIC_DIR, f);
        if (fs.existsSync(p)) {
          const s = fs.statSync(p);
          return { exists: true, size: s.size, mtimeMs: s.mtimeMs, mtime: s.mtime };
        }
      } catch (e) {}
      return { exists: false, mtimeMs: 0 };
    };

    const status = {
      turnover: checkFile('turnover.xlsx'),
      croselling: checkFile('CROSSELLING-12102026.xlsx'),
      keep: checkFile('keep-data.csv'),
      archives: checkFile('monthly_archives.json'),
      targets: checkFile('monthly_targets.json'),
      photos: checkFile('sales_photos.json'),
      serverTime: new Date().toISOString()
    };
    res.writeHead(200, { 
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store, no-cache, must-revalidate'
    });
    res.end(JSON.stringify(status));
    return;
  }

  if (reqPath === '/' || reqPath === '') reqPath = '/index.html';

  const safePath = path.normalize(reqPath).replace(/^(\.\.[\/\\])+/, '');
  const filePath = path.join(PUBLIC_DIR, safePath);

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, {
      'Content-Type': contentType,
      'Content-Length': stats.size,
      'Cache-Control': 'no-cache'
    });

    const readStream = fs.createReadStream(filePath);
    readStream.pipe(res);
  });
});

server.listen(PORT, () => {
  console.log(`\n======================================================`);
  console.log(`🚀 SOSMED DAILY ACHIEVEMENT DASHBOARD BERJALAN!`);
  console.log(`📡 URL Akses: http://localhost:${PORT}`);
  console.log(`📁 File default otomatis termuat: CROSSELLING-12102026.xlsx`);
  console.log(`======================================================\n`);
});
