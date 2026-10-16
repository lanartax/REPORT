const fs = require('fs');
const path = require('path');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const checkFile = (fileName) => {
    try {
      const p = path.join(process.cwd(), fileName);
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

  return res.status(200).json(status);
};
