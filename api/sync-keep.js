const fs = require('fs');
const path = require('path');

const GOOGLE_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1V0pIa0NZCprJqRwnSnikFaLZ-dd50nY0uE3bjVYsUe0/export?format=csv&gid=1014933310';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  try {
    console.log('🔄 Syncing Google Sheets Keep Data live from URL...');
    const response = await fetch(GOOGLE_SHEET_CSV_URL, { redirect: 'follow' });
    if (!response.ok) {
      throw new Error(`Google Sheets HTTP status ${response.status}`);
    }

    const csvText = await response.text();
    const localFile = path.join(process.cwd(), 'keep-data.csv');

    try {
      fs.writeFileSync(localFile, csvText, 'utf8');
    } catch (e) {
      // In read-only serverless environment, writing to disk might be skipped
    }

    return res.status(200).json({
      success: true,
      message: 'Data keep berhasil disinkronkan langsung dari Google Sheets!',
      rowsSample: csvText.slice(0, 150)
    });
  } catch (err) {
    console.error("Gagal sync Google Sheets:", err);
    return res.status(500).json({ error: err.message });
  }
};
