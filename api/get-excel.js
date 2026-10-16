const fs = require('fs');
const path = require('path');
const { getDb, initDb } = require('../lib/db.js');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, DELETE, OPTIONS');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const fileType = (req.query && req.query.type) ? req.query.type : 'turnover';
  const targetFileName = fileType === 'croselling' ? 'CROSSELLING-12102026.xlsx' : 'turnover.xlsx';

  // Handle DELETE request (untuk fitur Kembali ke Default)
  if (req.method === 'DELETE') {
    try {
      const sql = getDb();
      if (sql) {
        await initDb(sql);
        await sql`DELETE FROM uploaded_files WHERE file_key = ${fileType}`;
      }
      return res.status(200).json({ success: true, message: `File ${fileType} di database telah di-reset ke default.` });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // 1. Cek dari Database PostgreSQL (Neon.tech) terlebih dahulu
  try {
    const sql = getDb();
    if (sql) {
      await initDb(sql);
      const rows = await sql`SELECT file_data, timestamp, updated_at FROM uploaded_files WHERE file_key = ${fileType}`;
      if (rows && rows.length > 0 && rows[0].file_data) {
        const buffer = Buffer.from(rows[0].file_data, 'base64');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${targetFileName}"`);
        res.setHeader('X-Custom-File', 'true');
        res.setHeader('X-File-Timestamp', rows[0].timestamp || '');
        return res.status(200).send(buffer);
      }
    }
  } catch(dbErr) {
    console.warn("DB get-excel fallback:", dbErr.message);
  }

  // 2. Fallback ke file default di disk lokal jika belum ada upload di DB
  try {
    const localPath = path.join(process.cwd(), targetFileName);
    if (fs.existsSync(localPath)) {
      const buffer = fs.readFileSync(localPath);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${targetFileName}"`);
      res.setHeader('X-Custom-File', 'false');
      return res.status(200).send(buffer);
    }
  } catch(e) {}

  return res.status(404).json({ error: 'File tidak ditemukan' });
};
