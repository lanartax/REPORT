const fs = require('fs');
const path = require('path');
const { getDb, initDb } = require('../lib/db.js');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const fileType = (req.query && req.query.type) ? req.query.type : 'turnover';
  const targetFileName = fileType === 'croselling' ? 'CROSSELLING-12102026.xlsx' : 'turnover.xlsx';
  const targetPath = path.join(process.cwd(), targetFileName);

  try {
    let chunks = [];
    if (req.body && Buffer.isBuffer(req.body)) {
      chunks = [req.body];
    } else {
      for await (const chunk of req) {
        chunks.push(chunk);
      }
    }
    const buffer = Buffer.concat(chunks);
    if (!buffer || buffer.length === 0) {
      return res.status(400).json({ error: 'Data file kosong.' });
    }

    // 1. Simpan ke Database PostgreSQL (Neon.tech) agar PERMANEN di serverless & multi-device
    let dbSaved = false;
    try {
      const sql = getDb();
      if (sql) {
        await initDb(sql);
        const base64Data = buffer.toString('base64');
        const updateTs = new Date().toISOString();
        await sql`
          INSERT INTO uploaded_files (file_key, file_name, file_data, timestamp, updated_at)
          VALUES (${fileType}, ${targetFileName}, ${base64Data}, ${updateTs}, NOW())
          ON CONFLICT (file_key) DO UPDATE SET
            file_name = EXCLUDED.file_name,
            file_data = EXCLUDED.file_data,
            timestamp = EXCLUDED.timestamp,
            updated_at = NOW();
        `;
        dbSaved = true;
        console.log(`✅ File ${targetFileName} berhasil disimpan ke PostgreSQL database (${(buffer.length/1024/1024).toFixed(2)} MB)!`);
      }
    } catch (dbErr) {
      console.warn("Gagal simpan ke Neon DB:", dbErr.message);
    }

    // 2. Simpan juga ke disk lokal jika writable (misal saat running di local-server / localhost)
    try {
      fs.writeFileSync(targetPath, buffer);
    } catch(e) {
      // Di serverless read-only disk tidak apa-apa karena sudah tersimpan di database PostgreSQL
    }

    return res.status(200).json({
      success: true,
      message: `File ${targetFileName} berhasil disimpan ke database server (${buffer.length} bytes)`,
      dbSaved,
      size: buffer.length
    });
  } catch (err) {
    console.error("Gagal menyimpan Excel:", err);
    return res.status(500).json({ error: err.message });
  }
};
