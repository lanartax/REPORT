const fs = require('fs');
const path = require('path');

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

    try {
      fs.writeFileSync(targetPath, buffer);
    } catch(e) {
      console.warn("Disk write skipped on read-only serverless environment");
    }

    return res.status(200).json({
      success: true,
      message: `File ${targetFileName} berhasil diproses (${buffer.length} bytes)`,
      size: buffer.length
    });
  } catch (err) {
    console.error("Gagal menyimpan Excel:", err);
    return res.status(500).json({ error: err.message });
  }
};
