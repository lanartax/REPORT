const fs = require('fs');
const path = require('path');
const { getDb, initDb } = require('../lib/db');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const sql = getDb();
  if (sql) await initDb(sql);

  const localFile = path.join(process.cwd(), 'sales_photos.json');

  if (req.method === 'GET') {
    if (sql) {
      try {
        const rows = await sql`SELECT pic_name, photo_url FROM sales_photos`;
        const result = {};
        rows.forEach(r => {
          result[r.pic_name] = r.photo_url;
        });
        return res.status(200).json(result);
      } catch (err) {
        console.error("Neon photos fetch error:", err);
      }
    }

    try {
      if (fs.existsSync(localFile)) {
        const data = fs.readFileSync(localFile, 'utf8');
        return res.status(200).json(JSON.parse(data));
      }
    } catch (e) {}
    return res.status(200).json({});
  }

  if (req.method === 'POST') {
    let payload = req.body;
    if (typeof payload === 'string') {
      try { payload = JSON.parse(payload); } catch(e){}
    }

    if (sql && payload) {
      try {
        for (const [picName, photoUrl] of Object.entries(payload)) {
          if (!photoUrl) {
            await sql`DELETE FROM sales_photos WHERE pic_name = ${picName}`;
          } else {
            await sql`
              INSERT INTO sales_photos (pic_name, photo_url, updated_at)
              VALUES (${picName}, ${photoUrl}, NOW())
              ON CONFLICT (pic_name) DO UPDATE
              SET photo_url = EXCLUDED.photo_url,
                  updated_at = NOW()
            `;
          }
        }
        return res.status(200).json({ success: true, storage: 'neon' });
      } catch (err) {
        console.error("Neon photos save error:", err);
      }
    }

    try {
      fs.writeFileSync(localFile, JSON.stringify(payload, null, 2), 'utf8');
      return res.status(200).json({ success: true, storage: 'file' });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
