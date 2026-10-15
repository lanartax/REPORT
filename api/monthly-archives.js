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

  const localFile = path.join(process.cwd(), 'monthly_archives.json');

  if (req.method === 'GET') {
    if (sql) {
      try {
        const rows = await sql`SELECT data FROM monthly_archives ORDER BY created_at DESC`;
        const list = rows.map(r => r.data);
        return res.status(200).json(list);
      } catch (err) {
        console.error("Neon archives fetch error:", err);
      }
    }

    try {
      if (fs.existsSync(localFile)) {
        const data = fs.readFileSync(localFile, 'utf8');
        return res.status(200).json(JSON.parse(data));
      }
    } catch (e) {}
    return res.status(200).json([]);
  }

  if (req.method === 'POST') {
    let payload = req.body;
    if (typeof payload === 'string') {
      try { payload = JSON.parse(payload); } catch(e){}
    }

    if (sql && Array.isArray(payload)) {
      try {
        for (const item of payload) {
          const archiveId = String(item.id || `${item.year}-${item.month}`);
          await sql`
            INSERT INTO monthly_archives (id, month, year, title, data, created_at)
            VALUES (${archiveId}, ${item.month || null}, ${item.year || null}, ${item.title || ''}, ${JSON.stringify(item)}, NOW())
            ON CONFLICT (id) DO UPDATE
            SET month = EXCLUDED.month,
                year = EXCLUDED.year,
                title = EXCLUDED.title,
                data = EXCLUDED.data
          `;
        }
        return res.status(200).json({ success: true, storage: 'neon' });
      } catch (err) {
        console.error("Neon archives save error:", err);
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
