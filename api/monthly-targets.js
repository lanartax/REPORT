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

  const localFile = path.join(process.cwd(), 'monthly_targets.json');

  if (req.method === 'GET') {
    if (sql) {
      try {
        const rows = await sql`SELECT key, month, year, team_target, sales_targets FROM monthly_targets`;
        const result = {};
        rows.forEach(r => {
          result[r.key] = {
            month: r.month,
            year: r.year,
            teamTarget: Number(r.team_target),
            salesTargets: r.sales_targets
          };
        });
        return res.status(200).json(result);
      } catch (err) {
        console.error("Neon targets fetch error:", err);
      }
    }

    // Fallback to local file
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
        for (const [key, val] of Object.entries(payload)) {
          await sql`
            INSERT INTO monthly_targets (key, month, year, team_target, sales_targets, updated_at)
            VALUES (${key}, ${val.month || 10}, ${val.year || 2026}, ${val.teamTarget || 0}, ${JSON.stringify(val.salesTargets || {})}, NOW())
            ON CONFLICT (key) DO UPDATE
            SET month = EXCLUDED.month,
                year = EXCLUDED.year,
                team_target = EXCLUDED.team_target,
                sales_targets = EXCLUDED.sales_targets,
                updated_at = NOW()
          `;
        }
        return res.status(200).json({ success: true, storage: 'neon' });
      } catch (err) {
        console.error("Neon targets save error:", err);
      }
    }

    // Fallback to local file
    try {
      fs.writeFileSync(localFile, JSON.stringify(payload, null, 2), 'utf8');
      return res.status(200).json({ success: true, storage: 'file' });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
