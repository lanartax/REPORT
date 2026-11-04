const fs = require('fs');
const path = require('path');
const os = require('os');
const { getDb, initDb } = require('../lib/db');

const LOCAL_RESOLVE_FILE = path.join(os.tmpdir(), 'omni_resolved_chats.json');

function getLocalResolved() {
  try {
    if (fs.existsSync(LOCAL_RESOLVE_FILE)) {
      return JSON.parse(fs.readFileSync(LOCAL_RESOLVE_FILE, 'utf8')) || {};
    }
  } catch (e) {}
  return {};
}

function setLocalResolved(phone, isResolved, resolvedBy) {
  try {
    const data = getLocalResolved();
    if (isResolved) {
      data[phone] = { isResolved: true, resolvedAt: Date.now(), resolvedBy: resolvedBy || 'Admin' };
    } else {
      delete data[phone];
    }
    fs.writeFileSync(LOCAL_RESOLVE_FILE, JSON.stringify(data, null, 2), 'utf8');
  } catch (e) {}
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const sql = getDb();

  // GET: Ambil daftar nomor yang di-resolve secara permanen
  if (req.method === 'GET') {
    let resolvedMap = {};
    if (sql) {
      try {
        await initDb(sql);
        const rows = await sql`
          SELECT phone, is_resolved, resolved_at, resolved_by
          FROM omni_chat_status
          WHERE is_resolved = TRUE
        `;
        for (const r of rows) {
          resolvedMap[r.phone] = {
            isResolved: true,
            resolvedAt: Number(r.resolved_at || Date.now()),
            resolvedBy: r.resolved_by || 'Admin'
          };
        }
      } catch (err) {
        console.error('Neon DB omni-resolve GET Error:', err);
        resolvedMap = getLocalResolved();
      }
    } else {
      resolvedMap = getLocalResolved();
    }

    return res.status(200).json({
      success: true,
      resolvedChats: resolvedMap
    });
  }

  // POST: Simpan atau buka kembali (unresolve) chat secara permanen
  if (req.method === 'POST') {
    try {
      const body = req.body || {};
      let phone = String(body.phone || body.picKey || '').replace(/\D/g, '');
      if (phone.startsWith('08')) phone = '62' + phone.slice(1);
      else if (phone.startsWith('8')) phone = '62' + phone;

      if (!phone || phone.length < 9) {
        return res.status(400).json({ success: false, error: 'Nomor telepon tidak valid' });
      }

      const isResolved = body.isResolved === true || body.action === 'resolve';
      const resolvedBy = String(body.resolvedBy || 'Admin').trim();
      const now = Date.now();

      if (sql) {
        try {
          await initDb(sql);
          await sql`
            INSERT INTO omni_chat_status (phone, is_resolved, resolved_at, resolved_by, updated_at)
            VALUES (${phone}, ${isResolved}, ${now}, ${resolvedBy}, ${now})
            ON CONFLICT (phone) DO UPDATE SET
              is_resolved = EXCLUDED.is_resolved,
              resolved_at = EXCLUDED.resolved_at,
              resolved_by = EXCLUDED.resolved_by,
              updated_at = EXCLUDED.updated_at
          `;
        } catch (dbErr) {
          console.error('Neon DB omni-resolve POST Error:', dbErr);
        }
      }

      setLocalResolved(phone, isResolved, resolvedBy);

      return res.status(200).json({
        success: true,
        phone: phone,
        isResolved: isResolved,
        resolvedAt: now,
        resolvedBy: resolvedBy
      });
    } catch (err) {
      console.error('omni-resolve Error:', err);
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};

