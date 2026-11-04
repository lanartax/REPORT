const fs = require('fs');
const path = require('path');
const os = require('os');
const { getDb, initDb } = require('../lib/db');

const LOCAL_MESSAGES_FILE = path.join(os.tmpdir(), 'watzap_messages.json');

function getLocalMessages() {
  try {
    if (fs.existsSync(LOCAL_MESSAGES_FILE)) {
      return JSON.parse(fs.readFileSync(LOCAL_MESSAGES_FILE, 'utf8')) || [];
    }
  } catch (e) {}
  return [];
}

function appendLocalMessage(msg) {
  try {
    const list = getLocalMessages();
    list.push(msg);
    fs.writeFileSync(LOCAL_MESSAGES_FILE, JSON.stringify(list.slice(-500), null, 2), 'utf8');
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

  // GET: Sinkronisasi pesan masuk untuk frontend (Ultra-fast Neon Postgres)
  if (req.method === 'GET') {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    const since = parseInt(url.searchParams.get('since') || '0', 10);
    const phoneFilter = (url.searchParams.get('phone') || '').replace(/\D/g, '');

    let msgs = [];
    if (sql) {
      try {
        await initDb(sql);
        if (since > 0) {
          if (phoneFilter) {
            const last9 = phoneFilter.slice(-9);
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp"
              FROM watzap_inbox
              WHERE is_group = FALSE 
                AND created_at > ${since}
                AND (phone = ${phoneFilter} OR phone LIKE ${'%' + last9})
              ORDER BY created_at ASC
              LIMIT 200
            `;
          } else {
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp"
              FROM watzap_inbox
              WHERE is_group = FALSE AND created_at > ${since}
              ORDER BY created_at ASC
              LIMIT 200
            `;
          }
        } else {
          // Fetch initial recent messages
          if (phoneFilter) {
            const last9 = phoneFilter.slice(-9);
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp"
              FROM watzap_inbox
              WHERE is_group = FALSE AND (phone = ${phoneFilter} OR phone LIKE ${'%' + last9})
              ORDER BY created_at ASC
              LIMIT 500
            `;
          } else {
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp"
              FROM watzap_inbox
              WHERE is_group = FALSE
              ORDER BY created_at ASC
              LIMIT 500
            `;
          }
        }
      } catch (dbErr) {
        console.error('Neon DB GET Error, falling back to local:', dbErr);
        msgs = getLocalMessages();
        if (since > 0) msgs = msgs.filter(m => (m.timestamp || 0) > since);
      }
    } else {
      msgs = getLocalMessages();
      if (since > 0) msgs = msgs.filter(m => (m.timestamp || 0) > since);
    }

    // Strict filter: pastikan bukan grup dan jadikan timestamp sebagai Number
    const filtered = (msgs || []).filter(m => {
      const p = String(m.phone || '');
      return !m.isGroup && !p.startsWith('120363') && p.length <= 14;
    }).map(m => ({
      ...m,
      timestamp: Number(m.timestamp || m.created_at || Date.now())
    }));

    return res.status(200).json({
      success: true,
      count: filtered.length,
      messages: filtered
    });
  }

  // POST: Webhook payload dari WatZap saat ada chat masuk
  if (req.method === 'POST') {
    try {
      const body = req.body || {};
      const data = body.data || body;

      const rawChatId = String(data.chat_id || data.phone || data.from || data.sender || data.phone_no || body.phone || '');
      const cleanPhone = rawChatId.replace(/@.*$/, '').replace(/\D/g, '');

      // DETEKSI GRUP WHATSAPP: TOLAK & ABAIKAN GRUP
      const isGroup = rawChatId.includes('@g.us') || cleanPhone.startsWith('120363') || cleanPhone.length > 14;
      if (isGroup) {
        return res.status(200).json({
          status: 200,
          message: 'Group message ignored (Personal PIC chats only)'
        });
      }

      // Validasi nomor WhatsApp personal (minimal 9 digit, maksimal 14 digit)
      if (!cleanPhone || cleanPhone.length < 9) {
        return res.status(200).json({ status: 200, message: 'Invalid phone ignored' });
      }

      const text = String(data.message_body || data.message_text || data.message || data.body || data.text || body.message || '').trim();
      const senderName = String(data.name || data.pushName || '').trim();
      const isFromMe = data.is_from_me === true || body.is_from_me === true;
      const messageId = String(data.message_id || data.id || ('wz_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7)));

      let ts = Date.now();
      if (data.timestamp || body.timestamp) {
        const parsed = Number(data.timestamp || body.timestamp);
        ts = parsed > 10000000000 ? parsed : parsed * 1000;
      }

      const senderType = isFromMe ? 'ADMIN' : 'PIC';

      if (text) {
        if (sql) {
          try {
            await initDb(sql);
            await sql`
              INSERT INTO watzap_inbox (message_id, phone, sender_name, message, sender, is_group, created_at)
              VALUES (${messageId}, ${cleanPhone}, ${senderName}, ${text}, ${senderType}, FALSE, ${ts})
              ON CONFLICT (message_id) DO NOTHING
            `;
          } catch (dbErr) {
            console.error('Neon DB POST Insert Error:', dbErr);
          }
        }
        appendLocalMessage({
          messageId: messageId,
          phone: cleanPhone,
          message: text,
          sender: senderType,
          senderName: senderName,
          isGroup: false,
          timestamp: ts
        });
      }

      return res.status(200).json({
        status: 200,
        message: 'Personal chat processed successfully'
      });
    } catch (err) {
      console.error('Webhook error:', err);
      return res.status(200).json({ status: 200, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
