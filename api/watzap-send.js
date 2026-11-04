const fs = require('fs');
const path = require('path');
const { getDb, initDb } = require('../lib/db');

const WATZAP_API_KEY = 'VMUXOT68QD5TQ84V';
const WATZAP_NUMBER_KEY = 'gqLbEDAZ4lReCCME';
const MESSAGES_FILE = path.join(process.cwd(), 'watzap_messages.json');

function appendStoredMessage(msg) {
  try {
    let list = [];
    if (fs.existsSync(MESSAGES_FILE)) {
      list = JSON.parse(fs.readFileSync(MESSAGES_FILE, 'utf8')) || [];
    }
    list.push(msg);
    const trimmed = list.slice(-500);
    fs.writeFileSync(MESSAGES_FILE, JSON.stringify(trimmed, null, 2), 'utf8');
  } catch (e) { }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  try {
    const body = req.body || {};
    const rawPhone = body.phone || body.phone_no || '';
    const message = (body.message || '').trim();

    if (!rawPhone || !message) {
      return res.status(400).json({ success: false, message: 'Nomor telepon dan pesan wajib diisi!' });
    }

    let cleanPhone = String(rawPhone).replace(/\D/g, '');
    if (cleanPhone.startsWith('0')) {
      cleanPhone = '62' + cleanPhone.slice(1);
    } else if (!cleanPhone.startsWith('62')) {
      cleanPhone = '62' + cleanPhone;
    }

    const mediaUrl = body.mediaUrl || body.imageUrl || '';
    const replyTo = body.replyTo || null;
    const endpoint = mediaUrl ? 'https://api.watzap.id/v1/send_image_url' : 'https://api.watzap.id/v1/send_message';

    const payload = mediaUrl ? {
      api_key: WATZAP_API_KEY,
      number_key: WATZAP_NUMBER_KEY,
      phone_no: cleanPhone,
      url: mediaUrl,
      message: message,
      wait_until_send: '0'
    } : {
      api_key: WATZAP_API_KEY,
      number_key: WATZAP_NUMBER_KEY,
      phone_no: cleanPhone,
      message: message,
      wait_until_send: '0'
    };

    console.log(`📤 [WATZAP SENDING MESSAGE] To: ${cleanPhone} (${mediaUrl ? 'Image' : 'Text'})`);

    const wzRes = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const wzJson = await wzRes.json();
    console.log('WatZap send result:', wzJson);

    const now = Number(body.timestamp) || Date.now();
    const msgId = String(body.messageId || ('out_' + now + '_' + Math.random().toString(36).slice(2, 7)));

    // Simpan ke database Neon PostgreSQL secara realtime
    const sql = getDb();
    if (sql) {
      try {
        await initDb(sql);
        await sql`
          INSERT INTO watzap_inbox (message_id, phone, sender_name, message, sender, is_group, created_at, media_url, message_type, reply_to)
          VALUES (${msgId}, ${cleanPhone}, 'Admin', ${message}, 'ADMIN', FALSE, ${now}, ${mediaUrl || null}, ${mediaUrl ? 'image' : 'text'}, ${replyTo ? JSON.stringify(replyTo) : null})
          ON CONFLICT (message_id) DO NOTHING
        `;
      } catch (dbErr) {
        console.error('Neon DB Outgoing Insert Error:', dbErr);
      }
    }

    // Backup ke local messages file
    appendStoredMessage({
      messageId: msgId,
      phone: cleanPhone,
      message: message,
      sender: 'ADMIN',
      type: 'text',
      timestamp: now,
      mediaUrl: mediaUrl,
      replyTo: replyTo
    });

    return res.status(200).json({
      success: wzJson.status === '200' || wzJson.status === 200 || wzJson.ack === 'successfully',
      watzap: wzJson,
      phone: cleanPhone,
      messageId: msgId,
      timestamp: now
    });
  } catch (err) {
    console.error('Error sending WatZap message:', err);
    return res.status(500).json({ success: false, message: err.message });
  }
};
