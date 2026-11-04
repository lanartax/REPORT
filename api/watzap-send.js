const fs = require('fs');
const path = require('path');
const { getDb, initDb } = require('../lib/db');

const WATZAP_API_KEY = 'VMUXOT68QD5TQ84V';
const WATZAP_NUMBER_KEY = 'gqLbEDAZ4lReCCME';
const MESSAGES_FILE = path.join(process.cwd(), 'watzap_messages.json');

async function resolvePhoneOrLid(raw) {
  let str = String(raw).trim();
  if (str.endsWith('@lid') || (str.length >= 13 && !str.startsWith('628') && !str.startsWith('08') && !str.startsWith('8'))) {
    try {
      const q = str.includes('@') ? str : `${str}@lid`;
      const res = await fetch(`http://178.83.188.100:3088/api/contacts?session=default&contactId=${encodeURIComponent(q)}`, {
        headers: { 'X-Api-Key': 'gadgetmart123' },
        signal: AbortSignal.timeout(3000)
      });
      if (res.ok) {
        const d = await res.json();
        const num = d.id || '';
        let clean = String(num).replace(/@.*$/, '').replace(/\D/g, '');
        if (clean.length >= 9 && clean.length <= 15) {
          if (clean.startsWith('08')) clean = '62' + clean.slice(1);
          else if (clean.startsWith('8')) clean = '62' + clean;
          return clean;
        }
      }
    } catch (e) {}
  }
  let clean = str.replace(/@.*$/, '').replace(/\D/g, '');
  if (clean.startsWith('0')) clean = '62' + clean.slice(1);
  else if (clean.startsWith('8')) clean = '62' + clean;
  else if (!clean.startsWith('62')) clean = '62' + clean;
  return clean;
}

function appendStoredMessage(msg) {
  try {
    let list = [];
    if (fs.existsSync(MESSAGES_FILE)) {
      list = JSON.parse(fs.readFileSync(MESSAGES_FILE, 'utf8')) || [];
    }
    const idx = list.findIndex(m => m.messageId && m.messageId === msg.messageId);
    if (idx >= 0) {
      list[idx] = { ...list[idx], ...msg };
    } else {
      list.push(msg);
    }
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

    const cleanPhone = await resolvePhoneOrLid(rawPhone);
    const now = Number(body.timestamp) || Date.now();
    let msgId = String(body.messageId || ('out_' + now + '_' + Math.random().toString(36).slice(2, 7)));

    const mediaUrl = body.mediaUrl || body.imageUrl || '';
    const replyTo = body.replyTo || null;
    console.log(`📤 [OUTGOING MESSAGE] To: ${cleanPhone} (${mediaUrl ? 'Image' : 'Text'})`);

    let sentSuccessfully = false;
    let gatewaySource = 'WATZAP';

    // 1. Prioritaskan Pengiriman via WAHA VPS (Super Cepat & Native Quoted Reply)
    try {
      const wahaUrl = mediaUrl ? 'http://178.83.188.100:3088/api/sendImage' : 'http://178.83.188.100:3088/api/sendText';
      const wahaPayload = mediaUrl ? {
        session: 'default',
        chatId: `${cleanPhone}@c.us`,
        file: { url: mediaUrl },
        caption: message
      } : {
        session: 'default',
        chatId: `${cleanPhone}@c.us`,
        text: message
      };

      if (replyTo && (replyTo.id || replyTo.messageId)) {
        wahaPayload.reply_to = replyTo.id || replyTo.messageId;
      }

      const wahaRes = await fetch(wahaUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Api-Key': 'gadgetmart123'
        },
        body: JSON.stringify(wahaPayload),
        signal: AbortSignal.timeout(6000)
      });

      if (wahaRes.ok) {
        sentSuccessfully = true;
        gatewaySource = 'WAHA_VPS';
        const wData = await wahaRes.json().catch(() => ({}));
        if (wData.id) {
          msgId = String(wData.id);
        }
        console.log(`✅ [WAHA VPS] Pesan berhasil terkirim ke WhatsApp ${cleanPhone}!`);
      } else {
        console.warn(`⚠️ [WAHA VPS] HTTP ${wahaRes.status}, beralih fallback ke WatZap API...`);
      }
    } catch (wahaErr) {
      console.warn(`⚠️ [WAHA VPS] Gagal: ${wahaErr.message}, beralih fallback ke WatZap API...`);
    }

    // 2. Fallback ke WatZap API jika WAHA belum merespons
    if (!sentSuccessfully) {
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

      const wzRes = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(10000)
      });
      const wzJson = await wzRes.json();
      console.log('WatZap send result:', wzJson);
    }

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
