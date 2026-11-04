const fs = require('fs');
const path = require('path');
const os = require('os');

const CLOUD_STORE_URL = 'https://api.restful-api.dev/objects/ff808181a09d98f701a1a032afdb18bc';
const LOCAL_MESSAGES_FILE = path.join(os.tmpdir(), 'watzap_messages.json');

async function getStoredMessages() {
  try {
    const res = await fetch(CLOUD_STORE_URL);
    if (res.ok) {
      const json = await res.json();
      if (json && json.data && Array.isArray(json.data.messages)) {
        return json.data.messages;
      }
    }
  } catch (e) {}

  try {
    if (fs.existsSync(LOCAL_MESSAGES_FILE)) {
      return JSON.parse(fs.readFileSync(LOCAL_MESSAGES_FILE, 'utf8')) || [];
    }
  } catch (e) {}

  return [];
}

async function appendStoredMessage(newMsg) {
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const currentList = await getStoredMessages();
      const exists = currentList.some(m => m.messageId === newMsg.messageId);
      if (exists) return true;

      // Filter out any legacy group messages if present
      const personalOnly = currentList.filter(m => {
        const p = String(m.phone || '');
        return !m.isGroup && !p.startsWith('120363') && p.length <= 14;
      });

      personalOnly.push(newMsg);
      const trimmed = personalOnly.slice(-1000);

      const res = await fetch(CLOUD_STORE_URL, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: 'gadgetmart_watzap_inbox',
          data: { messages: trimmed }
        })
      });

      try {
        fs.writeFileSync(LOCAL_MESSAGES_FILE, JSON.stringify(trimmed, null, 2), 'utf8');
      } catch (e) {}

      if (res.ok) return true;
    } catch (e) {
      await new Promise(r => setTimeout(r, 200));
    }
  }
  return false;
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  // GET: Fetch incoming messages for frontend sync (HANYA CHAT PERSONAL)
  if (req.method === 'GET') {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    const since = parseInt(url.searchParams.get('since') || '0', 10);
    const phoneFilter = (url.searchParams.get('phone') || '').replace(/\D/g, '');

    let msgs = await getStoredMessages();
    // Strict filter: abaikan grup
    msgs = msgs.filter(m => {
      const p = String(m.phone || '');
      return !m.isGroup && !p.startsWith('120363') && p.length <= 14;
    });

    if (since > 0) {
      msgs = msgs.filter(m => (m.timestamp || 0) > since);
    }
    if (phoneFilter) {
      const last9 = phoneFilter.slice(-9);
      msgs = msgs.filter(m => {
        const clean = String(m.phone || '').replace(/\D/g, '');
        return clean === phoneFilter || (last9 && clean.endsWith(last9));
      });
    }

    return res.status(200).json({
      success: true,
      count: msgs.length,
      messages: msgs
    });
  }

  // POST: Webhook payload from WatZap
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
      const senderName = data.name || data.pushName || '';
      const isFromMe = data.is_from_me === true || body.is_from_me === true;
      const messageId = data.message_id || data.id || ('wz_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7));

      let ts = Date.now();
      if (data.timestamp || body.timestamp) {
        const parsed = Number(data.timestamp || body.timestamp);
        ts = parsed > 10000000000 ? parsed : parsed * 1000;
      }

      if (text) {
        await appendStoredMessage({
          messageId: messageId,
          phone: cleanPhone,
          message: text,
          sender: isFromMe ? 'ADMIN' : 'PIC',
          senderName: senderName,
          isGroup: false,
          type: data.type || body.type || 'text',
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
