const fs = require('fs');
const path = require('path');
const os = require('os');

const CLOUD_STORE_URL = 'https://api.restful-api.dev/objects/ff808181a09d98f701a1a032afdb18bc';
const LOCAL_MESSAGES_FILE = path.join(os.tmpdir(), 'watzap_messages.json');

async function getStoredMessages() {
  // 1. Try cloud store first (persistent across all Vercel instances)
  try {
    const res = await fetch(CLOUD_STORE_URL);
    if (res.ok) {
      const json = await res.json();
      if (json && json.data && Array.isArray(json.data.messages)) {
        return json.data.messages;
      }
    }
  } catch (e) {}

  // 2. Fallback to local file
  try {
    if (fs.existsSync(LOCAL_MESSAGES_FILE)) {
      return JSON.parse(fs.readFileSync(LOCAL_MESSAGES_FILE, 'utf8')) || [];
    }
  } catch (e) {}

  return [];
}

async function saveStoredMessages(list) {
  const trimmed = list.slice(-1000);

  // 1. Save to cloud store
  try {
    await fetch(CLOUD_STORE_URL, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: 'gadgetmart_watzap_inbox',
        data: { messages: trimmed }
      })
    });
  } catch (e) {}

  // 2. Save to local fallback
  try {
    fs.writeFileSync(LOCAL_MESSAGES_FILE, JSON.stringify(trimmed, null, 2), 'utf8');
  } catch (e) {}
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  // GET: Fetch incoming messages for frontend sync
  if (req.method === 'GET') {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    const since = parseInt(url.searchParams.get('since') || '0', 10);
    const phoneFilter = (url.searchParams.get('phone') || '').replace(/\D/g, '');

    let msgs = await getStoredMessages();
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
      console.log('📥 [WATZAP WEBHOOK RECEIVED]:', JSON.stringify(body));

      let rawPhone = body.phone || body.from || body.sender || body.phone_no || body.wa_number || '';
      let text = body.message_text || body.message || body.body || body.text || '';
      const messageId = body.message_id || body.id || ('wz_' + Date.now());

      if (!rawPhone && body.data) {
        rawPhone = body.data.phone || body.data.from || body.data.sender || body.data.phone_no || '';
      }
      if (!text && body.data) {
        text = body.data.message_text || body.data.message || body.data.body || body.data.text || '';
      }

      let ts = Date.now();
      if (body.timestamp) {
        const parsed = Number(body.timestamp);
        ts = parsed > 10000000000 ? parsed : parsed * 1000;
      }

      const cleanPhone = String(rawPhone).replace(/\D/g, '');

      if (cleanPhone && text) {
        const msgs = await getStoredMessages();
        const exists = msgs.some(m => m.messageId === messageId);
        if (!exists) {
          msgs.push({
            messageId: messageId,
            phone: cleanPhone,
            message: text,
            sender: 'PIC',
            type: body.type || 'text',
            timestamp: ts
          });
          await saveStoredMessages(msgs);
        }
      }

      return res.status(200).json({
        status: 200,
        message: 'Webhook processed successfully'
      });
    } catch (err) {
      console.error('Webhook error:', err);
      return res.status(200).json({ status: 200, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
