const fs = require('fs');
const path = require('path');

const MESSAGES_FILE = path.join(process.cwd(), 'watzap_messages.json');

function getStoredMessages() {
  try {
    if (fs.existsSync(MESSAGES_FILE)) {
      const raw = fs.readFileSync(MESSAGES_FILE, 'utf8');
      return JSON.parse(raw) || [];
    }
  } catch (e) {}
  return [];
}

function saveStoredMessages(list) {
  try {
    // Keep at most 1000 recent messages
    const trimmed = list.slice(-1000);
    fs.writeFileSync(MESSAGES_FILE, JSON.stringify(trimmed, null, 2), 'utf8');
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

    let msgs = getStoredMessages();
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

      const rawPhone = body.phone || body.from || body.sender || body.phone_no || '';
      const text = body.message_text || body.message || body.body || body.text || '';
      const messageId = body.message_id || body.id || ('wz_' + Date.now());

      let ts = Date.now();
      if (body.timestamp) {
        const parsed = Number(body.timestamp);
        ts = parsed > 10000000000 ? parsed : parsed * 1000;
      }

      const cleanPhone = String(rawPhone).replace(/\D/g, '');

      if (cleanPhone && text) {
        const msgs = getStoredMessages();
        // Check duplicate message_id
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
          saveStoredMessages(msgs);
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

