const fs = require('fs');
const path = require('path');

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

    const payload = {
      api_key: WATZAP_API_KEY,
      number_key: WATZAP_NUMBER_KEY,
      phone_no: cleanPhone,
      message: message,
      wait_until_send: '1'
    };

    console.log(`📤 [WATZAP SENDING MESSAGE] To: ${cleanPhone}`);

    const wzRes = await fetch('https://api.watzap.id/v1/send_message', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const wzJson = await wzRes.json();
    console.log('WatZap send_message result:', wzJson);

    // Record to local message log
    appendStoredMessage({
      messageId: 'out_' + Date.now(),
      phone: cleanPhone,
      message: message,
      sender: 'ADMIN',
      type: 'text',
      timestamp: Date.now()
    });

    return res.status(200).json({
      success: wzJson.status === '200' || wzJson.status === 200 || wzJson.ack === 'successfully',
      watzap: wzJson,
      phone: cleanPhone
    });
  } catch (err) {
    console.error('Error sending WatZap message:', err);
    return res.status(500).json({ success: false, message: err.message });
  }
};

