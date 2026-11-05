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
  } catch (e) { }
  return [];
}

function appendLocalMessage(msg) {
  try {
    const list = getLocalMessages();
    const existingIdx = list.findIndex(m => m.messageId && m.messageId === msg.messageId);
    if (existingIdx >= 0) {
      list[existingIdx] = { ...list[existingIdx], ...msg };
    } else {
      list.push(msg);
    }
    fs.writeFileSync(LOCAL_MESSAGES_FILE, JSON.stringify(list.slice(-500), null, 2), 'utf8');
  } catch (e) { }
}

// Parser cerdas untuk mendeteksi kutipan teks WhatsApp (> *Pengirim*:\n> Isi...)
function parseQuotedFromText(rawText) {
  if (!rawText || typeof rawText !== 'string') return { cleanText: rawText || '', replyTo: null };
  const trimmed = rawText.trim();
  if (!trimmed.startsWith('>')) return { cleanText: rawText, replyTo: null };

  // Format 1: > *Sender*:\n> Quoted message\n\nActual reply
  const m1 = trimmed.match(/^>(?: \*(.*?)\*:?)?\s*([\s\S]*?)\n\n([\s\S]*)$/);
  if (m1) {
    const sender = (m1[1] || 'Pesan').trim();
    const quoted = m1[2].replace(/^>\s?/gm, '').trim();
    const actual = m1[3].trim();
    return {
      cleanText: actual,
      replyTo: { senderName: sender, text: quoted }
    };
  }

  // Format 2: > [Sender] Quoted message\n\nActual reply
  const m2 = trimmed.match(/^> \[(.*?)\]\s*([\s\S]*?)\n\n([\s\S]*)$/);
  if (m2) {
    const sender = (m2[1] || 'Pesan').trim();
    const quoted = m2[2].replace(/^>\s?/gm, '').trim();
    const actual = m2[3].trim();
    return {
      cleanText: actual,
      replyTo: { senderName: sender, text: quoted }
    };
  }

  // Format 3: > Quoted message\n\nActual reply
  const m3 = trimmed.match(/^> ([\s\S]*?)\n\n([\s\S]*)$/);
  if (m3) {
    const quoted = m3[1].replace(/^>\s?/gm, '').trim();
    const actual = m3[2].trim();
    return {
      cleanText: actual,
      replyTo: { senderName: 'Pesan', text: quoted }
    };
  }

  return { cleanText: rawText, replyTo: null };
}

// Cache & resolver untuk memetakan WhatsApp LID (@lid) menjadi nomor telepon asli (@c.us)
const wahaLidCache = new Map();
const wahaContactNameCache = new Map();

async function resolveWahaLidToPhone(lidJid) {
  if (!lidJid) return '';
  const str = String(lidJid).trim();
  if (!str.endsWith('@lid')) {
    let p = str.replace(/@.*$/, '').replace(/\D/g, '');
    if (p.startsWith('08')) p = '62' + p.slice(1);
    else if (p.startsWith('8')) p = '62' + p;
    return p;
  }

  if (wahaLidCache.has(str)) {
    return wahaLidCache.get(str);
  }

  try {
    const res = await fetch(`http://178.83.188.100:3088/api/contacts?session=default&contactId=${encodeURIComponent(str)}`, {
      headers: { 'X-Api-Key': 'gadgetmart123' },
      signal: AbortSignal.timeout(3500)
    });
    if (res.ok) {
      const data = await res.json();
      const rawNum = data.id || '';
      let clean = String(rawNum).replace(/@.*$/, '').replace(/\D/g, '');
      if (clean && clean.length >= 9 && clean.length <= 15) {
        if (clean.startsWith('08')) clean = '62' + clean.slice(1);
        else if (clean.startsWith('8')) clean = '62' + clean;
        wahaLidCache.set(str, clean);
        const cName = String(data.pushname || data.name || data.shortName || '').trim();
        if (cName) {
          wahaContactNameCache.set(str, cName);
          wahaContactNameCache.set(clean, cName);
        }
        return clean;
      }
    }
  } catch (e) {
    console.warn('LID resolve warning:', e.message);
  }

  let fallback = str.replace(/@.*$/, '').replace(/\D/g, '');
  if (fallback.startsWith('08')) fallback = '62' + fallback.slice(1);
  else if (fallback.startsWith('8')) fallback = '62' + fallback;
  return fallback;
}

// Helper untuk men-download gambar dari WAHA VPS dan mengubahnya menjadi Base64 Data URI permanen
async function fetchAndEncodeMedia(url, fallbackMime = 'image/jpeg') {
  if (!url || typeof url !== 'string') return '';
  if (url.startsWith('data:')) return url;

  // Ganti localhost:3000 jika ada ke VPS IP 178.83.188.100:3088
  let targetUrl = url.replace(/http:\/\/localhost:3000/g, 'http://178.83.188.100:3088');

  try {
    const res = await fetch(targetUrl, {
      headers: { 'X-Api-Key': 'gadgetmart123' },
      signal: AbortSignal.timeout(7000)
    });
    if (res.ok) {
      const mime = res.headers.get('content-type') || fallbackMime;
      const arrayBuf = await res.arrayBuffer();
      const base64 = Buffer.from(arrayBuf).toString('base64');
      return `data:${mime};base64,${base64}`;
    }
  } catch (e) {
    console.warn('fetchAndEncodeMedia warning:', e.message);
  }
  return targetUrl;
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
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp", media_url as "mediaUrl", message_type as "messageType", reply_to as "replyTo"
              FROM watzap_inbox
              WHERE is_group = FALSE 
                AND created_at > ${since}
                AND (phone = ${phoneFilter} OR phone LIKE ${'%' + last9})
              ORDER BY created_at ASC
              LIMIT 200
            `;
          } else {
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp", media_url as "mediaUrl", message_type as "messageType", reply_to as "replyTo"
              FROM watzap_inbox
              WHERE is_group = FALSE AND created_at > ${since}
              ORDER BY created_at ASC
              LIMIT 200
            `;
          }
        } else {
          // Fetch initial recent messages (ambil 500 pesan TERBARU, lalu urutkan ASC untuk playback)
          if (phoneFilter) {
            const last9 = phoneFilter.slice(-9);
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp", media_url as "mediaUrl", message_type as "messageType", reply_to as "replyTo"
              FROM watzap_inbox
              WHERE is_group = FALSE AND (phone = ${phoneFilter} OR phone LIKE ${'%' + last9})
              ORDER BY created_at DESC
              LIMIT 500
            `;
          } else {
            msgs = await sql`
              SELECT message_id as "messageId", phone, sender_name as "senderName", message, sender, is_group as "isGroup", created_at as "timestamp", media_url as "mediaUrl", message_type as "messageType", reply_to as "replyTo"
              FROM watzap_inbox
              WHERE is_group = FALSE
              ORDER BY created_at DESC
              LIMIT 500
            `;
          }
          if (Array.isArray(msgs)) msgs.reverse();
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

    // Strict filter: pastikan bukan grup dan jadikan timestamp sebagai Number & phone ternormalisasi
    const filtered = (msgs || []).filter(m => {
      const p = String(m.phone || '');
      return !m.isGroup && !p.startsWith('120363') && p.length <= 14;
    }).map(m => {
      let p = String(m.phone || '').replace(/\D/g, '');
      if (p.startsWith('08')) p = '62' + p.slice(1);
      else if (p.startsWith('8')) p = '62' + p;

      let replyToObj = m.replyTo || null;
      let msgText = m.message || '';

      // Parsing otomatis kutipan jika ada di dalam teks
      if (!replyToObj && msgText.startsWith('>')) {
        const parsed = parseQuotedFromText(msgText);
        if (parsed.replyTo) {
          replyToObj = parsed.replyTo;
          msgText = parsed.cleanText;
        }
      }

      return {
        ...m,
        phone: p,
        message: msgText,
        text: msgText,
        timestamp: Number(m.timestamp || m.created_at || Date.now()),
        mediaUrl: m.mediaUrl || m.media_url || '',
        messageType: m.messageType || m.message_type || (m.mediaUrl || m.media_url ? 'image' : 'text'),
        replyTo: replyToObj
      };
    });

    let resolvedChatsMap = {};
    if (sql) {
      try {
        const statusRows = await sql`
          SELECT phone, is_resolved, resolved_at, resolved_by
          FROM omni_chat_status
          WHERE is_resolved = TRUE
        `;
        for (const row of statusRows) {
          resolvedChatsMap[row.phone] = {
            isResolved: true,
            resolvedAt: Number(row.resolved_at || Date.now()),
            resolvedBy: row.resolved_by || 'Admin'
          };
        }
      } catch (statusErr) {}
    }

    return res.status(200).json({
      success: true,
      count: filtered.length,
      messages: filtered,
      resolvedChats: resolvedChatsMap
    });
  }

  // POST: Webhook payload dari WatZap saat ada chat masuk
  if (req.method === 'POST') {
    try {
      const body = req.body || {};

      // Cek apakah aksi untuk resolve / unresolve chat secara permanen
      if (body.action === 'resolve_chat' || body.action === 'unresolve_chat') {
        let targetPhone = String(body.phone || body.picKey || '').replace(/\D/g, '');
        if (targetPhone.startsWith('08')) targetPhone = '62' + targetPhone.slice(1);
        else if (targetPhone.startsWith('8')) targetPhone = '62' + targetPhone;

        const isResolved = body.action === 'resolve_chat' || body.isResolved === true;
        const resolvedBy = String(body.resolvedBy || 'Admin').trim();
        const now = Date.now();

        if (sql && targetPhone) {
          try {
            await initDb(sql);
            await sql`
              INSERT INTO omni_chat_status (phone, is_resolved, resolved_at, resolved_by, updated_at)
              VALUES (${targetPhone}, ${isResolved}, ${now}, ${resolvedBy}, ${now})
              ON CONFLICT (phone) DO UPDATE SET
                is_resolved = EXCLUDED.is_resolved,
                resolved_at = EXCLUDED.resolved_at,
                resolved_by = EXCLUDED.resolved_by,
                updated_at = EXCLUDED.updated_at
            `;
          } catch (dbErr) {
            console.error('Neon DB resolve_chat Error:', dbErr);
          }
        }
        return res.status(200).json({ success: true, phone: targetPhone, isResolved, resolvedAt: now });
      }

      const data = body.data || body;

      // Debug logger: Simpan raw webhook payload ke watzap_webhook_logs untuk analisa real-time
      if (sql) {
        try {
          await initDb(sql);
          await sql`
            INSERT INTO watzap_webhook_logs (payload, created_at)
            VALUES (${JSON.stringify(body)}, ${Date.now()})
          `;
        } catch (logErr) { }
      }

      // =========================================================================
      // DUKUNGAN RESMI WAHA (WhatsApp HTTP API di VPS) DENGAN FULL QUOTED REPLY
      // =========================================================================
      if (body.event && (body.event === 'message' || body.event === 'message.any') && body.payload) {
        const p = body.payload;
        const isFromMe = Boolean(p.fromMe);
        const senderType = isFromMe ? 'ADMIN' : 'PIC';

        // Tentukan JID lawan bicara (PIC/Customer)
        // Jika dikirim dari bot/admin (fromMe=true), lawan bicaranya ada di p.to / remote
        // Jika pesan masuk (fromMe=false), lawan bicaranya ada di p.from / remote
        const targetJid = isFromMe
          ? String(p.to || p._data?.id?.remote || p._data?.to || '')
          : String(p.from || p._data?.id?.remote || '');

        // Abaikan grup WhatsApp & WhatsApp Status/Broadcast
        const isGroupOrBroadcast = targetJid.includes('@g.us') ||
          targetJid.includes('broadcast') ||
          String(p.from || '').includes('@g.us') ||
          String(p.from || '').includes('broadcast') ||
          String(p.to || '').includes('@g.us') ||
          String(p.to || '').includes('broadcast');

        if (isGroupOrBroadcast) {
          return res.status(200).json({ status: 200, message: 'Group or broadcast ignored' });
        }

        // Resolusi nomor telepon asli dari WhatsApp LID (@lid -> @c.us)
        let cleanPhone = await resolveWahaLidToPhone(targetJid);
        if (!cleanPhone || cleanPhone.length < 9) {
          cleanPhone = targetJid.replace(/@.*$/, '').replace(/\D/g, '');
        }
        if (cleanPhone.startsWith('08')) cleanPhone = '62' + cleanPhone.slice(1);
        else if (cleanPhone.startsWith('8')) cleanPhone = '62' + cleanPhone;

        if (!cleanPhone || cleanPhone.length < 9) {
          return res.status(200).json({ status: 200, message: 'Invalid phone or system contact ignored' });
        }

        const messageId = String(p.id || ('waha_' + Date.now()));

        // Tentukan nama pengirim
        let senderName = '';
        if (isFromMe) {
          senderName = 'Admin';
        } else {
          senderName = String(p._data?.notifyName || p.pushname || p.author || '').trim();
          if (!senderName) {
            senderName = wahaContactNameCache.get(targetJid) || wahaContactNameCache.get(cleanPhone) || 'PIC';
          }
        }

        let ts = Date.now();
        if (p.timestamp) {
          ts = Number(p.timestamp) > 10000000000 ? Number(p.timestamp) : Number(p.timestamp) * 1000;
        }

        // Ekstraksi & resolusi media gambar/foto agar menjadi Data URI permanen
        let mediaUrl = '';
        if (p.hasMedia && p.media) {
          if (typeof p.media === 'string') mediaUrl = p.media;
          else if (p.media.url) mediaUrl = p.media.url;
          else if (p.media.data) {
            const mime = p.media.mimetype || 'image/jpeg';
            mediaUrl = `data:${mime};base64,${p.media.data}`;
          }
        }
        if (mediaUrl && !mediaUrl.startsWith('data:')) {
          mediaUrl = await fetchAndEncodeMedia(mediaUrl, p.media?.mimetype || 'image/jpeg');
        }

        const isImage = Boolean(p.hasMedia || mediaUrl);
        const messageType = isImage ? 'image' : 'text';
        let text = String(p.body || '').trim();
        if (!text && isImage) text = '[Foto/Gambar]';

        // QUOTE RESMI DARI WAHA (100% Mengambil Teks yang Di-quote Pengirim!)
        let replyTo = null;
        if (p.replyTo && (p.replyTo.body || p.replyTo.id)) {
          const quotedSenderJid = String(p.replyTo.participant || '');
          let quotedCleanPhone = '';
          if (quotedSenderJid) {
            quotedCleanPhone = await resolveWahaLidToPhone(quotedSenderJid);
          }
          let quotedSender = 'Admin';
          if (quotedCleanPhone) {
            quotedSender = quotedCleanPhone === cleanPhone ? (senderName || 'PIC') : 'Admin';
          } else {
            quotedSender = isFromMe ? (senderName || 'PIC') : 'Admin';
          }

          replyTo = {
            senderName: quotedSender,
            text: String(p.replyTo.body || '').trim() || 'Pesan'
          };
        }

        if (sql) {
          await initDb(sql);
          await sql`
            INSERT INTO watzap_inbox (message_id, phone, sender_name, message, sender, is_group, created_at, media_url, message_type, reply_to, raw_payload)
            VALUES (${messageId}, ${cleanPhone}, ${senderName}, ${text}, ${senderType}, FALSE, ${ts}, ${mediaUrl || null}, ${messageType}, ${replyTo ? JSON.stringify(replyTo) : null}, ${JSON.stringify(body)})
            ON CONFLICT (message_id) DO UPDATE SET
              phone = EXCLUDED.phone,
              media_url = COALESCE(EXCLUDED.media_url, watzap_inbox.media_url),
              message_type = EXCLUDED.message_type,
              reply_to = COALESCE(EXCLUDED.reply_to, watzap_inbox.reply_to),
              sender_name = COALESCE(EXCLUDED.sender_name, watzap_inbox.sender_name)
          `;
        }

        appendLocalMessage({
          messageId: messageId,
          phone: cleanPhone,
          message: text,
          sender: senderType,
          senderName: senderName,
          isGroup: false,
          timestamp: ts,
          mediaUrl: mediaUrl,
          messageType: messageType,
          replyTo: replyTo
        });

        return res.status(200).json({
          status: 200,
          source: 'WAHA',
          message: 'WAHA chat processed successfully',
          messageId: messageId,
          hasReply: Boolean(replyTo),
          hasMedia: Boolean(mediaUrl)
        });
      }

      // Deteksi format Meta WABA vs WatZap Unofficial
      let msgItem = data;
      let contactName = '';
      if (body.entry && body.entry[0] && body.entry[0].changes && body.entry[0].changes[0] && body.entry[0].changes[0].value) {
        const val = body.entry[0].changes[0].value;
        if (Array.isArray(val.messages) && val.messages.length > 0) {
          msgItem = val.messages[0];
          if (Array.isArray(val.contacts) && val.contacts[0] && val.contacts[0].profile) {
            contactName = val.contacts[0].profile.name || '';
          }
        }
      }

      const rawChatId = String(
        msgItem.chat_id || msgItem.phone || msgItem.from || msgItem.sender || msgItem.phone_no ||
        data.chat_id || data.phone || data.from || data.sender || data.phone_no ||
        body.phone || body.from || ''
      );
      let cleanPhone = rawChatId.replace(/@.*$/, '').replace(/\D/g, '');

      // DETEKSI GRUP WHATSAPP: TOLAK & ABAIKAN GRUP
      const isGroup = rawChatId.includes('@g.us') || cleanPhone.startsWith('120363') || cleanPhone.length > 14;
      if (isGroup) {
        return res.status(200).json({
          status: 200,
          message: 'Group message ignored (Personal PIC chats only)'
        });
      }

      // Normalisasi nomor ke format 628...
      if (cleanPhone.startsWith('08')) {
        cleanPhone = '62' + cleanPhone.slice(1);
      } else if (cleanPhone.startsWith('8') && cleanPhone.length >= 9 && cleanPhone.length <= 13) {
        cleanPhone = '62' + cleanPhone;
      }

      // Validasi nomor WhatsApp personal (minimal 9 digit, maksimal 14 digit)
      if (!cleanPhone || cleanPhone.length < 9) {
        return res.status(200).json({ status: 200, message: 'Invalid phone ignored' });
      }

      const messageId = String(
        msgItem.id || data.message_id || data.id || body.messageId ||
        ('wz_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7))
      );

      // 1. Ekstraksi Gambar / Media URL
      let mediaUrl = '';
      if (typeof msgItem.image === 'string') mediaUrl = msgItem.image;
      else if (msgItem.image && (msgItem.image.url || msgItem.image.link)) mediaUrl = msgItem.image.url || msgItem.image.link;

      if (!mediaUrl && typeof data.image === 'string') mediaUrl = data.image;
      else if (!mediaUrl && data.image && (data.image.url || data.image.link)) mediaUrl = data.image.url || data.image.link;

      if (!mediaUrl) {
        mediaUrl = String(
          msgItem.media_url || msgItem.url || msgItem.file_url || msgItem.link || msgItem.file ||
          data.media_url || data.url || data.file_url || data.link || data.file ||
          (data.media && (data.media.url || data.media.link || data.media)) ||
          (data.message_payload && String(data.message_payload).startsWith('http') ? data.message_payload : '') ||
          body.image || body.media_url || body.url || ''
        ).trim();
      }

      // 2. Ekstraksi Tipe Pesan & Deteksi Media
      const hasMediaFlag = data.has_media === true || body.has_media === true || msgItem.has_media === true;
      const mediaMime = String(data.media_mime || msgItem.media_mime || body.media_mime || '').toLowerCase();
      const rawType = String(
        msgItem.type || msgItem.message_type || data.message_type || data.type || body.message_type ||
        (mediaUrl ? 'image' : 'text')
      ).toLowerCase();
      const isImage = rawType.includes('image') || rawType.includes('foto') || hasMediaFlag || mediaMime.includes('image') || Boolean(mediaUrl) || String(data.message_body) === '[Image]';
      const messageType = isImage ? 'image' : 'text';

      // 3. JIKA WATZAP MENGIRIM MEDIA TANPA URL -> LANGSUNG UNDUH VIA API /v1/download_media
      if (isImage && !mediaUrl && messageId) {
        try {
          const downloadPayload = {
            api_key: process.env.WATZAP_API_KEY || 'VMUXOT68QD5TQ84V',
            number_key: data.number_key || body.number_key || process.env.WATZAP_NUMBER_KEY || 'gqLbEDAZ4lReCCME',
            chat_id: rawChatId || `${cleanPhone}@s.whatsapp.net`,
            message_id: messageId
          };
          const dlRes = await fetch('https://api.watzap.id/v1/download_media', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(downloadPayload)
          });
          const dlData = await dlRes.json();
          if (dlData && dlData.url) {
            mediaUrl = dlData.url;
            console.log(`📸 [WatZap Webhook] Berhasil unduh media untuk ${cleanPhone} (${mediaUrl.length} chars)`);
          }
        } catch (dlErr) {
          console.error('⚠️ [WatZap Webhook] Gagal download media dari WatZap:', dlErr.message);
        }
      }

      // 4. Ekstraksi Pesan Teks
      let text = String(
        msgItem.text?.body || msgItem.message_body || msgItem.message ||
        data.message_body || data.message_text || data.message || data.body || data.text ||
        body.message || ''
      ).trim();

      if (!text && isImage) {
        text = '[Foto/Gambar]';
      }

      // 5. Ekstraksi Quoted Reply (Hanya jika asli ada di context atau format tanda kutip >)
      let replyTo = null;
      const ctx = msgItem.context || data.context || data.quoted || data.quoted_message || data.quoted_body || data.reply_to || null;
      if (ctx && typeof ctx === 'object') {
        const quotedText = ctx.body || ctx.message || ctx.text || ctx.quoted_body || '';
        const quotedSender = ctx.name || ctx.sender_name || ctx.sender || (ctx.from === cleanPhone ? (data.name || 'PIC') : 'Admin');
        if (quotedText) {
          replyTo = { senderName: quotedSender, text: quotedText };
        }
      }

      // Jika pesan teks menggunakan format kutipan WhatsApp bawaan (> *Nama*: ...)
      if (!replyTo && text.startsWith('>')) {
        const parsed = parseQuotedFromText(text);
        if (parsed.replyTo) {
          replyTo = parsed.replyTo;
          text = parsed.cleanText;
        }
      }

      const senderName = String(contactName || data.name || data.pushName || msgItem.name || '').trim();
      const isFromMe = data.is_from_me === true || body.is_from_me === true || msgItem.from_me === true;

      let ts = Date.now();
      if (msgItem.timestamp || data.timestamp || body.timestamp) {
        const parsed = Number(msgItem.timestamp || data.timestamp || body.timestamp);
        ts = parsed > 10000000000 ? parsed : parsed * 1000;
      }

      const senderType = isFromMe ? 'ADMIN' : 'PIC';

      // Pastikan pesan selalu disimpan
      if (!text) text = isImage ? '[Foto/Gambar]' : '[Pesan Masuk]';

      if (sql) {
        try {
          await initDb(sql);
          await sql`
            INSERT INTO watzap_inbox (message_id, phone, sender_name, message, sender, is_group, created_at, media_url, message_type, reply_to, raw_payload)
            VALUES (${messageId}, ${cleanPhone}, ${senderName}, ${text}, ${senderType}, FALSE, ${ts}, ${mediaUrl || null}, ${messageType}, ${replyTo ? JSON.stringify(replyTo) : null}, ${JSON.stringify(body)})
            ON CONFLICT (message_id) DO UPDATE SET
              media_url = COALESCE(EXCLUDED.media_url, watzap_inbox.media_url),
              message_type = EXCLUDED.message_type,
              reply_to = COALESCE(EXCLUDED.reply_to, watzap_inbox.reply_to)
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
        timestamp: ts,
        mediaUrl: mediaUrl,
        messageType: messageType,
        replyTo: replyTo
      });

      return res.status(200).json({
        status: 200,
        message: 'Personal chat processed successfully',
        messageId: messageId,
        messageType: messageType,
        hasMedia: Boolean(mediaUrl),
        hasReply: Boolean(replyTo)
      });
    } catch (err) {
      console.error('Webhook error:', err);
      return res.status(200).json({ status: 200, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
