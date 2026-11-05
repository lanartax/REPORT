module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const file = url.searchParams.get('file');

  if (!file) {
    return res.status(400).send('Parameter file is required');
  }

  try {
    // Bersihkan path file agar aman
    const cleanFile = file.replace(/^[/\\]+/, '');
    const wahaUrl = `http://178.83.188.100:3088/api/files/${cleanFile}`;

    const upstream = await fetch(wahaUrl, {
      headers: { 'X-Api-Key': 'gadgetmart123' },
      signal: AbortSignal.timeout(10000)
    });

    if (!upstream.ok) {
      return res.status(upstream.status).send(`Failed to fetch file from WAHA: ${upstream.statusText}`);
    }

    const contentType = upstream.headers.get('content-type') || 'image/jpeg';
    res.setHeader('Content-Type', contentType);
    res.setHeader('Cache-Control', 'public, max-age=31536000, immutable');

    const arrayBuf = await upstream.arrayBuffer();
    return res.status(200).send(Buffer.from(arrayBuf));
  } catch (err) {
    console.error('Media proxy error:', err);
    return res.status(500).send('Internal media proxy error');
  }
};
