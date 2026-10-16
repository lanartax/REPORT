const GOOGLE_API_KEY_DEFAULT = 'AQ.Ab8RN6Lvfri-gb4jU1QFWZYHQyb3pprbRkGJCoV020aWIWCsXA';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  let body = req.body || {};
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch(e){}
  }

  const prompt = body.prompt || '';
  const systemInstruction = body.systemInstruction || '';
  const apiKey = body.apiKey || process.env.GEMINI_API_KEY || GOOGLE_API_KEY_DEFAULT;
  const requestedModel = body.model || 'gemini-3.5-flash-lite';

  if (!prompt) {
    return res.status(400).json({ success: false, error: 'Prompt is required' });
  }

  const modelsToTry = [
    requestedModel,
    'gemini-3.5-flash-lite',
    'gemini-3.1-flash-lite',
    'gemini-3.8-flash'
  ];
  const uniqueModels = Array.from(new Set(modelsToTry.filter(Boolean)));

  const fullPromptText = systemInstruction 
    ? `${systemInstruction}\n\nPertanyaan User: "${prompt}"` 
    : prompt;

  for (const m of uniqueModels) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 12000);

      const url = `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${apiKey}`;
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: fullPromptText }] }]
        }),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      const data = await response.json();
      if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) {
        const text = data.candidates[0].content.parts.map(p => p.text || '').join('');
        return res.status(200).json({
          success: true,
          modelUsed: m,
          text: text
        });
      } else if (data.error) {
        console.warn(`Gemini server model ${m} returned error:`, data.error.message || data.error);
      }
    } catch (err) {
      console.warn(`Fetch error for model ${m}:`, err.message);
    }
  }

  return res.status(500).json({
    success: false,
    error: 'Semua model Gemini sedang sibuk atau tidak merespons.'
  });
};
